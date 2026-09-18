#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""语义化 HTML → 原生可编辑 PPTX（文字可编辑，非截图版）。

用法:
  python html2editable_pptx.py input.html output.pptx

原理:
  1. playwright 打开 HTML（1280x720 视口），先让所有 .slide active 并等过渡结束
     （消除非活跃页 transform: translateY(24px) 导致的坐标偏移——标题偏移根源）
  2. 读取每页元素的语义结构: tag/class/文本/位置/字号/颜色/字体栈/行距/runs
  3. python-pptx 按类型重建原生元素（文本框 + 形状），全部可编辑

关键修复记录:
  - 字体: 不硬编码"Microsoft YaHei"（本机无此字体会被替换成 Noto → 换行全变 → 偏移）。
    改用 pick_font() 取 HTML font-family 栈第一个具体家族，Windows/WPS→微软雅黑、
    macOS→苹方、Linux→Noto Sans SC，与浏览器渲染 HTML 一致。
  - 文本框默认 internal margin(0.1/0.05in) 会让文字相对框偏移 → 全部置 0。
  - auto_size = SHAPE_TO_FIT_TEXT，让 WPS 按文字自适应高度，避免溢出。
  - 行距跟随 HTML line-height ratio（如 1.7/1.6）。
  - 标题拆 runs 保留 .hl 强调色。
  - 嵌套文本（.txt 含 <small>）用 nodeText 递归拼接，block 子元素前后换行。

依赖: playwright + python-pptx + Pillow
"""
import sys, pathlib, os, re
from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

# emoji 字符范围：只匹配"真 emoji"（彩色 emoji 区 + 变体选择符）。
# ⚠️ 不要包含箭头(→ U+2192)/间隔号(· U+00B7)/破折号(— U+2014)/★✓✕ 等普通符号——
# 这些中文字体（微软雅黑/PingFang/Noto CJK）都有字形，拆到 Segoe UI Emoji 反而会乱码
# （Segoe UI Emoji 不含箭头等字形）。
EMOJI_RE = re.compile(r"[\U0001F000-\U0001FAFF\uFE0F]")
EMOJI_FONT = "Segoe UI Emoji"  # Windows/WPS 彩色渲染 emoji；Mac WPS 自动 fallback Apple Color Emoji

HTML_W, HTML_H = 1280, 720
SLIDE_W, SLIDE_H = Inches(10), Inches(5.625)

# 通用字体家族（跳过，取后面具体的）
GENERIC_FONTS = {
    "sans-serif", "serif", "monospace", "cursive", "fantasy", "system-ui",
    "ui-sans-serif", "ui-serif", "ui-monospace", "-apple-system", "blinkmacsystemfont",
    "segoe ui", "roboto", "helvetica neue", "helvetica", "arial", "noto sans",
    "sans", "pingfang sc", "microsoft yahei", "yahei",
}

def pick_font(stack):
    """优先选目标平台常用字体（Windows/WPS→微软雅黑），否则取栈中第一个具体家族。

    用户在 Windows/WPS 打开 PPTX 时，浏览器渲染 HTML 会按栈回退到 Microsoft YaHei，
    所以 PPTX 写 Microsoft YaHei 与 HTML 在目标机器上的观感最一致。
    """
    stack_lower = (stack or "").lower()
    if "microsoft yahei" in stack_lower or '"microsoft yahei"' in stack_lower:
        return "Microsoft YaHei"
    if "pingfang sc" in stack_lower:
        return "PingFang SC"
    if "noto sans sc" in stack_lower:
        return "Noto Sans SC"
    for part in (stack or "").split(","):
        name = part.strip().strip('"').strip("'").strip()
        if name and name.lower() not in GENERIC_FONTS:
            return name
    return "Microsoft YaHei"

def px_to_in(x, y, w, h):
    return (Inches(x / HTML_W * 10), Inches(y / HTML_H * 5.625),
            Inches(w / HTML_W * 10), Inches(h / HTML_H * 5.625))

def _parse_css_color(css):
    """解析 rgb()/rgba()/hex → (r, g, b, a)，失败返回 None。"""
    if not css:
        return None
    s = css.strip()
    if s.startswith("rgb"):
        inner = s[s.index("(") + 1 : s.rindex(")")]
        parts = [p.strip() for p in inner.split(",")]
        if len(parts) < 3:
            return None
        r, g, b = int(float(parts[0])), int(float(parts[1])), int(float(parts[2]))
        a = float(parts[3]) if len(parts) > 3 else 1.0
        return (r, g, b, a)
    if s.startswith("#"):
        s = s[1:]
        if len(s) == 3:
            s = "".join(c * 2 for c in s)
        if len(s) >= 6:
            return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16), 1.0)
    return None

def to_rgb(css, bg=None):
    """css → RGBColor。rgba 半透明色与 bg（默认深色底）混合，避免透明色盖住背景。"""
    c = _parse_css_color(css)
    if not c:
        return RGBColor(0xE8, 0xEC, 0xF4)
    r, g, b, a = c
    if a < 1.0:
        br, bg_2, bb = bg if bg else (0x0E, 0x15, 0x26)
        r = int(br * (1 - a) + r * a)
        g = int(bg_2 * (1 - a) + g * a)
        b = int(bb * (1 - a) + b * a)
    return RGBColor(max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))

def main():
    html_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else str(pathlib.Path(html_path).with_suffix("_editable.pptx"))
    # 背景图（可选第 3 参）：照片/网格图任意，缺省用 /tmp/bg_grid.png
    bg_image = sys.argv[3] if len(sys.argv) > 3 else "/tmp/bg_grid.png"

    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={"width": HTML_W, "height": HTML_H})
        pg.goto(f"file://{pathlib.Path(html_path).resolve()}")
        pg.wait_for_timeout(600)

        # ⚠️ 关键：所有 slide active + 等过渡结束，消除 translateY(24px) 坐标偏移
        pg.evaluate("document.querySelectorAll('.slide').forEach(s => s.classList.add('active'))")
        pg.wait_for_timeout(700)

        slides_data = pg.evaluate("""
        () => {
          const NL = String.fromCharCode(10);
          const out = [];
          const CONTAINER = ['hl','foot','avatar','card','vs-card','step','route','arrow',
                             'grid','g2','g3','g4','list','item','two','dot','pager','hint',
                             'progress','slide','deck'];
          // 递归拼接文本：文本节点直取、<br>转行、block 子元素前后换行
          const nodeText = (n) => {
            if (n.nodeType === 3) return n.textContent;
            if (n.nodeName === 'BR') return NL;
            if (n.nodeName === 'A') return '';  // 链接文本由 a 单独抓取（带 href）
            const cs = getComputedStyle(n);
            const d = cs.display;
            const inner = [...n.childNodes].map(nodeText).join('');
            return (d === 'block' || d === 'flex' || d === 'grid')
                   ? NL + inner.trim() + NL : inner;
          };

          document.querySelectorAll('.slide').forEach((slide, si) => {
            const page = { index: si, elements: [], blocks: [] };
            const absorbParents = [];
            slide.querySelectorAll('*').forEach(el => {
              const direct = [...el.childNodes].some(n => n.nodeType === 3 && n.textContent.trim());
              const kids = [...el.children].some(c => c.textContent.trim());
              if (direct && kids) absorbParents.push(el);
            });
            slide.querySelectorAll('*').forEach(el => {
              const cs = getComputedStyle(el);
              const r = el.getBoundingClientRect();
              if (r.width < 1 || r.height < 1) return;
              const clsStr = typeof el.className === 'string' ? el.className : '';
              if (el.hasAttribute && el.hasAttribute('data-html2pptx-ignore')) return;
              if (el.closest('h1,h2')) return;
              if (el.closest('table')) return;  // 表格由 table 单独处理，单元格不单独抓
              if (CONTAINER.some(c => clsStr.includes(c))) return;
              const absorbedBy = absorbParents.find(p => p !== el && p.contains(el));
              if (absorbedBy) return;
              const hasDirectText = [...el.childNodes].some(n => n.nodeType === 3 && n.textContent.trim());
              if (!hasDirectText) return;
              const text = [...el.childNodes].map(nodeText).join('')
                           .replace(new RegExp(NL + '{2,}', 'g'), NL).trim();
              if (!text) return;
              const lh = cs.lineHeight;
              page.elements.push({
                tag: el.tagName.toLowerCase(),
                cls: clsStr,
                text: text.slice(0, 500),
                href: el.tagName === 'A' ? (el.href || '') : null,
                x: r.x, y: r.y, w: r.width, h: r.height,
                fontSize: parseFloat(cs.fontSize) || 16,
                color: cs.color,
                bold: cs.fontWeight === '700' || cs.fontWeight === 'bold',
                align: cs.textAlign,
                fontFamily: cs.fontFamily,
                lineHeightRatio: (lh !== 'normal' && lh !== '') ? parseFloat(lh) / (parseFloat(cs.fontSize) || 16) : null
              });
            });
            // h1/h2 整段标题，按子节点拆 runs 保留 .hl 颜色
            slide.querySelectorAll('h1,h2').forEach(h => {
              const cs = getComputedStyle(h);
              const r = h.getBoundingClientRect();
              if (r.width < 1 || r.height < 1) return;
              const runs = [];
              h.childNodes.forEach(n => {
                if (n.nodeType === 3) {
                  if (n.textContent.trim()) runs.push({ text: n.textContent, color: cs.color });
                } else if (n.nodeType === 1 && n.textContent.trim()) {
                  runs.push({ text: n.textContent, color: getComputedStyle(n).color });
                }
              });
              const lh = cs.lineHeight;
              page.elements.push({
                tag: h.tagName.toLowerCase(), cls: '', text: h.textContent.trim().slice(0, 500),
                x: r.x, y: r.y, w: r.width, h: r.height,
                fontSize: parseFloat(cs.fontSize) || 34,
                color: cs.color, bold: true, align: cs.textAlign,
                fontFamily: cs.fontFamily,
                lineHeightRatio: (lh !== 'normal' && lh !== '') ? parseFloat(lh) / (parseFloat(cs.fontSize) || 34) : null,
                runs: runs
              });
            });
            // 背景块检测：任何有背景色或边框的块级元素 → 画圆角矩形（卡片/步骤卡等）
            slide.querySelectorAll('*').forEach(el => {
              const cs = getComputedStyle(el);
              const r = el.getBoundingClientRect();
              if (r.width < 20 || r.height < 20) return;   // 忽略小圆点
              if (el.closest('h1,h2')) return;
              if (el.closest('table')) return;  // 表格单元格的边框不画（表格用原生 add_table）
              if (el.hasAttribute && el.hasAttribute('data-html2pptx-ignore')) return;
              const bg = cs.backgroundColor;
              const hasBg = bg && bg !== 'transparent' && bg !== 'rgba(0, 0, 0, 0)';
              const bt = cs.borderTopWidth || '0px';
              const hasBorder = parseFloat(bt) > 0 && cs.borderTopStyle !== 'none';
              if (!hasBg && !hasBorder) return;
              // 跳过自身是布局根（slide 本身不画）
              if (el.classList.contains('slide')) return;
              const radius = parseFloat(cs.borderTopLeftRadius) || 0;
              page.blocks.push({
                x: r.x, y: r.y, w: r.width, h: r.height,
                bg: hasBg ? bg : null,
                border: hasBorder ? cs.borderTopColor : null,
                radius: radius
              });
            });
            // 图片 <img>：无文本元素，单独抓（本地路径）
            slide.querySelectorAll('img').forEach(el => {
              const r = el.getBoundingClientRect();
              if (r.width < 1 || r.height < 1) return;
              if (el.closest('h1,h2')) return;
              let src = el.currentSrc || el.src || '';
              if (src.startsWith('file://')) src = decodeURIComponent(src.replace('file://', ''));
              page.elements.push({
                tag: 'img', cls: '', text: '', src: src,
                x: r.x, y: r.y, w: r.width, h: r.height,
                fontSize: 16, color: '#e8ecf4', bold: false, align: 'left',
                fontFamily: '', lineHeightRatio: null
              });
            });
            // 视频 <video>：无文本元素，单独抓（本地路径）
            slide.querySelectorAll('video').forEach(el => {
              const r = el.getBoundingClientRect();
              if (r.width < 1 || r.height < 1) return;
              if (el.closest('h1,h2')) return;
              let src = el.currentSrc || (el.querySelector('source') ? el.querySelector('source').src : '') || el.src || '';
              if (src.startsWith('file://')) src = decodeURIComponent(src.replace('file://', ''));
              page.elements.push({
                tag: 'video', cls: '', text: '', src: src,
                x: r.x, y: r.y, w: r.width, h: r.height,
                fontSize: 16, color: '#e8ecf4', bold: false, align: 'left',
                fontFamily: '', lineHeightRatio: null
              });
            });
            // 超链接 <a>：单独抓（带 href + 自身精确位置），父元素文本已跳过它
            slide.querySelectorAll('a').forEach(el => {
              const r = el.getBoundingClientRect();
              if (r.width < 1 || r.height < 1) return;
              if (!el.textContent.trim()) return;
              if (el.closest('h1,h2')) return;
              const cs = getComputedStyle(el);
              page.elements.push({
                tag: 'a', cls: '', text: el.textContent.trim().slice(0, 200),
                href: el.href || '',
                x: r.x, y: r.y, w: r.width, h: r.height,
                fontSize: parseFloat(cs.fontSize) || 16,
                color: cs.color, bold: cs.fontWeight === '700',
                align: cs.textAlign, fontFamily: cs.fontFamily,
                lineHeightRatio: null
              });
            });
            // 表格 <table>：抓行列结构与合并单元格 → PPT 原生表格
            slide.querySelectorAll('table').forEach(tbl => {
              const r = tbl.getBoundingClientRect();
              if (r.width < 1 || r.height < 1) return;
              if (tbl.closest('h1,h2')) return;
              if (tbl.hasAttribute && tbl.hasAttribute('data-html2pptx-ignore')) return;
              const rows = [];
              tbl.querySelectorAll('tr').forEach(tr => {
                const cells = [];
                tr.querySelectorAll('th,td').forEach(td => {
                  cells.push({
                    text: td.textContent.trim().slice(0, 200),
                    rowspan: parseInt(td.getAttribute('rowspan')) || 1,
                    colspan: parseInt(td.getAttribute('colspan')) || 1
                  });
                });
                rows.push(cells);
              });
              if (!rows.length) return;
              page.elements.push({
                tag: 'table', cls: '', text: '', table: rows,
                x: r.x, y: r.y, w: r.width, h: r.height,
                fontSize: 16, color: '#e8ecf4', bold: false, align: 'left',
                fontFamily: '', lineHeightRatio: null
              });
            });
            // Tabler SVG 图标：整段抓取，currentColor 替换成实际色，放大渲染
            slide.querySelectorAll('svg').forEach(el => {
              const r = el.getBoundingClientRect();
              if (r.width < 1 || r.height < 1) return;
              if (el.closest('h1,h2')) return;
              const color = getComputedStyle(el).color;
              let svgHtml = el.outerHTML.replace(/currentColor/g, color);
              svgHtml = svgHtml.replace(/width="[^"]*"/, 'width="256"').replace(/height="[^"]*"/, 'height="256"');
              page.elements.push({
                tag: 'svg', cls: 'icon', text: '', svg: svgHtml,
                x: r.x, y: r.y, w: r.width, h: r.height,
                fontSize: 16, color: color, bold: false, align: 'left',
                fontFamily: '', lineHeightRatio: null
              });
            });
            out.push(page);
          });
          return out;
        }
        """)
        # SVG 图标 → 透明背景 PNG（playwright 渲染，和 HTML 观感一致）
        ic = 0
        for page in slides_data:
            for el in page["elements"]:
                if el.get("svg"):
                    ic += 1
                    png = f"/tmp/_ic_{ic}.png"
                    pg.set_content(
                        f'<style>html,body{{margin:0;background:transparent}}</style>'
                        f'<body style="margin:0;background:transparent">{el["svg"]}</body>'
                    )
                    pg.wait_for_timeout(60)
                    pg.locator("svg").screenshot(path=png, omit_background=True)
                    el["png"] = png
        b.close()

    def classify(el):
        cls = el["cls"]; tag = el["tag"]
        if "kicker" in cls: return "kicker"
        if tag in ("h1", "h2") or "title" in cls: return "title"
        if "vs-card" in cls: return "vs-card"
        if "card" in cls: return "card"
        if "dot" in cls: return "dot"
        if tag == "h3": return "card-title"
        if tag == "p" or "sub" in cls: return "body"
        if "foot" in cls or "pager" in cls or "hint" in cls: return "footer"
        if tag in ("img", "svg"): return "image"
        if tag == "video": return "video"
        if tag == "a": return "link"
        if tag == "table": return "table"
        return "body"

    def add_placeholder(slide, x, y, w, h, label):
        """文件缺失时画占位框：灰底 + 标签文字，先占好位置。"""
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0x1C, 0x26, 0x3E)
        box.line.color.rgb = RGBColor(0x3A, 0x46, 0x60)
        box.line.width = Pt(1)
        tf = box.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_ = tf.paragraphs[0]
        p_.alignment = PP_ALIGN.CENTER
        r_ = p_.add_run()
        r_.text = label
        r_.font.size = Pt(12)
        r_.font.color.rgb = RGBColor(0x8B, 0x96, 0xAD)
        r_.font.name = "Microsoft YaHei"

    def add_text_runs(para, text, font, fs, bold, color, href=None):
        """把文本写入段落：emoji 单独 run 用 emoji 字体，其余用指定字体（避免乱码）。"""
        if not text:
            run = para.add_run()
            run.text = ""
            run.font.size = Pt(fs)
            run.font.name = font
            return
        pos = 0
        for m in EMOJI_RE.finditer(text):
            if m.start() > pos:
                _run(para, text[pos:m.start()], font, fs, bold, color, href)
            _run(para, m.group(), EMOJI_FONT, fs, bold, color, href)
            pos = m.end()
        if pos < len(text):
            _run(para, text[pos:], font, fs, bold, color, href)

    def _run(para, text, font, fs, bold, color, href=None):
        run = para.add_run()
        run.text = text
        run.font.size = Pt(fs)
        run.font.bold = bold
        run.font.color.rgb = to_rgb(color)
        run.font.name = font
        if href:
            try:
                run.hyperlink.address = href  # PPT 里文字可点击跳转
                if run.font.color is None or str(run.font.color.rgb) == "E8ECF4":
                    run.font.color.rgb = RGBColor(0x4C, 0x8D, 0xFF)  # 链接蓝
            except Exception:
                pass

    def add_runs(para, el, font, fs, default_color):
        """把 el 的文本写入段落（标题用 runs 保留颜色，其余按 emoji 拆 run）。"""
        runs = el.get("runs")
        if runs:
            for r in runs:
                run = para.add_run()
                run.text = r["text"]  # 保留空格（Voyager + 旅行者 之间）
                run.font.size = Pt(fs)
                run.font.bold = el["bold"]
                run.font.color.rgb = to_rgb(r.get("color") or default_color)
                run.font.name = font
        else:
            for line in el["text"].split("\n"):
                add_text_runs(para, line, font, fs, el["bold"], el["color"], el.get("href"))

    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    blank = prs.slide_layouts[6]

    for page in slides_data:
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = RGBColor(0x0E, 0x15, 0x26)
        # HTML 背景（body::before 伪元素/纹理画不出来）→ 用预生成背景图铺底。
        # 背景图 = 第 3 个参数，或缺省 /tmp/bg_grid.png（网格纹理，make_bg.py 生成）。
        # 照片背景：把照片存成该路径（或传第 3 参）即可，任意图片都行。
        if bg_image and os.path.exists(bg_image):
            slide.shapes.add_picture(bg_image, 0, 0, width=SLIDE_W, height=SLIDE_H)

        # 背景块（卡片/步骤卡等视觉方框）：先画在底层，文字叠在上面
        for blk in sorted(page.get("blocks", []), key=lambda b: (b["y"], b["x"])):
            bx, by, bw, bh = px_to_in(blk["x"], blk["y"], blk["w"], blk["h"])
            shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, bh)
            if blk.get("bg"):
                shp.fill.solid()
                shp.fill.fore_color.rgb = to_rgb(blk["bg"])
            else:
                shp.fill.background()
            if blk.get("border"):
                shp.line.color.rgb = to_rgb(blk["border"])
                shp.line.width = Pt(1)
            else:
                shp.line.fill.background()
            if blk.get("radius") and blk["w"] > 0 and blk["h"] > 0:
                adj = min(0.5, (blk["radius"] / min(blk["w"], blk["h"])) * 2)
                shp.adjustments[0] = adj

        els = sorted(page["elements"], key=lambda e: (e["y"], e["x"]))
        for el in els:
            k = classify(el)
            if k == "dot":
                continue
            x, y, w, h = px_to_in(el["x"], el["y"], el["w"], el["h"])


            if k == "image":
                src_img = el.get("png") or el.get("src", "")
                if src_img and os.path.exists(src_img):
                    try:
                        slide.shapes.add_picture(src_img, x, y, width=w, height=h)
                    except Exception as e:
                        print(f"  图片插入失败: {e}")
                        add_placeholder(slide, x, y, w, h, "图片占位")
                else:
                    add_placeholder(slide, x, y, w, h, "图片占位")
                continue
            if k == "table":
                rows = el.get("table") or []
                if rows:
                    nrows = len(rows)
                    ncols = max(len(r) for r in rows)
                    try:
                        gfx = slide.shapes.add_table(nrows, ncols, x, y, w, h)
                        tbl = gfx.table
                        # 禁用默认样式 banding 保留简洁：python-pptx 默认表格有样式，先填内容
                        used = [[False] * ncols for _ in range(nrows)]
                        for ri, row in enumerate(rows):
                            ci = 0
                            for cell in row:
                                while ci < ncols and used[ri][ci]:
                                    ci += 1
                                if ci >= ncols:
                                    break
                                rs = max(1, min(cell.get("rowspan") or 1, nrows - ri))
                                cs = max(1, min(cell.get("colspan") or 1, ncols - ci))
                                c = tbl.cell(ri, ci)
                                c.text = cell.get("text", "")
                                for p_ in c.text_frame.paragraphs:
                                    for r_ in p_.runs:
                                        r_.font.size = Pt(12)
                                        r_.font.color.rgb = RGBColor(0xE8, 0xEC, 0xF4)
                                        r_.font.name = pick_font(el.get("fontFamily")) or "Microsoft YaHei"
                                if rs > 1 or cs > 1:
                                    c.merge(tbl.cell(ri + rs - 1, ci + cs - 1))
                                for rr in range(ri, ri + rs):
                                    for cc in range(ci, ci + cs):
                                        used[rr][cc] = True
                                ci += cs
                    except Exception as e:
                        print(f"  表格插入失败: {e}")
                continue
            if k == "video":
                src_v = el.get("src", "")
                if src_v and os.path.exists(src_v):
                    try:
                        slide.shapes.add_movie(src_v, x, y, width=w, height=h)
                    except Exception as e:
                        print(f"  视频插入失败: {e}")
                        add_placeholder(slide, x, y, w, h, "视频占位")
                else:
                    add_placeholder(slide, x, y, w, h, "视频占位")
                continue

            # 文本元素
            tb = slide.shapes.add_textbox(x, y, w, h)
            tf = tb.text_frame
            tf.word_wrap = True
            # ⚠️ 文本框默认 internal margin(0.1/0.05in) 会让文字相对框偏移 → 清零
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE if k == "card-title" else MSO_ANCHOR.TOP

            fs = max(8, el["fontSize"] * 0.5625)
            if k == "title":
                fs = max(fs, 24)
            elif k == "kicker":
                fs = min(fs, 14)
            elif k == "footer":
                fs = 9

            font = pick_font(el.get("fontFamily"))
            lhr = el.get("lineHeightRatio")

            # 多段落：按 \n 拆（标题 runs 走 add_runs 单段逻辑）
            if k == "title" and el.get("runs"):
                p_ = tf.paragraphs[0]
                if lhr: p_.line_spacing = lhr
                p_.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(el["align"], PP_ALIGN.LEFT)
                add_runs(p_, el, font, fs, el["color"])
            else:
                lines = el["text"].split("\n")
                for li, line in enumerate(lines):
                    p_ = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
                    if lhr: p_.line_spacing = lhr
                    p_.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(el["align"], PP_ALIGN.LEFT)
                    add_text_runs(p_, line, font, fs, el["bold"], el["color"], el.get("href"))

        # 页码
        fb = slide.shapes.add_textbox(Inches(8.6), Inches(5.25), Inches(1.0), Inches(0.3))
        ftf = fb.text_frame
        ftf.margin_left = ftf.margin_right = ftf.margin_top = ftf.margin_bottom = 0
        fp = ftf.paragraphs[0]
        fp.alignment = PP_ALIGN.RIGHT
        fr = fp.add_run(); fr.text = f"{page['index']+1}"
        fr.font.size = Pt(10); fr.font.color.rgb = RGBColor(0x8B, 0x96, 0xAD)

    prs.save(out_path)
    print(f"已生成 {out_path}（{len(slides_data)} 页，原生可编辑文字）")

if __name__ == "__main__":
    main()
