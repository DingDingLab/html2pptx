# html2pptx — HTML 幻灯片 → 可编辑 PowerPoint

把**语义化 HTML 幻灯片**一键转换成**原生可编辑的 PPTX**——不是截图！所有文字都是 PPT 里的真文本框，能在 WPS / PowerPoint 里直接改字、加批注。

灵感来自主人的思路：HTML 用语义标签（h1=标题、div+class=卡片/列表）标注内容，转换器按标签生成 PPT 原生元素，排版保真 + 文字可编辑。

## ✨ 特性

- **文字可编辑**：h1/h2/kicker/正文/卡片全部是原生 TextFrame
- **卡片自动识别**：任何有背景/边框的 div → 圆角矩形底框（不用枚举 class）
- **Tabler 矢量图标**：HTML 内联 SVG → 透明 PNG 插入，HTML 与 PPT 观感统一
- **背景图铺底**：网格纹理/照片背景一键铺底
- **图片 / 视频**：文件缺失自动画「占位框」先占位置
- **超链接**：`<a href>` → PPT 可点击链接
- **原生表格**：`<table>` → PPT 原生表格（支持 rowspan/colspan 合并）
- **符号不乱码**：箭头/间隔号留在中文字体，emoji 单独拆 run 用 emoji 字体

## 🚀 快速开始

### 依赖

```bash
pip install playwright python-pptx Pillow
playwright install chromium
```

### 用法

```bash
python scripts/html2editable_pptx.py input.html output.pptx [背景图]
# 缺省背景图 = /tmp/bg_grid.png（可用 scripts/make_bg.py 生成）
python scripts/make_bg.py                  # 生成深蓝网格背景
python scripts/html2editable_pptx.py 我的幻灯片.html 成品.pptx /path/photo.jpg
```

## 📐 HTML 约定

每页一个 `<section class="slide">`，语义标注：

| HTML | PPT 里变成 |
|---|---|
| `.slide` | 一页 |
| `.kicker` | 页眉小字 |
| `h1 / h2` | 页标题 |
| `.sub / p` | 正文文本框 |
| 有背景/边框的 div | 圆角矩形卡片 |
| `img` / `video` | 图片 / 视频（缺失→占位框） |
| `a href` | 可点击超链接 |
| `table` | 原生表格（含合并） |
| `<svg>`（Tabler 图标） | 透明 PNG 图标 |

参考模板：`templates/presentation.html`（深蓝+橙 7 页示例）。

## 🎨 图标：下载第三方符号，保持 HTML 与 PPT 一致

**核心方法：下载开源图标库（Tabler）→ 内联 SVG 进 HTML → 转换时自动识别 SVG → PPT 里符号一致。**

### 获取图标（两种方式）

**方式 A：在线复制（单个图标）**
1. 打开 https://tabler.io/icons 搜索图标名（如 `compass`、`bike`）
2. 点击图标 → 复制 SVG 代码
3. 内联进 HTML（去注释头、设 width/height 如 28、stroke 保持 `currentColor` 由 CSS 控制颜色）

**方式 B：本地图标库（批量/离线，推荐）**
```bash
git clone --depth 1 --filter=blob:none --sparse https://github.com/tabler/tabler-icons.git
cd tabler-icons
git sparse-checkout set icons/outline    # 只拉 5130 个 SVG，几 MB
ls icons/outline | grep 关键词           # 查图标名（自行车=bike 不是 bicycle）
```

### HTML 里放图标

```html
<!-- 内联 SVG，stroke 用 currentColor，颜色跟随 CSS -->
<div class="ic"><svg xmlns="http://www.w3.org/2000/svg" width="28" height="28"
  viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
  stroke-linecap="round" stroke-linejoin="round"><path d="..."/></svg></div>
```

转换时脚本识别 `<svg>` → 透明 PNG 插入 PPT，HTML 与 PPT 用同一套图标。**Tabler Icons 为 MIT 许可**（免费商用、可再分发，注明来源即可）。本项目不内置图标本体，图标由使用方从 Tabler 获取。

忽略元素：加 `data-html2pptx-ignore` 属性，转换时跳过。

## 📁 目录

```
├── scripts/
│   ├── html2editable_pptx.py   # 主转换器（可编辑版，推荐）
│   ├── html2pptx.py            # 截图版（100% 保真但文字不可编辑）
│   └── make_bg.py              # 网格背景生成器
├── templates/
│   └── presentation.html       # 幻灯片模板
├── SKILL.md                    # Agent skill（给 AI 助手用的操作手册）
└── README.md
```

## ⚠️ 已知坑（都修过，写在 SKILL.md 里）

- 非活跃 `.slide` 的 `translateY(24px)` 过渡会让坐标集体偏移 → 抓坐标前先全部 active
- 字体缺失→字宽变化→换行→位置级联偏移 → 优先用 PPT 预装字体（微软雅黑/Noto）
- 文本框默认内边距会偏位 → 清零
- 内联 `<a>` 会被父元素吸收丢 href → 单独抓取

## 📄 License

MIT © Nadin Chen (DingDingLab)
