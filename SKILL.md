---
name: ai-ppt
description: AI 做好看的 PPT/幻灯片。触发：主人说「做 PPT」「HTML 幻灯片」。含工具选型、审美、模板。
---

# AI 做好看的 PPT

## 核心信条
- **工具决定下限，审美决定上限**。AI 负责排版体力活，主人负责逻辑和审美。
- 好看 = 做减法（低饱和、留白、字少、克制）。
- 交付形态三种：①在线 AI 工具（快）②HTML 幻灯片（视觉可控、技术流，主人最爱）③python-pptx（批量数据汇报）。

## 触发场景
- 主人要做 PPT/演示/汇报/课堂展示
- 主人研究 AI PPT 工具
- 主人要「好看」的幻灯片，且不排斥技术流

## 一、工具选型速查（2026-09 实测）
| 需求 | 首选 | 备选 |
|---|---|---|
| 视觉惊艳/对外 | Gamma（视觉天花板） | Beautiful AI（排版稳） |
| 数据 100% 保真 | 即触AI PPT（唯一保真模式） | — |
| 用自己的模板 | 即触AI PPT（免费8个） | aippt（免费1个） |
| 长文/论文 | Kimi PPT | 天工 AI |
| 语音输入 | 讯飞智文 | — |
| 已有 WPS 生态 | WPS AI / Copilot（插件型） | — |
| 设计资源多 | Canva AI | — |

实测坑：
- Gamma 导出 PPTX 会文本框偏移/丢字体 → 网页端演示
- 大多数工具会「优化改写」数据（15.6%→近16%）→ 重要数据逐页核对
- aippt 免费版上传文档要会员
- 生成速度：即触 12-15s 最快，Kimi/Canva 25-40s 最慢

## 二、五步工作流
1. **喂素材**：大纲（最好）/ 主题关键词 / 文档 URL。大纲阶段就把对比、步骤、因果写清楚 → AI 自动转表格/流程图
2. **设参数**（决定质量的一步）：
   - 页数略少于大纲节点数（防注水）
   - 受众写具体：「给非技术背景的 CTO 汇报，突出资源投入和风险评估」＞「给领导」
   - 语气匹配：正式商务=数据图表；教育培训=解释拆解
3. **生成**：AI 拆解→填充→视觉匹配→配图→备注
4. **审查**：核对数字、换关键页配图、自然语言局部改（「第三页图表加数据标签」）
5. **导出**：PPTX/PDF/图片/在线链接，导出后翻一遍查字体丢失偏移

## 三、雅 vs 俗 6 原则（优设）
1. 新颖 > 大众：避免红配黄、蓝配白、描边字、金属字、默认模板
2. 低调 > 高调：字号不必大，信息有主次；「不喧哗，自有声」
3. 简约 > 繁杂：元素少、颜色少（黑白灰+1强调色）、字体≤2种、留白1/3
4. 抽象 > 具象：抽象几何/图标 > 直白大图
5. 有细节 > 没细节：对齐、间距、图标风格统一
6. 弱商业 > 强商业：别堆卖点，克制高级

## 四、提示词模板（审美写进 prompt）
```
请生成 N 页 PPT，主题「X」，受众 [具体描述]。
设计风格：简约现代，低饱和配色（主色 + 1 个强调色），
字体不超过两种，每页留白充足，标题层级清晰，
多用抽象几何图形与图标代替大图，图表用统一配色。
不要使用：高饱和渐变色块、描边立体字、剪贴画风格。
```
**反向排除比正向要求更管用**——明确告诉 AI「不要什么」。

## 五、手写 HTML 幻灯片（技术流，首选交付方式）
模板：`templates/presentation.html`（已验证，深蓝+橙 7 页示例）

做法：
1. 复制模板，改内容（每页一个 `<section class="slide">`）
2. 设计基调：深色底 `#0e1526` + 强调色 `#f5a623`，低饱和、留白多、卡片圆角+细边框
3. 翻页交互已内嵌（空格/方向键/点击 + 进度条 + 页码），零依赖单文件
4. 图标用 **Tabler 内联 SVG**（库在 `~/tabler-icons/icons/outline/`，5130 个，见下方「Tabler 图标库」）；要动效可换 Morphicons（做前端时提醒主人）

**验证流程（必须做）**：服务器无 GUI 浏览器，用 scrapling venv 的 playwright 渲染截图验证：
```bash
~/scrapling-venv/bin/python -c "from playwright.sync_api import sync_playwright; p=sync_playwright().start(); b=p.chromium.launch(); pg=b.new_page(); pg.goto('file:///path/to.html'); pg.wait_for_timeout(800); pg.screenshot(path='/tmp/slide.png'); b.close()"
```
再用 vision_analyze 看截图确认排版/配色/可读性，有问题就改。

## 六、交付：HTML 最后怎么交
HTML 不能直接交作业？三条路，按老师/场合选：

1. **HTML → PDF（最常用）**：浏览器打开 → Ctrl+P → 另存 PDF。排版 100% 保真，PDF 几乎哪都能交。
2. **HTML → 可编辑 PPTX（推荐，文字可改）**：`scripts/html2editable_pptx.py`——playwright 读 DOM 语义结构（tag/class/文本/位置/样式）→ python-pptx 重建原生 TextFrame（标题/正文/卡片/列表/图片全可编辑）。验证过：改字读回成功。**这是主人的思路：HTML 用语义标签（h1=标题、div+class=卡片/列表/正文）标注，转换器按标签生成原生元素。**
3. **HTML → 图片版 PPTX（保真但文字死）**：`scripts/html2pptx.py`（逐页截图合成）。排版 100% 保真但不可编辑——仅当不需要改字时用。
4. **HTML → 可编辑 PPTX（样式打折）**：Marp CLI（`marp --pptx`）或 pandoc 转。

**最佳组合**：现场演示/自己讲 → HTML（浏览器全屏最帅）；交作业 → 可编辑版 PPTX（html2editable_pptx.py）；急用不改字 → 截图版。

### 语义标注约定（html2editable_pptx.py 的映射规则）
| HTML 元素/class | PPT 里变成 |
|---|---|
| `.slide` | 一页（深色背景） |
| `.kicker` | 页眉小字 |
| `h1 / h2 / .title` | 页标题（拆 runs 保留 .hl 高亮色） |
| `.sub / p` | 正文文本框 |
| 任意有背景/边框的 div（.card/.step/.vs-card…） | **自动检测** → 圆角矩形底框（不枚举 class） |
| `h3` | 卡片标题 |
| `.foot / .pager / .hint` | 页脚 |
| `img` | 图片（文件缺失 → 灰色「图片占位」框） |
| `video` | 视频嵌入（add_movie，文件缺失 → 「视频占位」框） |
| `a href` | 文本超链接（run.hyperlink，蓝色可点击，文字可编辑） |
| `<svg>`（Tabler 图标） | SVG → 透明 PNG 图片插入 |

### 转换器要点（html2editable_pptx.py，2026-09-18 实测踩坑全记录）
1. **坐标坑（标题偏移根源）**：非活跃 `.slide` 有 `transform: translateY(24px)`，getBoundingClientRect 返回带偏移的坐标 → 抓坐标前必须先 `document.querySelectorAll('.slide').forEach(s => s.classList.add('active'))` + 等 700ms 过渡结束。
2. **字体**：pick_font() 从 HTML font-family 栈**优先选 Microsoft YaHei**（Windows/WPS 必有，与浏览器在目标机渲染 HTML 一致），回退 PingFang SC / Noto Sans SC。本机（服务器）无雅黑会替换成 Noto，但目标机器是 Windows WPS 就没问题。
3. **文本框默认内边距**（约 0.1/0.05in）会让文字相对框偏移 → margins 全部置 0。
4. **auto_size = SHAPE_TO_FIT_TEXT**：文字多时框自动撑高，不溢出错位。
5. **行距**：p.line_spacing = HTML line-height ratio（1.7/1.6），多行文字间距一致。
6. **符号处理**：
   - 箭头 → / 间隔号 · / 破折号 —：**留在中文字体 run**，正常显示；⚠️ 不要拆到 emoji 字体（Segoe UI Emoji 无箭头字形，拆了变方块）
   - 真 emoji（U+1F000-U+1FAFF）：单独拆 run 用 Segoe UI Emoji（Windows/WPS 彩色渲染）
7. **半透明背景**：rgba 色与深色底混合（blend）再画，避免透明色盖住背景。
8. **嵌套文本**（.txt 含 `<small>`）：nodeText 递归拼接，block 子元素前后换行，避免父子框重叠。
9. **容器去重**：CONTAINER 黑名单（.hl/.foot/.avatar/.card/.step/.route…）+ 吸收父元素逻辑，避免标题/容器重复文本。
10. **背景块自动检测**：computed style 有背景色或边框的块级元素（尺寸 >20px）→ 画 ROUNDED_RECTANGLE（背景色填充 + 边框 + 圆角 adjustment=radius/min(w,h)*2），文字叠在框上。这就是「HTML 里的小方框」在 PPT 里还原的原理，不用枚举 class。
11. **背景图铺底（HTML 背景还原）**：HTML 背景 = body 纯色 + body::before 伪元素的纹理（网格/渐变）——伪元素不在 DOM 里抓不到。解法：预生成一张背景图，每页 slide 创建后先 `add_picture(背景图, 0,0, 全尺寸)` 铺最底层，其他元素叠上层。
    - 网格纹理：`scripts/make_bg.py` 生成（默认深蓝 #0e1526 + 3% 白线 + 48px 间距，可调色/透明度/间距）→ /tmp/bg_grid.png
    - 照片背景：任意图片，传转换脚本第 3 参 `html2editable_pptx.py input.html out.pptx /path/photo.jpg`，或缺省用 /tmp/bg_grid.png
    - 3% 白网格其实极淡，纯色背景也可以——按主人喜好选
12. **图片/视频/超链接**（2026-09-18 实测）：
    - `<img>`：无文本元素主循环抓不到 → 单独 `querySelectorAll('img')` 抓 src（file:// 转本地路径）；文件存在 add_picture，缺失画占位框（灰底 RECTANGLE + 「图片占位」文字，位置先占好手动替换）
    - `<video>`：同样单独抓，add_movie 嵌入（python-pptx 支持）；缺失画「视频占位」框
    - `<a href>`：⚠️ 内联 a 会被父元素吸收丢 href——nodeText 对 A 返回空（不吸收），a 单独抓（自带 rect 精确位置 + href）；创建时 run.hyperlink.address = href（链接蓝 #4C8DFF）。父文本框留空 a 位置，a 独立框覆盖显示，位置精确
13. **原生表格**（借鉴 joker-duzhong/html-to-pptx，2026-09-18）：
    - `<table>` 单独抓（行列 + rowspan/colspan）→ `add_table` 原生 PPT 表格，合并单元格用 `cell.merge()`（先建 used 网格跳过被合并区）
    - ⚠️ 表格的 th/td 有边框会被「背景块检测」误判成卡片 → blocks 检测必须 `closest('table')` 跳过
    - 单元格文本默认 12pt 浅色雅黑
14. **忽略属性**：HTML 元素加 `data-html2pptx-ignore` → 转换时跳过（主循环/背景块/a/img/video/table 全部检查）

### 借鉴参考（2026-09-18 调研）
- **微软 html2pptx**（researchstudio，2.8k stars）：同路线标杆。流程 = playwright 抓 DOM → python-pptx 原生形状 → **soffice 渲染 PPTX → 和 HTML 截图对比** → **AI 视觉自动质检**（audit.json，可聚合找系统性 bug）。字体缺失→字宽变化→换行→级联偏移是他们 GOTCHAS 第一条（与我们踩的坑一致）；PPT-safe 字体（Arial/Calibri/雅黑）最稳，自定义字体需 embed。
- **joker-duzhong/html-to-pptx**（前端库）：原生表格 ✓ 已借鉴；原生图表（data-pptx-chart-config）、CSS 动画→PPT 动画、Canvas 处理，未做（低优先级）。
- **dom-to-pptx**（2 星桌面应用）：pixel accuracy，参考价值低。
- 未来可做：soffice 自动渲染验证（装 LibreOffice）+ vision 对比质检脚本（可用 PIL 像素 diff + vision_analyze）。

### Tabler 图标库（HTML 与 PPT 统一图标方案，主人定）
- 库位置：`~/tabler-icons/icons/outline/`（5130 个 SVG，MIT）
- **HTML 用法**：图标 SVG 内联进 HTML——去注释头、width/height 设目标尺寸（如 28）、stroke 保持 currentColor 由 CSS 控制颜色
- **转换**：脚本识别 `<svg>` → currentColor 替换为实际计算色 → playwright 渲染成透明 PNG（`omit_background=True` + transparent CSS）→ add_picture 插入
- **为什么不用 emoji 当图标**：各平台渲染不一（Windows/Mac/Linux/WPS 各画各的、缺字乱码）；图标库矢量统一、可换色、可缩放、有设计感
- 查图标名：`ls ~/tabler-icons/icons/outline | grep 关键词`（自行车=bike 不是 bicycle）
- python-pptx 1.0.2 不支持直接 add_picture SVG（Pillow 不认），必须转 PNG

坑：容器（.card/.foot）自身 textContent 含子元素文本 → 只画形状跳过文本，否则文字重复（已由 CONTAINER 黑名单 + blocks 自动检测解决）。

## 八、独立 CLI 工具（不依赖 agent）
`html2editable_pptx.py` 已独立成命令 **`~/bin/html2pptx`**（shebang 指向 scrapling-venv）：
```
html2pptx input.html [output.pptx] [背景图]
```
- output 缺省 = input_editable.pptx；背景图可选（照片/网格，缺省 /tmp/bg_grid.png）
- 依赖：~/scrapling-venv（playwright + python-pptx + PIL）；HTML 图标用 Tabler 内联 SVG
- 脚本更新后同步三处：~/bin/html2pptx、~/html2editable_pptx.py、skill scripts/

## 九、避坑清单
- [ ] 重要数据逐页核对（AI 偷偷改数字）
- [ ] 关键配图手动换
- [ ] Gamma/Tome 导出 PPTX 检查偏移丢字体
- [ ] 机密数据先脱敏再传在线工具
- [ ] AI 是效率倍增器，不是决策者——人工把关不能跳

## 验收标准
- [ ] 按主人场景推荐了正确工具/路线
- [ ] HTML 幻灯片：渲染截图验证过、翻页正常、可读性 OK
- [ ] 内容遵循雅 vs 俗原则（低饱和、留白、字少）
