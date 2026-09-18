#!/usr/bin/env python
"""生成 HTML 幻灯片同款网格背景图 → /tmp/bg_grid.png（html2editable_pptx.py 自动铺底）。

用法:
  python make_bg.py [背景色] [线透明度] [网格间距px] [输出]

默认生成 Voyager 模板同款：深蓝 #0e1526 + 3% 白线 + 48px 间距。

照片背景不靠这个：直接传照片路径给转换脚本第 3 参即可。
依赖: Pillow
"""
import sys
from PIL import Image, ImageDraw

def hex2rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def main():
    bg_hex = sys.argv[1] if len(sys.argv) > 1 else "#0e1526"
    alpha = float(sys.argv[2]) if len(sys.argv) > 2 else 0.03
    spacing = int(sys.argv[3]) if len(sys.argv) > 3 else 48
    out = sys.argv[4] if len(sys.argv) > 4 else "/tmp/bg_grid.png"

    W, H = 1280, 720
    bg = hex2rgb(bg_hex)
    # 半透明白线混到背景色
    line = tuple(int(c * (1 - alpha) + 255 * alpha) for c in bg)

    img = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(img)
    for x in range(0, W, spacing):
        d.line([(x, 0), (x, H)], fill=line, width=1)
    for y in range(0, H, spacing):
        d.line([(0, y), (W, y)], fill=line, width=1)
    img.save(out)
    print(f"✅ 网格背景图已生成 {out}（{bg_hex} + {int(alpha*100)}%白线 {spacing}px）")

if __name__ == "__main__":
    main()
