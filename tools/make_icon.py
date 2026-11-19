# -*- coding: utf-8 -*-
"""生成应用图标 assets/app.ico（外加一张 512px 的 app.png 预览图）。

设计口径：
  · 圆角方块 + 对角蓝色渐变，跟界面主色同源
  · 左上角一层柔光，避免纯平色看起来发闷
  · 图形只用白色单色，笔画粗、圆头圆角 —— 16×16 下依然能认出来

所有尺寸都从 S=2048 的超采样母图缩出来，保证边缘干净。

用法（只在改图标时需要，需要额外的 Pillow）：
    pip install pillow
    py -3.14 tools/make_icon.py
"""
import os

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageOps

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")

S = 2048                     # 超采样母图边长
RADIUS = 0.225               # 圆角占边长的比例（Windows 11 风格偏方的 squircle）
GRAD_FROM = (76, 141, 255)   # #4C8DFF 左上
GRAD_TO = (31, 79, 224)      # #1F4FE0 右下
ICO_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)
PNG_SIZE = 512


# ---------------------------------------------------------------- 基础元件
def squircle_mask(size, radius_frac=RADIUS):
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle((0, 0, size - 1, size - 1), radius=int(size * radius_frac), fill=255)
    return m


def diagonal_gradient(size, c_from, c_to):
    """左上 → 右下的线性渐变（把竖直渐变转 45° 后裁中间）"""
    g = Image.linear_gradient("L")                      # 256x256，上黑下白
    g = g.rotate(45, resample=Image.BICUBIC, expand=True)
    w, h = g.size
    g = g.crop(((w - 256) // 2, (h - 256) // 2,
                (w - 256) // 2 + 256, (h - 256) // 2 + 256))
    g = g.resize((size, size), Image.BICUBIC)
    return Image.composite(Image.new("RGB", (size, size), c_to),
                           Image.new("RGB", (size, size), c_from), g)


def tile(size=S):
    """渐变圆角方块 + 左上柔光"""
    mask = squircle_mask(size)
    base = diagonal_gradient(size, GRAD_FROM, GRAD_TO).convert("RGBA")
    base.putalpha(mask)

    # 左上柔光：radial 渐变裁到方块范围内，再压低透明度
    rg = ImageOps.invert(Image.radial_gradient("L"))     # 中心 255
    rg = rg.resize((size * 2, size * 2), Image.BICUBIC)
    cx, cy = int(size * 0.20), int(size * 0.14)
    box = (cx - size // 2, cy - size // 2, cx + size // 2, cy + size // 2)
    glow = rg.crop(box).point(lambda v: int(v * 0.30))
    glow = ImageChops.multiply(glow, mask)
    white = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    white.putalpha(glow)
    return Image.alpha_composite(base, white)


def stroke(draw, pts, width, fill):
    """圆头折线：中间段用 joint=curve，两端补圆点，否则端头是平的"""
    draw.line(pts, fill=fill, width=int(width), joint="curve")
    r = width / 2.0
    for x, y in (pts[0], pts[-1]):
        draw.ellipse((x - r, y - r, x + r, y + r), fill=fill)


def with_shadow(layer, blur, offset, alpha):
    """给白色图形垫一层柔和投影，避免压在蓝底上发飘"""
    a = layer.split()[3].point(lambda v: int(v * alpha))
    sh = Image.new("RGBA", layer.size, (0, 0, 0, 0))
    sh.putalpha(a)
    sh = sh.filter(ImageFilter.GaussianBlur(blur))
    sh = ImageChops.offset(sh, int(offset[0]), int(offset[1]))
    return Image.alpha_composite(sh, layer)


# ---------------------------------------------------------------- 图形
def build(size=S):
    """渐变圆角方块 + 白色圆头对勾"""
    img = tile(size)
    layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    pts = [(size * 0.238, size * 0.506),
           (size * 0.436, size * 0.694),
           (size * 0.768, size * 0.306)]
    stroke(d, pts, size * 0.115, (255, 255, 255, 255))
    return Image.alpha_composite(img, with_shadow(layer, size * 0.022,
                                                  (0, size * 0.015), 0.30))


# ---------------------------------------------------------------- 输出
def write_ico(img, path):
    # PIL 的 ICO 保存器自己按 sizes 从母图缩出每一档，不要再传 append_images（会被忽略）
    img.save(path, format="ICO", sizes=[(n, n) for n in ICO_SIZES])
    with Image.open(path) as chk:
        return sorted(w for w, _ in chk.info.get("sizes", []))


def main():
    os.makedirs(ASSETS, exist_ok=True)
    master = build()

    ico = os.path.join(ASSETS, "app.ico")
    sizes = write_ico(master, ico)

    png = os.path.join(ASSETS, "app.png")
    master.resize((PNG_SIZE, PNG_SIZE), Image.LANCZOS).save(png)

    print(f"已写入 {ico}  尺寸: {sizes}")
    print(f"已写入 {png}  {PNG_SIZE}x{PNG_SIZE}")


if __name__ == "__main__":
    main()
