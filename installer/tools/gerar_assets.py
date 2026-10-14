# -*- coding: utf-8 -*-
"""Gera o ícone e as imagens do assistente do instalador a partir do logo OCA.

Uso (opcional - as imagens já estão versionadas em installer/assets):
    python installer/tools/gerar_assets.py

Requer Python 3 + Pillow (pip install pillow). Só é necessário rodar de novo
se o logo em OCA.extension/lib/oca_ds/assets/logo/logo.png mudar.
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, "..", ".."))
LOGO = os.path.join(ROOT, "OCA.extension", "lib", "oca_ds", "assets", "logo", "logo.png")
OUT = os.path.join(ROOT, "installer", "assets")

BRAND = (47, 186, 167)        # #2FBAA7 - cor do logo (lib/oca_ds/tokens.py)
ACCENT = (29, 114, 102)       # #1D7266
WHITE = (255, 255, 255)

FONT_CANDIDATES = [
    "C:/Windows/Fonts/segoeuib.ttf", "C:/Windows/Fonts/arialbd.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]
FONT_REGULAR = [
    "C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/arial.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]


def font(candidates, size):
    for path in candidates:
        if os.path.isfile(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def logo_tinted(color, size):
    """Logo recolorido (mantém o alfa) e redimensionado para caber em size x size."""
    src = Image.open(LOGO).convert("RGBA")
    solid = Image.new("RGBA", src.size, color + (255,))
    solid.putalpha(src.getchannel("A"))
    solid.thumbnail((size, size), Image.LANCZOS)
    return solid


def vertical_gradient(w, h, top, bottom):
    img = Image.new("RGB", (w, h), top)
    draw = ImageDraw.Draw(img)
    for y in range(h):
        t = y / float(max(h - 1, 1))
        c = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
        draw.line([(0, y), (w, y)], fill=c)
    return img


def wizard_large(scale):
    w, h = int(164 * scale), int(314 * scale)
    img = vertical_gradient(w, h, BRAND, ACCENT)
    logo = logo_tinted(WHITE, int(96 * scale))
    img.paste(logo, ((w - logo.width) // 2, int(64 * scale)), logo)
    draw = ImageDraw.Draw(img)
    f1 = font(FONT_CANDIDATES, int(30 * scale))
    f2 = font(FONT_REGULAR, int(11 * scale))
    for text, f, y in ((u"OCA", f1, 178), (u"Automatizações", f2, 218), (u"em BIM", f2, 234),
                       (u"Extensão pyRevit", f2, 280)):
        tw = draw.textlength(text, font=f)
        draw.text(((w - tw) / 2, int(y * scale)), text, font=f, fill=WHITE)
    img.save(os.path.join(OUT, "wizard_large_{}.bmp".format(int(scale * 100))))


def wizard_small(scale):
    w, h = int(55 * scale), int(58 * scale)
    img = Image.new("RGB", (w, h), WHITE)
    logo = logo_tinted(BRAND, int(48 * scale))
    img.paste(logo, ((w - logo.width) // 2, (h - logo.height) // 2), logo)
    img.save(os.path.join(OUT, "wizard_small_{}.bmp".format(int(scale * 100))))


def icon():
    sizes = [16, 24, 32, 48, 64, 128, 256]
    base = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
    logo = logo_tinted(BRAND, 240)
    base.paste(logo, ((256 - logo.width) // 2, (256 - logo.height) // 2), logo)
    base.save(os.path.join(OUT, "oca_setup.ico"), sizes=[(s, s) for s in sizes])


def main():
    if not os.path.isfile(LOGO):
        sys.exit("Logo não encontrado: " + LOGO)
    if not os.path.isdir(OUT):
        os.makedirs(OUT)
    for s in (1.0, 1.5, 2.0):
        wizard_large(s)
        wizard_small(s)
    icon()
    print("Imagens geradas em " + OUT)


if __name__ == "__main__":
    main()
