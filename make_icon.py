"""Génère icon.ico et icon.png (logo Morphe) avec Pillow."""
from PIL import Image, ImageDraw, ImageFilter

S = 1024  # rendu en grand puis réduit pour des bords lisses


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def build():
    c1, c2 = (124, 58, 237), (6, 182, 212)  # violet -> cyan
    # fond dégradé diagonal
    grad = Image.new("RGB", (S, S))
    px = grad.load()
    for y in range(S):
        for x in range(S):
            px[x, y] = lerp(c1, c2, (x + y) / (2 * S))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * 0.23), fill=255)
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    img.paste(grad, (0, 0), mask)

    # reflet doux en haut
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse((-S * 0.2, -S * 0.7, S * 1.2, S * 0.45), fill=(255, 255, 255, 38))
    img = Image.alpha_composite(img, Image.composite(glow, Image.new("RGBA", (S, S), (0, 0, 0, 0)), mask))

    # glyphe "M" : un trait continu aux jointures arrondies
    layer = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    pts = [(0.25, 0.72), (0.25, 0.30), (0.50, 0.58), (0.75, 0.30), (0.75, 0.72)]
    pts = [(x * S, y * S) for x, y in pts]
    w = int(S * 0.105)
    d.line(pts, fill=(255, 255, 255, 255), width=w, joint="curve")
    for x, y in (pts[0], pts[-1]):
        d.ellipse((x - w / 2, y - w / 2, x + w / 2, y + w / 2), fill=(255, 255, 255, 255))
    # petite étincelle (le "changement de forme")
    cx, cy, r = 0.50 * S, 0.30 * S, S * 0.045
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(255, 255, 255, 255))

    shadow = layer.filter(ImageFilter.GaussianBlur(S * 0.02))
    shadow = Image.eval(shadow, lambda v: int(v * 0.35))
    sh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    sh.paste((20, 10, 60, 255), (0, int(S * 0.012)), shadow.split()[3])
    img = Image.alpha_composite(img, sh)
    img = Image.alpha_composite(img, layer)
    return img


if __name__ == "__main__":
    big = build()
    big.resize((512, 512), Image.LANCZOS).save("icon.png")
    big.save("icon.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("icon.ico + icon.png créés")
