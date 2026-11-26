"""Draw the app icon (a page with a 'W') and save a multi-size .ico."""

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def draw(size: int = 256) -> Image.Image:
    s = size / 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # Page with folded corner
    d.rounded_rectangle([40 * s, 16 * s, 216 * s, 240 * s], radius=18 * s, fill="#ffffff",
                        outline="#c9d1d9", width=max(1, int(4 * s)))
    d.polygon([(168 * s, 16 * s), (216 * s, 64 * s), (168 * s, 64 * s)], fill="#d0d7de")
    # Blue badge with W
    d.rounded_rectangle([16 * s, 96 * s, 176 * s, 224 * s], radius=20 * s, fill="#185abd")
    try:
        font = ImageFont.load_default(size=int(104 * s))
    except TypeError:  # very old Pillow
        font = ImageFont.load_default()
    d.text((96 * s, 160 * s), "W", fill="white", font=font, anchor="mm")
    return img


def main(out: str) -> None:
    draw().save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"wrote {Path(out).resolve()}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "packaging/doc2word.ico")
