#!/usr/bin/env python3
import json
import sys
from pathlib import Path

from PIL import Image, ImageOps


def bits_to_hex(bits):
    value = 0
    for bit in bits:
        value = (value << 1) | int(bit)
    width = max(1, (len(bits) + 3) // 4)
    return f"{value:0{width}x}"


def average_hash(image):
    resized = ImageOps.grayscale(image).resize((8, 8), Image.Resampling.LANCZOS)
    values = list(resized.getdata())
    avg = sum(values) / len(values)
    return bits_to_hex([value >= avg for value in values])


def difference_hash(image):
    resized = ImageOps.grayscale(image).resize((9, 8), Image.Resampling.LANCZOS)
    rows = [list(resized.crop((0, y, 9, y + 1)).getdata()) for y in range(8)]
    bits = []
    for row in rows:
      for x in range(8):
        bits.append(row[x] > row[x + 1])
    return bits_to_hex(bits)


def fingerprint(path):
    with Image.open(path) as image:
        image = ImageOps.exif_transpose(image).convert("RGB")
        return {
            "path": str(path),
            "width": image.width,
            "height": image.height,
            "ahash": average_hash(image),
            "dhash": difference_hash(image),
        }


def main():
    result = {}
    for raw in sys.argv[1:]:
        path = Path(raw)
        try:
            result[str(path)] = fingerprint(path)
        except Exception as exc:
            result[str(path)] = {"path": str(path), "error": str(exc)}
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
