#!/usr/bin/env python3
from __future__ import annotations

import sys
from pathlib import Path

from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


def main() -> int:
    if len(sys.argv) < 3:
        raise SystemExit("Usage: images_to_pdf.py <output.pdf> <slide-01.png> ...")
    output = Path(sys.argv[1])
    images = [Path(value) for value in sys.argv[2:]]
    page = (13.333333 * 72, 7.5 * 72)
    document = canvas.Canvas(str(output), pagesize=page, pageCompression=1)
    for image in images:
        document.drawImage(
            ImageReader(str(image)),
            0,
            0,
            width=page[0],
            height=page[1],
            preserveAspectRatio=True,
            mask="auto",
        )
        document.showPage()
    document.save()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
