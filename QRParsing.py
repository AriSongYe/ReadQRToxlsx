"""Validated scanner input and optional JPEG/PNG QR decoding."""
from dataclasses import dataclass
import re

CATEGORIES = ("일반", "신학생", "선교사", "개척교회", "북한")


@dataclass(frozen=True)
class QRData:
    category: int
    identifier: str

    @property
    def text(self):
        return f"{self.category}-{self.identifier}"


def parse_qr(value):
    match = re.fullmatch(r"([1-5])\s*-\s*([0-9]+)", str(value).strip())
    if not match:
        raise ValueError("QR 형식은 항목번호(1~5)-ID입니다. 예: 1-2386")
    return QRData(int(match[1]), match[2])


def decode_image(filepath):
    import zxingcpp
    from PIL import Image, ImageOps

    with Image.open(filepath) as source:
        image = ImageOps.exif_transpose(source).convert("RGB")
        results = zxingcpp.read_barcodes(image, formats=zxingcpp.BarcodeFormat.QRCode)
    codes = []
    for result in results:
        try:
            code = parse_qr(result.text).text
        except ValueError:
            continue
        if code not in codes:
            codes.append(code)
    if not codes:
        raise ValueError("사진에서 지원하는 QR을 찾지 못했습니다. QR이 선명한 사진을 선택해 주세요.")
    return codes
