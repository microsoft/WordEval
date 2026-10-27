"""families / charts / rendered_image with explicit evidence dependencies."""

from typing import Dict, List, Any
import struct
import zlib


def png_nonwhite_ratio(path: str) -> Dict[str, Any]:
    """Measure a Word-exported PNG using only the standard library."""
    with open(path, "rb") as stream:
        payload = stream.read()
    if not payload.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError("chart export is not a PNG")
    position = 8
    idat: List[bytes] = []
    palette = None
    width = height = bit_depth = color_type = interlace = None
    while position < len(payload):
        length = struct.unpack(">I", payload[position : position + 4])[0]
        chunk_type = payload[position + 4 : position + 8]
        chunk = payload[position + 8 : position + 8 + length]
        position += 12 + length
        if chunk_type == b"IHDR":
            (
                width,
                height,
                bit_depth,
                color_type,
                _compression,
                _filter,
                interlace,
            ) = struct.unpack(">IIBBBBB", chunk)
        elif chunk_type == b"PLTE":
            palette = [tuple(chunk[index : index + 3]) for index in range(0, len(chunk), 3)]
        elif chunk_type == b"IDAT":
            idat.append(chunk)
    if not width or not height or bit_depth != 8 or interlace != 0:
        raise ValueError("unsupported chart PNG format")
    bytes_per_pixel = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}.get(color_type)
    if bytes_per_pixel is None or (color_type == 3 and not palette):
        raise ValueError("unsupported chart PNG color type")

    def paeth(left: int, above: int, upper_left: int) -> int:
        estimate = left + above - upper_left
        distances = (
            abs(estimate - left),
            abs(estimate - above),
            abs(estimate - upper_left),
        )
        if distances[0] <= distances[1] and distances[0] <= distances[2]:
            return left
        return above if distances[1] <= distances[2] else upper_left

    raw = zlib.decompress(b"".join(idat))
    stride = width * bytes_per_pixel
    previous = bytearray(stride)
    offset = 0
    nonwhite = 0
    for _row in range(height):
        filter_type = raw[offset]
        offset += 1
        encoded = raw[offset : offset + stride]
        offset += stride
        decoded = bytearray(stride)
        for index, value in enumerate(encoded):
            left = decoded[index - bytes_per_pixel] if index >= bytes_per_pixel else 0
            above = previous[index]
            upper_left = previous[index - bytes_per_pixel] if index >= bytes_per_pixel else 0
            predictor = {
                0: 0,
                1: left,
                2: above,
                3: (left + above) // 2,
                4: paeth(left, above, upper_left),
            }.get(filter_type)
            if predictor is None:
                raise ValueError("unsupported chart PNG filter")
            decoded[index] = (value + predictor) & 0xFF
        for index in range(0, stride, bytes_per_pixel):
            if color_type == 6:
                red, green, blue, alpha = decoded[index : index + 4]
            elif color_type == 2:
                red, green, blue = decoded[index : index + 3]
                alpha = 255
            elif color_type == 3:
                red, green, blue = palette[decoded[index]]
                alpha = 255
            elif color_type == 0:
                red = green = blue = decoded[index]
                alpha = 255
            else:
                red = green = blue = decoded[index]
                alpha = decoded[index + 1]
            if alpha > 10 and min(red, green, blue) < 245:
                nonwhite += 1
        previous = decoded
    ratio = nonwhite / (width * height)
    return {
        "width": width,
        "height": height,
        "nonwhite_pixels": nonwhite,
        "nonwhite_ratio": round(ratio, 6),
        "visible": nonwhite >= 32 and ratio >= 0.002,
    }
