"""Consistent decoding for files and byte streams."""

import codecs
import re


def decode_text(data: bytes | str, encoding: str | None = None, *, html=False) -> str:
    if isinstance(data, str):
        return data.lstrip('\ufeff')
    if encoding:
        return data.decode(encoding).lstrip('\ufeff')
    for bom, codec in ((codecs.BOM_UTF32_LE, 'utf-32'), (codecs.BOM_UTF32_BE, 'utf-32'),
                       (codecs.BOM_UTF16_LE, 'utf-16'), (codecs.BOM_UTF16_BE, 'utf-16'),
                       (codecs.BOM_UTF8, 'utf-8-sig')):
        if data.startswith(bom):
            return data.decode(codec).lstrip('\ufeff')
    if html:
        declaration = re.search(br'<meta\b[^>]*charset\s*=\s*[\"\x27\s]*([\w-]+)', data[:4096], re.I)
        if declaration:
            return data.decode(declaration[1].decode('ascii')).lstrip('\ufeff')
    for codec in ('utf-8', 'cp1252', 'latin-1'):
        try:
            return data.decode(codec)
        except UnicodeDecodeError:
            pass
    raise ValueError('Unable to decode text')  # latin-1 accepts every byte
