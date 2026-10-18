"""Relationship-aware Open XML package reading, with explicit failures."""

import io
import posixpath
import xml.etree.ElementTree as ET
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass
from urllib.parse import unquote

R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


@contextmanager
def open_package(source):
    if isinstance(source, zipfile.ZipFile):
        yield source
        return
    try:
        with zipfile.ZipFile(io.BytesIO(source) if isinstance(source, bytes) else source) as archive:
            yield archive
    except zipfile.BadZipFile as exc:
        raise ValueError('Invalid or encrypted Office package (expected a readable ZIP file)') from exc


def read_xml(archive, part):
    try:
        root = ET.fromstring(archive.read(part))
    except KeyError as exc:
        raise ValueError(f'Missing Office package part: {part}') from exc
    except ET.ParseError as exc:
        raise ValueError(f'Invalid XML in {part}: {exc}') from exc
    # Strict and Transitional use equivalent names for the elements we read.
    namespaces = {
        'http://purl.oclc.org/ooxml/wordprocessingml/main': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'http://purl.oclc.org/ooxml/spreadsheetml/main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'http://purl.oclc.org/ooxml/presentationml/main': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'http://purl.oclc.org/ooxml/drawingml/main': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'http://purl.oclc.org/ooxml/officeDocument/relationships': R,
    }
    def normalized(name):
        for old, new in namespaces.items():
            name = name.replace('{' + old + '}', '{' + new + '}')
        return name
    for elem in root.iter():
        elem.tag = normalized(elem.tag)
        elem.attrib = {normalized(key): value for key, value in elem.attrib.items()}
    return root


@dataclass(frozen=True)
class Relationship:
    kind: str
    target: str
    external: bool = False


def relationships(archive, part=''):
    rel_part = (posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
                if part else '_rels/.rels')
    if rel_part not in archive.namelist():
        return {}
    result = {}
    for elem in read_xml(archive, rel_part):
        rid, target = elem.get('Id'), elem.get('Target')
        if not rid or target is None:
            raise ValueError(f'Invalid relationship in {rel_part}')
        external = elem.get('TargetMode') == 'External'
        if not external:
            target = unquote(target)
            target = posixpath.normpath(target.lstrip('/') if target.startswith('/') else
                                       posixpath.join(posixpath.dirname(part), target))
            if target == '..' or target.startswith('../'):
                raise ValueError(f'Relationship escapes package root in {rel_part}')
        if rid in result:
            raise ValueError(f'Duplicate relationship {rid} in {rel_part}')
        result[rid] = Relationship(elem.get('Type', '').rsplit('/', 1)[-1], target, external)
    return result


def related_part(archive, rels, kind, fallback=None):
    for rel in rels.values():
        if rel.kind == kind and not rel.external:
            if rel.target not in archive.namelist():
                raise ValueError(f'Missing {kind} part: {rel.target}')
            return rel.target
    if fallback and fallback in archive.namelist():
        return fallback
    return None


def main_part(archive, expected_root):
    part = related_part(archive, relationships(archive), 'officeDocument')
    if not part:
        raise ValueError('Missing officeDocument relationship; not a valid Office document')
    root = read_xml(archive, part)
    if root.tag != expected_root:
        raise ValueError(f'Wrong Office document type in {part}')
    return part, root
