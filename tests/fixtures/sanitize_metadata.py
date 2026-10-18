"""Remove account metadata from the synthetic Office-generated test fixtures."""

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile


def sanitize(directory):
    for path in Path(directory).glob('office.*'):
        if path.suffix not in {'.docx', '.xlsx', '.pptx'}:
            continue
        temporary = path.with_suffix('.tmp')
        with zipfile.ZipFile(path) as source, zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as target:
            for name in source.namelist():
                content = source.read(name)
                if name in {'docProps/core.xml', 'docProps/app.xml'}:
                    root = ET.fromstring(content)
                    for elem in root:
                        field = elem.tag.rsplit('}', 1)[-1]
                        if field in {'creator', 'lastModifiedBy'}:
                            elem.text = 'mdify synthetic fixture'
                        elif field in {'Company', 'Manager', 'HyperlinkBase', 'lastPrinted'}:
                            elem.text = ''
                    content = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                target.writestr(name, content)
        temporary.replace(path)


if __name__ == '__main__':
    sanitize(Path(__file__).parent)
