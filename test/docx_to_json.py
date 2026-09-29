"""Turns a .docx into the JSON tree test/mock-docs.js loads.

Only keeps what the converter looks at: paragraphs vs. list items, list
nesting level, page breaks, tables, and per-run text/color/bold.

    python3 test/docx_to_json.py teacher.docx > test/fixtures/teacher.json
"""
import json
import sys
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def val(el, name):
    child = el.find(W + name) if el is not None else None
    return None if child is None else child.get(W + 'val', 'true')


def truthy(v):
    return v is not None and v not in ('0', 'false')


def paragraph(p):
    ppr = p.find(W + 'pPr')
    num = ppr.find(W + 'numPr') if ppr is not None else None
    out = {'type': 'PARAGRAPH', 'runs': [], 'pageBreak': False}
    if num is not None and num.find(W + 'numId') is not None:
        out['type'] = 'LIST_ITEM'
        out['level'] = int(val(num, 'ilvl') or 0)
        out['listId'] = val(num, 'numId')
    for r in p.iter(W + 'r'):
        rpr = r.find(W + 'rPr')
        color = val(rpr, 'color')
        color = None if color in (None, 'auto', 'true') else '#' + color.lower()
        bold = truthy(val(rpr, 'b'))
        hl = val(rpr, 'highlight')
        bg = None if hl in (None, 'none', 'true') else hl
        text = ''
        for child in r:
            if child.tag == W + 't':
                text += child.text or ''
            elif child.tag == W + 'tab':
                text += '\t'
            elif child.tag == W + 'br' and child.get(W + 'type') == 'page':
                out['pageBreak'] = True
        if text:
            out['runs'].append({'text': text, 'color': color, 'bold': bold, 'bg': bg})
    return out


def container(el):
    children = []
    for child in el:
        if child.tag == W + 'p':
            children.append(paragraph(child))
        elif child.tag == W + 'tbl':
            rows = []
            for tr in child.findall(W + 'tr'):
                rows.append([container(tc) for tc in tr.findall(W + 'tc')])
            children.append({'type': 'TABLE', 'rows': rows})
    return children


def main(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    json.dump(container(root.find(W + 'body')), sys.stdout, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main(sys.argv[1])
