#!/usr/bin/env python3
"""Check a CommonLit teacher or student copy (.docx) against the brand and
template rules in this skill.

    python3 check_docx.py lesson.docx --mode teacher
    python3 check_docx.py lesson-student.docx --mode student

Prints problems as ERROR (must fix), WARN (probably fix) or INFO, and exits
with status 1 if there are errors. Uses only the Python standard library.
"""
import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'

SPEC = json.load(open(os.path.join(os.path.dirname(__file__), '..', 'assets', 'style-spec.json')))
ANSWER = SPEC['colors']['answer'].lower().lstrip('#')
TEXT_OK = {c.lower().lstrip('#') for c in SPEC['colors']['text'] + SPEC['colors'].get('smallPrint', [])}
LINK_OK = {c.lower().lstrip('#') for c in SPEC['colors']['link']}
NEVER_TEXT = {c.lower().lstrip('#') for c in SPEC['colors']['neverForText']}
FONTS_OK = set(SPEC['fonts']['allowed'])
FONTS_LEGACY = set(SPEC['fonts']['legacy'])

NAME_LINE = re.compile(r'^\s*name\s*:?\s*(_{3,}|.*\bclass\s*:)', re.I)
TEACHER_TEXT = [
    (re.compile(r'teacher copy', re.I), '"TEACHER COPY" label'),
    (re.compile(r'answers (are )?in blue', re.I), '"Answers in blue" notice'),
    (re.compile(r'to ensure (test|assessment) security', re.I), 'test-security note'),
    (re.compile(r'^\s*(notes? to (the )?teachers?|teacher notes?)\b', re.I), 'Notes to Teacher'),
    (re.compile(r'this page does not appear on the student copy', re.I), 'cover-page note'),
]
OPTIONAL_Q = re.compile(r'^\s*\*\s*[A-Z0-9]{1,2}[.)]\s')
COPYRIGHT = re.compile(r'(©|\(c\))\s*CommonLit,?\s*Inc', re.I)
CC_LINE = 'Unless otherwise noted, this content is licensed under the CC BY-NC-SA 4.0 license'


class Report:
    def __init__(self):
        self.items = []

    def add(self, level, msg, where=''):
        self.items.append((level, msg, where))

    def count(self, msg, where, sample=''):
        """Many hits of the same WARN: report once with a count."""
        key = (msg, where)
        self.counts = getattr(self, 'counts', {})
        if key not in self.counts:
            self.counts[key] = [0, sample]
        self.counts[key][0] += 1

    def print(self):
        for (msg, where), (n, sample) in getattr(self, 'counts', {}).items():
            self.items.append(('WARN', f'{msg}: {n} place(s)', where + (f': e.g. "{short(sample, 30)}"' if sample else '')))
        order = {'ERROR': 0, 'WARN': 1, 'INFO': 2}
        seen = set()
        for level, msg, where in sorted(self.items, key=lambda x: order[x[0]]):
            key = (level, msg)
            if key in seen and not where:
                continue
            seen.add(key)
            print(f'{level:5}  {msg}' + (f'  [{where}]' if where else ''))
        errors = sum(1 for i in self.items if i[0] == 'ERROR')
        warns = sum(1 for i in self.items if i[0] == 'WARN')
        print(f'\n{errors} error(s), {warns} warning(s)')
        return errors


def val(el, name, attr='val'):
    if el is None:
        return None
    child = el.find(W + name)
    return None if child is None else child.get(W + attr, 'true')


def text_of(node):
    return ''.join(t.text or '' for t in node.iter(W + 't'))


def run_text(r):
    return ''.join(t.text or '' for t in r.findall(W + 't'))


def short(s, n=60):
    s = ' '.join(s.split())
    return s if len(s) <= n else s[:n - 1] + '…'


def parts(z):
    """document.xml plus header/footer/footnote parts, with their types."""
    rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
    out = {'document': [('word/document.xml', ET.fromstring(z.read('word/document.xml')))]}
    ids = {}
    for rel in rels.iter(PKG + 'Relationship'):
        kind = rel.get('Type', '').rsplit('/', 1)[-1]
        target = 'word/' + rel.get('Target', '').lstrip('/').replace('word/', '', 1)
        ids[rel.get('Id')] = target
        if kind in ('header', 'footer', 'footnotes', 'endnotes') and target in z.namelist():
            out.setdefault(kind, []).append((target, ET.fromstring(z.read(target))))
    return out, ids


def check_runs(root, where, mode, rep):
    hyperlink_runs = {id(r) for h in root.iter(W + 'hyperlink') for r in h.iter(W + 'r')}
    for p in root.iter(W + 'p'):
        ptext = text_of(p)
        for r in p.iter(W + 'r'):
            t = run_text(r)
            rpr = r.find(W + 'rPr')
            color = (val(rpr, 'color') or '').lower()
            font = val(rpr, 'rFonts', 'ascii')
            bold = val(rpr, 'b')
            hl = val(rpr, 'highlight')
            if hl and hl != 'none':
                rep.add('ERROR' if mode == 'student' else 'INFO',
                        f'highlight ({hl})' + (' on a footnote number' if r.find(W + 'footnoteReference') is not None else ''),
                        f'{where}: "{short(ptext)}"')
            if not t.strip():
                continue
            if font and font not in FONTS_OK:
                level = 'INFO' if font in FONTS_LEGACY else 'WARN'
                rep.add(level, f'font "{font}" is not a CommonLit document font' +
                        (' (older template)' if font in FONTS_LEGACY else ''))
            if bold and bold not in ('0', 'false') and (font or '').startswith('Open Sans') \
                    and 'Medium' not in font and 'SemiBold' not in font:
                rep.count('Bold applied to Open Sans (use the Open Sans Medium font instead)', where, t)
            if color and color != 'auto':
                if color == ANSWER and id(r) in hyperlink_runs:
                    rep.add('INFO', 'a link is colored #0000FF (answer blue); link blue #1155CC/#035FE6 avoids confusion',
                            f'{where}: "{short(t, 40)}"')
                elif color == ANSWER:
                    if mode == 'student':
                        rep.add('ERROR', 'answer-blue text left in the student copy', f'{where}: "{short(t)}"')
                elif color in NEVER_TEXT:
                    rep.add('ERROR', f'text in accent color #{color.upper()} (type must be ink)', f'{where}: "{short(t, 40)}"')
                elif color in LINK_OK:
                    if id(r) not in hyperlink_runs:
                        rep.add('WARN', f'link-blue #{color.upper()} on text that is not a link (answers must be #0000FF)',
                                f'{where}: "{short(t, 40)}"')
                elif color not in TEXT_OK and color != 'ffffff':
                    rep.add('WARN', f'text color #{color.upper()} is not ink/black', f'{where}: "{short(t, 40)}"')
        # Paragraph-mark formatting: what a student types on this line inherits it.
        ppr = p.find(W + 'pPr')
        mark = ppr.find(W + 'rPr') if ppr is not None else None
        if mode == 'student' and mark is not None:
            if (val(mark, 'color') or '').lower() == ANSWER:
                rep.count('hidden answer blue on line endings (anything a student types there will be blue)', where)
            if val(mark, 'highlight') not in (None, 'none'):
                rep.count('hidden highlight on line endings', where)


def body_blocks(body):
    return [el for el in body if el.tag in (W + 'p', W + 'tbl')]


def first_student_index(blocks):
    for i, el in enumerate(blocks):
        if el.tag == W + 'tbl':
            cell = el.find(f'.//{W}tc')
            if cell is not None and re.match(r'\s*name\b', text_of(cell), re.I):
                return i
        elif NAME_LINE.search(text_of(el)):
            return i
    return -1


def check_document(doc, mode, rep):
    body = doc.find(W + 'body')
    blocks = body_blocks(body)
    idx = first_student_index(blocks)
    if idx < 0:
        rep.add('ERROR', 'no Name / Class line or table found (the student part must start with it)')
    elif mode == 'student':
        before = [b for b in blocks[:idx] if text_of(b).strip()]
        if before:
            rep.add('ERROR', f'{len(before)} block(s) with text before Name / Class (cover page left?)',
                    f'"{short(text_of(before[0]))}"')
    elif mode == 'teacher' and idx == 0:
        rep.add('WARN', 'no cover page (Lesson Overview) before Name / Class')

    paras = list(body.iter(W + 'p'))
    texts = [text_of(p) for p in paras]
    if mode == 'teacher':
        if not any(re.match(r'\s*TEACHER COPY:', t) for t in texts):
            rep.add('INFO', 'no "TEACHER COPY: <title>" line (fine in newer templates, where the header says Teacher Copy)')
        if not any(re.search(r'answers (are )?in blue', t, re.I) for t in texts):
            rep.add('WARN', 'no "Answers in blue" notice')
        blue = sum(1 for r in body.iter(W + 'r') if (val(r.find(W + 'rPr'), 'color') or '').lower() == ANSWER and run_text(r).strip())
        if not blue:
            rep.add('WARN', 'no answers in #0000FF found')
    else:
        for t in texts:
            for pat, what in TEACHER_TEXT:
                if pat.search(t):
                    rep.add('ERROR', f'teacher-only text left: {what}', f'"{short(t)}"')
            if OPTIONAL_Q.match(t):
                rep.add('WARN', 'optional (*) question left in the student copy', f'"{short(t)}"')

    # Empty bullets: a list item with no text shows a stray bullet.
    for p in paras:
        ppr = p.find(W + 'pPr')
        if ppr is not None and ppr.find(W + 'numPr') is not None and not text_of(p).strip() \
                and p.find(f'.//{W}drawing') is None:
            rep.add('WARN', 'empty bulleted/numbered line (shows a stray bullet)')

    # Headings in ALL CAPS
    for p in paras:
        sty = val(p.find(W + 'pPr'), 'pStyle') or ''
        t = text_of(p).strip()
        if sty.lower().startswith(('heading', 'title')) and len(t) > 6 and t.isupper():
            rep.add('WARN', 'ALL CAPS heading (use sentence or title case)', f'"{short(t)}"')

    # First-page header (large logo on page 1)
    sects = list(body.iter(W + 'sectPr'))
    first = sects[0] if sects else None
    if first is not None:
        has_first = any(h.get(W + 'type') == 'first' for h in first.findall(W + 'headerReference'))
        if not (first.find(W + 'titlePg') is not None and has_first):
            rep.add('WARN', 'page 1 has no "different first page" header (large logo)')
    return sects


def check_footers(footers, mode, rep):
    texts = []
    for path, root in footers:
        for p in root.iter(W + 'p'):
            t = text_of(p)
            texts.append(t)
            instr = ' '.join((i.text or '') for i in p.iter(W + 'instrText'))
            instr += ' '.join(f.get(W + 'instr', '') for f in p.iter(W + 'fldSimple'))
            if re.search(r'\bPAGE\b', instr):
                tabs_in_text = sum(len(r.findall(W + 'tab')) for r in p.iter(W + 'r'))
                ppr = p.find(W + 'pPr')
                jc = val(ppr, 'jc')
                right_stop = ppr is not None and any(
                    tab.get(W + 'val') in ('right', 'end') for tab in ppr.iter(W + 'tab'))
                if tabs_in_text > 1:
                    rep.add('WARN', f'page number pushed right with {tabs_in_text} tabs; use one right tab stop',
                            path)
                elif not (right_stop or jc in ('right', 'end')):
                    rep.add('WARN', 'page number may not be right-aligned', path)
    joined = '\n'.join(texts)
    if mode == 'teacher':
        if not COPYRIGHT.search(joined):
            rep.add('INFO', 'no "© CommonLit, Inc." line in the footer')
    else:
        if COPYRIGHT.search(joined):
            rep.add('ERROR', '"© CommonLit, Inc." footer left in the student copy (use the CC BY-NC-SA line)')
        if CC_LINE not in joined:
            rep.add('ERROR', 'CC BY-NC-SA 4.0 footer line missing')


def check_headers(headers, mode, rep):
    for path, root in headers:
        t = text_of(root)
        if mode == 'student' and re.search(r'teacher copy', t, re.I):
            rep.add('ERROR', 'header still says "Teacher Copy"', path)
        if mode == 'teacher' and re.search(r'student copy', t, re.I):
            rep.add('WARN', 'teacher copy header says "Student Copy"', path)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('docx')
    ap.add_argument('--mode', choices=['teacher', 'student'], required=True)
    args = ap.parse_args()

    rep = Report()
    with zipfile.ZipFile(args.docx) as z:
        found, _ = parts(z)
        doc = found['document'][0][1]
        check_document(doc, args.mode, rep)
        check_runs(doc, 'body', args.mode, rep)
        for kind in ('header', 'footer', 'footnotes', 'endnotes'):
            for path, root in found.get(kind, []):
                check_runs(root, path, args.mode, rep)
        check_headers(found.get('header', []), args.mode, rep)
        check_footers(found.get('footer', []), args.mode, rep)
        if args.mode == 'student' and any(n.startswith('word/comments') for n in z.namelist()):
            rep.add('ERROR', 'the file still contains comments')
        name = os.path.basename(args.docx)
        if args.mode == 'student' and re.search(r'teacher', name, re.I):
            rep.add('WARN', 'file name still says TEACHER', name)

    print(f'{os.path.basename(args.docx)} ({args.mode} copy)\n')
    sys.exit(1 if rep.print() else 0)


if __name__ == '__main__':
    main()
