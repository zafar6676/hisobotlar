#!/usr/bin/env python3
from __future__ import annotations
import re, zipfile, subprocess
from pathlib import Path
from collections import defaultdict
import xml.etree.ElementTree as ET

ROOT = Path('.')
SKIP_PARTS = {'.git', 'old'}
EXTS = {'.xlsx', '.xltx', '.xlsm', '.docx', '.doc'}
CODE_RE = re.compile(r'^(?:[A-Z]{1,3}-?\d{1,4}|\d{2,4})$')


def should_skip(path: Path) -> bool:
    return any(p in SKIP_PARTS for p in path.parts)


def looks_like_code(tok: str) -> bool:
    tok = tok.strip().upper().replace(' ', '')
    if not tok or len(tok) > 8 or not CODE_RE.match(tok):
        return False
    if tok.isdigit() and 1900 <= int(tok) <= 2100:
        return False
    return True


def get_report_tag(path: Path) -> str:
    m = re.search(r'(\d{1,3})', str(path))
    return m.group(1) if m else path.stem[:20]


def extract_docx_rows(path: Path):
    try:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read('word/document.xml'))
    except Exception:
        return []
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    rows=[]
    for tr in root.findall('.//w:tr',ns):
        r=[''.join(t.text or '' for t in tc.findall('.//w:t',ns)).strip() for tc in tr.findall('.//w:tc',ns)]
        if any(r): rows.append(r)
    for p in root.findall('.//w:p',ns):
        t=''.join(x.text or '' for x in p.findall('.//w:t',ns)).strip()
        if t: rows.append([t])
    return rows


def extract_xlsx_rows(path: Path):
    try:
        with zipfile.ZipFile(path) as zf:
            sns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            shared=[]
            if 'xl/sharedStrings.xml' in zf.namelist():
                sroot=ET.fromstring(zf.read('xl/sharedStrings.xml'))
                for si in sroot.findall('.//s:si',sns):
                    shared.append(''.join(t.text or '' for t in si.findall('.//s:t',sns)))
            rows=[]
            for sf in sorted(n for n in zf.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')):
                root=ET.fromstring(zf.read(sf))
                for row in root.findall('.//s:sheetData/s:row',sns):
                    vals=[]
                    for c in row.findall('s:c',sns):
                        t=c.attrib.get('t')
                        v=c.find('s:v',sns)
                        if v is None or v.text is None:
                            isel=c.find('s:is',sns)
                            txt=''.join(x.text or '' for x in isel.findall('.//s:t',sns)) if isel is not None else ''
                            vals.append(txt.strip()); continue
                        val=v.text
                        if t=='s':
                            try: val=shared[int(val)]
                            except: pass
                        vals.append(str(val).strip())
                    if any(vals): rows.append(vals)
            return rows
    except Exception:
        return []


def extract_doc_rows(path: Path):
    try:
        cp=subprocess.run(['strings',str(path)],capture_output=True,text=True,timeout=20)
    except Exception:
        return []
    lines=[ln.strip() for ln in cp.stdout.splitlines() if ln.strip()]
    return [[ln] for ln in lines[:2000]]


def find_codes(rows):
    out=set()
    for row in rows:
        non=[c for c in row if c and str(c).strip()]
        if not non: continue
        has_alpha=any(any(ch.isalpha() for ch in c) for c in non)
        for cell in non[:8]:
            for part in re.split(r'[\s,;:()\[\]{}]+',str(cell).replace('\n',' ')):
                p=part.strip(' ."\'“”‘’')
                if looks_like_code(p):
                    if p.isdigit() and not has_alpha and len(non)<2:
                        continue
                    out.add(p)
    return out


def main():
    files=sorted([p for p in ROOT.rglob('*') if p.is_file() and p.suffix.lower() in EXTS and not should_skip(p)])
    code_to_reports=defaultdict(set)
    stats={'parsed':0,'failed':0}

    for p in files:
        ext=p.suffix.lower()
        rows=extract_xlsx_rows(p) if ext in {'.xlsx','.xltx','.xlsm'} else extract_docx_rows(p) if ext=='.docx' else extract_doc_rows(p)
        if not rows:
            stats['failed']+=1; continue
        stats['parsed']+=1
        rep=get_report_tag(p)
        for c in find_codes(rows):
            code_to_reports[c].add(rep)

    lines=[
        '# Field Key Taxonomy (All Reports)',
        '',
        'First column is the field key (`code`). Third column shows where the same key appears in other reports.',
        '',
        f"- Files discovered: **{len(files)}**",
        f"- Files parsed successfully: **{stats['parsed']}**",
        f"- Files with parse issues: **{stats['failed']}**",
        f"- Unique field keys found: **{len(code_to_reports)}**",
        '',
        '| Code (field key) | Primary report | Used in other report(s) |',
        '|---|---|---|'
    ]

    def sk(c): return (0,int(c)) if c.isdigit() else (1,c)
    for code in sorted(code_to_reports,key=sk):
        reps=sorted(code_to_reports[code], key=lambda x:(len(x),x))
        primary=reps[0]
        others=', '.join(reps[1:]) if len(reps)>1 else '—'
        lines.append(f"| `{code}` | `{primary}` | {others} |")

    Path('FIELD_KEY_TAXONOMY.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print(f"Wrote FIELD_KEY_TAXONOMY.md with {len(code_to_reports)} keys")

if __name__=='__main__':
    main()
