#!/usr/bin/env python3
"""Inspect or patch confirmed text in a selected source template; no fallback layout."""
import argparse
import copy
import json
from pathlib import Path
import sys
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from zipfile import ZipFile
from io import BytesIO
from lxml import etree

def load(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))

def string(value, label, optional=False):
    if not isinstance(value, str) or (not optional and not value.strip()):
        raise ValueError(label+' must be a string')
    if any(ord(c)<32 and c not in '\t\n\r' for c in value):
        raise ValueError('invalid XML character')
    return value

def body_paragraphs(doc):
    return list(doc.element.body.iter(qn('w:p')))


def inspect(path):
    doc=Document(path)
    return [{'index':i,'text':''.join(t.text or '' for t in p.iter(qn('w:t'))),
             'complex':any(e.tag in {qn('w:drawing'),qn('w:pict'),qn('w:fldChar'),qn('w:hyperlink'),qn('w:ins'),qn('w:del')} for e in p.iter())}
            for i,p in enumerate(body_paragraphs(doc))]


def patch(source, edits, output):
    if Path(source).resolve()==Path(output).resolve(): raise ValueError('preserve the original; choose a different output')
    if not isinstance(edits,dict) or not isinstance(edits.get('edits'),list):
        raise ValueError('edits need an edits list')
    if edits.get('confirmed') is not True:
        raise ValueError('final resume text must be confirmed before generating Word')
    if edits.get('template_id') not in {f'T{i:02d}' for i in range(1,10)}:
        raise ValueError('select a template T01–T09 first')
    if Path(output).exists(): raise ValueError('output already exists; choose a new filename')
    doc=Document(source);paras=body_paragraphs(doc);seen=set()
    for change in edits['edits']:
        if not isinstance(change,dict): raise ValueError('change must be an object')
        i=change.get('index')
        if type(i) is not int or not 0<=i<len(paras) or i in seen: raise ValueError('invalid or duplicate paragraph index')
        seen.add(i);para=paras[i]
        if any(c.tag not in (qn('w:pPr'),qn('w:r')) for c in para): raise ValueError('complex paragraph; use a structure-aware editor')
        runs=list(para.findall(qn('w:r')))
        if any(c.tag not in (qn('w:rPr'),qn('w:t')) for run in runs for c in run):
            raise ValueError('paragraph contains drawing/field/break; cannot safely replace as plain text')
        text=''.join(t.text or '' for t in para.iter(qn('w:t')))
        if text!=change.get('expected'): raise ValueError('source text changed; inspect again before patching')
        string(change.get('replacement'),'replacement',optional=True)
        formatting=copy.deepcopy(runs[0].find(qn('w:rPr'))) if runs and runs[0].find(qn('w:rPr')) is not None else None
        for run in runs: para.remove(run)
        new=OxmlElement('w:r')
        if formatting is not None:new.append(formatting)
        t=OxmlElement('w:t');t.set(qn('xml:space'),'preserve');t.text=change['replacement'];new.append(t);para.append(new)
    # Keep every other package part byte-identical, including drawings and relationships.
    buffer=BytesIO()
    with ZipFile(source) as src, ZipFile(buffer,'w') as dst:
        for entry in src.infolist():
            data=etree.tostring(doc.element,xml_declaration=True,encoding='UTF-8',standalone=True) if entry.filename=='word/document.xml' else src.read(entry.filename)
            dst.writestr(entry,data)
    Path(output).parent.mkdir(parents=True,exist_ok=True)
    with Path(output).open('xb') as out: out.write(buffer.getvalue())
    return {'file':str(output),'patched_paragraphs':len(seen),'note':'paragraph-level formatting preserved; mixed inline emphasis may require reapplication; render and inspect'}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    sub=parser.add_subparsers(dest='mode',required=True)
    i=sub.add_parser('inspect');i.add_argument('input')
    e=sub.add_parser('patch');e.add_argument('input');e.add_argument('edits');e.add_argument('output')
    a=parser.parse_args()
    try:
        result=inspect(a.input) if a.mode=='inspect' else patch(a.input,load(a.edits),a.output)
        print(json.dumps(result,ensure_ascii=False,indent=2))
    except (OSError,ValueError,TypeError,KeyError) as exc:
        print(f'Cannot generate: {exc}',file=sys.stderr);return 2
    return 0

if __name__=='__main__':sys.exit(main())
