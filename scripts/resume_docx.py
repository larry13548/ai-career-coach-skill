#!/usr/bin/env python3
"""Build editable resumes or patch confirmed text in an existing Word file."""
import argparse
import copy
import json
import os
from pathlib import Path
import sys
import tempfile
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_BREAK

ROOT = Path(__file__).resolve().parents[1]
FONT = 'Microsoft YaHei'


def load(path):
    with open(path, encoding='utf-8') as f: return json.load(f)


def string(value, label, optional=False):
    if not isinstance(value, str) or (not optional and not value.strip()):
        raise ValueError(f'{label} must be a {"possibly empty " if optional else "nonempty "}string')
    return value


def color(value):
    if not isinstance(value, str) or len(value) != 6 or any(c not in '0123456789abcdefABCDEF' for c in value):
        raise ValueError('accent must be a six-digit RGB value')
    return value.upper()


def save_new(doc, output):
    output = Path(output)
    if output.exists(): raise ValueError('output already exists; use a new filename')
    output.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive final creation prevents accidental overwrite, including a racing process.
    fd, tmp = tempfile.mkstemp(suffix='.docx', dir=output.parent)
    os.close(fd)
    try:
        doc.save(tmp)
        with open(tmp, 'rb') as src, output.open('xb') as dst:
            dst.write(src.read())
    finally:
        Path(tmp).unlink(missing_ok=True)


def font(run, size=11, bold=False, rgb='000000'):
    run.font.name = FONT; run.font.size = Pt(size); run.bold = bold
    props = run._element.get_or_add_rPr()
    rf = props.rFonts
    for k in ('ascii', 'hAnsi', 'eastAsia', 'cs'): rf.set(qn('w:'+k), FONT)
    run.font.color.rgb = RGBColor.from_string(rgb)


def p(container, text, bold=False, bullet=False, keep=False):
    a = container.add_paragraph()
    pf = a.paragraph_format
    pf.line_spacing = Pt(18); pf.space_before = Pt(0); pf.space_after = Pt(9)
    pf.keep_with_next = keep; pf.widow_control = True
    if bullet:
        pf.left_indent = Cm(.35); pf.first_line_indent = Cm(-.35)
        text = '• ' + text
    font(a.add_run(text), bold=bold)
    return a


def heading(container, text, accent=None, icon=None):
    a = container.add_paragraph(style='Heading 1')
    a.paragraph_format.space_before = Pt(22); a.paragraph_format.space_after = Pt(9)
    a.paragraph_format.line_spacing = Pt(20); a.paragraph_format.keep_with_next = True
    if icon:
        a.paragraph_format.line_spacing=Pt(22)
        a.add_run().add_picture(icon,width=Cm(.5));font(a.add_run('  '))
    elif accent: font(a.add_run('●  '), size=10, rgb=accent)
    font(a.add_run(text), size=13, bold=True)
    return a


def shade(cell, fill):
    pr = cell._tc.get_or_add_tcPr(); e = OxmlElement('w:shd'); e.set(qn('w:fill'), fill); pr.append(e)


def prep_cell(cell):
    first = cell.paragraphs[0]
    first.paragraph_format.line_spacing = Pt(1); first.paragraph_format.space_after = Pt(0)
    first.paragraph_format.space_before = Pt(0)
    font(first.add_run(''), size=1)
    margins = OxmlElement('w:tcMar')
    for edge in ('top','left','bottom','right'):
        el = OxmlElement('w:'+edge); el.set(qn('w:w'), '150'); el.set(qn('w:type'), 'dxa'); margins.append(el)
    cell._tc.get_or_add_tcPr().append(margins)


def identity(doc, container, data):
    if data.get('photo'):
        from docx.image.image import Image
        image=Image.from_file(data['photo'])
        table=container.add_table(rows=1,cols=2)
        table.autofit=False
        for cell,col,width in zip(table.rows[0].cells,table.columns,[13.2,3.4]):
            col.width=Cm(width);cell.width=Cm(width);prep_cell(cell)
        left,right=table.rows[0].cells
        no_photo=dict(data);no_photo.pop('photo');identity(doc,left,no_photo)
        right.paragraphs[0].paragraph_format.line_spacing=1.0
        right.paragraphs[0].add_run().add_picture(data['photo'],width=Cm(min(2.4,3.2*image.width/image.height)))
        return
    a = container.add_paragraph(style='Title')
    a.paragraph_format.line_spacing = Pt(32); a.paragraph_format.space_after = Pt(9)
    a.paragraph_format.keep_with_next = True
    font(a.add_run(data['name']), size=24)
    p(container, '求职意向  '+data['target'], bold=True, keep=True)
    if data['contact']: p(container, '  |  '.join(data['contact']))


def sections(container, items, accent):
    for section in items:
        h = heading(container, section['title'], accent, section.get('icon'))
        if section.get('page_break_before'): h.paragraph_format.page_break_before = True
        for item in section['entries']:
            title = '  |  '.join(v for v in (item.get('heading'),item.get('period')) if v)
            if title: p(container, title, bold=True, keep=bool(item.get('detail') or item.get('bullets')))
            if item.get('detail'): p(container, item['detail'], keep=bool(item.get('bullets')))
            for text in item.get('bullets', []): p(container, text, bullet=True)


def corner_art(doc, accent):
    # Editable vector shapes, not a full-page raster. Enable only when matching reference artwork.
    # First page top-left and each page bottom-right; report differences from the supplied source.
    from lxml import etree
    sec = doc.sections[0]
    def add_lines(part, points, prefix):
        para = part.paragraphs[0]
        for i, coords in enumerate(points):
            x1,y1,x2,y2=coords
            xml = f'''<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:pict><v:line id="{prefix}{i}" from="{x1}pt,{y1}pt" to="{x2}pt,{y2}pt" strokecolor="#{accent}" strokeweight="5pt" style="position:absolute;z-index:-251654144;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"/></w:pict></w:r>'''
            para._p.append(etree.fromstring(xml.encode()))
    sec.different_first_page_header_footer = True
    add_lines(sec.first_page_header, [(12,12,90,12),(12,12,12,90)], 'careerTop')
    width = sec.page_width.pt; height = sec.page_height.pt
    lines = [(width-12,height-90,width-12,height-12),(width-90,height-12,width-12,height-12)]
    add_lines(sec.first_page_footer, lines,'careerBottomFirst')
    add_lines(sec.footer,lines,'careerBottom')


def background_images(doc, artworks):
    """Anchor original decorative image assets behind editable text using explicit coordinates."""
    from docx.image.image import Image
    sec=doc.sections[0];sec.different_first_page_header_footer=True
    for art in artworks:
        targets=[sec.first_page_header,sec.header] if art['scope']=='all' else [sec.first_page_header] if art['scope']=='first' else [sec.header]
        image=Image.from_file(art['path'])
        width=Cm(art['width_cm']);height=int(width*image.height/image.width)
        if Cm(art['x_cm'])+width>sec.page_width or Cm(art['y_cm'])+height>sec.page_height:
            raise ValueError('background artwork exceeds page bounds; supply measured coordinates')
        for part in targets:
            inline=part.paragraphs[0].add_run().add_picture(art['path'],width=width)._inline
            anchor=OxmlElement('wp:anchor')
            for k,v in {'distT':'0','distB':'0','distL':'0','distR':'0','simplePos':'0','relativeHeight':'0','behindDoc':'1','locked':'0','layoutInCell':'1','allowOverlap':'1'}.items():anchor.set(k,v)
            simple=OxmlElement('wp:simplePos');simple.set('x','0');simple.set('y','0');anchor.append(simple)
            for axis,value in [('H',art['x_cm']),('V',art['y_cm'])]:
                position=OxmlElement('wp:position'+axis);position.set('relativeFrom','page')
                offset=OxmlElement('wp:posOffset');offset.text=str(int(Cm(value)));position.append(offset);anchor.append(position)
            for tag in ['wp:extent','wp:effectExtent']:
                elem=inline.find(qn(tag))
                if elem is not None:anchor.append(copy.deepcopy(elem))
            anchor.append(OxmlElement('wp:wrapNone'))
            for tag in ['wp:docPr','wp:cNvGraphicFramePr','a:graphic']:
                elem=inline.find(qn(tag))
                if elem is not None:anchor.append(copy.deepcopy(elem))
            inline.getparent().replace(inline,anchor)


def validate_resume(data):
    if not isinstance(data, dict):
        raise ValueError('resume must be an object')
    string(data.get('name'), 'name'); string(data.get('target'), 'target')
    if 'photo' in data: string(data['photo'],'photo')
    if not isinstance(data.get('contact'), list): raise ValueError('contact must be a list')
    for t in data['contact']: string(t, 'contact item')
    if not isinstance(data.get('sections'), list):
        raise ValueError('sections must be a list')
    for sec in data['sections']:
        string(sec.get('title'),'section title')
        if 'icon' in sec: string(sec['icon'],'section icon')
        if sec.get('column','main') not in ('main','sidebar'): raise ValueError('invalid column')
        if 'page_break_before' in sec and type(sec['page_break_before']) is not bool: raise ValueError('page_break_before must be boolean')
        if not isinstance(sec.get('entries'),list) or not sec['entries']: raise ValueError('remove empty sections')
        for ent in sec['entries']:
            if not isinstance(ent,dict): raise ValueError('entry must be an object')
            for k in ('heading','period','detail'):
                if k in ent: string(ent[k],k,optional=True)
            bs=ent.get('bullets',[])
            if not isinstance(bs,list): raise ValueError('bullets must be a list')
            for t in bs: string(t,'bullet')
            if not any(ent.get(k) for k in ('heading','detail','bullets')): raise ValueError('empty entry')
    visual=data.get('visual',{})
    if not isinstance(visual,dict): raise ValueError('visual must be an object')
    if 'accent' in visual: color(visual['accent'])
    artworks=visual.get('backgrounds',[])
    if not isinstance(artworks,list): raise ValueError('backgrounds must be a list')
    for art in artworks:
        if not isinstance(art,dict):raise ValueError('background must be an object')
        string(art.get('path'),'background path')
        if art.get('scope') not in ('first','rest','all'):raise ValueError('invalid artwork scope')
        import math
        for key in ('x_cm','y_cm','width_cm'):
            value=art.get(key)
            if type(value) not in (int,float) or not math.isfinite(value) or value<0 or (key=='width_cm' and value==0):raise ValueError('invalid artwork coordinate')
    for k in ('section_markers','corners'):
        if k in visual and type(visual[k]) is not bool: raise ValueError(f'{k} must be boolean')


def build(data, output, layout):
    if not isinstance(data,dict): raise ValueError('resume must be an object')
    data=copy.deepcopy(data)
    for key in ('name','target'):
        if data.get(key) in (None,''): data[key]='待补充'
    if data.get('photo') in (None,''): data.pop('photo',None)
    data.setdefault('contact',[]); data.setdefault('sections',[])
    if isinstance(data['contact'],list):
        import re
        contact_text=' '.join(str(x) for x in data['contact'])
        if not re.search(r'手机|电话|[+\d][\d ()-]{6,}',contact_text): data['contact'].append('手机：待补充')
        if not re.search(r'邮箱|邮件|@',contact_text): data['contact'].append('邮箱：待补充')
        if not data.get('photo'): data['contact'].append('照片：待上传')
    validate_resume(data)
    doc=Document(ROOT/'assets'/f'resume-{layout}.docx')
    for child in list(doc._element.body):
        if child.tag != qn('w:sectPr'): doc._element.body.remove(child)
    visual=data.get('visual',{})
    accent=color(visual.get('accent','FFC000')) if visual.get('section_markers') else None
    if layout == 'banner':
        table=doc.add_table(rows=1,cols=1); table.autofit=False; table.columns[0].width=Cm(17.2)
        c=table.cell(0,0); prep_cell(c); shade(c,'EEF0F2'); identity(doc,c,data)
    else: identity(doc,doc,data)
    if layout == 'sidebar':
        side=[s for s in data['sections'] if s.get('column')=='sidebar']
        main=[s for s in data['sections'] if s.get('column')!='sidebar']
        # Missing sections are allowed; do not block draft delivery.
        table=doc.add_table(rows=1,cols=2); table.autofit=False
        for cell, col, width in zip(table.rows[0].cells, table.columns, [5.1,12.1]):
            col.width=Cm(width);cell.width=Cm(width);prep_cell(cell)
        a,b=table.rows[0].cells; shade(a,'F0F1F2')
        sections(a,side,accent);sections(b,main,accent)
    else: sections(doc,data['sections'],accent)
    if visual.get('backgrounds'): background_images(doc,visual['backgrounds'])
    if visual.get('corners'): corner_art(doc,color(visual.get('accent','FFC000')))
    # Small continuation header provides context without repeating an entire experience.
    sec=doc.sections[0];sec.different_first_page_header_footer=True
    hp=sec.header.paragraphs[0]
    font(hp.add_run(data['name']+'  |  '+data['target']+'  |  续页'),size=9,rgb='666666')
    doc.core_properties.author='';doc.core_properties.title=data['name']+' '+data['target']
    save_new(doc,output)
    return {'file':str(output),'layout':layout,'visual_preservation':'requires_rendered_comparison','text_editable':True}


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
    doc=Document(source);paras=body_paragraphs(doc);seen=set()
    for change in edits['edits']:
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
    save_new(doc,output)
    return {'file':str(output),'patched_paragraphs':len(seen),'note':'paragraph-level formatting preserved; mixed inline emphasis may require reapplication; render and inspect'}


def main():
    parser=argparse.ArgumentParser(description=__doc__);sub=parser.add_subparsers(dest='mode',required=True)
    b=sub.add_parser('build');b.add_argument('input');b.add_argument('output');b.add_argument('--layout',choices=['single-column','sidebar','banner'],default='single-column')
    i=sub.add_parser('inspect');i.add_argument('input')
    e=sub.add_parser('patch');e.add_argument('input');e.add_argument('edits');e.add_argument('output')
    a=parser.parse_args()
    try:
        if a.mode=='build':
            data=load(a.input)
            base=Path(a.input).resolve().parent
            def absolute(value):return str(base/value) if not Path(value).is_absolute() else value
            if data.get('photo'):data['photo']=absolute(data['photo'])
            for sec in data.get('sections',[]):
                if sec.get('icon'):sec['icon']=absolute(sec['icon'])
            for art in data.get('visual',{}).get('backgrounds',[]):art['path']=absolute(art['path'])
            result=build(data,a.output,a.layout)
        elif a.mode=='inspect': result=inspect(a.input)
        else:result=patch(a.input,load(a.edits),a.output)
        print(json.dumps(result,ensure_ascii=False,indent=2))
    except (OSError,ValueError,TypeError,KeyError) as exc:
        print(f'Cannot generate: {exc}',file=sys.stderr);return 2
    return 0

if __name__=='__main__':sys.exit(main())
