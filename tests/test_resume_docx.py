import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest
from zipfile import ZipFile
from docx import Document
from docx.shared import Pt
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from resume_docx import build,inspect,patch

ROOT=Path(__file__).resolve().parents[1]


class ResumeTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.dir=Path(self.tmp.name)
        self.data=json.loads((ROOT/'tests/fixtures/resume.json').read_text())

    def test_all_layouts_preserve_confirmed_content_and_spacing(self):
        for layout in ('single-column','sidebar','banner'):
            f=self.dir/(layout+'.docx');build(self.data,f,layout)
            texts='\n'.join(x['text'] for x in inspect(f))
            for sec in self.data['sections']:
                self.assertIn(sec['title'],texts)
                for e in sec['entries']:
                    for v in [e.get('heading',''),e.get('period',''),e.get('detail','')]+e.get('bullets',[]):self.assertIn(v,texts)
            d=Document(f)
            from docx.oxml.ns import qn
            for p in d.element.body.iter(qn('w:p')):
                text=''.join(t.text or '' for t in p.iter(qn('w:t')))
                if text.startswith('• '):
                    spacing=p.find(qn('w:pPr')).find(qn('w:spacing'))
                    self.assertEqual(spacing.get(qn('w:line')),'360')
                    self.assertEqual(spacing.get(qn('w:after')),'180')
            with self.assertRaises(ValueError):build(self.data,f,layout)

    def test_unconfirmed_draft_and_missing_identity_do_not_block(self):
        d=copy.deepcopy(self.data);d['confirmed']=False
        d.pop('name');d['contact']=[]
        f=self.dir/'draft.docx';build(d,f,'single-column')
        text='\n'.join(p['text'] for p in inspect(f))
        for value in ('待补充','手机：待补充','邮箱：待补充','照片：待上传'):self.assertIn(value,text)
        self.assertIn(d['sections'][0]['title'],text)

    def test_patch_preserves_unedited_content_and_refuses_stale_text(self):
        src=self.dir/'source.docx';build(self.data,src,'single-column')
        rows=inspect(src);index=next(x['index'] for x in rows if x['text']==self.data['name'])
        edits={'edits':[{'index':index,'expected':self.data['name'],'replacement':'合成候选人乙'}]}
        dest=self.dir/'patched.docx';patch(src,edits,dest)
        after=inspect(dest)
        for before,a in zip(rows,after):self.assertEqual(a['text'],'合成候选人乙' if before['index']==index else before['text'])
        with ZipFile(src) as a,ZipFile(dest) as b:
            for name in a.namelist():
                if name=='word/document.xml':continue
                if name.endswith(('.xml','.rels')):
                    from lxml import etree
                    self.assertEqual(etree.tostring(etree.fromstring(a.read(name)),method='c14n'),etree.tostring(etree.fromstring(b.read(name)),method='c14n'))
                else:self.assertEqual(a.read(name),b.read(name))
        edits['edits'][0]['expected']='过期文本'
        with self.assertRaises(ValueError):patch(src,edits,self.dir/'bad.docx')

    def test_patch_rejects_drawing_paragraph(self):
        src=self.dir/'complex.docx';d=Document(ROOT/'assets/resume-single-column.docx')
        from docx.oxml import OxmlElement
        p=d.add_paragraph('图文混合');p.add_run()._r.append(OxmlElement('w:drawing'));d.save(src)
        row=inspect(src)[-1];self.assertTrue(row['complex'])
        with self.assertRaises(ValueError):patch(src,{'confirmed':True,'edits':[{'index':row['index'],'expected':row['text'],'replacement':'新文字'}]},self.dir/'bad.docx')

    def test_vector_decoration_is_not_rasterized_body(self):
        d=copy.deepcopy(self.data);d['visual']={'accent':'FFC000','section_markers':True,'corners':True}
        f=self.dir/'decorated.docx';build(d,f,'single-column')
        with ZipFile(f) as z:
            self.assertIn(b'careerTop',z.read('word/header1.xml'))
            self.assertFalse(any(n.startswith('word/media/') for n in z.namelist()))
        self.assertIn(self.data['name'],'\n'.join(p['text'] for p in inspect(f)))

    def test_photo_icons_and_background_keep_editable_text(self):
        import struct,zlib
        def chunk(t,body):return struct.pack('>I',len(body))+t+body+struct.pack('>I',zlib.crc32(t+body)&0xffffffff)
        image=self.dir/'test.png'
        image.write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',20,20,8,2,0,0,0))+chunk(b'IDAT',zlib.compress((b'\0'+b'\xff\xc0\0'*20)*20))+chunk(b'IEND',b''))
        d=copy.deepcopy(self.data);d['photo']=str(image)
        d['sections'][0]['icon']=str(image)
        d['visual']={'backgrounds':[{'path':str(image),'x_cm':.3,'y_cm':.3,'width_cm':.5,'scope':'all'}]}
        for layout in ('single-column','sidebar','banner'):
            f=self.dir/('images-'+layout+'.docx');build(d,f,layout)
            self.assertIn(self.data['name'],'\n'.join(p['text'] for p in inspect(f)))
            with ZipFile(f) as z:
                self.assertTrue(any(n.startswith('word/media/') for n in z.namelist()))
                self.assertTrue(any(b'behindDoc="1"' in z.read(n) for n in z.namelist() if n.startswith('word/header') and n.endswith('.xml')))


if __name__=='__main__':unittest.main()
