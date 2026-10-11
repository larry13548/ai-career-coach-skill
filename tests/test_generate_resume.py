import copy
import json
from pathlib import Path
import sys
import tempfile
import time
import unittest
from unittest.mock import patch
from zipfile import ZipFile
from docx import Document
from docx.oxml.ns import qn
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import generate_resume as g
ROOT=Path(__file__).resolve().parents[1]

class FastGenerationTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.dir=Path(self.tmp.name)
        self.data=json.loads((ROOT/'tests/fixtures/fast-resume.json').read_text())
    def text(self,path):
        d=Document(path)
        return '\n'.join(t.text or '' for t in d.element.body.iter(qn('w:t')))
    def test_all_nine_preserve_both_summaries_and_remove_empty_module(self):
        for i in range(1,10):
            self.data['template_id']=f'T{i:02d}';f=self.dir/f'{i}.docx';r=g.fill(self.data,f)
            s=self.text(f)
            for line in self.data['summary_original']+self.data['summary_additions']:self.assertIn(line,s)
            self.assertNotIn('荣誉奖项',s);self.assertNotIn('视频创作实践',s);self.assertNotIn('{{',s)
            self.assertIn('照片待上传',s);self.assertTrue(r['content_verified']);self.assertEqual(r['visual_status'],'not_verified')
            with ZipFile(f) as z:
                xml=z.read('word/document.xml')
                self.assertNotIn(b'AlternateContent',xml);self.assertNotIn(b'txbxContent',xml)
    def test_refuses_unconfirmed_and_unselected(self):
        for key,value in [('confirmed',False),('template_id','T10')]:
            d=copy.deepcopy(self.data);d[key]=value
            with self.assertRaises(ValueError):g.fill(d,self.dir/'bad.docx')
    def test_custom_practice_needs_explicit_permission(self):
        d=copy.deepcopy(self.data);d['sections'].append({'kind':'practice','title':'视频创作实践','entries':[{'detail':'拍摄一次'}]})
        with self.assertRaises(ValueError):g.fill(d,self.dir/'bad.docx')
        d['sections'][-1]['independent_section_confirmed']=True
        g.fill(d,self.dir/'ok.docx');self.assertIn('视频创作实践',self.text(self.dir/'ok.docx'))
    def test_summary_not_silently_omitted(self):
        self.data['omit_summary']=True
        with self.assertRaises(ValueError):g.fill(self.data,self.dir/'bad.docx')
        self.data['summary_omission_confirmed']=True;g.fill(self.data,self.dir/'ok.docx')
        self.assertNotIn('个人介绍',self.text(self.dir/'ok.docx'))
    def test_existing_file_never_overwritten(self):
        f=self.dir/'x.docx';f.write_bytes(b'keep')
        with self.assertRaises(ValueError):g.fill(self.data,f)
        self.assertEqual(f.read_bytes(),b'keep')
    def test_session_starts_once_and_expires(self):
        state=self.dir/'session.json';g.start(state)
        with self.assertRaises(FileExistsError):g.start(state)
        state.write_text(json.dumps({'started_at':time.time()-301,'attempts':0}))
        inp=self.dir/'input.json';inp.write_text(json.dumps(self.data))
        with self.assertRaises(TimeoutError):g.generate(inp,self.dir/'bad.docx',state)
        self.assertFalse((self.dir/'bad.docx').exists())
    def test_at_most_two_corrections(self):
        state=self.dir/'session.json';g.start(state);inp=self.dir/'input.json';inp.write_text(json.dumps(self.data))
        for i in range(3):self.assertEqual(g.generate(inp,self.dir/f'{i}.docx',state)['attempt'],i+1)
        with self.assertRaises(ValueError):g.generate(inp,self.dir/'fourth.docx',state)
    def test_one_engine_failure_does_not_claim_visual_success(self):
        with patch.object(g,'run_command',side_effect=TimeoutError('budget')) as run:
            result=g.render(self.dir/'resume.docx',{'started_at':time.time()},soffice='/fixed/soffice')
            self.assertEqual(run.call_count,1);self.assertEqual(result['visual_status'],'not_verified')
    def test_subprocess_timeout_uses_remaining_budget(self):
        with patch.object(g.subprocess,'run') as run:
            g.run_command(['renderer'],{'started_at':time.time()-280},60)
            self.assertGreater(run.call_args.kwargs['timeout'],0)
            self.assertLessEqual(run.call_args.kwargs['timeout'],17)
    def test_picture_has_one_inline_reference(self):
        from PIL import Image
        photo=self.dir/'photo.png';Image.new('RGB',(120,180),'gray').save(photo)
        self.data['photo']=str(photo);out=self.dir/'photo.docx';g.fill(self.data,out)
        self.assertEqual(len(Document(out).inline_shapes),1)
        self.assertNotIn('照片待上传',self.text(out))

if __name__=='__main__':unittest.main()
