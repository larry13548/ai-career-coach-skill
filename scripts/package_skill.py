#!/usr/bin/env python3
"""Package only executable skill resources, never candidate files or local docs."""
import hashlib
import json
from pathlib import Path
import zipfile

ROOT=Path(__file__).resolve().parents[1]
VERSION='1.1.0'

def package():
    names=['SKILL.md','README.md','requirements.txt','scripts/score_match.py','scripts/resume_docx.py']
    names += [str(p.relative_to(ROOT)) for folder in ('references','assets') for p in sorted((ROOT/folder).iterdir()) if p.is_file() and p.suffix in ('.md','.docx')]
    for name in names:
        p=ROOT/name
        if p.is_symlink() or not p.is_file():raise ValueError(f'Invalid resource: {name}')
    out=ROOT/'dist';out.mkdir(exist_ok=True)
    manifest={'version':VERSION,'files':{name:hashlib.sha256((ROOT/name).read_bytes()).hexdigest() for name in names}}
    manifest_path=out/'manifest.json';manifest_path.write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
    archive=out/f'ai-career-coach-skill-{VERSION}.zip'
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
        for name in names:z.write(ROOT/name,'ai-career-coach-skill/'+name)
        z.write(manifest_path,'ai-career-coach-skill/manifest.json')
    with zipfile.ZipFile(archive) as z:
        if z.testzip():raise ValueError('Archive integrity failure')
        assert all(not any(part in n for part in ('private-samples','tests/','validation/')) for n in z.namelist())
    print(archive)
    return archive

if __name__=='__main__':package()
