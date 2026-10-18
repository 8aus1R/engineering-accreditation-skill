"""Read-only verification of a delivery inventory and CSV file references.

This verifies file identity, layout of folders, and references, not grades or page rendering.
"""
import argparse
import csv
import hashlib
import json
from pathlib import Path

def inside(root,relative):
    value=Path(relative)
    if value.is_absolute():
        raise ValueError('Inventory/CSV path must be relative: '+str(relative))
    path=(root/value).resolve()
    if not path.is_relative_to(root) or path==root:
        raise ValueError('Path escapes delivery root: '+str(relative))
    return path

def verify(root,inventory='文件与分数核验记录.json',csv_files=(),path_column='文件',require_binary_doc=False):
    root=Path(root).resolve(strict=True)
    data=json.loads(inside(root,inventory).read_text(encoding='utf-8-sig'))
    records=data['documents_inventory']
    expected={}
    for record in records:
        path=inside(root,record['file'])
        if path in expected:raise ValueError('Duplicate inventory document: '+record['file'])
        if path.suffix.lower() not in ('.doc','.docx'):raise ValueError('Not a Word document: '+record['file'])
        raw=path.read_bytes()
        if hashlib.sha256(raw).hexdigest()!=record['sha256']:
            raise ValueError('SHA-256 mismatch: '+record['file'])
        if require_binary_doc and (path.suffix.lower()!='.doc' or raw[:8]!=bytes.fromhex('D0CF11E0A1B11AE1')):
            raise ValueError('Not binary Word DOC: '+record['file'])
        expected[path]=record
    if not expected:raise ValueError('Empty document inventory')
    if 'documents' in data and data['documents']!=len(expected):raise ValueError('Document count disagrees with inventory')
    found={p.resolve() for p in root.rglob('*') if p.is_file() and p.suffix.lower() in ('.doc','.docx')}
    if found!=set(expected):
        raise ValueError('Missing or extra Word files: '+json.dumps({'missing':[str(p) for p in set(expected)-found],'extra':[str(p) for p in found-set(expected)]},ensure_ascii=False))
    folders=[]
    for folder in data.get('assignment_folders',[]):
        directory=inside(root,folder['folder'])
        count=sum(p.parent==directory for p in expected)
        if not directory.is_dir() or count!=folder['documents']:raise ValueError('Assignment folder count mismatch: '+folder['folder'])
        folders.append({'folder':folder['folder'],'documents':count})
    csv_results=[]
    for name in csv_files:
        references=set()
        count=0
        with inside(root,name).open(encoding='utf-8-sig',newline='') as f:
            reader=csv.DictReader(f)
            if not reader.fieldnames or path_column not in reader.fieldnames:raise ValueError('CSV missing path column: '+name)
            for row in reader:
                path=inside(root,row[path_column])
                if path not in expected:raise ValueError('CSV references unknown document: '+row[path_column])
                references.add(path)
                count+=1
        if references!=set(expected):raise ValueError('CSV does not cover every document: '+name)
        csv_results.append({'file':name,'rows':count,'document_references':len(references)})
    return {'status':'PASS','root':str(root),'documents':len(expected),'all_hashes_match':True,'no_extra_word_files':True,'folders':folders,'csv':csv_results,'grades_and_rendering_checked_by_this_script':False}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',required=True)
    parser.add_argument('--inventory',default='文件与分数核验记录.json')
    parser.add_argument('--csv',dest='csv_files',action='append',default=[])
    parser.add_argument('--path-column',default='文件')
    parser.add_argument('--require-binary-doc',action='store_true')
    args=parser.parse_args()
    print(json.dumps(verify(args.root,args.inventory,args.csv_files,args.path_column,args.require_binary_doc),ensure_ascii=False))
