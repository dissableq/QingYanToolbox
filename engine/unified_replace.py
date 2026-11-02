"""Preview and run one rule set across document text and file/folder names."""
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import os
import zipfile
import uuid
from lxml import etree as ET
from engine.core import FileResult
from engine.safety import atomic_save, backup_file, validate_filename, replace_text_nodes
from engine.store import data_dir, _atomic_write
from engine import backups

TEXT_EXTS = {'.txt','.md','.csv','.tsv','.log','.json','.xml','.html','.htm','.srt',
             '.ini','.yaml','.yml','.css','.js','.py','.bat','.sql','.toml','.rtf'} - {'.rtf'}
OFFICE_EXTS = {'.docx','.docm','.xlsx','.xlsm','.pptx','.pptm'}
SUPPORTED = TEXT_EXTS | OFFICE_EXTS
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


class Cancelled(Exception):pass


def validate_options(options):
    root = Path(options.get('root','')).absolute()
    if not options.get('root') or not root.is_dir():raise ValueError('请选择有效的目标文件夹')
    if backups.is_backup_path(root) or not backups._no_links(root):raise ValueError('请选择普通文件夹，备份目录和链接目录不参与任务')
    if not any(options.get(k,False) for k in ('contents','file_names','folder_names')):
        raise ValueError('请至少勾选一种处理范围')
    rules = options.get('rules',[])
    if not rules:raise ValueError('请添加至少一组查找和替换规则')
    for i, rule in enumerate(rules,1):
        if not isinstance(rule, dict) or not isinstance(rule.get('find'), str) or not rule['find']:
            raise ValueError(f'第 {i} 行查找内容不能为空')
        if not isinstance(rule.get('replace'),str):raise ValueError(f'第 {i} 行替换内容必须是文字')
    return root, rules


def apply_rules(text, rules):
    count = 0
    for rule in rules:
        find, replacement = rule['find'], rule['replace']
        if find == replacement:continue
        count += text.count(find)
        text = text.replace(find, replacement)
    return text, count


def transformed(path, rules):
    """Return changed payload only; preserve unrelated package parts and formats."""
    path = Path(path);ext = path.suffix.lower()
    if ext in TEXT_EXTS:
        from engine.ops import detect_encoding
        raw = path.read_bytes();bom = b''
        for marker,codec in ((b'\xff\xfe','utf-16-le'),(b'\xfe\xff','utf-16-be'),(b'\xef\xbb\xbf','utf-8')):
            if raw.startswith(marker):bom=marker;encoding=codec;break
        else:encoding=detect_encoding(str(path))
        body = raw[len(bom):].decode(encoding, errors='strict')
        updated,hits = apply_rules(body,rules)
        return (bom+updated.encode(encoding,errors='strict') if updated!=body else None),hits
    if ext not in OFFICE_EXTS:return None,0
    changes={};count=0
    parser=ET.XMLParser(resolve_entities=False,no_network=True)
    with zipfile.ZipFile(path) as z:
        if len(z.namelist())!=len(set(z.namelist())):raise ValueError('文档含重名内部部件，请先用 Office 另存修复')
        for info in z.infolist():
            name=info.filename
            if not name.endswith('.xml'):continue
            family='word' if ext.startswith('.doc') else 'ppt' if ext.startswith('.ppt') else 'xl'
            if not name.startswith(family+'/'):continue
            raw=z.read(info);tree=ET.fromstring(raw,parser);groups=[]
            if family in ('word','ppt'):
                ns=W if family=='word' else A
                groups=[list(p.iter('{'+ns+'}t')) for p in tree.iter('{'+ns+'}p')]
            else:
                groups += [list(si.iter('{'+S+'}t')) for si in tree.iter('{'+S+'}si')]
                groups += [list(si.iter('{'+S+'}t')) for si in tree.iter('{'+S+'}is')]
                for cell in tree.iter('{'+S+'}c'):
                    if cell.get('t')=='str' and cell.find('{'+S+'}f') is None:
                        v=cell.find('{'+S+'}v')
                        if v is not None:groups.append([v])
            before=ET.tostring(tree);hits=0
            for nodes in groups:
                for rule in rules:
                    if rule['find']!=rule['replace']:
                        hits+=replace_text_nodes(nodes,rule['find'],rule['replace'])
            after=ET.tostring(tree)
            if before!=after:
                changes[name]=ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
                count+=hits
    return changes or None,count


def _write_payload(path,payload):
    if isinstance(payload,bytes):
        atomic_save(lambda tmp:Path(tmp).write_bytes(payload),str(path))
    else:
        def save(tmp):
            with zipfile.ZipFile(path) as zin,zipfile.ZipFile(tmp,'w') as zout:
                zout.comment=zin.comment
                for info in zin.infolist():zout.writestr(info,payload.get(info.filename,zin.read(info)))
        atomic_save(save,str(path))


def _mapped(path,old,new):
    try:return str(Path(new)/Path(path).relative_to(old))
    except ValueError:return str(path)


def make_plan(options,cancel=lambda:False,progress=lambda *args:None):
    root,rules=validate_options(options);paths=[]
    for dp,dirs,files in backups.walk_sources(root):
        paths.extend((Path(dp)/f,False) for f in sorted(files))
        if options.get('folder_names'):paths.extend((Path(dp)/d,True) for d in sorted(dirs))
        if not options.get('recursive',True):break
    if options.get('folder_names') and options.get('include_root'):paths.append((root,True))
    rows=[];name_targets=defaultdict(list)
    for i,(path,is_dir) in enumerate(paths):
        if cancel():raise Cancelled()
        progress(i,len(paths),path.name,0)
        row=dict(path=str(path),directory=is_dir,name=path.name,new_name=path.name,
                 final_path=str(path),hits=0,content=False,error='',note='')
        rename=options.get('folder_names') if is_dir else options.get('file_names')
        if rename:
            base=path.name if is_dir else path.stem
            new,_=apply_rules(base,rules)
            row['new_name']=new if is_dir else new+path.suffix
            try:
                validate_filename(row['new_name'])
                if is_dir and row['new_name']=='备份':raise ValueError('“备份”为程序保留目录名')
                if row['new_name']!=path.name:
                    target=path.with_name(row['new_name'])
                    if target.exists() and os.path.normcase(str(target))!=os.path.normcase(str(path)):
                        raise ValueError('目标已存在：'+str(target))
            except ValueError as exc:row['error']=str(exc)
        if not is_dir:
            stat=path.stat();row['signature']=[stat.st_size,stat.st_mtime_ns]
            if options.get('contents'):
                if path.suffix.lower() in SUPPORTED:
                    try:
                        payload,hits=transformed(path,rules)
                        row['hits']=hits;row['content']=payload is not None
                    except Exception as exc:row['error']='无法读取内容：'+str(exc)
                else:row['note']='该格式仅处理名称；不修改内部内容'
        if not row['error']:
            name_targets[os.path.normcase(str(path.with_name(row['new_name'])))].append(row)
        rows.append(row)
    for group in name_targets.values():
        if len(group)>1:
            for row in group:row['error']='多项将产生同名目标，请调整规则'
    dirs=sorted((r for r in rows if r['directory'] and not r['error'] and r['new_name']!=r['name']),key=lambda r:len(Path(r['path']).parts),reverse=True)
    for row in rows:
        if row['error']:continue
        final=str(Path(row['path']).with_name(row['new_name']))
        for d in dirs:
            if d is not row:final=_mapped(final,d['path'],str(Path(d['path']).with_name(d['new_name'])))
        row['final_path']=final
    return dict(root=str(root),options=options,rows=rows)


def execute(plan,cancel=lambda:False,progress=lambda *args:None,on_result=lambda r:None):
    rows=plan['rows'];rules=plan['options']['rules'];root=Path(plan['root'])
    # Reject all preview errors before any write, and require a fresh preview if input changed.
    errors=[r for r in rows if r['error']]
    if errors:raise ValueError(f'预览中有 {len(errors)} 项错误，请解决后重新预览')
    for row in rows:
        path=Path(row['path'])
        if not path.exists():raise ValueError('预览后路径已变化，请重新预览：'+str(path))
        if not row['directory']:
            stat=path.stat()
            if [stat.st_size,stat.st_mtime_ns]!=row['signature']:
                raise ValueError('预览后文件已变化，请重新预览：'+str(path))
        if row['new_name']!=row['name']:
            dst=path.with_name(row['new_name'])
            if dst.exists() and os.path.normcase(str(dst))!=os.path.normcase(str(path)):
                raise ValueError('目标名称已被占用，请重新预览：'+str(dst))
    if cancel():raise Cancelled()
    changed=[r for r in rows if r['content'] or r['new_name']!=r['name']]
    if not changed:return [],[],False,''
    task_id=datetime.now().strftime('%Y%m%d_%H%M%S_')+uuid.uuid4().hex[:8]
    actions=[];results=[];failed=set();done=0
    total=sum(int(r['content'])+int(r['new_name']!=r['name']) for r in changed)
    journal=backups.create_journal(root,task_id,{'task_id':task_id,'rules':rules,'actions':[]})
    def checkpoint():
        data={'task_id':task_id,'rules':rules,'actions':actions}
        _atomic_write(journal,data)
        # Journal contents change as the job proceeds. Refresh its registered hash.
        with backups._lock:
            saved=backups._load()
            for entry in saved:
                if entry['path']==journal:entry['sha256']=backups._hash(journal);entry['bytes']=Path(journal).stat().st_size
            _atomic_write(str(backups._index()),saved)
        _atomic_write(str(Path(data_dir())/'last_undo.json'),data)
    def emit(r):
        nonlocal done
        results.append(r);on_result(r);done+=1;progress(done,total,Path(r.path).name,0)
    try:
        backups.migrate_legacy(root,recursive=plan['options'].get('recursive',True))
        for row in changed:
            if cancel():break
            if not row['content']:continue
            path=row['path']
            try:
                payload,hits=transformed(path,rules)
                if payload is None:raise ValueError('内容已变化，请重新预览')
                bak=backup_file(path)
                action={'src':bak,'dst':path,'kind':'restore'}
                # Journal before the atomic replace retains recovery even on interrupted writes.
                actions.append(action);checkpoint()
                _write_payload(path,payload)
                emit(FileResult(path,'ok',f'内容替换 {hits} 处',path,undo=[action]))
            except Exception as exc:
                failed.add(path);emit(FileResult(path,'fail',str(exc)))
        renames=sorted((r for r in changed if r['new_name']!=r['name']),
                       key=lambda r:(r['directory'],-len(Path(r['path']).parts)))
        for row in renames:
            if cancel():break
            path=row['path'];dst=str(Path(path).with_name(row['new_name']))
            if path in failed:
                emit(FileResult(path,'skipped','内容处理失败，名称保持原样'));continue
            try:
                if Path(dst).exists() and os.path.normcase(dst)!=os.path.normcase(path):raise FileExistsError('目标名称已存在')
                os.rename(path,dst)
                action={'src':path,'dst':dst,'kind':'move'};actions.append(action)
                journal=_mapped(journal,path,dst)
                backups.remap(path,dst)
                for result in results:
                    if result.output:result.output=_mapped(result.output,path,dst)
                checkpoint()
                emit(FileResult(path,'ok','文件夹已改名' if row['directory'] else '文件已改名',dst,undo=[action]))
            except Exception as exc:emit(FileResult(path,'fail',str(exc)))
    finally:
        checkpoint()
        rec={'task_id':task_id,'time':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
             'kind':'unified','title':'统一替换','options':dict(plan['options'],op_name='内容与名称统一替换'),
             'total':len(results),'ok':sum(r.status=='ok' for r in results),
             'fail':sum(r.status=='fail' for r in results),'skipped':sum(r.status=='skipped' for r in results),
             'cancelled':bool(cancel()),'duration_sec':0,'preview_only':False,
             'files':[dict(path=r.path,status=r.status,detail=r.detail,output=r.output,outputs=[]) for r in results]}
        _atomic_write(str(Path(data_dir())/'records'/(task_id+'.json')),rec)
    return results,actions,bool(cancel()),task_id
