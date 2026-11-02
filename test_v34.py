"""Real content, backup lifecycle, and multi-step recovery regressions."""
import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
from pathlib import Path
import tempfile
import unittest
import json
import zipfile
from unittest.mock import patch
import docx
import openpyxl
from pptx import Presentation
from pptx.util import Inches
from engine import backups,unified_replace as u
from engine.safety import backup_file
from engine.core import collect_files
from engine.undo import undo_actions
from engine.store import _atomic_write,data_dir


class Tests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)/'old项目';self.root.mkdir()
        self.env=patch.dict(os.environ,APPDATA=str(Path(self.temp.name)/'data'));self.env.start()
    def tearDown(self):self.env.stop();self.temp.cleanup()
    def options(self,**kw):
        o=dict(root=str(self.root),rules=[dict(find='old',replace='new'),dict(find='公司',replace='企业')],contents=True,file_names=True,folder_names=True,recursive=True,include_root=False);o.update(kw);return o
    def snapshot(self):return {p.relative_to(self.root).as_posix():p.read_bytes() for p in self.root.rglob('*') if p.is_file() and '备份' not in p.parts}

    def test_versioned_backups_cleanup_and_untracked_files(self):
        p=self.root/'原文件.txt';p.write_bytes(b'one');a=Path(backup_file(str(p)));p.write_bytes(b'two');b=Path(backup_file(str(p)))
        self.assertEqual(a.parent.name,'备份');self.assertNotEqual(a,b);self.assertEqual(a.read_bytes(),b'one');self.assertEqual(b.read_bytes(),b'two')
        other=a.parent/'我的文件.txt';other.write_text('personal');a.write_bytes(b'changed by user')
        removed,errors=backups.cleanup({r['id'] for r in backups.existing()})
        self.assertEqual(len(removed),1);self.assertEqual(len(errors),1);self.assertTrue(a.exists());self.assertTrue(other.exists());self.assertEqual(p.read_bytes(),b'two')

    def test_migrate_legacy_preserves_undo(self):
        p=self.root/'report.txt';p.write_bytes(b'after');old=Path(str(p)+'.qingyanbak');old.write_bytes(b'before')
        _atomic_write(str(Path(data_dir())/'last_undo.json'),{'actions':[dict(src=str(old),dst=str(p),kind='restore')]})
        moved=backups.migrate_legacy(self.root);self.assertEqual(len(moved),1);self.assertFalse(old.exists())
        actions=json.loads((Path(data_dir())/'last_undo.json').read_text(encoding='utf-8'))['actions'];_,remaining,errors=undo_actions(actions)
        self.assertFalse(remaining,errors);self.assertFalse(errors);self.assertEqual(p.read_bytes(),b'before')

    def test_preview_execute_all_formats_nested_names_and_exact_undo(self):
        folder=self.root/'old公司';folder.mkdir()
        (folder/'old.txt').write_bytes('old公司\r\nkeep'.encode('utf-8-sig'))
        d=docx.Document();r=d.add_paragraph().add_run('ol');r.bold=True;d.paragraphs[0].add_run('d公司');d.sections[0].header.paragraphs[0].text='old';d.save(folder/'old.docx')
        wb=openpyxl.Workbook();wb.active['A1']='old公司';wb.active['B1']='="old"';wb.active['A1'].font=openpyxl.styles.Font(bold=True);wb.save(folder/'old.xlsx');wb.close()
        pr=Presentation();slide=pr.slides.add_slide(pr.slide_layouts[6]);shape=slide.shapes.add_textbox(Inches(1),Inches(1),Inches(4),Inches(1));shape.text='old公司';pr.save(folder/'old.pptx')
        (folder/'old.pdf').write_bytes(b'opaque pdf');snapshot=self.snapshot()
        plan=u.make_plan(self.options(include_root=True));self.assertFalse(any(r['error'] for r in plan['rows']));self.assertEqual(self.snapshot(),snapshot);self.assertFalse((folder/'备份').exists())
        results,actions,cancelled,_=u.execute(plan);self.assertFalse(cancelled);self.assertFalse([r for r in results if r.status=='fail'])
        newroot=self.root.with_name('new项目');target=newroot/'new企业';self.assertTrue(target.exists())
        self.assertEqual((target/'new.txt').read_bytes(),'new企业\r\nkeep'.encode('utf-8-sig'))
        d=docx.Document(target/'new.docx');self.assertEqual(d.paragraphs[0].text,'new企业');self.assertTrue(d.paragraphs[0].runs[0].bold);self.assertEqual(d.sections[0].header.paragraphs[0].text,'new')
        wb=openpyxl.load_workbook(target/'new.xlsx');self.assertEqual(wb.active['A1'].value,'new企业');self.assertEqual(wb.active['B1'].value,'="old"');self.assertTrue(wb.active['A1'].font.bold);wb.close()
        self.assertEqual(Presentation(target/'new.pptx').slides[0].shapes[0].text,'new企业');self.assertEqual((target/'new.pdf').read_bytes(),b'opaque pdf')
        self.assertTrue(all(Path(r['path']).exists() for r in backups.existing()))
        _,remaining,errors=undo_actions(actions);self.assertFalse(remaining,errors);self.assertEqual(self.snapshot(),snapshot)

    def test_conflicts_invalid_names_and_changed_input_block_before_writes(self):
        p=self.root/'old.txt';p.write_text('old');(self.root/'new.txt').write_text('keep')
        plan=u.make_plan(self.options());self.assertTrue(any(r['error'] for r in plan['rows']))
        with self.assertRaises(ValueError):u.execute(plan)
        self.assertEqual(p.read_text(),'old');self.assertFalse((self.root/'备份').exists())
        plan=u.make_plan(self.options(rules=[dict(find='old',replace='bad/name')]))
        self.assertTrue(any(r['error'] for r in plan['rows']))
        plan=u.make_plan(self.options(file_names=False));p.write_text('changed after preview')
        with self.assertRaises(ValueError):u.execute(plan)
        self.assertFalse((self.root/'备份').exists())

    def test_cancel_and_failed_file_are_recoverable(self):
        for i in range(8):(self.root/f'old{i}.txt').write_text('old')
        plan=u.make_plan(self.options());done=[]
        results,actions,cancelled,_=u.execute(plan,cancel=lambda:len(done)>=2,on_result=done.append)
        self.assertTrue(cancelled);self.assertEqual(len(results),2)
        _,left,errors=undo_actions(actions);self.assertFalse(left,errors)
        self.assertTrue(all(p.read_text()=='old' for p in self.root.glob('*.txt')))
        plan=u.make_plan(self.options())
        original=u._write_payload
        def fail_one(path,payload):
            if str(path).endswith('old0.txt'):raise PermissionError('occupied fixture')
            original(path,payload)
        with patch.object(u,'_write_payload',fail_one):results,actions,_,_=u.execute(plan)
        self.assertTrue(any(r.status=='fail' for r in results));self.assertTrue((self.root/'old0.txt').exists());self.assertTrue((self.root/'new1.txt').exists())
        _,left,errors=undo_actions(actions);self.assertFalse(left,errors)

    def test_backups_excluded_and_recursive_option(self):
        (self.root/'old.txt').write_text('old');sub=self.root/'oldsub';sub.mkdir();(sub/'old.txt').write_text('old')
        backup_file(str(self.root/'old.txt'))
        paths=collect_files(str(self.root),True);self.assertEqual(len(paths),2)
        plan=u.make_plan(self.options(recursive=False));self.assertFalse(any('备份' in r['path'] for r in plan['rows']));self.assertEqual(len(plan['rows']),2)
        plan=u.make_plan(self.options(contents=False,file_names=False));self.assertEqual(sum(r['new_name']!=r['name'] for r in plan['rows']),1)

    def test_rules_order_and_bom_preserved(self):
        for codec,bom in [('utf-16-be',b'\xfe\xff'),('utf-16-le',b'\xff\xfe'),('gb18030',b'')]:
            p=self.root/(codec+'.txt');p.write_bytes(bom+'old公司\r\n'.encode(codec));payload,hits=u.transformed(p,[dict(find='old',replace='公司'),dict(find='公司',replace='企业')]);self.assertEqual(payload,bom+'企业企业\r\n'.encode(codec));self.assertEqual(hits,3)

    def test_cleanup_after_directory_rename_removes_historical_restore_actions(self):
        folder=self.root/'old';folder.mkdir();p=folder/'old.txt';p.write_text('old')
        _,actions,_,_=u.execute(u.make_plan(self.options()))
        rows=backups.existing();removed,errors=backups.cleanup({r['id'] for r in rows})
        self.assertFalse(errors)
        remaining=json.loads((Path(data_dir())/'last_undo.json').read_text(encoding='utf-8'))['actions']
        self.assertTrue(remaining);self.assertFalse(any(a.get('kind')=='restore' for a in remaining))
        _,left,errors=undo_actions(remaining);self.assertFalse(left,errors)
        self.assertEqual(p.read_text(),'new')

    def test_migrated_in_memory_undo_resolves_legacy_path(self):
        p=self.root/'a.txt';p.write_bytes(b'after');old=Path(str(p)+'.qingyanbak');old.write_bytes(b'before')
        actions=[dict(src=str(old),dst=str(p),kind='restore')]
        backups.migrate_legacy(self.root)
        _,left,errors=undo_actions(actions);self.assertFalse(left,errors);self.assertEqual(p.read_bytes(),b'before')


if __name__=='__main__':unittest.main(verbosity=2)
