"""Content-level regressions for the 3.2 redesign. All files are temporary."""
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

import docx
from docx.shared import RGBColor
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
import openpyxl
from engine import office_ops as O, advanced_ops as A, ops
from engine.core import BatchWorker, TaskSpec, FileResult
import installer_entry as installer


class QualityTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix='qy_quality_')
        self.root = Path(self.tmp.name)
        self.out = self.root / 'out'
        self.out.mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def test_word_replace_preserves_surrounding_styles_and_picture(self):
        p = self.root / 'mixed.docx'
        img = self.root / 'red.png'
        Image.new('RGB', (16, 16), 'red').save(img)
        d = docx.Document()
        para = d.add_paragraph()
        para.add_run('Keep ').italic = True
        para.add_run('tar').bold = True
        para.add_run('get').font.color.rgb = RGBColor(255, 0, 0)
        para.add_run(' suffix').underline = True
        para.add_run().add_picture(str(img))
        d.save(p)
        r = O.docx_replace(str(p), dict(find='target', replace='NEW', out_dir=str(self.out)))
        self.assertEqual(r.status, 'ok')
        result = docx.Document(r.output)
        self.assertEqual(result.paragraphs[0].text, 'Keep NEW suffix')
        self.assertTrue(next(x for x in result.paragraphs[0].runs if 'suffix' in x.text).underline)
        self.assertTrue(next(x for x in result.paragraphs[0].runs if 'Keep' in x.text).italic)
        self.assertEqual(len(result.inline_shapes), 1)

    def test_word_regex_replacement_windows_path_is_literal(self):
        p = self.root / 'r.docx'
        d = docx.Document(); d.add_paragraph('ROOT'); d.save(p)
        r = O.docx_replace(str(p), dict(find='R.O.T|ROOT', regex=True, replace=r'C:\new', out_dir=str(self.out)))
        self.assertEqual(docx.Document(r.output).paragraphs[0].text, r'C:\new')

    def test_remove_word_images_preserves_same_run_text_and_valid_package(self):
        image = self.root / 'pic.png'
        Image.new('RGB', (20,20), 'red').save(image)
        p = self.root / 'pictures.docx'
        d = docx.Document()
        run = d.add_paragraph().add_run('before ')
        run.add_picture(str(image)); run.add_text(' after')
        d.sections[0].header.paragraphs[0].add_run().add_picture(str(image))
        d.save(p)
        result = O.docx_remove_images(str(p), {'out_dir':str(self.out)})
        self.assertEqual(result.status, 'ok')
        reopened = docx.Document(result.output)
        self.assertEqual(reopened.paragraphs[0].text, 'before  after')
        self.assertEqual(len(reopened.inline_shapes), 0)

    def test_ppt_split_retains_existing_files_and_all_outputs(self):
        p=self.root/'slides.pptx'; prs=Presentation()
        for i in range(3): prs.slides.add_slide(prs.slide_layouts[6])
        prs.save(p)
        sentinel=self.out/'slides_part001.pptx'; sentinel.write_bytes(b'keep')
        result=A.pptx_split(str(p), {'ppt_chunk':1,'out_dir':str(self.out)})
        self.assertEqual(sentinel.read_bytes(), b'keep')
        self.assertEqual(len(result.outputs),3)
        for output in result.outputs: self.assertEqual(len(Presentation(output).slides),1)

    def test_undo_partial_failure_keeps_pending_actions_for_retry(self):
        from engine.undo import undo_actions
        src=self.root/'original.txt'; dst=self.root/'renamed.txt'
        src.write_text('occupied'); dst.write_text('moved')
        backup=self.root/'saved.bak'; edited=self.root/'edited.txt'
        backup.write_text('before'); edited.write_text('after')
        actions=[{'src':str(src),'dst':str(dst)},
                 {'src':str(backup),'dst':str(edited),'kind':'restore'}]
        completed,pending,errors=undo_actions(actions)
        self.assertEqual(len(completed),1); self.assertEqual(pending,actions[:1])
        self.assertTrue(errors); self.assertEqual(edited.read_text(),'before')
        self.assertEqual(src.read_text(),'occupied')
        src.unlink()
        completed,pending,errors=undo_actions(pending)
        self.assertFalse(errors); self.assertFalse(pending)
        self.assertEqual(src.read_text(),'moved')

    def test_ppt_replace_preserves_unmatched_style(self):
        p = self.root / 'mixed.pptx'
        prs = Presentation(); sl = prs.slides.add_slide(prs.slide_layouts[6])
        para = sl.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2)).text_frame.paragraphs[0]
        para.add_run().text = 'target '
        end = para.add_run(); end.text = 'suffix'; end.font.italic = True
        prs.save(p)
        r = O.pptx_replace(str(p), dict(find='target', replace='NEW', out_dir=str(self.out)))
        runs = Presentation(r.output).slides[0].shapes[0].text_frame.paragraphs[0].runs
        self.assertEqual(''.join(x.text for x in runs), 'NEW suffix')
        self.assertTrue(next(x for x in runs if 'suffix' in x.text).font.italic)

    def test_word_split_never_removes_existing_user_files(self):
        p = self.root / 'book.docx'
        d = docx.Document(); d.add_paragraph('one'); d.add_paragraph('two'); d.save(p)
        sentinel = self.out / 'book_part001.docx'; sentinel.write_bytes(b'USER ORIGINAL')
        r = A.docx_split_by_paragraphs(str(p), dict(docx_chunk=1, out_dir=str(self.out)))
        self.assertEqual(r.status, 'ok')
        self.assertEqual(sentinel.read_bytes(), b'USER ORIGINAL')
        self.assertNotEqual(r.output, str(sentinel))

    def test_empty_undo_replaces_previous_task(self):
        manifest = self.root / 'undo.json'
        manifest.write_text(json.dumps({'task_id':'old','actions':[{'src':'old','dst':'new'}]}))
        worker = BatchWorker(TaskSpec('test','test', ['file']), lambda p,o: FileResult(p,'ok'), str(self.root/'records'), str(manifest))
        worker.run()
        self.assertEqual(json.loads(manifest.read_text())['actions'], [])

    def test_independent_backups_do_not_invalidate_previous_undo(self):
        p = self.root / 'word.docx'
        d = docx.Document(); d.add_paragraph('first'); d.save(p)
        r1 = O.docx_replace(str(p), dict(find='first', replace='second'))
        original = Path(r1.undo[0]['src']).read_bytes()
        r2 = O.docx_replace(str(p), dict(find='second', replace='third'))
        self.assertNotEqual(r1.undo[0]['src'], r2.undo[0]['src'])
        self.assertEqual(Path(r1.undo[0]['src']).read_bytes(), original)

    def test_excel_delete_rows_updates_formula_references(self):
        p=self.root/'formulas.xlsx'; wb=openpyxl.Workbook(); ws=wb.active; ws.title='Data'
        ws['A1']=10; ws['A3']=20; ws['B3']='=$A$1+A3'; ws['B3'].number_format='0.00'
        wb.create_sheet('Summary')['A1']="='Data'!A3"
        wb.save(p)
        r=O.xlsx_delete_rows(str(p),dict(row_mode='blank',out_dir=str(self.out)))
        result=openpyxl.load_workbook(r.output)
        self.assertEqual(result['Data']['B2'].value,'=$A$1+A2')
        self.assertEqual(result['Summary']['A1'].value,"='Data'!A2")
        self.assertEqual(result['Data']['B2'].number_format,'0.00')
        result.close(); wb.close()

    def test_duplicate_rows_keep_distinct_types(self):
        p=self.root/'types.xlsx'; wb=openpyxl.Workbook(); ws=wb.active
        ws.append([1]); ws.append(['1']); wb.save(p)
        r=O.xlsx_delete_rows(str(p),dict(row_mode='dup',out_dir=str(self.out)))
        self.assertEqual(r.status,'skipped')
        wb.close()

    def test_rename_rejects_path_escape(self):
        p=self.root/'a.txt'; p.write_text('keep')
        with self.assertRaises(ValueError):
            ops.build_rename_plan([str(p)],dict(mode='prefix',prefix='../'))

    def test_excel_split_preserves_styles_images_and_formulas(self):
        p=self.root/'styled.xlsx'; wb=openpyxl.Workbook(); ws=wb.active
        ws['A1']='header'; ws['A1'].font=openpyxl.styles.Font(bold=True,color='FF0000')
        ws['A2']=10; ws['B2']='=A2*2'; ws.merge_cells('C1:D1'); ws.column_dimensions['A'].width=29
        image=self.root/'img.png'; Image.new('RGB',(20,20),'red').save(image)
        ws.add_image(openpyxl.drawing.image.Image(str(image)),'E1'); wb.save(p); wb.close()
        r=A.xlsx_split(str(p),dict(out_dir=str(self.out),xlsx_split_mode='sheet'))
        result=openpyxl.load_workbook(r.output); sheet=result.active
        self.assertTrue(sheet['A1'].font.bold); self.assertEqual(sheet['B2'].value,'=A2*2')
        self.assertEqual(sheet.column_dimensions['A'].width,29)
        self.assertEqual(len(sheet._images),1); self.assertIn('C1:D1',sheet.merged_cells)
        result.close()

    def test_excel_merge_preserves_styles_and_cross_sheet_formulas(self):
        p=self.root/'book.xlsx'; wb=openpyxl.Workbook(); ws=wb.active; ws.title='Data'
        ws['A1']=12; ws['A1'].number_format='0.00'; ws['A1'].font=openpyxl.styles.Font(bold=True)
        wb.create_sheet('Summary')['A1']="='Data'!A1*2"; wb.save(p); wb.close()
        target=self.out/'merged.xlsx'; A.xlsx_merge_files([str(p)],str(target))
        result=openpyxl.load_workbook(target)
        self.assertTrue(result['book_Data']['A1'].font.bold)
        self.assertEqual(result['book_Data']['A1'].number_format,'0.00')
        self.assertEqual(result['book_Summary']['A1'].value,"='book_Data'!A1*2")
        result.close()

    def test_word_merge_preserves_images_and_styles(self):
        from engine.merge_ops import merge_files
        image=self.root/'img.png'; Image.new('RGB',(20,20),'red').save(image)
        files=[]
        for i in range(2):
            p=self.root/f'file{i}.docx'; d=docx.Document(); d.add_paragraph(f'title{i}').runs[0].bold=True
            d.add_picture(str(image)); d.save(p); files.append(str(p))
        r=merge_files(files[0],dict(files=files,type='Word 合并(追加段落)',destination=str(self.out/'combined.docx')))
        result=docx.Document(r.output)
        self.assertEqual(len(result.inline_shapes),2)
        self.assertTrue(all(p.runs[0].bold for p in result.paragraphs if p.text.startswith('title')))

    def test_atomic_failure_keeps_destination(self):
        from engine.safety import atomic_save
        target=self.root/'original.docx'; target.write_bytes(b'original')
        def failing_save(path):
            Path(path).write_bytes(b'partial'); raise ValueError('simulated failure')
        with self.assertRaises(ValueError): atomic_save(failing_save,str(target))
        self.assertEqual(target.read_bytes(),b'original')

    def test_aes256_pdf_decryption_and_wrong_password(self):
        import pymupdf
        p=self.root/'aes.pdf'; document=pymupdf.open()
        document.new_page().insert_text((72,72),'AES contents')
        document.save(p,encryption=pymupdf.PDF_ENCRYPT_AES_256,user_pw='secret',owner_pw='owner')
        document.close(); original=p.read_bytes()
        wrong=O.pdf_ops(str(p),dict(pdf_op='decrypt',password='wrong',out_dir=str(self.out)))
        self.assertEqual(wrong.status,'fail');self.assertEqual(p.read_bytes(),original)
        result=O.pdf_ops(str(p),dict(pdf_op='decrypt',password='secret',out_dir=str(self.out)))
        self.assertEqual(result.status,'ok',result.detail)
        with pymupdf.open(result.output) as reopened:
            self.assertFalse(reopened.needs_pass)
            self.assertIn('AES contents',reopened[0].get_text())

    def test_rename_excel_rejects_path_escape(self):
        p=self.root/'a.txt'; p.write_text('keep')
        r=ops.rename_by_excel_map(str(p),{'a.txt':'../escape.txt'})
        self.assertEqual(r.status,'fail'); self.assertTrue(p.exists())

    def payload(self, name='payload.zip', valid=True):
        p=self.root/name
        with zipfile.ZipFile(p,'w') as z:
            z.writestr('青雁批量工具箱v3.exe' if valid else 'missing.txt', b'new program')
            z.writestr('_internal/data.txt', b'resource')
        return str(p)

    def test_cancelled_upgrade_preserves_old_install(self):
        target=self.root/'installed'; target.mkdir(); old=target/'青雁批量工具箱v3.exe'; old.write_bytes(b'old')
        with patch.object(installer,'create_shortcuts'):
            rc=installer.do_install(self.payload(),str(target),{'exe':old.name},should_cancel=lambda:True)
        self.assertEqual(rc,-1); self.assertEqual(old.read_bytes(),b'old')

    def test_invalid_payload_does_not_modify_existing_install(self):
        target=self.root/'installed'; target.mkdir(); keep=target/'keep.txt'; keep.write_text('keep')
        with patch.object(installer,'create_shortcuts'):
            rc=installer.do_install(self.payload(valid=False),str(target),{'exe':'青雁批量工具箱v3.exe'})
        self.assertNotEqual(rc,0); self.assertEqual(sorted(x.name for x in target.iterdir()),['keep.txt'])

    def test_install_zip_rejects_path_traversal(self):
        p=self.root/'bad.zip'
        with zipfile.ZipFile(p,'w') as z: z.writestr('../outside.txt','bad')
        with self.assertRaises(ValueError): installer.extract_payload(str(p),str(self.root/'target'))

    def test_failed_commit_restores_all_old_program_bytes(self):
        target=self.root/'installed'; target.mkdir()
        exe=target/'青雁批量工具箱v3.exe'; exe.write_bytes(b'old exe')
        old=target/'_internal';old.mkdir();(old/'data.txt').write_bytes(b'old data')
        original_replace=os.replace
        def fail_once(src,dst):
            if '.qingyan-stage-' in str(src) and str(dst).endswith('data.txt'):
                raise PermissionError('simulated file in use')
            return original_replace(src,dst)
        with patch.object(installer.os,'replace',side_effect=fail_once):
            rc=installer.do_install(self.payload(),str(target),{'exe':exe.name,'shortcuts':False})
        self.assertEqual(rc,1);self.assertEqual(exe.read_bytes(),b'old exe')
        self.assertEqual((old/'data.txt').read_bytes(),b'old data')

    def test_upgrade_removes_only_obsolete_owned_files(self):
        target=self.root/'installed'
        meta={'exe':'青雁批量工具箱v3.exe','shortcuts':False}
        self.assertEqual(installer.do_install(self.payload(),str(target),meta),0)
        keep=target/'personal.txt';keep.write_text('keep')
        new=self.root/'next.zip'
        with zipfile.ZipFile(new,'w') as z:z.writestr(meta['exe'],b'next')
        self.assertEqual(installer.do_install(str(new),str(target),meta),0)
        self.assertFalse((target/'_internal'/'data.txt').exists())
        self.assertEqual(keep.read_text(),'keep')

    def test_uninstall_unrecognized_directory_is_refused(self):
        target=self.root/'ordinary'; target.mkdir(); keep=target/'keep.txt'; keep.write_text('keep')
        with patch.object(installer,'remove_shortcuts'), patch.object(installer.subprocess,'run'):
            with self.assertRaises(ValueError): installer.uninstall(str(target))
        self.assertEqual(keep.read_text(),'keep')

    def test_uninstall_keeps_non_application_files(self):
        target=self.root/'installed'
        with patch.object(installer,'create_shortcuts'),patch.object(installer,'remove_shortcuts'):
            self.assertEqual(installer.do_install(self.payload(),str(target),{'exe':'青雁批量工具箱v3.exe'}),0)
            keep=target/'my-notes.txt'; keep.write_text('keep')
            installer.uninstall(str(target))
        self.assertEqual(keep.read_text(),'keep')


if __name__=='__main__': unittest.main(verbosity=2)
