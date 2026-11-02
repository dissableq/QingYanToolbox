"""Check document content and the actual step-editor window introduced in 3.3."""
import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
from pathlib import Path
import tempfile
import unittest
import zipfile
import posixpath
from lxml import etree as ET
import docx
from pptx import Presentation
from PySide6.QtWidgets import QApplication, QDialog
from PySide6.QtCore import QTimer
from app.main import MainWindow
from engine import advanced_ops as ops

APP = QApplication.instance() or QApplication([])


def check_relationships(test, path):
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        for name in names:
            if not name.endswith('.rels'):continue
            base = name.split('/_rels/')[0] if '/_rels/' in name else ''
            for relation in ET.fromstring(archive.read(name)):
                if relation.get('TargetMode') == 'External':continue
                target = relation.get('Target', '')
                target = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))
                test.assertIn(target, names)


class Regressions(unittest.TestCase):
    def test_comments_remove_content_and_keep_document(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp)/'review.docx'
            document = docx.Document()
            run = document.add_paragraph('正文要保留').runs[0]
            run.bold = True
            document.add_comment(run, '批注正文需要删除', author='Reviewer')
            document.save(path)
            original = path.read_bytes()
            result = ops.docx_remove_comments(str(path), {'out_dir':str(Path(temp)/'out')})
            self.assertEqual(result.status, 'ok')
            output = docx.Document(result.output)
            self.assertEqual(output.paragraphs[0].text, '正文要保留')
            self.assertTrue(output.paragraphs[0].runs[0].bold)
            self.assertEqual(len(output.comments), 0)
            self.assertEqual(path.read_bytes(), original)
            check_relationships(self, result.output)

    def test_word_and_ppt_macros_produce_readable_regular_documents(self):
        for family, extension, maker, processor, opener in (
            ('word','.docm',docx.Document,ops.docx_remove_macros,docx.Document),
            ('ppt','.pptm',Presentation,ops.pptx_remove_macros,Presentation),
        ):
            with self.subTest(family=family), tempfile.TemporaryDirectory() as temp:
                path=Path(temp)/('macro'+extension)
                document=maker()
                if family=='word':document.add_paragraph('保留正文')
                else:document.slides.add_slide(document.slide_layouts[6])
                document.save(path)
                main = 'document' if family=='word' else 'presentation'
                with zipfile.ZipFile(path) as archive:parts={n:archive.read(n) for n in archive.namelist()}
                relation_name=family+'/_rels/'+main+'.xml.rels'
                tree=ET.fromstring(parts[relation_name])
                ET.SubElement(tree,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
                              Id='rIdAuditMacro',Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject',Target='vbaProject.bin')
                parts[relation_name]=ET.tostring(tree)
                parts[family+'/vbaProject.bin']=b'INERT TEST DATA'
                tree=ET.fromstring(parts['[Content_Types].xml'])
                ET.SubElement(tree,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                              PartName='/'+family+'/vbaProject.bin',ContentType='application/vnd.ms-office.vbaProject')
                for item in tree:
                    if item.get('PartName')=='/'+family+'/'+main+'.xml':
                        item.set('ContentType','application/vnd.ms-'+('word.document' if family=='word' else 'powerpoint.presentation')+'.macroEnabled.main+xml')
                parts['[Content_Types].xml']=ET.tostring(tree)
                with zipfile.ZipFile(path,'w') as archive:
                    for name,data in parts.items():archive.writestr(name,data)
                result=processor(str(path),{'out_dir':str(Path(temp)/'out')})
                self.assertEqual(result.status,'ok')
                self.assertEqual(Path(result.output).suffix,'.docx' if family=='word' else '.pptx')
                opener(result.output)
                check_relationships(self,result.output)
                with zipfile.ZipFile(result.output) as archive:self.assertFalse(any('vba' in n.lower() for n in archive.namelist()))

    def test_step_editor_shows_parameters_and_cancel_restores_page(self):
        window=MainWindow();window.show();APP.processEvents()
        page=window.pages[5];pipeline=window.pages[17]
        window.goto_page(17);errors=[]
        def inspect():
            try:
                self.assertTrue(page.isVisible())
                self.assertTrue(page.file_edit.isVisible())
                self.assertTrue(page._widgets['op'].isVisible())
                self.assertTrue(page.action_bar.isHidden())
                self.assertTrue(page._splitter.widget(1).isHidden())
            except AssertionError as error:errors.append(error)
            APP.activeModalWidget().reject()
        QTimer.singleShot(30,inspect)
        dialog=pipeline._show_page_in_dialog(page,'参数窗口验证')
        self.assertEqual(dialog.result(),QDialog.Rejected)
        self.assertFalse(errors,errors)
        self.assertEqual(window.stack.indexOf(page),5)
        self.assertFalse(page.action_bar.isHidden())
        window.close()


if __name__=='__main__':unittest.main()
