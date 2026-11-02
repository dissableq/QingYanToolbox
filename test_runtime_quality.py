"""Async pipeline, cancellation and malformed-file regressions using disposable data."""
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch
import docx
from PySide6.QtCore import QEventLoop, QTimer
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import QApplication, QMessageBox
from app.main import MainWindow
from engine.core import FileResult, TaskSpec, BatchWorker
from engine import office_ops as office

APP = QApplication.instance() or QApplication([])


def wait_until(predicate, timeout=15):
    end=time.monotonic()+timeout
    while not predicate():
        if time.monotonic()>end: raise AssertionError('任务未在限定时间内完成')
        loop=QEventLoop(); QTimer.singleShot(10,loop.quit); loop.exec()
    APP.processEvents()


class RuntimeQualityTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='qy_runtime_')
        self.root=Path(self.tmp.name)
        self.win=MainWindow()
        self.word=next(p for p in self.win.pages if getattr(p,'kind','')=='word')
        self.pipe=next(p for p in self.win.pages if getattr(p,'title','')=='任务流水线')
        self.patches=[patch.object(QMessageBox,n,return_value=QMessageBox.Yes)
                      for n in ('warning','information','question')]
        for p in self.patches:p.start()

    def tearDown(self):
        for p in self.win.pages:
            w=getattr(p,'worker',None)
            if w is not None and w.isRunning():w.cancel();w.wait(3000)
        APP.processEvents()
        from app.backup_dialog import BackupDialog
        with patch.object(BackupDialog, 'exec', return_value=BackupDialog.Accepted):self.win.close()
        for p in self.patches:p.stop()
        self.tmp.cleanup()

    def step(self,operation,**opts):
        self.word.restore_state({'op':operation,**opts})
        return {'page':self.word.title,'op':operation,'page_obj':self.word,
                'ui_state':self.word.capture_state()}

    def test_pipeline_split_passes_every_output_to_next_stage(self):
        path=self.root/'source.docx'; d=docx.Document()
        for i in range(3):d.add_paragraph('target '+str(i))
        d.save(path)
        self.pipe._steps=[self.step('按段落拆分',file_edit=str(path),docx_chunk=1,out_dir=str(self.root/'split')),
                          self.step('查找替换',find='target',replace='changed',out_dir=str(self.root/'final'))]
        self.pipe.run_pipeline();wait_until(lambda:self.pipe.run_btn.isEnabled())
        files=list((self.root/'final').glob('*.docx'))
        self.assertEqual(len(files),3,self.pipe.status.text())
        self.assertEqual(sorted(docx.Document(p).paragraphs[0].text for p in files),
                         ['changed 0','changed 1','changed 2'])

    def test_pipeline_keeps_unchanged_file_for_later_stage(self):
        path=self.root/'source.docx';d=docx.Document();d.add_paragraph('target');d.save(path)
        self.pipe._steps=[self.step('删除空白段落',file_edit=str(path),out_dir=str(self.root/'first')),
                          self.step('查找替换',find='target',replace='next',out_dir=str(self.root/'second'))]
        self.pipe.run_pipeline();wait_until(lambda:self.pipe.run_btn.isEnabled())
        files=list((self.root/'second').glob('*.docx'))
        self.assertEqual(len(files),1,self.pipe.status.text())
        self.assertEqual(docx.Document(files[0]).paragraphs[0].text,'next')

    def test_cancel_keeps_completed_results_and_skips_remaining(self):
        calls=[]
        def process(path, options):
            calls.append(path);worker.cancel();return FileResult(path,'ok')
        worker=BatchWorker(TaskSpec('test','test',['first','second']),process,str(self.root/'records'))
        worker.start();wait_until(lambda:not worker.isRunning())
        self.assertEqual(calls,['first']);self.assertEqual(len(worker.results),1)

    def test_no_to_close_preserves_running_window(self):
        class Busy:
            def isRunning(self):return True
            def cancel(self):raise AssertionError('不应取消')
        self.word.worker=Busy()
        try:
            with patch.object(QMessageBox,'question',return_value=QMessageBox.No):
                event=QCloseEvent();self.win.closeEvent(event)
            self.assertFalse(event.isAccepted())
        finally:self.word.worker=None

    def test_global_guard_covers_gap_between_pipeline_steps(self):
        self.win._task_running=True
        self.word._start_worker(TaskSpec('test','test',['file']),lambda p,o:FileResult(p))
        self.assertIsNone(self.word.worker)
        self.win._task_running=False

    def test_bad_office_files_fail_without_changing_bytes(self):
        for extension,processor,options in [('.docx',office.docx_replace,{'find':'x','replace':'y'}),
                                           ('.xlsx',office.xlsx_replace,{'find':'x','replace':'y'}),
                                           ('.pptx',office.pptx_replace,{'find':'x','replace':'y'})]:
            p=self.root/('broken'+extension);p.write_bytes(b'PK truncated')
            worker=BatchWorker(TaskSpec('test','test',[str(p)],options),processor,str(self.root/'records'))
            worker.run()
            self.assertEqual(worker.results[0].status,'fail')
            self.assertEqual(p.read_bytes(),b'PK truncated')


if __name__=='__main__':unittest.main(verbosity=2)
