"""Exercise every UI operation with disposable representative inputs and reopen outputs."""
import json
import os
from pathlib import Path
import shutil
import tempfile
import time
import zipfile
from unittest.mock import patch

os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
from PySide6.QtCore import QEventLoop, QTimer
from PySide6.QtWidgets import QApplication, QLineEdit, QPlainTextEdit, QSpinBox, QComboBox, QMessageBox
import docx
import openpyxl
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import pymupdf
from app.main import MainWindow
from app.pages.base import FunctionPage

app=QApplication.instance() or QApplication([])
window=MainWindow()
reports=[]
messages=[]

def dialog(*args,**kwargs):
    messages.append(str(args[2]) if len(args)>2 else str(args)); return QMessageBox.Yes

def wait(page):
    worker=getattr(page,'worker',None)
    if worker is None:return
    loop=QEventLoop(); timer=QTimer()
    timer.timeout.connect(lambda:loop.quit() if not worker.isRunning() else None)
    timer.start(20); QTimer.singleShot(90000,loop.quit); loop.exec(); timer.stop()
    if worker.isRunning():
        worker.cancel(); worker.wait(); raise TimeoutError('task exceeded 90 seconds')
    app.processEvents()

def check_output(path):
    p=Path(path)
    if not p.is_file():return
    ext=p.suffix.lower()
    if ext=='.docx': docx.Document(p)
    elif ext in ('.xlsx','.xlsm'):
        wb=openpyxl.load_workbook(p); wb.close()
    elif ext=='.pptx': Presentation(p)
    elif ext=='.pdf':
        with pymupdf.open(p) as document:
            if not document.is_encrypted:
                assert len(document)>0
                document[0].get_pixmap(matrix=pymupdf.Matrix(.2,.2))
    elif ext in ('.png','.jpg','.jpeg','.bmp','.webp','.tif','.gif'):
        with Image.open(p) as image:image.verify()
    elif ext=='.zip':
        with zipfile.ZipFile(p) as z:assert z.testzip() is None
    elif ext=='.json':json.loads(p.read_text(encoding='utf-8-sig'))

with tempfile.TemporaryDirectory(prefix='qingyan_surface_') as tmp:
    root=Path(tmp); samples=root/'samples'; samples.mkdir()
    Image.new('RGB',(100,80),'#2463AA').save(samples/'picture.png')
    d=docx.Document(); d.add_heading('青雁测试 target',0); d.add_paragraph('target 保留格式').runs[0].bold=True
    d.add_paragraph(''); d.add_paragraph('target 正文'); d.add_picture(str(samples/'picture.png'))
    d.add_table(rows=2,cols=2).cell(0,0).text='target 表格'
    d.sections[0].header.paragraphs[0].text='target 页眉'; d.save(samples/'sample.docx')
    wb=openpyxl.Workbook(); ws=wb.active; ws.title='Sheet1'; ws.append(['name','value'])
    ws.append(['target',12]); ws.append(['target',12]); ws.append([None,None]); ws.append(['last',22])
    ws['A1'].font=openpyxl.styles.Font(bold=True); wb.create_sheet('Empty'); wb.save(samples/'sample.xlsx');wb.close()
    prs=Presentation()
    for i in range(3):
        slide=prs.slides.add_slide(prs.slide_layouts[6])
        if i<2:
            slide.shapes.add_textbox(Inches(1),Inches(1),Inches(7),Inches(1)).text='target 测试'
            slide.shapes.add_picture(str(samples/'picture.png'),Inches(1),Inches(2))
    prs.save(samples/'sample.pptx')
    pdf=pymupdf.open()
    for i in range(3):
        page=pdf.new_page()
        if i<2:
            page.insert_text((72,72),'target sample')
            for x in (72,180,280):page.draw_line((x,100),(x,160))
            for y in (100,130,160):page.draw_line((72,y),(280,y))
            for x,y,text in ((80,120,'Name'),(190,120,'Value'),(80,150,'target'),(190,150,'12')):page.insert_text((x,y),text)
            page.insert_image(pymupdf.Rect(72,200,172,280),filename=str(samples/'picture.png'))
    pdf.save(samples/'sample.pdf');pdf.close()
    (samples/'sample.txt').write_bytes('target 测试\r\n\r\n重复\r\n重复\r\n'.encode('utf-8'))
    (samples/'sample.md').write_text('# target 标题\n\n**格式** 内容',encoding='utf-8')
    (samples/'sample.html').write_text('<h1>target</h1><p>内容</p>',encoding='utf-8')
    (samples/'sample.json').write_text('[{"name":"target","value":12}]',encoding='utf-8')
    (samples/'sample.csv').write_text('name,value\ntarget,12\n',encoding='utf-8')
    with zipfile.ZipFile(samples/'sample.zip','w') as z:z.writestr('target.txt','content')

    for page in window.pages:
        if not isinstance(page,FunctionPage):continue
        combo=page._op_widget()
        operations=[combo.itemText(i) for i in range(combo.count())] if combo else list(page.OPS)
        for index,op in enumerate(operations):
            case=root/f'{page.kind}_{index}'; src=case/'input'; out=case/'output'
            src.mkdir(parents=True);out.mkdir()
            messages.clear(); page.worker=None
            if combo:combo.setCurrentText(op)
            # Independent source files for every operation.
            ext={'word':'.docx','excel':'.xlsx','ppt':'.pptx','pdf':'.pdf','text':'.txt',
                 'image':'.png','rename':'.txt','files':'.txt','print':'.txt','template':'.docx'}.get(page.kind,'.pdf')
            sample='picture.png' if ext=='.png' else 'sample'+ext
            if page.kind=='text':
                for token,suffix in [('JSON','.json'),('Markdown','.md'),('HTML→','.html')]:
                    if token in op:sample='sample'+suffix
            if page.kind=='convert':
                from engine.convert_ops import supported_exts
                key=page.collect_options()['convert']; supported=supported_exts(key)
                sample=next((f for f in ['sample.docx','sample.xlsx','sample.pptx','sample.csv','sample.html','sample.txt','sample.pdf','picture.png']
                             if Path(f).suffix in supported),'sample.pdf')
            if page.kind=='files' and '解压' in op:sample='sample.zip'
            source=src/sample;shutil.copy2(samples/sample,source)
            if page.kind=='files':
                (src/'empty.txt').write_bytes(b'');(src/'empty-folder').mkdir()
            if page.kind=='rename' and '文件夹' in op:
                (src/'folder').mkdir()
            wb=openpyxl.Workbook(); ws=wb.active; ws.append(['原','新']);ws.append(['target','changed']);rules=case/'rules.xlsx';wb.save(rules);wb.close()
            wb=openpyxl.Workbook(); ws=wb.active;ws.append(['path','target','type']);ws.append([str(source),str(out/'copied.txt'),'file']); listing=case/'list.xlsx';wb.save(listing);wb.close()
            wb=openpyxl.Workbook();ws=wb.active;ws.append(['name','value']);ws.append(['target','12']); data=case/'data.xlsx';wb.save(data);wb.close()
            values={'find':'target','replace':'changed','wm':'测试水印','wm_text':'测试水印','header':'测试页眉','footer':'页脚',
                    'bg_color':'#EFF4FF','meta_author':'Test Author','meta_title':'Test Title','sheet_names':'Empty',
                    'sheet_new_name':'Inserted','pages_spec':'2','seal_edit':str(samples/'picture.png'),'pwd':'secret',
                    'link_url':'https://example.invalid','crop_margin':'5,5,5,5','rules_edit':str(rules),
                    'lines_spec':'2','line_match':'重复','line_new':'changed','insert_text':'inserted','prefix':'new_',
                    'suffix':'_done','new_ext':'.log','excel_edit':str(rules),'excel_file':str(listing),'merge_target':str(out),
                    'out_dir':str(out),'pdf_out_dir':str(out),'tpl_edit':str(source),'data_edit':str(data),'name_col':'name',
                    'fixed_time':'2026-10-01 12:00:00','ts_fixed':'2026-10-01 12:00:00','targets':str(out),'copy_targets':str(out),'crop':'5,5,5,5','grid':'2,2','erase_rect':'0,0,20,20','erase_color':'#FFFFFF'}
            for name,value in values.items():
                w=page._widgets.get(name)
                if isinstance(w,QLineEdit):w.setText(value)
                elif isinstance(w,QPlainTextEdit):w.setPlainText(value)
            for name in ('docx_chunk','rows_per','p_chunk','chunk','lines_per','width','height'):
                w=page._widgets.get(name)
                if isinstance(w,QSpinBox):w.setValue(2)
            page.file_edit.setText(str(src) if page.kind in ('files','rename') else str(source))
            page.ext_edit.setText('');page.open_out_cb.setChecked(False)
            if page.kind=='print':
                printer=page._widgets['printer'];pos=printer.findText('Microsoft Print to PDF')
                if pos<0:
                    reports.append(dict(page=page.title,op=op,status='UNAVAILABLE',detail='virtual printer unavailable'));continue
                printer.setCurrentIndex(pos)
            start=time.monotonic()
            try:
                opts=page.collect_options(); proc=page.processor(); assert callable(proc)
                state=page.capture_state(); page.restore_state(state)
                preview=page.build_preview(page.current_files(),opts)
                if preview:page._split_preview(preview)
                with patch.object(QMessageBox,'question',dialog),patch.object(QMessageBox,'information',dialog),patch.object(QMessageBox,'warning',dialog):
                    page.do_run();wait(page)
                results=page.worker.results if page.worker else []
                fails=[r.detail for r in results if r.status=='fail']
                if not results:raise AssertionError('未执行: '+'; '.join(messages))
                if fails:raise AssertionError('; '.join(fails))
                for path in out.rglob('*'):check_output(path)
                if source.exists():check_output(source)
                reports.append(dict(page=page.title,op=op,status='PASS',ok=sum(r.status=='ok' for r in results),skipped=sum(r.status=='skipped' for r in results),seconds=round(time.monotonic()-start,2)))
            except Exception as e:
                reports.append(dict(page=page.title,op=op,status='FAIL',detail=type(e).__name__+': '+str(e)))
            print(json.dumps(reports[-1],ensure_ascii=False),flush=True)
    window.close()

report_path=Path(os.environ['APPDATA']).parent/'surface_results.json'
report_path.write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')
failures=[r for r in reports if r['status']=='FAIL']
print(f'RESULT: {len(reports)-len(failures)} checked, {len(failures)} failed')
raise SystemExit(bool(failures))
