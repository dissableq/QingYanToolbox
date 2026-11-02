# -*- coding: utf-8 -*-
"""打包环境冒烟测试。

把本文件放到 exe 同级(或用 --runtime-tmpdir 指向的临时目录)运行,
验证打包产物里的引擎功能真的可用 —— 尤其是 Excel / PDF 这类
延迟 import、静态分析容易漏掉的依赖。
"""
import os
import sys
import tempfile
import json


def main():
    d = tempfile.mkdtemp(prefix='qy_smoke_')
    results = []

    def chk(name, cond, detail=''):
        results.append((name, bool(cond), detail))
        print(('PASS ' if cond else 'FAIL ') + name +
              ((' :: ' + str(detail)) if not cond else ''))

    # ---- Excel (openpyxl)
    try:
        import openpyxl
        from engine import office_ops
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 1
        ws['A2'] = 2
        ws['A3'] = '=A1+A2'
        f = os.path.join(d, 'a.xlsx')
        wb.save(f)
        r = office_ops.xlsx_formulas_to_values(f, {'out_dir': d})
        chk('openpyxl 可用且公式转值不丢数据',
            r.status == 'skipped' and '缓存' in r.detail, r.status + r.detail)
        r = office_ops.xlsx_delete_rows(f, {'row_mode': 'blank',
                                           'out_dir': d})
        chk('Excel 删行可执行', r.status in ('ok', 'skipped'), r.status)
    except Exception as e:
        chk('openpyxl / Excel 功能', False, f'{type(e).__name__}: {e}')

    # ---- PDF (pypdf + reportlab + CJK 字体)
    try:
        from reportlab.pdfgen import canvas
        from pypdf import PdfReader
        from engine import office_ops
        p = os.path.join(d, 'a.pdf')
        c = canvas.Canvas(p)
        c.showPage()
        c.save()
        r = office_ops.pdf_ops(p, {'pdf_op': 'extract', 'page_from': '1',
                                   'page_to': '1', 'out_dir': d})
        chk('pypdf/reportlab 可用(PDF 提取)', r.status == 'ok',
            r.status + r.detail)
        r = office_ops.pdf_watermark(p, {'wm_text': '青雁水印',
                                         'out_dir': d})
        t = PdfReader(r.output).pages[0].extract_text()
        chk('PDF 中文水印正常(非 ■■)', '青雁' in t and '■' not in t,
            repr(t[:16]))
        r = office_ops.pdf_ops(p, {'pdf_op': 'encrypt', 'password': '123',
                                   'out_dir': d})
        chk('PDF 加密产物确实加密',
            r.status == 'ok' and PdfReader(r.output).is_encrypted, r.detail)
    except Exception as e:
        chk('pypdf / reportlab / PDF 功能', False, f'{type(e).__name__}: {e}')

    # ---- PyMuPDF (PDF 解析/渲染/表格线, 保真转换的底座)
    #    reportlab/pypdf 覆盖不到它: PDF→Word/Excel/HTML/PPT 全走 pymupdf,
    #    它缺了这四个方向会整片失效, 所以必须单独验。
    try:
        import pymupdf
        from engine import convert_ops
        pdf = os.path.join(d, 't.pdf')
        doc = pymupdf.open()
        pg = doc.new_page()
        pg.insert_text((72, 60), '标题测试', fontsize=18,
                       fontname='china-s')
        rows_src = (('名称', '数量', '单价'),
                    ('苹果', '3', '5.50'),
                    ('香蕉', '12', '3.25'))
        xs = (72, 160, 260, 360)
        # N 行数据要 N+1 条横线; 文字基线取行带中点偏下, 因为
        # insert_text 的 y 是基线, 词框约 [y-12, y], 中线必须落在
        # 两条横线之间才归位到本行
        top, rh = 90.0, 30.0
        ys = [top + i * rh for i in range(len(rows_src) + 1)]
        for y in ys:
            pg.draw_line(pymupdf.Point(xs[0], y), pymupdf.Point(xs[-1], y))
        for x in xs:
            pg.draw_line(pymupdf.Point(x, ys[0]), pymupdf.Point(x, ys[-1]))
        for ri, row in enumerate(rows_src):
            for ci, t in enumerate(row):
                pg.insert_text((xs[ci] + 4, (ys[ri] + ys[ri + 1]) / 2 + 4),
                               t, fontsize=10, fontname='china-s')
        doc.save(pdf)
        doc.close()
        back = pymupdf.open(pdf)[0].get_text()
        chk('PyMuPDF 可打开自建 PDF 且 CJK 文本可回读',
            '标题测试' in back, repr(back[:16]))
        aes = os.path.join(d, 'aes.pdf')
        with pymupdf.open(pdf) as encrypted:
            encrypted.save(aes, encryption=pymupdf.PDF_ENCRYPT_AES_256,
                           user_pw='secret', owner_pw='owner')
        decrypted = office_ops.pdf_ops(aes, {'pdf_op':'decrypt', 'password':'secret', 'out_dir':d})
        chk('AES-256 PDF 解密组件完整', decrypted.status == 'ok', decrypted.detail)
        r = convert_ops.convert_pdf_to_word(pdf, {'out_dir': d})
        chk('PyMuPDF 可用(PDF→Word 可编辑)', r.status == 'ok', r.detail)
        r = convert_ops.convert_pdf_to_excel(pdf, {'out_dir': d})
        chk('PyMuPDF 可用(PDF→Excel 表格识别)', r.status == 'ok', r.detail)
        import openpyxl
        rows = list(openpyxl.load_workbook(r.output).worksheets[0].
                    iter_rows(values_only=True))
        chk('PDF 表格文字归位且数字可求和',
            bool(rows) and rows[0][0] == '名称' and rows[-1][1] == 12,
            str(rows))
    except Exception as e:
        chk('PyMuPDF / PDF 保真转换', False, f'{type(e).__name__}: {e}')

    # ---- PPT (python-pptx + 克隆关系)
    try:
        from pptx import Presentation
        from engine import advanced_ops
        pr = Presentation()
        pr.slides.add_slide(pr.slide_layouts[5]).shapes.title.text = 'T'
        pp = os.path.join(d, 'a.pptx')
        pr.save(pp)
        n = advanced_ops.pptx_merge_files([pp], os.path.join(d, 'm.pptx'))
        chk('python-pptx 可用且合并成功', n == 1, str(n))
        chk('合并产物可重新打开',
            len(Presentation(os.path.join(d, 'm.pptx')).slides._sldIdLst) == 1)
    except Exception as e:
        chk('python-pptx / PPT 功能', False, f'{type(e).__name__}: {e}')

    # ---- Word (python-docx)
    try:
        import docx
        from engine import office_ops
        wd = os.path.join(d, 'a.docx')
        d1 = docx.Document()
        d1.add_paragraph('打包环境测试')
        d1.save(wd)
        r = office_ops.docx_watermark(wd, {'wm_text': 'WM', 'out_dir': d})
        chk('python-docx 可用且加水印', r.status == 'ok', r.detail)
        from engine.merge_ops import merge_files
        merged = os.path.join(d, 'merged.docx')
        result = merge_files(wd, {'files':[wd, wd], 'type':'Word 合并(追加段落)',
                                 'destination':merged})
        chk('Word 合并依赖及资源完整', result.status == 'ok' and
            sum(p.text == '打包环境测试' for p in docx.Document(merged).paragraphs) == 2)
    except Exception as e:
        chk('python-docx / Word 功能', False, f'{type(e).__name__}: {e}')

    # ---- 图片 (Pillow)
    try:
        from PIL import Image
        from engine import ops
        img = os.path.join(d, 'a.png')
        Image.new('RGB', (120, 80), (10, 100, 10)).save(img)
        r = ops.process_image(img, {'op': 'resize', 'width': 60,
                                    'height': 0, 'out_dir': d})
        chk('Pillow 可用且图片缩放成功', r.status == 'ok', r.detail)
    except Exception as e:
        chk('Pillow / 图片功能', False, f'{type(e).__name__}: {e}')

    # ---- 文本
    try:
        from engine import ops
        t1 = os.path.join(d, 'a.txt')
        with open(t1, 'w', encoding='utf-8') as fh:
            fh.write('aa  \nbb  \n')
        r = ops.process_text(t1, {'op': 'trim', 'out_dir': d})
        chk('文本工具可用', r.status in ('ok', 'skipped'), r.status)
    except Exception as e:
        chk('文本工具', False, f'{type(e).__name__}: {e}')

    try:
        from pathlib import Path
        from engine import unified_replace, backups
        from engine.undo import undo_actions
        folder=Path(d)/'统一任务';folder.mkdir();p=folder/'旧名.txt';p.write_text('旧名保留',encoding='utf-8')
        opts=dict(root=str(folder),rules=[dict(find='旧名',replace='新名')],contents=True,file_names=True,folder_names=True,recursive=True)
        plan=unified_replace.make_plan(opts);task_results,actions,_,_=unified_replace.execute(plan)
        chk('统一任务内容与名称替换', (folder/'新名.txt').read_text(encoding='utf-8')=='新名保留' and all(r.status=='ok' for r in task_results))
        chk('备份目录创建', (folder/'备份').is_dir())
        _,remaining,errors=undo_actions(actions)
        chk('统一任务撤销恢复',not remaining and not errors and p.read_text(encoding='utf-8')=='旧名保留')
        backups.cleanup({r['id'] for r in backups.existing() if r['path'].startswith(d+os.sep)})
    except Exception as exc:chk('统一任务与备份',False,str(exc))

    if os.environ.get('QINGYAN_SMOKE_GUI'):
        try:
            from PySide6.QtWidgets import QApplication
            from PySide6.QtGui import QFontDatabase
            from app.main import MainWindow
            app = QApplication.instance() or QApplication([])
            for font in ('msyh.ttc', 'msyhbd.ttc'):
                QFontDatabase.addApplicationFont(os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', font))
            window = MainWindow(); window.resize(1320, 860); window.show()
            for index in range(len(window.pages)):
                window.goto_page(index); app.processEvents()
            chk('安装版 22 个页面创建并切换', len(window.pages) == 22)
            window.goto_page(0); app.processEvents()
            chk('安装版首页渲染', window.grab().save(os.environ['QINGYAN_SMOKE_GUI']))
            window.apply_theme('dark', '#315CF5', 1.25)
            app.processEvents(); window.close()
            chk('安装版主题和字号切换', True)
        except Exception as exc:
            chk('安装版界面', False, str(exc))
    passed = sum(1 for _n, ok, _d in results if ok)
    print(f'\nSMOKE RESULT: {passed} passed, {len(results) - passed} failed')
    report = os.environ.get('QINGYAN_SMOKE_REPORT')
    if report:
        with open(report, 'w', encoding='utf-8') as fh:
            json.dump({'passed':passed, 'failed':len(results)-passed, 'checks':results},
                      fh, ensure_ascii=False, indent=2)
    return 0 if passed == len(results) else 1


if __name__ == '__main__':
    sys.exit(main())
