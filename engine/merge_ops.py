"""Non-GUI merge operations, suitable for BatchWorker."""
import os
from engine.core import FileResult
from engine.safety import atomic_save
from engine.ops import _uniq_path, read_text_auto


def merge_files(_path, options):
    files = options['files']; kind = options['type']; out = options['destination']
    if not files:
        raise ValueError('请选择文件')
    out = _uniq_path(out) if kind != 'PDF 每 N 页拆分' else out
    if kind == 'Word 合并(追加段落)':
        from docx import Document
        from docxcompose.composer import Composer
        composer = Composer(Document(files[0]))
        for path in files[1:]: composer.append(Document(path))
        atomic_save(composer.save, out)
    elif kind == 'Excel 合并(多 Sheet)':
        from engine.workbooks import merge_workbooks
        merge_workbooks(files, out)
    elif kind == 'PPT 合并':
        from engine.advanced_ops import pptx_merge_files
        pptx_merge_files(files, out)
    elif kind == 'PDF 合并':
        from engine.office_ops import pdf_merge_files
        pdf_merge_files(files, out)
    elif kind == 'PDF 每 N 页拆分':
        from pypdf import PdfReader, PdfWriter
        outputs = []
        for path in files:
            reader = PdfReader(path)
            chunk = max(1, int(options['chunk']))
            for i in range(0, len(reader.pages), chunk):
                writer = PdfWriter()
                for page in reader.pages[i:i+chunk]: writer.add_page(page)
                target = _uniq_path(os.path.join(out, os.path.splitext(os.path.basename(path))[0]+f'_part{i//chunk+1:03d}.pdf'))
                atomic_save(writer.write, target); outputs.append(target)
        return FileResult(files[0], 'ok', f'已拆分为 {len(outputs)} 份', out, outputs=outputs)
    elif kind == '文本合并':
        def save(path):
            with open(path,'w',encoding='utf-8',newline='') as f:
                f.write('\n'.join(read_text_auto(path) for path in files))
        atomic_save(save, out)
    elif kind in ('图片拼接长图','图片合并 PDF'):
        from engine.advanced_ops import image_merge_files
        image_merge_files(files,out,'长图' if kind=='图片拼接长图' else 'PDF')
    else:
        raise ValueError('不支持的合并方式')
    return FileResult(files[0],'ok',f'已处理 {len(files)} 个文件',out)
