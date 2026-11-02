"""Remove VBA package parts together with their relationships and content types."""
import os
import posixpath
import zipfile
from lxml import etree as ET
from engine.core import FileResult
from engine.safety import atomic_save


def remove_macros(path, options, family):
    from engine.advanced_ops import _out, _uniq_out
    try:
        with zipfile.ZipFile(path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
    except (zipfile.BadZipFile, OSError) as error:
        return FileResult(path, 'fail', '无法读取 Office 文件：' + str(error))
    removed = {name for name in parts if 'vba' in name.lower()}
    if not removed:
        return FileResult(path, 'skipped', '未发现宏')
    for name in removed:
        del parts[name]
    regular_type = {
        'word': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        'ppt': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    }[family]
    for name, data in list(parts.items()):
        if name.endswith('.rels'):
            tree = ET.fromstring(data)
            base = name.split('/_rels/')[0] if '/_rels/' in name else ''
            for relation in list(tree):
                target = relation.get('Target', '')
                resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))
                if 'vba' in relation.get('Type', '').lower() or (relation.get('TargetMode') != 'External' and resolved in removed):
                    tree.remove(relation)
            parts[name] = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        elif name == '[Content_Types].xml':
            tree = ET.fromstring(data)
            for item in list(tree):
                content_type = item.get('ContentType', '')
                if item.get('PartName', '').lstrip('/') in removed or 'vbaproject' in content_type.lower():
                    tree.remove(item)
                elif 'macroEnabled.main+xml' in content_type:
                    item.set('ContentType', regular_type)
            parts[name] = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    regular_ext = '.docx' if family == 'word' else '.pptx'
    if os.path.splitext(path)[1].lower() != regular_ext:
        out_dir = options.get('out_dir') or os.path.dirname(os.path.abspath(path))
        os.makedirs(out_dir, exist_ok=True)
        stem = os.path.splitext(os.path.basename(path))[0]
        out = _uniq_out(os.path.join(out_dir, stem + '_nomacro' + regular_ext))
        undo = []
    else:
        out, undo = _out(path, options, '_nomacro')
    def save(destination):
        with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
    atomic_save(save, out)
    return FileResult(path, 'ok', f'已移除宏及关联，保存为无宏 {regular_ext} 文件', out, undo=undo)
