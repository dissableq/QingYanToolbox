"""Remove Word comments, markers and package relationships as one operation."""
import posixpath
import zipfile
from lxml import etree as ET
from engine.core import FileResult
from engine.safety import atomic_save


def remove_comments(path, options):
    from engine.advanced_ops import _out
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    removed = {name for name in parts if name.startswith('word/') and
               posixpath.basename(name).lower().startswith(('comments', 'people.xml'))}
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    changed = len(removed)
    for name in removed:del parts[name]
    for name, data in list(parts.items()):
        if name.endswith('.rels'):
            tree = ET.fromstring(data)
            base = name.split('/_rels/')[0] if '/_rels/' in name else ''
            for relation in list(tree):
                target = relation.get('Target', '')
                target = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))
                if relation.get('TargetMode') != 'External' and target in removed:tree.remove(relation)
            parts[name] = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        elif name == '[Content_Types].xml':
            tree = ET.fromstring(data)
            for item in list(tree):
                if item.get('PartName', '').lstrip('/') in removed:tree.remove(item)
            parts[name] = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        elif name.startswith('word/') and name.endswith('.xml'):
            tree = ET.fromstring(data)
            for element in list(tree.iter(w+'commentRangeStart',w+'commentRangeEnd',w+'commentReference')):
                parent = element.getparent()
                if parent is not None:parent.remove(element);changed += 1
            parts[name] = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    if not changed:return FileResult(path, 'skipped', '未发现批注')
    out, undo = _out(path, options, '_nocmt')
    def save(destination):
        with zipfile.ZipFile(destination,'w',zipfile.ZIP_DEFLATED) as archive:
            for name,data in parts.items():archive.writestr(name,data)
    atomic_save(save,out)
    return FileResult(path,'ok','已删除批注正文、标记及关联信息',out,undo=undo)
