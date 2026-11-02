"""Shared file and OOXML safeguards."""
from bisect import bisect_left, bisect_right
import os
import re
import tempfile
import time


def backup_file(path):
    from engine.backups import create_backup
    return create_backup(path)


def atomic_save(save, destination):
    """Only replace the destination after the serializer has finished."""
    destination = os.path.abspath(destination)
    fd, tmp = tempfile.mkstemp(prefix='.qingyan_', suffix=os.path.splitext(destination)[1],
                               dir=os.path.dirname(destination))
    os.close(fd)
    try:
        save(tmp)
        owner = getattr(save, '__self__', None)
        archive = getattr(owner, 'vba_archive', None)
        if archive is not None:
            archive.close()
        for attempt in range(8):
            try:
                os.replace(tmp, destination)
                break
            except PermissionError:
                if attempt == 7:
                    raise
                time.sleep(0.05 * (attempt + 1))
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def validate_filename(name):
    if (not isinstance(name, str) or not name or name in ('.', '..')
            or re.search(r'[<>:"/\\|?*\x00-\x1f]', name)
            or name.endswith((' ', '.'))
            or re.fullmatch(r'(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])',
                            name.split('.')[0], re.I)):
        raise ValueError('文件名包含无效字符、路径或系统保留名称: ' + str(name))
    return name


def replace_text_nodes(nodes, find, replacement, regex=False):
    """Edit matched text ranges only. Keep surrounding runs and drawings intact."""
    if not find:
        return 0
    nodes = list(nodes)
    texts = [n.text or '' for n in nodes]
    full = ''.join(texts)
    pattern = re.compile(find if regex else re.escape(find))
    matches = list(pattern.finditer(full))
    offsets = []
    pos = 0
    for text in texts:
        offsets.append(pos)
        pos += len(text)
    for match in reversed(matches):
        start, end = match.span()
        if not nodes:
            break
        first = max(0, bisect_right(offsets, start) - 1)
        last = max(first, bisect_right(offsets, max(start, end - 1)) - 1)
        a = nodes[first].text or ''
        b = nodes[last].text or ''
        if first == last:
            nodes[first].text = a[:start-offsets[first]] + replacement + a[end-offsets[first]:]
        else:
            nodes[first].text = a[:start-offsets[first]] + replacement
            for i in range(first + 1, last):
                nodes[i].text = ''
            nodes[last].text = b[end-offsets[last]:]
        for i in range(first, last + 1):
            nodes[i].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return len(matches)


def update_row_references(wb, deletions):
    """Map A1 references after row removal, including references on other sheets."""
    from openpyxl.formula import Tokenizer
    cell_re = re.compile(r'(\$?[A-Za-z]{1,3})(\$?)([1-9][0-9]*)$')

    def reference(value, current_sheet):
        prefix, sep, area = value.rpartition('!')
        sheet = prefix.strip("'").replace("''", "'") if sep else current_sheet
        if '[' in sheet or ':' in sheet:
            return value
        rows = deletions.get(sheet)
        if not rows:
            return value
        parts = (area if sep else value).split(':')
        cells = [cell_re.fullmatch(p) for p in parts]
        if len(parts) > 2 or not all(cells):
            return value
        mapped = []
        for i, m in enumerate(cells):
            row = int(m[3])
            if len(cells) == 1 and row in rows:
                return (prefix + '!' if sep else '') + '#REF!'
            new = row - (bisect_left(rows, row) if i == 0 else bisect_right(rows, row))
            mapped.append((m[1] + m[2], new))
        if len(mapped) == 2 and mapped[1][1] < mapped[0][1]:
            return (prefix + '!' if sep else '') + '#REF!'
        return (prefix + '!' if sep else '') + ':'.join(col + str(row) for col, row in mapped)

    for ws in wb.worksheets:
        for row in ws:
            for cell in row:
                if cell.data_type != 'f' or not isinstance(cell.value, str):
                    continue
                tokens = Tokenizer(cell.value).items
                cell.value = '=' + ''.join(reference(t.value, ws.title)
                    if t.type == 'OPERAND' and t.subtype == 'RANGE' else t.value for t in tokens)
