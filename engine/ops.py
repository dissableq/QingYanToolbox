# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 各批量功能的具体处理器 (文件/文本/图片/重命名)。

每个函数签名统一: (path, options) -> FileResult。
全部离线: 不发起任何网络请求。
预览模式: options 里带 "_preview_plan" 列表时, 处理器只返回计划不落盘。
v3 增强: 移动/重命名/删除类操作在 FileResult.undo 携带撤销动作,
任务结束后自动写撤销清单, 台账页可一键撤销。
"""
from __future__ import annotations

from engine.safety import validate_filename, backup_file
from engine.backups import walk_sources

import csv
import hashlib
import json
import os
import re
import shutil
import time
import zipfile
from collections import defaultdict
from datetime import date, datetime
from html import escape
from html.parser import HTMLParser

from engine.core import FileResult, collect_files
from engine.store import TrashManifest, data_dir

# --------------------------------------------------------------------- 常量
IMG_EXTS = frozenset({'.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tif',
                      '.tiff', '.gif'})
TXT_EXTS = frozenset({'.txt', '.md', '.csv', '.log', '.json', '.xml',
                      '.html', '.htm', '.srt', '.ini', '.yaml', '.yml',
                      '.css', '.js', '.py', '.bat'})
ENC_MAP = {'UTF-8': 'utf-8', 'GBK': 'gbk', 'GB18030': 'gb18030',
           'UTF-16': 'utf-16'}
CN_PUNCT_MAP = {',': '，', '.': '。', ';': '；', ':': '：',
                '?': '？', '!': '！'}
_CN_RE = re.compile('[\u4e00-\u9fff]')


def default_trash_dir() -> str:
    d = os.path.join(data_dir(), 'trash')
    os.makedirs(d, exist_ok=True)
    return d


def _now_stamp() -> str:
    return datetime.now().strftime('%Y%m%d_%H%M%S_%f')


def _uniq_path(p: str) -> str:
    """输出文件名冲突时追加 (1)(2), 避免静默覆盖已有文件。"""
    if not os.path.exists(p):
        return p
    b, e = os.path.splitext(p)
    k = 1
    while os.path.exists(f'{b}({k}){e}'):
        k += 1
    return f'{b}({k}){e}'


# ---------------------------------------------------------------- 编码探测
def detect_encoding(path: str, sample_size: int = 262144) -> str:
    """编码探测: BOM 检查 → 严格解码候选(utf-8/gb18030) → chardet。

    chardet 对短中文 GBK 文本常误判为 iso-8859-x, 故先做严格解码候选。
    gb18030 是 GBK/GB2312 的超集, 严格解码成功即可确信是简中编码。

    安全要求: 采样块按字节截断, 很可能正好切断一个多字节字符。此时
    utf-8 严格解码必然失败, 而 gb18030 却常能"成功"解出乱码 —— 于是
    UTF-8 文件被误判成 GBK, 后续读写成乱码。所以判定失败时要把截断
    处回退到字符边界再判, 仍判不出就宁可报 utf-8(与原来一致)并让调用
    方以 errors='replace' 兜底, 绝不能"自信地"给出错误编码。
    """
    try:
        with open(path, 'rb') as fh:
            raw = fh.read(sample_size)
            # 纯 ASCII 前缀是"合法 UTF-8", 直接采信会把后面整段 GBK 判成
            # UTF-8 并读成 U+FFFD。所以样本全 ASCII 时继续往后读, 直到
            # 出现非 ASCII 字节(判据才有信息量)或读够上限。
            if raw and not any(b > 0x7f for b in raw):
                extra = fh.read(4 * 1024 * 1024)
                if extra:
                    raw += extra
    except OSError:
        return 'utf-8'
    if not raw:
        return 'utf-8'
    if raw.startswith(b'\xef\xbb\xbf'):
        return 'utf-8-sig'
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return 'utf-16'

    def _strict_ok(buf: bytes, enc: str) -> bool:
        try:
            buf.decode(enc, errors='strict')
            return True
        except (UnicodeDecodeError, ValueError):
            return False

    for enc in ('utf-8', 'gb18030'):
        if _strict_ok(raw, enc):
            return enc
        # 采样块可能切在字符中间: 逐字节回退到前一个合法边界再试。
        # 只回退有限字节, 覆盖 UTF-8 最多 4 字节 / GB18030 最多 4 字节。
        for cut in range(1, min(5, len(raw))):
            if _strict_ok(raw[:-cut], enc):
                return enc
    # 两种候选都不行, 可能是 Latin-1/Big5 等。交给 chardet。
    try:
        import chardet
        r = chardet.detect(raw)
        enc = (r or {}).get('encoding')
        if enc:
            enc = enc.lower()
            if enc == 'ascii':
                return 'utf-8'
            return enc
    except Exception:
        pass
    return 'utf-8'


def read_text_auto(fp: str) -> str:
    """自动探测编码读取文本(复用 detect_encoding, 容错回退)。"""
    enc = None
    try:
        enc = detect_encoding(fp)
    except Exception:
        enc = None
    for e in (enc, 'utf-8', 'gbk'):
        if not e:
            continue
        try:
            with open(fp, 'r', encoding=e, errors='strict') as fh:
                return fh.read()
        except Exception:
            continue
    try:
        with open(fp, 'r', encoding='utf-8', errors='replace') as fh:
            return fh.read()
    except Exception:
        return ''


# ------------------------------------------------------------ 全角/半角
def _full_to_half(s: str) -> str:
    """全角字符(字母数字标点)转半角; 全角空格转半角空格。"""
    out = []
    for ch in s:
        code = ord(ch)
        if code == 12288:
            out.append(' ')
        elif 65281 <= code <= 65374:
            out.append(chr(code - 65248))
        else:
            out.append(ch)
    return ''.join(out)


def _half_to_full(s: str) -> str:
    """半角字母数字标点转全角(空格除外)。"""
    out = []
    for ch in s:
        code = ord(ch)
        if 33 <= code <= 126:
            out.append(chr(code + 65248))
        else:
            out.append(ch)
    return ''.join(out)


def _normalize_cn_punct(s: str) -> str:
    """英文标点 → 中文标点(仅当相邻为中文时, 避免破坏 URL/代码)。"""
    chars = list(s)
    n = len(chars)
    for i, ch in enumerate(chars):
        if ch not in CN_PUNCT_MAP:
            continue
        prev_cn = i > 0 and _CN_RE.search(chars[i - 1] or '')
        next_cn = i + 1 < n and _CN_RE.search(chars[i + 1] or '')
        if prev_cn or next_cn:
            chars[i] = CN_PUNCT_MAP[ch]
    return ''.join(chars)


def text_fullwidth_convert(path: str, options: dict) -> FileResult:
    """全角/半角转换 + 中文标点统一 (办公文本清理高频操作)。

    UI 传的是 op='fw_half'/'fw_full'/'fw_punct'(一个下拉项一个 op),
    早期这里却只读 fw_mode, 导致三种功能全部退化成"全角→半角"。
    现在两种键名都认: op 优先, 其次 fw_mode(兼容旧方案文件)。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext not in TXT_EXTS:
        return FileResult(path, 'skipped', '非支持文本格式')
    op = str(options.get('op', '') or '')
    by_op = {'fw_half': '半角', 'fw_full': '全角', 'fw_punct': '中文标点'}
    mode = by_op.get(op) or options.get('fw_mode', '半角')
    fn = {'半角': _full_to_half, '全角': _half_to_full,
          '中文标点': _normalize_cn_punct}.get(mode)
    if fn is None:
        return FileResult(path, 'fail', '未知转换模式 ' + str(mode))
    enc = detect_encoding(path)
    with open(path, 'r', encoding=enc, errors='replace') as fh:
        content = fh.read()
    new_content = fn(content)
    if new_content == content:
        return FileResult(path, 'skipped', '无需转换(内容未变化)')
    out_dir = options.get('out_dir')
    undo: list = []
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, os.path.basename(path))
        # 输出到源目录时等同于就地修改, 必须备份, 否则源文件被无提示覆盖
        if os.path.abspath(out) == os.path.abspath(path):
            bak = backup_file(path)
            undo.append({'src': bak, 'dst': path, 'kind': 'restore'})
        elif os.path.exists(out):
            out = _uniq_path(out)
    else:
        bak = backup_file(path)
        undo.append({'src': bak, 'dst': path, 'kind': 'restore'})
        out = path
    tmp = out + '.qytmp'
    try:
        with open(tmp, 'w', encoding=enc, newline='') as fh:
            fh.write(new_content)
        os.replace(tmp, out)
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        raise
    return FileResult(path, 'ok', '转换完成(' + mode + ')', out, undo=undo)


# ------------------------------------------------------------ 空文件清理
def cleanup_empty(path: str, options: dict) -> FileResult:
    """空文件夹/空文件清理: 递归扫描并移入回收备份区(可恢复)。

    options:
        mode: "empty_dirs" 空文件夹 / "empty_files" 空文件 / "both" 两者
        recursive: 多轮清理空文件夹(父目录因清空而变空也移除)
    """
    mode = options.get('empty_mode', 'both')
    trash_dir = options.get('trash_dir') or default_trash_dir()
    manifest = TrashManifest()

    def _to_trash(p: str) -> str:
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        dst = os.path.join(trash_dir, stamp + '_' + os.path.basename(p))
        shutil.move(p, dst)
        manifest.record(os.path.basename(dst), p)
        return dst

    undo: list = []
    if os.path.isfile(path):
        if os.path.getsize(path) == 0:
            dst = _to_trash(path)
            undo.append({'src': path, 'dst': dst})
            return FileResult(path, 'ok', '空文件已移入回收备份区', dst, undo=undo)
        return FileResult(path, 'skipped', '目标不是空文件')

    moved_files = moved_dirs = 0
    if mode in ('empty_files', 'both'):
        for dirpath, _dirs, files in walk_sources(path):
            for fn in files:
                fp = os.path.join(dirpath, fn)
                try:
                    if os.path.isfile(fp) and os.path.getsize(fp) == 0:
                        dst = _to_trash(fp)
                        undo.append({'src': fp, 'dst': dst})
                        moved_files += 1
                except OSError:
                    continue
    if mode in ('empty_dirs', 'both'):
        changed = True
        while changed:
            changed = False
            for dirpath, dirs, _files in walk_sources(path, topdown=False):
                for dn in dirs:
                    dp = os.path.join(dirpath, dn)
                    try:
                        if not os.listdir(dp):
                            dst = _to_trash(dp)
                            undo.append({'src': dp, 'dst': dst})
                            moved_dirs += 1
                            changed = True
                    except OSError:
                        continue

    if moved_files == 0 and moved_dirs == 0:
        return FileResult(path, 'skipped', '未发现空文件夹/空文件')
    detail = (f'清理空文件 {moved_files} 个, 空文件夹 {moved_dirs} 个'
              f'(均进回收备份区)')
    return FileResult(path, 'ok', detail, path, undo=undo)


# ------------------------------------------------------------ 重复文件
def _md5_of(fp: str) -> str:
    h = hashlib.md5()
    with open(fp, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1048576), b''):
            h.update(chunk)
    return h.hexdigest()


def find_duplicate_files(path: str, options: dict) -> FileResult:
    """重复文件查找: 按 大小→首块→MD5 三级分组, 输出 CSV 报告。

    options:
      auto_keep: bool  自动删除重复副本(每组保留最早修改的一份, 移入回收备份区)
      out_dir:   str   报告输出目录
    """
    root = path if os.path.isdir(path) else os.path.dirname(path)
    by_size: dict = defaultdict(list)
    for dirpath, _dirs, files in walk_sources(root):
        for fn in files:
            fp = os.path.join(dirpath, fn)
            try:
                by_size[os.path.getsize(fp)].append(fp)
            except OSError:
                continue

    candidates = [v for v in by_size.values() if len(v) > 1]
    by_head: dict = defaultdict(list)
    for group in candidates:
        for fp in group:
            try:
                with open(fp, 'rb') as fh:
                    head = fh.read(65536)
                by_head[(os.path.getsize(fp), hashlib.md5(head).hexdigest())].append(fp)
            except OSError:
                continue

    dup_groups: list = []
    for group in by_head.values():
        if len(group) < 2:
            continue
        by_md5: dict = defaultdict(list)
        for fp in group:
            try:
                by_md5[_md5_of(fp)].append(fp)
            except OSError:
                continue
        for g in by_md5.values():
            if len(g) > 1:
                dup_groups.append(sorted(g, key=os.path.getmtime))

    total_files = sum(len(g) - 1 for g in dup_groups)
    wasted = sum(sum(os.path.getsize(f) for f in g[1:]) for g in dup_groups)
    out_dir = options.get('out_dir') or (os.path.dirname(root) if os.path.isfile(path) else root)
    os.makedirs(out_dir, exist_ok=True)
    title = os.path.basename(root.rstrip('\\/')) or '重复文件'
    out = os.path.join(out_dir, title + '_重复文件报告.csv')
    with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
        w = csv.writer(fh)
        w.writerow(('重复组', '文件', '大小(字节)', '处理建议'))
        for gi, g in enumerate(dup_groups, 1):
            for i, fpath in enumerate(g):
                w.writerow((f'组{gi}', fpath, os.path.getsize(fpath),
                            '保留(最早)' if i == 0 else '重复'))

    cleaned = 0
    undo: list = []
    if options.get('auto_keep'):
        trash_dir = options.get('trash_dir') or default_trash_dir()
        manifest = TrashManifest()
        for g in dup_groups:
            for dup in g[1:]:
                try:
                    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
                    dst = os.path.join(trash_dir, stamp + '_' + os.path.basename(dup))
                    shutil.move(dup, dst)
                    manifest.record(os.path.basename(dst), dup)
                    undo.append({'src': dup, 'dst': dst})
                    cleaned += 1
                except Exception:
                    continue

    detail = (f'发现 {len(dup_groups)} 组重复, 共 {total_files} 个多余文件, '
              f'浪费 {wasted // 1024} KB')
    if options.get('auto_keep'):
        detail += f', 已自动清理 {cleaned} 个(进回收备份区)'
    return FileResult(path, 'ok', detail, out, undo=undo)


# ------------------------------------------------------------ 文件清单
def files_export_excel(path: str, options: dict) -> FileResult:
    """文件清单导出 Excel (办公用户友好: 含序号/文件名/类型/大小/修改时间)。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    root = path if os.path.isdir(path) else os.path.dirname(path)
    title = os.path.basename(root.rstrip('\\/')) or '文件清单'
    rows = []
    for dirpath, _dirs, files in sorted(walk_sources(root)):
        for fn in sorted(files):
            fp = os.path.join(dirpath, fn)
            try:
                st = os.stat(fp)
            except OSError:
                continue
            ext = os.path.splitext(fn)[1].lstrip('.').upper() or '无扩展名'
            rows.append((len(rows) + 1, fn, ext, round(st.st_size / 1024, 1),
                         datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M'),
                         fp))
    out_dir = options.get('out_dir') or root
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, title + '_文件清单.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = '文件清单'
    headers = ('序号', '文件名', '类型', '大小(KB)', '修改时间', '完整路径')
    ws.append(headers)
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='0F6E8C')
    for row in rows:
        ws.append(list(row))
    for col, w in (('A', 6), ('B', 32), ('C', 10), ('D', 10), ('E', 18), ('F', 60)):
        ws.column_dimensions[col].width = w
    wb.save(out)
    return FileResult(path, 'ok', f'清单 {len(rows)} 行 → {out}', out)


# ------------------------------------------------------------ 批量重命名
CLEANUP_PATTERNS = (
    '\\s*[-_ ]?副本\\d*\\s*$',
    '\\s*[-_ ]?最终版\\d*\\s*$',
    '(?i)\\s*[-_ ]?final\\d*\\s*$',
    '(?i)\\s*[-_ ]?new\\d*\\s*$',
    '\\s*[(（]\\d+[)）]\\s*$',
    '\\s*[-_ ]?\\d+\\s*$',
)


def _safe_new_name(dirpath: str, old: str, new: str) -> str:
    """避免与现有文件冲突; 冲突时追加序号。"""
    candidate = os.path.join(dirpath, new)
    if old and os.path.abspath(candidate) == os.path.abspath(old):
        return candidate
    i = 1
    base, ext = os.path.splitext(new)
    while os.path.exists(candidate):
        candidate = os.path.join(dirpath, f'{base}_{i}{ext}')
        i += 1
    return candidate


def build_rename_plan(files: list, options: dict) -> list:
    """返回 [(old_path, new_path, new_name), ...]

    模式: sequence / replace / prefix(前后缀) / add_date / ext(扩展名替换)
          / cleanup(冗余命名清理)

    计划阶段即做冲突消解: 目标名与磁盘现状或计划内其他条目撞名时,
    自动追加 _1/_2 序号, 保证计划可全部执行(v3 修复)。
    """
    mode = options.get('mode', 'sequence')
    seq_start = int(options.get('seq_start', 1) or 1)
    seq_pad = int(options.get('seq_pad', 3) or 3)
    find = options.get('find', '')
    replace = options.get('replace', '')
    prefix = options.get('prefix', '')
    suffix = options.get('suffix', '')
    ext_case = options.get('ext_case', 'keep')
    new_ext_raw = options.get('new_ext', '')
    plan: list = []
    # 计划内已占用的目标名(按目录分组), 初始含磁盘现状
    used: dict = {}
    for f in files:
        dirpath, old = os.path.split(f)
        used.setdefault(os.path.abspath(dirpath),
                        {os.path.normcase(n) for n in os.listdir(dirpath)})
    for f in files:
        dirpath, old = os.path.split(f)
        base, ext = os.path.splitext(old)
        new_base = base
        if mode == 'sequence':
            seq = seq_start + len(plan)
            # 前缀必须生效: UI 有"前缀"输入框且提示写明了"可加前缀",
            # 早期只有「文件夹序号编排」认, 文件序号编排默默丢掉。
            new_base = prefix + str(seq).zfill(seq_pad)
        elif mode == 'replace':
            new_base = base.replace(find, replace)
        elif mode == 'prefix':
            new_base = prefix + base + suffix
        elif mode == 'add_date':
            stamp = date.today().strftime('%Y%m%d')
            # UI 的后缀框对「加日期」可见且会收集, 早期只拼日期, 用户填的
            # 说明文字被静默丢弃 (suffix='备份' 得到 report_20261026.txt)。
            new_base = f'{base}_{stamp}{suffix}'
        elif mode == 'cleanup':
            new_base = base
            for pat in options.get('cleanup_patterns', CLEANUP_PATTERNS):
                new_base = re.sub(pat, '', new_base).strip()
            new_base = re.sub(r'\s*[-_ ]+$', '', new_base)
        elif mode == 'ext':
            if new_ext_raw:
                ne = new_ext_raw if new_ext_raw.startswith('.') else '.' + new_ext_raw
                new_base = base
                ext = ne
        if ext_case == 'lower':
            ext = ext.lower()
        elif ext_case == 'upper':
            ext = ext.upper()
        new_name = new_base + ext
        validate_filename(new_name)
        if not new_name or new_name == old:
            new_name = old
        # 计划内查重: 同目录下目标名已被其他条目/磁盘占用则追加序号
        taken = used.setdefault(os.path.abspath(dirpath), set())
        if new_name != old and os.path.normcase(new_name) in taken:
            i = 1
            nb, ne = os.path.splitext(new_name)
            while (os.path.normcase(f'{nb}_{i}{ne}') in taken):
                i += 1
            new_name = f'{nb}_{i}{ne}'
        new_path = os.path.abspath(os.path.join(dirpath, new_name))
        taken.discard(os.path.normcase(old))      # 原名腾出
        taken.add(os.path.normcase(new_name))     # 占用新名
        plan.append((f, new_path, new_name))
    return plan


def process_rename(path: str, options: dict) -> FileResult:
    plan = options.get('_rename_plan') or []
    for old, new, new_name in plan:
        if os.path.abspath(old) == os.path.abspath(path):
            if os.path.abspath(old) == os.path.abspath(new):
                return FileResult(path, 'skipped', '名称未变化')
            os.rename(old, new)
            undo = [{'src': old, 'dst': new}]
            return FileResult(path, 'ok', '重命名为 ' + new_name, new, undo=undo)
    return FileResult(path, 'fail', '重命名计划中未找到该文件')


def rename_by_excel_map(path: str, mapping: dict) -> FileResult:
    old = os.path.basename(path)
    new = mapping.get(old)
    if not new:
        return FileResult(path, 'skipped', '对照表中无此文件')
    if new == old:
        return FileResult(path, 'skipped', '名称未变化')
    try:
        validate_filename(new)
    except ValueError as e:
        return FileResult(path, 'fail', str(e))
    newp = os.path.join(os.path.dirname(os.path.abspath(path)), new)
    if os.path.exists(newp):
        return FileResult(path, 'fail', '目标名已存在: ' + new)
    os.rename(path, newp)
    undo = [{'src': path, 'dst': newp}]
    return FileResult(path, 'ok', '重命名为 ' + new, newp, undo=undo)


def load_excel_map(xlsx_path: str) -> dict:
    """Excel 对照表: 第1列=原名, 第2列=新名(首行为表头自动跳过)。"""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    ws = wb.active
    mapping: dict = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if len(row) >= 2 and row[0] is not None and row[1] is not None:
            old = str(row[0]).strip()
            new = str(row[1]).strip()
            if old and new:
                mapping[old] = new
    wb.close()
    return mapping


# ------------------------------------------------------------ 图片处理
def save_img(im, out_path: str, fmt: str, quality: int = 90) -> None:
    params = {}
    if fmt in ('JPEG', 'WEBP'):
        if im.mode in ('RGBA', 'P', 'LA'):
            im = im.convert('RGB')
        params['quality'] = int(quality)
    elif fmt == 'PNG' and im.mode == 'CMYK':
        im = im.convert('RGBA')
    im.save(out_path, fmt, **params)


def _make_text_layer(size, text: str, pos: str, opacity: int,
                     font_size: int):
    from PIL import Image, ImageDraw, ImageFont
    W, H = size
    layer = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    fs = max(12, int(font_size))
    font = None
    for fp in ('C:\\Windows\\Fonts\\msyh.ttc', 'C:\\Windows\\Fonts\\simhei.ttf',
               'C:\\Windows\\Fonts\\simsun.ttc', 'C:\\Windows\\Fonts\\arial.ttf'):
        if os.path.exists(fp):
            try:
                font = ImageFont.truetype(fp, fs)
                break
            except Exception:
                continue
    if font is None:
        font = ImageFont.load_default()
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    except Exception:
        tw, th = fs * len(text), fs
    margin = 10
    pos_map = {
        'tl': (margin, margin), 'tc': ((W - tw) // 2, margin),
        'tr': (W - tw - margin, margin), 'cl': (margin, (H - th) // 2),
        'cc': ((W - tw) // 2, (H - th) // 2), 'cr': (W - tw - margin, (H - th) // 2),
        'bl': (margin, H - th - margin), 'bc': ((W - tw) // 2, H - th - margin),
        'br': (W - tw - margin, H - th - margin),
    }
    x, y = pos_map.get(pos, pos_map['br'])
    alpha = max(0, min(255, int(opacity * 255 / 100)))
    draw.text((x, y), text, font=font,
              fill=(255, 255, 255, alpha), stroke_width=2,
              stroke_fill=(0, 0, 0, alpha))
    return layer


def process_image(path: str, options: dict) -> FileResult:
    from PIL import Image, ImageOps

    ext = os.path.splitext(path)[1].lower()
    if ext not in IMG_EXTS:
        return FileResult(path, 'skipped', '非支持图片格式')
    op = options.get('op', 'convert')
    out_dir = options.get('out_dir')
    base = os.path.splitext(os.path.basename(path))[0]
    target_dir = out_dir if out_dir else os.path.dirname(path)
    os.makedirs(target_dir, exist_ok=True)

    def out(suffix: str, new_ext: str = None) -> str:
        """输出路径。suffix 里若已带扩展名(如 '_clean.png'), 不再叠加原扩展名,
        否则会产出 photo_clean.png.jpg 这种双扩展名文件(内容与名字不符)。"""
        suf, suf_ext = os.path.splitext(suffix)
        if suf_ext:
            e = suf_ext
        else:
            e = new_ext if new_ext else ext
        p = os.path.abspath(os.path.join(target_dir, base + suf + e))
        # 同名不同扩展名的源图(photo.png / photo.jpg)会算出同一输出名,
        # 后处理的会覆盖先处理的, 造成结果丢失
        return _uniq_path(p) if os.path.abspath(p) != os.path.abspath(path) \
            else p

    im0 = Image.open(path)
    try:
        im0 = ImageOps.exif_transpose(im0)
    except Exception:
        pass

    if op == 'convert':
        target_fmt = str(options.get('format', 'PNG')).upper()
        fmt_map = {'JPG': 'JPEG', 'JPEG': 'JPEG', 'PNG': 'PNG', 'WEBP': 'WEBP',
                   'BMP': 'BMP', 'TIF': 'TIFF', 'TIFF': 'TIFF'}
        pil_fmt = fmt_map.get(target_fmt)
        if not pil_fmt:
            return FileResult(path, 'fail', '未知目标格式 ' + target_fmt)
        new_ext = {'JPEG': '.jpg', 'PNG': '.png', 'WEBP': '.webp',
                   'BMP': '.bmp', 'TIFF': '.tif'}[pil_fmt]
        o = out('_converted', new_ext)
        save_img(im0.copy(), o, pil_fmt, int(options.get('quality', 90) or 90))
        return FileResult(path, 'ok', '转换为 ' + target_fmt, o)

    if op == 'strip':
        im = im0.copy()
        im.info = {}
        o = out('_clean.png')
        save_img(im, o, 'PNG')
        return FileResult(path, 'ok', '已剥离元数据', o)

    if op == 'compress':
        quality = int(options.get('quality', 80) or 80)
        im = im0
        max_w = options.get('max_width') or 0
        if max_w and isinstance(max_w, (int, float)) and max_w > 0 and im.width > int(max_w):
            ratio = int(max_w) / im.width
            im = im.resize((int(max_w), max(1, round(im.height * ratio))),
                           Image.LANCZOS)
        if im.mode != 'RGB':
            im = im.convert('RGB')
        o = out('_compressed.jpg')
        save_img(im, o, 'JPEG', quality)
        so, sn = os.path.getsize(path), os.path.getsize(o)
        pct = round((1 - sn / so) * 100) if so else 0
        return FileResult(path, 'ok',
                          f'压缩 {so} → {sn} 字节({pct}%)', o)

    if op == 'resize':
        w = int(options.get('width') or 0)
        h = int(options.get('height') or 0)
        if w <= 0 and h <= 0:
            return FileResult(path, 'fail', '未填写目标尺寸')
        im = im0
        ow, oh = im.size
        if w > 0 and h > 0 and not options.get('keep_ratio', True):
            im = im.resize((w, h), Image.LANCZOS)
        elif w > 0:
            nh = max(1, round(oh * w / ow))
            im = im.resize((w, nh), Image.LANCZOS)
        else:
            nw = max(1, round(ow * h / oh))
            im = im.resize((nw, h), Image.LANCZOS)
        o = out('_resized.png')
        save_img(im, o, 'PNG')
        return FileResult(path, 'ok', f'调整为 {im.width}x{im.height}', o)

    if op == 'watermark':
        text = options.get('wm_text') or '青雁'
        pos = options.get('wm_pos', 'br')
        opacity = int(options.get('wm_opacity', 60) or 60)
        font_size = int(options.get('wm_size', 24) or 24)
        im = im0.convert('RGBA')
        layer = _make_text_layer(im.size, text, pos, opacity, font_size)
        im = Image.alpha_composite(im, layer)
        o = out('_wm.png')
        save_img(im, o, 'PNG')
        return FileResult(path, 'ok', f"加水印'{text}'", o)

    if op == 'rotate':
        angle = float(options.get('angle', 90) or 90)
        im = im0.rotate(angle, expand=True)
        o = out('_rotated.png')
        save_img(im, o, 'PNG')
        return FileResult(path, 'ok', f'旋转 {angle:g} 度', o)

    if op == 'flip':
        mode = options.get('flip', '水平')
        im = im0.transpose(Image.FLIP_TOP_BOTTOM if mode == '垂直'
                           else Image.FLIP_LEFT_RIGHT)
        o = out('_flip.png')
        save_img(im, o, 'PNG')
        return FileResult(path, 'ok', f'{mode}翻转', o)

    if op == 'crop':
        raw = str(options.get('crop', '') or '').strip()
        parts = [x.strip() for x in raw.split(',')]
        if len(parts) != 4:
            return FileResult(path, 'fail', '裁剪边距格式应为: 左,上,右,下')
        try:
            l, t, r, b = (int(x) for x in parts)
        except ValueError:
            return FileResult(path, 'fail', '裁剪边距格式应为: 左,上,右,下')
        W, H = im0.size
        if l + r >= W or t + b >= H:
            return FileResult(path, 'fail', f'裁剪边距超过图片尺寸 {W}x{H}')
        im = im0.crop((l, t, W - r, H - b))
        o = out('_crop.png')
        save_img(im, o, 'PNG')
        return FileResult(path, 'ok', f'裁剪为 {im.width}x{im.height}', o)

    if op == 'color':
        mode = options.get('color_mode', '灰度')
        m = {'灰度': 'L', '黑白': '1', 'RGBA': 'RGBA', 'RGB': 'RGB'}.get(mode)
        if not m:
            return FileResult(path, 'fail', '未知颜色模式 ' + str(mode))
        im = im0.convert(m)
        o = out('_color.png')
        save_img(im, o, 'PNG')
        return FileResult(path, 'ok', '颜色模式 ' + mode, o)

    if op == 'split':
        raw = str(options.get('grid', '') or '').strip()
        parts = [x.strip() for x in raw.split(',')]
        if len(parts) != 2:
            return FileResult(path, 'fail', '网格格式应为: 行,列')
        try:
            rows_n, cols_n = int(parts[0]), int(parts[1])
        except ValueError:
            return FileResult(path, 'fail', '网格格式应为: 行,列')
        if rows_n < 1 or cols_n < 1:
            return FileResult(path, 'fail', '行列必须 >= 1')
        # 上限保护: 99x99 会生成 9801 个文件, 拖垮磁盘与后续处理
        if rows_n * cols_n > 400:
            return FileResult(path, 'fail',
                              f'分割块数 {rows_n * cols_n} 过多(上限 400), '
                              '请减小行列数')
        W, H = im0.size
        if cols_n > W or rows_n > H:
            return FileResult(path, 'fail',
                              f'图片仅 {W}x{H}, 无法按 {rows_n}x{cols_n} 分割')
        cw, ch = W // cols_n, H // rows_n
        if cw < 1 or ch < 1:
            return FileResult(path, 'fail', '图片太小, 无法按此网格分割')
        outs = []
        n = 0
        for ri in range(rows_n):
            for ci in range(cols_n):
                piece = im0.crop((ci * cw, ri * ch,
                                  (ci + 1) * cw if ci < cols_n - 1 else W,
                                  (ri + 1) * ch if ri < rows_n - 1 else H))
                o = out(f'_r{ri + 1}c{ci + 1}.png', '.png')
                save_img(piece, o, 'PNG')
                outs.append(o)
                n += 1
        return FileResult(path, 'ok', f'分割为 {n} 块', outs[0] if outs else '')

    return FileResult(path, 'fail', '未知图片操作 ' + str(op))


# ------------------------------------------------------------ 文本处理
def md_to_html(md: str) -> str:
    out: list = []
    in_code = False
    for ln in md.splitlines():
        esc = escape(ln)
        s = esc.strip()
        if s.startswith('```'):
            out.append('</code></pre>' if in_code else '<pre><code>')
            in_code = not in_code
            continue
        if in_code:
            out.append(esc)
            continue
        if s.startswith('### '):
            out.append('<h3>' + s[4:] + '</h3>')
        elif s.startswith('## '):
            out.append('<h2>' + s[3:] + '</h2>')
        elif s.startswith('# '):
            out.append('<h1>' + s[2:] + '</h1>')
        elif s.startswith('- ') or s.startswith('* '):
            out.append('<li>' + s[2:] + '</li>')
        elif s:
            t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', esc)
            t = re.sub(r'\*(.+?)\*', r'<em>\1</em>', t)
            t = re.sub(r'\[(.+?)\]\((.+?)\)', r"<a href='\2'>\1</a>", t)
            out.append('<p>' + t + '</p>')
        else:
            out.append('')
    body = '\n'.join(out)
    return ("<!DOCTYPE html>\n<html>\n<head><meta charset='utf-8'></head>\n"
            '<body>\n' + body + '\n</body>\n</html>')


class _TextExtract(HTMLParser):
    def __init__(self):
        super().__init__()
        self.lines: list = []

    def handle_data(self, data):
        s = data.strip()
        if s:
            self.lines.append(s)


def html_to_text(content: str) -> str:
    p = _TextExtract()
    p.feed(content)
    return '\n'.join(p.lines)


def process_text(path: str, options: dict) -> FileResult:
    """文本处理入口: 收集就地修改产生的撤销动作后透传。"""
    undo_actions: list = []
    r = _process_text_impl(path, options, undo_actions)
    if undo_actions:
        r.undo = undo_actions
    return r


def _process_text_impl(path: str, options: dict,
                       undo_actions: list) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in TXT_EXTS:
        return FileResult(path, 'skipped', '非支持文本格式')
    op = options.get('op', 'replace')
    enc_from = options.get('enc_from', '自动识别')
    enc = ENC_MAP.get(enc_from) if enc_from and enc_from != '自动识别' else None
    if not enc:
        enc = detect_encoding(path)

    def _read() -> str:
        # newline='' 保留原始换行(不转换)。默认的 newline=None 会把 \r\n
        # 统一成 \n, 写回时 CRLF 文件就被静默改成 LF —— 对 .bat/.csv
        # 是实质破坏, 而且状态还报 ok。
        with open(path, 'r', encoding=enc, errors='replace',
                  newline='') as fh:
            return fh.read()

    def _read_lines() -> tuple:
        """按行读取 → (逻辑行列表, 原换行符)。逻辑行不带 \\r\\n。

        按行操作(去空行/去重/清行尾)必须先把 \r\n 归一, 否则 \r 会被
        当成行尾空白一起剥掉, 造成"清理了 2 行"的假报告。最后写回时
        再按原换行符拼回去, 不让 CRLF 文件被静默改成 LF。
        """
        raw = _read()
        eol = '\r\n' if '\r\n' in raw else '\n'
        text = raw.replace('\r\n', '\n').replace('\r', '\n')
        return text.split('\n'), eol

    def _write(out_path: str, text: str, encoding: str = None) -> None:
        """写文本: 先写临时文件再原子替换, 避免中途失败把原文件截断。"""
        use_enc = encoding or enc
        tmp = out_path + '.qytmp'
        try:
            with open(tmp, 'w', encoding=use_enc, newline='') as fh:
                fh.write(text)
            os.replace(tmp, out_path)
        except Exception:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
            raise

    def _out_path(default_suffix: str = '') -> str:
        """决定输出路径。

        - 未指定 out_dir: 就地修改, 先备份 .qingyanbak 并登记撤销
        - 指定 out_dir 且与源文件同目录: 同样走备份/撤销(否则会静默覆盖
          源文件, 又没有任何恢复手段)
        - 指定 out_dir 且文件已存在: 加序号避免覆盖别人的文件
        - default_suffix: 输出到新扩展名(txt→html), 不再原地改源文件
        """
        out_dir = options.get('out_dir')
        if not out_dir:
            bak = backup_file(path)
            undo_actions.append(
                {'src': bak, 'dst': path, 'kind': 'restore'})
            if default_suffix:
                stem = os.path.splitext(path)[0]
                return _uniq_path(stem + default_suffix)
            return path
        out_dir = os.path.abspath(out_dir)
        os.makedirs(out_dir, exist_ok=True)
        name = os.path.basename(path)
        if default_suffix:
            name = os.path.splitext(name)[0] + default_suffix
        out = os.path.join(out_dir, name)
        # 与源文件是同一个文件 → 必须备份, 否则等于无提示覆盖源文件
        if os.path.abspath(out) == os.path.abspath(path):
            bak = backup_file(path)
            undo_actions.append(
                {'src': bak, 'dst': path, 'kind': 'restore'})
            return path
        if os.path.exists(out):
            out = _uniq_path(out)
        return out

    # ---------------------------------------------------------- 查找替换
    if op == 'replace':
        find = options.get('find', '')
        if not find:
            return FileResult(path, 'skipped', '未填写查找内容')
        replace = options.get('replace', '')
        use_re = bool(options.get('regex', False))
        content = _read()
        if use_re:
            try:
                re.compile(find)
            except re.error as e:
                return FileResult(path, 'fail', '正则表达式无效: ' + str(e))
            hits = re.findall(find, content)
            n = len(hits)
            # 替换串是用户输入的普通文本, 其中的 \n \1 等不能被 re.sub
            # 当成转义(用户想插入 Windows 路径 C:\new 时会把 \n 变成换行)。
            # 但 \1..\9 是用户可能有意用的反向引用, 保留。
            safe_repl = re.sub(r'(?<!\\)\\(?![1-9])', r'\\\\', replace)
            try:
                new_content = re.sub(find, safe_repl, content)
            except re.error as e:
                return FileResult(path, 'fail', '替换内容无效: ' + str(e))
        else:
            n = content.count(find)
            new_content = content.replace(find, replace)
        if n == 0:
            return FileResult(path, 'skipped', '未匹配到内容')
        out = _out_path()
        _write(out, new_content)
        return FileResult(path, 'ok', f'替换 {n} 处' + ('(正则)' if use_re else ''), out)

    if op in ('del_blank', 'del_dup', 'trim'):
        lines, eol = _read_lines()
        # 末尾换行会在 split 后留一个空串, 它不是"空行"。单独记下是否有
        # 尾换行, 处理完再补回, 否则每次都会多删一行并丢掉文件末尾换行。
        tail = lines[-1] if lines and lines[-1] == '' else None
        if tail is not None:
            lines = lines[:-1]
        if op == 'del_blank':
            keep = [ln for ln in lines if ln.strip()]
            n = len(lines) - len(keep)
            detail = f'删除空行 {n} 行'
        elif op == 'del_dup':
            seen = set()
            keep = []
            for ln in lines:
                k = ln.strip()
                if k not in seen:
                    seen.add(k)
                    keep.append(ln)
            n = len(lines) - len(keep)
            detail = f'删除重复行 {n} 行'
        else:
            keep = [ln.rstrip() for ln in lines]
            n = sum(1 for a, b in zip(lines, keep) if a != b)
            detail = f'清理行尾空格 {n} 行'
        if n == 0:
            return FileResult(path, 'skipped', '没有需要处理的行')
        out = _out_path()
        new_content = eol.join(keep)
        if tail is not None and keep:
            new_content += eol
        _write(out, new_content)
        return FileResult(path, 'ok', detail, out)

    if op == 'crlf':
        content = _read()
        to = options.get('eol', 'lf')
        if to == 'lf':
            new_content = content.replace('\r\n', '\n')
        else:
            new_content = re.sub(r'(?<!\r)\n', '\r\n',
                                 content.replace('\r\n', '\n'))
        if new_content == content:
            return FileResult(path, 'skipped', '换行符已符合要求')
        out = _out_path()
        _write(out, new_content)
        return FileResult(path, 'ok', '换行符转换 ' + to.upper(), out)

    if op == 'bom':
        # BOM 必须按字节判断: utf-8-sig 读取时 Python 已吞掉 BOM, 用文本
        # 判断永远为 False(导致"移除"永远 skipped、"添加"产生双 BOM)。
        try:
            with open(path, 'rb') as fh:
                head = fh.read(4)
        except OSError as e:
            return FileResult(path, 'fail', f'读取失败: {e}')
        has_bom = head.startswith(b'\xef\xbb\xbf')
        to = options.get('bom_to', 'add')
        if to == 'add':
            if has_bom:
                return FileResult(path, 'skipped', 'BOM 状态已符合要求')
            if enc not in ('utf-8', 'utf-8-sig'):
                return FileResult(path, 'skipped',
                                  f'BOM 只适用于 UTF-8(当前 {enc})')
            content = _read()
            out = _out_path()
            _write(out, '\ufeff' + content.lstrip('\ufeff'), 'utf-8')
            return FileResult(path, 'ok', '添加 BOM', out)
        if not has_bom:
            return FileResult(path, 'skipped', 'BOM 状态已符合要求')
        content = _read()
        out = _out_path()
        _write(out, content.lstrip('\ufeff'), 'utf-8')
        return FileResult(path, 'ok', '移除 BOM', out)

    if op == 'encoding':
        target = options.get('enc_to', 'UTF-8')
        target_enc = ENC_MAP.get(target, 'utf-8')
        cur_norm = 'utf-8-sig' if enc == 'utf-8-sig' else enc
        if cur_norm == target_enc:
            return FileResult(path, 'skipped', '编码已相同')
        content = _read()
        out = _out_path()
        # 编码转换用临时文件: 目标编码表示不了某些字符时会抛异常, 直接写
        # 会把源文件截断为 0 字节(实测 UTF-8→GBK 遇生僻字就把文件清空了)
        try:
            _write(out, content, target_enc)
        except UnicodeEncodeError as e:
            if os.path.abspath(out) != os.path.abspath(path):
                try:
                    os.remove(out)
                except OSError:
                    pass
            return FileResult(
                path, 'fail',
                f'字符无法用 {target} 表示({e.reason} @ 第{e.start}字节), '
                '已放弃转换, 原文件未改动; 可先删除特殊字符再转')
        return FileResult(path, 'ok', f'{enc} → {target}', out)

    if op == 'split_lines':
        content = _read()
        lines = content.split('\n')
        if not lines or (len(lines) == 1 and not lines[0]):
            return FileResult(path, 'skipped', '空文件')
        # 末尾换行产生的空串不算一份, 否则会多出空文件
        if len(lines) > 1 and lines[-1].strip() == '':
            lines = lines[:-1]
        if not lines:
            return FileResult(path, 'skipped', '空文件')
        try:
            per = max(1, int(options.get('lines_per', 1000) or 1000))
        except (TypeError, ValueError):
            per = 1000
        out_dir = options.get('out_dir') or os.path.dirname(path)
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(path))[0]
        ext = os.path.splitext(path)[1] or '.txt'
        # 不同目录下的同名文件拆到同一输出目录会互相覆盖, 先探测并加区分
        existing = [f for f in os.listdir(out_dir)
                    if f.startswith(base + '_part')]
        tag = ''
        if existing:
            k = 2
            while any(f.startswith(f'{base}_{k}_part') for f in existing):
                k += 1
            tag = f'_{k}'
        n_parts = 0
        last = ''
        for i in range(0, len(lines), per):
            n_parts += 1
            part = os.path.join(
                out_dir, f'{base}{tag}_part{n_parts:03d}{ext}')
            _write(part, '\n'.join(lines[i:i + per]) + '\n')
            last = part
        return FileResult(path, 'ok',
                          f'拆为 {n_parts} 份(每份 {per} 行)', last)

    if op == 'txt2html':
        content = _read()
        body = '\n'.join('<p>' + escape(ln) + '</p>'
                         for ln in content.splitlines())
        html = ("<!DOCTYPE html>\n<html>\n<head><meta charset='utf-8'><title>"
                + escape(os.path.basename(path)) + "</title></head>\n<body>\n"
                + body + '\n</body>\n</html>')
        out = _out_path('.html')
        _write(out, html, 'utf-8')
        return FileResult(path, 'ok', 'txt → html', out)

    if op == 'md2html':
        content = _read()
        html = md_to_html(content)
        out = _out_path('.html')
        _write(out, html, 'utf-8')
        return FileResult(path, 'ok', 'md → html', out)

    if op == 'html2txt':
        content = _read()
        text = html_to_text(content)
        out = _out_path('.txt')
        with open(out, 'w', encoding='utf-8', newline='') as fh:
            fh.write(text)
        return FileResult(path, 'ok', 'html → txt', out)

    if op == 'json_fmt':
        content = _read()
        try:
            obj = json.loads(content)
        except Exception as e:
            return FileResult(path, 'fail', 'JSON 无效: ' + str(e))
        out = _out_path()
        with open(out, 'w', encoding='utf-8', newline='') as fh:
            json.dump(obj, fh, ensure_ascii=False, indent=2)
        return FileResult(path, 'ok', 'JSON 已格式化', out)

    return FileResult(path, 'fail', '未知文本操作 ' + str(op))


# ------------------------------------------------------------ 文件整理
def process_fileops(path: str, options: dict) -> FileResult:
    op = options.get('op', 'list')
    out_dir = options.get('out_dir')

    if op == 'list':
        root = path if os.path.isdir(path) else os.path.dirname(path)
        title = os.path.basename(root.rstrip('\\/')) or '文件清单'
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir or root, title + '_文件清单.csv')
        rows = []
        for dirpath, _dirs, files in sorted(walk_sources(root)):
            for fn in sorted(files):
                fp = os.path.join(dirpath, fn)
                try:
                    size = os.path.getsize(fp)
                except OSError:
                    size = 0
                rows.append((fp, os.path.relpath(fp, root), size))
        with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
            w = csv.writer(fh)
            w.writerow(('完整路径', '相对路径', '大小(字节)'))
            w.writerows(rows)
        return FileResult(path, 'ok', f'清单 {len(rows)} 条', out)

    if op == 'pack':
        base = os.path.splitext(os.path.basename(path))[0]
        target_dir = out_dir or os.path.dirname(path)
        os.makedirs(target_dir, exist_ok=True)
        zip_path = os.path.abspath(os.path.join(target_dir, base + '.zip'))
        # 输入本身是 .zip 时, base + '.zip' 恰好等于原文件: 以 'w' 打开会
        # 先把源压缩包截断, 然后把正在写的 zip 当内容打进去, 原内容全丢。
        if zip_path == os.path.abspath(path):
            zip_path = os.path.abspath(os.path.join(
                target_dir, base + '_打包.zip'))
        # 已有同名 zip 也不能直接覆盖, 改名保存
        zip_path = _uniq_path(zip_path)
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                if os.path.isfile(path):
                    zf.write(path, os.path.basename(path))
                else:
                    root_parent = os.path.dirname(os.path.abspath(path))
                    for dirpath, _dirs, files in walk_sources(path):
                        for fn in files:
                            fp = os.path.join(dirpath, fn)
                            # 输出文件若落在被遍历目录内, 会把正在写的 zip
                            # 自己打进去(zip 包内嵌自身), 必须排除
                            if os.path.abspath(fp) == zip_path:
                                continue
                            zf.write(fp, os.path.relpath(fp, root_parent))
        except Exception as e:
            try:
                if os.path.exists(zip_path) and zip_path != path:
                    os.remove(zip_path)
            except OSError:
                pass
            return FileResult(path, 'fail', f'打包失败: {e}')
        return FileResult(path, 'ok', '已打包', zip_path)

    if op == 'unzip':
        return process_zip(path, options)

    if op == 'by_ext':
        root = path if os.path.isdir(path) else os.path.dirname(path)
        n = 0
        undo: list = []
        for dirpath, _dirs, files in walk_sources(root):
            for fn in files:
                fp = os.path.join(dirpath, fn)
                d = os.path.join(root, os.path.splitext(fn)[1].lower().lstrip('.') or '无扩展名')
                if os.path.abspath(dirpath) == os.path.abspath(d):
                    continue
                os.makedirs(d, exist_ok=True)
                dst = os.path.join(d, fn)
                if os.path.exists(dst):
                    continue
                shutil.move(fp, dst)
                undo.append({'src': fp, 'dst': dst})
                n += 1
        if n == 0:
            return FileResult(path, 'skipped', '没有需要分类的文件')
        return FileResult(path, 'ok', f'按扩展名归类 {n} 个文件', root, undo=undo)

    if op == 'by_date':
        root = path if os.path.isdir(path) else os.path.dirname(path)
        n = 0
        undo: list = []
        for dirpath, _dirs, files in walk_sources(root):
            for fn in files:
                fp = os.path.join(dirpath, fn)
                day = datetime.date(datetime.fromtimestamp(os.path.getmtime(fp))).isoformat()
                d = os.path.join(root, day)
                if os.path.abspath(dirpath) == os.path.abspath(d):
                    continue
                os.makedirs(d, exist_ok=True)
                dst = os.path.join(d, fn)
                if os.path.exists(dst):
                    continue
                shutil.move(fp, dst)
                undo.append({'src': fp, 'dst': dst})
                n += 1
        if n == 0:
            return FileResult(path, 'skipped', '没有需要分类的文件')
        return FileResult(path, 'ok', f'按日期归类 {n} 个文件', root, undo=undo)

    if op == 'copy':
        targets = [t.strip() for t in str(options.get('targets', '') or '').split(';') if t.strip()]
        if not targets:
            return FileResult(path, 'fail', '未填写目标目录')
        n = 0
        for t in targets:
            os.makedirs(t, exist_ok=True)
            if os.path.isdir(path):
                # 目录源: 保留相对结构, 避免不同子目录同名文件互相覆盖
                for dirpath, _dirs, files in walk_sources(path):
                    for fn in files:
                        fp = os.path.join(dirpath, fn)
                        rel = os.path.relpath(fp, path)
                        dst = os.path.join(t, rel)
                        os.makedirs(os.path.dirname(dst) or t, exist_ok=True)
                        shutil.copy2(fp, dst)
                        n += 1
            else:
                shutil.copy2(path, os.path.join(t, os.path.basename(path)))
                n += 1
        return FileResult(path, 'ok', f'复制到 {len(targets)} 个目录, 共 {n} 个文件', targets[0])

    if op == 'mkdir':
        target_dir = out_dir or os.path.dirname(path)
        d = os.path.join(target_dir, os.path.splitext(os.path.basename(path))[0])
        os.makedirs(d, exist_ok=True)
        return FileResult(path, 'ok', '创建同名文件夹', d)

    if op == 'delete':
        trash_dir = options.get('trash_dir') or default_trash_dir()
        manifest = TrashManifest()
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        dst = os.path.join(trash_dir, stamp + '_' + os.path.basename(path))
        shutil.move(path, dst)
        manifest.record(os.path.basename(dst), path)
        undo = [{'src': path, 'dst': dst}]
        return FileResult(path, 'ok', '已移入回收备份区(可恢复)', dst, undo=undo)

    return FileResult(path, 'fail', '未知文件操作 ' + str(op))


# ------------------------------------------------------------ 文档转换
def process_doc(path: str, options: dict) -> FileResult:
    import docx

    ext = os.path.splitext(path)[1].lower()
    target = options.get('target', 'docx')
    out_dir = options.get('out_dir')
    base = os.path.splitext(os.path.basename(path))[0]
    target_dir = out_dir or os.path.dirname(path)
    os.makedirs(target_dir, exist_ok=True)

    if target == 'docx':
        out = os.path.join(target_dir, base + '.docx')
        if ext == '.docx':
            return FileResult(path, 'skipped', '已是 docx')
        if ext in ('.txt', '.md', '.log'):
            enc = options.get('encoding', 'utf-8')
            with open(path, 'r', encoding=enc, errors='replace') as fh:
                text = fh.read()
            d = docx.Document()
            for para in text.splitlines():
                d.add_paragraph(para)
            d.save(out)
            return FileResult(path, 'ok', 'txt 转 docx', out)
        if ext == '.csv':
            with open(path, 'r', encoding='utf-8-sig', errors='replace', newline='') as fh:
                rows = list(csv.reader(fh, delimiter='\t'))
            d = docx.Document()
            for row in rows:
                d.add_paragraph('\t'.join(row))
            d.save(out)
            return FileResult(path, 'ok', 'csv 转 docx(文本表)', out)
        if ext in ('.html', '.htm'):
            with open(path, 'r', encoding='utf-8', errors='replace') as fh:
                text = html_to_text(fh.read())
            d = docx.Document()
            for para in text.splitlines():
                d.add_paragraph(para)
            d.save(out)
            return FileResult(path, 'ok', 'html 转 docx', out)
        return FileResult(path, 'skipped', f'暂不支持 {ext} → docx')

    if target == 'txt':
        out = os.path.join(target_dir, base + '.txt')
        if ext == '.docx':
            d = docx.Document(path)
            text = '\n'.join(p.text for p in d.paragraphs)
            with open(out, 'w', encoding='utf-8', newline='') as fh:
                fh.write(text)
            return FileResult(path, 'ok', 'docx 转 txt', out)
        if ext in ('.html', '.htm'):
            with open(path, 'r', encoding='utf-8', errors='replace') as fh:
                text = html_to_text(fh.read())
            with open(out, 'w', encoding='utf-8', newline='') as fh:
                fh.write(text)
            return FileResult(path, 'ok', 'html 转 txt', out)
        return FileResult(path, 'skipped', f'暂不支持 {ext} → txt')

    return FileResult(path, 'fail', '未知目标格式 ' + str(target))


# ------------------------------------------------------------ zip 处理
def process_zip(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext != '.zip':
        return FileResult(path, 'skipped', '仅支持 zip(含 GBK 文件名修复)')
    out_dir = options.get('out_dir') or os.path.join(
        os.path.dirname(path), os.path.splitext(os.path.basename(path))[0])
    os.makedirs(out_dir, exist_ok=True)
    n = 0
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            name = info.filename
            try:
                fixed = name.encode('cp437').decode('gbk')
            except Exception:
                fixed = name
            target = os.path.join(out_dir, fixed)
            if info.is_dir():
                os.makedirs(target, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(target) or out_dir, exist_ok=True)
            with zf.open(info) as src, open(target, 'wb') as dst:
                shutil.copyfileobj(src, dst)
            n += 1
    return FileResult(path, 'ok', f'解压 {n} 个文件', out_dir)


# ------------------------------------------------------------ A1-A5 高频
def file_hasher(path: str, options: dict) -> FileResult:
    """A1 批量文件哈希: MD5/SHA1/SHA256 → CSV。

    options: alg(md5/sha1/sha256), out_dir
    入参为单个文件时按其父目录计算(逐文件调用的兼容)。
    UI 传的是 hash_alg, 早期只读 alg 导致选什么都算 MD5, 两个键都认。
    """
    alg = str(options.get('alg') or options.get('hash_alg')
              or 'md5').lower()
    # UI 显示的是大写 MD5/SHA1/SHA256, 归一化
    alg = alg.replace('-', '').replace('_', '')
    if alg not in ('md5', 'sha1', 'sha256'):
        return FileResult(path, 'fail', '未知算法: ' + alg)
    hfun = getattr(hashlib, alg)
    out_dir = options.get('out_dir')
    scope = os.path.dirname(path) if os.path.isfile(path) else path
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    rows = []
    fails = 0
    for dirpath, _dirs, files in walk_sources(scope):
        for fn in sorted(files):
            fp = os.path.join(dirpath, fn)
            try:
                h = hfun()
                with open(fp, 'rb') as fh:
                    for chunk in iter(lambda: fh.read(1048576), b''):
                        h.update(chunk)
                rows.append((os.path.relpath(fp, scope), h.hexdigest()))
            except Exception as e:
                rows.append((os.path.relpath(fp, scope), 'ERROR: ' + str(e)))
                fails += 1
    stamp = time.strftime('%Y%m%d_%H%M%S')
    csv_path = os.path.join(out_dir or scope, f'hash_{alg}_{stamp}.csv')
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as fh:
        w = csv.writer(fh)
        w.writerow(('文件(相对路径)', alg.upper()))
        w.writerows(rows)
    ok = len(rows) - fails
    detail = f'已计算 {ok} 个文件的 {alg.upper()}，{fails} 个失败'
    return FileResult(path, 'ok', detail, csv_path)


def file_lister(path: str, options: dict) -> FileResult:
    """A2 目录清单: 递归文件列表 → Excel/CSV/TXT。"""
    fmt = options.get('fmt', 'xlsx')
    out_dir = options.get('out_dir') or (os.path.dirname(path) if os.path.isfile(path) else path)
    os.makedirs(out_dir, exist_ok=True)
    root = path if os.path.isdir(path) else os.path.dirname(path)
    rows = []
    for dirpath, _dirs, files in walk_sources(root):
        for fn in sorted(files):
            fp = os.path.join(dirpath, fn)
            try:
                st = os.stat(fp)
                size = st.st_size
                mtime = time.strftime('%Y-%m-%d %H:%M:%S',
                                      time.localtime(st.st_mtime))
            except OSError:
                size, mtime = 0, ''
            ext = os.path.splitext(fn)[1].lstrip('.').upper() or '—'
            rows.append((os.path.relpath(fp, root), fn, size, mtime, ext))
    stamp = time.strftime('%Y%m%d_%H%M%S')
    out_path = os.path.join(out_dir, '文件清单_' + stamp + '.' + fmt)
    try:
        if fmt == 'xlsx':
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = '文件清单'
            ws.append(('相对路径', '文件名', '大小(字节)', '修改时间', '类型'))
            for row in rows:
                ws.append(list(row))
            wb.save(out_path)
        elif fmt == 'csv':
            with open(out_path, 'w', encoding='utf-8-sig', newline='') as fh:
                w = csv.writer(fh)
                w.writerow(('相对路径', '文件名', '大小(字节)', '修改时间', '类型'))
                w.writerows(rows)
        else:
            with open(out_path, 'w', encoding='utf-8') as fh:
                fh.write('相对路径\t文件名\t大小(字节)\t修改时间\t类型\n')
                for r in rows:
                    fh.write('\t'.join(str(x) for x in r) + '\n')
    except Exception as e:
        return FileResult(path, 'fail', '清单生成失败: ' + str(e))
    return FileResult(path, 'ok', f'已列出 {len(rows)} 个文件 → {out_path}', out_path)


def text_stats(path: str, options: dict) -> FileResult:
    """A3 文本统计: 行数/字数/字符数(去空白) → Excel。"""
    from openpyxl import Workbook

    out_dir = options.get('out_dir') or (os.path.dirname(path) if os.path.isfile(path) else path)
    os.makedirs(out_dir, exist_ok=True)
    scope = os.path.dirname(path) if os.path.isfile(path) else path
    files = collect_files(scope, recursive=True, ext_filter=sorted(TXT_EXTS))
    rows = []
    for fp in files:
        try:
            txt = read_text_auto(fp)
        except Exception:
            continue
        lines = txt.count('\n') + (1 if txt and not txt.endswith('\n') else 0)
        chars = len(txt)
        nonspace = sum(1 for ch in txt if not ch.isspace())
        words = len([w for w in re.split(r'\s+', txt.strip()) if w])
        rows.append((os.path.basename(fp), lines, words, chars, nonspace))
    stamp = time.strftime('%Y%m%d_%H%M%S')
    out = os.path.join(out_dir, '文本统计_' + stamp + '.xlsx')
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = '文本统计'
        ws.append(('文件', '行数', '单词/分词数', '字符数(含空白)', '非空白字符数'))
        for row in rows:
            ws.append(list(row))
        ws.append(('合计', sum(r[1] for r in rows), sum(r[2] for r in rows),
                   sum(r[3] for r in rows), sum(r[4] for r in rows)))
        wb.save(out)
    except Exception as e:
        return FileResult(path, 'fail', 'Excel 生成失败: ' + str(e))
    return FileResult(path, 'ok', f'已统计 {len(rows)} 个文本文件', out)


_TOUCH_NAME_RE = re.compile(r'(20\d{2})[-_/年]?(\d{1,2})[-_/月]?(\d{1,2})')


def touch_timestamps(path: str, options: dict) -> FileResult:
    """A4 修改时间批改: 设为指定时间 或 从文件名提取日期。"""
    mode = options.get('mode', 'fixed')
    files = collect_files(path, recursive=True) if os.path.isdir(path) else [path]
    if mode == 'fixed':
        fixed = str(options.get('ts_fixed', '') or '').strip()
        try:
            ts = time.mktime(time.strptime(fixed, '%Y-%m-%d %H:%M:%S'))
        except Exception:
            return FileResult(path, 'fail', '时间格式应为 YYYY-MM-DD HH:MM:SS')
        n = 0
        for fp in files:
            try:
                os.utime(fp, (ts, ts))
                n += 1
            except OSError:
                continue
        return FileResult(path, 'ok', f'已把 {n} 个文件的修改时间设为 {fixed}')
    # from_name
    n = skipped = 0
    pat = _TOUCH_NAME_RE
    for fp in files:
        m = pat.search(os.path.basename(fp))
        if not m:
            skipped += 1
            continue
        y, mo, d = int(m.group(1)), min(12, int(m.group(2))), min(31, int(m.group(3)))
        try:
            ts = time.mktime((y, mo, d, 12, 0, 0, 0, 0, -1))
            os.utime(fp, (ts, ts))
            n += 1
        except (OSError, OverflowError):
            skipped += 1
    return FileResult(path, 'ok',
                      f'按文件名日期改写 {n} 个；无法提取 {skipped} 个')


def pdf_blank_and_count(path: str, options: dict) -> FileResult:
    """A5 PDF: mode=blank 生成空白页 PDF; mode=count 批量页数统计 → Excel。"""
    from openpyxl import Workbook

    mode = options.get('mode', 'count')
    out_dir = options.get('out_dir') or (os.path.dirname(path) if os.path.isfile(path) else path)
    os.makedirs(out_dir, exist_ok=True)
    stamp = time.strftime('%Y%m%d_%H%M%S')

    if mode == 'blank':
        pages = max(1, min(500, int(options.get('pdf_pages', 1) or 1)))
        out_path = os.path.join(out_dir, f'空白页_{stamp}.pdf')
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            cv = canvas.Canvas(out_path, pagesize=A4)
            for _ in range(pages):
                cv.showPage()
            cv.save()
        except Exception as e:
            return FileResult(path, 'fail', '生成失败: ' + str(e))
        return FileResult(path, 'ok', f'已生成 {pages} 页空白 PDF', out_path)

    # count
    try:
        from pypdf import PdfReader
    except Exception as e:
        return FileResult(path, 'fail', 'pypdf 不可用: ' + str(e))
    scope = os.path.dirname(path) if os.path.isfile(path) else path
    rows = []
    for dirpath, _dirs, files in walk_sources(scope):
        for fn in sorted(files):
            if not fn.lower().endswith('.pdf'):
                continue
            fp = os.path.join(dirpath, fn)
            try:
                n = len(PdfReader(fp).pages)
            except Exception:
                n = -1
            rows.append((os.path.relpath(fp, scope), n))
    out = os.path.join(out_dir, 'PDF页数统计_' + stamp + '.xlsx')
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = 'PDF页数'
        ws.append(('文件(相对路径)', '页数'))
        for r in rows:
            ws.append(list(r))
        ws.append(('合计', sum(r[1] for r in rows if r[1] >= 0)))
        wb.save(out)
    except Exception as e:
        return FileResult(path, 'fail', 'Excel 生成失败: ' + str(e))
    return FileResult(path, 'ok', f'已统计 {len(rows)} 个 PDF', out)
