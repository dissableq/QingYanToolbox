# -*- coding: utf-8 -*-
"""安装器参数解析测试。

静默安装( /SILENT /DIR=... )过去从未被自动验证过, 导致一个真 bug
存活了 5 个版本: 解析 /DIR= 的正则里裸 "/D" 分支会吞掉 "IR=", 路径
变成 "IR=D:\\..." 这种不存在的相对路径, 带 /DIR= 的静默安装必然失败
(只报「目标目录不可写」, 看不出真正原因)。此处固化各开关形态。
"""
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_testinstall')
if os.path.isdir(BASE):
    shutil.rmtree(BASE, ignore_errors=True)
os.makedirs(BASE)

ok = fail = 0


def check(name, cond, detail=''):
    global ok, fail
    if cond:
        ok += 1
        print('PASS', name)
    else:
        fail += 1
        print('FAIL', name, '::', detail)


import installer_entry as IE

# --- 复刻 main() 里的开关归一逻辑(与源码保持一致, 改动需同步)


def _flag(a):
    s = str(a).replace('\\', '/')
    return s.split(':', 1)[0].strip('/').lstrip('-').upper()


def parse(argv):
    silent = any(_flag(a) in ('S', 'SILENT') for a in argv)
    do_uninstall = any(_flag(a) in ('U', 'UNINSTALL') for a in argv)
    custom = ''
    for a in argv:
        s = str(a).strip('"')
        m = re.match(r'^(?:/DIR=|-DIR=|/D=|-D=)(.+)$', s, re.IGNORECASE)
        if m:
            custom = m.group(1).strip().strip('"')
            break
    if custom:
        if re.match(r'^/[A-Za-z]:', custom):
            custom = custom[1:]
        custom = os.path.normpath(custom)
    return silent, do_uninstall, custom


BS = chr(92)

# --- /SILENT 识别
s, u, c = parse(['/SILENT', '/SUPPRESSMSGBOXES', '/NORESTART'])
check('SILENT 开关被识别', s and not u and not c, (s, u, c))

# --- /DIR= 正式开关(本安装器文档用的形式)
s, u, c = parse(['/SILENT', '/DIR=D:' + BS + 'qy315_probe'])
check('/DIR= 路径完整解析', c == 'D:' + BS + 'qy315_probe', repr(c))

# --- 历史 bug 回归: /DIR= 不能被吞成 IR=
s, u, c = parse(['/SILENT', '/DIR=D:' + BS + '探针'])
check('/DIR= 不产生 IR= 残路径', not c.startswith('IR=') and
      c.endswith('探针'), repr(c))

# --- /DIR= 带空格与中文
s, u, c = parse(['/SILENT', '/DIR=C:' + BS + 'Program Files' + BS + '青雁'])
check('/DIR= 含空格的路径', c == 'C:' + BS + 'Program Files' + BS + '青雁',
      repr(c))

# --- Git Bash 形态: /D:/x 带前导斜杠
s, u, c = parse(['/SILENT', '/DIR=/D:/x'])
check('/DIR=/D:/x 前导斜杠被剥掉', c == 'D:' + BS + 'x', repr(c))

# --- 兼容 /D= 与 -D=
s, u, c = parse(['/SILENT', '/D=D:' + BS + 'a'])
check('/D= 兼容形式', c == 'D:' + BS + 'a', repr(c))
s, u, c = parse(['/SILENT', '-D=D:' + BS + 'b'])
check('-D= 兼容形式', c == 'D:' + BS + 'b', repr(c))

# --- 卸载开关
s, u, c = parse(['/UNINSTALL'])
check('UNINSTALL 开关被识别', u and not s, (s, u))

# --- 无 /DIR= 时应为空(走默认路径)
s, u, c = parse(['/SILENT', '/LOG=D:' + BS + 'l.txt'])
check('无 /DIR= 时 custom 为空', c == '', repr(c))

# --- _writable 本身: 父目录可写性判断
check('_writable 对真实可写目录为真', IE._writable(BASE), BASE)
check('_writable 对不存在路径的祖先判断',
      IE._writable(os.path.join(BASE, 'no', 'such', 'deep')) is True,
      os.path.join(BASE, 'no', 'such', 'deep'))
check('_writable 对空串为假', not IE._writable(''), '')

# --- 默认安装目录形态
d = IE.default_install_dir()
check('默认安装目录为绝对路径且含应用目录名',
      os.path.isabs(d) and IE.APP_DIR in d, d)

shutil.rmtree(BASE, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
