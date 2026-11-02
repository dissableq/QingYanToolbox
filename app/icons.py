# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — QPainter 自绘 20x20 线性图标库。

零外部依赖、离线可用, 圆头线帽 + 单强调色风格, 与全局 QSS 一致。
v3: 新增 sun/moon(主题切换)/undo/refresh/eta 等图标。
"""
from __future__ import annotations

import math

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (QColor, QFont, QIcon, QPainter, QPainterPath,
                           QPen, QPixmap)

_CACHE: dict = {}
_DEFAULT_COLOR = '#1A2430'


def set_default_color(color: str) -> None:
    """主题切换时更新图标默认色, 并清空缓存。"""
    global _DEFAULT_COLOR, _CACHE
    if _DEFAULT_COLOR != color:
        _DEFAULT_COLOR = color
        _CACHE = {}


def _painter(pm):
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    return p


def _pen(color: str, width: float) -> QPen:
    pen = QPen(QColor(color))
    pen.setWidthF(width)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    return pen


def _doc_frame(p, r: QRectF, color: str, fold: bool = True):
    """文档外框 + 右下折角。"""
    p.setPen(_pen(color, 1.8))
    p.setBrush(Qt.NoBrush)
    path = QPainterPath()
    path.moveTo(r.left(), r.top())
    path.lineTo(r.right() - 4, r.top())
    path.lineTo(r.right(), r.top() + 4)
    path.lineTo(r.right(), r.bottom())
    path.lineTo(r.left(), r.bottom())
    path.closeSubpath()
    p.drawPath(path)
    if fold:
        p.drawLine(QPointF(r.right() - 4, r.top()),
                   QPointF(r.right() - 4, r.top() + 4))
        p.drawLine(QPointF(r.right() - 4, r.top() + 4),
                   QPointF(r.right(), r.top() + 4))


def _paint(name: str, pm, color: str):
    p = _painter(pm)
    r = QRectF(2.5, 2.5, 15, 15)
    f = QFont('Segoe UI', 7)
    f.setPixelSize(8)
    f.setBold(True)

    if name == 'word':
        _doc_frame(p, r, color)
        p.setPen(_pen(color, 2.2))
        p.drawLine(QPointF(6, 7.5), QPointF(7.6, 12.5))
        p.drawLine(QPointF(7.6, 12.5), QPointF(10, 8.5))
        p.drawLine(QPointF(10, 8.5), QPointF(12.4, 12.5))
        p.drawLine(QPointF(12.4, 12.5), QPointF(14, 7.5))
    elif name == 'excel':
        _doc_frame(p, r, color)
        p.setPen(_pen(color, 2.0))
        p.drawLine(QPointF(6.5, 7), QPointF(13.5, 13))
        p.drawLine(QPointF(13.5, 7), QPointF(6.5, 13))
    elif name == 'ppt':
        _doc_frame(p, r, color)
        p.setPen(_pen(color, 2.0))
        p.setBrush(Qt.NoBrush)
        p.drawRect(QRectF(6, 7.5, 8, 5))
        p.drawLine(QPointF(10, 12.5), QPointF(10, 14.5))
    elif name == 'pdf':
        _doc_frame(p, r, color, fold=False)
        p.setPen(_pen(color, 1.6))
        p.drawText(QRectF(3, 6.5, 14, 8), Qt.AlignCenter, 'PDF')
    elif name == 'text':
        _doc_frame(p, r, color)
        p.setPen(_pen(color, 1.5))
        for y in (7, 10, 13):
            p.drawLine(QPointF(5.5, y), QPointF(14.5, y))
    elif name == 'image':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(QRectF(2.5, 3.5, 15, 13), 2, 2)
        p.drawEllipse(QPointF(7, 8), 1.4, 1.4)
        path = QPainterPath()
        path.moveTo(4, 14)
        path.lineTo(9, 9.5)
        path.lineTo(12.5, 13)
        path.lineTo(14.5, 11)
        path.lineTo(17, 14)
        p.drawPath(path)
    elif name == 'rename':
        p.setPen(_pen(color, 1.8))
        p.drawLine(QPointF(5.5, 5.5), QPointF(5.5, 17))
        p.drawLine(QPointF(10, 3), QPointF(10, 14.5))
        p.drawLine(QPointF(3.5, 5.5), QPointF(8.8, 5.5))
        p.drawLine(QPointF(7.5, 14.5), QPointF(13, 14.5))
        p.drawLine(QPointF(14.5, 11.2), QPointF(17, 15.2))
        p.drawLine(QPointF(17, 15.2), QPointF(15.5, 16.5))
    elif name == 'files':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        path = QPainterPath()
        path.moveTo(3, 6)
        path.lineTo(8, 6)
        path.lineTo(9.5, 8)
        path.lineTo(17, 8)
        path.lineTo(17, 16)
        path.lineTo(3, 16)
        path.closeSubpath()
        p.drawPath(path)
    elif name == 'template':
        _doc_frame(p, r, color)
        p.setPen(_pen(color, 1.5))
        p.drawLine(QPointF(6, 7.5), QPointF(14, 7.5))
        p.drawLine(QPointF(6, 10), QPointF(12, 10))
        p.drawLine(QPointF(6, 12.5), QPointF(13.5, 12.5))
    elif name == 'merge':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        path = QPainterPath()
        path.moveTo(3, 4.5)
        path.lineTo(8, 4.5)
        path.lineTo(12, 10)
        path.lineTo(12, 10)
        path.moveTo(3, 15.5)
        path.lineTo(8, 15.5)
        path.lineTo(12, 10)
        path.moveTo(12, 10)
        path.lineTo(17, 10)
        p.drawPath(path)
        p.drawLine(QPointF(14.8, 7.8), QPointF(17, 10))
        p.drawLine(QPointF(14.8, 12.2), QPointF(17, 10))
    elif name == 'tools':
        p.setPen(_pen(color, 1.8))
        p.drawEllipse(QPointF(8.5, 8.5), 4.5, 4.5)
        p.drawLine(QPointF(11.8, 11.8), QPointF(16.5, 16.5))
    elif name == 'records':
        _doc_frame(p, r, color)
        p.setPen(_pen(color, 1.5))
        p.drawLine(QPointF(6, 7.5), QPointF(14, 7.5))
        p.drawLine(QPointF(6, 10), QPointF(14, 10))
        p.drawLine(QPointF(6, 12.5), QPointF(11, 12.5))
    elif name == 'trash':
        p.setPen(_pen(color, 1.8))
        p.drawLine(QPointF(4.5, 6), QPointF(15.5, 6))
        p.drawLine(QPointF(8, 6), QPointF(8, 4.5))
        p.drawLine(QPointF(12, 6), QPointF(12, 4.5))
        p.drawLine(QPointF(8, 4.5), QPointF(12, 4.5))
        path = QPainterPath()
        path.moveTo(5.5, 6)
        path.lineTo(6.5, 16.5)
        path.lineTo(13.5, 16.5)
        path.lineTo(14.5, 6)
        p.drawPath(path)
        for x in (8, 10, 12):
            p.drawLine(QPointF(x, 8), QPointF(x, 14))
    elif name == 'pipeline':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(QPointF(5, 5), 2.2, 2.2)
        p.drawEllipse(QPointF(15, 10), 2.2, 2.2)
        p.drawEllipse(QPointF(5, 15), 2.2, 2.2)
        p.drawLine(QPointF(7, 6), QPointF(13, 9.2))
        p.drawLine(QPointF(13, 10.8), QPointF(7, 14))
    elif name == 'star':
        path = QPainterPath()
        cx, cy, R, r2 = 10, 10.5, 7.2, 3.0
        for i in range(10):
            rad = math.pi / 2 + i * math.pi / 5
            rad2 = math.pi / 2 + (i + 0.5) * math.pi / 5
            if i == 0:
                path.moveTo(cx + R * math.cos(rad), cy - R * math.sin(rad))
            else:
                path.lineTo(cx + R * math.cos(rad), cy - R * math.sin(rad))
            path.lineTo(cx + r2 * math.cos(rad2), cy - r2 * math.sin(rad2))
        path.closeSubpath()
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(color))
        p.drawPath(path)
    elif name == 'home':
        p.setPen(_pen(color, 1.8))
        path = QPainterPath()
        path.moveTo(3, 10)
        path.lineTo(10, 3.5)
        path.lineTo(17, 10)
        p.drawPath(path)
        p.drawRect(QRectF(6, 10, 8, 6.5))
    elif name == 'help':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(QPointF(10, 10), 7, 7)
        p.setFont(f)
        p.drawText(QRectF(3, 3, 14, 14), Qt.AlignCenter, '?')
    elif name == 'settings':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(QPointF(10, 10), 3, 3)
        for i in range(8):
            a = i * math.pi / 4
            x1 = 10 + 5.0 * math.cos(a)
            y1 = 10 + 5.0 * math.sin(a)
            x2 = 10 + 7.0 * math.cos(a)
            y2 = 10 + 7.0 * math.sin(a)
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))
    elif name == 'about':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(QPointF(10, 10), 7, 7)
        p.drawPoint(QPointF(10, 8.2))
        p.drawLine(QPointF(10, 10.5), QPointF(10, 14))
    elif name == 'clock':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(QPointF(10, 10), 6.5, 6.5)
        p.drawLine(QPointF(10, 6), QPointF(10, 10.5))
        p.drawLine(QPointF(10, 10.5), QPointF(13.2, 12.4))
    elif name == 'sun':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawEllipse(QPointF(10, 10), 3.2, 3.2)
        for i in range(8):
            a = i * math.pi / 4
            x1 = 10 + 5.2 * math.cos(a)
            y1 = 10 + 5.2 * math.sin(a)
            x2 = 10 + 7.2 * math.cos(a)
            y2 = 10 + 7.2 * math.sin(a)
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))
    elif name == 'moon':
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        path = QPainterPath()
        path.moveTo(13.5, 4)
        path.arcTo(QRectF(3, 3, 14, 14), 40, -260)
        p.drawPath(path)
    elif name == 'undo':
        p.setPen(_pen(color, 1.8))
        path = QPainterPath()
        path.moveTo(15, 14)
        path.arcTo(QRectF(4, 6, 11, 11), 0, 180)
        p.drawPath(path)
        p.drawLine(QPointF(4, 11.5), QPointF(4, 7))
        p.drawLine(QPointF(4, 7), QPointF(8.5, 7))
    elif name == 'refresh':
        p.setPen(_pen(color, 1.8))
        path = QPainterPath()
        path.moveTo(15.5, 8)
        path.arcTo(QRectF(4.5, 4.5, 11, 11), 30, 300)
        p.drawPath(path)
        p.drawLine(QPointF(15.5, 4), QPointF(15.5, 8))
        p.drawLine(QPointF(15.5, 8), QPointF(12, 8))
    elif name == 'print':
        # 打印机: 出纸 + 机身 + 托盘 + 指示灯
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawRect(QRectF(5.5, 3, 9, 4))
        p.drawRoundedRect(QRectF(3, 7, 14, 6.5), 1.5, 1.5)
        p.drawRect(QRectF(6.5, 12, 7, 4.5))
        p.drawPoint(QPointF(14.6, 9.2))
    elif name == 'convert':
        # 格式转换: 左右两个文档框 + 对向箭头(上右下左)
        p.setPen(_pen(color, 1.8))
        p.setBrush(Qt.NoBrush)
        p.drawRect(QRectF(2.5, 4, 7, 9))          # 左文档
        p.drawRect(QRectF(10.5, 7, 7, 9))         # 右文档
        p.drawLine(QPointF(5, 6.5), QPointF(9.5, 11.5))   # 箭头 左上→右下
        p.drawLine(QPointF(9.5, 11.5), QPointF(9.5, 9))
        p.drawLine(QPointF(9.5, 11.5), QPointF(7, 11.5))
        p.drawLine(QPointF(15, 8), QPointF(10.5, 8))      # 箭头 右上→左上
        p.drawLine(QPointF(10.5, 8), QPointF(10.5, 10))
        p.drawLine(QPointF(10.5, 8), QPointF(13, 8))
    elif name == 'play':
        p.setPen(Qt.NoPen)
        p.setBrush(QColor(color))
        path = QPainterPath()
        path.moveTo(6.5, 4)
        path.lineTo(16, 10)
        path.lineTo(6.5, 16)
        path.closeSubpath()
        p.drawPath(path)
    p.end()


def get_icon_pixmap(name: str, size: int = 20, color: str = None):
    """返回 (name) 图标的 QPixmap。color=None 时用主题默认色。"""
    if color is None:
        color = _DEFAULT_COLOR
    key = (name, size, color)
    if key in _CACHE:
        return _CACHE[key]
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    _paint(name, pm, color)
    _CACHE[key] = pm
    return pm


def get_icon(name: str, size: int = 20, color: str = None) -> QIcon:
    return QIcon(get_icon_pixmap(name, size, color))


NAV_ICONS = {
    '首页': 'home',
    'Word 工具': 'word',
    'Excel 工具': 'excel',
    'PPT 工具': 'ppt',
    'PDF 工具': 'pdf',
    '文本工具': 'text',
    '图片工具': 'image',
    '重命名工具': 'rename',
    '文件整理工具': 'files',
    '模板批量生成': 'template',
    '办公模板库': 'template',
    '合并与拆分': 'merge',
    '辅助工具': 'tools',
    '任务台账': 'records',
    '回收站': 'trash',
    '任务流水线': 'pipeline',
    '批量打印': 'print',
    '格式转换': 'convert',
    '帮助中心': 'help',
    '全局设置': 'settings',
    '关于': 'about',
}

GROUP_ICONS = {
    'Word 工具': 'word',
    'Excel 工具': 'excel',
    'PPT 工具': 'ppt',
    'PDF 工具': 'pdf',
    '文本工具': 'text',
    '图片工具': 'image',
    '重命名工具': 'rename',
    '文件整理工具': 'files',
    '常用工具': 'star',
    '更多': 'template',
}
