# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 主题系统。

设计: Müller-Brockmann 瑞士网格 — 8pt 基线, 严格左对齐, 黑白 + 单强调色。
v3 新增: 亮/暗双主题、强调色可配、字号缩放, 全部持久化到 ParamStore。
"""
from __future__ import annotations

ACCENT_PRESETS = (
    ('柔雾紫', '#635BDF'),
    ('青雁蓝', '#315CF5'),
    ('经典青蓝', '#0F6E8C'),
    ('黛蓝', '#2D5BA6'),
    ('翡翠', '#0B8457'),
    ('绛紫', '#7C4D8F'),
    ('石墨', '#37474F'),
)

LIGHT = {
    'INK': '#1A2430',
    'PAPER': '#FFFFFF',
    'PAPER2': '#F6F9FB',
    'PAPER3': '#FBFDFE',
    'LINE': '#E3EAF0',
    'LINE2': '#C6D8E2',
    'MUTED': '#7A858C',
    'FAILC': '#B3261E',
    'OKC': '#0B8457',
    'WARNC': '#B36A00',
    'SEL': '#E6F3F8',
    'HOVER': '#F1F7FA',
    'TIPBG': '#1A2430',
    'TIPFG': '#FFFFFF',
    'INPUT_BG': '#FFFFFF',
    'INPUT_FOCUS': '#FCFEFF',
}

DARK = {
    'INK': '#E8EEF2',
    'PAPER': '#151A20',
    'PAPER2': '#1B222B',
    'PAPER3': '#202834',
    'LINE': '#2C3642',
    'LINE2': '#3D4A59',
    'MUTED': '#8A97A3',
    'FAILC': '#E57373',
    'OKC': '#66BB8A',
    'WARNC': '#E0A24A',
    'SEL': '#1E3A47',
    'HOVER': '#232E3A',
    'TIPBG': '#0F1318',
    'TIPFG': '#E8EEF2',
    'INPUT_BG': '#1E252E',
    'INPUT_FOCUS': '#242D38',
}


def _mix(c1: str, c2: str, t: float) -> str:
    """线性混色: t=0 返回 c1, t=1 返回 c2。"""
    a = tuple(int(c1[i:i + 2], 16) for i in (1, 3, 5))
    b = tuple(int(c2[i:i + 2], 16) for i in (1, 3, 5))
    m = tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return '#%02X%02X%02X' % m


class Theme:
    """当前主题状态。MainWindow 持有实例, 切换时重建 QSS 并通知页面。"""

    def __init__(self, mode: str = 'light', accent: str = '#0F6E8C',
                 font_scale: float = 1.0):
        self.mode = mode if mode in ('light', 'dark') else 'light'
        self.accent = accent if accent.startswith('#') and len(accent) == 7 \
            else '#0F6E8C'
        self.font_scale = max(0.85, min(1.4, font_scale or 1.0))
        base = LIGHT if self.mode == 'light' else DARK
        self.c = dict(base)
        self.c['ACCENT'] = self.accent
        self.c['ACCENT_DARK'] = _mix(self.accent, '#000000', 0.25)
        self.c['ACCENT_SOFT'] = _mix(self.accent, base['PAPER'], 0.88)
        if self.mode == 'dark':
            # 暗色下强调色整体提亮, 保证文字/边框对比度
            self.c['ACCENT_BRIGHT'] = _mix(self.accent, '#FFFFFF', 0.45)
        else:
            self.c['ACCENT_BRIGHT'] = self.accent

    # -------------------------------------------------------------- helpers
    def px(self, n: int) -> int:
        """字号缩放后的像素值(仅用于字号)。"""
        return round(n * self.font_scale)

    def qss(self) -> str:
        c = self.c
        px = self.px
        return f'''
QMainWindow, QDialog {{ background: {c["PAPER"]}; }}
QWidget {{ font-family: "Microsoft YaHei UI"; font-size: {px(13)}px; color: {c["INK"]}; }}
QScrollArea {{ background: transparent; border: none; }}
QScrollArea > QWidget > QWidget {{ background: transparent; }}
QLabel {{ background: transparent; }}
QLineEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
  background: {c["INPUT_BG"]}; border: 1px solid {c["LINE"]};
  border-radius: 9px; padding: 6px 12px; min-height: 24px;
  selection-background-color: {c["ACCENT"]};
}}
QLineEdit:focus, QPlainTextEdit:focus, QSpinBox:focus, QComboBox:focus {{
  border: 1px solid {c["ACCENT"]}; background: {c["INPUT_FOCUS"]};
}}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox::down-arrow {{
  image: none; border-left: 5px solid transparent; border-right: 5px solid transparent;
  border-top: 6px solid {c["MUTED"]}; margin-right: 10px;
}}
QComboBox QAbstractItemView {{
  background: {c["PAPER"]}; border: 1px solid {c["LINE"]};
  selection-background-color: {c["SEL"]}; selection-color: {c["INK"]};
}}
QPushButton {{
  background: {c["PAPER"]}; border: 1px solid {c["LINE2"]}; border-radius: 9px;
  padding: 7px 20px; color: {c["INK"]};
}}
QPushButton:hover {{ background: {c["HOVER"]}; border-color: {c["ACCENT"]}; }}
QPushButton:pressed {{ background: {c["ACCENT_SOFT"]}; }}
QPushButton:disabled {{ color: {c["MUTED"]}; border-color: {c["LINE"]}; background: {c["PAPER2"]}; }}
QPushButton[accent="true"] {{
  background: {c["ACCENT"]}; color: white; border: none; font-weight: 600;
}}
QPushButton[accent="true"]:hover {{ background: {c["ACCENT_DARK"]}; }}
QPushButton[accent="true"]:disabled {{ background: {c["LINE2"]}; color: {c["PAPER"]}; }}
QTableWidget {{
  border: 1px solid {c["LINE"]}; border-radius: 12px; gridline-color: {c["LINE"]};
  background: {c["PAPER"]}; alternate-background-color: {c["PAPER2"]};
}}
QTableWidget::item {{ padding: 7px 10px; }}
QTableWidget::item:selected {{ background: {c["SEL"]}; color: {c["INK"]}; }}
QHeaderView::section {{
  background: {c["PAPER2"]}; color: {c["MUTED"]}; padding: 9px 12px; border: none;
  border-bottom: 1px solid {c["LINE"]}; font-weight: 600;
}}
QProgressBar {{
  border: none; border-radius: 7px; background: {c["PAPER2"]};
  text-align: center; height: 14px; font-size: {px(10)}px; color: {c["MUTED"]};
}}
QProgressBar::chunk {{ background: {c["ACCENT"]}; border-radius: 7px; }}
QTabWidget::pane {{ border: 1px solid {c["LINE"]}; border-radius: 8px; top: -1px; }}
QTabBar::tab {{
  background: {c["PAPER2"]}; padding: 7px 20px; margin-right: 2px;
  border-top-left-radius: 8px; border-top-right-radius: 8px; color: {c["MUTED"]};
}}
QTabBar::tab:selected {{
  background: {c["PAPER"]}; border-bottom: 2px solid {c["ACCENT"]};
  color: {c["ACCENT"]}; font-weight: 600;
}}
QCheckBox::indicator {{ width: 16px; height: 16px; border-radius: 5px; }}
QCheckBox::indicator:unchecked {{ border: 1px solid {c["LINE2"]}; background: {c["INPUT_BG"]}; }}
QCheckBox::indicator:unchecked:hover {{ border-color: {c["ACCENT"]}; }}
QCheckBox::indicator:checked {{ border: 1px solid {c["ACCENT"]}; background: {c["ACCENT"]}; }}
QGroupBox {{
  border: 1px solid {c["LINE"]}; border-radius: 14px; margin-top: 14px;
  padding: 18px 14px 12px 14px; background: {c["PAPER2"]};
}}
QGroupBox::title {{
  subcontrol-origin: margin; left: 14px; padding: 0 6px;
  font-weight: 600; color: {c["ACCENT_BRIGHT"]};
}}
QListWidget {{
  border: 1px solid {c["LINE"]}; border-radius: 8px; background: {c["PAPER"]};
  outline: none;
}}
QListWidget::item {{ padding: 5px 8px; border-radius: 6px; }}
QListWidget::item:selected {{ background: {c["SEL"]}; color: {c["INK"]}; }}
QScrollBar:vertical {{ background: transparent; width: 8px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {c["LINE2"]}; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: {c["ACCENT"]}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 8px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: {c["LINE2"]}; border-radius: 4px; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
QToolTip {{ background: {c["TIPBG"]}; color: {c["TIPFG"]}; border: none;
  padding: 5px 10px; border-radius: 6px; }}
QMenu {{ background: {c["PAPER"]}; border: 1px solid {c["LINE"]};
  border-radius: 8px; padding: 4px; }}
QMenu::item {{ padding: 6px 24px; border-radius: 6px; }}
QMenu::item:selected {{ background: {c["ACCENT"]}; color: white; }}
QMenu::separator {{ height: 1px; background: {c["LINE"]}; margin: 4px 8px; }}
QMessageBox {{ background: {c["PAPER"]}; }}
QInputDialog {{ background: {c["PAPER"]}; }}
'''

    # ------------------------------------------------------------ 语义色
    @property
    def INK(self):
        return self.c['INK']

    @property
    def MUTED(self):
        return self.c['MUTED']

    @property
    def ACCENT(self):
        return self.c['ACCENT']

    @property
    def FAILC(self):
        return self.c['FAILC']

    @property
    def OKC(self):
        return self.c['OKC']

    @property
    def WARNC(self):
        return self.c['WARNC']

    @property
    def ACCENT_BRIGHT(self):
        return self.c['ACCENT_BRIGHT']

    @property
    def ACCENT_SOFT(self):
        return self.c['ACCENT_SOFT']

    @property
    def PAPER(self):
        return self.c['PAPER']

    @property
    def PAPER2(self):
        return self.c['PAPER2']

    @property
    def LINE(self):
        return self.c['LINE']
