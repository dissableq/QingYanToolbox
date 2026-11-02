"""The 3.3 workspace: semantic styling shared by every page."""
import re
from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QWidget, QLabel, QPushButton, QFrame
from app import icons


def apply_design(main):
    th = main.theme
    for page in main.pages:
        for name, widget in vars(page).items():
            if name != 'main' and isinstance(widget, QWidget) and not widget.objectName():
                widget.setObjectName(type(page).__name__ + '_' + name)
    for widget in main.findChildren(QWidget):
        css = widget.styleSheet()
        if not widget.property('design_prepared'):
            if isinstance(widget, QLabel) and not widget.property('role'):
                size = re.search(r'font-size:\s*(\d+)px', css)
                if size and int(size[1]) >= 19:
                    widget.setProperty('role', 'title')
                elif ('font-weight:600' in css or 'font-weight:700' in css):
                    widget.setProperty('role', 'section')
                elif 'color:' in css:
                    widget.setProperty('role', 'muted')
            if isinstance(widget, QPushButton) and 'color:white' in css:
                widget.setProperty('accent', True)
            if (type(widget) is QFrame and 'background:' in css
                    and not widget.objectName()):
                widget.setProperty('card', True)
            widget.setProperty('design_prepared', True)
        if css:
            widget.setStyleSheet('')
    for i in range(main.nav.count()):
        name = main.nav.item(i).text()
        main.nav.item(i).setIcon(icons.get_icon(icons.NAV_ICONS.get(name, 'star'), 20, '#C4CDE0'))
    main.logo_btn.setText('青雁   ·   文件工作台')
    main.logo_btn.setObjectName('brand')
    main.nav.setObjectName('navigation')
    main.nav.setFixedWidth(204)
    main.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    from PySide6.QtWidgets import QApplication
    QApplication.instance().setStyleSheet(stylesheet(th))
    for page in main.pages:
        if hasattr(page, '_entries'):
            from app.home_workspace import refresh_home
            refresh_home(page)
        if hasattr(page, '_update_summary'):
            page._update_summary()


def stylesheet(th):
    dark = th.mode == 'dark'
    bg = '#111722' if dark else '#F5F5FA'
    surface = '#1B2433' if dark else '#FFFFFF'
    border = '#303E54' if dark else '#E3E8F1'
    ink = '#E8EDF7' if dark else '#17233A'
    muted = '#A3B0C7' if dark else '#64748B'
    soft = '#253653' if dark else '#EDF2FF'
    accent = th.ACCENT
    px = th.px
    assets = (Path(__file__).parent / 'assets').as_posix()
    return f'''
QWidget {{font-family:"Microsoft YaHei UI";font-size:{px(13)}px;color:{ink};}}
QMainWindow, QDialog {{background:{bg};}}
QWidget#workspace {{background:{bg};}}
QFrame#appHeader {{background:{surface};border-bottom:1px solid {border};}}
QFrame#taskFooter {{background:{surface};border-top:1px solid {border};}}
QLabel {{background:transparent;}}
QLabel[role="title"] {{font-size:{px(24)}px;font-weight:700;color:{ink};}}
QLabel[role="section"] {{font-size:{px(14)}px;font-weight:600;color:{ink};}}
QLabel[role="muted"] {{color:{muted};}}
QLabel[role="eyebrow"] {{color:{accent};font-size:{px(11)}px;font-weight:700;letter-spacing:2px;}}
QPushButton#brand {{font-size:{px(17)}px;font-weight:700;color:{ink};border:none;background:transparent;text-align:left;}}
QListWidget#navigation {{background:#23263D;border:none;border-radius:22px;padding:14px 8px;}}
QListWidget#navigation::item {{color:#C4CDE0;padding:10px 12px;margin:3px 1px;border-radius:14px;}}
QListWidget#navigation::item:hover {{background:#223650;}}
QListWidget#navigation::item:selected {{background:{accent};color:white;font-weight:600;}}
QScrollArea {{border:none;background:transparent;}}
QScrollArea > QWidget > QWidget {{background:transparent;}}
QGroupBox, QFrame[card="true"] {{background:{surface};border:1px solid {border};border-radius:20px;}}
QGroupBox {{margin-top:18px;padding:16px 12px 10px;}}
QGroupBox::title {{subcontrol-origin:margin;left:14px;padding:0 5px;color:{ink};font-weight:600;}}
QLineEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {{background:{surface};border:1px solid {border};border-radius:12px;padding:6px 9px;min-height:22px;selection-background-color:{accent};}}
QLineEdit:focus,QPlainTextEdit:focus,QComboBox:focus,QSpinBox:focus {{border:1px solid {accent};}}
QFrame[toolCard="true"] {{background:{surface};border:1px solid {border};border-radius:20px;}}
QFrame[toolCard="true"]:hover {{border-color:{accent};}}
QPushButton[cardTitle="true"] {{background:transparent;border:none;text-align:left;font-weight:600;padding:0;color:{ink};}}
QPushButton[cardTitle="true"]:hover {{color:{accent};}}
QPushButton[favorite="true"] {{padding:4px 8px;background:{bg};border:none;border-radius:12px;color:{muted};font-size:{px(11)}px;}}
QPushButton[favorite="true"]:checked {{background:{soft};color:{accent};}}
QLabel[role="emptyHint"] {{background:{surface};border:1px dashed {border};border-radius:18px;padding:18px;color:{muted};}}
QFrame[actionBar="true"] {{background:{surface};border:1px solid {border};border-radius:20px;}}
QLineEdit#featureSearch {{font-size:{px(15)}px;padding:12px 14px;}}
QComboBox::drop-down {{width:20px;border:none;}}
QComboBox::down-arrow {{image:url("{assets}/chevron.svg");width:12px;height:12px;}}
QComboBox QAbstractItemView {{background:{surface};color:{ink};selection-background-color:{soft};selection-color:{ink};}}
QPushButton {{background:{surface};border:1px solid {border};border-radius:12px;padding:9px 16px;}}
QPushButton:hover {{background:{soft};border-color:{accent};}}
QPushButton:pressed {{background:{soft};}}
QPushButton[accent="true"] {{background:{accent};color:white;border:1px solid {accent};font-weight:600;}}
QPushButton:disabled {{background:{bg};color:{muted};border-color:{border};}}
QPushButton[compact="true"] {{padding:0;}}
QPushButton[feature="true"] {{text-align:left;padding:14px 16px;background:{surface};font-weight:500;}}
QPushButton[feature="true"]:hover {{background:{soft};border-color:{accent};}}
QPushButton[category="true"]:checked {{background:{soft};color:{accent};border-color:{accent};}}
QTableWidget,QListWidget,QTreeWidget {{background:{surface};alternate-background-color:{bg};border:1px solid {border};border-radius:10px;gridline-color:{border};outline:none;}}
QTableWidget::item {{padding:8px;}}
QTableWidget::item:selected,QListWidget::item:selected {{background:{soft};color:{ink};}}
QListWidget::item {{padding:8px;}}
QHeaderView::section {{background:{surface};color:{muted};padding:10px;border:none;border-bottom:1px solid {border};font-weight:600;}}
QCheckBox {{spacing:6px;}}
QCheckBox::indicator {{width:16px;height:16px;border:1px solid {muted};border-radius:4px;background:{surface};}}
QCheckBox::indicator:checked {{background:{accent};border-color:{accent};image:url("{assets}/check.svg");}}
QProgressBar {{background:{soft};border:none;border-radius:5px;text-align:center;color:{muted};}}
QProgressBar::chunk {{background:{accent};border-radius:5px;}}
QTabWidget::pane {{border:1px solid {border};background:{surface};border-radius:10px;}}
QTabBar::tab {{padding:10px 16px;background:{bg};color:{muted};}}
QTabBar::tab:selected {{background:{surface};color:{accent};border-bottom:2px solid {accent};}}
QScrollBar:vertical {{background:transparent;width:9px;margin:2px;}}
QScrollBar::handle:vertical {{background:{border};border-radius:3px;min-height:30px;}}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical {{height:0;}}
QScrollBar:horizontal {{background:transparent;height:9px;margin:2px;}}
QScrollBar::handle:horizontal {{background:{border};min-width:30px;border-radius:3px;}}
QScrollBar::add-line:horizontal,QScrollBar::sub-line:horizontal {{width:0;}}
QSplitter::handle {{background:transparent;}}
QSplitter::handle:hover {{background:{border};}}
QMenu {{border-radius:12px;background:{surface};border:1px solid {border};padding:5px;}}
QMenu::item {{padding:8px 20px;}}
QMenu::item:selected {{background:{soft};color:{ink};}}
QToolTip {{background:#152238;color:#FFFFFF;padding:6px 10px;border:none;}}
'''
