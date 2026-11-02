"""Responsive tool cards with explicit favorite controls and natural height."""
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPainter, QPalette
from PySide6.QtWidgets import (QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QWidget, QScrollArea, QFrame, QSizePolicy, QLayout, QMessageBox,
    QStyle, QStyleOptionButton)
from engine.store import Favorites, load_recent
from app.home_registry import COMMON_ENTRIES


class CardGrid(QWidget):
    def __init__(self):
        super().__init__()
        self.grid=QGridLayout(self); self.grid.setContentsMargins(0,0,0,0); self.grid.setSpacing(12)
        self.cards=[]; self.columns=0
        self.setSizePolicy(QSizePolicy.Expanding,QSizePolicy.Fixed)

    def set_cards(self,cards):
        while self.grid.count():self.grid.takeAt(0)
        for card in self.cards:card.hide();card.deleteLater()
        self.cards=cards;self.columns=0;self.reflow()

    def reflow(self):
        columns=max(1,min(3,(max(self.width(),260)+12)//300))
        if columns!=self.columns:
            while self.grid.count():self.grid.takeAt(0)
            for i,card in enumerate(self.cards):self.grid.addWidget(card,i//columns,i%columns);card.show()
            for c in range(3):self.grid.setColumnStretch(c,1 if c<columns else 0)
            self.columns=columns
        self.grid.invalidate()
        height=max(self.grid.minimumSize().height(),self.grid.sizeHint().height()) if self.cards else 0
        self.setFixedHeight(height);self.updateGeometry()

    def resizeEvent(self,event):
        super().resizeEvent(event);self.reflow()


class WrappedTitle(QPushButton):
    """Keep a real accessible button while allowing two-line tool names."""
    def paintEvent(self, event):
        option = QStyleOptionButton()
        self.initStyleOption(option)
        option.text = ''
        painter = QPainter(self)
        self.style().drawControl(QStyle.CE_PushButton, option, painter, self)
        painter.setPen(self.palette().color(QPalette.ButtonText))
        painter.drawText(self.rect(), Qt.AlignLeft | Qt.AlignVCenter | Qt.TextWordWrap, self.text())


class ToolCard(QFrame):
    def __init__(self,home,page,operation,hint='',favorite=False):
        super().__init__();self.setProperty('toolCard',True)
        self.setSizePolicy(QSizePolicy.Expanding,QSizePolicy.Fixed)
        scale=home.main.theme.font_scale;self.setMinimumHeight(round(150*scale))
        layout=QVBoxLayout(self);layout.setContentsMargins(16,12,16,14);layout.setSpacing(6)
        row=QHBoxLayout();row.setSpacing(6)
        self.open=WrappedTitle(operation);self.open.setProperty('cardTitle',True)
        self.open.setSizePolicy(QSizePolicy.Ignored,QSizePolicy.Fixed)
        self.open.setObjectName('open_'+page+'_'+operation)
        self.open.setAccessibleName('打开 '+page+' '+operation)
        self.open.setFixedHeight(round(54*scale));self.open.setCursor(Qt.PointingHandCursor)
        self.open.setToolTip(page+' · '+operation)
        self.open.clicked.connect(lambda:home.main.goto_feature(page,operation))
        self.open.setContextMenuPolicy(Qt.CustomContextMenu)
        self.open.customContextMenuRequested.connect(lambda pos:home._fav_menu(pos,self.open,operation,page,operation))
        layout.addWidget(self.open)
        star=QPushButton('已收藏' if favorite else '收藏');star.setProperty('favorite',True)
        star.setObjectName('favorite_'+page+'_'+operation)
        star.setAccessibleName(('取消收藏 ' if favorite else '收藏 ')+page+' '+operation)
        star.setCheckable(True);star.setChecked(favorite)
        star.setMinimumSize(round(66*scale),round(32*scale));star.setCursor(Qt.PointingHandCursor)
        star.clicked.connect(lambda:toggle_favorite(home,operation,page,operation,favorite))
        text=QLabel(hint or page);text.setProperty('role','muted');text.setWordWrap(True)
        text.setMinimumHeight(round(32*scale));text.setAlignment(Qt.AlignVCenter|Qt.AlignLeft)
        row.addWidget(text,1);row.addWidget(star);layout.addLayout(row)


def build_home(home):
    home._fav=Favorites();home._buttons={};home._group_caps={};home._common_buttons=[];home._entries=[]
    for page in home.main.pages:
        if not hasattr(page,'OPS'):continue
        combo=page._op_widget()
        ops=[combo.itemText(i) for i in range(combo.count())] if combo else list(page.OPS)
        home._entries.extend((page.title,o,getattr(page,'OP_HINTS',{}).get(o,'')) for o in ops)
    outer=QVBoxLayout(home);outer.setContentsMargins(24,22,24,18);outer.setSpacing(14)
    title=QLabel('今天，少一点重复工作。');title.setProperty('role','title');outer.addWidget(title)
    intro=QLabel(f'{len(home._entries)} 个工具 · 找到工具，添加文件，预览后开始处理。');intro.setProperty('role','muted');outer.addWidget(intro)
    row=QHBoxLayout();row.setSpacing(10)
    home.search_edit=QLineEdit();home.search_edit.setObjectName('featureSearch');home.search_edit.setAccessibleName('搜索工具')
    home.search_edit.setClearButtonEnabled(True);home.search_edit.setPlaceholderText('搜索工具名称，例如：查找替换、拆分、转 PDF…');row.addWidget(home.search_edit,1)
    home.category=QComboBox();home.category.setObjectName('toolCategory');home.category.setAccessibleName('工具分类')
    home.category.addItems(['全部分类']+list(dict.fromkeys(e[0] for e in home._entries)));home.category.setMinimumWidth(150)
    row.addWidget(home.category);outer.addLayout(row)
    scroll=QScrollArea();scroll.setWidgetResizable(True);scroll.setObjectName('homeScroll')
    inner=QWidget();contents=QVBoxLayout(inner);contents.setContentsMargins(0,4,12,12);contents.setSpacing(14)
    contents.setSizeConstraint(QLayout.SetMinimumSize)
    def heading(text):
        label=QLabel(text);label.setProperty('role','section');contents.addWidget(label);return label
    home._fav_title=heading('我的收藏')
    home._fav_hint=QLabel('点击工具卡片上的「收藏」，常用工具就会出现在这里。')
    home._fav_hint.setProperty('role','emptyHint');home._fav_hint.setWordWrap(True);contents.addWidget(home._fav_hint)
    home.fav_area=CardGrid();home.fav_area.setObjectName('favoriteCards');contents.addWidget(home.fav_area)
    home._recent_title=heading('最近使用');home.recent_area=CardGrid();contents.addWidget(home.recent_area)
    home._recent_hint=QLabel();home._recent_hint.hide()
    home._tool_title=heading('全部工具');home.tool_area=CardGrid();contents.addWidget(home.tool_area)
    home.no_results=QLabel('没有找到相关工具，换个关键词试试。');home.no_results.setProperty('role','emptyHint');contents.addWidget(home.no_results)
    heading('工作流程与记录');shortcuts=QGridLayout()
    for i,(name,target,op) in enumerate(COMMON_ENTRIES):
        button=QPushButton(name);button.clicked.connect(lambda _=False,p=target,o=op:home.main.goto_feature(p,o))
        shortcuts.addWidget(button,i//3,i%3);home._common_buttons.append((button,name))
    contents.addLayout(shortcuts);contents.addStretch();scroll.setWidget(inner);outer.addWidget(scroll,1)
    home.search_edit.textChanged.connect(lambda:refresh_tools(home));home.category.currentTextChanged.connect(lambda:refresh_tools(home))
    home._filter_features=lambda text:refresh_tools(home)
    refresh_home(home)


def refresh_tools(home):
    query=home.search_edit.text().strip().casefold();category=home.category.currentText()
    favorites={(x['page'],x['name']) for x in home._fav.load()}
    entries=[e for e in home._entries if (category=='全部分类' or e[0]==category) and (not query or query in ' '.join(e).casefold())]
    cards=[];home._buttons={}
    for page,op,hint in entries:
        card=ToolCard(home,page,op,hint,(page,op) in favorites);cards.append(card);home._buttons[(page,op)]=(card.open,page,op,page)
    home.tool_area.set_cards(cards);home._tool_title.setText(f'{category if category!="全部分类" else "全部工具"}  ·  {len(entries)}')
    home.no_results.setVisible(not entries);searching=bool(query) or category!='全部分类'
    for w in (home._fav_title,home._fav_hint,home.fav_area,home._recent_title,home.recent_area):
        w.setVisible(not searching and (w is not home._fav_hint or not favorites))
    if not searching:home._recent_title.setVisible(bool(home.recent_area.cards))


def refresh_favorites(home):
    known={(p,o) for p,o,h in home._entries}
    cards=[ToolCard(home,x['page'],x['name'],x['page'],True) for x in home._fav.load() if (x['page'],x['name']) in known]
    home.fav_area.set_cards(cards);home._fav_title.setText(f'我的收藏  ·  {len(cards)}');home._fav_hint.setVisible(not cards)


def refresh_recent(home):
    known={(p,o) for p,o,h in home._entries};favorites={(x['page'],x['name']) for x in home._fav.load()}
    entries=[(p,o) for p,o,*_ in load_recent(home.main.RECORDS_DIR,6) if (p,o) in known]
    home.recent_area.set_cards([ToolCard(home,p,o,p,(p,o) in favorites) for p,o in entries]);home._recent_title.setVisible(bool(entries))


def refresh_home(home):
    refresh_favorites(home);refresh_recent(home);refresh_tools(home)


def toggle_favorite(home,name,page,op,remove):
    try:
        if remove:home._fav.remove(name,page)
        else:home._fav.add(name,page,op)
    except OSError:
        QMessageBox.warning(home,home.main.APP_NAME,'收藏未能保存，请检查用户数据目录是否可写。')
        QTimer.singleShot(0,lambda:refresh_home(home))
        return
    QTimer.singleShot(0,lambda:refresh_home(home))
    home.main.show_status(('已取消收藏：' if remove else '已收藏：')+name)
