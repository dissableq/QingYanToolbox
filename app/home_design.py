from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QVBoxLayout,QHBoxLayout,QGridLayout,QLabel,QLineEdit,
                              QComboBox,QPushButton,QWidget,QScrollArea)
from engine.store import Favorites
from app.home_registry import COMMON_ENTRIES


def build_home(self):
    self._fav=Favorites(); self._buttons={}; self._group_caps={}; self._common_buttons=[]
    v=QVBoxLayout(self); v.setContentsMargins(30,26,30,20); v.setSpacing(14)
    eyebrow=QLabel('QINGYAN  /  FILE WORKSPACE'); eyebrow.setProperty('role','eyebrow'); v.addWidget(eyebrow)
    title=QLabel('把重复工作，交给青雁。'); title.setProperty('role','title'); v.addWidget(title)
    intro=QLabel('文档、图片与文件整理，一个工作台完成。选择工具 → 添加文件 → 预览并执行。')
    intro.setProperty('role','muted'); intro.setWordWrap(True); v.addWidget(intro)
    row=QHBoxLayout()
    self.search_edit=QLineEdit(); self.search_edit.setObjectName('featureSearch')
    self.search_edit.setPlaceholderText('搜索全部工具，例如：水印、转 PDF、重命名…'); self.search_edit.setClearButtonEnabled(True)
    row.addWidget(self.search_edit,1)
    self.category=QComboBox(); self.category.addItem('全部分类'); self.category.setMinimumWidth(140); row.addWidget(self.category)
    v.addLayout(row)
    scroll=QScrollArea(); scroll.setWidgetResizable(True)
    inner=QWidget(); contents=QVBoxLayout(inner); contents.setContentsMargins(0,6,8,0); contents.setSpacing(14)
    self.fav_area=QWidget(); self.fav_layout=QGridLayout(self.fav_area)
    self._fav_hint=QLabel('常用工具可右键收藏，下次更快找到。'); self._fav_hint.setProperty('role','muted')
    contents.addWidget(self.fav_area); contents.addWidget(self._fav_hint)
    self.recent_layout=QGridLayout(); contents.addLayout(self.recent_layout)
    self._recent_hint=QLabel(''); self._recent_hint.hide(); contents.addWidget(self._recent_hint)
    self._rebuild_favs(); self._rebuild_recent()
    for page in self.main.pages:
        if not hasattr(page,'OPS'):
            continue
        combo=page._op_widget()
        operations=[combo.itemText(i) for i in range(combo.count())] if combo else list(page.OPS)
        group=page.title; self.category.addItem(group)
        cap=QLabel(f'{group}  ·  {len(operations)}'); cap.setProperty('role','section')
        self._group_caps[group]=cap; contents.addWidget(cap)
        grid=QGridLayout(); grid.setSpacing(10)
        for j,name in enumerate(operations):
            button=QPushButton(name+'  →'); button.setProperty('feature',True)
            button.setMinimumHeight(48); button.setCursor(Qt.PointingHandCursor)
            button.setToolTip(getattr(page,'OP_HINTS',{}).get(name,name))
            button.clicked.connect(lambda _=False,p=group,o=name:self.main.goto_feature(p,o))
            button.setContextMenuPolicy(Qt.CustomContextMenu)
            button.customContextMenuRequested.connect(lambda pos,b=button,n=name,g=group:self._fav_menu(pos,b,n,g,n))
            self._buttons[(group,name)]=(button,group,name,group)
            grid.addWidget(button,j//3,j%3)
        for i in range(3):grid.setColumnStretch(i,1)
        contents.addLayout(grid)
    more=QLabel('工作流程与记录'); more.setProperty('role','section'); contents.addWidget(more)
    grid=QGridLayout()
    for i,(name,target,op) in enumerate(COMMON_ENTRIES):
        b=QPushButton(name); b.setProperty('feature',True)
        b.clicked.connect(lambda _=False,p=target,o=op:self.main.goto_feature(p,o))
        self._common_buttons.append((b,name)); grid.addWidget(b,i//4,i%4)
    contents.addLayout(grid); contents.addStretch(1)
    self.no_results=QLabel('没有找到匹配的工具，试试其他关键词。'); self.no_results.setProperty('role','muted'); self.no_results.hide(); contents.addWidget(self.no_results)
    scroll.setWidget(inner); v.addWidget(scroll,1)
    def search():
        text=self.search_edit.text().strip().casefold(); category=self.category.currentText(); counts={}
        for b,g,n,target in self._buttons.values():
            visible=(category=='全部分类' or category==g) and (not text or text in (g+' '+n).casefold())
            b.setVisible(visible); counts[g]=counts.get(g,0)+int(visible)
        for g,cap in self._group_caps.items(): cap.setVisible(counts.get(g,0)>0)
        for b,n in self._common_buttons:b.setVisible(category=='全部分类' and (not text or text in n.casefold()))
        self.no_results.setVisible(sum(counts.values())==0)
    self.search_edit.textChanged.connect(search); self.category.currentTextChanged.connect(search)
    self._filter_features=lambda text:search()
