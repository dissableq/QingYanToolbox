"""A toolbar that wraps controls at narrow window widths."""
from PySide6.QtCore import QRect, QSize, Qt
from PySide6.QtWidgets import QLayout


class WrapRow(QLayout):
    def __init__(self, parent=None, spacing=8):
        super().__init__(parent)
        self.items = []
        self.setContentsMargins(0, 0, 0, 0)
        self.setSpacing(spacing)

    def addItem(self, item): self.items.append(item)
    def count(self): return len(self.items)
    def itemAt(self, index): return self.items[index] if 0 <= index < len(self.items) else None
    def takeAt(self, index): return self.items.pop(index) if 0 <= index < len(self.items) else None
    def addStretch(self, stretch=0): pass
    def expandingDirections(self): return Qt.Orientations(0)
    def hasHeightForWidth(self): return True
    def heightForWidth(self, width): return self._arrange(QRect(0, 0, width, 0), False)
    def sizeHint(self): return self.minimumSize()

    def minimumSize(self):
        size = QSize(0, 0)
        for item in self.items:
            if not item.isEmpty():size = size.expandedTo(item.minimumSize())
        l,t,r,b = self.getContentsMargins()
        return size + QSize(l+r,t+b)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._arrange(rect, True)

    def _arrange(self, rect, place):
        l,t,r,b = self.getContentsMargins()
        bounds=rect.adjusted(l,t,-r,-b);x=bounds.x();y=bounds.y();height=0
        for item in self.items:
            if item.isEmpty():continue
            size=item.sizeHint();size.setWidth(min(size.width(),max(1,bounds.width())))
            if x+size.width()>bounds.right()+1 and height:
                x=bounds.x();y+=height+self.spacing();height=0
            if place:item.setGeometry(QRect(x,y,size.width(),size.height()))
            x+=size.width()+self.spacing();height=max(height,size.height())
        return y+height-rect.y()+b
