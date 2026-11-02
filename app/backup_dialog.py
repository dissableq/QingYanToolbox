"""Review managed backups and decide what to keep on exit."""
import os
from pathlib import Path
from PySide6.QtWidgets import (QDialog,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,
    QTableWidget,QTableWidgetItem,QHeaderView,QFileDialog,QMessageBox)
from engine import backups


class BackupDialog(QDialog):
    def __init__(self,main,exiting=False):
        super().__init__(main);self.main=main;self.exiting=exiting
        self.setObjectName('backupDialog');self.setWindowTitle('退出前保留备份' if exiting else '原文件备份管理')
        self.resize(820,540);layout=QVBoxLayout(self);layout.setSpacing(14)
        self.info=QLabel();self.info.setWordWrap(True);layout.addWidget(self.info)
        note=QLabel('备份保存在原文件所在目录的“备份”文件夹。清理只删除下表记录的备份；清理后无法通过这些备份恢复原文件内容。')
        note.setWordWrap(True);layout.addWidget(note)
        self.table=QTableWidget(0,3);self.table.setObjectName('backupTable')
        self.table.setHorizontalHeaderLabels(['原文件 / 任务','备份位置','大小'])
        self.table.horizontalHeader().setSectionResizeMode(0,QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1,QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers);layout.addWidget(self.table,1)
        row=QHBoxLayout()
        def button(text,slot,accent=False):
            b=QPushButton(text);b.clicked.connect(slot);b.setProperty('accent',accent);row.addWidget(b);return b
        button('打开选中备份目录',self.open_folder)
        if not exiting:button('整理已有备份…',self.migrate)
        row.addStretch()
        if exiting:
            button('返回软件',self.reject)
            button('清理并退出',self.clean_exit)
            keep=button('保留并退出',self.accept,True);keep.setDefault(True)
        else:
            button('清理选中备份',self.clean_selected)
            button('关闭',self.accept,True)
        layout.addLayout(row);self.refresh()

    def refresh(self):
        self.rows=backups.existing(current_session=self.exiting)
        size=sum(r['bytes'] for r in self.rows)/1024/1024
        self.info.setText(f'本次使用产生或整理了 {len(self.rows)} 份备份（{size:.2f} MB），是否保留？' if self.exiting
                          else f'已记录 {len(self.rows)} 份备份（{size:.2f} MB）。可选择需要清理的项目。')
        self.table.setRowCount(len(self.rows))
        for i,r in enumerate(self.rows):
            original=Path(r['original']).name if r['original'] else '任务记录'
            if r.get('kind')=='journal':original+='（名称与任务记录）'
            path=Path(r['path']);location=str(Path(path.parent.parent.name)/'备份'/path.name)
            size=f"{r['bytes']} B" if r['bytes']<1024 else f"{r['bytes']/1024:.1f} KB"
            for j,value in enumerate((original,location,size)):
                cell=QTableWidgetItem(value);cell.setToolTip((r['original'],r['path'],size)[j]);self.table.setItem(i,j,cell)
        self.table.resizeRowsToContents()

    def open_folder(self):
        i=self.table.currentRow()
        if i<0 and self.rows:i=0
        if i>=0:
            try:os.startfile(str(Path(self.rows[i]['path']).parent))
            except OSError as exc:QMessageBox.warning(self,'无法打开',str(exc))

    def migrate(self):
        directory=QFileDialog.getExistingDirectory(self,'选择需要整理旧备份的文件夹')
        if not directory:return
        try:
            rows=backups.migrate_legacy(directory,recursive=True)
            QMessageBox.information(self,'整理完成',f'已把 {len(rows)} 份旧版备份移入各自的“备份”文件夹。')
            self.refresh()
        except Exception as exc:QMessageBox.warning(self,'备份整理未完成',str(exc))

    def clean(self,rows):
        try:
            removed,errors=backups.cleanup({r['id'] for r in rows})
            paths={r['path'] for r in removed}
            for page in self.main.pages:
                actions=getattr(page,'_banner_undo_actions',None)
                if actions is not None:page._banner_undo_actions=backups.without_cleaned_actions(actions,paths)
                if hasattr(page,'undo_actions'):page.undo_actions=backups.without_cleaned_actions(page.undo_actions,paths)
            self.refresh()
            if errors:QMessageBox.warning(self,'部分备份已保留','\n'.join(errors[:8]));return False
            return True
        except Exception as exc:QMessageBox.warning(self,'清理未完成',str(exc));return False

    def clean_exit(self):
        if self.clean(self.rows):self.accept()

    def clean_selected(self):
        selected=sorted({i.row() for i in self.table.selectedIndexes()})
        if not selected:QMessageBox.information(self,'选择备份','请先选中需要清理的备份。');return
        if QMessageBox.question(self,'清理备份',f'删除选中的 {len(selected)} 份备份？删除后无法恢复对应的原文件内容。',
                                QMessageBox.Yes|QMessageBox.No,QMessageBox.No)==QMessageBox.Yes:
            self.clean([self.rows[i] for i in selected])


def allow_exit(main):
    try:
        if not backups.existing(current_session=True):return True
        return BackupDialog(main,True).exec()==QDialog.Accepted
    except Exception as exc:
        QMessageBox.warning(main,'备份保留','备份清单暂时无法读取，备份已保留。\n'+str(exc))
        return True
