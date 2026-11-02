"""A single previewable task for content and name replacement."""
from pathlib import Path
import json
from PySide6.QtCore import QThread,Signal,QEvent
from PySide6.QtWidgets import (QWidget,QVBoxLayout,QHBoxLayout,QLabel,QLineEdit,QPushButton,
    QCheckBox,QTableWidget,QTableWidgetItem,QHeaderView,QFileDialog,QMessageBox,QGroupBox,QScrollArea,QLayout)
from app.layouts import WrapRow
from engine import unified_replace as engine
from engine.store import ParamStore
from engine.undo import undo_actions


class UnifiedWorker(QThread):
    progress=Signal(int,int,str,float)
    ready=Signal(object)
    error=Signal(str)
    def __init__(self,options,plan=None):
        super().__init__();self.options=options;self.plan=plan;self.stopped=False
    def cancel(self):self.stopped=True
    def run(self):
        try:
            result=engine.make_plan(self.options,lambda:self.stopped,self.progress.emit) if self.plan is None else engine.execute(self.plan,lambda:self.stopped,self.progress.emit)
            self.ready.emit(result)
        except engine.Cancelled:self.error.emit('已取消，未开始的项目保持原样')
        except Exception as exc:self.error.emit(str(exc))


class UnifiedPage(QWidget):
    title='统一替换';kind='unified';OPS=('内容与名称统一替换',)
    OP_HINTS={'内容与名称统一替换':'多组规则同时处理文件内容、文件名与文件夹名'}
    def _op_widget(self):return None
    def __init__(self,main):
        super().__init__();self.main=main;self.worker=None;self.plan=None;self.undo_actions=[]
        self.setAcceptDrops(True);self.setObjectName('unifiedPage')
        layout=QVBoxLayout(self);layout.setContentsMargins(24,20,24,18);layout.setSpacing(12)
        title=QLabel('一个任务，统一修改内容与名称');title.setProperty('role','title');layout.addWidget(title)
        hint=QLabel('先预览，再执行。多组规则从上到下依次替换；文件扩展名保持不变，备份目录自动跳过。')
        hint.setWordWrap(True);layout.addWidget(hint)
        folder=QHBoxLayout();self.root_edit=QLineEdit();self.root_edit.setObjectName('unified_root');self.root_edit.setPlaceholderText('选择目标文件夹，或拖入文件夹')
        folder.addWidget(self.root_edit,1);pick=QPushButton('选择文件夹');self.pick_btn=pick;pick.clicked.connect(self.pick);folder.addWidget(pick);layout.addLayout(folder)
        scope=WrapRow()
        self.checks={}
        for key,label,default in [('contents','文件内容',True),('file_names','文件名',True),('folder_names','文件夹名',True),('recursive','包含子文件夹',True),('include_root','也修改所选文件夹本身',False)]:
            cb=QCheckBox(label);cb.setObjectName('unified_'+key);cb.setChecked(default);cb.toggled.connect(self.invalidate);self.checks[key]=cb;scope.addWidget(cb)
        layout.addLayout(scope)
        formats=QLabel('内容支持 Word（docx/docm）、Excel（xlsx/xlsm）、PPT（pptx/pptm）和文本文件。PDF、图片、旧版 doc/xls/ppt 仅处理名称；Excel 公式保持不变。')
        formats.setProperty('role','muted');formats.setWordWrap(True);layout.addWidget(formats)
        scroll=QScrollArea();self.content_scroll=scroll;scroll.setObjectName('unified_contentScroll');scroll.setWidgetResizable(True)
        body=QWidget();sections=QVBoxLayout(body);sections.setContentsMargins(0,0,10,0);sections.setSpacing(12);sections.setSizeConstraint(QLayout.SetMinimumSize)
        rules_box=QGroupBox('01  替换规则');self.rules_box=rules_box;rv=QVBoxLayout(rules_box)
        self.rules=QTableWidget(0,2);self.rules.setObjectName('unified_rules');self.rules.setHorizontalHeaderLabels(['查找内容（区分大小写）','替换为（留空表示删除）'])
        self.rules.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch);self.rules.setMinimumHeight(130);rv.addWidget(self.rules)
        buttons=WrapRow()
        for label,slot in [('添加规则',self.add_rule),('删除选中规则',self.remove_rule),('上移',lambda:self.move_rule(-1)),('下移',lambda:self.move_rule(1)),('导入规则',self.import_rules),('导出规则',self.export_rules)]:
            b=QPushButton(label);b.clicked.connect(slot);buttons.addWidget(b)
        rv.addLayout(buttons);sections.addWidget(rules_box)
        result_box=QGroupBox('02  预览与结果');self.result_box=result_box;result_box.setMinimumHeight(240);pv=QVBoxLayout(result_box)
        self.summary=QLabel('填写规则后点击“预览统一任务”。');self.summary.setWordWrap(True);self.summary.setObjectName('unified_summary');pv.addWidget(self.summary)
        self.table=QTableWidget(0,4);self.table.setObjectName('unified_results');self.table.setHorizontalHeaderLabels(['类型 / 状态','原路径','完成后路径','内容 / 说明'])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers);self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1,QHeaderView.Stretch);self.table.horizontalHeader().setSectionResizeMode(2,QHeaderView.Stretch)
        self.table.setColumnWidth(0,100);self.table.setColumnWidth(3,230);self.table.setMinimumHeight(130);pv.addWidget(self.table);sections.addWidget(result_box,1)
        scroll.setWidget(body);layout.addWidget(scroll,1)
        actions=WrapRow()
        for attr,label,slot in [('preview_btn','预览统一任务',self.preview),('run_btn','执行统一任务',self.run_task),('cancel_btn','取消任务',self.cancel),('undo_btn','撤销本次统一任务',self.undo)]:
            b=QPushButton(label);b.setObjectName('unified_'+attr);b.clicked.connect(slot);setattr(self,attr,b);actions.addWidget(b)
        self.run_btn.setProperty('accent',True);self.run_btn.setEnabled(False);self.cancel_btn.setEnabled(False);self.undo_btn.setEnabled(False)
        layout.addLayout(actions)
        self.root_edit.textChanged.connect(self.invalidate)
        saved=ParamStore().load_ui().get('unified_task',{})
        self.root_edit.setText(saved.get('root',''))
        for key,cb in self.checks.items():cb.setChecked(saved.get(key,cb.isChecked()))
        for rule in saved.get('rules') or [{'find':'','replace':''}]:self.add_rule(rule)

    def values(self):
        return [{'find':self.rules.cellWidget(i,0).text(),'replace':self.rules.cellWidget(i,1).text()} for i in range(self.rules.rowCount())]
    def options(self):return dict(root=self.root_edit.text().strip(),rules=self.values(),**{k:cb.isChecked() for k,cb in self.checks.items()})
    def invalidate(self,*_):
        self.plan=None
        if hasattr(self,'run_btn'):self.run_btn.setEnabled(False)
        if hasattr(self,'summary'):self.summary.setText('设置已变化，请重新预览。')
    def add_rule(self,rule=None):
        interactive=isinstance(rule,bool)
        rule=rule if isinstance(rule,dict) else {'find':'','replace':''};row=self.rules.rowCount();self.rules.insertRow(row)
        for j,key in enumerate(('find','replace')):
            edit=QLineEdit(rule[key]);edit.setObjectName(f'unified_{key}_{row}');edit.setProperty('rule_cell',(row,j));edit.installEventFilter(self);edit.textChanged.connect(self.invalidate);self.rules.setCellWidget(row,j,edit)
        self.rules.setRowHeight(row,42);self.invalidate()
        if interactive:
            self.rules.scrollToBottom();self.rules.cellWidget(row,0).setFocus()
    def eventFilter(self,watched,event):
        cell=watched.property('rule_cell')
        if cell is not None and event.type()==QEvent.FocusIn:self.rules.setCurrentCell(*cell)
        return super().eventFilter(watched,event)
    def set_rules(self,rules):
        self.rules.setRowCount(0)
        for rule in rules:self.add_rule(rule)
    def remove_rule(self):
        row=self.rules.currentRow()
        if row>=0:
            rules=self.values();rules.pop(row);self.set_rules(rules or [{'find':'','replace':''}])
    def move_rule(self,delta):
        row=self.rules.currentRow();new=row+delta
        if 0<=row<self.rules.rowCount() and 0<=new<self.rules.rowCount():
            rules=self.values();rules[row],rules[new]=rules[new],rules[row];self.set_rules(rules);self.rules.setCurrentCell(new,0)
    def pick(self):
        folder=QFileDialog.getExistingDirectory(self,'选择统一替换的目标文件夹')
        if folder:self.root_edit.setText(folder)
    def import_rules(self):
        path,_=QFileDialog.getOpenFileName(self,'导入替换规则','','规则文件 (*.json)')
        if not path:return
        try:
            rows=json.loads(Path(path).read_text(encoding='utf-8'))
            if not isinstance(rows,list) or not rows or not all(isinstance(r,dict) and isinstance(r.get('find'),str) and r['find'] and isinstance(r.get('replace'),str) for r in rows):raise ValueError('规则需要包含非空的 find 和文字 replace')
            self.set_rules(rows)
        except Exception as exc:QMessageBox.warning(self,'导入失败',str(exc))
    def export_rules(self):
        path,_=QFileDialog.getSaveFileName(self,'导出替换规则','替换规则.json','规则文件 (*.json)')
        if path:
            try:Path(path).write_text(json.dumps(self.values(),ensure_ascii=False,indent=2),encoding='utf-8')
            except OSError as exc:QMessageBox.warning(self,'导出失败',str(exc))
    def busy(self,on):
        for w in (self.root_edit,self.pick_btn,self.rules_box,self.preview_btn,*self.checks.values()):w.setEnabled(not on)
        self.run_btn.setEnabled(not on and self.plan is not None and not any(r['error'] for r in self.plan['rows']))
        self.undo_btn.setEnabled(not on and bool(self.undo_actions))
        self.cancel_btn.setEnabled(on);self.main.set_running(on,'统一替换')
    def start(self,plan=None):
        if getattr(self.main,'_task_running',False):QMessageBox.information(self,'任务进行中','请等待当前任务完成。');return
        try:
            options=self.options();engine.validate_options(options);ParamStore().save_ui({'unified_task':options})
        except Exception as exc:QMessageBox.warning(self,'检查设置',str(exc));return
        self.worker=UnifiedWorker(options,plan);self.worker.progress.connect(self.main.on_progress)
        self.worker.ready.connect(self.completed);self.worker.error.connect(self.failed)
        self.busy(True);self.worker.start()
    def preview(self):self.plan=None;self.start()
    def run_task(self):
        if self.plan is None:return
        changes=sum(r['content'] or r['name']!=r['new_name'] for r in self.plan['rows'])
        if not changes:QMessageBox.information(self,'无需处理','没有找到需要修改的内容或名称。');return
        if QMessageBox.question(self,'执行统一替换',f'将修改 {changes} 个文件或文件夹。原文件内容先备份到各自的“备份”文件夹，随后修改内容与名称。\n确认按当前预览执行？',QMessageBox.Yes|QMessageBox.No,QMessageBox.No)==QMessageBox.Yes:self.start(self.plan)
    def cancel(self):
        if self.worker:self.worker.cancel();self.summary.setText('正在取消，当前文件完成保存后停止。')
    def fill(self,rows):
        self.table.setRowCount(len(rows))
        for i,row in enumerate(rows):
            for j,value in enumerate(row):
                display=str(value)
                if j in (1,2) and value:
                    try:display=str(Path(value).relative_to(Path(self.root_edit.text()).parent))
                    except ValueError:display=str(Path(value).name)
                item=QTableWidgetItem(display);item.setToolTip(str(value));self.table.setItem(i,j,item)
        self.table.resizeRowsToContents()
    def completed(self,result):
        if isinstance(result,dict):
            self.plan=result;rows=result['rows'];changes=sum(r['content'] or r['new_name']!=r['name'] for r in rows);errors=sum(bool(r['error']) for r in rows)
            unsupported=sum(bool(r['note']) for r in rows)
            self.fill([('错误' if r['error'] else '文件夹' if r['directory'] else '文件',r['path'],r['final_path'],r['error'] or (f"内容 {r['hits']} 处；" if r['content'] else '')+(r['note'] or '')) for r in rows if r['content'] or r['new_name']!=r['name'] or r['error'] or r['note']])
            self.summary.setText(f'已扫描 {len(rows)} 项：{changes} 项将修改，{errors} 项错误，{unsupported} 项格式仅支持改名。预览未修改文件。')
        else:
            results,self.undo_actions,cancelled,task_id=result;self.plan=None
            self.fill([(dict(ok='成功',fail='失败',skipped='跳过')[r.status],r.path,r.output,r.detail) for r in results])
            self.summary.setText(('已取消' if cancelled else '任务完成')+f"：成功 {sum(r.status=='ok' for r in results)} 步，失败 {sum(r.status=='fail' for r in results)} 步。已完成的操作可撤销。")
            self.undo_btn.setEnabled(bool(self.undo_actions));self.main.records_dirty=True
        self.busy(False)
        self.main.show_status(self.summary.text())
        self.content_scroll.ensureWidgetVisible(self.result_box)
    def failed(self,message):
        self.plan=None;self.busy(False);self.summary.setText(message);self.main.show_status(message);QMessageBox.warning(self,'任务提示',message)
    def undo(self):
        if getattr(self.main,'_task_running',False):return
        if not self.undo_actions:return
        completed,remaining,errors=undo_actions(self.undo_actions);self.undo_actions=remaining;self.undo_btn.setEnabled(bool(remaining))
        from engine.store import data_dir,_atomic_write
        p=Path(data_dir())/'last_undo.json'
        if p.exists():
            data=json.loads(p.read_text(encoding='utf-8'));data['actions']=[a for a in data.get('actions',[]) if a not in completed];_atomic_write(str(p),data)
        self.invalidate();self.summary.setText(f'已撤销 {len(completed)} 步。'+('\n'.join(errors) if errors else '可用的撤销步骤已完成。'));self.main.show_status(self.summary.text())
    def dragEnterEvent(self,event):
        if event.mimeData().hasUrls():event.acceptProposedAction()
    def dropEvent(self,event):
        paths=[u.toLocalFile() for u in event.mimeData().urls()]
        if paths and Path(paths[0]).is_dir() and not getattr(self.main,'_task_running',False):self.root_edit.setText(paths[0]);event.acceptProposedAction()
