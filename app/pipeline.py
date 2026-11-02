"""Sequential background pipeline controller. All widget access stays on the GUI thread."""
import os
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QFileDialog, QMessageBox
from engine.core import BatchWorker, TaskSpec, collect_files
from engine.store import _atomic_write


def start(page):
    if getattr(page.main, '_task_running', False) or any(getattr(p,'worker',None) is not None and p.worker.isRunning() for p in page.main.pages):
        QMessageBox.information(page,page.main.APP_NAME,'请等待当前任务完成。'); return
    steps = [dict(s) for s in page._steps]
    if not steps:
        QMessageBox.information(page,page.main.APP_NAME,'请先添加至少一个步骤'); return
    first = steps[0].get('page_obj')
    if first is None:
        QMessageBox.warning(page,page.main.APP_NAME,'步骤对应的工具不存在'); return
    snapshot = first.capture_state()
    try:
        first.restore_state(steps[0].get('ui_state') or steps[0].get('options',{}))
        files = [p.strip() for p in first.file_edit.text().split('||') if os.path.exists(p.strip())]
        if not hasattr(first, 'DIR_LEVEL_OPS') and not hasattr(first, '_collect_mode_options'):
            files = first.current_files()
    finally:
        first.restore_state(snapshot)
    if not files:
        directory = QFileDialog.getExistingDirectory(page,'选择第一步的输入文件夹')
        if not directory: return
        files = [directory]
    if not files:
        QMessageBox.information(page,page.main.APP_NAME,'没有可处理的文件'); return
    page._pipeline_active = True
    page._pipeline_state = dict(steps=steps,index=0,files=files,undo=[],results=[],cancelled=False)
    page.run_btn.setEnabled(False)
    page.cancel_pipeline_btn.setEnabled(True)
    page.main.set_running(True,'任务流水线')
    advance(page)


def advance(page):
    state = page._pipeline_state
    if state['cancelled'] or state['index'] >= len(state['steps']):
        finish(page); return
    if not state['files']:
        finish(page,'上一步没有可用输出，后续步骤已停止'); return
    step = state['steps'][state['index']]
    tool = step.get('page_obj')
    if tool is None:
        finish(page,'步骤对应的工具不存在'); return
    snapshot = tool.capture_state()
    try:
        tool.restore_state(step.get('ui_state') or step.get('options',{}))
        missing = tool.validate_required()
        if missing: raise ValueError('请补充参数：'+'、'.join(missing))
        options = tool.collect_options()
        options['op_name'] = step.get('op') or options.get('op_name', '')
        proc = tool.processor()
        if proc is None: raise ValueError('该步骤没有可执行操作')
        proc = tool._guarded(proc)
        options['out_dir'] = options.get('out_dir') or os.path.join(
            os.path.dirname(state['files'][0]), f"qingyan_pipe_{state['index']+1}")
        options['op_name'] = step.get('op') or options.get('op_name','')
        files = list(state['files'])
        op = options.get('op_name', '')
        dir_op = op in getattr(tool, 'DIR_LEVEL_OPS', ())
        batch_op = op in ('打包为 zip', '多源复制到一处')
        dir_mode = options.get('mode') in ('dir_seq', 'excel_dir')
        if not dir_op and not batch_op and not dir_mode:
            files = list(dict.fromkeys(f for p in files for f in
                (collect_files(p, True) if os.path.isdir(p) else [p])))
        if batch_op:
            options['_src_paths'] = files
            files = files[:1]
        if dir_mode:
            files = [f for p in files for f in
                ([os.path.join(p, n) for n in sorted(os.listdir(p))
                  if os.path.isdir(os.path.join(p, n))] if os.path.isdir(p) else [])]
        if op == '冗余命名清理':
            from engine import ops
            options['_rename_plan'] = ops.build_rename_plan(files,
                {'mode': 'cleanup', 'cleanup_patterns': ops.CLEANUP_PATTERNS})
        if hasattr(tool,'_collect_mode_options'):
            from engine import ops
            if options.get('mode') in ('excel','excel_dir'):
                options['_excel_map'] = ops.load_excel_map(options['excel_map'])
            else:
                options['_rename_plan'] = ops.build_rename_plan(files,options)
            tool._dir_counter = 0
        if hasattr(tool,'_gen_proc'):
            import openpyxl
            wb = openpyxl.load_workbook(options['data'],read_only=True,data_only=True)
            try:
                values = list(wb.active.values)
                headers = [str(x or '') for x in values[0]]
                options['_rows'] = [dict(zip(headers,row)) for row in values[1:] if any(v is not None for v in row)]
            finally: wb.close()
            files = [f'row{i}' for i in range(len(options['_rows']))]
            def proc(path, options, generator=tool._gen_proc):
                return generator(path,{**options,'_row_idx':int(path[3:])})
    except Exception as e:
        finish(page,f'步骤 {state["index"]+1} 未执行：{e}'); return
    finally:
        tool.restore_state(snapshot)
    worker = BatchWorker(TaskSpec('pipe',step['page'],files,options),proc,page.main.RECORDS_DIR,None)
    page.worker = worker
    worker.progress.connect(page.main.on_progress)
    def results_ready(results, task_id):
        state['results'].extend(results); state['undo'].extend(worker.undo_actions)
        outputs=[]
        for result in results:
            if result.status == 'skipped' and os.path.exists(result.path):
                outputs.append(result.path)
                continue
            if result.status != 'ok': continue
            if '删除文件' in op or op.startswith('清理') or op == '自动清理重复文件':
                if os.path.exists(result.path): outputs.append(result.path)
                continue
            for output in (result.outputs or ([result.output] if result.output else [])):
                if os.path.isdir(output): outputs.append(output)
                elif os.path.isfile(output): outputs.append(output)
        state['files'] = list(dict.fromkeys(outputs))
        state['index'] += 1
    def cancelled():
        state['results'].extend(worker.results); state['undo'].extend(worker.undo_actions)
        state['cancelled'] = True
    worker.finished_all.connect(results_ready)
    worker.cancelled.connect(cancelled)
    worker.finished.connect(lambda:QTimer.singleShot(0,lambda:advance(page)))
    page.status.setText(f'正在执行步骤 {state["index"]+1}/{len(state["steps"])} · {step.get("op",step["page"])}')
    worker.start()


def finish(page, error=''):
    from uuid import uuid4
    state = page._pipeline_state
    try:
        _atomic_write(page.main.UNDO_MANIFEST, {'task_id':'pipeline_'+uuid4().hex,'actions':state['undo']})
    except OSError as exc:
        error = '撤销记录保存失败：' + str(exc)
    page.main.records_dirty = True
    page.main.set_running(False)
    page.run_btn.setEnabled(True); page.cancel_pipeline_btn.setEnabled(False)
    ok = sum(r.status=='ok' for r in state['results'])
    failed = sum(r.status=='fail' for r in state['results'])
    label = '已取消' if state['cancelled'] else '已完成'
    if error or failed: label = '部分未完成'
    page.status.setText(f'{label} · 成功 {ok} / 失败 {failed}' + (' · '+error if error else ''))
    page._pipeline_active = False


def cancel(page):
    state = getattr(page,'_pipeline_state',None)
    if state is None: return
    state['cancelled'] = True
    if getattr(page,'worker',None) is not None: page.worker.cancel()
