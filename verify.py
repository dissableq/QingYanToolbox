from pathlib import Path
import os, sys, subprocess, json, time
sys.stdout.reconfigure(encoding='utf-8')
root = Path(__file__).resolve().parent
phase = sys.argv[1] if len(sys.argv) > 1 else 'verification'
logs = root / '.verification_data' / phase
logs.mkdir(parents=True, exist_ok=True)
names = sys.argv[2:] or ['test_engine.py','test_extra.py','test_flow.py','test_convert.py',
    'test_regress.py','test_regress2.py','test_print_virtual.py','test_installer_args.py',
    'test_gui_integration.py','test_offline.py','test_quality.py','test_runtime_quality.py','test_surface.py','test_v33_regressions.py','test_v34.py']
report=[]
for name in names:
    env=os.environ.copy()
    env.update(APPDATA=str(logs.resolve() / (name + '_data')),
               LOCALAPPDATA=str(logs.resolve() / 'local'), QT_QPA_PLATFORM='offscreen',
               PYTHONUTF8='1', PYTHONDONTWRITEBYTECODE='1')
    start=time.monotonic()
    print('RUN', name, flush=True)
    try:
        p=subprocess.run([sys.executable, '-B', str(root/name)], cwd=root, env=env,
                         capture_output=True, timeout=240, encoding='utf-8', errors='replace')
        output=p.stdout+'\n'+p.stderr
        rc=p.returncode
    except subprocess.TimeoutExpired as e:
        output='TIMEOUT\n'+str(e.stdout)+str(e.stderr); rc=-1
    (logs/(name+'.log')).write_text(output,encoding='utf-8')
    summary=[s for s in output.splitlines() if any(k in s for k in
                ('RESULT:', 'FAIL', 'ALL PASS', 'CLEAN', 'Traceback', 'Error:'))]
    passed=rc==0 and not any(s.startswith('FAIL') for s in output.splitlines())
    report.append(dict(test=name,exit=rc,passed=passed,seconds=round(time.monotonic()-start,1),summary=summary))
    print(json.dumps(report[-1],ensure_ascii=False),flush=True)
    (logs/'summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
sys.exit(0 if all(r['passed'] for r in report) else 1)
