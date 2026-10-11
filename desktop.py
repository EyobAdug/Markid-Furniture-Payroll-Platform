"""Desktop launcher: runs Markid Payroll in its own window. Works fully offline; syncs with the online server when configured."""
import os, sys, socket, threading, time, secrets, webbrowser

def data_dir():
    if sys.platform == 'win32': base = os.environ.get('APPDATA') or os.path.expanduser('~')
    elif sys.platform == 'darwin': base = os.path.expanduser('~/Library/Application Support')
    else: base = os.environ.get('XDG_DATA_HOME') or os.path.expanduser('~/.local/share')
    d = os.path.join(base, 'MarkidPayroll'); os.makedirs(d, exist_ok=True); return d

def main():
    d = data_dir()
    os.environ['MARKID_MODE'] = 'client'
    os.environ.setdefault('MARKID_DB', os.path.join(d, 'markid.db'))
    os.environ.setdefault('MARKID_STATE', os.path.join(d, 'sync_state.db'))
    kf = os.path.join(d, 'secret.key')
    if not os.path.exists(kf): open(kf, 'w').write(secrets.token_hex(32))
    os.environ['SECRET_KEY'] = open(kf).read().strip()
    import app as A
    s = socket.socket(); s.bind(('127.0.0.1', 0)); port = s.getsockname()[1]; s.close()
    from werkzeug.serving import make_server
    srv = make_server('127.0.0.1', port, A.app, threaded=True)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    A.start_autosync()
    url = f'http://127.0.0.1:{port}'
    print(f'Markid Payroll running at {url}  (data in {d})', flush=True)
    try:
        import webview
        try: webview.settings['ALLOW_DOWNLOADS'] = True          # lets the Excel exports be saved
        except Exception: pass
        webview.create_window('Markid Payroll', url, width=1360, height=860, min_size=(1000, 640))
        webview.start()
    except Exception as e:                                       # no native window available: use the normal browser
        print('Native window unavailable (%s) - opening your browser instead.' % e, flush=True)
        webbrowser.open(url)
        while True: time.sleep(3600)
    srv.shutdown()

if __name__ == '__main__': main()
