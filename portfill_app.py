#!/usr/bin/env python3
"""PortFill - the window. Runs a small local web page (only on this computer, no internet needed) and opens it
like an app. Start with:  python portfill_app.py
"""
import os, sys, json, shutil, tempfile, threading, time, subprocess, secrets, socket, webbrowser, platform, uuid
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from flask import Flask, request, jsonify, abort, Response
from master import Master
from forms import process_workbook
from portfill import fill_all

# When packaged (PyInstaller) the data files sit in sys._MEIPASS
HERE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
HOME = os.environ.get("PORTFILL_HOME") or os.path.join(os.path.expanduser("~"), ".portfill")
SETTINGS = os.path.join(HOME, "settings.json")
TOKEN = secrets.token_hex(16)
SESSION = tempfile.mkdtemp(prefix="portfill_")
FORMS = {}            # id -> {"name", "path", "analysis"}
LAST_PING = {"t": None}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# ---------------------------------------------------------------- settings
def default_out():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    return os.path.join(docs if os.path.isdir(docs) else os.path.expanduser("~"), "PortFill")

def load_settings():
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        d = {}
    d.setdefault("master_path", "")
    d.setdefault("out_root", default_out())
    return d

def save_settings(d):
    os.makedirs(HOME, exist_ok=True)
    with open(SETTINGS, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2)

# ---------------------------------------------------------------- master + analysis
def read_master(path):
    """Returns (summary dict, error text)."""
    if not path:
        return None, ""
    if not os.path.isfile(path):
        return None, "File not found. Check the path."
    try:
        m = Master(path)
    except Exception as e:
        return None, f"Can't read this file as a master ({e.__class__.__name__})."
    if "Ship" not in m.wb.sheetnames or "Voyage" not in m.wb.sheetnames:
        return None, "This doesn't look like the PortFill master (sheets 'Ship' and 'Voyage' not found)."
    eta = m.fields.get("eta")
    try:
        eta_txt = eta.strftime("%d-%m-%Y")
    except Exception:
        eta_txt = ""
    return {
        "path": path, "name": os.path.basename(path), "ship": str(m.fields.get("ship_name", "")),
        "imo": str(m.fields.get("imo", "")), "port": str(m.fields.get("port", "")), "eta": eta_txt,
        "folder": m.folder_name(), "crew": len(m.crew), "ports": len(m.ports), "sheets": len(m.wb.sheetnames),
    }, ""

def analyse_form(master_path, path):
    m = Master(master_path)
    tmp = tempfile.mkdtemp(prefix="pf_an_")
    try:
        _out, results, notes = process_workbook(m, path, tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    rec = [r.kind for r in results if r.kind != "not recognised"]
    return {"recognised": len(rec), "total": len(results), "kinds": rec, "notes": notes}

def state():
    s = load_settings()
    master, err = read_master(s["master_path"])
    forms = []
    for fid, f in FORMS.items():
        if master and f.get("analysis_for") != (s["master_path"], os.path.getmtime(s["master_path"])):
            try:
                f["analysis"] = analyse_form(s["master_path"], f["path"])
                f["error"] = ""
            except Exception as e:
                f["analysis"], f["error"] = None, f"Can't read this file ({e.__class__.__name__})."
            f["analysis_for"] = (s["master_path"], os.path.getmtime(s["master_path"]))
        forms.append({"id": fid, "name": f["name"], "analysis": f.get("analysis"), "error": f.get("error", ""),
                      "waiting": master is None})
    return {"master": master, "master_error": err, "master_path": s["master_path"], "out_root": s["out_root"],
            "forms": forms, "licence": "Test version: licence check not active yet",
            "target": os.path.join(s["out_root"], master["folder"]) if master else s["out_root"]}

# ---------------------------------------------------------------- native dialogs (best effort)
def native_dialog(kind):
    """kind: 'file' or 'folder'. Returns a path or None if no dialog is available / cancelled."""
    try:
        sysname = platform.system()
        if sysname == "Darwin":
            # No file-type filter: on some macOS versions a filter by extension greys out every Excel file.
            # The program checks the file itself afterwards.
            what = "choose file with prompt \"Choose the master Excel file\"" if kind == "file" else "choose folder with prompt \"Choose the results folder\""
            out = subprocess.run(["osascript", "-e", f"POSIX path of ({what})"], capture_output=True, text=True, timeout=600)
            return out.stdout.strip() or None
        if sysname == "Windows":
            if kind == "file":
                ps = ("Add-Type -AssemblyName System.Windows.Forms;$d=New-Object System.Windows.Forms.OpenFileDialog;"
                      "$d.Filter='Excel files|*.xlsx;*.xlsm';if($d.ShowDialog() -eq 'OK'){$d.FileName}")
            else:
                ps = ("Add-Type -AssemblyName System.Windows.Forms;$d=New-Object System.Windows.Forms.FolderBrowserDialog;"
                      "if($d.ShowDialog() -eq 'OK'){$d.SelectedPath}")
            out = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True, timeout=600)
            return out.stdout.strip() or None
        for cmd in (["zenity", "--file-selection"] + (["--directory"] if kind == "folder" else []),):
            out = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
            return out.stdout.strip() or None
    except Exception:
        return None
    return None

def open_folder(path):
    if platform.system() == "Windows":
        os.startfile(path)  # noqa
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])

# ---------------------------------------------------------------- routes
@app.before_request
def guard():
    if request.method != "GET" and request.headers.get("X-PortFill") != TOKEN:
        abort(403)

@app.route("/")
def index():
    with open(os.path.join(HERE, "ui.html"), encoding="utf-8") as f:
        html = f.read().replace("__TOKEN__", TOKEN)
    return Response(html, mimetype="text/html")

@app.route("/api/state")
def api_state():
    return jsonify(state())

@app.route("/api/ping", methods=["GET", "POST"])
def api_ping():
    LAST_PING["t"] = time.time()
    return jsonify(ok=True)

@app.route("/api/master", methods=["POST"])
def api_master():
    s = load_settings()
    s["master_path"] = (request.json or {}).get("path", "").strip().strip('"')
    save_settings(s)
    return jsonify(state())

@app.route("/api/outdir", methods=["POST"])
def api_outdir():
    s = load_settings()
    p = (request.json or {}).get("path", "").strip().strip('"')
    if p:
        s["out_root"] = p
        save_settings(s)
    return jsonify(state())

@app.route("/api/choose", methods=["POST"])
def api_choose():
    kind = (request.json or {}).get("kind", "file")
    p = native_dialog(kind)
    if not p:
        return jsonify(ok=False, message="Couldn't open a file dialog here. Type or paste the path instead.")
    s = load_settings()
    s["master_path" if kind == "file" else "out_root"] = p
    save_settings(s)
    return jsonify(ok=True, **state())

@app.route("/api/forms", methods=["POST"])
def api_forms_add():
    for f in request.files.getlist("files"):
        name = os.path.basename((f.filename or "form.xlsx").replace("\\", "/"))
        if not name.lower().endswith((".xlsx", ".xlsm")):
            continue
        fid = uuid.uuid4().hex[:8]
        d = os.path.join(SESSION, fid)
        os.makedirs(d)
        path = os.path.join(d, name)
        f.save(path)
        FORMS[fid] = {"name": name, "path": path}
    return jsonify(state())

@app.route("/api/forms/<fid>", methods=["DELETE"])
def api_forms_del(fid):
    f = FORMS.pop(fid, None)
    if f:
        shutil.rmtree(os.path.dirname(f["path"]), ignore_errors=True)
    return jsonify(state())

@app.route("/api/fill", methods=["POST"])
def api_fill():
    s = load_settings()
    master, err = read_master(s["master_path"])
    if not master:
        return jsonify(ok=False, message=err or "Choose the master file first.")
    if not FORMS:
        return jsonify(ok=False, message="Add at least one port form first.")
    try:
        res = fill_all(s["master_path"], [f["path"] for f in FORMS.values()], s["out_root"])
    except PermissionError:
        return jsonify(ok=False, message="Can't save in this folder. Choose another folder.")
    except Exception as e:
        return jsonify(ok=False, message=f"Something went wrong while filling ({e.__class__.__name__}: {e}).")
    return jsonify(ok=True, **res)

@app.route("/api/open", methods=["POST"])
def api_open():
    p = (request.json or {}).get("path", "")
    if os.path.isdir(p):
        open_folder(p)
        return jsonify(ok=True)
    return jsonify(ok=False)

@app.route("/api/quit", methods=["POST"])
def api_quit():
    threading.Timer(0.3, lambda: os._exit(0)).start()
    return jsonify(ok=True)

# ---------------------------------------------------------------- start
def free_port(preferred=8765):
    for p in (preferred, 0):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", p))
                return s.getsockname()[1]
            except OSError:
                continue

def open_window(url):
    """Open like an app window (Edge/Chrome app mode) if possible, else the default browser."""
    try:
        sysname = platform.system()
        if sysname == "Windows":
            for exe in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                        r"C:\Program Files\Google\Chrome\Application\chrome.exe"):
                if os.path.exists(exe):
                    subprocess.Popen([exe, f"--app={url}", "--window-size=760,900"])
                    return
        elif sysname == "Darwin":
            for app_name in ("Google Chrome", "Microsoft Edge"):
                if os.path.exists(f"/Applications/{app_name}.app"):
                    subprocess.Popen(["open", "-na", app_name, "--args", f"--app={url}", "--window-size=760,900"])
                    return
    except Exception:
        pass
    webbrowser.open(url)

def watchdog():
    """Stop the program when the window is closed (no ping for 25 seconds)."""
    while True:
        time.sleep(3)
        t = LAST_PING["t"]
        if t is not None and time.time() - t > 25:
            os._exit(0)

def main():
    # A packaged windowed app has no console: give print/logging somewhere harmless to write
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")
    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    port = free_port()
    url = f"http://127.0.0.1:{port}/"
    if not os.environ.get("PORTFILL_NO_WATCHDOG"):
        threading.Thread(target=watchdog, daemon=True).start()
    if not os.environ.get("PORTFILL_NO_OPEN"):
        threading.Timer(1.0, lambda: open_window(url)).start()
    print(f"PortFill is running at {url}  (close the window to stop)")
    app.run(host="127.0.0.1", port=port, threaded=True, debug=False)

if __name__ == "__main__":
    main()
