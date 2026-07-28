import os
import sys
import json
import time
import queue
import threading
import io
from pathlib import Path
from flask import Flask, render_template, request, jsonify, Response, send_from_directory

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).parent

try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import win32com.client
    HAS_COM = True
except ImportError:
    HAS_COM = False

app = Flask(__name__)
MEDIA_DIR = BASE_DIR / 'media'
MEDIA_DIR.mkdir(exist_ok=True)
PPTX_CACHE = MEDIA_DIR / '.pptx_cache'
PPTX_CACHE.mkdir(exist_ok=True)

sse_clients = []
sse_lock = threading.Lock()

BACKGROUND_FILE = MEDIA_DIR / '.background.json'
current_background = None

def _load_background():
    global current_background
    if BACKGROUND_FILE.exists():
        try:
            with open(BACKGROUND_FILE) as f:
                d = json.load(f)
                current_background = d.get('file')
        except Exception:
            current_background = None
    else:
        current_background = None

def _save_background(filename=None):
    global current_background
    current_background = filename
    if filename:
        BACKGROUND_FILE.write_text(json.dumps({'file': filename}, ensure_ascii=False), encoding='utf-8')
    else:
        BACKGROUND_FILE.unlink(missing_ok=True)

_load_background()

MEDIA_EXTENSIONS = {
    'image': ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.svg', '.ico'),
    'video': ('.mp4', '.webm', '.mov', '.avi', '.mkv', '.flv', '.wmv'),
    'audio': ('.mp3', '.wav', '.ogg', '.aac', '.flac', '.m4a', '.wma'),
    'pdf': ('.pdf',),
    'ppt': ('.pptx', '.ppt'),
}

def get_media_type(ext):
    for mtype, exts in MEDIA_EXTENSIONS.items():
        if ext in exts:
            return mtype
    return 'unknown'

def get_media_files():
    files = []
    if not MEDIA_DIR.exists():
        return files
    for f in sorted(MEDIA_DIR.iterdir()):
        if not f.is_file():
            continue
        ext = f.suffix.lower()
        mtype = get_media_type(ext)
        if mtype == 'unknown':
            continue
        size = f.stat().st_size
        display_type = 'document' if mtype in ('pdf', 'ppt') else mtype
        files.append({
            'name': f.name,
            'path': f.name,
            'type': display_type,
            'subtype': mtype,
            'ext': ext,
            'size': size,
        })
    return files

def broadcast(cmd):
    with sse_lock:
        dead = []
        for q in sse_clients:
            try:
                q.put_nowait(cmd)
            except queue.Full:
                dead.append(q)
        for q in dead:
            sse_clients.remove(q)

# ---- PPT rendering (COM) ----

def _render_ppt_com(filename):
    import shutil
    stem = Path(filename).stem
    cache_dir = PPTX_CACHE / stem
    cache_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = cache_dir / 'manifest.json'

    if manifest_path.exists():
        with open(manifest_path, encoding='utf-8') as f:
            m = json.load(f)
        if m.get('renderer') == 'com':
            return m

    if cache_dir.exists():
        shutil.rmtree(str(cache_dir))
    cache_dir.mkdir(parents=True)

    ppt_path = str((MEDIA_DIR / filename).resolve())
    powerpoint = win32com.client.Dispatch('Powerpoint.Application')
    pres = None
    try:
        pres = powerpoint.Presentations.Open(ppt_path, ReadOnly=True, WithWindow=True)
        slide_count = pres.Slides.Count
        slides_info = []

        for i in range(1, slide_count + 1):
            slide = pres.Slides(i)
            slide_path = str((cache_dir / f'slide_{i-1}.png').resolve())
            slide.Export(slide_path, 'PNG', 1920, 1080)
            slides_info.append({
                'index': i - 1,
                'image': f'/pptx_cache/{stem}/slide_{i-1}.png',
            })

        manifest = {'count': slide_count, 'file': filename, 'renderer': 'com', 'slides': slides_info}
        with open(manifest_path, 'w', encoding='utf-8') as f:
            json.dump(manifest, f)
        return manifest
    finally:
        if pres:
            try:
                pres.Close()
            except Exception:
                pass
        try:
            powerpoint.Quit()
        except Exception:
            pass


# ---- PDF rendering ----

def _render_pdf(filename):
    import shutil
    stem = Path(filename).stem
    cache_dir = PPTX_CACHE / stem
    cache_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = cache_dir / 'manifest.json'

    if manifest_path.exists():
        with open(manifest_path, encoding='utf-8') as f:
            m = json.load(f)
        if m.get('renderer') == 'pymupdf':
            return m

    if cache_dir.exists():
        shutil.rmtree(str(cache_dir))
    cache_dir.mkdir(parents=True)

    pdf_path = str((MEDIA_DIR / filename).resolve())
    pdf_doc = fitz.open(pdf_path)
    slide_count = pdf_doc.page_count
    slides_info = []

    for i in range(slide_count):
        page = pdf_doc[i]
        pix = page.get_pixmap(dpi=192)
        slide_path = cache_dir / f'slide_{i}.png'
        pix.save(str(slide_path))
        slides_info.append({
            'index': i,
            'image': f'/pptx_cache/{stem}/slide_{i}.png',
        })
    pdf_doc.close()

    manifest = {'count': slide_count, 'file': filename, 'renderer': 'pymupdf', 'slides': slides_info}
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f)
    return manifest


def _get_cached_manifest(filename):
    stem = Path(filename).stem
    cache_dir = PPTX_CACHE / stem
    manifest_path = cache_dir / 'manifest.json'
    if manifest_path.exists():
        with open(manifest_path, encoding='utf-8') as f:
            m = json.load(f)
        if m.get('slides'):
            return m
    return None


# ---- Routes ----

@app.route('/')
def control():
    return render_template('control.html')

@app.route('/output')
def output():
    return render_template('output.html')

@app.route('/api/media')
def api_media():
    return jsonify(get_media_files())

@app.route('/api/play', methods=['POST'])
def api_play():
    data = request.get_json()
    filename = data.get('path') or data.get('file') or data.get('name')
    subtype = data.get('subtype', data.get('type', 'unknown'))

    cmd = {
        'action': 'play',
        'file': filename,
        'type': data.get('type', 'unknown'),
        'subtype': subtype,
        'path': '/media/' + filename,
    }

    result = {'ok': True}

    if subtype in ('pdf', 'ppt'):
        cached = _get_cached_manifest(filename)
        if not cached and subtype == 'ppt' and HAS_COM:
            try:
                _render_ppt_com(filename)
                cached = _get_cached_manifest(filename)
            except Exception as e:
                return jsonify({'error': f'PPT转换失败: {e}'}), 400
        if cached:
            cmd['slideCount'] = cached['count']
            cmd['slideIndex'] = 0
            cmd['slideImage'] = cached['slides'][0]['image']
            result['slideCount'] = cached['count']
            result['slides'] = cached['slides']
            broadcast(cmd)
            return jsonify(result)
        return jsonify({'error': '请重新上传'}), 400

    broadcast(cmd)
    return jsonify(result)

@app.route('/api/pptx_nav', methods=['POST'])
def api_pptx_nav():
    data = request.get_json()
    broadcast({
        'action': 'pptx_nav',
        'slideIndex': data.get('slide'),
    })
    return jsonify({'ok': True})

@app.route('/api/background', methods=['GET', 'POST', 'DELETE'])
def api_background():
    if request.method == 'GET':
        return jsonify({'file': current_background})
    if request.method == 'DELETE':
        _save_background(None)
        broadcast({'action': 'background_clear'})
        return jsonify({'ok': True})
    data = request.get_json()
    filename = (data or {}).get('file')
    if not filename:
        return jsonify({'error': 'no file'}), 400
    filepath = MEDIA_DIR / filename
    if not filepath.exists():
        return jsonify({'error': 'file not found'}), 404
    _save_background(filename)
    broadcast({'action': 'background_set', 'file': filename})
    return jsonify({'ok': True})

@app.route('/api/background_toggle', methods=['POST'])
def api_background_toggle():
    broadcast({'action': 'background_toggle'})
    return jsonify({'ok': True})

@app.route('/api/stop', methods=['POST'])
def api_stop():
    broadcast({'action': 'stop'})
    return jsonify({'ok': True})

@app.route('/api/pause', methods=['POST'])
def api_pause():
    broadcast({'action': 'pause'})
    return jsonify({'ok': True})

@app.route('/api/progress', methods=['POST'])
def api_progress():
    data = request.get_json()
    broadcast({
        'action': 'progress',
        'currentTime': data.get('currentTime', 0),
        'duration': data.get('duration', 0),
        'paused': data.get('paused', False),
        'file': data.get('file', ''),
    })
    return jsonify({'ok': True})

@app.route('/api/seek', methods=['POST'])
def api_seek():
    data = request.get_json()
    broadcast({
        'action': 'seek',
        'time': data.get('time', 0),
    })
    return jsonify({'ok': True})

@app.route('/api/upload', methods=['POST'])
def api_upload():
    if 'file' not in request.files:
        return jsonify({'error': 'no file'}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({'error': 'empty filename'}), 400
    f.save(MEDIA_DIR / f.filename)
    ext = Path(f.filename).suffix.lower()
    subtype = None
    for mt, exts in MEDIA_EXTENSIONS.items():
        if ext in exts:
            subtype = mt
            break
    if subtype == 'pdf' and HAS_PYMUPDF:
        _render_pdf(f.filename)
    elif subtype == 'ppt':
        if HAS_COM:
            try:
                _render_ppt_com(f.filename)
            except Exception as e:
                return jsonify({'ok': True, 'name': f.filename, 'convert_warning': f'PPT转换失败，点击播放时将重试: {e}'}), 200
    return jsonify({'ok': True, 'name': f.filename})

@app.route('/api/delete', methods=['POST'])
def api_delete():
    data = request.get_json()
    filepath = MEDIA_DIR / data['file']
    if filepath.exists():
        filepath.unlink()
    return jsonify({'ok': True})

@app.route('/media/<path:filename>')
def serve_media(filename):
    ext = Path(filename).suffix.lower()
    resp = send_from_directory(MEDIA_DIR, filename)
    if ext == '.pdf':
        resp.headers['Content-Disposition'] = 'inline'
        resp.headers['X-Content-Type-Options'] = 'nosniff'
    return resp

@app.route('/pptx_cache/<path:filename>')
def serve_pptx_cache(filename):
    return send_from_directory(PPTX_CACHE, filename)

@app.route('/events')
def events():
    q = queue.Queue(maxsize=100)
    with sse_lock:
        sse_clients.append(q)

    def generate():
        try:
            if current_background:
                yield 'data: ' + json.dumps({'action': 'background_set', 'file': current_background}) + '\n\n'
            while True:
                try:
                    cmd = q.get(timeout=30)
                    yield 'data: ' + json.dumps(cmd) + '\n\n'
                except queue.Empty:
                    yield ': heartbeat\n\n'
        except GeneratorExit:
            with sse_lock:
                if q in sse_clients:
                    sse_clients.remove(q)

    return Response(
        generate(),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache, no-transform',
            'Connection': 'keep-alive',
        }
    )

if __name__ == '__main__':
    import webbrowser
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        local_ip = s.getsockname()[0]
    except Exception:
        local_ip = '127.0.0.1'
    finally:
        s.close()
    print('=' * 60)
    print('  Media Controller - 媒体控制器')
    print('=' * 60)
    print(f'  媒体目录: {MEDIA_DIR}')
    print(f'  PDF渲染: {"PyMuPDF" if HAS_PYMUPDF else "不可用"}')
    print(f'  PPT渲染: {"COM (需安装Office)" if HAS_COM else "不可用"}')
    print()
    print(f'  控制面板 [主屏]:  http://localhost:5000')
    print(f'  输出屏幕 [扩展]:  http://localhost:5000/output')
    print(f'  局域网控制面板:  http://{local_ip}:5000')
    print(f'  局域网扩展屏:    http://{local_ip}:5000/output')
    print()
    print('  [使用说明]')
    print('  1. 在主屏浏览器打开 [控制面板]')
    print('  2. 在扩展屏浏览器打开 [输出屏幕] 并按 F11 全屏')
    print('  3. 点击任意媒体文件 -> 扩展屏播放')
    print('  4. PDF文件: 控制面板出现翻页控制器')
    print('=' * 60)
    threading.Timer(1.5, lambda: webbrowser.open('http://localhost:5000')).start()
    threading.Timer(1.5, lambda: webbrowser.open('http://localhost:5000/output')).start()
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
