import os, shutil, subprocess, sys
from pathlib import Path

BASE = Path(__file__).parent
DIST = BASE / 'dist' / 'MediaController'


def build():
    if DIST.exists():
        shutil.rmtree(DIST)
    if (BASE / 'build').exists():
        shutil.rmtree(BASE / 'build')

    print('Building with PyInstaller...')
    subprocess.run([
        sys.executable, '-m', 'PyInstaller',
        '--name', 'MediaController',
        '--distpath', str(BASE / 'dist'),
        '--onedir',
        '--add-data', f'templates{os.pathsep}templates',
        '--noconfirm',
        'app.py',
    ], check=True)

    media_dir = DIST / 'media'
    media_dir.mkdir(exist_ok=True)
    readme = media_dir / '将媒体文件放入此文件夹.txt'
    readme.write_text('将图片、视频、音频、PDF 文件放入此文件夹', encoding='utf-8')

    exe_name = 'MediaController.exe' if sys.platform == 'win32' else 'MediaController'
    print(f'\nBuild complete: {DIST}')
    print(f'  Run: {DIST / exe_name}')

    total_size = sum(os.path.getsize(os.path.join(r, f)) for r,_,fs in os.walk(DIST) for f in fs)
    print(f'  Total size: {total_size/1024/1024:.0f}M')


if __name__ == '__main__':
    build()
