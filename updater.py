"""
updater.py — автообновление через GitHub Releases.

Схема:
  1. GET https://api.github.com/repos/OWNER/REPO/releases/latest
  2. Сравниваем tag_name (v2.0.0) с локальной версией
  3. Если новее — скачиваем .exe из assets по прямой ссылке
  4. Bat-скрипт заменяет exe и перезапускает программу

Преимущество перед Google Drive: прямые ссылки без редиректов и
страниц подтверждения. Работает с первого запроса.
"""

import os
import re
import sys
import threading
import tempfile
import urllib.request
from pathlib import Path

# ── Настройки ─────────────────────────────────────────────────────────────────
GITHUB_REPO    = 'ТВОЙ_ЛОГИН/ctv-document-suite'   # ← заменить после создания репо
APP_VERSION    = '2.0.0'
CHECK_INTERVAL = 20 * 60   # секунд (20 минут)
# ─────────────────────────────────────────────────────────────────────────────

RELEASES_URL = f'https://api.github.com/repos/{GITHUB_REPO}/releases/latest'

if getattr(sys, 'frozen', False):
    CURRENT_EXE = Path(sys.executable)
else:
    CURRENT_EXE = Path(sys.argv[0]).resolve()

_LOCAL_VER_FILE = CURRENT_EXE.parent / 'version_local.txt'


def _ver(v: str) -> tuple:
    try:
        return tuple(int(x) for x in re.sub(r'[^\d.]', '', v).split('.') if x)
    except Exception:
        return (0,)


def get_local_version() -> str:
    try:
        if _LOCAL_VER_FILE.exists():
            for line in _LOCAL_VER_FILE.read_text('utf-8').splitlines():
                if line.startswith('version='):
                    return line.split('=', 1)[1].strip()
    except Exception:
        pass
    return APP_VERSION


def _save_local_version(v: str):
    try:
        _LOCAL_VER_FILE.write_text(f'version={v}\n', encoding='utf-8')
    except Exception:
        pass


def check_for_update():
    """
    Запрашивает GitHub API и возвращает dict с данными релиза если
    версия новее локальной. Иначе None.

    Структура возвращаемого dict:
      version     — строка вида '2.1.0'
      download_url — прямая ссылка на .exe из assets
      changelog   — текст тела релиза (что нового)
    """
    try:
        req = urllib.request.Request(
            RELEASES_URL,
            headers={
                'User-Agent':  f'CTV-Suite/{APP_VERSION}',
                'Accept':      'application/vnd.github.v3+json',
            }
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            import json
            data = json.loads(resp.read().decode('utf-8'))

        tag      = data.get('tag_name', '0.0.0')          # 'v2.1.0'
        version  = tag.lstrip('v')                         # '2.1.0'
        body     = data.get('body', '')                    # changelog
        assets   = data.get('assets', [])

        # Ищем .exe в assets
        exe_url = None
        for asset in assets:
            if asset.get('name', '').endswith('.exe'):
                exe_url = asset.get('browser_download_url')
                break

        if not exe_url:
            return None

        if _ver(version) > _ver(get_local_version()):
            return {
                'version':      version,
                'download_url': exe_url,
                'changelog':    body,
            }

    except Exception:
        pass
    return None


def download_and_apply(download_url: str, new_version: str,
                       progress_cb=None) -> bool:
    """
    Скачивает новый exe напрямую с GitHub и запускает bat-скрипт замены.
    GitHub отдаёт файл без редиректов и подтверждений.
    """
    new_exe = CURRENT_EXE.parent / (CURRENT_EXE.stem + '_update.exe')
    new_exe.unlink(missing_ok=True)

    try:
        if progress_cb:
            progress_cb(2, 'Подключение к GitHub…')

        req = urllib.request.Request(
            download_url,
            headers={
                'User-Agent': f'CTV-Suite/{APP_VERSION}',
                'Accept':     'application/octet-stream',
            }
        )
        with urllib.request.urlopen(req, timeout=180) as resp:
            total      = int(resp.headers.get('Content-Length') or 0)
            downloaded = 0
            CHUNK      = 65_536

            with open(new_exe, 'wb') as f:
                while True:
                    chunk = resp.read(CHUNK)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_cb and total > 0:
                        pct = max(3, min(95, int(downloaded * 93 / total)))
                        mb  = downloaded / 1_048_576
                        progress_cb(pct, f'Скачивание… {mb:.1f} МБ')

        if new_exe.stat().st_size < 100_000:
            new_exe.unlink(missing_ok=True)
            if progress_cb:
                progress_cb(-1, 'Ошибка: файл повреждён')
            return False

        if progress_cb:
            progress_cb(97, 'Установка…')

        # Bat-скрипт: ждёт завершения текущего процесса → заменяет exe → запускает
        bat = Path(tempfile.mktemp(suffix='.bat'))
        bat.write_text(
            f'@echo off\n'
            f'chcp 65001 >nul\n'
            f':wait\n'
            f'tasklist /FI "PID eq {os.getpid()}" 2>nul | '
            f'find /I "{os.getpid()}" >nul\n'
            f'if not errorlevel 1 (\n'
            f'    timeout /T 1 /NOBREAK >nul\n'
            f'    goto wait\n'
            f')\n'
            f'move /Y "{new_exe}" "{CURRENT_EXE}"\n'
            f'start "" "{CURRENT_EXE}"\n'
            f'del "%~f0"\n',
            encoding='utf-8'
        )

        import subprocess
        subprocess.Popen(
            ['cmd.exe', '/C', str(bat)],
            creationflags=subprocess.CREATE_NO_WINDOW,
            close_fds=True,
        )

        _save_local_version(new_version)

        if progress_cb:
            progress_cb(100, 'Готово! Перезапуск…')

        return True

    except Exception as e:
        new_exe.unlink(missing_ok=True)
        if progress_cb:
            progress_cb(-1, f'Ошибка: {e}')
        return False


def quit_for_update():
    sys.exit(0)


class AutoUpdater:
    def __init__(self, tk_root, on_update_found):
        self._root     = tk_root
        self._callback = on_update_found
        self._stop     = threading.Event()
        self._thread   = threading.Thread(target=self._loop, daemon=True)

    def start(self):
        self._thread.start()

    def stop(self):
        self._stop.set()

    def _loop(self):
        self._stop.wait(30)
        while not self._stop.is_set():
            try:
                info = check_for_update()
                if info:
                    self._root.after(0, self._callback, info)
                    return
            except Exception:
                pass
            self._stop.wait(CHECK_INTERVAL)
