#!/usr/bin/env python3
"""
setup_github.py — одноразовый скрипт настройки.

Запусти один раз на своей машине:
    python setup_github.py

Что делает:
  1. Создаёт приватный репозиторий ctv-document-suite
  2. Загружает все файлы из папки ctv_suite/
  3. Создаёт релиз v2.0.0
  4. Прикрепляет CTV_Document_Suite.exe к релизу
  5. Обновляет GITHUB_REPO в updater.py

После этого для каждого нового обновления используй release.py
"""

import json
import os
import sys
import urllib.request
import urllib.error
from pathlib import Path

TOKEN = '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'
REPO_NAME = 'ctv-document-suite'


def api(method: str, path: str, data=None, raw=False):
    url = f'https://api.github.com{path}'
    body = json.dumps(data).encode() if data else None
    req  = urllib.request.Request(url, data=body, method=method)
    req.add_header('Authorization', f'token {TOKEN}')
    req.add_header('Accept', 'application/vnd.github.v3+json')
    if body:
        req.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(req) as r:
            return r.read() if raw else json.loads(r.read())
    except urllib.error.HTTPError as e:
        err = json.loads(e.read())
        raise RuntimeError(f'GitHub API {e.code}: {err.get("message")}') from e


def upload_asset(upload_url: str, filepath: Path):
    """Загружает файл как asset к релизу."""
    # upload_url вида: https://uploads.github.com/repos/.../assets{?name,label}
    clean_url = upload_url.split('{')[0]
    url = f'{clean_url}?name={filepath.name}'
    data = filepath.read_bytes()
    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization', f'token {TOKEN}')
    req.add_header('Content-Type', 'application/octet-stream')
    req.add_header('Accept', 'application/vnd.github.v3+json')
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def main():
    script_dir = Path(__file__).parent
    suite_dir  = script_dir

    if not suite_dir.exists():
        print(f'❌  Папка {suite_dir} не найдена.')
        sys.exit(1)

    # ── 1. Получаем username ──────────────────────────────────────────────────
    print('🔑  Проверяю токен…')
    user = api('GET', '/user')
    username = user['login']
    print(f'✅  Вошёл как: {username}')

    # ── 2. Создаём репозиторий ────────────────────────────────────────────────
    print(f'\n📁  Создаю репозиторий {username}/{REPO_NAME}…')
    try:
        repo = api('POST', '/user/repos', {
            'name':        REPO_NAME,
            'description': 'CTV Document Suite — генератор КП и Каталогов',
            'private':     True,
            'auto_init':   True,
        })
        print(f'✅  Репозиторий создан: {repo["html_url"]}')
    except RuntimeError as e:
        if 'already exists' in str(e):
            print('ℹ️   Репозиторий уже существует, продолжаю…')
            repo = api('GET', f'/repos/{username}/{REPO_NAME}')
        else:
            raise

    # ── 3. Обновляем GITHUB_REPO в updater.py ────────────────────────────────
    updater_path = suite_dir / 'updater.py'
    if updater_path.exists():
        text = updater_path.read_text('utf-8')
        text = text.replace(
            "GITHUB_REPO    = 'laskinss27-cmyk/ctv-document-suite'",
            f"GITHUB_REPO    = '{username}/{REPO_NAME}'"
        )
        updater_path.write_text(text, encoding='utf-8')
        print(f'✅  updater.py: GITHUB_REPO = {username}/{REPO_NAME}')

    # ── 4. Загружаем файлы через Git API ─────────────────────────────────────
    print('\n📤  Загружаю файлы…')

    # Получаем SHA последнего коммита
    ref    = api('GET', f'/repos/{username}/{REPO_NAME}/git/ref/heads/main')
    sha    = ref['object']['sha']
    tree_r = api('GET', f'/repos/{username}/{REPO_NAME}/git/trees/{sha}')
    base_sha = tree_r['sha']

    # Собираем все файлы
    import base64
    blobs = []
    for f in suite_dir.rglob('*'):
        if f.is_file() and '__pycache__' not in str(f) and f.suffix != '.pyc':
            rel = f.relative_to(script_dir).as_posix()
            try:
                content = base64.b64encode(f.read_bytes()).decode()
                blob = api('POST', f'/repos/{username}/{REPO_NAME}/git/blobs', {
                    'content':  content,
                    'encoding': 'base64',
                })
                blobs.append({'path': rel, 'mode': '100644',
                               'type': 'blob', 'sha': blob['sha']})
                print(f'  + {rel}')
            except Exception as ex:
                print(f'  ⚠ {rel}: {ex}')

    # Создаём дерево
    tree = api('POST', f'/repos/{username}/{REPO_NAME}/git/trees', {
        'base_tree': base_sha,
        'tree':      blobs,
    })

    # Коммит
    commit = api('POST', f'/repos/{username}/{REPO_NAME}/git/commits', {
        'message': 'Initial commit: CTV Document Suite v2.0.0',
        'tree':    tree['sha'],
        'parents': [sha],
    })

    # Обновляем ветку
    api('PATCH', f'/repos/{username}/{REPO_NAME}/git/refs/heads/main', {
        'sha': commit['sha'],
    })
    print('✅  Код загружен')

    # ── 5. Создаём релиз v2.0.0 ──────────────────────────────────────────────
    print('\n🏷️   Создаю релиз v2.0.0…')
    try:
        release = api('POST', f'/repos/{username}/{REPO_NAME}/releases', {
            'tag_name':         'v2.0.0',
            'target_commitish': 'main',
            'name':             'v2.0.0 — Новый редактор',
            'body':             '• Новый редактор товаров и работ\n'
                                '• Исправлен скролл и ресайз\n'
                                '• Система обновлений через GitHub',
            'draft':            False,
            'prerelease':       False,
        })
        print(f'✅  Релиз создан: {release["html_url"]}')
    except RuntimeError as e:
        if 'already_exists' in str(e):
            print('ℹ️   Релиз уже существует')
            releases = api('GET', f'/repos/{username}/{REPO_NAME}/releases')
            release  = next(r for r in releases if r['tag_name'] == 'v2.0.0')
        else:
            raise

    # ── 6. Прикрепляем exe ───────────────────────────────────────────────────
    exe_candidates = [
        script_dir / 'dist' / 'CTV_Document_Suite.exe',
        script_dir / 'CTV_Document_Suite.exe',
    ]
    exe_path = next((p for p in exe_candidates if p.exists()), None)

    if exe_path:
        print(f'\n📎  Прикрепляю {exe_path.name} к релизу…')
        asset = upload_asset(release['upload_url'], exe_path)
        print(f'✅  Asset загружен: {asset["browser_download_url"]}')
    else:
        print('\n⚠️   exe не найден. После сборки запусти:')
        print(f'    python release.py --version 2.0.0 --exe dist/CTV_Document_Suite.exe')

    # ── Итог ─────────────────────────────────────────────────────────────────
    print(f"""
╔══════════════════════════════════════════════════════════╗
  ✅  Всё готово!

  Репозиторий : {repo['html_url']}
  Релизы      : {repo['html_url']}/releases

  Что дальше:
  1. Собери exe: pyinstaller CTV_Suite.spec
  2. Залей в релиз: python release.py --version 2.0.0
                      --exe dist/CTV_Document_Suite.exe
  3. Раздай exe пользователям — дальше обновления автоматически
╚══════════════════════════════════════════════════════════╝
""")


if __name__ == '__main__':
    main()
