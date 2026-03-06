#!/usr/bin/env python3
"""
release.py — публикация нового обновления на GitHub.

Использование:
    python release.py --version 2.1.0 --exe dist/CTV_Document_Suite.exe

Что делает:
  1. Создаёт тег и релиз v2.1.0 на GitHub
  2. Загружает exe как asset
  3. Пользователи получат уведомление в течение 20 минут

Примечание: updater.py в программе сравнивает версии через GitHub API,
скачивает exe напрямую из assets — без редиректов и подтверждений.
"""

import argparse
import json
import sys
import urllib.request
import urllib.error
from pathlib import Path

TOKEN     = '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'
REPO_NAME = 'ctv-document-suite'


def api(method, path, data=None):
    url  = f'https://api.github.com{path}'
    body = json.dumps(data).encode() if data else None
    req  = urllib.request.Request(url, data=body, method=method)
    req.add_header('Authorization',  f'token {TOKEN}')
    req.add_header('Accept', 'application/vnd.github.v3+json')
    if body:
        req.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        err = json.loads(e.read())
        raise RuntimeError(f'GitHub {e.code}: {err.get("message")}') from e


def upload_asset(upload_url, filepath: Path):
    url  = upload_url.split('{')[0] + f'?name={filepath.name}'
    data = filepath.read_bytes()
    req  = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization',  f'token {TOKEN}')
    req.add_header('Content-Type',   'application/octet-stream')
    req.add_header('Accept', 'application/vnd.github.v3+json')
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--version', required=True,
                    help='Номер версии, напр. 2.1.0')
    ap.add_argument('--exe',     required=True,
                    help='Путь к .exe файлу')
    ap.add_argument('--notes',   default='',
                    help='Что нового (необязательно)')
    args = ap.parse_args()

    exe_path = Path(args.exe)
    if not exe_path.exists():
        print(f'❌  Файл не найден: {exe_path}')
        sys.exit(1)

    version = args.version.lstrip('v')
    tag     = f'v{version}'

    # Получаем username
    user     = api('GET', '/user')
    username = user['login']
    repo_path = f'/repos/{username}/{REPO_NAME}'

    print(f'📦  Публикую {tag} для {username}/{REPO_NAME}…')

    # Создаём релиз
    body = args.notes or f'Обновление {tag}'
    try:
        release = api('POST', f'{repo_path}/releases', {
            'tag_name':         tag,
            'target_commitish': 'main',
            'name':             tag,
            'body':             body,
            'draft':            False,
            'prerelease':       False,
        })
    except RuntimeError as e:
        if 'already_exists' in str(e):
            print(f'⚠️   Релиз {tag} уже существует, добавляю asset…')
            releases = api('GET', f'{repo_path}/releases')
            release  = next((r for r in releases
                              if r['tag_name'] == tag), None)
            if not release:
                raise
        else:
            raise

    # Загружаем exe
    print(f'⬆️   Загружаю {exe_path.name} ({exe_path.stat().st_size//1024} КБ)…')
    asset = upload_asset(release['upload_url'], exe_path)

    print(f"""
✅  Готово!

   Релиз  : {release['html_url']}
   Файл   : {asset['browser_download_url']}
   Размер : {asset['size'] // 1024} КБ

Пользователи получат уведомление в течение 20 минут.
""")


if __name__ == '__main__':
    main()
