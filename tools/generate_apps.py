#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""data/chapterNN.json から章別学習アプリとダッシュボード (index.html) を生成する。

使い方:
    python3 tools/generate_apps.py

新しい章を追加する手順:
    1. docx を source_docs/ に置く
    2. python3 tools/parse_vintage_docx.py source_docs/chapterNN_章名.docx
    3. python3 tools/generate_apps.py
"""

import glob
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES = os.path.join(ROOT, 'tools', 'templates')
DATA_DIR = os.path.join(ROOT, 'data')


def main():
    with open(os.path.join(TEMPLATES, 'chapter_template.html'), encoding='utf-8') as f:
        chapter_tpl = f.read()
    with open(os.path.join(TEMPLATES, 'index_template.html'), encoding='utf-8') as f:
        index_tpl = f.read()

    chapters_info = []
    for path in sorted(glob.glob(os.path.join(DATA_DIR, 'chapter*.json'))):
        with open(path, encoding='utf-8') as f:
            chapter = json.load(f)

        file_name = f'chapter{chapter["number"]:02d}.html'
        app_id = f'vintage_chapter{chapter["number"]:02d}'

        html = (chapter_tpl
                .replace('__FULL_TITLE__', chapter['fullTitle'])
                .replace('__APP_ID__', app_id)
                .replace('__DATA__', json.dumps(chapter['problems'], ensure_ascii=False, indent=1)))

        out = os.path.join(ROOT, file_name)
        with open(out, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f'{file_name}: {chapter["fullTitle"]} ({len(chapter["problems"])}問)')

        chapters_info.append({
            'number': chapter['number'],
            'title': chapter['title'],
            'file': file_name,
            'problemCount': len(chapter['problems']),
        })

    index_html = index_tpl.replace(
        '__CHAPTERS_INFO__', json.dumps(chapters_info, ensure_ascii=False))
    with open(os.path.join(ROOT, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(index_html)
    print(f'index.html: {len(chapters_info)}章を掲載')


if __name__ == '__main__':
    main()
