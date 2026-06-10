#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Vintage 4th Edition の docx (章ごとの問題ファイル) を構造化 JSON に変換する。

使い方:
    python3 tools/parse_vintage_docx.py source_docs/chapter01_時制.docx ...

docx の形式:
    第１章　時制            ← 章タイトル
    ■001                   ← 問題番号
    基本 / 発展             ← 任意タグ
    ●超頻出                ← 任意タグ
    (問題文 1〜複数行)
    ①xx　②yy　③zz　④ww   ← 4択の場合
    #③                     ← 解答 (行頭 #)
    (和訳 1〜複数行)

問題タイプ:
    choice   : 4択空所補充 (①〜④の選択肢行あり)
    error    : 誤り指摘 (問題文に ①<u>...</u> を含む)
    ordering : 整序英作文 (英文中に ( a / b / c ) 形式)
    fill     : 空所補充記述 ((　　) のみで選択肢なし)
    rewrite  : 書き換えなどの自由記述
"""

import json
import os
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

CIRCLED = ['①', '②', '③', '④', '⑤', '⑥']
BLANK_RE = re.compile(r'(?:\(\s*[　\s]*\)|\(　+\)|（\s*[　\s]*）|（　+）)')
BLANK_RUN_RE = re.compile(r'(?:[（(][　\s]*[)）][ 　]*){1,}')
# 整序英作文の語群: 全角・半角どちらのカッコにも対応 (■579, ■656)
ORDERING_RE = re.compile(r'[（(]([^()（）]*\s/\s[^()（）]*)[)）]')
# 頭文字ヒント付き空所 '( l　)' (■1082 など)
HINT_BLANK_RE = re.compile(r'\(\s*[a-z][　]+\)')

# 元データの誤植の修正 (問題番号 -> 正しい解答)
ANSWER_OVERRIDES = {
    # ■556: 語群は 'an honest person' だが解答が 'an honest woman' になっている
    '556': 'What do you think made an honest person like Jane',
}


def docx_lines(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml')
    body = ET.fromstring(xml).find(W + 'body')
    lines = []
    for p in body.iter(W + 'p'):
        text = ''.join(t.text or '' for t in p.iter(W + 't'))
        lines.append(text.rstrip())
    return lines


def is_japanese(line):
    return any('぀' <= ch <= 'ヿ' or '一' <= ch <= '鿿'
               or ch in '。，、．「」' for ch in line)


def ja_char_count(line):
    return sum(1 for ch in line if '぀' <= ch <= 'ヿ' or '一' <= ch <= '鿿'
               or ch in '。，、．「」')


def split_choices(line):
    """'①is　②has　③were　④has been' -> ['is','has','were','has been']"""
    parts = re.split(r'[①②③④]', line)
    return [p.strip(' 　') for p in parts[1:] if p.strip(' 　')]


def normalize_answer_index(ans):
    """'③' -> 2"""
    for i, c in enumerate(CIRCLED):
        if c in ans:
            return i
    return None


def strip_fuyou(answer):
    """'have been to Nara（gone不要）' -> 'have been to Nara' (不足の注記も除去)"""
    return re.sub(r'[（(][^（()）]*(?:不要|不足)[)）]', '', answer).strip()


def fill_blanks(text, answer):
    """連続する空所 (　　) (　　) ... を answer で置換して正解文を作る。"""
    answer = answer.split('/')[0].strip()  # 'have/need' -> 'have'
    result = BLANK_RUN_RE.sub(answer + ' ', text, count=1)
    # 残りの空所も同じ語で埋める(通常は1か所のみ)
    result = BLANK_RUN_RE.sub(answer + ' ', result)
    return re.sub(r'\s+([,.!?;:])', r'\1', re.sub(r'  +', ' ', result)).strip()


def fill_answer_blanks(text, answer):
    """fill 問題の正解文を作る。会話表現 (■1330 'time, see' など) では
    カンマ区切りの解答が複数の空所に順番に対応する。"""
    ans = re.sub(r'\[[^\]]*\]', '', answer.split('/')[0]).strip()  # 別解 [..] を除去
    parts = [p.strip() for p in re.split(r'[，,]', ans) if p.strip()]
    runs = BLANK_RUN_RE.findall(text)
    if len(parts) > 1 and len(parts) == len(runs):
        it = iter(parts)
        result = BLANK_RUN_RE.sub(lambda _: next(it) + ' ', text)
    else:
        result = BLANK_RUN_RE.sub((ans if len(parts) <= 1 else parts[0]) + ' ', text)
    # '... = 言い換え' の別表現や話者ラベルは音声用の文からは外す
    result = re.split(r'\s*=\s*', result)[0]
    result = re.sub(r'^[AB]\s*[:：]\s*', '', result.strip())
    return re.sub(r'\s+([,.!?;:])', r'\1', re.sub(r'  +', ' ', result)).strip()


def strip_ab_label(line):
    return re.sub(r'^\([ab]\)\s*', '', line).strip()


def build_error_sentence(question, answer):
    """誤り指摘問題の正解文を作る。answer 例: '②→would'"""
    m = re.match(r'([①②③④])\s*→\s*(.+)', answer)
    correction = None
    wrong_idx = None
    if m:
        wrong_idx = CIRCLED.index(m.group(1))
        correction = m.group(2).strip()
    sentence = question
    segs = re.findall(r'([①②③④])<u>(.*?)</u>', sentence)
    for marker, seg in segs:
        idx = CIRCLED.index(marker)
        rep = correction if (wrong_idx is not None and idx == wrong_idx) else seg
        sentence = sentence.replace(f'{marker}<u>{seg}</u>', rep, 1)
    return re.sub(r'\s+', ' ', sentence).strip(), wrong_idx, correction


def parse_problem(num, lines):
    tags = []
    body = []
    answer = None
    post = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if answer is None and line in ('基本', '発展'):
            tags.append(line)
            continue
        if answer is None and line.startswith('●'):
            tags.append(line.lstrip('●'))
            continue
        if line.startswith('#') and answer is None:
            answer = line[1:].strip()
            continue
        if answer is None:
            # 「■次の(a)，(b)の文が...」のような指示行の ■ は除去
            body.append(line.lstrip('■').strip() if line.startswith('■') else line)
        else:
            post.append(line)

    if answer is None and body:
        # '#' が抜けた '(a) any　(b) No' 形式の解答行を救済する
        last = body[-1]
        if re.match(r'^\(a\)', last) and '(b)' in last and not BLANK_RE.search(last):
            answer = last
            body = body[:-1]

    if answer is None:
        raise ValueError(f'問題 {num}: 解答行 (#...) が見つかりません')

    answer = ANSWER_OVERRIDES.get(num, answer)

    # 選択肢行 (①で始まる行)。意味一致問題などでは選択肢が複数行に分かれる。
    # 下線部の用法選択問題 (■1559) では下線 <u> を含む英文の選択肢が4行並ぶ
    marked = [l for l in body if re.match(r'^[①②③④]', l) and '<u>' in l]
    underlined_as_choices = len(marked) >= 2
    choice_lines = []
    q_lines = []
    for line in body:
        if re.match(r'^[①②③④]', line) and (
                '<u>' not in line or underlined_as_choices):
            choice_lines.append(line)
        else:
            q_lines.append(line)

    question = '\n'.join(q_lines)
    # 日本語文字が3つ以上あれば日本語行とみなす ('DVD' など英字混じりにも対応)
    ja_q = [l for l in q_lines if ja_char_count(l) >= 3 and '<u>' not in l]
    en_q = [l for l in q_lines if l not in ja_q]
    translation = '\n'.join(post)

    prob = {
        'id': num,
        'tags': tags,
        'question': question,
        'rawAnswer': answer,
        'ja': translation or '\n'.join(ja_q),
    }

    has_blank = bool(BLANK_RE.search(question) or HINT_BLANK_RE.search(question))
    # 誤り指摘は ①<u>...</u> 形式。番号なしの <u> は下線部言い換えなどの4択
    has_numbered_underline = bool(re.search(r'[①②③④]<u>', question))
    has_ordering = bool(ORDERING_RE.search(question))

    if has_numbered_underline:
        prob['type'] = 'error'
        en_src = '\n'.join(en_q) if en_q else question
        sentence, wrong_idx, correction = build_error_sentence(en_src, answer)
        prob['en'] = sentence
        prob['answerIndex'] = wrong_idx
        prob['correction'] = correction
    elif choice_lines:
        prob['type'] = 'choice'
        prob['choices'] = split_choices(''.join(choice_lines))
        idx = normalize_answer_index(answer)
        prob['answerIndex'] = idx
        # 正解文: (b) 行があればそれを優先
        en_lines = [l for l in en_q if BLANK_RE.search(l)]
        b_lines = [l for l in en_lines if l.startswith('(b)')]
        base = strip_ab_label(b_lines[0] if b_lines else (en_lines[0] if en_lines else ''))
        if idx is not None and base and len(prob['choices']) > idx:
            prob['en'] = fill_blanks(base, prob['choices'][idx])
        else:
            # 空所なし (下線部言い換え・意味一致など) は英文をそのまま使う
            plain = [re.sub(r'</?u>', '', l) for l in en_q if re.search(r'[a-zA-Z]{3,}', l)]
            prob['en'] = base or (plain[0] if plain else '')
            # 日本語文に対する英文選択問題は正解の選択肢を正解英文とする
            if not prob['en'] and idx is not None and len(prob['choices']) > idx:
                prob['en'] = prob['choices'][idx]
    elif has_ordering:
        prob['type'] = 'ordering'
        clean = strip_fuyou(answer)
        # ことわざ問題 (■1309) は問題文で与え済みの部分が '(A bird) ... (in the bush.)'
        # のように括弧で示されるため、並べ替え対象の語句だけを解答とする
        if re.search(r'\([^()]+\)', clean):
            clean = re.sub(r'\s+', ' ', re.sub(r'\([^()]+\)', '', clean)).strip()
        prob['answer'] = clean
        m = ORDERING_RE.search(question)
        prob['pieces'] = [p.strip() for p in m.group(1).split('/')]
        ordering_line = next((l for l in en_q if ORDERING_RE.search(l)), '')
        rest = ORDERING_RE.sub('', ordering_line).strip(' 　.,!?')
        blank_line = next((l for l in en_q
                           if BLANK_RUN_RE.search(l) and not ORDERING_RE.search(l)), '')
        if not rest and blank_line:
            # 「(　　) (　　) ... he would lose the game.」形式 (■495):
            # 語群行とは別の空所行に解答を流し込む
            sentence = BLANK_RUN_RE.sub(clean + ' ', blank_line, count=1)
        elif ordering_line.startswith(('(a)', '(b)')):
            # (a)/(b) 書き換え形式 (■1094) は語群を含む行のみを正解文に使う
            sentence = ORDERING_RE.sub(lambda _: f' {clean} ', strip_ab_label(ordering_line))
        else:
            # 整序部分が複数行にまたがる場合 (■579) があるため英文行を結合して置換する
            # 「mother（have / ...」のようにカッコ前後に空白がない原文に備えて空白を補う
            sentence = ORDERING_RE.sub(lambda _: f' {clean} ', ' '.join(en_q))
        prob['en'] = re.sub(r'\s+([,.!?;:])', r'\1', re.sub(r'\s+', ' ', sentence)).strip()
        note = re.search(r'[（(]([^（()）]*不要)[)）]', answer)
        if note:
            prob['note'] = note.group(1)
        # 1語不足: 並べ替え UI で解けるように不足語をチップに加える
        missing = re.search(r'[（(]([^（()）]*?)不足[)）]', answer)
        if missing and missing.group(1).strip():
            prob['pieces'].append(missing.group(1).strip())
            prob['note'] = '1語不足の語もチップに含まれています'
    elif has_blank:
        prob['type'] = 'fill'
        prob['answer'] = answer
        en_lines = [l for l in en_q if BLANK_RE.search(l) or HINT_BLANK_RE.search(l)]
        b_lines = [l for l in en_lines if l.startswith('(b)')]
        # 会話表現では空所が複数行 (両話者の発話) にまたがるため結合する
        base = strip_ab_label(b_lines[0] if b_lines else ' '.join(en_lines))
        # '(a) any　(b) No' のように (a)(b) 個別の解答を持つ場合は (b) の文を採用
        ab = re.match(r'\(a\)\s*(.+?)[ 　]+\(b\)\s*(.+)', answer)
        if ab and b_lines:
            prob['en'] = fill_blanks(base, ab.group(2).strip())
        elif base and HINT_BLANK_RE.search(base):
            # 頭文字ヒント付き空所 '( i　) ( d　)' は解答の語を順に流し込む
            words = answer.split('/')[0].split()
            blanks = len(HINT_BLANK_RE.findall(base))
            if len(words) == blanks:
                it = iter(words)
                prob['en'] = re.sub(r'\s+([,.!?;:])', r'\1', re.sub(
                    r'  +', ' ', HINT_BLANK_RE.sub(lambda _: next(it), base))).strip()
            else:
                prob['en'] = re.sub(r'\s+([,.!?;:])', r'\1', re.sub(
                    r'  +', ' ', HINT_BLANK_RE.sub(' '.join(words), base, count=1)
                    .replace('( ', '').replace('　)', ''))).strip()
        else:
            prob['en'] = fill_answer_blanks(base, answer) if base else ''
    else:
        prob['type'] = 'rewrite'
        prob['answer'] = answer
        prob['en'] = re.sub(r'\s+\.', '.', answer).strip()

    return prob


def parse_docx(path):
    lines = docx_lines(path)
    title = next((l.strip() for l in lines if l.strip()), '')
    m = re.match(r'第([０-９0-9]+)章[　\s]*(.+)', title)
    if not m:
        raise ValueError(f'{path}: 章タイトルが見つかりません: {title!r}')
    number = int(unicodedata.normalize('NFKC', m.group(1)))
    name = m.group(2).strip()

    problems = []
    current_num = None
    current = []
    for line in lines:
        s = line.strip()
        # \d は全角数字にもマッチするため半角に限定 (「■２文が…」のような指示行と区別)
        mm = re.match(r'^■([0-9]+)', s)
        if mm:
            if current_num is not None:
                problems.append(parse_problem(current_num, current))
            current_num = mm.group(1)
            current = []
        elif current_num is not None:
            current.append(line)
    if current_num is not None:
        problems.append(parse_problem(current_num, current))

    return {
        'number': number,
        'title': name,
        'fullTitle': f'第{number}章　{name}',
        'problems': problems,
    }


def main():
    paths = sys.argv[1:]
    if not paths:
        print(__doc__)
        sys.exit(1)
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    outdir = os.path.join(root, 'data')
    os.makedirs(outdir, exist_ok=True)
    for path in paths:
        chapter = parse_docx(path)
        out = os.path.join(outdir, f'chapter{chapter["number"]:02d}.json')
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(chapter, f, ensure_ascii=False, indent=2)
        types = {}
        for p in chapter['problems']:
            types[p['type']] = types.get(p['type'], 0) + 1
        print(f'{out}: {chapter["fullTitle"]} 問題数={len(chapter["problems"])} {types}')


if __name__ == '__main__':
    main()
