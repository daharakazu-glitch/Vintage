# Vintage 学習アプリ

英文法・語法問題集「Vintage (4th Edition)」の章別学習アプリと進捗ダッシュボードです。

## 構成

| パス | 内容 |
| --- | --- |
| `index.html` | 学習ダッシュボード（Google ログインで進捗を記録・表示） |
| `chapterNN.html` | 章別学習アプリ（自動生成） |
| `data/chapterNN.json` | パース済みの問題データ |
| `source_docs/` | 元の docx ファイル |
| `tools/parse_vintage_docx.py` | docx → JSON パーサー |
| `tools/generate_apps.py` | JSON → HTML 生成スクリプト |
| `tools/templates/` | アプリ・ダッシュボードのテンプレート |

## アプリの機能

- **問題演習**: 4択・誤り指摘は選択ボタン、整序英作文は語句チップの並べ替え、空所補充・書き換えは解答表示＋自己採点
- **音声**: 解答確認後に正解英文の音声再生（通常／ゆっくり）
- **録音評価**: マイクで英文を読み上げると音声認識で採点（80点以上で自動マスター）
- **ランダムテスト**: 選択式問題から10問出題
- **PDF ダウンロード**: チェックした問題から6種類のPDFを作成
  1. 完全な問題リスト（解答・和訳付き）
  2. 4択・誤り指摘テスト
  3. 空所補充テスト（選択肢なし）
  4. 整序英作文テスト
  5. 和文英訳テスト
  6. ランダム混合テスト
- **進捗記録**: Firebase (Google ログイン) で章ごとのマスター状況・テスト結果を保存し、ダッシュボードに表示

## 章の追加手順（第6章〜第30章）

1. 章の docx を `source_docs/` に置く（例: `source_docs/chapter06_動名詞.docx`）
2. パースして JSON を生成:

   ```bash
   python3 tools/parse_vintage_docx.py source_docs/chapter06_動名詞.docx
   ```

3. アプリとダッシュボードを再生成:

   ```bash
   python3 tools/generate_apps.py
   ```

ダッシュボードには全30章の枠があり、生成済みの章だけがリンクとして有効になります。

## docx の形式

```
第１章　時制          ← 章タイトル（1行目）
■001                 ← 問題番号
基本                  ← 任意タグ（基本／発展）
●超頻出              ← 任意タグ
(問題文 1〜複数行)
①xx　②yy　③zz　④ww  ← 4択の場合の選択肢
#③                   ← 解答（行頭 #）
(和訳 1〜複数行)
```

誤り指摘問題は問題文中に `①<u>...</u>` 形式、整序英作文は `(to / at / high school / ...)` 形式で記述します。

## 進捗データ

Firestore の `leap_users/{uid}/progress/vintage_chapterNN` に保存します
（既存の LEAP アプリと同じコレクションを使い、`vintage_` プレフィックスで区別）。
