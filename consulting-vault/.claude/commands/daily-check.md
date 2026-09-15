---
description: 直近24時間の inbox 振り分け結果と inbox の残りを1〜3行で要約する（ファイルは動かさない）
allowed-tools: Read, Glob, Grep, Bash(tail:*), Bash(ls:*), Bash(date:*)
---

# /daily-check — 振り分け結果の確認

このコマンドは vault ルート（`CLAUDE.md` がある場所）で実行する。**読むだけ。ファイルの移動・改名・削除・編集は一切しない。**

## 手順
1. `date "+%Y-%m-%dT%H:%M:%S"` で現在時刻を確認する。
2. `tail -n 500 04_logs/inbox_sort.jsonl` でログを読み、`ts` が現在から24時間以内の行だけを使う。
   - 1行1イベント: `{"ts","action","src","dest","client","type","reason"}`。`action` は moved / skipped / conflict / unsorted / error。
   - ログが無い・24時間以内の行が無いときは「直近24時間の振り分けなし（inbox_sort.py が動いていない可能性）」と扱う。
3. `ls -1 00_inbox/` と `ls -1 00_inbox/_要確認/` で、振り分けられずに残っているファイルを確認する（`_done/` `_要確認/` のフォルダ自体と `.gitkeep` は数えない）。
4. 未振り分け・衝突のファイルがあるときは、`01_clients/*/_profile.md` の `client:` と `aliases:` を読み、どう直せば振り分けられるかを考える。

## 出力（1〜3行。前置き不要）
- 1行目: 振り分け結果。`moved` をクライアント別・種別別に数える。種別の表記は meeting→会議、slack→Slack、line→LINE、mail→メール、doc→資料、memo→メモ。
  例: 「サンプル商事に会議1件・Slack3件、テスト工務店にメール1件を振り分け。」
- 2行目: 未振り分け（unsorted / conflict / error と、00_inbox・_要確認 に残っているファイル）の件数と理由。無ければ「未振り分けなし。」
  例: 「未振り分け1件（2026-09-14_mtg_サンプル.md: 種別 mtg が不明）。」
- 3行目（未振り分けがあるときだけ）: ファイル名の直し方の**提案**。規則は `YYYY-MM-DD_<種別>_<相手>_<件名>.md`、種別は meeting / slack / line / mail / doc / memo、相手は `_profile.md` の `client` か `aliases` と完全一致。
  例: 「→ `2026-09-14_meeting_サンプル_定例.md` に改名すると サンプル商事 に振り分けられます。」
- `_要確認/` の衝突（同名の raw が既にある）は改名案ではなく「raw の既存ファイルと中身を見比べて判断が必要」と書く。
- 改名・移動は提案だけ。実行するかは利用者が決める。
