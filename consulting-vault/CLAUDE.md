# CLAUDE.md — コンサル業務記録 vault

## このvaultの目的
コンサルタントの実際の業務記録（録音の文字起こし・チャット・メール・資料・メモ）を蓄積し、
そこから議事録・提案書・Q&A記録・稼働ログ・月次レポートを生成してストックする。

## 大原則（すべてに優先する）
**ソースに遡れないことは書かない。書けない項目は空欄か「要確認」のまま残す。**
それらしい文章で埋めるくらいなら、空欄で返す。空欄は失敗ではない。

---

## フォルダと4つのルール
```
00_inbox/            入口。利用者が置く。_done/ = 振り分け済みの元ファイル、_要確認/ = 衝突など
01_clients/<クライアント名>/
  _profile.md        利用者が手で書く前提知識
  raw/               振り分け済みの生データ（ソース）
  ads/YYYY-MM/       広告マネージャのエクスポート（CSV）。raw と同じくソースなので編集しない
  YYYY-MM/           生成物（worklog.md / minutes/ / proposals/ / qa/ / report.md / ads/＝広告レポート）
    sent/            実際に送付した確定版と 送付記録.md
02_templates/        テンプレート
03_scripts/          Pythonスクリプト
04_logs/             処理ログ（jsonl）
99_samples/          動作確認用のダミーデータ
```
1. `raw/` は編集しない。ソースとして固定されている。
2. `YYYY-MM/` 配下（`sent/` を除く）は生成物。作り直してよい（前版はGitに残る）。
3. `sent/` は「納品した事実」。読むだけ。書き込み・上書き・削除をしない。
4. `_profile.md` は利用者が書く。Claude は書き換えない（誤りに気づいたら指摘だけする）。

クライアントのフォルダ名は `_profile.md` の `client:` と同じ。

---

## 生成ルール（資料を作るときは必ず守る）
1. **sources 以外を書かない。** frontmatter（ファイル先頭の設定欄）の `sources:` に列挙したファイルの内容だけで書く。
   先に読むファイルを決めて `sources:` に書き、それから本文を書く。使わなかったファイルは `sources:` から外す。
   `_profile.md` の内容（背景・目標KPI）を本文に使ったら `_profile.md` も `sources:` に入れる（文体の指定に使うだけなら不要）。
2. **書けない項目は「要確認」。** 見出しは消さずに残し、本文に「要確認」と書く。一般論や推測で埋めない。
3. **数字はソースにあるものだけ。** 根拠のない効果試算は「要試算」と書く。Claude が試算した場合は
   「試算（生成時点の仮定：〜）」と明記し、仮定をすべて書く。所要時間は raw の `estimated_minutes` を使い、「（推定）」を付ける。
4. **文体は `_profile.md` の `style:` に従う。** 指定が空なら「です・ます調。簡潔に。」
5. **末尾にフッターを付ける。** 下記の形式で、`sources:` の全ファイルを Obsidian リンク（拡張子なし）で並べる。

### 生成物の frontmatter（この形式以外にしない）
```yaml
---
type: minutes
client: サンプル商事
date: 2026-09-14
generated_at: 2026-10-01T09:00:00
sources:
  - raw/2026-09-14_meeting_サンプル商事_月次定例.md
---
```
- `type:` は worklog / minutes / proposal / qa / report のどれか。
- `date:` は対象日（YYYY-MM-DD）。worklog と report は対象月（YYYY-MM）。
- frontmatter の中に `#` コメントを書かない。
- `sources:` のパスは**クライアントフォルダからの相対パス**（`raw/…` / `2026-09/minutes/…` / `2026-09/worklog.md` / `_profile.md`）。
- `sources:` はブロックリスト形式（`  - パス`）で書く。

### フッター（本文の最後に必ず付ける）
```
---
本資料は以下の記録から作成: [[2026-09-14_meeting_サンプル商事_月次定例]] / [[_profile]]
```

### その他の書き方
- テンプレートの見出し構成を変えない。見出しの順番も変えない。
- テンプレートの `<!-- 書き方: … -->` コメントに従い、生成物からはコメントを消す。`{{ }}` を残さない。
- 1つの資料に別クライアントの情報を混ぜない。

---

## テンプレートと保存先
| 資料 | テンプレート | 保存先（クライアントフォルダ内） |
|---|---|---|
| 稼働ログ | `02_templates/worklog.md` | `YYYY-MM/worklog.md` |
| 議事録（会議ごと） | `02_templates/minutes.md` | `YYYY-MM/minutes/YYYY-MM-DD_<会議名>.md` |
| 提案書 | `02_templates/proposal.md` | `YYYY-MM/proposals/YYYY-MM-DD_<テーマ>.md` |
| Q&A記録（相談1件1ファイル） | `02_templates/qa.md` | `YYYY-MM/qa/YYYY-MM-DD_<テーマ>.md` |
| 月次レポート | `02_templates/report.md` | `YYYY-MM/report.md` |
| クライアントプロフィール | `02_templates/_profile.md` | `_profile.md`（利用者が書く） |

### 作る順番（依存関係）
1. `worklog.md`（raw から。まず `worklog.py` で作る）
2. `minutes/`（meeting の raw から）、`qa/`（slack / line / mail の raw から）
3. `proposals/`（議事録やチャットで**実際に提案したもの**だけ。新しい提案を考えない）
4. `report.md`（上の生成物から。raw は KPI 実績値のときだけ直接使う）

### 作業分類の語彙（これ以外の語を使わない）
準備 / 会議 / 調査 / 提案 / 相談対応 / フォロー / 資料作成
判断できないときは「要確認」。

### worklog.md の扱い
- worklog.md は `python3 03_scripts/worklog.py YYYY-MM` が丸ごと作る（テンプレートは形式の見本）。
- Claude が書き直してよいのは **「要点」列と「作業分類」列だけ**。スクリプトは要点に件名、作業分類に種別からの初期値（meeting→会議 / slack・line→相談対応 / mail→フォロー / doc→資料作成 / memo→フォロー）を入れているので、raw の中身を読んで直す。
- 日付・種別・相手・所要目安（推定）・ソースの各列、行の数と順番、「月末集計（推定）」、推定根拠、frontmatter、フッターは変えない（機械算出値）。
- 提案の本数は worklog には入れない（worklog.py は提案書より先に動くため）。月次レポートで `proposals/` のファイル数を数える。

---

## スクリプト（すべて vault ルートで実行）
| コマンド | 役割 |
|---|---|
| `python3 03_scripts/inbox_sort.py [--dry-run] [--notify]` | `00_inbox/` をクライアント別 `raw/` に振り分ける。`--dry-run` は移動せず結果だけ表示 |
| `python3 03_scripts/worklog.py YYYY-MM` | 指定月の `worklog.md` を raw から作る |
| `python3 03_scripts/check_sources.py YYYY-MM` | 指定月の生成物の `sources:` が実在するか・フッターがあるか・「要確認」の残数を確認する |
| `python3 03_scripts/send_log.py --client <名> --month YYYY-MM --to <宛先> --files <ファイル...>` | 送付記録を `sent/送付記録.md` に追記する |

| `python3 03_scripts/ad_report/ad_analyze.py --client <名> --month YYYY-MM` | `01_clients/<名>/ads/YYYY-MM/` の広告マネージャ CSV を集計し、`YYYY-MM/ads/analysis.json` と所見の雛形を作る |
| `python3 03_scripts/ad_report/build_deck.py --client <名> --month YYYY-MM` | analysis.json と insights.json から広告運用レポート（PowerPoint）を作る。python-pptx が必要 |
| `python3 03_scripts/ad_report/check_deck.py --client <名> --month YYYY-MM` | 所見の未解決プレースホルダ・生の数字の混入・要確認の数を検査する |

- `send_log.py` は利用者が送付した後に、**利用者の指示があったときだけ**実行する。
- 広告レポートでは **Claude は数字を1つも書かない**。所見の数字は `{{total.cpa}}` のようなプレースホルダで書き、build_deck.py が analysis.json の値に置き換える（詳細は `/ad-report` の手順）。
- スクリプトがエラーを出したら、スクリプトを書き換えずにエラー内容を利用者に報告する。

## スラッシュコマンド
- `/monthly [YYYY-MM]` … 指定月（省略時は前月）の全クライアント分の資料を生成し、最後に `check_sources.py` で確認する。
- `/daily-check` … `inbox_sort.py` の振り分け結果（`04_logs/`）を要約して報告する。
- `/ad-report <クライアント名> [YYYY-MM]` … 広告データを集計し、所見を書き、広告運用レポート（PowerPoint 60〜90枚）を作って検査する。生成先は `01_clients/<名>/YYYY-MM/ads/`。

---

## 禁止事項
- `raw/` と `sent/` のファイルを編集・上書き・削除・移動すること。
- `sources:` に無い情報を書くこと。記憶・一般論・他クライアントの情報で補うこと。
- 数字を創作すること（金額・件数・時間・KPI・効果）。「要試算」「要確認」で止める。
- Web検索の結果を出典URLなしで書くこと。Web検索は利用者が頼んだときだけ使い、URLを必ず本文に残す。
- ファイルを削除すること。`00_inbox/` の中身は `inbox_sort.py` が移動する。Claude が手で動かしたり消したりしない。
- `_profile.md` と `02_templates/` を利用者の指示なしに書き換えること。
- `git push`・リモートの追加。Git はローカルのみ。コミットも利用者が頼んだときだけ。
- `.env` を読む・表示すること。APIキーを本文やログに書くこと。
- メール送信・チャット投稿など、外部に何かを送ること（送付は利用者が行う）。
- 音声ファイルや原本ファイル（PDF・Excel等）を外部サービスに送ること。

## 迷ったとき
- ソースに書いてあるか分からない → 書かずに「要確認」。
- ファイル名からクライアントが判定できない → 振り分けず、利用者に聞く。
- ルール同士がぶつかる → 「大原則」を優先し、利用者に確認する。
