---
type: worklog
client: {{client}}
date: {{month}}
generated_at: {{generated_at}}
sources:
  - {{source_path}}
---

# 稼働ログ {{client}} {{month}}

<!-- 書き方: このファイルは形式の見本。実物は python3 03_scripts/worklog.py YYYY-MM が丸ごと作る。Claude が書き直してよいのは「要点」「作業分類」の2列だけ。 -->

## 活動一覧

| 日付 | 種別 | 作業分類 | 相手 | 要点 | 所要目安（推定） | ソース |
|---|---|---|---|---|---|---|
| {{MM-DD}} | {{type}} | {{category}} | {{client}} | {{summary}} | {{estimated}} | [[{{source_file}}]] |

<!-- 書き方: 作業分類は 準備 / 会議 / 調査 / 提案 / 相談対応 / フォロー / 資料作成 のどれか1つ（判断できなければ「要確認」）。要点は raw の中身から1行・40字以内。日付・種別・相手・所要目安・ソースは変えない。 -->

## 月末集計（推定）

| 区分 | 件数 | 推定時間 |
|---|---|---|
| 会議 | {{meeting_count}}件 | {{meeting_time}} |
| チャット相談 | {{chat_count}}件 | {{chat_time}} |
| メール | {{mail_count}}通 | {{mail_time}} |
| 資料 | {{doc_count}}件 | {{doc_time}} |
| メモ | {{memo_count}}件 | {{memo_time}} |
| **合計** | **{{total_count}}件** | **{{total_time}}** |

※ 所要目安はメッセージ数・録音の長さから機械的に算出した推定値（実際の作業時間ではない）。

<!-- 書き方: 月末集計と推定根拠は worklog.py の機械算出値。Claude は変えない。提案の本数はここに入れず、月次レポートで proposals/ のファイル数を数える。 -->

---
本資料は以下の記録から作成: {{sources_links}}
