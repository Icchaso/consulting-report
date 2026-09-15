# HANDOFF — コンサル業務記録ツール（レポート自動生成）

最終更新: 2026-09-14

## 1. 目的
メタ広告運用コンサルの業務記録（会議録音・Slack・メモ）と広告データ（広告マネージャのCSV）から、
議事録・提案書・Q&A・稼働ログ・月次レポート、そして **広告運用レポート（PowerPoint 50〜100枚）** を自動生成する。
大原則:「ソースに遡れないことは書かない」。数字はスクリプトが計算し、Claude は文章だけ（広告は数字をプレースホルダで書く）。
三好さんと共同制作。形ができたら GitHub に上げて三好さんを招待し、いっちゃんは途中で外れる予定。

## 2. 現状（2026-09-15 時点・GitHub: Icchaso/consulting-report）
- 設計書 STEP1〜4 は完成し、ダミーデータで最後まで通った
  - `inbox_sort.py`（文字起こし .md/.txt・Slackエクスポートの振り分け）→ `worklog.py` → `/monthly`（議事録・提案書・Q&A・月次レポート）→ `check_sources.py`（ERROR 0）
- 広告レポート（設計書に無い追加要件）は試作が完成
  - `ad_report/ad_analyze.py` → `insights.json`（Claude が所見をプレースホルダで記入）→ `build_deck.py`（python-pptx・ネイティブグラフ）→ `check_deck.py`
  - サンプルで82枚。Quick Look での目視は良好。**PowerPoint / Keynote の実機では未確認**（作業時は画面ロックで Keynote が応答しなかった）
- 自動テスト: 記録系56件（`/usr/bin/python3` 3.9）＋ 広告系26件（python-pptx が入った python3）すべてパス
- vault 内の `01_clients/サンプル商事` `テスト工務店` は架空サンプル。入口・生成物もサンプルを流した結果が入っている

## 3. 次にやること
- [x] `/ad-report サンプル商事 2026-09` を本物の vault でヘッドレス実行 → 82枚・所見40/53枠・check_deck ERROR 0 / WARN 0（2026-09-14）
- [ ] 改善: 次月施策スライドの「根拠にした記録」に、その月の議事録・提案書が**全部**並ぶ（在庫の話など広告と無関係な資料名が先方に出る）→ 所見の sources に実際に使ったものだけ出す
- [ ] 改善: 所見の地の文に少数の件数（「最も多い2日」「2つ」）が残る。check_deck の生数字判定を件数表現にも広げるか検討
- [ ] PowerPoint 実機で開いて見た目確認（フォント「游ゴシック」の見え方・グラフの数値書式）
- [ ] **いっちゃん判断**: 三好さんの実際のレポート（パワポ）の実物 or 目次 → 構成・デザインを寄せる
- [ ] **いっちゃん判断**: 納品形式（pptx / Googleスライド / PDF）、会社テンプレートの有無
- [x] GitHub: https://github.com/Icchaso/consulting-report （Private）に push（2026-09-15, d10c352）。ルートの .gitignore で実データ（01_clients のサンプル2社以外・00_inbox・04_logs・.env）を除外済み
- [ ] 三好さんをコラボレーターに招待（**三好さんの GitHub ユーザー名待ち**）
- [ ] いっちゃんが抜けるとき: Settings → Transfer ownership で三好さんへ移管（三好さんは1日以内に承認）→ 自動でコラボレーターになるので自分を外す。移管後に README の `git clone` の URL を新しい持ち主に書き換える（旧URLも自動転送はされる）
- [ ] サンプルの記録（在庫の話）を広告運用コンサルの会議内容に差し替え、広告レポートの次月施策とつながるデモにする
- [ ] 実データでの検証: 広告マネージャの実エクスポートで列名を確定（column_map.json）
- [ ] 設計書 STEP6（LINE / Gmail 取り込み、定時実行、通知、Meta Marketing API 連携）は運用が回ってから

## 4. 注意
- `raw/` `ads/`（ソース）と `sent/`（納品の事実）は編集禁止。`.claude/settings.json` の deny でも防いでいる
- 記録系スクリプトは標準ライブラリのみ・Python 3.9 互換。広告の `build_deck.py` だけ python-pptx が必要（`pip3 install -r 03_scripts/ad_report/requirements.txt`）
- python-pptx は棒＋折れ線の2軸複合グラフを作れない。日本語フォントは `a:ea` を別途設定している
- Meta の計測仕様変更（2026-01-12 ビュースルー7日/28日廃止、2026-03 クリックスルー＝リンククリックのみ）でレポートの前月比に段差が出うる → 「計測上の注意」スライドに自動記載
- `結果` 列は目的ごとに意味が違うので合算しない。リーチは足せない（期間合計CSVが無いと参考値）
- ツール内に個人パス・個人名を入れない（三好さんに渡すため）。仕上げ時に grep で確認する
- ヘッドレス実行の例: `claude -p "/monthly 2026-09" --permission-mode acceptEdits --allowedTools "Read" "Write" "Edit" "Glob" "Grep" "Bash(python3:*)" "Bash(date:*)" "Bash(ls:*)"`

## 5. ファイル
- 設計書: `コンサル業務記録ツール_設計書/`（図解PDF・詳細版MD/PDF）
- データ契約（全スクリプト共通の約束）: `開発メモ/データ契約.md`
- vault: `consulting-vault/`
  - `CLAUDE.md`（生成ルール・禁止事項）/ `README.md`（利用者向け使い方）
  - `02_templates/`（議事録・提案書・Q&A・稼働ログ・月次レポート・プロフィール）
  - `03_scripts/`（inbox_sort / worklog / check_sources / send_log / config.json / tests）
  - `03_scripts/ad_report/`（ad_analyze / build_deck / check_deck / column_map.json / deck_config.json / tests / tools/make_sample.py）
  - `.claude/commands/`（monthly / daily-check / ad-report）
  - `99_samples/`（ダミーの録音・Slack・広告CSV）

## 6. 検証
```bash
cd consulting-vault
/usr/bin/python3 -m unittest discover -s 03_scripts/tests          # 記録系 56件
python3 -m unittest discover -s 03_scripts/ad_report/tests         # 広告系 26件（python-pptx 必要）
python3 03_scripts/check_sources.py 2026-09                        # 生成物の出典チェック（ERROR 0 で exit 0）
python3 03_scripts/ad_report/check_deck.py --client サンプル商事 --month 2026-09
```
