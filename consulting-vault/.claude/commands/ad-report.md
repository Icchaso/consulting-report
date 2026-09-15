---
description: Meta広告のエクスポートから広告運用レポート（PowerPoint 60〜90枚）を作る。数字はスクリプトが計算し、Claude は所見の文章だけをプレースホルダで書く
argument-hint: <クライアント名> [YYYY-MM]
allowed-tools: Bash(python3:*), Bash(date:*), Bash(ls:*), Read, Write, Edit, Glob, Grep
---

# /ad-report — 広告運用レポート（PowerPoint）の作成

引数: `$ARGUMENTS`（1つ目 = クライアント名（`01_clients/` のフォルダ名）、2つ目 = `YYYY-MM`。月が空なら前月）

このコマンドは vault ルート（`CLAUDE.md` がある場所）で実行する。以下の手順を上から順に、省略せずに実行すること。

## 絶対に守るルール（大原則「ソースに遡れないことは書かない」の広告版）
- **数字を1つも書かない。** 金額・件数・率・倍率・日付の数字は、すべて `{{キー}}` のプレースホルダで書く（例: `CPAは{{total.cpa}}で前月比{{mom.cpa.pct}}`）。build_deck.py が analysis.json の値に置き換える。日付も `{{day.2026-09-21.label}}` `{{week.W3.label}}` で書ける。
- **使えるキーは analysis.json の `values` にあるものだけ。** 存在しないキーを書くとビルドがエラーで止まる。迷ったら `insights.template.json` の `placeholders` を見るか、`Grep` で analysis.json を検索して確かめる。キーを推測で作らない。
- **所見ごとに `sources` を必ず書く。** analysis.json のキー（`total.cpa` / `campaign.C1` など）か、クライアントフォルダからの相対パス（`2026-09/minutes/2026-09-17_月次定例.md` など）。
- **書けないスライドは `text` を空のままにする。** スライドに「所見：要確認」と出る。一般論・推測・それらしい文章で埋めない。空欄は失敗ではない。
- **次月施策（`next_actions`）は、その月の議事録・提案書に実際に書いてある施策だけ。** 新しい施策を考えない。記録が無ければ空のまま。
- 文体は `_profile.md` の `style:` に従う（空なら「です・ます調。簡潔に。」）。`_profile.md` の「注意点」（伏せる社名など）を守る。
- `raw/` と `sent/`、`_profile.md`、`02_templates/`、`03_scripts/` は書き換えない。analysis.json と insights.template.json も手で直さない（作り直すときはスクリプトを再実行する）。
- スクリプトがエラーを出したら、スクリプトを書き換えずに利用者に報告する（insights.json の直しで解決するものだけ自分で直す）。

## 手順

### 0. 準備
1. `date "+%Y-%m-%d"` で今日の日付を確認し、対象月を決める（引数の2つ目が `YYYY-MM` ならそれ、空なら前月）。クライアント名が無い・形式が違うときは止めて利用者に聞く。以下 `<client>` `<月>` と書く。
2. `CLAUDE.md` と `01_clients/<client>/_profile.md` を読む（`style:`、目標KPI、注意点）。
3. `ls 01_clients/<client>/ads/<月>/` で広告データがあるか確認する。無ければ止めて、次を利用者に伝える:
   - 置き場所: `01_clients/<client>/ads/<月>/`（前月比を出すなら前月分も `ads/<前月>/` に）
   - 必須: 広告マネージャで「レベル=広告」「内訳=日（期間の分割=1日）」で書き出した CSV（キャンペーン名・広告セット名・広告名・日・消化金額・インプレッション等の列）
   - 任意: 内訳（年齢・性別／配置／デバイス／地域）の CSV、キャンペーン単位の期間合計 CSV（正しいリーチ用）
   - 目標対比を出すなら `_profile.md` の frontmatter に `ad_kpi_targets: [cpa=3000, roas=3.0]` を書いてもらう（Claude は `_profile.md` を書き換えない）

### 1. 集計
`python3 03_scripts/ad_report/ad_analyze.py --client <client> --month <月>` を実行する。
- `01_clients/<client>/<月>/ads/analysis.json` と `insights.template.json` ができる。
- `[WARN]`（未知の列・読み飛ばした行・内訳と明細のずれ など）は最後の報告に全部書く。

### 2. 材料を読む
1. `insights.template.json` を読む（大きいので分けて読んでよい）。`slides` の各スライドID・`title`・`hint`・`max_chars`・`placeholders`（使えるキーと今の値の見本）を把握する。
2. analysis.json は大きいので全部は読まない。必要な値は template の `placeholders` の見本で確認し、足りなければ `Grep` で `"campaign.C2.cpa"` のようにキーを検索する。
3. `01_clients/<client>/<月>/minutes/`・`proposals/`・`qa/` があれば全部読む（次月施策・所見の根拠）。

### 3. 所見を書く（insights.json）
`insights.template.json` を丸ごとコピーして、同じフォルダに `insights.json` として書く。書き換えてよいのは各スライドの `text` と `sources` だけ（`title` `hint` `placeholders` などは残してよい。ビルドは読まない）。
- `text`: 置換後で `max_chars` 程度。改行で段落・箇条書きになる。事実（数字の大小・増減・順位）と、そこから言える評価だけを書く。
- 優先して書くスライド: `summary`（今月の結論）/ `mom` / `campaigns` と各 `campaign.C*` / `ranking` / 各内訳（`age_gender` `placement` `device` `region`）/ `issues` / `next_actions`。広告セット・広告ごとのスライドは、データから明確に言えることがあるものだけでよい。
- 注意:
  - 目的（`campaign.C*.result_label`）が違うキャンペーンを CPA だけで比べない。認知目的のキャンペーンは CPA が高くて当然。
  - 値に「（参考値）」が付くもの（日別合算のリーチ・フリークエンシー）は、確定値のように書かない。
  - `mom.*.trend` は「改善／悪化／増加／減少／横ばい」の文字（指標の良し悪しの向きで判定済み）。自分で逆の評価を書かない。
  - CV が少ない広告の CPA は振れやすい（ランキングは CV が `min_cv_for_ranking` 件以上だけ）。
  - 外部要因（季節・競合・アルゴリズム等）の推測を書かない。書くなら議事録・Q&A にある場合だけで、そのファイルを sources に入れる。

### 4. ビルド
`python3 03_scripts/ad_report/build_deck.py --client <client> --month <月>` を実行する。
- 「python-pptx がありません」と出たら止めて、利用者に `pip3 install -r 03_scripts/ad_report/requirements.txt` を案内する（自分でインストールしない）。
- 「解決できないプレースホルダ」エラーは insights.json を直して（キーを analysis.json で確かめて）再実行する。

### 5. 検査
`python3 03_scripts/ad_report/check_deck.py --client <client> --month <月>` を実行する。
- ERROR が出たら insights.json を直し → 手順4（ビルド）→ 手順5（検査）をやり直す（最大3回）。3回で直らなければ止めて報告する。
- WARN「生の数字が混ざっている」は、その数字をプレースホルダに書き換えて直す（日付の「9月」などの月名は WARN にならない）。
- 「要確認」が残るのは正常。無理に埋めない。

## 最後に利用者へ報告すること（前置き不要・箇条書き）
1. 生成物: `01_clients/<client>/<月>/ads/<client>_<月>_広告運用レポート.pptx` のパス、スライド枚数とセクション別の枚数（check_deck.py の出力）
2. 所見: 書いたスライド数と「要確認」の数、要確認の主な理由（データ不足・議事録なし 等）
3. 集計の警告: ad_analyze.py の `[WARN]`、check_deck.py の WARN
4. レビューで見るべき点（例）:
   - エグゼクティブサマリーと所見・課題の文章が、表・グラフの数字と矛盾していないか（各スライドのノートに出どころ）
   - 目標対比（`_profile.md` の `ad_kpi_targets`）が今月の目標として正しいか
   - 「計測上の注意」スライド（アトリビューション設定・Meta の計測仕様の変更）が今回の比較に当てはまるか
   - 次月施策が議事録・提案書の内容どおりか、先方に出してよい表現か
   - 送付前に PowerPoint で開いて、文字のはみ出し・グラフの表示を目で確認する
