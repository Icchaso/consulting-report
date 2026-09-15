#!/bin/bash
# HTML を画像（PNG）に変換する。README の図・資料サンプル画像の書き出しに使う。
# 使い方: html2png.sh <入力.html> <出力.png> [幅] [高さ] [倍率]
#   例:   html2png.sh docs/src/overview.html docs/images/overview.png 1000 700 2
#
# 待ち時間の上限は 90秒。変えたいときは環境変数 TIMEOUT で指定する。
set -euo pipefail

IN="${1:?入力HTMLのパスを指定してください}"
OUT="${2:?出力PNGのパスを指定してください}"
W="${3:-1080}"
H="${4:-1350}"
SCALE="${5:-1}"
TIMEOUT="${TIMEOUT:-90}"

[ -f "$IN" ] || { echo "エラー: 入力HTMLが見つかりません: $IN" >&2; exit 1; }

CHROME=""
for c in \
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  "/Applications/Chromium.app/Contents/MacOS/Chromium" \
  "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge" \
  "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"; do
  [ -x "$c" ] && { CHROME="$c"; break; }
done
[ -n "$CHROME" ] || {
  echo "エラー: Google Chrome が見つかりません。" >&2
  echo "　　　  https://www.google.com/chrome/ から入れてから、もう一度お試しください（5分ほどで終わります）。" >&2
  exit 1; }

# HTML が指している画像が、本当にそこにあるか確かめる。
# 無いまま書き出しても Chrome はエラーを出さず、写真の入っていない画像ができてしまう。
# それを「できました」と報告すると、写真なしの画像がそのまま投稿される。
IN_DIR="$(cd "$(dirname "$IN")" && pwd)"
refs="$(grep -Eo 'src=("[^"]*"|'"'"'[^'"'"']*'"'"')' "$IN" | sed -E 's/^src=.//; s/.$//' || true)"
missing=""
while IFS= read -r ref; do
  [ -n "$ref" ] || continue
  case "$ref" in http://*|https://*|//*|data:*) continue ;; esac
  case "$ref" in /*) refpath="$ref" ;; *) refpath="${IN_DIR}/${ref}" ;; esac
  [ -f "$refpath" ] || missing="${missing}${ref}
"
done <<REFS
$refs
REFS

if [ -n "$missing" ]; then
  echo "エラー: 指定された画像が見つかりません:" >&2
  printf '%s' "$missing" | sed 's/^/　　　  /' >&2
  echo "　　　  ファイル名をもう一度確かめてください。" >&2
  echo "　　　  （このまま進めると、写真の入っていない画像ができあがります）" >&2
  exit 1
fi

mkdir -p "$(dirname "$OUT")"
ABS_IN="$(cd "$(dirname "$IN")" && pwd)/$(basename "$IN")"

# 一時ファイルに書き出してから入れ替える。
# 出力先に前回の画像が残っていると、書き出し前に「大きさが変わらない＝書き終わった」と誤判定し、
# 古い画像のまま「できました」と言ってしまうため（2026-09-15 に実際に起きた）。
TMP_OUT="$(dirname "$OUT")/.tmp-$$-$(basename "$OUT")"
rm -f "$TMP_OUT"

# 使い捨ての作業用プロファイルを使う。
# これを指定しないと、利用者が普段開いている Chrome と設定フォルダを奪い合い、
# 数分間なにも起きない（固まったように見える）ことがある。
PROFILE="$(mktemp -d "${TMPDIR:-/tmp}/doc-shot.XXXXXX")"
trap 'rm -rf "$PROFILE"' EXIT

# Chrome を後ろで動かし、画像が書き終わった時点で止める。
# Chrome はスクリーンショットを書いたあと、すぐに終了しないことがある。
# 終了を待つと数分かかるので、書き終わりを見て自分で止める。
"$CHROME" --headless --disable-gpu --hide-scrollbars \
  --user-data-dir="$PROFILE" --no-first-run --no-default-browser-check \
  --force-device-scale-factor="$SCALE" \
  --window-size="${W},${H}" \
  --screenshot="$TMP_OUT" \
  "file://${ABS_IN}" >/dev/null 2>&1 &
PID=$!
# Chrome を確実に止める。起動直後は TERM を無視することがあるので、効かなければ KILL する。
stop_chrome() {
  kill -TERM "$PID" 2>/dev/null || true
  n=0
  while kill -0 "$PID" 2>/dev/null && [ "$n" -lt 6 ]; do sleep 0.5; n=$(( n + 1 )); done
  kill -KILL "$PID" 2>/dev/null || true
  wait "$PID" 2>/dev/null || true
}
trap 'stop_chrome; rm -rf "$PROFILE"; rm -f "$TMP_OUT"' EXIT

max=$(( TIMEOUT * 2 ))   # 0.5秒きざみで待つ
i=0
prev=-1
done_ok=0
while [ "$i" -lt "$max" ]; do
  if [ -s "$TMP_OUT" ]; then
    now=$(wc -c < "$TMP_OUT" | tr -d ' ')
    # 大きさが前回と同じなら、書き終わったとみなす
    if [ "$now" = "$prev" ]; then done_ok=1; break; fi
    prev="$now"
  fi
  kill -0 "$PID" 2>/dev/null || { done_ok=1; break; }   # Chrome が自分で終わった
  sleep 0.5
  i=$(( i + 1 ))
done

stop_chrome

if [ ! -s "$TMP_OUT" ] || [ "$done_ok" != "1" ]; then
  echo "エラー: 画像の書き出しが ${TIMEOUT}秒 以内に終わりませんでした: $OUT" >&2
  echo "　　　  Chrome を一度すべて終了してから、もう一度お試しください。" >&2
  exit 1
fi

# 壊れた画像でないか確かめる
sips -g pixelWidth "$TMP_OUT" >/dev/null 2>&1 || { echo "エラー: 画像が壊れています: $OUT" >&2; exit 1; }
mv -f "$TMP_OUT" "$OUT"

echo "できました: $OUT ($(sips -g pixelWidth -g pixelHeight "$OUT" | awk '/pixelWidth/{w=$2}/pixelHeight/{h=$2}END{print w"x"h}'))"
