#!/usr/bin/env python3
"""広告レポート（ad_analyze / build_deck / check_deck）で共有する小道具。

- vault ルート・クライアントフォルダの解決（macOS の NFD ファイル名に注意して NFC で比較）
- frontmatter の簡易パーサ（03_scripts/worklog.py と同じ仕様。他担当のファイルに依存しないよう複製）
- プレースホルダ {{campaign.C1.cpa}} の解決
- 所見の文章に紛れ込んだ「生の数字」の検出
- 04_logs/runs.jsonl への実行記録

標準ライブラリのみ・Python 3.9 で動作。
"""
from __future__ import annotations

import datetime as dt
import json
import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")

# プレースホルダの書式: {{ key }}。key は英数字と _ - + . : のみ（. が階層の区切り）
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]*?)\s*\}\}")
PLACEHOLDER_KEY_RE = re.compile(r"^[A-Za-z0-9_\-+:.]+$")

# 生の数字の検出（プレースホルダを取り除いたあとの文章に対して使う）
_DIGIT = "[0-9０-９]"
_NUM = _DIGIT + "[0-9０-９,，.．]*"
RAW_NUMBER_RES = [
    re.compile(r"[¥￥$＄]\s*" + _DIGIT),
    re.compile(_NUM + r"\s*(%|％|円|件|回|倍|pt|ｐｔ|ポイント|人|名|万|千|億|本|枚|秒|割|クリック|imp|IMP|CV|cv)"),
    re.compile(r"(CPA|CPC|CPM|CTR|CVR|ROAS|ＣＰＡ|ＲＯＡＳ)\s*[:：は]?\s*" + _DIGIT),
]

# 表示に使う記号
NA_DISPLAY = "—"
UNKNOWN_INSIGHT = "所見：要確認"


# ---------------------------------------------------------------- 基本

def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def default_vault() -> Path:
    """03_scripts/ad_report/ad_common.py → vault ルート"""
    return Path(__file__).resolve().parent.parent.parent


def script_dir() -> Path:
    return Path(__file__).resolve().parent


def now_iso() -> str:
    return dt.datetime.now().replace(microsecond=0).isoformat()


def append_run_log(vault: Path, script: str, summary: dict) -> None:
    """データ契約8章: 04_logs/runs.jsonl に {"ts","script","summary"} を1行追記。"""
    log_dir = vault / "04_logs"
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        line = json.dumps({"ts": now_iso(), "script": script, "summary": summary},
                          ensure_ascii=False)
        with open(log_dir / "runs.jsonl", "a", encoding="utf-8", newline="\n") as f:
            f.write(line + "\n")
    except OSError:
        pass  # ログが書けなくても本処理は止めない


def resolve_vault(arg: Optional[str]) -> Path:
    return Path(arg).expanduser().resolve() if arg else default_vault()


def find_client_dir(vault: Path, client: str) -> Optional[Path]:
    root = vault / "01_clients"
    if not root.is_dir():
        return None
    want = nfc(client.strip())
    for p in root.iterdir():
        if p.is_dir() and nfc(p.name) == want:
            return p
    return None


def month_shift(month: str, delta: int) -> str:
    y, m = int(month[:4]), int(month[5:7])
    idx = y * 12 + (m - 1) + delta
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def previous_month_of_today() -> str:
    return month_shift(dt.date.today().strftime("%Y-%m"), -1)


def rel_to(path: Path, base: Path) -> str:
    try:
        return nfc(path.resolve().relative_to(base.resolve()).as_posix())
    except ValueError:
        return nfc(path.resolve().as_posix())


def load_json(path: Path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


_LEAF_OBJ_RE = re.compile(r"\{\n\s+([^{}\[\]]*?)\n\s*\}")
_LEAF_ARR_RE = re.compile(r"\[\n\s+([^{}\[\]]*?)\n\s*\]")


def dumps_compact(data) -> str:
    """indent 付きで書くが、入れ子の無い {…} と […] は1行にまとめる（読みやすさとサイズの両立）。"""
    s = json.dumps(data, ensure_ascii=False, indent=1)
    s = _LEAF_OBJ_RE.sub(lambda m: "{" + re.sub(r"\n\s+", " ", m.group(1)) + "}", s)
    s = _LEAF_ARR_RE.sub(lambda m: "[" + re.sub(r"\n\s+", " ", m.group(1)) + "]", s)
    return s


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(dumps_compact(data))
        f.write("\n")


# ---------------------------------------------------------------- frontmatter（worklog.py と同仕様）

def _parse_scalar(raw: str):
    v = raw.strip()
    if v == "":
        return None
    if len(v) >= 2 and v[0] == v[-1] and v[0] in ("'", '"'):
        return v[1:-1]
    if v.startswith("[") and v.endswith("]"):
        inner = v[1:-1].strip()
        if inner == "":
            return []
        items = []
        for part in inner.split(","):
            p = part.strip()
            if len(p) >= 2 and p[0] == p[-1] and p[0] in ("'", '"'):
                p = p[1:-1]
            if p != "":
                items.append(p)
        return items
    return v


def parse_frontmatter(text: str) -> Tuple[Optional[dict], str]:
    """先頭の --- ～ --- を簡易YAMLとして読む（key: value / 空値=None / [a, b] / ブロックリスト）。"""
    if text.startswith("\ufeff"):
        text = text[1:]
    lines = text.split("\n")
    if not lines or lines[0].strip() != "---":
        return None, text
    end = None
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
    if end is None:
        return None, text
    data: dict = {}
    last_key: Optional[str] = None
    for line in lines[1:end]:
        line = line.rstrip("\r")
        if line.strip() == "" or line.lstrip().startswith("#"):
            continue
        stripped = line.strip()
        if stripped.startswith("- ") or stripped == "-":
            if last_key is not None:
                item = _parse_scalar(stripped[1:])
                cur = data.get(last_key)
                if not isinstance(cur, list):
                    cur = []
                if item is not None:
                    cur.append(item if not isinstance(item, list) else ", ".join(item))
                data[last_key] = cur
            continue
        if ":" not in line:
            continue
        key, _, value = line.partition(":")
        key = key.strip()
        if key == "":
            continue
        data[key] = _parse_scalar(value)
        last_key = key
    body = "\n".join(lines[end + 1:])
    return data, body


# ---------------------------------------------------------------- プレースホルダ

def insight_text(slot: Optional[dict]) -> str:
    """insights.json の1スロットから本文を取り出す（text は文字列か文字列の配列）。"""
    if not isinstance(slot, dict):
        return ""
    t = slot.get("text", "")
    if isinstance(t, list):
        t = "\n".join(str(x) for x in t if str(x).strip() != "")
    return str(t or "").strip()


def insight_sources(slot: Optional[dict]) -> List[str]:
    if not isinstance(slot, dict):
        return []
    s = slot.get("sources") or []
    if isinstance(s, str):
        s = [s]
    return [str(x).strip() for x in s if str(x).strip() != ""]


def find_placeholders(text: str) -> List[str]:
    return [m.group(1) for m in PLACEHOLDER_RE.finditer(text)]


def resolve_placeholders(text: str, values: Dict[str, dict]) -> Tuple[str, List[str]]:
    """{{key}} を values[key]["display"] に置き換える。戻り値: (置換後, 解決できなかったキー)

    閉じていない {{ や }} も「解決できなかった」として返す。
    """
    unknown: List[str] = []

    def repl(m: "re.Match") -> str:
        key = m.group(1).strip()
        entry = values.get(key)
        if not PLACEHOLDER_KEY_RE.match(key) or entry is None:
            unknown.append(key)
            return m.group(0)
        disp = str(entry.get("display", NA_DISPLAY))
        if entry.get("reference"):
            disp += "（参考値）"  # 足し算できない値（リーチ等）の合算は、文章でも参考値と明記する
        return disp

    out = PLACEHOLDER_RE.sub(repl, text)
    rest = PLACEHOLDER_RE.sub("", text)
    if "{{" in rest or "}}" in rest:
        unknown.append("（閉じていない {{ または }} があります）")
    return out, unknown


def find_raw_numbers(text: str) -> List[str]:
    """プレースホルダ以外の場所に書かれた「数字＋単位」等を返す（重複なし・出現順）。"""
    stripped = PLACEHOLDER_RE.sub(" ", text)
    found: List[str] = []
    for rx in RAW_NUMBER_RES:
        for m in rx.finditer(stripped):
            s = m.group(0).strip()
            if s not in found:
                found.append(s)
    return found


def is_file_source(src: str) -> bool:
    """sources の要素がファイルパス（クライアントフォルダからの相対）か、analysis.json のキーか。"""
    return "/" in src or src.endswith((".md", ".txt", ".csv", ".json", ".xlsx"))


def analysis_key_exists(key: str, values: Dict[str, dict], analysis: Optional[dict] = None) -> bool:
    """sources に書かれた analysis.json のキーが実在するか。

    values のキーと完全一致 / values のスコープ（例 campaign.C1）/ analysis.json の構造をたどれるパス
    （例 rankings.zero_cv）のどれかなら実在とみなす。
    """
    k = key.strip()
    if k.startswith("analysis:"):
        k = k[len("analysis:"):]
    if k in values:
        return True
    prefix = k.rstrip(".") + "."
    if any(v.startswith(prefix) for v in values):
        return True
    if analysis is not None:
        cur = analysis
        for part in k.split("."):
            if isinstance(cur, dict) and part in cur:
                cur = cur[part]
            elif isinstance(cur, list) and part.isdigit() and int(part) < len(cur):
                cur = cur[int(part)]
            else:
                return False
        return True
    return False


# ---------------------------------------------------------------- 文字幅（レイアウトの見積もり用）

def text_units(s: str) -> float:
    """全角=1.0、半角=0.55 として文字列の幅を em 単位で見積もる。"""
    total = 0.0
    for ch in s:
        if unicodedata.east_asian_width(ch) in ("W", "F", "A"):
            total += 1.0
        elif ch.isupper() or ch in "%@#&mwMW¥":
            total += 0.68
        else:
            total += 0.55
    return total
