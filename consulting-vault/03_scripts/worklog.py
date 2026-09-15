#!/usr/bin/env python3
"""稼働ログ（worklog.md）を raw の frontmatter から機械的に作るスクリプト。

何をするか:
  01_clients/<クライアント>/raw/ にある記録ファイルの frontmatter（date / type / title /
  estimated_minutes / estimate_basis）を読み、対象月の分だけを1ファイル1行の表にして
  01_clients/<クライアント>/<YYYY-MM>/worklog.md に書き出す（生成物なので毎回上書き）。
  表の下に「月末集計（推定）」と推定根拠の一覧を付ける。
  対象月の raw が1件もないクライアントは worklog.md を作らず「記録なし」と表示する。

  raw の中身（本文）は読まない。要点列・作業分類列は機械的な初期値なので、
  /monthly で Claude が raw を読んで書き直す。所要目安と集計はこのスクリプトの値が正。

使い方（vault ルートで実行）:
  python3 03_scripts/worklog.py                  # 前月分・全クライアント
  python3 03_scripts/worklog.py 2026-09          # 2026年9月分
  python3 03_scripts/worklog.py 2026-09 --client サンプル商事
  python3 03_scripts/worklog.py 2026-09 --vault /path/to/consulting-vault

標準ライブラリのみ・Python 3.9 で動作。実行ごとに 04_logs/runs.jsonl に1行追記する。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

SCRIPT_NAME = "worklog"
RAW_SUFFIXES = (".md", ".txt")
MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
FILENAME_DATE_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})_")

# 種別 → 作業分類の機械的な初期値（/monthly で書き直してよい）
CATEGORY_BY_TYPE = {
    "meeting": "会議",
    "slack": "相談対応",
    "line": "相談対応",
    "mail": "フォロー",
    "doc": "資料作成",
    "memo": "フォロー",
}

# 月末集計の区分（表示順）: (区分名, 対象の種別, 件数の単位)
SUMMARY_GROUPS = [
    ("会議", ("meeting",), "件"),
    ("チャット相談", ("slack", "line"), "件"),
    ("メール", ("mail",), "通"),
    ("資料", ("doc",), "件"),
    ("メモ", ("memo",), "件"),
]

UNKNOWN = "要確認"


# ---------------------------------------------------------------- 共通の小道具

def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def default_vault() -> Path:
    return Path(__file__).resolve().parent.parent


def now_iso() -> str:
    return dt.datetime.now().replace(microsecond=0).isoformat()


def append_run_log(vault: Path, summary: dict) -> None:
    log_dir = vault / "04_logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    line = json.dumps({"ts": now_iso(), "script": SCRIPT_NAME, "summary": summary},
                      ensure_ascii=False)
    with open(log_dir / "runs.jsonl", "a", encoding="utf-8", newline="\n") as f:
        f.write(line + "\n")


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
    """先頭の --- ～ --- を簡易YAMLとして読む。

    対応: `key: value` / 空値=None / インラインリスト `[a, b]` / ブロックリスト `  - a`。
    戻り値: (frontmatter の dict または None, 本文)
    """
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


def fmt_minutes(total: int) -> str:
    """58 → 「58分」、70 → 「1時間10分」、120 → 「2時間」。"""
    total = int(total)
    if total < 60:
        return f"{total}分"
    h, m = divmod(total, 60)
    return f"{h}時間{m}分" if m else f"{h}時間"


def previous_month(today: dt.date) -> str:
    first = today.replace(day=1)
    last_prev = first - dt.timedelta(days=1)
    return last_prev.strftime("%Y-%m")


def cell(s: str) -> str:
    """Markdown の表のセルに入れられるように | と改行を逃がす。"""
    return str(s).replace("\r", " ").replace("\n", " ").replace("|", "\\|").strip()


def link_name(filename: str) -> str:
    """Obsidian のリンク名。.md は拡張子なし、それ以外（.txt 等）は拡張子付き。"""
    p = Path(filename)
    return p.stem if p.suffix.lower() == ".md" else p.name


def to_int(v) -> Optional[int]:
    if v is None or isinstance(v, list):
        return None
    try:
        return int(round(float(str(v).strip())))
    except ValueError:
        return None


# ---------------------------------------------------------------- 本体

def list_client_dirs(vault: Path) -> List[Path]:
    clients_root = vault / "01_clients"
    if not clients_root.is_dir():
        return []
    dirs = [p for p in clients_root.iterdir()
            if p.is_dir() and not p.name.startswith((".", "_"))]
    return sorted(dirs, key=lambda p: nfc(p.name))


def collect_rows(client_dir: Path, month: str) -> Tuple[List[dict], List[str]]:
    """対象月の raw を集めて行データにする。戻り値: (行, 警告)"""
    rows: List[dict] = []
    warnings: List[str] = []
    raw_dir = client_dir / "raw"
    if not raw_dir.is_dir():
        return rows, warnings
    client_name = nfc(client_dir.name)
    for p in raw_dir.iterdir():
        if not p.is_file() or p.name.startswith(".") or p.suffix.lower() not in RAW_SUFFIXES:
            continue
        fname = nfc(p.name)
        try:
            text = p.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError) as e:
            warnings.append(f"読めないファイルを飛ばしました: {fname}（{e}）")
            continue
        fm, _ = parse_frontmatter(text)
        fm = fm or {}
        date = str(fm.get("date") or "").strip()
        date_from_name = False
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", date):
            m = FILENAME_DATE_RE.match(fname)
            if not m:
                warnings.append(f"date が読めないため飛ばしました: {fname}")
                continue
            date = m.group(1)
            date_from_name = True
        if not date.startswith(month + "-"):
            continue
        if date_from_name:
            warnings.append(f"frontmatter に date が無いのでファイル名の日付を使いました: {fname}")
        rtype = str(fm.get("type") or "").strip()
        if rtype == "":
            parts = Path(fname).stem.split("_", 3)
            rtype = parts[1] if len(parts) > 1 else ""
        title = fm.get("title")
        title = "" if title is None or isinstance(title, list) else str(title).strip()
        basis = fm.get("estimate_basis")
        basis = "" if basis is None or isinstance(basis, list) else str(basis).strip()
        rows.append({
            "date": date,
            "type": rtype,
            "category": CATEGORY_BY_TYPE.get(rtype, UNKNOWN),
            "client": nfc(str(fm.get("client") or client_name)),
            "title": title,
            "minutes": to_int(fm.get("estimated_minutes")),
            "basis": basis,
            "filename": fname,
        })
    rows.sort(key=lambda r: (r["date"], r["filename"]))
    return rows, warnings


def summarize(rows: List[dict]) -> Tuple[List[Tuple[str, int, str, int]], int, int, int]:
    """戻り値: ([(区分, 件数, 単位, 推定分)], 合計件数, 合計推定分, 所要目安が不明の件数)"""
    groups = []
    known_types = set()
    for label, types, unit in SUMMARY_GROUPS:
        known_types.update(types)
        rs = [r for r in rows if r["type"] in types]
        groups.append((label, len(rs), unit, sum(r["minutes"] or 0 for r in rs)))
    others = [r for r in rows if r["type"] not in known_types]
    if others:
        groups.append(("その他", len(others), "件", sum(r["minutes"] or 0 for r in others)))
    total_min = sum(r["minutes"] or 0 for r in rows)
    unknown = sum(1 for r in rows if r["minutes"] is None)
    return groups, len(rows), total_min, unknown


def render(client: str, month: str, rows: List[dict]) -> str:
    out: List[str] = []
    out.append("---")
    out.append("type: worklog")
    out.append(f"client: {client}")
    out.append(f"date: {month}")
    out.append(f"generated_at: {now_iso()}")
    out.append("sources:")
    for r in rows:
        out.append(f"  - raw/{r['filename']}")
    out.append("---")
    out.append("")
    out.append(f"# 稼働ログ {client} {month}")
    out.append("")
    out.append("<!-- このファイルは 03_scripts/worklog.py が raw の frontmatter から機械的に生成したもの。")
    out.append("     「要点」列と「作業分類」列は /monthly で Claude が raw の中身を読んで書き直してよい（行の追加・削除・並べ替えはしない）。")
    out.append("     作業分類は 準備 / 会議 / 調査 / 提案 / 相談対応 / フォロー / 資料作成 のどれか1つ。")
    out.append("     「所要目安（推定）」列と「月末集計（推定）」「推定根拠」は機械算出値なので書き換えないこと。")
    out.append("     直したい場合は raw の記録を見直してから worklog.py を再実行する。 -->")
    out.append("")
    out.append("## 活動一覧")
    out.append("")
    out.append("| 日付 | 種別 | 作業分類 | 相手 | 要点 | 所要目安（推定） | ソース |")
    out.append("|---|---|---|---|---|---|---|")
    for r in rows:
        mm_dd = r["date"][5:]
        minutes = fmt_minutes(r["minutes"]) if r["minutes"] is not None else UNKNOWN
        out.append("| " + " | ".join([
            cell(mm_dd),
            cell(r["type"] or UNKNOWN),
            cell(r["category"]),
            cell(r["client"]),
            cell(r["title"] or UNKNOWN),
            cell(minutes),
            f"[[{link_name(r['filename'])}]]",
        ]) + " |")
    out.append("")

    groups, total_n, total_min, unknown = summarize(rows)
    out.append("## 月末集計（推定）")
    out.append("")
    out.append("| 区分 | 件数 | 推定時間 |")
    out.append("|---|---|---|")
    for label, n, unit, mins in groups:
        out.append(f"| {label} | {n}{unit} | {fmt_minutes(mins)} |")
    out.append(f"| **合計** | **{total_n}件** | **{fmt_minutes(total_min)}** |")
    out.append("")
    out.append("※ 所要目安はメッセージ数・録音の長さから機械的に算出した推定値（実際の作業時間ではない）。"
               "係数は 03_scripts/config.json で変更できる。")
    if unknown:
        out.append(f"※ 所要目安が{UNKNOWN}の {unknown}件 は合計時間に含めていない。")
    out.append("")
    out.append("<details>")
    out.append(f"<summary>推定根拠（{total_n}件）</summary>")
    out.append("")
    for r in rows:
        minutes = fmt_minutes(r["minutes"]) if r["minutes"] is not None else "算出なし"
        basis = r["basis"] or "根拠の記載なし"
        out.append(f"- {r['date'][5:]} {r['type'] or '種別不明'} {link_name(r['filename'])}: "
                   f"{minutes}（{basis}）")
    out.append("")
    out.append("</details>")
    out.append("")
    out.append("---")
    out.append("本資料は以下の記録から作成: "
               + " / ".join(f"[[{link_name(r['filename'])}]]" for r in rows))
    out.append("")
    return "\n".join(out)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="raw の frontmatter から稼働ログ（<クライアント>/<YYYY-MM>/worklog.md）を生成する。",
    )
    parser.add_argument("month", nargs="?", help="対象月 YYYY-MM（省略時は前月）")
    parser.add_argument("--client", help="このクライアントだけ処理する（フォルダ名）")
    parser.add_argument("--vault", help="vault ルートのパス（省略時はこのスクリプトの1つ上）")
    args = parser.parse_args(argv)

    month = args.month or previous_month(dt.date.today())
    if not MONTH_RE.match(month):
        print(f"エラー: 月は YYYY-MM 形式で指定してください（受け取った値: {month}）", file=sys.stderr)
        return 2
    vault = Path(args.vault).expanduser().resolve() if args.vault else default_vault()
    if not (vault / "01_clients").is_dir():
        print(f"エラー: {vault} に 01_clients/ がありません。--vault を確認してください", file=sys.stderr)
        return 2

    client_dirs = list_client_dirs(vault)
    if args.client:
        want = nfc(args.client.strip())
        client_dirs = [d for d in client_dirs if nfc(d.name) == want]
        if not client_dirs:
            print(f"エラー: クライアント「{want}」のフォルダが 01_clients/ にありません", file=sys.stderr)
            return 2

    print(f"worklog: 対象月 {month}")
    summary: Dict[str, object] = {"month": month, "clients": {}, "no_records": [], "warnings": 0}
    for cdir in client_dirs:
        client = nfc(cdir.name)
        rows, warnings = collect_rows(cdir, month)
        for w in warnings:
            print(f"  [WARN] {client}: {w}", file=sys.stderr)
        summary["warnings"] = int(summary["warnings"]) + len(warnings)  # type: ignore[arg-type]
        if not rows:
            print(f"  {client}: 記録なし（worklog.md は作成しません）")
            summary["no_records"].append(client)  # type: ignore[union-attr]
            continue
        out_dir = cdir / month
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / "worklog.md"
        with open(out_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(render(client, month, rows))
        _, total_n, total_min, unknown = summarize(rows)
        rel = out_path.relative_to(vault).as_posix()
        extra = f"、所要目安{UNKNOWN} {unknown}件" if unknown else ""
        print(f"  {client}: {total_n}件 → {rel}（推定合計 {fmt_minutes(total_min)}{extra}）")
        summary["clients"][client] = {  # type: ignore[index]
            "count": total_n, "total_minutes": total_min,
            "unknown_minutes": unknown, "path": nfc(rel),
        }
    append_run_log(vault, summary)
    return 0


if __name__ == "__main__":
    sys.exit(main())
