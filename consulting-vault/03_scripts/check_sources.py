#!/usr/bin/env python3
"""生成物の「根拠が辿れるか」を検査するスクリプト。

何をするか:
  01_clients/<クライアント>/<YYYY-MM>/ 配下の .md（sent/ は除外）を全部開いて、次を確認する。
    ERROR … frontmatter が無い / type・client・sources が無い / sources が空 /
            sources のファイルが実在しない（クライアントフォルダからの相対パスで確認）/
            本文末尾に「本資料は以下の記録から作成:」フッターが無い
    WARN  … 本文中の [[リンク]] の先が vault 内に無い / client がフォルダ名と違う
    INFO  … 「要確認」「要試算」の出現数と行番号（月次レビューで見る。エラーではない）
  HTMLコメント（<!-- -->）の中は「要確認」の数え上げとリンク検査の対象外。
  結果を表で表示し、ERROR が1件でもあれば終了コード 1、無ければ 0 を返す。
  ファイルは一切書き換えない（04_logs/runs.jsonl への実行記録の追記のみ）。

使い方（vault ルートで実行）:
  python3 03_scripts/check_sources.py 2026-09
  python3 03_scripts/check_sources.py 2026-09 --client サンプル商事
  python3 03_scripts/check_sources.py 2026-09 --vault /path/to/consulting-vault

標準ライブラリのみ・Python 3.9 で動作。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

SCRIPT_NAME = "check_sources"
MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
FOOTER_RE = re.compile(r"^本資料は以下の記録から作成\s*[:：]")
HTML_COMMENT_RE = re.compile(r"<!--.*?-->", re.S)
FENCE_RE = re.compile(r"^\s*(```|~~~)")
LINK_RE = re.compile(r"!?\[\[([^\[\]]+?)\]\]")
MARKERS = ("要確認", "要試算")
REQUIRED_KEYS = ("type", "client", "sources")
VALID_TYPES = ("worklog", "minutes", "proposal", "qa", "report")
SKIP_DIRS = {".git", ".obsidian", ".trash", "__pycache__"}


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


def parse_frontmatter(text: str) -> Tuple[Optional[dict], str, int]:
    """先頭の --- ～ --- を簡易YAMLとして読む。

    対応: `key: value` / 空値=None / インラインリスト `[a, b]` / ブロックリスト `  - a`。
    戻り値: (frontmatter の dict または None, 本文, 本文の開始行番号(1始まり))
    """
    if text.startswith("\ufeff"):
        text = text[1:]
    lines = text.split("\n")
    if not lines or lines[0].strip() != "---":
        return None, text, 1
    end = None
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
    if end is None:
        return None, text, 1
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
    return data, body, end + 2


# ---------------------------------------------------------------- 検査

class Finding:
    def __init__(self, level: str, file: str, message: str) -> None:
        self.level = level
        self.file = file
        self.message = message


def blank_comments(body: str) -> str:
    """HTMLコメントを空白に置き換える（改行は残して行番号がずれないようにする）。"""
    return HTML_COMMENT_RE.sub(lambda m: re.sub(r"[^\n]", " ", m.group(0)), body)


def strip_code_fences(lines: List[str]) -> List[str]:
    """``` で囲まれたコードブロックの中身を空行にする（リンク検査の対象外にする）。"""
    out = []
    in_fence = False
    for ln in lines:
        if FENCE_RE.match(ln):
            in_fence = not in_fence
            out.append("")
            continue
        out.append("" if in_fence else ln)
    return out


def build_vault_index(vault: Path) -> Tuple[Set[str], Set[str]]:
    """vault 内の全ファイルの「名前（拡張子あり/なし）」と「vault 相対パス」の集合を作る。"""
    names: Set[str] = set()
    rel_paths: Set[str] = set()
    for root, dirs, files in os.walk(vault):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for fn in files:
            p = Path(root) / fn
            name = nfc(fn)
            names.add(name)
            stem = nfc(p.stem)
            if p.suffix.lower() == ".md":
                names.add(stem)
            rel = nfc(p.relative_to(vault).as_posix())
            rel_paths.add(rel)
            if rel.lower().endswith(".md"):
                rel_paths.add(rel[:-3])
    return names, rel_paths


def link_exists(target: str, names: Set[str], rel_paths: Set[str]) -> bool:
    t = nfc(target.split("|", 1)[0].split("#", 1)[0].strip())
    if t == "":
        return True  # [[#見出し]] のような同一ファイル内リンク
    if "/" in t:
        t = t.lstrip("/")
        return t in rel_paths or any(p.endswith("/" + t) for p in rel_paths)
    return t in names


def find_existing(base: Path, rel: str) -> Optional[Path]:
    """base からの相対パス rel が実在するか（NFC/NFD の違いを吸収して）調べる。"""
    direct = base / rel
    if direct.exists():
        return direct
    cur = base
    for part in Path(rel).parts:
        if part in (".", ""):
            continue
        if part == "..":
            cur = cur.parent
            continue
        if not cur.is_dir():
            return None
        want = nfc(part)
        hit = None
        for child in cur.iterdir():
            if nfc(child.name) == want:
                hit = child
                break
        if hit is None:
            return None
        cur = hit
    return cur


def check_file(path: Path, client_dir: Path, vault: Path,
               names: Set[str], rel_paths: Set[str]) -> Tuple[List[Finding], Dict[str, List[int]]]:
    findings: List[Finding] = []
    disp = nfc(path.relative_to(vault / "01_clients").as_posix())
    marker_lines: Dict[str, List[int]] = {m: [] for m in MARKERS}

    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as e:
        findings.append(Finding("ERROR", disp, f"ファイルを読めない（{e}）"))
        return findings, marker_lines

    fm, body, body_start = parse_frontmatter(text)
    if fm is None:
        findings.append(Finding("ERROR", disp, "frontmatter（先頭の --- で囲んだ設定欄）が無い"))
        fm = {}
    else:
        for key in REQUIRED_KEYS:
            if key not in fm:
                findings.append(Finding("ERROR", disp, f"frontmatter に {key} が無い"))
        if "type" in fm and not fm.get("type"):
            findings.append(Finding("ERROR", disp, "type が空"))
        elif fm.get("type") and str(fm["type"]) not in VALID_TYPES:
            findings.append(Finding("WARN", disp,
                                    f"type「{fm['type']}」が想定外（{' / '.join(VALID_TYPES)} のどれか）"))
        if "client" in fm and not fm.get("client"):
            findings.append(Finding("ERROR", disp, "client が空"))
        elif fm.get("client") and nfc(str(fm["client"])) != nfc(client_dir.name):
            findings.append(Finding("WARN", disp,
                                    f"client「{fm['client']}」がフォルダ名「{nfc(client_dir.name)}」と違う"))

        sources = fm.get("sources")
        if "sources" in fm:
            if sources is None or sources == [] or sources == "":
                findings.append(Finding("ERROR", disp, "sources が空"))
            else:
                if not isinstance(sources, list):
                    sources = [sources]
                client_root = client_dir.resolve()
                for src in sources:
                    src = str(src).strip()
                    if src == "":
                        continue
                    target = (client_dir / src).resolve()
                    try:
                        target.relative_to(client_root)
                    except ValueError:
                        findings.append(Finding("ERROR", disp,
                                                f"sources がクライアントフォルダの外を指している: {src}"))
                        continue
                    hit = find_existing(client_dir, src)
                    if hit is None or not hit.is_file():
                        findings.append(Finding("ERROR", disp, f"sources のファイルが存在しない: {src}"))

    # 本文の検査（HTMLコメントは対象外）
    visible = blank_comments(body)
    lines = visible.split("\n")

    non_blank = [ln.strip() for ln in lines if ln.strip() != ""]
    if not non_blank or not FOOTER_RE.match(non_blank[-1]):
        findings.append(Finding("ERROR", disp, "本文末尾に「本資料は以下の記録から作成:」フッターが無い"))

    for i, ln in enumerate(lines):
        for m in MARKERS:
            c = ln.count(m)
            if c:
                marker_lines[m].extend([body_start + i] * c)

    missing_links: List[str] = []
    for ln in strip_code_fences(lines):
        for m in LINK_RE.finditer(ln):
            target = m.group(1)
            if not link_exists(target, names, rel_paths):
                t = target.split("|", 1)[0].strip()
                if t not in missing_links:
                    missing_links.append(t)
    for t in missing_links:
        findings.append(Finding("WARN", disp, f"リンク先が vault 内に無い: [[{t}]]"))

    counts = {m: len(v) for m, v in marker_lines.items()}
    if any(counts.values()):
        parts = []
        for m in MARKERS:
            if counts[m]:
                uniq = sorted(set(marker_lines[m]))
                where = ", ".join(f"L{n}" for n in uniq[:10]) + (" …" if len(uniq) > 10 else "")
                parts.append(f"{m} {counts[m]}件（{where}）")
        findings.append(Finding("INFO", disp, " / ".join(parts)))
    return findings, marker_lines


def list_client_dirs(vault: Path) -> List[Path]:
    root = vault / "01_clients"
    if not root.is_dir():
        return []
    dirs = [p for p in root.iterdir() if p.is_dir() and not p.name.startswith((".", "_"))]
    return sorted(dirs, key=lambda p: nfc(p.name))


def target_files(month_dir: Path) -> List[Path]:
    files = []
    for root, dirs, fnames in os.walk(month_dir):
        dirs[:] = [d for d in dirs if nfc(d) != "sent" and not d.startswith(".")]
        for fn in fnames:
            if fn.lower().endswith(".md") and not fn.startswith("."):
                files.append(Path(root) / fn)
    return sorted(files, key=lambda p: nfc(p.as_posix()))


def cell(s: str) -> str:
    return str(s).replace("\n", " ").replace("|", "\\|")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="指定月の生成物（sent/ 以外）の sources・フッター・リンク・要確認の数を検査する。"
                    "ERROR があれば終了コード 1。",
    )
    parser.add_argument("month", help="対象月 YYYY-MM")
    parser.add_argument("--client", help="このクライアントだけ検査する（フォルダ名）")
    parser.add_argument("--vault", help="vault ルートのパス（省略時はこのスクリプトの1つ上）")
    args = parser.parse_args(argv)

    month = args.month
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

    names, rel_paths = build_vault_index(vault)
    findings: List[Finding] = []
    marker_total = {m: 0 for m in MARKERS}
    n_files = 0
    no_output: List[str] = []
    for cdir in client_dirs:
        month_dir = cdir / month
        files = target_files(month_dir) if month_dir.is_dir() else []
        if not files:
            no_output.append(nfc(cdir.name))
            continue
        for f in files:
            n_files += 1
            fs, marks = check_file(f, cdir, vault, names, rel_paths)
            findings.extend(fs)
            for m in MARKERS:
                marker_total[m] += len(marks[m])

    order = {"ERROR": 0, "WARN": 1, "INFO": 2}
    findings.sort(key=lambda x: (order.get(x.level, 9), x.file))
    n_err = sum(1 for x in findings if x.level == "ERROR")
    n_warn = sum(1 for x in findings if x.level == "WARN")
    n_info = sum(1 for x in findings if x.level == "INFO")

    print(f"check_sources: 対象月 {month}（検査 {n_files} ファイル、sent/ は対象外）")
    for c in no_output:
        print(f"  {c}: {month}/ に生成物なし")
    if findings:
        print()
        print("| 区分 | ファイル | 内容 |")
        print("|---|---|---|")
        for x in findings:
            print(f"| {x.level} | {cell(x.file)} | {cell(x.message)} |")
    print()
    print(f"結果: ERROR {n_err} / WARN {n_warn} / INFO {n_info}"
          f"（要確認 {marker_total['要確認']}件 / 要試算 {marker_total['要試算']}件）")
    if n_err:
        print("→ ERROR を直して再実行してください。")
    else:
        print("→ ERROR なし。")

    append_run_log(vault, {
        "month": month,
        "client": nfc(args.client) if args.client else None,
        "files": n_files,
        "error": n_err, "warn": n_warn, "info": n_info,
        "要確認": marker_total["要確認"], "要試算": marker_total["要試算"],
    })
    return 1 if n_err else 0


if __name__ == "__main__":
    sys.exit(main())
