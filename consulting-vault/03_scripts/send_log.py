#!/usr/bin/env python3
"""クライアントへ送付した事実を sent/送付記録.md に追記するスクリプト。

何をするか:
  01_clients/<クライアント>/<YYYY-MM>/sent/ に**既に置いてある**確定版ファイルについて、
  送付日時・宛先・ファイル名・SHA256（ファイルの指紋）の先頭12桁・メモを
  sent/送付記録.md の表に1ファイル1行で追記する。
  - 送付記録.md が無ければ表の見出し付きで新規作成する
  - 既存の行は一切変更しない（ファイル末尾への追記のみ）
  - ファイルのコピー・移動はしない。sent/ に無いファイルを指定するとエラーで終了する
  - --dry-run なら追記する予定の行を表示するだけで、何も書き込まない

使い方（vault ルートで実行。利用者が実際に送付した後に使う）:
  python3 03_scripts/send_log.py --client サンプル商事 --month 2026-09 \\
      --to "佐藤 様（sample-shoji.example）" --files 月次レポート_2026-09.pdf 議事録_0914.pdf \\
      [--note "メールで送付"] [--dry-run] [--vault PATH]

標準ライブラリのみ・Python 3.9 で動作。実行ごとに 04_logs/runs.jsonl に1行追記する（--dry-run 時は書かない）。
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import List, Optional, Tuple

SCRIPT_NAME = "send_log"
MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
RECORD_NAME = "送付記録.md"
HEADER = "| 送付日時 | 宛先 | 添付ファイル | SHA256先頭12桁 | メモ |"
SEPARATOR = "|---|---|---|---|---|"


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


def cell(s: str) -> str:
    return str(s).replace("\r", " ").replace("\n", " ").replace("|", "\\|").strip()


def sha256_12(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()[:12]


def find_child(parent: Path, name: str) -> Optional[Path]:
    """parent 直下から NFC で一致する名前のファイル/フォルダを探す。"""
    if not parent.is_dir():
        return None
    want = nfc(name)
    direct = parent / name
    if direct.exists():
        return direct
    for child in parent.iterdir():
        if nfc(child.name) == want:
            return child
    return None


def resolve_in_sent(sent_dir: Path, spec: str) -> Tuple[Optional[Path], str]:
    """--files の1件を sent/ 内のファイルに解決する。戻り値: (パス or None, エラー理由)"""
    s = spec.strip()
    if s == "":
        return None, "空のファイル名"
    p = Path(s)
    sent_root = sent_dir.resolve()
    if p.is_absolute():
        # フルパスで渡された場合も、sent/ の中を指していれば受け付ける
        try:
            rel_parts = p.resolve().relative_to(sent_root).parts
        except ValueError:
            return None, "sent/ の外のファイルは記録できません"
    else:
        parts = list(p.parts)
        if parts and nfc(parts[0]) == "sent":
            parts = parts[1:]
        if any(x == ".." for x in parts):
            return None, "「..」を含むパスは使えません"
        rel_parts = tuple(parts)
    cur = sent_dir
    for part in rel_parts:
        nxt = find_child(cur, part)
        if nxt is None:
            return None, "sent/ にありません（先に確定版を sent/ に置いてください）"
        cur = nxt
    if not cur.is_file():
        return None, "ファイルではありません"
    if nfc(cur.name) == RECORD_NAME and cur.parent.resolve() == sent_root:
        return None, f"{RECORD_NAME} 自体は添付ファイルとして記録できません"
    return cur, ""


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="sent/ に置いた確定版の送付記録を sent/送付記録.md に追記する（既存行は変更しない）。",
    )
    parser.add_argument("--client", required=True, help="クライアント名（01_clients/ のフォルダ名）")
    parser.add_argument("--month", required=True, help="対象月 YYYY-MM（どの月の sent/ か）")
    parser.add_argument("--to", required=True, help="宛先（例: \"佐藤 様 sato@sample-shoji.example\"）")
    parser.add_argument("--files", required=True, nargs="+",
                        help="sent/ に置いてあるファイル名（複数可）")
    parser.add_argument("--note", default="", help="メモ（任意）")
    parser.add_argument("--dry-run", action="store_true", help="追記せず、追記予定の行を表示するだけ")
    parser.add_argument("--vault", help="vault ルートのパス（省略時はこのスクリプトの1つ上）")
    args = parser.parse_args(argv)

    if not MONTH_RE.match(args.month):
        print(f"エラー: --month は YYYY-MM 形式で指定してください（受け取った値: {args.month}）",
              file=sys.stderr)
        return 2
    vault = Path(args.vault).expanduser().resolve() if args.vault else default_vault()
    clients_root = vault / "01_clients"
    if not clients_root.is_dir():
        print(f"エラー: {vault} に 01_clients/ がありません。--vault を確認してください", file=sys.stderr)
        return 2
    if args.to.strip() == "":
        print("エラー: --to（宛先）が空です", file=sys.stderr)
        return 2

    client = nfc(args.client.strip())
    client_dir = find_child(clients_root, client)
    if client_dir is None or not client_dir.is_dir():
        print(f"エラー: クライアント「{client}」のフォルダが 01_clients/ にありません", file=sys.stderr)
        return 1
    month_dir = find_child(client_dir, args.month)
    sent_dir = find_child(month_dir, "sent") if month_dir is not None else None
    if sent_dir is None or not sent_dir.is_dir():
        print(f"エラー: 01_clients/{client}/{args.month}/sent/ がありません。"
              "先に確定版ファイルを sent/ に置いてください", file=sys.stderr)
        return 1

    resolved: List[Tuple[str, str]] = []
    errors: List[str] = []
    for spec in args.files:
        path, reason = resolve_in_sent(sent_dir, spec)
        if path is None:
            errors.append(f"  {spec}: {reason}")
            continue
        rel = nfc(path.relative_to(sent_dir).as_posix())
        resolved.append((rel, sha256_12(path)))
    if errors:
        print("エラー: 次のファイルを記録できません。何も追記していません。", file=sys.stderr)
        for e in errors:
            print(e, file=sys.stderr)
        return 1

    sent_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    rows = [
        "| " + " | ".join([sent_at, cell(args.to), cell(name), f"`{digest}`", cell(args.note)]) + " |"
        for name, digest in resolved
    ]

    record = sent_dir / RECORD_NAME
    existing = find_child(sent_dir, RECORD_NAME)
    if existing is not None:
        record = existing
    rel_record = nfc(record.relative_to(vault).as_posix())

    if args.dry_run:
        print(f"[dry-run] {rel_record} に次の {len(rows)} 行を追記する予定（書き込みはしていません）:")
        if existing is None:
            print("[dry-run] （送付記録.md が無いので、表の見出し付きで新規作成します）")
        for r in rows:
            print(r)
        return 0

    if existing is None:
        head = "\n".join([
            f"# 送付記録 {nfc(client_dir.name)} {args.month}",
            "",
            "<!-- 03_scripts/send_log.py が行を追記する。納品した事実の記録なので、既存の行を書き換えない・消さない。 -->",
            "",
            HEADER,
            SEPARATOR,
            "",
        ])
        with open(record, "x", encoding="utf-8", newline="\n") as f:
            f.write(head)
    else:
        # 末尾が改行で終わっていない場合だけ改行を足す（既存の内容には触れない）
        with open(record, "rb") as f:
            f.seek(0, 2)
            size = f.tell()
            last = b""
            if size:
                f.seek(-1, 2)
                last = f.read(1)
        if size and last != b"\n":
            with open(record, "a", encoding="utf-8", newline="\n") as f:
                f.write("\n")
    with open(record, "a", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(rows) + "\n")

    print(f"{rel_record} に {len(rows)} 行を追記しました（宛先: {args.to}）")
    for name, digest in resolved:
        print(f"  {name}  SHA256先頭12桁 {digest}")
    append_run_log(vault, {
        "client": nfc(client_dir.name), "month": args.month, "to": args.to,
        "files": [{"name": n, "sha256_12": d} for n, d in resolved],
        "note": args.note, "record": rel_record,
    })
    return 0


if __name__ == "__main__":
    sys.exit(main())
