#!/usr/bin/env python3
"""inbox_sort.py — 00_inbox の中身を整形して、クライアント別の raw/ に振り分ける。

■ 何をするか
  00_inbox/ に置かれた文字起こし・メモ・Slack エクスポートを読み、
  01_clients/<クライアント>/raw/ に「frontmatter（日付・種別・所要時間の推定など）＋元の本文」の
  Markdown として保存する。処理が済んだ元ファイルは 00_inbox/_done/<実行日>/ に移す。
  判定できないものは inbox に残し、理由をログに書く。

■ 使い方
  python3 03_scripts/inbox_sort.py              # 振り分けを実行
  python3 03_scripts/inbox_sort.py --dry-run    # 何が起きるかだけ表示（ファイル・ログは一切変えない）
  python3 03_scripts/inbox_sort.py --notify     # 結果を macOS 通知でも出す
  python3 03_scripts/inbox_sort.py --vault PATH # vault の場所を指定（テスト用。既定は 03_scripts の1つ上＝vault ルート）

■ 振り分けルール
  1. ファイル名が「YYYY-MM-DD_<種別>_<相手>_<件名>.md / .txt」の形のもの
     - 種別: meeting / slack / line / mail / doc / memo（件名は省略可。アンダースコアを含んでもよい）
     - <相手> が 01_clients/*/_profile.md の client: か aliases: と完全一致（NFC 正規化後）したら振り分ける
     - 保存名は YYYY-MM-DD_<種別>_<client正式名>_<件名>.md（相手が別名でも正式名に揃える）
     - meeting は本文のタイムスタンプ（行頭 or 括弧内のものだけ）の最大値から duration_min を出す
  2. Slack エクスポート（zip、または展開済みフォルダ。channels.json / users.json / <チャンネル>/<日付>.json）
     - チャンネル名を _profile.md の slack_channels と照合し、1チャンネル1日1ファイルの Markdown にする
     - 参加・退出などのシステムメッセージは除外。<@U123> は名前に置換。スレッド返信は親の下にインデント
     - どのクライアントにも紐づかないチャンネルは無視し、未振り分けとしてログに残す
  3. それ以外（規則外のファイル名・相手が判定できない・未対応形式）は inbox に残し、未振り分けとしてログに残す

■ 守っていること
  - raw は絶対に上書きしない。同名の raw があり内容（imported_at を除く）が同じなら skipped、
    違えば新しい方を 00_inbox/_要確認/ に置いて conflict とする
  - 00_inbox/_done/ と 00_inbox/_要確認/ と隠しファイル（.DS_Store 等）は処理しない
  - 1件でエラーが起きても他のファイルの処理は続ける（エラーはログに残し、元ファイルは inbox に残す）
  - 同じ入力で何度実行しても壊れない（冪等）
  - 標準ライブラリのみ・Python 3.9 で動く

■ ログ
  04_logs/inbox_sort.jsonl … 1行1イベント（moved / skipped / conflict / unsorted / error）
  04_logs/runs.jsonl       … 実行ごとのサマリー1行
"""
from __future__ import annotations

import argparse
import json
import math
import re
import shutil
import subprocess
import sys
import unicodedata
import zipfile
from datetime import date, datetime
from pathlib import Path, PurePosixPath
from typing import Any, Dict, List, Optional, Tuple

SCRIPT_NAME = "inbox_sort"
TYPES = ("meeting", "slack", "line", "mail", "doc", "memo")
TYPE_LABELS = {
    "meeting": "会議",
    "slack": "Slack",
    "line": "LINE",
    "mail": "メール",
    "doc": "資料",
    "memo": "メモ",
}
TEXT_EXTS = (".md", ".txt")
DONE_DIR = "_done"
REVIEW_DIR = "_要確認"
SKIP_NAMES = {DONE_DIR, REVIEW_DIR}

DEFAULT_CONFIG: Dict[str, Any] = {
    "timezone": "Asia/Tokyo",
    "estimate": {
        "meeting_chars_per_minute": 300,
        "chat_minutes_per_message": 2,
        "chat_min_minutes": 10,
        "mail_minutes": 10,
        "doc_minutes": 30,
        "memo_minutes": 5,
    },
}

# raw frontmatter のキー順（データ契約5章）
RAW_KEYS = (
    "date",
    "type",
    "client",
    "title",
    "source_file",
    "imported_at",
    "duration_min",
    "message_count",
    "estimated_minutes",
    "estimate_basis",
)

# Slack で「人の発言」として扱う subtype（これ以外の subtype は channel_join 等のシステムメッセージとみなして除外）
SLACK_HUMAN_SUBTYPES = {None, "", "thread_broadcast", "file_share", "me_message", "bot_message"}

FILENAME_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})_([^_]+)_(.+)$")
SLACK_DAY_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})\.json$")


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


# ---------------------------------------------------------------------------
# 設定・プロフィール
# ---------------------------------------------------------------------------

def load_config(vault: Path) -> Dict[str, Any]:
    """03_scripts/config.json を読む。無い・壊れている・値が変なときは内蔵デフォルトを使う。"""
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    path = vault / "03_scripts" / "config.json"
    if not path.is_file():
        return cfg
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return cfg
    if isinstance(data.get("timezone"), str) and data["timezone"].strip():
        cfg["timezone"] = data["timezone"].strip()
    est = data.get("estimate")
    if isinstance(est, dict):
        for key in cfg["estimate"]:
            val = est.get(key)
            if isinstance(val, (int, float)) and not isinstance(val, bool) and val > 0:
                cfg["estimate"][key] = val
    return cfg


def _parse_scalar(raw: str) -> Optional[str]:
    v = raw.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in ("'", '"'):
        v = v[1:-1].strip()
    return v if v else None


def parse_simple_frontmatter(text: str) -> Dict[str, Any]:
    """データ契約4章の簡易YAMLサブセット（key: value / インラインリスト [a, b] / 空は None）を読む。"""
    text = text.lstrip("﻿").replace("\r\n", "\n")
    lines = text.split("\n")
    if not lines or lines[0].strip() != "---":
        return {}
    result: Dict[str, Any] = {}
    for line in lines[1:]:
        if line.strip() == "---":
            break
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if ":" not in line:
            continue
        key, _, value = line.partition(":")
        key = key.strip()
        if not key:
            continue
        value = value.strip()
        if value.startswith("[") and value.endswith("]"):
            items = [_parse_scalar(x) for x in value[1:-1].split(",")]
            result[key] = [x for x in items if x]
        else:
            result[key] = _parse_scalar(value)
    return result


class Client:
    def __init__(self, name: str, folder: Path, names: List[str], slack_channels: List[str]):
        self.name = name
        self.folder = folder
        self.names = names
        self.slack_channels = slack_channels


def _as_list(v: Any) -> List[str]:
    if v is None:
        return []
    if isinstance(v, list):
        return [str(x) for x in v if x]
    return [str(v)]


def load_clients(vault: Path) -> Tuple[List[Client], Dict[str, List[Client]], Dict[str, List[Client]]]:
    """01_clients/*/_profile.md を読み、(クライアント一覧, 名前→候補, Slackチャンネル→候補) を返す。"""
    clients: List[Client] = []
    base = vault / "01_clients"
    if not base.is_dir():
        return clients, {}, {}
    for folder in sorted(base.iterdir(), key=lambda p: nfc(p.name)):
        profile = folder / "_profile.md"
        if not folder.is_dir() or folder.name.startswith(".") or not profile.is_file():
            continue
        try:
            fm = parse_simple_frontmatter(profile.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError):
            continue
        cval = fm.get("client")
        name = nfc(cval) if isinstance(cval, str) and cval else nfc(folder.name)
        names = [name] + [nfc(a) for a in _as_list(fm.get("aliases"))]
        channels = [nfc(c).lstrip("#") for c in _as_list(fm.get("slack_channels"))]
        clients.append(Client(name, folder, list(dict.fromkeys(names)), channels))
    name_map: Dict[str, List[Client]] = {}
    chan_map: Dict[str, List[Client]] = {}
    for c in clients:
        for n in c.names:
            name_map.setdefault(n, [])
            if c not in name_map[n]:
                name_map[n].append(c)
        for ch in c.slack_channels:
            chan_map.setdefault(ch, [])
            if c not in chan_map[ch]:
                chan_map[ch].append(c)
    return clients, name_map, chan_map


# ---------------------------------------------------------------------------
# ファイル名・本文の解析
# ---------------------------------------------------------------------------

def parse_inbox_filename(filename: str) -> Tuple[Optional[Dict[str, str]], str, str]:
    """規則「YYYY-MM-DD_<種別>_<相手>_<件名>.md」を解析。(結果, 規則外の理由, 短い理由) を返す。"""
    name = nfc(filename)
    p = PurePosixPath(name)
    ext = p.suffix.lower()
    if ext not in TEXT_EXTS:
        return None, "未対応の形式（.md / .txt / Slack エクスポート以外）", "未対応の形式"
    stem = name[: -len(p.suffix)]
    parts = stem.split("_", 3)
    if len(parts) < 3 or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", parts[0]):
        return None, "ファイル名が規則外（YYYY-MM-DD_<種別>_<相手>_<件名>.md の形ではない）", "ファイル名が規則外"
    date_s, typ, party = parts[0], parts[1], parts[2]
    title = parts[3] if len(parts) == 4 else ""
    try:
        date.fromisoformat(date_s)
    except ValueError:
        return None, f"ファイル名の日付が不正（{date_s}）", "日付が不正"
    if typ not in TYPES:
        return None, f"種別「{typ}」が規則外（{' / '.join(TYPES)} のどれか）", f"種別「{typ}」が規則外"
    if not party.strip():
        return None, "ファイル名に相手がない", "相手がない"
    return {"date": date_s, "type": typ, "party": party, "title": title, "ext": ext}, "", ""


_TS = r"(\d{1,3}):(\d{2})(?::(\d{2}))?(?:[.,](\d{1,3}))?"
# 括弧内のタイムスタンプ: [00:58:10] (58:10) （12:34） [00:00:05.120 --> 00:00:09.000]
_BRACKET_TS_RE = re.compile(r"[\[\(（]\s*" + _TS + r"(?:\s*-->\s*" + _TS + r")?\s*[\]\)）]")
# 行頭のタイムスタンプ: 00:12:34 鈴木: … / 00:00:05,120 --> 00:00:09,000（直後が空白か行末のときだけ）
_LINESTART_TS_RE = re.compile(r"^\s*(?:[-*>]\s+)?" + _TS + r"(?:\s*-->\s*" + _TS + r")?(?=\s|$)")


def _ts_seconds(groups: Tuple[Optional[str], ...]) -> Optional[float]:
    a, b, c, frac = groups
    if a is None or b is None:
        return None
    if c is None:  # MM:SS
        mins, secs = int(a), int(b)
        if secs >= 60:
            return None
        total = mins * 60 + secs
    else:  # HH:MM:SS
        hours, mins, secs = int(a), int(b), int(c)
        if mins >= 60 or secs >= 60:
            return None
        total = hours * 3600 + mins * 60 + secs
    if frac:
        total += int(frac) / (10 ** len(frac))
    return float(total)


def max_timestamp_seconds(text: str) -> Optional[float]:
    """行頭または括弧内のタイムスタンプの最大値（秒）。本文中の「22:00に集合」等は拾わない。"""
    best: Optional[float] = None
    for line in text.split("\n"):
        found: List[Tuple[Optional[str], ...]] = []
        for m in _BRACKET_TS_RE.finditer(line):
            g = m.groups()
            found.append(g[0:4])
            found.append(g[4:8])
        m = _LINESTART_TS_RE.match(line)
        if m:
            g = m.groups()
            found.append(g[0:4])
            found.append(g[4:8])
        for grp in found:
            sec = _ts_seconds(grp)
            if sec is not None and (best is None or sec > best):
                best = sec
    return best


def format_hms(seconds: float) -> str:
    s = int(seconds)
    return f"{s // 3600:02d}:{(s % 3600) // 60:02d}:{s % 60:02d}"


_CHAT_LINE_RE = re.compile(r"^\s*(?:[-*]\s+)?(?:午前|午後)?\d{1,2}:\d{2}(?::\d{2})?(?=\s)")


def count_chat_messages(text: str) -> int:
    """規則名で置かれた slack / line のテキストのメッセージ数を数える（行頭が時刻の行を1件と数える）。"""
    return sum(1 for line in text.split("\n") if _CHAT_LINE_RE.match(line))


def estimate(typ: str, body: str, cfg: Dict[str, Any], message_count: Optional[int] = None) -> Tuple[Optional[int], Optional[int], int, str]:
    """(duration_min, message_count, estimated_minutes, estimate_basis) を返す。"""
    est = cfg["estimate"]
    if typ == "meeting":
        sec = max_timestamp_seconds(body)
        if sec is not None and sec > 0:
            duration = max(1, int(sec / 60 + 0.5))
            return duration, None, duration, f"文字起こしの最終タイムスタンプ {format_hms(sec)}"
        chars = len(re.sub(r"\s", "", body))
        per = est["meeting_chars_per_minute"]
        minutes = max(1, math.ceil(chars / per))
        return None, None, minutes, f"タイムスタンプなし。本文{chars}文字÷{per:g}を切り上げ"
    if typ in ("slack", "line"):
        per = est["chat_minutes_per_message"]
        floor = est["chat_min_minutes"]
        if message_count is None:
            message_count = count_chat_messages(body)
        if message_count <= 0:
            return None, None, int(math.ceil(floor)), f"メッセージ数を判定できず最低値{floor:g}分"
        raw_min = message_count * per
        minutes = int(math.ceil(max(raw_min, floor)))
        basis = f"メッセージ{message_count}件×{per:g}分"
        if raw_min < floor:
            basis += f"（最低{floor:g}分を適用）"
        return None, message_count, minutes, basis
    key = f"{typ}_minutes"
    minutes = int(math.ceil(est.get(key, DEFAULT_CONFIG["estimate"].get(key, 10))))
    return None, None, minutes, f"{typ} は一律{minutes}分"


def build_raw(meta: Dict[str, Any], body: str) -> str:
    lines = ["---"]
    for key in RAW_KEYS:
        val = meta.get(key)
        lines.append(f"{key}: " + ("" if val is None else str(val)))
    lines.append("---")
    return "\n".join(lines) + "\n" + body


def comparable(text: str) -> str:
    """比較用: 先頭 frontmatter の imported_at 行だけを除いた内容。"""
    if not text.startswith("---\n"):
        return text
    end = text.find("\n---\n", 3)
    if end < 0:
        return text
    fm = text[4:end].split("\n")
    fm = [ln for ln in fm if not ln.startswith("imported_at:")]
    return "---\n" + "\n".join(fm) + text[end:]


def decode_text(data: bytes) -> str:
    """UTF-8（BOM付き可）→ だめなら cp932。改行は LF に揃える。"""
    for enc in ("utf-8-sig", "cp932"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("文字コードを判定できない（UTF-8 / Shift_JIS 以外）")
    return text.replace("\r\n", "\n").replace("\r", "\n")


# ---------------------------------------------------------------------------
# Slack エクスポート
# ---------------------------------------------------------------------------

class SlackSource:
    """展開済みフォルダ / zip の差を吸収する読み取り窓口。"""

    def __init__(self, path: Path):
        self.path = path
        self.zip: Optional[zipfile.ZipFile] = None
        self.prefix = ""  # zip 内で channels.json があるフォルダ
        self.root: Optional[Path] = None  # 展開済みフォルダで channels.json があるフォルダ
        if path.is_dir():
            if (path / "channels.json").is_file():
                self.root = path
            else:
                subs = [p for p in path.iterdir() if p.is_dir() and (p / "channels.json").is_file()]
                if len(subs) == 1:
                    self.root = subs[0]
        elif path.suffix.lower() == ".zip" and zipfile.is_zipfile(str(path)):
            self.zip = zipfile.ZipFile(str(path))
            cands = [n for n in self.zip.namelist() if PurePosixPath(n).name == "channels.json"]
            if cands:
                best = min(cands, key=lambda n: n.count("/"))
                self.prefix = best[: -len("channels.json")]
            else:
                self.close()
                self.zip = None

    @property
    def valid(self) -> bool:
        return self.root is not None or self.zip is not None

    def close(self) -> None:
        if self.zip is not None:
            self.zip.close()

    def read_json(self, rel: str) -> Any:
        if self.root is not None:
            p = self.root / rel
            if not p.is_file():
                return None
            return json.loads(p.read_text(encoding="utf-8"))
        assert self.zip is not None
        name = self.prefix + rel
        try:
            data = self.zip.read(name)
        except KeyError:
            return None
        return json.loads(data.decode("utf-8-sig"))

    def channel_days(self) -> Dict[str, List[Tuple[str, str]]]:
        """{チャンネル名(NFC): [(日付, 相対パス), ...]}"""
        out: Dict[str, List[Tuple[str, str]]] = {}
        if self.root is not None:
            for d in sorted(self.root.iterdir()):
                if not d.is_dir() or d.name.startswith("."):
                    continue
                for f in sorted(d.iterdir()):
                    m = SLACK_DAY_RE.match(f.name)
                    if m:
                        out.setdefault(nfc(d.name), []).append((m.group(1), f"{d.name}/{f.name}"))
        else:
            assert self.zip is not None
            for n in sorted(self.zip.namelist()):
                if not n.startswith(self.prefix) or n.endswith("/"):
                    continue
                parts = n[len(self.prefix):].split("/")
                if len(parts) != 2:
                    continue
                m = SLACK_DAY_RE.match(parts[1])
                if m and not parts[0].startswith("."):
                    out.setdefault(nfc(parts[0]), []).append((m.group(1), "/".join(parts)))
        for k in out:
            out[k].sort()
        return out


def slack_user_names(users: Any) -> Dict[str, str]:
    """users.json → {user_id: 表示名}。real_name → display_name → name の順で最初に空でないもの。"""
    out: Dict[str, str] = {}
    if not isinstance(users, list):
        return out
    for u in users:
        if not isinstance(u, dict) or not u.get("id"):
            continue
        prof = u.get("profile") or {}
        for cand in (u.get("real_name"), prof.get("real_name"), prof.get("display_name"), u.get("name")):
            if isinstance(cand, str) and cand.strip():
                out[u["id"]] = cand.strip()
                break
    return out


def slack_speaker(msg: Dict[str, Any], users: Dict[str, str]) -> str:
    uid = msg.get("user")
    if uid and uid in users:
        return users[uid]
    prof = msg.get("user_profile") or {}
    for cand in (prof.get("real_name"), prof.get("display_name"), prof.get("name"), msg.get("username"),
                 (msg.get("bot_profile") or {}).get("name")):
        if isinstance(cand, str) and cand.strip():
            return cand.strip()
    return uid or "不明"


def slack_text(text: str, users: Dict[str, str], channels: Dict[str, str]) -> str:
    def repl(m: "re.Match[str]") -> str:
        inner = m.group(1)
        if inner.startswith("@"):
            uid, _, label = inner[1:].partition("|")
            return "@" + (users.get(uid) or label or uid)
        if inner.startswith("#"):
            cid, _, label = inner[1:].partition("|")
            return "#" + (label or channels.get(cid, cid))
        if inner.startswith("!"):
            key, _, label = inner[1:].partition("|")
            if key.startswith("subteam^"):
                return "@" + (label.lstrip("@") or "グループ")
            return "@" + (label.lstrip("@") or key)
        url, _, label = inner.partition("|")
        return f"{label} ({url})" if label and label != url else url

    text = re.sub(r"<([^<>\n]+)>", repl, text or "")
    return text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")


def is_human_message(msg: Any) -> bool:
    return isinstance(msg, dict) and msg.get("type", "message") == "message" and msg.get("subtype") in SLACK_HUMAN_SUBTYPES and bool(msg.get("ts"))


def _ts_float(ts: Any) -> float:
    try:
        return float(ts)
    except (TypeError, ValueError):
        return 0.0


def render_slack_day(channel: str, day: str, messages: List[Dict[str, Any]], users: Dict[str, str],
                     channels: Dict[str, str], tz: Any, parent_index: Dict[str, str]) -> Tuple[str, int]:
    """1チャンネル1日分を Markdown にする。(本文, メッセージ数) を返す。"""
    msgs = sorted([m for m in messages if is_human_message(m)], key=lambda m: _ts_float(m["ts"]))
    by_ts = {m["ts"]: m for m in msgs}
    top: List[Dict[str, Any]] = []
    replies: Dict[str, List[Dict[str, Any]]] = {}
    for m in msgs:
        tts = m.get("thread_ts")
        if tts and tts != m["ts"] and tts in by_ts:
            replies.setdefault(tts, []).append(m)
        else:
            top.append(m)

    def hhmm(ts: str) -> str:
        return datetime.fromtimestamp(_ts_float(ts), tz).strftime("%H:%M")

    def fmt(m: Dict[str, Any], indent: str, note: str = "") -> List[str]:
        body = slack_text(m.get("text", ""), users, channels)
        files = [f.get("name") or f.get("title") for f in (m.get("files") or []) if isinstance(f, dict)]
        if files:
            body = (body + " " if body else "") + " ".join(f"[添付: {n}]" for n in files if n)
        lines = body.split("\n") if body else [""]
        head = f"{indent}- {hhmm(m['ts'])} **{slack_speaker(m, users)}**: {note}{lines[0]}"
        return [head] + [f"{indent}  {ln}" for ln in lines[1:]]

    out = [f"# Slack #{channel} {day}", ""]
    for m in top:
        tts = m.get("thread_ts")
        note = ""
        if tts and tts != m["ts"]:
            origin = parent_index.get(tts)
            note = f"↳ スレッド返信（元の投稿: {origin}） " if origin else "↳ スレッド返信（元の投稿はこのエクスポートに含まれない） "
        out.extend(fmt(m, "", note))
        for r in replies.get(m["ts"], []):
            out.extend(fmt(r, "  "))
    return "\n".join(out) + "\n", len(msgs)


# ---------------------------------------------------------------------------
# 実行本体
# ---------------------------------------------------------------------------

class Runner:
    def __init__(self, vault: Path, dry_run: bool):
        self.vault = vault
        self.inbox = vault / "00_inbox"
        self.dry_run = dry_run
        self.now = datetime.now().replace(microsecond=0)
        self.today = self.now.date().isoformat()
        self.cfg = load_config(vault)
        self.clients, self.name_map, self.chan_map = load_clients(vault)
        self.events: List[Dict[str, Any]] = []
        self.planned: Dict[str, str] = {}  # dry-run 用: この実行で書く予定のファイル
        self.tz = self._tz()

    def _tz(self) -> Any:
        try:
            from zoneinfo import ZoneInfo
            return ZoneInfo(self.cfg["timezone"])
        except Exception:
            return None  # ローカル時刻

    # --- 補助 ------------------------------------------------------------
    def rel(self, p: Optional[Path]) -> Optional[str]:
        if p is None:
            return None
        try:
            return nfc(str(p.relative_to(self.vault)))
        except ValueError:
            return nfc(str(p))

    def log(self, action: str, src: Any, dest: Any = None, client: Optional[str] = None,
            typ: Optional[str] = None, reason: str = "", short: str = "") -> None:
        """イベントを1件記録。short はサマリー（標準出力・通知）用の短い理由で、jsonl には書かない。"""
        self.events.append({
            "ts": datetime.now().replace(microsecond=0).isoformat(),
            "action": action,
            "src": src if isinstance(src, str) or src is None else self.rel(src),
            "dest": dest if isinstance(dest, str) or dest is None else self.rel(dest),
            "client": client,
            "type": typ,
            "reason": reason,
            "_short": short or reason,
        })

    def exists(self, p: Path) -> bool:
        return str(p) in self.planned or p.exists()

    def read_existing(self, p: Path) -> str:
        if str(p) in self.planned:
            return self.planned[str(p)]
        return p.read_text(encoding="utf-8")

    def write_new(self, p: Path, content: str) -> None:
        """新規作成のみ（既存は絶対に上書きしない）。"""
        if self.dry_run:
            self.planned[str(p)] = content
            return
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "x", encoding="utf-8", newline="\n") as f:
            f.write(content)

    def unique_path(self, directory: Path, name: str, is_dir: bool = False) -> Path:
        cand = directory / name
        if not self.exists(cand):
            return cand
        pp = PurePosixPath(name)
        stem, suffix = (name[: -len(pp.suffix)], pp.suffix) if pp.suffix and not is_dir else (name, "")
        n = 2
        while True:
            cand = directory / f"{stem}_{n}{suffix}"
            if not self.exists(cand):
                return cand
            n += 1

    def archive(self, src: Path) -> Path:
        """元ファイル（フォルダ）を 00_inbox/_done/<実行日>/ に移す。同名があれば連番。"""
        done_dir = self.inbox / DONE_DIR / self.today
        dest = self.unique_path(done_dir, src.name, is_dir=src.is_dir())
        if self.dry_run:
            self.planned[str(dest)] = ""
        else:
            done_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dest))
        return dest

    def place_raw(self, client: Client, raw_name: str, content: str) -> Tuple[str, Path, str]:
        """raw に置く。(action, dest, reason)。同名があれば skipped / conflict。"""
        raw_path = client.folder / "raw" / raw_name
        if not self.exists(raw_path):
            self.write_new(raw_path, content)
            return "moved", raw_path, ""
        existing = self.read_existing(raw_path)
        if comparable(existing) == comparable(content):
            return "skipped", raw_path, "同名の raw が既にあり内容も同一"
        review_dir = self.inbox / REVIEW_DIR
        # 同じ衝突ファイルが既に _要確認 にあれば増やさない（冪等）
        n = 1
        while True:
            cand = review_dir / (raw_name if n == 1 else f"{raw_name[:-3]}_{n}.md")
            if not self.exists(cand):
                self.write_new(cand, content)
                break
            if comparable(self.read_existing(cand)) == comparable(content):
                break
            n += 1
        return "conflict", cand, f"同名の raw（{self.rel(raw_path)}）が既にあり内容が異なる。raw は変更せず新しい方を _要確認 に置いた"

    # --- 入力の種類ごと ---------------------------------------------------
    def resolve_client(self, party: str) -> Tuple[Optional[Client], str, str]:
        name = nfc(party)
        cands = self.name_map.get(name, [])
        if not cands:
            return None, f"相手「{name}」がどの _profile.md の client / aliases とも一致しない", f"未登録の相手「{name}」"
        if len(cands) > 1:
            return (None, f"相手「{name}」が複数のクライアント（{'、'.join(c.name for c in cands)}）に一致する",
                    f"相手「{name}」が複数社に一致")
        return cands[0], "", ""

    def handle_text_file(self, path: Path) -> None:
        info, reason, short = parse_inbox_filename(path.name)
        if info is None:
            self.log("unsorted", path, reason=reason, short=short)
            return
        client, reason, short = self.resolve_client(info["party"])
        if client is None:
            self.log("unsorted", path, typ=info["type"], reason=reason, short=short)
            return
        body = decode_text(path.read_bytes())
        duration, msg_count, est_min, basis = estimate(info["type"], body, self.cfg)
        title = nfc(info["title"])
        raw_name = f"{info['date']}_{info['type']}_{client.name}" + (f"_{title}" if title else "") + ".md"
        meta = {
            "date": info["date"],
            "type": info["type"],
            "client": client.name,
            "title": title,
            "source_file": nfc(path.name),
            "imported_at": self.now.isoformat(),
            "duration_min": duration,
            "message_count": msg_count,
            "estimated_minutes": est_min,
            "estimate_basis": basis,
        }
        action, dest, why = self.place_raw(client, raw_name, build_raw(meta, body))
        self.log(action, path, dest, client.name, info["type"], why)
        self.archive(path)

    def handle_slack(self, path: Path, src: SlackSource) -> bool:
        """Slack エクスポートを処理。元を _done に移してよければ True。"""
        users = slack_user_names(src.read_json("users.json"))
        chans_raw = src.read_json("channels.json") or []
        channels = {c.get("id"): c.get("name") for c in chans_raw if isinstance(c, dict) and c.get("id")}
        days = src.channel_days()
        matched = False
        had_error = False
        for channel, day_files in days.items():
            chan_src = f"{self.rel(path)}/{channel}"
            cands = self.chan_map.get(channel, [])
            if len(cands) != 1:
                if not cands:
                    reason = f"チャンネル #{channel} はどのクライアントの slack_channels にも一致しない"
                    short = "紐づくクライアントなし"
                else:
                    reason = f"チャンネル #{channel} が複数のクライアント（{'、'.join(c.name for c in cands)}）に一致する"
                    short = "複数社に一致"
                self.log("unsorted", chan_src, typ="slack", reason=reason, short=short)
                continue
            client = cands[0]
            matched = True
            # 日をまたぐスレッド返信の「元の投稿」表示用に、チャンネル内の全投稿を索引化
            loaded: List[Tuple[str, str, Optional[List[Dict[str, Any]]]]] = []
            parent_index: Dict[str, str] = {}
            for day, relp in day_files:
                try:
                    data = src.read_json(relp)
                    if not isinstance(data, list):
                        raise ValueError("JSON がメッセージの配列ではない")
                    loaded.append((day, relp, data))
                    for m in data:
                        if is_human_message(m):
                            t = datetime.fromtimestamp(_ts_float(m["ts"]), self.tz).strftime("%m-%d %H:%M")
                            parent_index[m["ts"]] = f"{t} {slack_speaker(m, users)}"
                except Exception as e:  # noqa: BLE001
                    had_error = True
                    self.log("error", f"{self.rel(path)}/{relp}", client=client.name, typ="slack", reason=f"読み込み失敗: {e}")
            for day, relp, data in loaded:
                item_src = f"{self.rel(path)}/{relp}"
                try:
                    body, count = render_slack_day(channel, day, data or [], users, channels, self.tz, parent_index)
                    if count == 0:
                        self.log("skipped", item_src, client=client.name, typ="slack", reason="人の発言がない（システムメッセージのみ）")
                        continue
                    _, _, est_min, basis = estimate("slack", body, self.cfg, message_count=count)
                    raw_name = f"{day}_slack_{client.name}_{channel}.md"
                    meta = {
                        "date": day,
                        "type": "slack",
                        "client": client.name,
                        "title": channel,
                        "source_file": nfc(f"{path.name}/{relp}"),
                        "imported_at": self.now.isoformat(),
                        "duration_min": None,
                        "message_count": count,
                        "estimated_minutes": est_min,
                        "estimate_basis": basis,
                    }
                    action, dest, why = self.place_raw(client, raw_name, build_raw(meta, body))
                    self.log(action, item_src, dest, client.name, "slack", why)
                except Exception as e:  # noqa: BLE001
                    had_error = True
                    self.log("error", item_src, client=client.name, typ="slack", reason=f"{type(e).__name__}: {e}")
        if not matched and not had_error:
            self.log("unsorted", path, typ="slack", reason="Slack エクスポートにクライアントに紐づくチャンネルがない（inbox に残した）",
                     short="紐づくチャンネルなし")
        return matched and not had_error

    def handle_entry(self, path: Path) -> None:
        name = nfc(path.name)
        if path.is_dir() or path.suffix.lower() == ".zip":
            src = SlackSource(path)
            if not src.valid:
                what = "フォルダ" if path.is_dir() else "zip"
                self.log("unsorted", path, reason=f"未対応の{what}（channels.json を含む Slack エクスポートではない）",
                         short=f"未対応の{what}")
                return
            try:
                ok = self.handle_slack(path, src)
            finally:
                src.close()
            if ok:
                self.archive(path)
            return
        if PurePosixPath(name).suffix.lower() in TEXT_EXTS:
            self.handle_text_file(path)
            return
        self.log("unsorted", path, reason="未対応の形式（.md / .txt / Slack エクスポート以外）", short="未対応の形式")

    def run(self) -> int:
        if not self.inbox.is_dir():
            print(f"00_inbox が見つかりません: {self.inbox}", file=sys.stderr)
            return 1
        entries = sorted(
            (p for p in self.inbox.iterdir() if not p.name.startswith(".") and nfc(p.name) not in SKIP_NAMES),
            key=lambda p: nfc(p.name),
        )
        for p in entries:
            try:
                self.handle_entry(p)
            except Exception as e:  # noqa: BLE001  1件の失敗で全体を止めない
                self.log("error", p, reason=f"{type(e).__name__}: {e}")
        return 1 if any(e["action"] == "error" for e in self.events) else 0

    # --- 結果 --------------------------------------------------------------
    def summary(self) -> Dict[str, Any]:
        counts = {a: 0 for a in ("moved", "skipped", "conflict", "unsorted", "error")}
        by_client: Dict[str, Dict[str, int]] = {}
        unsorted_items = []
        for e in self.events:
            counts[e["action"]] = counts.get(e["action"], 0) + 1
            if e["action"] == "moved":
                by_client.setdefault(e["client"], {})
                by_client[e["client"]][e["type"]] = by_client[e["client"]].get(e["type"], 0) + 1
            if e["action"] == "unsorted":
                unsorted_items.append({"src": e["src"], "reason": e["reason"], "_short": e["_short"]})
        message = self.message(counts, by_client, unsorted_items)
        for u in unsorted_items:
            u.pop("_short", None)
        return {"dry_run": self.dry_run, **counts, "by_client": by_client, "unsorted_items": unsorted_items,
                "message": message}

    def message(self, counts: Dict[str, int], by_client: Dict[str, Dict[str, int]], unsorted_items: List[Dict[str, Any]]) -> str:
        parts = []
        if by_client:
            cl = []
            for cname in sorted(by_client):
                t = by_client[cname]
                cl.append(cname + " " + "・".join(f"{TYPE_LABELS.get(k, k)}{t[k]}件" for k in TYPES if k in t))
            parts.append("振り分け: " + " / ".join(cl))
        else:
            parts.append("振り分け: なし")
        if counts["skipped"]:
            parts.append(f"既存と同一でスキップ{counts['skipped']}件")
        if counts["conflict"]:
            parts.append(f"要確認{counts['conflict']}件（00_inbox/_要確認/ を見てください）")
        if unsorted_items:
            reasons = []
            for u in unsorted_items:
                label = u["src"] or ""
                if label.startswith("00_inbox/"):
                    label = label[len("00_inbox/"):]
                reasons.append(f"{label}: {u['_short']}")
            parts.append(f"未振り分け{len(unsorted_items)}件（" + " / ".join(reasons) + "）")
        if counts["error"]:
            parts.append(f"エラー{counts['error']}件（04_logs/inbox_sort.jsonl を確認）")
        msg = " ｜ ".join(parts)
        return ("[dry-run] " + msg) if self.dry_run else msg

    def write_logs(self, summary: Dict[str, Any]) -> None:
        if self.dry_run:
            return
        logs = self.vault / "04_logs"
        logs.mkdir(parents=True, exist_ok=True)
        with open(logs / "inbox_sort.jsonl", "a", encoding="utf-8", newline="\n") as f:
            for e in self.events:
                row = {k: v for k, v in e.items() if not k.startswith("_")}
                f.write(json.dumps(row, ensure_ascii=False) + "\n")
        with open(logs / "runs.jsonl", "a", encoding="utf-8", newline="\n") as f:
            f.write(json.dumps({"ts": self.now.isoformat(), "script": SCRIPT_NAME, "summary": summary}, ensure_ascii=False) + "\n")


def notify(message: str) -> None:
    """macOS 通知。osascript が無い環境では何もしない。"""
    if not shutil.which("osascript"):
        return
    text = message.replace("\\", "\\\\").replace('"', '\\"')
    if len(text) > 240:
        text = text[:237] + "…"
    try:
        subprocess.run(["osascript", "-e", f'display notification "{text}" with title "inbox 振り分け"'],
                       check=False, timeout=10, capture_output=True)
    except Exception:  # noqa: BLE001
        pass


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        prog="inbox_sort.py",
        description="00_inbox の文字起こし・メモ・Slack エクスポートを整形し、01_clients/<クライアント>/raw/ に振り分ける。",
    )
    ap.add_argument("--dry-run", action="store_true", help="何が起きるかだけ表示。ファイル移動・書き込み・ログ追記を一切しない")
    ap.add_argument("--notify", action="store_true", help="結果のサマリーを macOS 通知でも出す")
    ap.add_argument("--vault", type=Path, default=None, help="vault ルートのパス（既定: このスクリプトがある 03_scripts の1つ上）")
    args = ap.parse_args(argv)

    vault = (args.vault or Path(__file__).resolve().parent.parent).resolve()
    runner = Runner(vault, args.dry_run)
    code = runner.run()
    summary = runner.summary()
    runner.write_logs(summary)

    for e in runner.events:
        arrow = f" → {e['dest']}" if e.get("dest") else ""
        why = f"（{e['reason']}）" if e.get("reason") else ""
        print(f"  [{e['action']}] {e['src']}{arrow}{why}")
    print(summary["message"])
    if args.notify:
        notify(summary["message"])
    return code


if __name__ == "__main__":
    sys.exit(main())
