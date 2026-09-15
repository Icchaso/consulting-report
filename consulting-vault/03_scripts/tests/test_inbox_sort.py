"""inbox_sort.py のテスト（標準ライブラリ unittest のみ）。

実行: vault ルートで
    /usr/bin/python3 -m unittest discover -s 03_scripts/tests -v

本物の 00_inbox / 01_clients には一切触らない。毎回、一時ディレクトリに小さな vault を組み立て、
99_samples/inbox をコピーして inbox_sort.py を実行し、結果を確かめる。
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import unittest
import zipfile
from datetime import date
from pathlib import Path
from typing import Dict, List

sys.dont_write_bytecode = True

TESTS_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = TESTS_DIR.parent
REAL_VAULT = SCRIPTS_DIR.parent
SCRIPT = SCRIPTS_DIR / "inbox_sort.py"
CONFIG = SCRIPTS_DIR / "config.json"
SAMPLES = REAL_VAULT / "99_samples" / "inbox"

sys.path.insert(0, str(SCRIPTS_DIR))
import inbox_sort  # noqa: E402

SAMPLE = "サンプル商事"
KOUMUTEN = "テスト工務店"

PROFILE_SAMPLE = """---
client: サンプル商事
aliases: [サンプル商事, サンプル]
slack_channels: [sample-shoji]
email_domains: [sample-shoji.example]
line_names: [サンプル商事 佐藤]
contract: 月額顧問（月2回定例＋チャット相談）
monthly_fee:
contact: 佐藤 様（営業部長）
style: です・ます調。簡潔に。結論から。
---
目標KPI: 在庫回転日数（テスト用の最小プロフィール）
"""

PROFILE_KOUMUTEN = """---
client: テスト工務店
aliases: [テスト工務店, テスト]
slack_channels: [test-koumuten]
email_domains: [test-koumuten.example]
monthly_fee:
---
"""

LOG_KEYS = {"ts", "action", "src", "dest", "client", "type", "reason"}


def make_vault(root: Path, samples: bool = True, config: bool = True) -> Path:
    vault = root / "vault"
    for name, text in ((SAMPLE, PROFILE_SAMPLE), (KOUMUTEN, PROFILE_KOUMUTEN)):
        d = vault / "01_clients" / name
        d.mkdir(parents=True)
        (d / "_profile.md").write_text(text, encoding="utf-8")
    (vault / "03_scripts").mkdir(parents=True)
    if config:
        shutil.copy2(str(CONFIG), str(vault / "03_scripts" / "config.json"))
    if samples:
        shutil.copytree(str(SAMPLES), str(vault / "00_inbox"))
    else:
        (vault / "00_inbox").mkdir(parents=True)
    return vault


def run(vault: Path, *args: str) -> subprocess.CompletedProcess:
    env = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
    return subprocess.run([sys.executable, str(SCRIPT), "--vault", str(vault), *args],
                          capture_output=True, text=True, env=env)


def raw_dir(vault: Path, client: str) -> Path:
    return vault / "01_clients" / client / "raw"


def frontmatter(path: Path) -> Dict[str, object]:
    return inbox_sort.parse_simple_frontmatter(path.read_text(encoding="utf-8"))


def body_of(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    end = text.index("\n---\n", 3)
    return text[end + len("\n---\n"):]


def read_jsonl(path: Path) -> List[dict]:
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def snapshot(root: Path) -> Dict[str, str]:
    out = {}
    for p in sorted(root.rglob("*")):
        rel = unicodedata.normalize("NFC", str(p.relative_to(root)))
        out[rel] = hashlib.sha256(p.read_bytes()).hexdigest() if p.is_file() else "<dir>"
    return out


def today() -> str:
    return date.today().isoformat()


class SampleSortingTest(unittest.TestCase):
    """99_samples/inbox を一度振り分けて、行き先と frontmatter を確かめる。"""

    @classmethod
    def setUpClass(cls) -> None:
        cls.tmp = tempfile.TemporaryDirectory()
        cls.vault = make_vault(Path(cls.tmp.name))
        cls.result = run(cls.vault)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.tmp.cleanup()

    def test_exit_code(self) -> None:
        self.assertEqual(self.result.returncode, 0, self.result.stderr)

    def test_raw_files(self) -> None:
        self.assertEqual(sorted(p.name for p in raw_dir(self.vault, SAMPLE).iterdir()), [
            "2026-09-03_meeting_サンプル商事_月次定例.md",
            "2026-09-08_slack_サンプル商事_sample-shoji.md",
            "2026-09-12_slack_サンプル商事_sample-shoji.md",
            "2026-09-15_memo_サンプル商事_電話.md",
            "2026-09-17_meeting_サンプル商事_月次定例2回目.md",
        ])
        self.assertEqual([p.name for p in raw_dir(self.vault, KOUMUTEN).iterdir()],
                         ["2026-09-10_meeting_テスト工務店_集客相談.md"])

    def test_meeting_macwhisper(self) -> None:
        path = raw_dir(self.vault, SAMPLE) / "2026-09-03_meeting_サンプル商事_月次定例.md"
        fm = frontmatter(path)
        self.assertEqual(fm["date"], "2026-09-03")
        self.assertEqual(fm["type"], "meeting")
        self.assertEqual(fm["client"], SAMPLE)
        self.assertEqual(fm["title"], "月次定例")
        self.assertEqual(fm["source_file"], "2026-09-03_meeting_サンプル商事_月次定例.md")
        self.assertEqual(fm["duration_min"], "50")  # 最終タイムスタンプ 00:50:12.300
        self.assertIsNone(fm["message_count"])
        self.assertEqual(fm["estimated_minutes"], "50")
        self.assertIn("00:50:12", fm["estimate_basis"])
        self.assertRegex(fm["imported_at"], r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$")
        # 本文は元ファイルのまま
        original = (SAMPLES / "2026-09-03_meeting_サンプル商事_月次定例.md").read_text(encoding="utf-8")
        self.assertEqual(body_of(path), original)

    def test_frontmatter_key_order(self) -> None:
        path = raw_dir(self.vault, SAMPLE) / "2026-09-15_memo_サンプル商事_電話.md"
        lines = path.read_text(encoding="utf-8").split("\n")
        self.assertEqual(lines[0], "---")
        keys = [ln.split(":", 1)[0] for ln in lines[1:11]]
        self.assertEqual(keys, list(inbox_sort.RAW_KEYS))
        self.assertEqual(lines[11], "---")

    def test_alias_uses_official_name(self) -> None:
        path = raw_dir(self.vault, SAMPLE) / "2026-09-17_meeting_サンプル商事_月次定例2回目.md"
        fm = frontmatter(path)
        self.assertEqual(fm["client"], SAMPLE)
        self.assertEqual(fm["title"], "月次定例2回目")
        self.assertEqual(fm["source_file"], "2026-09-17_meeting_サンプル_月次定例2回目.md")
        self.assertEqual(fm["duration_min"], "28")  # (00:28:10)
        self.assertEqual(fm["estimated_minutes"], "28")

    def test_txt_meeting(self) -> None:
        path = raw_dir(self.vault, KOUMUTEN) / "2026-09-10_meeting_テスト工務店_集客相談.md"
        fm = frontmatter(path)
        self.assertEqual(fm["type"], "meeting")
        self.assertEqual(fm["client"], KOUMUTEN)
        self.assertEqual(fm["title"], "集客相談")
        self.assertEqual(fm["source_file"], "2026-09-10_meeting_テスト工務店_集客相談.txt")
        self.assertEqual(fm["duration_min"], "31")  # 00:31:20。本文中の「22:00過ぎ」は拾わない
        self.assertIn("22:00過ぎ", body_of(path))

    def test_memo(self) -> None:
        fm = frontmatter(raw_dir(self.vault, SAMPLE) / "2026-09-15_memo_サンプル商事_電話.md")
        self.assertEqual(fm["type"], "memo")
        self.assertEqual(fm["title"], "電話")
        self.assertIsNone(fm["duration_min"])
        self.assertIsNone(fm["message_count"])
        self.assertEqual(fm["estimated_minutes"], "5")

    def test_slack_frontmatter(self) -> None:
        d = raw_dir(self.vault, SAMPLE)
        fm1 = frontmatter(d / "2026-09-08_slack_サンプル商事_sample-shoji.md")
        fm2 = frontmatter(d / "2026-09-12_slack_サンプル商事_sample-shoji.md")
        for fm, day in ((fm1, "2026-09-08"), (fm2, "2026-09-12")):
            self.assertEqual(fm["date"], day)
            self.assertEqual(fm["type"], "slack")
            self.assertEqual(fm["client"], SAMPLE)
            self.assertEqual(fm["title"], "sample-shoji")
            self.assertIsNone(fm["duration_min"])
        self.assertEqual(fm1["message_count"], "4")  # channel_join を除いた件数
        self.assertEqual(fm1["estimated_minutes"], "10")  # 4件×2分=8分 → 最低10分
        self.assertEqual(fm2["message_count"], "6")
        self.assertEqual(fm2["estimated_minutes"], "12")
        self.assertEqual(fm1["source_file"], "slack_export/sample-shoji/2026-09-08.json")

    def test_slack_body(self) -> None:
        d = raw_dir(self.vault, SAMPLE)
        b1 = body_of(d / "2026-09-08_slack_サンプル商事_sample-shoji.md")
        b2 = body_of(d / "2026-09-12_slack_サンプル商事_sample-shoji.md")
        # channel_join は除外
        self.assertNotIn("has joined", b1)
        # メンションは名前に置換（real_name → display_name → name の順で解決）
        self.assertNotIn("<@", b1 + b2)
        self.assertIn("- 10:15 **佐藤 部長**: @山田 さん、在庫の持ち方で相談です。", b1)
        self.assertIn("@佐藤 部長 購買から", b2)
        self.assertIn("**tanaka-sample**", b1)
        # スレッド返信は親の下にインデント（複数行の本文も字下げ）
        self.assertIn("\n  - 10:40 **山田**: ご相談ありがとうございます。", b1)
        self.assertIn("\n    まずは定番色だけ店舗に寄せて", b1)
        self.assertIn("\n  - 11:05 **佐藤 部長**: なるほど", b1)
        self.assertIn("\n  - 09:45 **山田**: はい、3か月分で大丈夫です。", b2)
        # 日をまたぐスレッド返信には元の投稿を注記
        self.assertIn("↳ スレッド返信（元の投稿: 09-08 10:15 佐藤 部長）", b2)
        # HTML エスケープの復元と添付ファイル名
        self.assertIn("<品番> ごとに", b1)
        self.assertIn("[添付: 棚番号一覧.xlsx]", b1)

    def test_random_channel_unsorted(self) -> None:
        events = read_jsonl(self.vault / "04_logs" / "inbox_sort.jsonl")
        rnd = [e for e in events if (e["src"] or "").endswith("slack_export/random")]
        self.assertEqual(len(rnd), 1)
        self.assertEqual(rnd[0]["action"], "unsorted")
        self.assertIn("random", rnd[0]["reason"])
        for client in (SAMPLE, KOUMUTEN):
            self.assertFalse(any("random" in p.name for p in raw_dir(self.vault, client).iterdir()))

    def test_unsorted_files_stay_in_inbox(self) -> None:
        inbox = self.vault / "00_inbox"
        left = sorted(unicodedata.normalize("NFC", p.name) for p in inbox.iterdir())
        self.assertEqual(left, ["2026-09-22_meeting_未登録社_打合せ.md", "_done", "録音 2026-09-20 14.30.md"])
        events = read_jsonl(self.vault / "04_logs" / "inbox_sort.jsonl")
        reasons = {e["src"]: e["reason"] for e in events if e["action"] == "unsorted"}
        self.assertIn("未登録社", reasons["00_inbox/2026-09-22_meeting_未登録社_打合せ.md"])
        self.assertIn("規則外", reasons["00_inbox/録音 2026-09-20 14.30.md"])

    def test_done_folder(self) -> None:
        done = self.vault / "00_inbox" / "_done" / today()
        self.assertEqual(sorted(unicodedata.normalize("NFC", p.name) for p in done.iterdir()), [
            "2026-09-03_meeting_サンプル商事_月次定例.md",
            "2026-09-10_meeting_テスト工務店_集客相談.txt",
            "2026-09-15_memo_サンプル商事_電話.md",
            "2026-09-17_meeting_サンプル_月次定例2回目.md",
            "slack_export",
        ])
        self.assertTrue((done / "slack_export" / "channels.json").is_file())

    def test_logs(self) -> None:
        events = read_jsonl(self.vault / "04_logs" / "inbox_sort.jsonl")
        self.assertEqual(len(events), 9)
        for e in events:
            self.assertEqual(set(e), LOG_KEYS)
            self.assertIn(e["action"], {"moved", "skipped", "conflict", "unsorted", "error"})
        self.assertEqual(sum(e["action"] == "moved" for e in events), 6)
        self.assertEqual(sum(e["action"] == "unsorted" for e in events), 3)
        runs = read_jsonl(self.vault / "04_logs" / "runs.jsonl")
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0]["script"], "inbox_sort")
        self.assertEqual(runs[0]["summary"]["moved"], 6)
        self.assertEqual(runs[0]["summary"]["by_client"][SAMPLE], {"meeting": 2, "memo": 1, "slack": 2})

    def test_stdout_summary(self) -> None:
        last = self.result.stdout.strip().splitlines()[-1]
        self.assertTrue(last.startswith("振り分け: サンプル商事 会議2件・Slack2件・メモ1件 / テスト工務店 会議1件"), last)
        self.assertIn("未振り分け3件", last)
        self.assertIn("slack_export/random: 紐づくクライアントなし", last)


class RerunTest(unittest.TestCase):
    """冪等性・raw を上書きしないこと。"""

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.vault = make_vault(Path(self.tmp.name))
        first = run(self.vault)
        self.assertEqual(first.returncode, 0, first.stderr)
        self.raw_before = snapshot(self.vault / "01_clients")

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def last_run_events(self, n: int) -> List[dict]:
        return read_jsonl(self.vault / "04_logs" / "inbox_sort.jsonl")[-n:]

    def test_rerun_without_new_input(self) -> None:
        r = run(self.vault)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(snapshot(self.vault / "01_clients"), self.raw_before)
        self.assertEqual(len(read_jsonl(self.vault / "04_logs" / "runs.jsonl")), 2)

    def test_same_file_again_is_skipped(self) -> None:
        name = "2026-09-03_meeting_サンプル商事_月次定例.md"
        shutil.copy2(str(SAMPLES / name), str(self.vault / "00_inbox" / name))
        shutil.copytree(str(SAMPLES / "slack_export"), str(self.vault / "00_inbox" / "slack_export"))
        r = run(self.vault)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(snapshot(self.vault / "01_clients"), self.raw_before)
        events = read_jsonl(self.vault / "04_logs" / "inbox_sort.jsonl")
        skipped = [e for e in events if e["action"] == "skipped"]
        self.assertEqual(len(skipped), 3)  # 会議1 + Slack 2日分
        done = self.vault / "00_inbox" / "_done" / today()
        self.assertTrue((done / "2026-09-03_meeting_サンプル商事_月次定例_2.md").is_file())
        self.assertTrue((done / "slack_export_2").is_dir())
        self.assertIn("既存と同一でスキップ3件", r.stdout)

    def test_changed_file_goes_to_review(self) -> None:
        name = "2026-09-03_meeting_サンプル商事_月次定例.md"
        changed = (SAMPLES / name).read_text(encoding="utf-8") + "[00:51:00.000 --> 00:51:05.000] 佐藤: 追記です。\n"
        (self.vault / "00_inbox" / name).write_text(changed, encoding="utf-8")
        r = run(self.vault)
        self.assertEqual(r.returncode, 0, r.stderr)
        # raw は変わらない
        self.assertEqual(snapshot(self.vault / "01_clients"), self.raw_before)
        review = self.vault / "00_inbox" / "_要確認" / name
        self.assertTrue(review.is_file())
        self.assertIn("追記です。", review.read_text(encoding="utf-8"))
        ev = self.last_run_events(3)
        conflicts = [e for e in ev if e["action"] == "conflict"]
        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts[0]["dest"], f"00_inbox/_要確認/{name}")
        self.assertIn("要確認1件", r.stdout)
        # 同じ差分ファイルをもう一度入れても _要確認 は増えない
        (self.vault / "00_inbox" / name).write_text(changed, encoding="utf-8")
        run(self.vault)
        self.assertEqual([p.name for p in (self.vault / "00_inbox" / "_要確認").iterdir()], [name])
        self.assertEqual(snapshot(self.vault / "01_clients"), self.raw_before)


class DryRunTest(unittest.TestCase):
    def test_dry_run_changes_nothing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp))
            before = snapshot(vault)
            r = run(vault, "--dry-run")
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertEqual(snapshot(vault), before)
            self.assertFalse((vault / "04_logs").exists())
            last = r.stdout.strip().splitlines()[-1]
            self.assertTrue(last.startswith("[dry-run] 振り分け: サンプル商事 会議2件・Slack2件・メモ1件"), last)


class NfdFilenameTest(unittest.TestCase):
    def test_nfd_name_is_sorted_and_saved_as_nfc(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp), samples=False)
            nfc_name = "2026-09-25_memo_サンプル商事_プレゼン準備.md"
            nfd_name = unicodedata.normalize("NFD", nfc_name)
            self.assertNotEqual(nfc_name, nfd_name)
            (vault / "00_inbox" / nfd_name).write_text("ブラウザで資料確認。\n", encoding="utf-8")
            r = run(vault)
            self.assertEqual(r.returncode, 0, r.stderr)
            names = os.listdir(str(raw_dir(vault, SAMPLE)))
            self.assertEqual(len(names), 1)
            self.assertEqual(names[0], nfc_name)  # NFC で保存されている
            fm = frontmatter(raw_dir(vault, SAMPLE) / names[0])
            self.assertEqual(fm["client"], SAMPLE)
            self.assertEqual(fm["title"], "プレゼン準備")
            self.assertEqual(fm["source_file"], nfc_name)


class ZipExportTest(unittest.TestCase):
    def test_zip_export(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp), samples=False)
            zpath = vault / "00_inbox" / "slack_export.zip"
            with zipfile.ZipFile(str(zpath), "w") as z:
                for p in sorted((SAMPLES / "slack_export").rglob("*.json")):
                    z.write(str(p), "slack_export/" + str(p.relative_to(SAMPLES / "slack_export")))
            r = run(vault)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertEqual(sorted(p.name for p in raw_dir(vault, SAMPLE).iterdir()), [
                "2026-09-08_slack_サンプル商事_sample-shoji.md",
                "2026-09-12_slack_サンプル商事_sample-shoji.md",
            ])
            self.assertFalse(zpath.exists())
            self.assertTrue((vault / "00_inbox" / "_done" / today() / "slack_export.zip").is_file())
            fm = frontmatter(raw_dir(vault, SAMPLE) / "2026-09-08_slack_サンプル商事_sample-shoji.md")
            self.assertEqual(fm["source_file"], "slack_export.zip/sample-shoji/2026-09-08.json")


class RobustnessTest(unittest.TestCase):
    def test_broken_slack_day_does_not_stop_others(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp))
            (vault / "00_inbox" / "slack_export" / "sample-shoji" / "2026-09-13.json").write_text("{壊れた", encoding="utf-8")
            r = run(vault)
            self.assertEqual(r.returncode, 1)  # エラーがあれば終了コード1
            events = read_jsonl(vault / "04_logs" / "inbox_sort.jsonl")
            self.assertEqual(sum(e["action"] == "error" for e in events), 1)
            # 他のファイルは振り分けられている
            self.assertEqual(len(list(raw_dir(vault, SAMPLE).iterdir())), 5)
            self.assertEqual(len(list(raw_dir(vault, KOUMUTEN).iterdir())), 1)
            # エラーを含むエクスポートは inbox に残す（直して再実行すれば済んだ分は skipped になる）
            self.assertTrue((vault / "00_inbox" / "slack_export").is_dir())
            self.assertIn("エラー1件", r.stdout)

    def test_ignored_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp), samples=False)
            inbox = vault / "00_inbox"
            (inbox / ".DS_Store").write_bytes(b"\x00\x01")
            (inbox / ".gitkeep").write_text("", encoding="utf-8")
            (inbox / "_要確認").mkdir()
            (inbox / "_要確認" / "2026-09-01_memo_サンプル商事_保留.md").write_text("保留\n", encoding="utf-8")
            (inbox / "_done").mkdir()
            (inbox / "_done" / "2026-09-02_memo_サンプル商事_済み.md").write_text("済み\n", encoding="utf-8")
            before = snapshot(inbox)
            r = run(vault)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertEqual(snapshot(inbox), before)
            self.assertEqual(read_jsonl(vault / "04_logs" / "inbox_sort.jsonl"), [])
            self.assertFalse(raw_dir(vault, SAMPLE).exists())

    def test_unsupported_entries_stay(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp), samples=False)
            inbox = vault / "00_inbox"
            (inbox / "2026-09-05_meeting_サンプル商事_録音.m4a").write_bytes(b"fake")
            (inbox / "2026-09-05_call_サンプル商事_電話.md").write_text("x\n", encoding="utf-8")
            (inbox / "2026-13-05_memo_サンプル商事_日付不正.md").write_text("x\n", encoding="utf-8")
            (inbox / "資料フォルダ").mkdir()
            r = run(vault)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertEqual(len(list(inbox.iterdir())), 4)
            events = read_jsonl(vault / "04_logs" / "inbox_sort.jsonl")
            self.assertEqual([e["action"] for e in events], ["unsorted"] * 4)

    def test_config_coefficients(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp), config=False)
            cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
            cfg["estimate"]["chat_minutes_per_message"] = 5
            cfg["estimate"]["memo_minutes"] = 7
            (vault / "03_scripts" / "config.json").write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
            run(vault)
            d = raw_dir(vault, SAMPLE)
            self.assertEqual(frontmatter(d / "2026-09-08_slack_サンプル商事_sample-shoji.md")["estimated_minutes"], "20")
            self.assertEqual(frontmatter(d / "2026-09-15_memo_サンプル商事_電話.md")["estimated_minutes"], "7")

    def test_missing_config_uses_defaults(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            vault = make_vault(Path(tmp), config=False)
            run(vault)
            fm = frontmatter(raw_dir(vault, SAMPLE) / "2026-09-12_slack_サンプル商事_sample-shoji.md")
            self.assertEqual(fm["estimated_minutes"], "12")


class UnitTest(unittest.TestCase):
    """関数単位のテスト。"""

    def test_times_in_sentences_are_ignored(self) -> None:
        text = "\n".join([
            "[00:01] 山田: 始めます",
            "[00:03:10] 佐藤: 懇親会は22:00に集合で",
            "22:00に集合です（行頭でも直後が文字なら拾わない）",
            "鈴木: 18:30から 23:59 まで",
            "夜間バッチは毎日 22:00 に回ります",
        ])
        self.assertEqual(inbox_sort.max_timestamp_seconds(text), 190.0)

    def test_timestamp_formats(self) -> None:
        f = inbox_sort.max_timestamp_seconds
        self.assertEqual(f("[00:00:05.120 --> 00:00:09.000] 山田: こんにちは"), 9.0)
        self.assertEqual(f("00:12:34 鈴木: はい"), 754.0)
        self.assertEqual(f("**山田** (12:34): はい"), 754.0)
        self.assertEqual(f("**山田** (00:12:34): はい"), 754.0)
        self.assertEqual(f("[12:34] はい"), 754.0)
        self.assertEqual(f("[00:58:10]"), 3490.0)
        self.assertEqual(f("（58:10） 全角括弧"), 3490.0)
        self.assertEqual(f("1\n00:00:05,120 --> 00:59:09,000\nSRT"), 3549.0)
        self.assertIsNone(f("タイムスタンプのない文章。22:00に集合。"))

    def test_estimate(self) -> None:
        cfg = inbox_sort.load_config(Path("/nonexistent"))
        self.assertEqual(inbox_sort.estimate("meeting", "[00:58:10] 終わり", cfg),
                         (58, None, 58, "文字起こしの最終タイムスタンプ 00:58:10"))
        dur, cnt, est, basis = inbox_sort.estimate("meeting", "あ" * 601, cfg)
        self.assertEqual((dur, cnt, est), (None, None, 3))
        self.assertIn("601文字", basis)
        self.assertEqual(inbox_sort.estimate("slack", "", cfg, message_count=3)[2], 10)
        self.assertEqual(inbox_sort.estimate("line", "", cfg, message_count=12)[1:3], (12, 24))
        self.assertEqual(inbox_sort.estimate("line", "12:34\t佐藤\tこんにちは\n12:35\t山田\tどうも\n", cfg)[1], 2)
        self.assertEqual(inbox_sort.estimate("mail", "x", cfg)[2], 10)
        self.assertEqual(inbox_sort.estimate("doc", "x", cfg)[2], 30)
        self.assertEqual(inbox_sort.estimate("memo", "x", cfg)[2], 5)

    def test_parse_filename(self) -> None:
        p = inbox_sort.parse_inbox_filename
        info, _, _ = p("2026-09-14_meeting_サンプル商事_月次_定例.md")
        self.assertEqual((info["date"], info["type"], info["party"], info["title"]),
                         ("2026-09-14", "meeting", "サンプル商事", "月次_定例"))
        info, _, _ = p("2026-09-14_memo_サンプル商事.md")
        self.assertEqual((info["party"], info["title"]), ("サンプル商事", ""))
        self.assertEqual(p("2026-09-14_mail_テスト_見積.TXT")[0]["type"], "mail")
        self.assertIsNone(p("2026-09-14_call_サンプル商事_x.md")[0])
        self.assertIsNone(p("2026-02-30_memo_サンプル商事_x.md")[0])
        self.assertIsNone(p("録音 2026-09-20 14.30.md")[0])
        self.assertIsNone(p("2026-09-14_memo_サンプル商事_x.pdf")[0])

    def test_profile_parser(self) -> None:
        fm = inbox_sort.parse_simple_frontmatter(PROFILE_SAMPLE)
        self.assertEqual(fm["client"], SAMPLE)
        self.assertEqual(fm["aliases"], [SAMPLE, "サンプル"])
        self.assertEqual(fm["slack_channels"], ["sample-shoji"])
        self.assertIsNone(fm["monthly_fee"])
        self.assertEqual(fm["style"], "です・ます調。簡潔に。結論から。")

    def test_comparable_ignores_imported_at_only(self) -> None:
        a = "---\ndate: 2026-09-01\nimported_at: 2026-09-14T22:00:00\n---\n本文\n"
        b = "---\ndate: 2026-09-01\nimported_at: 2026-09-15T22:00:00\n---\n本文\n"
        c = "---\ndate: 2026-09-01\nimported_at: 2026-09-15T22:00:00\n---\n本文2\n"
        self.assertEqual(inbox_sort.comparable(a), inbox_sort.comparable(b))
        self.assertNotEqual(inbox_sort.comparable(a), inbox_sort.comparable(c))


if __name__ == "__main__":
    unittest.main()
