"""worklog.py / check_sources.py / send_log.py のテスト。

一時ディレクトリに vault を組み立てて、各スクリプトを --vault 付きで実行して確かめる。
実際の vault の 01_clients/ には何も書かない。

実行（vault ルートで）:
  /usr/bin/python3 -m unittest 03_scripts/tests/test_worklog_check_send.py -v
"""
from __future__ import annotations

import datetime as dt
import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent
WORKLOG = SCRIPTS / "worklog.py"
CHECK = SCRIPTS / "check_sources.py"
SEND = SCRIPTS / "send_log.py"

CLIENT_A = "サンプル商事"   # 「プ」を含むので NFD と NFC で見た目が同じでもバイト列が違う
CLIENT_B = "テスト工務店"


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, str(path))
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


def raw_text(date: str, rtype: str, client: str, title: str, fname: str,
             minutes, basis: str, body: str = "本文") -> str:
    m = "" if minutes is None else str(minutes)
    return (
        "---\n"
        f"date: {date}\n"
        f"type: {rtype}\n"
        f"client: {client}\n"
        f"title: {title}\n"
        f"source_file: {fname}\n"
        "imported_at: 2026-09-30T22:00:00\n"
        "duration_min: \n"
        "message_count: \n"
        f"estimated_minutes: {m}\n"
        f"estimate_basis: {basis}\n"
        "---\n"
        f"{body}\n"
    )


class VaultTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="vault_test_"))
        self.vault = self.tmp / "consulting-vault"
        for d in ("00_inbox", "02_templates", "03_scripts", "04_logs"):
            (self.vault / d).mkdir(parents=True)
        # クライアントAのフォルダ名は macOS が返しうる NFD 形で作る
        self.dir_a = self.vault / "01_clients" / unicodedata.normalize("NFD", CLIENT_A)
        self.dir_b = self.vault / "01_clients" / CLIENT_B
        for d in (self.dir_a, self.dir_b):
            (d / "raw").mkdir(parents=True)
            (d / "_profile.md").write_text("---\nclient: x\n---\n", encoding="utf-8")

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def add_raw(self, client_dir: Path, date: str, rtype: str, title: str, minutes,
                basis: str = "", client: str = CLIENT_A, nfd_name: bool = False) -> str:
        fname = f"{date}_{rtype}_{client}_{title or '無題'}.md"
        on_disk = unicodedata.normalize("NFD", fname) if nfd_name else fname
        (client_dir / "raw" / on_disk).write_text(
            raw_text(date, rtype, client, title, fname, minutes, basis), encoding="utf-8")
        return fname

    def run_script(self, script: Path, *args: str) -> subprocess.CompletedProcess:
        return subprocess.run(
            [sys.executable, str(script), *args, "--vault", str(self.vault)],
            capture_output=True, text=True, encoding="utf-8",
        )

    def runs_log(self):
        p = self.vault / "04_logs" / "runs.jsonl"
        if not p.exists():
            return []
        return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()]


# ====================================================================== worklog.py

class TestWorklog(VaultTestCase):
    def setUp(self) -> None:
        super().setUp()
        self.add_raw(self.dir_a, "2026-09-14", "meeting", "月次定例", 58,
                     "文字起こしの最終タイムスタンプ 00:58:10", nfd_name=True)
        self.add_raw(self.dir_a, "2026-09-02", "slack", "sample-shoji", 24, "メッセージ12件×2分")
        self.add_raw(self.dir_a, "2026-09-02", "line", "在庫の相談", 10, "メッセージ3件×2分（最低10分）")
        self.add_raw(self.dir_a, "2026-09-20", "meeting", "臨時打合せ", 72, "文字起こしの最終タイムスタンプ 01:12:00")
        self.add_raw(self.dir_a, "2026-09-21", "mail", "見積の件", 10, "メール1通=10分")
        self.add_raw(self.dir_a, "2026-09-25", "doc", "", None, "")  # title・所要目安が空
        self.add_raw(self.dir_a, "2026-09-26", "memo", "電話メモ", 5, "メモ1件=5分")
        self.add_raw(self.dir_a, "2026-08-31", "meeting", "前月の定例", 60, "…")   # 対象外の月
        self.add_raw(self.dir_a, "2026-10-01", "slack", "sample-shoji", 30, "…")  # 対象外の月
        # テスト工務店は 9月の記録なし（8月だけ）
        self.add_raw(self.dir_b, "2026-08-10", "mail", "問合せ", 10, "…", client=CLIENT_B)

    def test_fmt_minutes(self):
        mod = load_module(WORKLOG, "worklog_mod")
        self.assertEqual(mod.fmt_minutes(58), "58分")
        self.assertEqual(mod.fmt_minutes(70), "1時間10分")
        self.assertEqual(mod.fmt_minutes(120), "2時間")
        self.assertEqual(mod.fmt_minutes(0), "0分")

    def test_previous_month(self):
        mod = load_module(WORKLOG, "worklog_mod2")
        self.assertEqual(mod.previous_month(dt.date(2026, 10, 1)), "2026-09")
        self.assertEqual(mod.previous_month(dt.date(2026, 1, 15)), "2025-12")

    def test_generate_month_filter_and_totals(self):
        r = self.run_script(WORKLOG, "2026-09")
        self.assertEqual(r.returncode, 0, r.stderr)
        out = self.dir_a / "2026-09" / "worklog.md"
        self.assertTrue(out.exists())
        text = out.read_text(encoding="utf-8")

        # frontmatter（データ契約6章）
        self.assertTrue(text.startswith("---\ntype: worklog\n"))
        self.assertIn(f"client: {CLIENT_A}\n", text)
        self.assertIn("date: 2026-09\n", text)
        self.assertIn("generated_at: ", text)
        self.assertIn(f"  - raw/2026-09-14_meeting_{CLIENT_A}_月次定例.md\n", text)  # NFD名もNFCで出る
        self.assertNotIn("前月の定例", text)
        self.assertNotIn("2026-10-01", text)

        table = [ln for ln in text.splitlines() if ln.startswith("| 0")]
        self.assertEqual(len(table), 7)
        # 日付→ファイル名順（09-02 は line が slack より先）
        self.assertTrue(table[0].startswith("| 09-02 | line | 相談対応 |"))
        self.assertTrue(table[1].startswith("| 09-02 | slack | 相談対応 |"))
        self.assertIn("| 09-14 | meeting | 会議 | サンプル商事 | 月次定例 | 58分 | "
                      f"[[2026-09-14_meeting_{CLIENT_A}_月次定例]] |", text)
        self.assertIn("| 臨時打合せ | 1時間12分 |", text)
        self.assertIn("| 09-21 | mail | フォロー |", text)
        self.assertIn("| 09-25 | doc | 資料作成 | サンプル商事 | 要確認 | 要確認 |", text)
        self.assertIn("| 09-26 | memo | フォロー |", text)

        # 月末集計（推定）: 会議 58+72=130, チャット 24+10=34, メール 10, 資料 0(不明), メモ 5 → 179
        self.assertIn("## 月末集計（推定）", text)
        self.assertIn("| 会議 | 2件 | 2時間10分 |", text)
        self.assertIn("| チャット相談 | 2件 | 34分 |", text)
        self.assertIn("| メール | 1通 | 10分 |", text)
        self.assertIn("| 資料 | 1件 | 0分 |", text)
        self.assertIn("| メモ | 1件 | 5分 |", text)
        self.assertIn("| **合計** | **7件** | **2時間59分** |", text)
        self.assertIn("機械的に算出した推定値", text)
        self.assertIn("要確認の 1件 は合計時間に含めていない", text)
        self.assertIn("<details>", text)
        self.assertIn("メッセージ12件×2分", text)
        # 書き換えルールのHTMLコメント
        self.assertIn("<!-- このファイルは 03_scripts/worklog.py", text)
        self.assertIn("書き換えないこと", text)
        # フッター
        last = [ln for ln in text.splitlines() if ln.strip()][-1]
        self.assertTrue(last.startswith("本資料は以下の記録から作成: [[2026-09-02_line_"))

    def test_no_record_client(self):
        r = self.run_script(WORKLOG, "2026-09")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn(f"{CLIENT_B}: 記録なし", r.stdout)
        self.assertFalse((self.dir_b / "2026-09").exists())
        log = self.runs_log()
        self.assertEqual(log[-1]["script"], "worklog")
        self.assertEqual(log[-1]["summary"]["no_records"], [CLIENT_B])
        self.assertEqual(log[-1]["summary"]["clients"][CLIENT_A]["total_minutes"], 179)

    def test_client_filter_nfc(self):
        # --client は NFC、フォルダ名は NFD でも一致すること
        r = self.run_script(WORKLOG, "2026-08", "--client", CLIENT_A)
        self.assertEqual(r.returncode, 0, r.stderr)
        text = (self.dir_a / "2026-08" / "worklog.md").read_text(encoding="utf-8")
        self.assertIn("| 前月の定例 | 1時間 |", text)
        self.assertFalse((self.dir_b / "2026-08").exists())  # 絞り込みでBは処理しない

    def test_generated_worklog_passes_check_sources(self):
        self.run_script(WORKLOG, "2026-09")
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)

    def test_bad_month(self):
        r = self.run_script(WORKLOG, "2026-13")
        self.assertEqual(r.returncode, 2)


# ====================================================================== check_sources.py

GOOD_DOC = """---
type: minutes
client: サンプル商事
date: 2026-09-14
generated_at: 2026-10-01T09:00:00
sources:
  - raw/2026-09-14_meeting_サンプル商事_月次定例.md
---

# 議事録: サンプル商事 月次定例 2026-09-14

<!-- 書き方: このコメント内の要確認は数えない -->

## 参加者
- 山田、佐藤 部長

## 宿題・次回まで
- 見積の再提出（期限: 要確認）

## 期待効果
- 要試算

---
本資料は以下の記録から作成: [[2026-09-14_meeting_サンプル商事_月次定例]]
"""


class TestCheckSources(VaultTestCase):
    def setUp(self) -> None:
        super().setUp()
        self.raw_name = self.add_raw(self.dir_a, "2026-09-14", "meeting", "月次定例", 58, "…")
        self.minutes_dir = self.dir_a / "2026-09" / "minutes"
        self.minutes_dir.mkdir(parents=True)

    def write_doc(self, text: str, name: str = "2026-09-14_月次定例.md") -> Path:
        p = self.minutes_dir / name
        p.write_text(text, encoding="utf-8")
        return p

    def test_good_document_exit0(self):
        self.write_doc(GOOD_DOC)
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("ERROR 0", r.stdout)

    def test_marker_count(self):
        self.write_doc(GOOD_DOC)
        r = self.run_script(CHECK, "2026-09")
        # コメント内の「要確認」は数えず、本文の1件だけ
        # 行番号はファイル先頭（frontmatter 含む）からの行
        self.assertIn("要確認 1件（L18）", r.stdout)
        self.assertIn("要試算 1件（L21）", r.stdout)
        self.assertIn("| INFO |", r.stdout)
        log = self.runs_log()[-1]
        self.assertEqual(log["script"], "check_sources")
        self.assertEqual(log["summary"]["要確認"], 1)
        self.assertEqual(log["summary"]["要試算"], 1)

    def test_missing_source_exit1(self):
        self.write_doc(GOOD_DOC.replace("raw/2026-09-14_meeting", "raw/2026-09-15_meeting"))
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 1)
        self.assertIn("sources のファイルが存在しない: raw/2026-09-15_meeting", r.stdout)

    def test_missing_footer(self):
        text = GOOD_DOC.split("\n---\n本資料は")[0] + "\n"
        self.write_doc(text)
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 1)
        self.assertIn("フッターが無い", r.stdout)

    def test_footer_not_at_end(self):
        # フッターの後ろに本文が続いていたら「末尾に無い」扱い
        self.write_doc(GOOD_DOC + "\n追記された段落\n")
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 1)
        self.assertIn("フッターが無い", r.stdout)

    def test_inline_list_sources(self):
        text = GOOD_DOC.replace(
            "sources:\n  - raw/2026-09-14_meeting_サンプル商事_月次定例.md\n",
            "sources: [raw/2026-09-14_meeting_サンプル商事_月次定例.md, _profile.md]\n")
        self.write_doc(text)
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        text2 = text.replace("_profile.md]", "_profile_none.md]")
        self.write_doc(text2)
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 1)
        self.assertIn("存在しない: _profile_none.md", r.stdout)

    def test_missing_frontmatter_keys(self):
        self.write_doc("# frontmatter なし\n\n---\n本資料は以下の記録から作成: [[x]]\n", "a.md")
        self.write_doc(GOOD_DOC.replace("client: サンプル商事\n", ""), "b.md")
        self.write_doc(GOOD_DOC.replace(
            "sources:\n  - raw/2026-09-14_meeting_サンプル商事_月次定例.md\n", "sources: \n"), "c.md")
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 1)
        self.assertIn("minutes/a.md | frontmatter（先頭の --- で囲んだ設定欄）が無い", r.stdout)
        self.assertIn("minutes/b.md | frontmatter に client が無い", r.stdout)
        self.assertIn("minutes/c.md | sources が空", r.stdout)

    def test_broken_link_is_warn(self):
        self.write_doc(GOOD_DOC.replace("## 参加者", "## 参加者\n参照: [[存在しないノート]]"))
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 0, r.stdout)
        self.assertIn("| WARN |", r.stdout)
        self.assertIn("[[存在しないノート]]", r.stdout)

    def test_sent_is_excluded(self):
        self.write_doc(GOOD_DOC)
        sent = self.dir_a / "2026-09" / "sent"
        sent.mkdir()
        (sent / "送付記録.md").write_text("# 送付記録\n| a |\n", encoding="utf-8")  # frontmatter なし
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 0, r.stdout)
        self.assertNotIn("送付記録", r.stdout)

    def test_source_outside_client_dir(self):
        self.write_doc(GOOD_DOC.replace(
            "raw/2026-09-14_meeting_サンプル商事_月次定例.md", "../テスト工務店/_profile.md"))
        r = self.run_script(CHECK, "2026-09")
        self.assertEqual(r.returncode, 1)
        self.assertIn("クライアントフォルダの外", r.stdout)


# ====================================================================== send_log.py

class TestSendLog(VaultTestCase):
    def setUp(self) -> None:
        super().setUp()
        self.sent = self.dir_a / "2026-09" / "sent"
        self.sent.mkdir(parents=True)
        (self.sent / "月次レポート_2026-09.pdf").write_bytes(b"%PDF-1.4 dummy")
        (self.sent / "議事録_0914.pdf").write_bytes(b"%PDF-1.4 minutes")
        self.record = self.sent / "送付記録.md"

    def send(self, *files: str, extra=()):
        return self.run_script(SEND, "--client", CLIENT_A, "--month", "2026-09",
                               "--to", "佐藤 様", "--files", *files, *extra)

    def test_create_and_append(self):
        r = self.send("月次レポート_2026-09.pdf", "議事録_0914.pdf", extra=("--note", "メールで送付"))
        self.assertEqual(r.returncode, 0, r.stderr)
        text = self.record.read_text(encoding="utf-8")
        self.assertIn("| 送付日時 | 宛先 | 添付ファイル | SHA256先頭12桁 | メモ |", text)
        import hashlib
        h = hashlib.sha256(b"%PDF-1.4 dummy").hexdigest()[:12]
        self.assertIn(f"| 佐藤 様 | 月次レポート_2026-09.pdf | `{h}` | メールで送付 |", text)
        rows = [ln for ln in text.splitlines() if ln.startswith("| 20")]
        self.assertEqual(len(rows), 2)
        log = self.runs_log()[-1]
        self.assertEqual(log["script"], "send_log")
        self.assertEqual(log["summary"]["files"][0]["sha256_12"], h)

    def test_existing_rows_unchanged(self):
        self.send("月次レポート_2026-09.pdf")
        before = self.record.read_bytes()
        r = self.send("議事録_0914.pdf", extra=("--note", "追送"))
        self.assertEqual(r.returncode, 0, r.stderr)
        after = self.record.read_bytes()
        self.assertTrue(after.startswith(before), "既存の内容が変わっている")
        self.assertEqual(after.count("\n".encode()), before.count("\n".encode()) + 1)

    def test_existing_without_trailing_newline(self):
        original = "# 手で作った送付記録\n\n| 送付日時 | 宛先 | 添付ファイル | SHA256先頭12桁 | メモ |\n|---|---|---|---|---|\n| 2026-09-01 10:00 | 旧 | a.pdf | `000000000000` | |"
        self.record.write_text(original, encoding="utf-8")
        r = self.send("議事録_0914.pdf")
        self.assertEqual(r.returncode, 0, r.stderr)
        text = self.record.read_text(encoding="utf-8")
        self.assertTrue(text.startswith(original + "\n| 20"))

    def test_file_not_in_sent(self):
        (self.dir_a / "2026-09" / "report.md").write_text("x", encoding="utf-8")
        r = self.send("report.md")
        self.assertEqual(r.returncode, 1)
        self.assertIn("sent/ にありません", r.stderr)
        self.assertFalse(self.record.exists(), "エラー時に送付記録を作ってはいけない")

    def test_partial_error_writes_nothing(self):
        self.send("月次レポート_2026-09.pdf")
        before = self.record.read_bytes()
        r = self.send("議事録_0914.pdf", "無いファイル.pdf")
        self.assertEqual(r.returncode, 1)
        self.assertEqual(self.record.read_bytes(), before)

    def test_no_copy_from_outside(self):
        outside = self.tmp / "外のファイル.pdf"
        outside.write_bytes(b"x")
        r = self.send(str(outside))
        self.assertEqual(r.returncode, 1)
        self.assertFalse((self.sent / "外のファイル.pdf").exists())

    def test_dry_run(self):
        r = self.send("月次レポート_2026-09.pdf", extra=("--dry-run",))
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("[dry-run]", r.stdout)
        self.assertFalse(self.record.exists())
        self.assertEqual(self.runs_log(), [])

    def test_nfd_filename_in_sent(self):
        nfd = unicodedata.normalize("NFD", "プレゼン資料.pdf")
        (self.sent / nfd).write_bytes(b"y")
        r = self.send("プレゼン資料.pdf")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("| プレゼン資料.pdf |", self.record.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
