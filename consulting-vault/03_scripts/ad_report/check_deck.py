#!/usr/bin/env python3
"""広告運用レポート（insights.json と pptx）を検査するスクリプト。ファイルは書き換えない。

検査する内容:
  ERROR … analysis.json が無い / insights.json が JSON として読めない /
          所見に解決できないプレースホルダ（analysis.json に無いキー）がある /
          所見が書いてあるのに sources が空 / sources のキーやファイルが実在しない /
          pptx が無い / ノートに「数字の出どころ」が無いスライドがある
  WARN  … 所見の文章に生の数字（数字＋単位、¥付きの金額など）が混ざっている /
          所見が目安の文字数を大きく超える / 対応するスライドの無い所見がある /
          pptx が analysis.json・insights.json より古い（再ビルドが必要）
  INFO  … 要確認（未記入の所見）の件数・スライド総数・セクション別の枚数

pptx の XML は標準ライブラリの ElementTree で読む（このツール自身が作ったファイルが対象。
defusedxml は標準ライブラリ外のため使わない。expat は外部エンティティを解決しない）。

結果を表で表示し、ERROR が1件でもあれば終了コード 1、無ければ 0。

使い方（vault ルートで実行）:
  python3 03_scripts/ad_report/check_deck.py --client サンプル商事 --month 2026-09
  python3 03_scripts/ad_report/check_deck.py --client サンプル商事 --month 2026-09 --vault /path/to/vault

標準ライブラリのみ・Python 3.9 で動作。実行ごとに 04_logs/runs.jsonl に1行追記する。
"""
from __future__ import annotations

import argparse
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ad_common as C  # noqa: E402

SCRIPT_NAME = "check_deck"
NOTES_MARK = "【数字の出どころ】"
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


class Finding:
    def __init__(self, level: str, where: str, message: str) -> None:
        self.level, self.where, self.message = level, where, message


def _text_of(root) -> Tuple[str, List[str]]:
    paras = []
    for p in root.iter(f"{{{NS['a']}}}p"):
        paras.append("".join(t.text or "" for t in p.iter(f"{{{NS['a']}}}t")))
    return "\n".join(paras), paras


def read_pptx(path: Path) -> List[dict]:
    """スライドを順番に読み、各スライドの本文とノートを返す。"""
    out = []
    with zipfile.ZipFile(path) as z:
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels.findall("rel:Relationship", NS)}
        lst = pres.find("p:sldIdLst", NS)
        for sld in (list(lst) if lst is not None else []):
            rid = sld.get(f"{{{NS['r']}}}id")
            spath = posixpath.normpath(posixpath.join("ppt", targets[rid]))
            sroot = ET.fromstring(z.read(spath))
            text, _ = _text_of(sroot)
            charts = 0
            for gd in sroot.iter(f"{{{NS['a']}}}graphicData"):
                if gd.get("uri", "").endswith("/chart"):
                    charts += 1
            notes = ""
            rel_path = posixpath.join(posixpath.dirname(spath), "_rels", posixpath.basename(spath) + ".rels")
            if rel_path in z.namelist():
                srels = ET.fromstring(z.read(rel_path))
                for r in srels.findall("rel:Relationship", NS):
                    if r.get("Type", "").endswith("/notesSlide"):
                        npath = posixpath.normpath(posixpath.join(posixpath.dirname(spath), r.get("Target")))
                        notes, _ = _text_of(ET.fromstring(z.read(npath)))
            sid = ""
            for line in notes.split("\n"):
                if line.startswith("slide_id:"):
                    sid = line.split(":", 1)[1].strip()
                    break
            out.append({"text": text, "notes": notes, "slide_id": sid, "charts": charts})
    return out


def section_of(slide_id: str) -> str:
    if not slide_id:
        return "（不明）"
    if slide_id in ("campaigns", "campaigns.chart"):
        return "キャンペーン比較"
    head = slide_id.split(".")[0]
    return SECTION_NAMES.get(head, head)


SECTION_NAMES = {
    "cover": "表紙", "toc": "目次", "divider": "中扉", "summary": "サマリー", "targets": "目標対比", "mom": "前月比較",
    "measurement_notes": "計測上の注意", "daily": "日別推移", "weekly": "週別・曜日別", "weekday": "週別・曜日別",
    "campaign": "キャンペーン別", "adset": "広告セット別", "ad": "広告別", "ranking": "上位・下位",
    "age_gender": "年齢×性別", "placement": "配置", "device": "デバイス", "region": "地域", "issues": "所見・課題",
    "next_actions": "次月施策", "appendix": "Appendix", "glossary": "Appendix",
}


def check(vault: Path, client_dir: Path, month: str, pptx: Optional[Path], ipath: Optional[Path]
          ) -> Tuple[List[Finding], dict]:
    f: List[Finding] = []
    stats: dict = {"slides": 0, "unknown_insights": 0, "raw_numbers": 0, "by_section": {}}
    ads_dir = client_dir / month / "ads"
    apath = ads_dir / "analysis.json"
    if not apath.is_file():
        f.append(Finding("ERROR", C.rel_to(apath, vault), "analysis.json が無い（ad_analyze.py を先に実行）"))
        return f, stats
    analysis = C.load_json(apath)
    values: Dict[str, dict] = analysis.get("values", {})
    slots = {s["id"]: s for s in analysis.get("insight_slots", [])}

    ipath = ipath or ads_dir / "insights.json"
    insights: dict = {}
    if ipath.is_file():
        try:
            insights = C.load_json(ipath)
        except ValueError as e:
            f.append(Finding("ERROR", ipath.name, f"JSON として読めない（{e}）"))
    else:
        f.append(Finding("INFO", ipath.name, "insights.json が無い（全スライドの所見が「要確認」になる）"))
    slides_in = insights.get("slides", {}) if isinstance(insights, dict) else {}
    if not isinstance(slides_in, dict):
        f.append(Finding("ERROR", ipath.name, "slides がオブジェクト（{ }）ではない"))
        slides_in = {}

    empty = 0
    for sid in slots:
        slot = slides_in.get(sid)
        text = C.insight_text(slot)
        where = f"insights: {sid}"
        if not text:
            empty += 1
            continue
        resolved, unknown = C.resolve_placeholders(text, values)
        for k in unknown:
            f.append(Finding("ERROR", where, f"解決できないプレースホルダ {{{{{k}}}}}（analysis.json の values に無い）"))
        raws = C.find_raw_numbers(text)
        if raws:
            stats["raw_numbers"] += len(raws)
            f.append(Finding("WARN", where, "生の数字が混ざっている: " + " / ".join(f"「{r}」" for r in raws[:6])
                             + "（数字は {{キー}} で書く）"))
        srcs = C.insight_sources(slot)
        if not srcs:
            f.append(Finding("ERROR", where, "所見があるのに sources が空（根拠のキーかファイルを書く）"))
        for s in srcs:
            if C.is_file_source(s):
                if not (client_dir / s).exists():
                    f.append(Finding("ERROR", where, f"sources のファイルが存在しない: {s}（クライアントフォルダからの相対パス）"))
            elif not C.analysis_key_exists(s, values, analysis):
                f.append(Finding("ERROR", where, f"sources のキーが analysis.json に無い: {s}"))
        mx = slots[sid].get("max_chars")
        if mx and len(resolved) > mx * 1.3:
            f.append(Finding("WARN", where, f"所見が長い（置換後 {len(resolved)}文字／目安 {mx}文字）。スライドで文字が小さくなる"))
    for sid in slides_in:
        if sid not in slots and C.insight_text(slides_in[sid]):
            f.append(Finding("WARN", f"insights: {sid}", "対応するスライドIDが無い（この所見は表示されない）"))
    stats["unknown_insights"] = empty
    stats["slots"] = len(slots)

    pptx = pptx or ads_dir / f"{C.nfc(client_dir.name)}_{month}_広告運用レポート.pptx"
    if not pptx.is_file():
        f.append(Finding("ERROR", pptx.name, "pptx が無い（build_deck.py を実行）"))
        return f, stats
    newest_src = max(apath.stat().st_mtime, ipath.stat().st_mtime if ipath.is_file() else 0)
    if pptx.stat().st_mtime + 1 < newest_src:
        f.append(Finding("WARN", pptx.name, "pptx が analysis.json / insights.json より古い（build_deck.py で作り直す）"))
    try:
        slides = read_pptx(pptx)
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as e:
        f.append(Finding("ERROR", pptx.name, f"pptx を読めない（{e}）"))
        return f, stats
    stats["slides"] = len(slides)
    stats["charts"] = sum(s["charts"] for s in slides)
    deck_unknown = 0
    for i, s in enumerate(slides, start=1):
        if NOTES_MARK not in s["notes"]:
            f.append(Finding("ERROR", f"スライド {i}", "ノートに「数字の出どころ」が無い"))
        deck_unknown += s["text"].count(C.UNKNOWN_INSIGHT)
        sec = section_of(s["slide_id"])
        stats["by_section"][sec] = stats["by_section"].get(sec, 0) + 1
        if "{{" in s["text"] or "}}" in s["text"]:
            f.append(Finding("ERROR", f"スライド {i}", "本文に置換されていない {{ }} が残っている"))
    stats["deck_unknown"] = deck_unknown
    return f, stats


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="広告運用レポートの insights.json と pptx を検査する（書き換えない）")
    ap.add_argument("--client", required=True)
    ap.add_argument("--month", help="対象月 YYYY-MM（省略時は前月）")
    ap.add_argument("--vault")
    ap.add_argument("--pptx", help="検査する pptx（既定: 標準の出力先）")
    ap.add_argument("--insights", help="insights.json の場所（既定: analysis.json と同じフォルダ）")
    args = ap.parse_args(argv)
    month = args.month or C.previous_month_of_today()
    if not C.MONTH_RE.match(month):
        print(f"エラー: 月は YYYY-MM 形式で指定してください（受け取った値: {month}）", file=sys.stderr)
        return 2
    vault = C.resolve_vault(args.vault)
    client_dir = C.find_client_dir(vault, args.client)
    if client_dir is None:
        print(f"エラー: クライアント「{args.client}」のフォルダが {vault / '01_clients'} にありません", file=sys.stderr)
        return 2
    findings, stats = check(vault, client_dir, month, Path(args.pptx) if args.pptx else None,
                            Path(args.insights) if args.insights else None)
    order = {"ERROR": 0, "WARN": 1, "INFO": 2}
    findings.sort(key=lambda x: order.get(x.level, 9))
    n_err = sum(1 for x in findings if x.level == "ERROR")
    n_warn = sum(1 for x in findings if x.level == "WARN")
    print(f"check_deck: {C.nfc(client_dir.name)} {month}")
    if findings:
        print()
        print("| 区分 | 場所 | 内容 |")
        print("|---|---|---|")
        for x in findings:
            print(f"| {x.level} | {x.where.replace('|', '/')} | {x.message.replace('|', '/')} |")
    print()
    print(f"スライド総数: {stats.get('slides', 0)}枚（ネイティブグラフ {stats.get('charts', 0)}個）")
    if stats.get("by_section"):
        print("セクション別: " + "、".join(f"{k} {v}" for k, v in stats["by_section"].items()))
    print(f"要確認（未記入の所見）: {stats.get('unknown_insights', 0)} / {stats.get('slots', 0)}枠"
          f"（スライド上の「所見：要確認」{stats.get('deck_unknown', 0)}件）")
    print(f"結果: ERROR {n_err} / WARN {n_warn}")
    print("→ ERROR を直して再実行してください。" if n_err else "→ ERROR なし。")
    C.append_run_log(vault, SCRIPT_NAME, {"client": C.nfc(client_dir.name), "month": month, "errors": n_err,
                                          "warnings": n_warn, "slides": stats.get("slides", 0),
                                          "unknown_insights": stats.get("unknown_insights", 0)})
    return 1 if n_err else 0


if __name__ == "__main__":
    sys.exit(main())
