#!/usr/bin/env python3
"""架空クライアント「サンプル商事」の Meta 広告エクスポート（CSV）を作るスクリプト。

すべて架空の数字。乱数シードは月ごとに固定（"sample-shoji-<YYYY-MM>"）なので、何度実行しても
同じバイト列のファイルができる（テストで再現性を確認している）。

作るもの（1か月あたり）:
  広告_日別.csv          広告 × 日 の明細（キャンペーン3 × 広告セット3 × 広告3 = 27本）
  期間合計_キャンペーン.csv  キャンペーン単位・期間全体（リーチ・フリークエンシーは期間全体の重複なし値）
  内訳_年齢性別.csv       キャンペーン × 年齢 × 性別 の月合計
  内訳_配置.csv           キャンペーン × プラットフォーム × 配置 の月合計
  内訳_デバイス.csv       キャンペーン × インプレッションデバイス の月合計
  内訳_地域.csv           キャンペーン × 地域 の月合計
列名は日本語UIのエクスポートを想定、UTF-8（BOM付き）、改行 LF。

想定した「それらしさ」:
  - 週末と 9/19〜9/23（シルバーウィーク）は消化・CVが増える。8月はお盆（8/13〜8/16）にCVRが落ちる
  - 「ブロード_女性25-54 / 新規_UGC動画_9月新作」は 9/15 開始（8月には無い）
  - 「既存顧客_リピート / RT_クーポン訴求」は 9/20 で停止（以降は行が無い＝配信なし）
  - 「RT_レビュー動画」は月の後半ほど CTR が落ちる（クリエイティブ疲れ）
  - 認知_動画 の「結果」はリーチ（結果インジケーター = reach）。購入はほぼ出ない。
    「興味関心_料理 / 認知_職人インタビュー」は購入 0 件（ゼロ除算の確認用）

使い方（vault ルートで）:
  python3 03_scripts/ad_report/tools/make_sample.py                  # 2026-08 と 2026-09 を 99_samples/ads/サンプル商事/ に作る
  python3 03_scripts/ad_report/tools/make_sample.py --out /tmp/x --months 2026-09

標準ライブラリのみ・Python 3.9 で動作。
"""
from __future__ import annotations

import argparse
import calendar
import csv
import datetime as dt
import random
import sys
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

CLIENT = "サンプル商事"
DEFAULT_MONTHS = ["2026-08", "2026-09"]
PURCHASE_IND = "actions:offsite_conversion.fb_pixel_purchase"
REACH_IND = "reach"
ATTR_CV = "7日間のクリックまたは1日間のビュー"

# (キャンペーン名, 結果の種類, 基準: 1広告1日の消化金額, CPM, リンクCTR, CVR, 平均購入単価, アトリビューション設定,
#  広告セット [(名前, 消化倍率, CVR倍率, [広告(名前, 動画か, CTR倍率, CVR倍率, 開始日, 終了日, 疲れ)])])
CAMPAIGNS = [
    ("新規獲得_コンバージョン", PURCHASE_IND, 5200, 1450.0, 0.0115, 0.021, 6800, ATTR_CV, [
        ("類似オーディエンス_購入者1%", 1.30, 1.20, [
            ("新規_動画_使い方15秒", True, 1.20, 1.00, 1, 31, 0.0),
            ("新規_静止画_レビュー訴求", False, 0.90, 1.15, 1, 31, 0.0),
            ("新規_カルーセル_人気5選", False, 1.00, 0.85, 1, 31, 0.0),
        ]),
        ("興味関心_キッチン雑貨", 1.00, 0.90, [
            ("新規_動画_使い方15秒", True, 1.10, 0.95, 1, 31, 0.0),
            ("新規_静止画_レビュー訴求", False, 0.95, 1.05, 1, 31, 0.0),
            ("新規_カルーセル_人気5選", False, 1.05, 0.80, 1, 31, 0.0),
        ]),
        ("ブロード_女性25-54", 0.85, 0.80, [
            ("新規_動画_使い方15秒", True, 1.00, 0.90, 1, 31, 0.0),
            ("新規_静止画_レビュー訴求", False, 0.85, 1.10, 1, 31, 0.0),
            ("新規_UGC動画_9月新作", True, 1.45, 1.25, 15, 31, 0.0),
        ]),
    ]),
    ("リターゲティング", PURCHASE_IND, 2100, 2250.0, 0.0160, 0.046, 7400, ATTR_CV, [
        ("カート放棄_7日", 0.80, 1.50, [
            ("RT_クーポン訴求", False, 1.00, 1.20, 1, 31, 0.0),
            ("RT_動的_閲覧商品", False, 1.30, 1.00, 1, 31, 0.0),
            ("RT_レビュー動画", True, 0.85, 0.90, 1, 31, 0.018),
        ]),
        ("サイト訪問_30日", 1.00, 1.00, [
            ("RT_クーポン訴求", False, 1.00, 1.10, 1, 31, 0.0),
            ("RT_動的_閲覧商品", False, 1.25, 0.95, 1, 31, 0.0),
            ("RT_レビュー動画", True, 0.80, 0.85, 1, 31, 0.015),
        ]),
        ("既存顧客_リピート", 0.90, 1.10, [
            ("RT_クーポン訴求", False, 1.05, 1.30, 1, 20, 0.0),
            ("RT_動的_閲覧商品", False, 1.20, 1.00, 1, 31, 0.0),
            ("RT_レビュー動画", True, 0.90, 0.95, 1, 31, 0.012),
        ]),
    ]),
    ("認知_動画", REACH_IND, 1750, 620.0, 0.0042, 0.0065, 6100, "", [
        ("ブロード_全国", 1.20, 1.00, [
            ("認知_ブランド紹介30秒", True, 1.00, 1.00, 1, 31, 0.0),
            ("認知_職人インタビュー", True, 0.90, 0.80, 1, 31, 0.0),
            ("認知_季節の新商品", True, 1.15, 1.20, 1, 31, 0.0),
        ]),
        ("興味関心_料理", 1.00, 1.10, [
            ("認知_ブランド紹介30秒", True, 1.05, 1.00, 1, 31, 0.0),
            ("認知_職人インタビュー", True, 0.95, 0.00, 1, 31, 0.0),
            ("認知_季節の新商品", True, 1.10, 1.10, 1, 31, 0.0),
        ]),
        ("類似_サイト訪問者3%", 0.90, 1.30, [
            ("認知_ブランド紹介30秒", True, 1.00, 1.10, 1, 31, 0.0),
            ("認知_職人インタビュー", True, 0.90, 0.90, 1, 31, 0.0),
            ("認知_季節の新商品", True, 1.20, 1.30, 1, 31, 0.0),
        ]),
    ]),
]

# 月ごとの全体の調子（消化倍率, CVR倍率）。8月は9月よりやや弱い。
MONTH_FACTOR = {"2026-08": (0.94, 0.90), "2026-09": (1.00, 1.00)}
# 曜日（月=0）ごとの（消化倍率, CVR倍率）
WEEKDAY_FACTOR = [(0.95, 0.96), (0.97, 0.97), (1.00, 1.00), (1.00, 1.00), (1.04, 1.02), (1.12, 1.10), (1.10, 1.08)]

MAIN_COLUMNS = [
    "レポート開始日", "レポート終了日", "キャンペーン名", "広告セット名", "広告名", "日",
    "結果", "結果インジケーター", "リーチ", "フリークエンシー", "インプレッション", "消化金額 (JPY)",
    "結果の単価", "リンクのクリック", "CTR(リンククリックスルー率)", "クリック(すべて)",
    "CPM(1,000インプレッション単価) (JPY)", "CPC(リンククリックの単価) (JPY)", "ランディングページビュー",
    "購入", "購入のコンバージョン値", "購入ROAS(広告費用対効果)", "動画の3秒再生数", "ThruPlay",
    "アトリビューション設定",
]

AGE_GENDER = [
    # (年齢, 性別, 消化の重み, CTR倍率, CVR倍率)
    ("18-24", "female", 0.050, 1.10, 0.60), ("18-24", "male", 0.030, 0.90, 0.45), ("18-24", "unknown", 0.004, 0.80, 0.40),
    ("25-34", "female", 0.140, 1.10, 0.95), ("25-34", "male", 0.070, 0.90, 0.70), ("25-34", "unknown", 0.006, 0.80, 0.60),
    ("35-44", "female", 0.180, 1.05, 1.30), ("35-44", "male", 0.080, 0.95, 0.90), ("35-44", "unknown", 0.006, 0.80, 0.70),
    ("45-54", "female", 0.160, 1.00, 1.35), ("45-54", "male", 0.070, 0.95, 0.95), ("45-54", "unknown", 0.005, 0.80, 0.70),
    ("55-64", "female", 0.080, 0.95, 1.10), ("55-64", "male", 0.040, 0.90, 0.80), ("55-64", "unknown", 0.003, 0.80, 0.60),
    ("65+", "female", 0.030, 0.90, 0.90), ("65+", "male", 0.020, 0.85, 0.70), ("65+", "unknown", 0.002, 0.80, 0.50),
]
PLACEMENTS = [
    ("instagram", "フィード", 0.24, 1.10, 1.10), ("instagram", "ストーリーズ", 0.16, 0.80, 0.90),
    ("instagram", "リール", 0.14, 0.90, 0.80), ("instagram", "発見タブ", 0.03, 0.90, 0.85),
    ("facebook", "フィード", 0.22, 1.00, 1.20), ("facebook", "リール", 0.05, 0.85, 0.70),
    ("facebook", "インストリーム動画", 0.04, 0.60, 0.50), ("facebook", "Marketplace", 0.04, 0.95, 0.90),
    ("audience_network", "ネイティブ、バナー、インタースティシャル", 0.06, 0.70, 0.25),
    ("messenger", "受信箱", 0.02, 0.60, 0.50),
]
DEVICES = [
    ("iphone", 0.60, 1.05, 1.10), ("android_smartphone", 0.29, 0.95, 0.85), ("ipad", 0.04, 0.90, 1.00),
    ("android_tablet", 0.02, 0.85, 0.70), ("desktop", 0.05, 0.80, 1.30),
]
REGIONS = [
    ("東京都", 0.22, 1.05, 1.10), ("神奈川県", 0.10, 1.00, 1.05), ("大阪府", 0.10, 1.00, 1.00),
    ("愛知県", 0.07, 0.95, 0.95), ("埼玉県", 0.07, 0.95, 0.95), ("千葉県", 0.06, 0.95, 0.95),
    ("福岡県", 0.05, 1.00, 0.90), ("北海道", 0.05, 0.95, 0.85), ("兵庫県", 0.05, 1.00, 1.00),
    ("京都府", 0.03, 1.00, 1.05), ("静岡県", 0.03, 0.95, 0.90), ("広島県", 0.02, 0.95, 0.90),
    ("宮城県", 0.02, 0.95, 0.85), ("その他", 0.13, 0.90, 0.80),
]


def fmt_num(v: float, digits: int = 0) -> str:
    if digits == 0:
        return str(int(round(v)))
    return f"{v:.{digits}f}"


def blank_if_zero(v: float, digits: int = 0) -> str:
    return "" if v == 0 else fmt_num(v, digits)


def is_sale_day(d: dt.date) -> Tuple[float, float]:
    if d.month == 9 and 19 <= d.day <= 23:
        return 1.20, 1.30   # シルバーウィーク
    if d.month == 8 and 13 <= d.day <= 16:
        return 0.95, 0.85   # お盆
    return 1.0, 1.0


def split_int(total: int, weights: Sequence[float]) -> List[int]:
    """total を weights の比で整数に分ける（最大剰余法。合計は必ず total に一致）。"""
    s = sum(weights)
    if total <= 0 or s <= 0:
        return [0] * len(weights)
    raw = [total * w / s for w in weights]
    base = [int(x) for x in raw]
    rest = total - sum(base)
    order = sorted(range(len(raw)), key=lambda i: (-(raw[i] - base[i]), i))
    for i in order[:rest]:
        base[i] += 1
    return base


def gen_month(month: str) -> Dict[str, List[List[str]]]:
    rng = random.Random(f"sample-shoji-{month}")
    y, m = int(month[:4]), int(month[5:7])
    ndays = calendar.monthrange(y, m)[1]
    first = dt.date(y, m, 1)
    last = dt.date(y, m, ndays)
    m_spend, m_cvr = MONTH_FACTOR.get(month, (1.0, 1.0))

    def u(a: float, b: float) -> float:
        return a + (b - a) * rng.random()

    main_rows: List[List[str]] = [MAIN_COLUMNS]
    camp_tot: Dict[str, Dict[str, float]] = {}
    for (cname, indicator, base_spend, cpm, ctr, cvr, aov, attr, adsets) in CAMPAIGNS:
        tot = camp_tot.setdefault(cname, {"spend": 0, "impressions": 0, "reach_daily": 0, "link_clicks": 0,
                                          "clicks_all": 0, "purchases": 0, "purchase_value": 0})
        for (sname, s_spend, s_cvr, ads) in adsets:
            for (aname, is_video, a_ctr, a_cvr, start, end, fatigue) in ads:
                if month != "2026-09" and start > 1:
                    continue  # 9月の新作は他の月には無い
                eff_end = end if month == "2026-09" else 31  # 停止は9月だけ
                for day in range(1, ndays + 1):
                    d = dt.date(y, m, day)
                    if day < start or day > eff_end:
                        continue
                    wd_s, wd_c = WEEKDAY_FACTOR[d.weekday()]
                    sale_s, sale_c = is_sale_day(d)
                    spend = round(base_spend * s_spend * m_spend * wd_s * sale_s * u(0.82, 1.18))
                    imps = round(spend / (cpm * u(0.9, 1.1)) * 1000)
                    reach = round(imps / u(1.05, 1.25))
                    fat = max(0.4, 1.0 - fatigue * (day - 1))
                    link = round(imps * ctr * a_ctr * fat * u(0.8, 1.2))
                    clicks_all = round(link * u(1.6, 2.0))
                    lpv = round(link * u(0.72, 0.88))
                    exp_cv = link * cvr * s_cvr * a_cvr * m_cvr * wd_c * sale_c * u(0.7, 1.3)
                    purchases = int(exp_cv + rng.random())
                    pvalue = round(purchases * aov * u(0.85, 1.2))
                    v3 = round(imps * u(0.2, 0.3)) if is_video else 0
                    thru = round(imps * u(0.06, 0.1)) if is_video else 0
                    results = purchases if indicator == PURCHASE_IND else reach
                    ds = d.isoformat()
                    main_rows.append([
                        ds, ds, cname, sname, aname, ds,
                        blank_if_zero(results), indicator, fmt_num(reach),
                        fmt_num(imps / reach if reach else 0, 2), fmt_num(imps), fmt_num(spend),
                        blank_if_zero(spend / results if results else 0, 2), blank_if_zero(link),
                        fmt_num(link / imps * 100 if imps else 0, 2), blank_if_zero(clicks_all),
                        fmt_num(spend / imps * 1000 if imps else 0, 2),
                        blank_if_zero(spend / link if link else 0, 2), blank_if_zero(lpv),
                        blank_if_zero(purchases), blank_if_zero(pvalue),
                        blank_if_zero(pvalue / spend if spend else 0, 2),
                        blank_if_zero(v3), blank_if_zero(thru), attr,
                    ])
                    tot["spend"] += spend
                    tot["impressions"] += imps
                    tot["reach_daily"] += reach
                    tot["link_clicks"] += link
                    tot["clicks_all"] += clicks_all
                    tot["purchases"] += purchases
                    tot["purchase_value"] += pvalue

    files: Dict[str, List[List[str]]] = {"広告_日別.csv": main_rows}
    fs, ls = first.isoformat(), last.isoformat()

    # 期間合計（キャンペーン単位）: リーチは期間全体の重複なし値として日別合算の約45%にする
    period = [["レポート開始日", "レポート終了日", "キャンペーン名", "結果", "結果インジケーター", "リーチ",
               "フリークエンシー", "インプレッション", "消化金額 (JPY)", "リンクのクリック", "クリック(すべて)",
               "購入", "購入のコンバージョン値", "アトリビューション設定"]]
    campaign_reach: Dict[str, int] = {}
    for (cname, indicator, *_rest) in CAMPAIGNS:
        t = camp_tot[cname]
        reach = round(t["reach_daily"] * u(0.40, 0.50))
        campaign_reach[cname] = reach
        results = t["purchases"] if indicator == PURCHASE_IND else reach
        attr = _rest[-2]
        period.append([fs, ls, cname, blank_if_zero(results), indicator, fmt_num(reach),
                       fmt_num(t["impressions"] / reach, 2), fmt_num(t["impressions"]), fmt_num(t["spend"]),
                       fmt_num(t["link_clicks"]), fmt_num(t["clicks_all"]), blank_if_zero(t["purchases"]),
                       blank_if_zero(t["purchase_value"]), attr])
    files["期間合計_キャンペーン.csv"] = period

    def breakdown(dim_cols: List[str], buckets: List[Tuple]) -> List[List[str]]:
        rows = [["レポート開始日", "レポート終了日", "キャンペーン名"] + dim_cols +
                ["結果", "結果インジケーター", "リーチ", "インプレッション", "消化金額 (JPY)",
                 "リンクのクリック", "クリック(すべて)", "購入", "購入のコンバージョン値"]]
        ndim = len(dim_cols)
        for (cname, indicator, *_rest) in CAMPAIGNS:
            t = camp_tot[cname]
            ws = [b[ndim] * u(0.85, 1.15) for b in buckets]
            spend = split_int(int(t["spend"]), ws)
            imps = split_int(int(t["impressions"]), ws)
            link = split_int(int(t["link_clicks"]), [w * b[ndim + 1] for w, b in zip(ws, buckets)])
            call = split_int(int(t["clicks_all"]), [w * b[ndim + 1] for w, b in zip(ws, buckets)])
            cvw = [w * b[ndim + 1] * b[ndim + 2] for w, b in zip(ws, buckets)]
            pur = split_int(int(t["purchases"]), cvw)
            pval = split_int(int(t["purchase_value"]), cvw)
            reach = split_int(int(campaign_reach[cname] * 1.08), ws)
            for i, b in enumerate(buckets):
                if spend[i] == 0 and imps[i] == 0:
                    continue
                results = pur[i] if indicator == PURCHASE_IND else reach[i]
                rows.append([fs, ls, cname] + [str(x) for x in b[:ndim]] +
                            [blank_if_zero(results), indicator, fmt_num(reach[i]), fmt_num(imps[i]),
                             fmt_num(spend[i]), blank_if_zero(link[i]), blank_if_zero(call[i]),
                             blank_if_zero(pur[i]), blank_if_zero(pval[i])])
        return rows

    files["内訳_年齢性別.csv"] = breakdown(["年齢", "性別"], AGE_GENDER)
    files["内訳_配置.csv"] = breakdown(["プラットフォーム", "配置"], PLACEMENTS)
    files["内訳_デバイス.csv"] = breakdown(["インプレッションデバイス"], DEVICES)
    files["内訳_地域.csv"] = breakdown(["地域"], REGIONS)
    return files


def write_month(out_root: Path, month: str) -> List[Path]:
    out_dir = out_root / month
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for name, rows in gen_month(month).items():
        p = out_dir / name
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, lineterminator="\n")
            w.writerows(rows)
        written.append(p)
    return written


def main(argv=None) -> int:
    default_out = Path(__file__).resolve().parents[3] / "99_samples" / "ads" / CLIENT
    ap = argparse.ArgumentParser(description="架空の Meta 広告エクスポート（CSV）を作る（乱数シード固定）")
    ap.add_argument("--out", default=str(default_out), help="出力先（この下に YYYY-MM/ を作る）")
    ap.add_argument("--months", nargs="+", default=DEFAULT_MONTHS, help="作る月（YYYY-MM）")
    args = ap.parse_args(argv)
    out = Path(args.out).expanduser()
    for month in args.months:
        for p in write_month(out, month):
            print(f"  {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
