#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Taiwan Lottery Analysis V2 updater.

- Fetches latest official draw data from Taiwan Lottery API used by the official site.
- Recomputes strategy recommendations.
- Adds rolling backtest vs deterministic random baseline.
- Generates a redesigned standalone HTML dashboard and a lightweight XLSX report.
"""
from __future__ import annotations

import itertools
import json
import math
import random
import time
import urllib.parse
import urllib.request
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from string import Template
from typing import Dict, List, Tuple
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]
HTML_PATH = ROOT / "台灣彩券號碼推薦器.html"
INDEX_PATH = ROOT / "index.html"
XLSX_PATH = ROOT / "台灣彩券分析報告.xlsx"
RAW_PATH = ROOT / "latest_official_draws.json"
AUDIT_PATH = ROOT / "演算法審查與UI優化建議.md"
API_BASE = "https://api.taiwanlottery.com/TLCAPIWeB"
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X) AppleWebKit/537.36 Safari/537.36"

@dataclass(frozen=True)
class GameSpec:
    key: str
    title: str
    endpoint: str
    response_key: str
    max_num: int
    special_max: int
    special_label: str
    draw_count: int = 96
    train_window: int = 50

GAMES = [
    GameSpec("lotto649", "大樂透", "Lotto649Result", "lotto649Res", 49, 49, "特別號"),
    GameSpec("superlotto", "威力彩", "SuperLotto638Result", "superLotto638Res", 38, 8, "第二區"),
]

STRATEGY_INFO = {
    "回測最佳": "依照最近 rolling backtest 表現，自動選擇平均命中較高的策略；若差距太小，視為沒有明顯優勢。",
    "均衡分散": "使用頻率、近期趨勢、遺漏期數三個訊號，但強制奇偶、高低區分散，避免全部押同一種型態。",
    "熱號觀察": "只看樣本內出現次數較高的號碼。這是觀察，不代表下一期更容易出現。",
    "冷號觀察": "只看較久未出的號碼。這容易落入賭徒謬誤，因此只作對照。",
    "近期趨勢": "偏重近 15 期出現較多的號碼。樣本很小，容易受短期噪音影響。",
    "冷熱混合": "混合熱號與冷號，讓組合比較分散，但不代表期望值提升。",
    "隨機基準": "用固定種子產生可重現的隨機組合，作為比較基準。",
}


def months_back(n: int = 18) -> List[str]:
    y, m = date.today().year, date.today().month
    out = []
    for _ in range(n):
        out.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            y -= 1
            m = 12
    return out


def api_get(path: str, params: dict) -> dict:
    url = f"{API_BASE}{path}?{urllib.parse.urlencode(params)}"
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Referer": "https://www.taiwanlottery.com/"})
            with urllib.request.urlopen(req, timeout=30) as r:
                payload = json.loads(r.read().decode("utf-8"))
            if payload.get("rtCode") != 0:
                raise RuntimeError(f"API failed: {payload.get('rtCode')} {payload.get('rtMsg')} {url}")
            content = payload.get("content")
            if not isinstance(content, dict):
                raise RuntimeError(f"API returned invalid content: {url}")
            return content
        except Exception as exc:
            last_error = exc
            if attempt < 2:
                time.sleep(2 ** attempt)
    raise RuntimeError(f"Taiwan Lottery API failed after 3 attempts: {url}") from last_error


def fetch_draws(spec: GameSpec) -> List[dict]:
    rows, seen = [], set()
    for ym in months_back():
        content = api_get(f"/Lottery/{spec.endpoint}", {"month": ym, "endMonth": ym, "pageNum": 1, "pageSize": 200})
        for row in content.get(spec.response_key, []):
            period = int(row["period"])
            if period in seen:
                continue
            seen.add(period)
            nums = [int(x) for x in row["drawNumberSize"]]
            rows.append({"period": period, "date": row["lotteryDate"][:10], "numbers": nums[:6], "special": nums[6]})
        if len(rows) >= spec.draw_count:
            break
    rows.sort(key=lambda x: x["period"], reverse=True)
    rows = rows[: spec.draw_count]
    if len(rows) < spec.draw_count:
        raise RuntimeError(f"{spec.title} draw data incomplete: expected {spec.draw_count}, got {len(rows)}")
    for row in rows:
        numbers = row["numbers"]
        if len(numbers) != 6 or len(set(numbers)) != 6:
            raise RuntimeError(f"{spec.title} period {row['period']} has invalid main numbers")
        if not all(1 <= n <= spec.max_num for n in numbers):
            raise RuntimeError(f"{spec.title} period {row['period']} has out-of-range main numbers")
        if not 1 <= row["special"] <= spec.special_max:
            raise RuntimeError(f"{spec.title} period {row['period']} has invalid {spec.special_label}")
        try:
            date.fromisoformat(row["date"])
        except ValueError as exc:
            raise RuntimeError(f"{spec.title} period {row['period']} has invalid date") from exc
    return rows


def score_table(spec: GameSpec, train: List[dict]) -> Dict[str, dict]:
    freq = Counter(n for d in train for n in d["numbers"])
    recent_n = min(15, len(train))
    recent = Counter(n for d in train[:recent_n] for n in d["numbers"])
    misses = {}
    for n in range(1, spec.max_num + 1):
        misses[n] = next((i for i, d in enumerate(train) if n in d["numbers"]), len(train))
    max_f = max(freq.get(n, 0) for n in range(1, spec.max_num + 1)) or 1
    max_m = max(misses.values()) or 1
    max_r = max(recent.get(n, 0) for n in range(1, spec.max_num + 1)) or 1
    avg_f = sum(freq.values()) / spec.max_num
    out = {}
    for n in range(1, spec.max_num + 1):
        f = freq.get(n, 0)
        r = recent.get(n, 0)
        miss = misses[n]
        freq_s = round(f / max_f * 100, 2)
        trend_s = round(r / max_r * 100, 2)
        overdue_s = round(miss / max_m * 100, 2)
        # V2 ranking score: balanced display score, not probability.
        # It does not invert overdue aggressively like V1; overdue is only a mild diversity signal.
        stability = 100 - min(100, abs(f - avg_f) / max(avg_f, 1) * 100)
        balanced_s = round(freq_s * 0.30 + trend_s * 0.25 + overdue_s * 0.15 + stability * 0.15 + 15, 2)
        out[str(n)] = {
            "f": f,
            "fp": round(f / len(train) * 100, 2),
            "recent": r,
            "m": miss,
            "fs": freq_s,
            "ts": trend_s,
            "os": overdue_s,
            "stability": round(stability, 2),
            "balanced": min(100, balanced_s),
        }
    return out


def choose_spread(candidates: List[int], max_num: int, scores: Dict[str, dict]) -> List[int]:
    """Pick 6 from ranked candidates with basic odd/even and low/high diversity."""
    chosen: List[int] = []
    low_cut = max_num // 2
    for n in candidates:
        if len(chosen) >= 6:
            break
        trial = chosen + [n]
        odd = sum(x % 2 for x in trial)
        low = sum(x <= low_cut for x in trial)
        remaining = 6 - len(trial)
        if odd > 4 or odd + remaining < 2:
            continue
        if low > 4 or low + remaining < 2:
            continue
        chosen.append(n)
    for n in candidates:
        if len(chosen) >= 6:
            break
        if n not in chosen:
            chosen.append(n)
    return sorted(chosen[:6])


def deterministic_random(spec: GameSpec, seed: int) -> List[int]:
    rng = random.Random(seed)
    return sorted(rng.sample(range(1, spec.max_num + 1), 6))


def recommendations(spec: GameSpec, train: List[dict], period_seed: int | None = None) -> Tuple[Dict[str, List[int]], Dict[str, dict]]:
    scores = score_table(spec, train)
    items = [(n, scores[str(n)]) for n in range(1, spec.max_num + 1)]
    hot_rank = [n for n, _ in sorted(items, key=lambda x: (-x[1]["f"], -x[1]["fs"], x[0]))]
    cold_rank = [n for n, _ in sorted(items, key=lambda x: (-x[1]["m"], -x[1]["stability"], x[0]))]
    trend_rank = [n for n, _ in sorted(items, key=lambda x: (-x[1]["recent"], -x[1]["balanced"], x[0]))]
    balanced_rank = [n for n, _ in sorted(items, key=lambda x: (-x[1]["balanced"], x[0]))]
    mix_rank = []
    for n in hot_rank[:12] + cold_rank[:12] + trend_rank[:12] + balanced_rank[:12]:
        if n not in mix_rank:
            mix_rank.append(n)
    recs = {
        "均衡分散": choose_spread(balanced_rank, spec.max_num, scores),
        "熱號觀察": sorted(hot_rank[:6]),
        "冷號觀察": sorted(cold_rank[:6]),
        "近期趨勢": sorted(trend_rank[:6]),
        "冷熱混合": choose_spread(mix_rank, spec.max_num, scores),
        "隨機基準": deterministic_random(spec, period_seed or train[0]["period"]),
    }
    return recs, scores


def backtest(spec: GameSpec, draws: List[dict]) -> dict:
    strategies = ["均衡分散", "熱號觀察", "冷號觀察", "近期趨勢", "冷熱混合", "隨機基準"]
    rows = {s: [] for s in strategies}
    limit = min(42, len(draws) - spec.train_window - 1)
    if limit < 8:
        limit = max(0, len(draws) - spec.train_window - 1)
    for i in range(limit):
        target = draws[i]
        train = draws[i + 1 : i + 1 + spec.train_window]
        recs, _ = recommendations(spec, train, target["period"])
        actual = set(target["numbers"])
        for s in strategies:
            pick = recs[s]
            rows[s].append({"period": target["period"], "date": target["date"], "pick": pick, "hit": len(actual & set(pick))})
    summary = []
    for s in strategies:
        hits = [r["hit"] for r in rows[s]]
        n = len(hits) or 1
        summary.append({
            "strategy": s,
            "tests": len(hits),
            "avgHit": round(sum(hits) / n, 3),
            "hit2": round(sum(h >= 2 for h in hits) / n * 100, 1),
            "hit3": round(sum(h >= 3 for h in hits) / n * 100, 1),
            "best": max(hits) if hits else 0,
        })
    summary.sort(key=lambda x: (-x["avgHit"], -x["hit3"], x["strategy"] != "隨機基準"))
    return {"window": spec.train_window, "tests": limit, "summary": summary, "rows": rows}


def compute_game(spec: GameSpec, draws: List[dict]) -> dict:
    train = draws[: spec.train_window]
    recs, scores = recommendations(spec, train, draws[0]["period"])
    bt = backtest(spec, draws)
    best = bt["summary"][0]["strategy"] if bt["summary"] else "均衡分散"
    # If best is only barely above random, keep default balanced and say evidence is weak.
    rand = next((x for x in bt["summary"] if x["strategy"] == "隨機基準"), None)
    best_obj = bt["summary"][0] if bt["summary"] else None
    edge_note = "回測樣本不足"
    if rand and best_obj:
        diff = round(best_obj["avgHit"] - rand["avgHit"], 3)
        edge_note = f"回測最佳比隨機基準平均多 {diff} 碼/期" if diff > 0 else "沒有明顯勝過隨機基準"
        if diff < 0.08:
            best = "均衡分散"
    recs["回測最佳"] = recs.get(best, recs["均衡分散"])
    sp_count = Counter(d["special"] for d in train)
    sp_sorted = sorted(sp_count.items(), key=lambda kv: (-kv[1], kv[0]))
    cost = {str(k): math.comb(k, 6) for k in range(6, 11)}
    return {
        "title": spec.title,
        "maxNum": spec.max_num,
        "specialLabel": spec.special_label,
        "date": draws[0]["date"],
        "period": draws[0]["period"],
        "draws": len(train),
        "historyTotal": len(draws),
        "last": draws[0],
        "scores": scores,
        "specialStats": {str(k): v for k, v in sp_sorted},
        "specialRecommend": sp_sorted[0][0] if sp_sorted else None,
        "recs": recs,
        "strategyInfo": STRATEGY_INFO,
        "backtest": bt,
        "edgeNote": edge_note,
        "cost": cost,
        "history": draws[: spec.train_window],
    }


def render_html(data: dict) -> str:
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    template = Template(r'''<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#090d18">
<meta name="description" content="大樂透與威力彩開獎資料、策略比較、rolling backtest 與號碼排序訊號。僅供娛樂與統計觀察。">
<meta property="og:title" content="台灣彩券策略儀表板">
<meta property="og:description" content="大樂透與威力彩策略比較、回測與最新開獎資料。">
<meta property="og:type" content="website">
<meta property="og:url" content="https://samyiqrs.github.io/taiwan-lottery-dashboard/">
<link rel="canonical" href="https://samyiqrs.github.io/taiwan-lottery-dashboard/">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='16' fill='%23ffd166'/%3E%3Ctext x='32' y='44' text-anchor='middle' font-size='38'%3E%E5%BD%A9%3C/text%3E%3C/svg%3E">
<title>台灣彩券策略儀表板</title>
<style>
:root{color-scheme:dark;--bg:#070a12;--panel:#111827;--panel2:#172033;--line:rgba(255,255,255,.11);--text:#f4f7ff;--muted:#9ba8c2;--gold:#ffd166;--blue:#5dd6ff;--green:#80ed99;--purple:#c69cff;--shadow:0 18px 60px rgba(0,0,0,.26)}
*{box-sizing:border-box}
.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}
html{scroll-behavior:smooth;-webkit-text-size-adjust:100%}
body{margin:0;min-width:320px;overflow-x:hidden;background:radial-gradient(circle at 18% 0%,#20335f 0,#090d18 36%,var(--bg) 100%);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,'PingFang TC','Microsoft JhengHei','Noto Sans TC',Segoe UI,sans-serif}
button,a{font:inherit}.wrap{width:min(100%,1180px);margin:0 auto;padding:28px max(18px,env(safe-area-inset-left)) calc(44px + env(safe-area-inset-bottom))}
.topbar{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:16px}.brand{display:flex;align-items:center;gap:10px;font-weight:900}.brand-mark{display:grid;place-items:center;width:36px;height:36px;border-radius:12px;background:linear-gradient(145deg,var(--gold),#ff9f1c);color:#111827;box-shadow:0 8px 24px rgba(255,159,28,.2)}
.top-actions{display:flex;align-items:center;gap:8px}.link-btn{display:inline-flex;align-items:center;justify-content:center;min-height:44px;padding:9px 13px;border:1px solid var(--line);border-radius:12px;color:var(--text);text-decoration:none;background:rgba(255,255,255,.04)}
.hero{display:grid;grid-template-columns:minmax(0,1.3fr) minmax(310px,.7fr);gap:18px;align-items:stretch}.card,.hero-main,.hero-side{min-width:0;background:linear-gradient(180deg,rgba(255,255,255,.075),rgba(255,255,255,.035));border:1px solid var(--line);box-shadow:var(--shadow);border-radius:24px}.hero-main{padding:clamp(22px,4vw,32px)}.hero-side{padding:22px;display:flex;flex-direction:column}.hero-side .btn{margin-top:auto}
h1{margin:0 0 10px;font-size:clamp(28px,4.3vw,42px);line-height:1.12;letter-spacing:.2px}h2{font-size:19px;margin:0}.subtitle{max-width:760px;color:var(--muted);line-height:1.75;margin:0}.pill{display:inline-flex;align-items:center;gap:8px;border:1px solid rgba(255,209,102,.4);color:#ffe2a1;background:rgba(255,209,102,.08);border-radius:999px;padding:8px 12px;font-size:13px;margin-bottom:15px}
.tabs{display:flex;gap:10px;margin:20px 0 16px}.tab,.strategy,.btn{min-height:44px;touch-action:manipulation}.tab{border:1px solid var(--line);background:#0e1526;color:var(--text);border-radius:999px;padding:11px 20px;cursor:pointer;font-weight:800}.tab.active{background:linear-gradient(135deg,var(--gold),#ff9f1c);color:#111827;border-color:transparent}.tab:focus-visible,.strategy:focus-visible,.btn:focus-visible,.link-btn:focus-visible,.table-wrap:focus-visible{outline:3px solid var(--blue);outline-offset:3px}
.grid{display:grid;gap:14px}.stats{grid-template-columns:repeat(4,minmax(0,1fr));margin:18px 0}.stat{min-width:0;padding:16px;border-radius:18px;background:rgba(255,255,255,.055);border:1px solid var(--line)}.k{font-size:12px;color:var(--muted);margin-bottom:7px}.v{font-size:clamp(15px,2vw,18px);font-weight:900;overflow-wrap:anywhere}.main{display:grid;grid-template-columns:minmax(0,.95fr) minmax(0,1.05fr);gap:16px;margin-top:16px}.card{padding:20px}.section-title{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:14px}.small{font-size:12px;color:var(--muted);line-height:1.6}.numbers{display:flex;gap:9px;flex-wrap:wrap;margin:14px 0}.ball{width:46px;height:46px;flex:0 0 46px;border-radius:50%;display:grid;place-items:center;font-weight:900;font-size:18px;color:#0f172a;background:linear-gradient(145deg,#fff4bf,#ffd166);box-shadow:0 8px 22px rgba(255,209,102,.18)}.ball.special{background:linear-gradient(145deg,#c8f5ff,#5dd6ff)}
.strategies{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:9px}.strategy{border:1px solid var(--line);background:#0d1424;color:var(--text);border-radius:14px;padding:11px 12px;cursor:pointer;text-align:left}.strategy.active{border-color:rgba(93,214,255,.75);box-shadow:0 0 0 2px rgba(93,214,255,.12) inset;background:rgba(93,214,255,.08)}.badge{flex:0 0 auto;font-size:11px;padding:5px 8px;border-radius:999px;background:rgba(128,237,153,.12);color:var(--green);border:1px solid rgba(128,237,153,.25)}.warn{font-size:12px;line-height:1.65;color:#ffddb0;background:rgba(255,159,28,.08);border:1px solid rgba(255,159,28,.2);padding:11px 12px;border-radius:14px}.btn{width:100%;border:0;border-radius:14px;padding:12px 15px;font-weight:900;cursor:pointer}.btn.primary{background:linear-gradient(135deg,var(--blue),#6c8cff);color:#07111f}.copy-status{min-height:20px;margin-top:8px;text-align:center}
.mobile-hint{display:none}.table-wrap{width:100%;overflow-x:auto;overscroll-behavior-inline:contain;border:1px solid var(--line);border-radius:15px;-webkit-overflow-scrolling:touch}table{width:100%;border-collapse:collapse;min-width:610px}th,td{padding:11px 10px;border-bottom:1px solid rgba(255,255,255,.07);font-size:12px;text-align:right;white-space:nowrap}th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){text-align:left}th{position:sticky;top:0;background:#172033;color:#c6d0e4}.rank1{color:var(--gold);font-weight:900}.meter{width:92px;height:7px;background:#202a3f;border-radius:999px;overflow:hidden}.meter i{display:block;height:100%;background:linear-gradient(90deg,var(--purple),var(--blue))}.footer-grid{grid-template-columns:minmax(0,.65fr) minmax(0,1.35fr);align-items:start;margin-top:16px}.cost{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.cost>div{padding:13px;border-radius:14px;background:rgba(255,255,255,.05);display:flex;gap:8px;justify-content:space-between}.cost b{color:var(--gold)}.cost small{color:var(--muted)}
.site-footer{margin-top:18px;padding:18px 4px 0;color:var(--muted);font-size:12px;line-height:1.7;text-align:center}.site-footer a{color:var(--blue)}
@media(max-width:900px){.hero,.main,.footer-grid{grid-template-columns:1fr}.hero-side{min-height:280px}.stats{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:600px){.wrap{padding-top:14px;padding-right:max(12px,env(safe-area-inset-right));padding-left:max(12px,env(safe-area-inset-left))}.topbar{align-items:flex-start}.top-actions{width:100%}.link-btn{flex:1}.hero-main,.hero-side,.card{padding:16px;border-radius:18px}.hero-side{min-height:0}.tabs{display:grid;grid-template-columns:1fr 1fr}.tab{width:100%;padding-inline:10px}.stats{gap:9px}.stat{padding:13px}.strategies{grid-template-columns:1fr 1fr}.strategy{text-align:center;padding-inline:7px}.ball{width:42px;height:42px;flex-basis:42px;font-size:17px}.section-title{align-items:flex-start}.footer-grid{gap:12px}.mobile-hint{display:block;margin:-4px 0 10px;color:var(--muted);font-size:12px;text-align:right}.table-wrap{margin-inline:0}.small,.warn,th,td{font-size:13px}}
@media(max-width:390px){.topbar{display:block}.top-actions{margin-top:10px}.strategies{grid-template-columns:1fr}.stats{grid-template-columns:1fr 1fr}.numbers{gap:7px}.ball{width:39px;height:39px;flex-basis:39px;font-size:16px}.cost{grid-template-columns:1fr}}
@media(prefers-reduced-motion:reduce){html{scroll-behavior:auto}*,*::before,*::after{animation-duration:.01ms!important;transition-duration:.01ms!important}}
</style>
</head>
<body>
<div class="wrap">
<header class="topbar"><div class="brand"><span class="brand-mark" aria-hidden="true">彩</span><span>台灣彩券策略儀表板</span></div><nav class="top-actions" aria-label="下載與說明"><a class="link-btn" href="台灣彩券分析報告.xlsx" download>下載分析報告</a></nav></header>
<main>
<section class="hero"><div class="hero-main"><div class="pill">V2｜回測優先，不假裝神準</div><h1>大樂透與威力彩<br>策略分析</h1><p class="subtitle">官方開獎資料自動更新。推薦邏輯採用策略比較、rolling backtest 與隨機基準；分數是排序訊號，不是中獎機率。</p><div class="tabs" aria-label="彩券種類"><button type="button" class="tab active" aria-pressed="true" data-game="lotto649">大樂透</button><button type="button" class="tab" aria-pressed="false" data-game="superlotto">威力彩</button></div><div class="warn">理性提醒：彩券開獎近似獨立隨機事件。若回測沒有穩定勝過隨機基準，就應視為娛樂參考，而不是預測優勢。</div></div><div class="hero-side"><h2>今日結論</h2><p class="small" id="todayConclusion" aria-live="polite"></p><div class="numbers" id="heroNums" aria-label="推薦號碼"></div><button type="button" class="btn primary" id="copyBtn">複製推薦號碼</button><div class="small copy-status" id="copyStatus" aria-live="polite"></div></div></section>
<div class="grid stats"><div class="stat"><div class="k">最新期號</div><div class="v" id="period"></div></div><div class="stat"><div class="k">最新開獎日</div><div class="v" id="date"></div></div><div class="stat"><div class="k">回測期數</div><div class="v" id="tests"></div></div><div class="stat"><div class="k">回測判斷</div><div class="v" id="edge"></div></div></div>
<section class="main"><div class="card"><div class="section-title"><h2>策略選擇</h2><span class="badge" id="currentStrategyBadge"></span></div><div class="strategies" id="strategies"></div><div class="warn" style="margin-top:14px" id="strategyDesc"></div><div class="section-title" style="margin-top:18px"><h2>最新開獎</h2></div><div class="numbers" id="lastNums" aria-label="最新開獎號碼"></div></div><div class="card"><div class="section-title"><h2>Rolling Backtest 策略比較</h2><span class="small">前 50 期預測下一期</span></div><span class="mobile-hint">左右滑動查看完整資料 →</span><div class="table-wrap" tabindex="0" aria-label="可橫向捲動的回測比較表"><table id="backtestTable"></table></div><p class="small" style="margin-top:10px">比較平均命中碼數、2 碼以上比例與 3 碼以上比例；若只小幅勝過隨機，不視為可靠優勢。</p></div></section>
<section class="grid footer-grid"><div class="card"><div class="section-title"><h2>包牌成本提醒</h2><span class="small">組合數，不含每注金額</span></div><div class="cost" id="cost"></div><p class="small">包牌會增加覆蓋率，也會等比例增加成本；不會改善單注期望值。</p></div><div class="card"><div class="section-title"><h2>號碼排序訊號</h2><span class="small">非機率</span></div><span class="mobile-hint">左右滑動查看完整資料 →</span><div class="table-wrap" tabindex="0" aria-label="可橫向捲動的號碼排序表"><table id="scoreTable"></table></div></div></section>
</main>
<footer class="site-footer">資料來源：台灣彩券官方公開資料。本站不隸屬於台灣彩券；內容僅供娛樂與統計觀察，請量力而為。</footer>
</div>
<script>const DATA=$DATA;
let game='lotto649';let strategy='回測最佳';let currentPick=[];
const order=['回測最佳','均衡分散','熱號觀察','冷號觀察','近期趨勢','冷熱混合','隨機基準'];
function $(id){return document.getElementById(id)}
function d(){return DATA[game]}
function balls(nums,special,label){let h='';nums.forEach(n=>h+=`<div class="ball" aria-label="號碼 ${n}">${n}</div>`);if(special)h+=`<div class="ball special" title="${label}" aria-label="${label} ${special}">${special}</div>`;return h}
function render(){let x=d();document.querySelectorAll('.tab').forEach(b=>{let active=b.dataset.game===game;b.classList.toggle('active',active);b.setAttribute('aria-pressed',String(active))});$('period').textContent=x.period;$('date').textContent=x.date;$('tests').textContent=x.backtest.tests+' 期';$('edge').textContent=x.edgeNote;$('todayConclusion').innerHTML=`${x.title} 目前採用 <b>${strategy}</b>。${x.edgeNote}。建議把這當作策略觀察，不是保證預測。`;currentPick=x.recs[strategy]||x.recs['均衡分散'];$('heroNums').innerHTML=balls(currentPick,x.specialRecommend,x.specialLabel);$('lastNums').innerHTML=balls(x.last.numbers,x.last.special,x.specialLabel);$('currentStrategyBadge').textContent=strategy;renderStrategies();renderBacktest();renderScore();renderCost();$('strategyDesc').textContent=x.strategyInfo[strategy]||'';document.title=`${x.title}｜台灣彩券策略儀表板`}
function renderStrategies(){$('strategies').innerHTML=order.map(s=>`<button type="button" class="strategy ${s===strategy?'active':''}" aria-pressed="${s===strategy}" data-strategy="${s}">${s}</button>`).join('');document.querySelectorAll('.strategy').forEach(b=>b.onclick=()=>{strategy=b.dataset.strategy;render()})}
function renderBacktest(){let rows=d().backtest.summary;let h='<caption class="sr-only">各策略 rolling backtest 比較</caption><thead><tr><th scope="col">排名</th><th scope="col">策略</th><th scope="col">平均命中</th><th scope="col">≥2碼</th><th scope="col">≥3碼</th><th scope="col">最佳</th></tr></thead><tbody>';rows.forEach((r,i)=>{h+=`<tr><td class="${i===0?'rank1':''}">${i+1}</td><td>${r.strategy}</td><td>${r.avgHit}</td><td>${r.hit2}%</td><td>${r.hit3}%</td><td>${r.best}</td></tr>`});$('backtestTable').innerHTML=h+'</tbody>'}
function renderScore(){let rows=Object.entries(d().scores).sort((a,b)=>b[1].balanced-a[1].balanced).slice(0,18);let h='<caption class="sr-only">號碼統計排序訊號</caption><thead><tr><th scope="col">排名</th><th scope="col">號碼</th><th scope="col">排序分</th><th scope="col">次數</th><th scope="col">近15期</th><th scope="col">遺漏</th><th scope="col">視覺</th></tr></thead><tbody>';rows.forEach(([n,s],i)=>{h+=`<tr><td>${i+1}</td><td><b>${n}</b></td><td>${s.balanced}</td><td>${s.f}</td><td>${s.recent}</td><td>${s.m}</td><td><div class="meter" aria-hidden="true"><i style="width:${s.balanced}%"></i></div></td></tr>`});$('scoreTable').innerHTML=h+'</tbody>'}
function renderCost(){let x=d(); $('cost').innerHTML=Object.entries(x.cost).map(([k,v])=>`<div><span>${k} 碼</span><b>${v}</b><small>注</small></div>`).join('')}
async function copyPick(){const text=`${d().title} ${strategy}: ${currentPick.join(', ')}｜${d().specialLabel} ${d().specialRecommend}`;let copied=false;try{await navigator.clipboard.writeText(text);copied=true}catch(e){const area=document.createElement('textarea');area.value=text;area.style.position='fixed';area.style.opacity='0';document.body.appendChild(area);area.select();copied=document.execCommand('copy');area.remove()}$('copyStatus').textContent=copied?'已複製到剪貼簿':'無法自動複製，請長按號碼手動選取';setTimeout(()=>{$('copyStatus').textContent=''},copied?2200:4500)}
document.querySelectorAll('.tab').forEach(b=>b.onclick=()=>{game=b.dataset.game;strategy='回測最佳';render()});$('copyBtn').onclick=copyPick;render();</script>
</body></html>''')
    return template.template.replace("$DATA", data_json)


def xcol(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet_xml(rows: List[List[object]]) -> str:
    lines = []
    for ri, row in enumerate(rows, 1):
        cells = []
        for ci, val in enumerate(row, 1):
            ref = f"{xcol(ci)}{ri}"
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cells.append(f'<c r="{ref}"><v>{val}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(val))}</t></is></c>')
        lines.append(f'<row r="{ri}">' + ''.join(cells) + '</row>')
    return '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + ''.join(lines) + '</sheetData></worksheet>'


def data_basis(data: dict) -> str:
    return "；".join(f"{data[spec.key]['title']} {data[spec.key]['date']} 第 {data[spec.key]['period']} 期" for spec in GAMES)


def stable_writestr(z: zipfile.ZipFile, filename: str, content: str) -> None:
    info = zipfile.ZipInfo(filename, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    z.writestr(info, content.encode("utf-8"))


def write_xlsx(data: dict) -> None:
    sheets: List[Tuple[str, List[List[object]]]] = []
    overview = [["台灣彩券策略儀表板 V2"], ["資料基準", data_basis(data)]]
    for spec in GAMES:
        g = data[spec.key]
        overview += [[g["title"], "最新期號", g["period"], "開獎日", g["date"], "預設策略", "回測最佳"], ["推薦", ", ".join(map(str, g["recs"]["回測最佳"])), g["specialLabel"], g["specialRecommend"], "回測判斷", g["edgeNote"]]]
    sheets.append(("總覽", overview))
    for spec in GAMES:
        g = data[spec.key]
        sheets.append((f"{g['title']}回測", [["策略", "測試期數", "平均命中", "≥2碼%", "≥3碼%", "最佳"]] + [[r["strategy"], r["tests"], r["avgHit"], r["hit2"], r["hit3"], r["best"]] for r in g["backtest"]["summary"]]))
        rows = [["排名", "號碼", "排序分", "出現次數", "近15期", "遺漏", "頻率分", "趨勢分", "遺漏訊號"]]
        for i, (n, s) in enumerate(sorted(g["scores"].items(), key=lambda kv: kv[1]["balanced"], reverse=True), 1):
            rows.append([i, int(n), s["balanced"], s["f"], s["recent"], s["m"], s["fs"], s["ts"], s["os"]])
        sheets.append((f"{g['title']}排序", rows))
        hist = [["期號", "日期", "N1", "N2", "N3", "N4", "N5", "N6", g["specialLabel"]]]
        for row in g["history"]:
            hist.append([row["period"], row["date"], *row["numbers"], row["special"]])
        sheets.append((f"{g['title']}歷史", hist))
    with zipfile.ZipFile(XLSX_PATH, "w", zipfile.ZIP_DEFLATED) as z:
        stable_writestr(z, "_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        types = ['<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
        rels = ['<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        wb = []
        for i, (name, rows) in enumerate(sheets, 1):
            stable_writestr(z, f"xl/worksheets/sheet{i}.xml", sheet_xml(rows))
            wb.append(f'<sheet name="{escape(name[:31])}" sheetId="{i}" r:id="rId{i}"/>')
            rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
            types.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        stable_writestr(z, "[Content_Types].xml", ''.join(types) + '</Types>')
        stable_writestr(z, "xl/_rels/workbook.xml.rels", ''.join(rels) + '</Relationships>')
        stable_writestr(z, "xl/workbook.xml", '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + ''.join(wb) + '</sheets></workbook>')


def write_audit(data: dict) -> None:
    lines = ["# 台灣彩券分析套件 V2：演算法與 UI 審查", "", f"資料基準：{data_basis(data)}", "", "## 本次 V2 實際修改", "", "1. 移除原本『AI 深度預測』式定位，改成策略儀表板。", "2. 推薦邏輯不再使用 V1 的『近期未遺漏加分』作為核心公式。", "3. 新增 rolling backtest，用前 50 期預測下一期，並與固定種子的隨機基準比較。", "4. 預設『回測最佳』若只小幅勝過隨機，會回退到均衡分散，不假裝有優勢。", "5. UI 重做為儀表板：今日結論、推薦號碼、最新開獎、回測比較、包牌成本、排序訊號。", "", "## 仍須保守解讀", "", "彩券開獎近似獨立隨機。即使某策略在短期回測略勝，也可能只是噪音；工具只能當娛樂與策略觀察。", ""]
    for spec in GAMES:
        g = data[spec.key]
        lines += [f"## {g['title']}", "", f"- 最新期號：{g['period']}", f"- 開獎日：{g['date']}", f"- 預設推薦：{', '.join(map(str, g['recs']['回測最佳']))}｜{g['specialLabel']} {g['specialRecommend']}", f"- 回測判斷：{g['edgeNote']}", "", "### 回測摘要", "", "| 策略 | 平均命中 | ≥2碼 | ≥3碼 | 最佳 |", "|---|---:|---:|---:|---:|"]
        for r in g["backtest"]["summary"]:
            lines.append(f"| {r['strategy']} | {r['avgHit']} | {r['hit2']}% | {r['hit3']}% | {r['best']} |")
        lines.append("")
    AUDIT_PATH.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    data = {}
    for spec in GAMES:
        data[spec.key] = compute_game(spec, fetch_draws(spec))
    RAW_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    html = render_html(data)
    HTML_PATH.write_text(html, encoding="utf-8")
    INDEX_PATH.write_text(html, encoding="utf-8")
    write_xlsx(data)
    write_audit(data)
    print(f"V2 更新完成：{datetime.now():%Y-%m-%d %H:%M:%S}")
    for spec in GAMES:
        g = data[spec.key]
        print(f"{g['title']}: {g['period']} {g['date']}｜回測最佳 {g['recs']['回測最佳']}｜{g['specialLabel']} {g['specialRecommend']}｜{g['edgeNote']}")
    print(f"HTML: {HTML_PATH}")
    print(f"Web: {INDEX_PATH}")
    print(f"XLSX: {XLSX_PATH}")
    print(f"Audit: {AUDIT_PATH}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
