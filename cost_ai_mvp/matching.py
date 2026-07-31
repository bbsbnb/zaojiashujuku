from __future__ import annotations

import re
import statistics
from typing import Any

from rapidfuzz import fuzz

from .utils import normalize_name, normalize_unit


REGION_PRIORITY = ["深圳", "广州", "东莞", "佛山"]


def search_prices(conn, text: str, region: str = "深圳", item_type: str | None = None, limit: int = 50) -> list[dict[str, Any]]:
    norm = normalize_name(text)
    params: list[Any] = []
    where = ["is_outlier = 0"]
    if item_type:
        where.append("item_type = ?")
        params.append(item_type)
    rows = [dict(r) for r in conn.execute(f"SELECT * FROM prices WHERE {' AND '.join(where)}", params).fetchall()]
    scored: list[dict[str, Any]] = []
    for row in rows:
        name = row.get("standard_name") or row.get("original_name") or ""
        spec = row.get("standard_spec") or row.get("original_spec") or ""
        candidate_text = f"{name} {spec}"
        score = fuzz.partial_ratio(norm, normalize_name(name))
        if norm in normalize_name(candidate_text):
            score = max(score, 85)
        if _has_spec_conflict(text, candidate_text):
            score -= 45
        score += _region_points(row.get("region"), region)
        score += _source_points(row.get("source_type"))
        if score >= 45:
            row["match_score"] = max(0, min(score, 100))
            scored.append(row)
    scored.sort(key=lambda r: (r["match_score"], r.get("price_date") or str(r.get("price_year") or "")), reverse=True)
    return scored[:limit]


def build_price_suggestion(rows: list[dict[str, Any]], region: str = "深圳", duration: str = "", payment_terms: str = "") -> dict[str, Any]:
    if not rows:
        return {
            "suggested_price": None,
            "reference_low": None,
            "reference_high": None,
            "confidence": "low",
            "risk_tags": ["NO_SAMPLE"],
            "risk_text": "未找到可靠参考价，建议人工询价。",
            "main_source_text": "无",
            "source_type": None,
        }
    clean = [r for r in rows if r.get("unit_price") is not None and float(r.get("match_score") or 100) >= 70]
    if not clean:
        return build_price_suggestion([], region, duration, payment_terms)

    prices = [float(r["unit_price"]) for r in clean]
    if len(prices) >= 5:
        sorted_prices = sorted(prices)
        low = sorted_prices[max(0, int(len(sorted_prices) * 0.2) - 1)]
        high = sorted_prices[min(len(sorted_prices) - 1, int(len(sorted_prices) * 0.8))]
    elif len(prices) >= 2:
        low, high = min(prices), max(prices)
    else:
        base = prices[0]
        pct = {"labor": 0.10, "material": 0.10, "machinery": 0.15, "equipment": 0.20}.get(clean[0].get("item_type"), 0.10)
        low, high = base * (1 - pct), base * (1 + pct)

    suggested = _weighted_median(clean)
    if any(k in (duration or "") for k in ["紧", "赶工", "短"]):
        suggested = max(suggested, (suggested + high) / 2)

    risk_tags: list[str] = []
    if len(clean) < 3:
        risk_tags.append("SAMPLE_LOW")
    if not any(r.get("region") == region for r in clean):
        risk_tags.append("REGION_DIFF")
    if any((r.get("is_tax_included") or "unknown") == "unknown" for r in clean[:3]):
        risk_tags.append("UNKNOWN_SCOPE")
    if any(k in (duration or "") for k in ["紧", "赶工", "短"]):
        risk_tags.append("TIGHT_SCHEDULE")
    if not any(r.get("source_type") == "inquiry" for r in clean):
        risk_tags.append("NO_RECENT_INQUIRY")

    confidence = "high" if len(clean) >= 3 and not {"REGION_DIFF", "UNKNOWN_SCOPE"}.intersection(risk_tags) else "medium"
    if len(clean) < 2 or "REGION_DIFF" in risk_tags:
        confidence = "low" if len(clean) < 2 else "medium"

    main_sources = []
    for r in clean[:3]:
        src = r.get("source_type") or "来源"
        area = r.get("region") or "未知地区"
        date = r.get("price_date") or r.get("price_year") or "未知时间"
        main_sources.append(f"{src}/{area}/{date}/{r.get('unit_price')}")
    return {
        "suggested_price": round(float(suggested), 2),
        "reference_low": round(float(low), 2),
        "reference_high": round(float(high), 2),
        "confidence": confidence,
        "risk_tags": risk_tags,
        "risk_text": _risk_text(risk_tags),
        "main_source_text": "；".join(main_sources),
        "source_type": clean[0].get("source_type"),
        "main_source_price_id": clean[0].get("id"),
    }


def candidates_for_item(conn, item: dict[str, Any], region: str = "深圳") -> tuple[list[dict[str, Any]], dict[str, Any]]:
    text = f"{item.get('standard_name') or item.get('original_name')} {item.get('standard_spec') or item.get('original_spec') or ''}"
    all_rows = search_prices(conn, text, region=region, item_type=item.get("item_type"), limit=100)
    unit = normalize_unit(item.get("original_unit"))
    filtered = []
    for row in all_rows:
        row_unit = normalize_unit(row.get("standard_unit") or row.get("original_unit"))
        score = float(row.get("match_score") or 0)
        if row_unit == unit:
            score += 20
        else:
            score -= 30
        row["match_score"] = max(0, min(100, score))
        if row["match_score"] >= 50:
            filtered.append(row)
    filtered.sort(key=lambda r: r["match_score"], reverse=True)
    top = filtered[:3]
    suggestion = build_price_suggestion(filtered[:10], region=region)
    candidates = []
    for idx, row in enumerate(top, start=1):
        cand_suggestion = build_price_suggestion([row], region=region)
        candidates.append(
            {
                "rank_no": idx,
                "source_price_id": row.get("id"),
                "source_quote_item_id": row.get("quote_item_id"),
                "standard_item_id": row.get("standard_item_id"),
                "candidate_name": row.get("standard_name") or row.get("original_name"),
                "candidate_spec": row.get("standard_spec") or row.get("original_spec") or "",
                "candidate_unit": row.get("standard_unit") or row.get("original_unit"),
                "candidate_price": row.get("unit_price"),
                "reference_low": cand_suggestion["reference_low"],
                "reference_high": cand_suggestion["reference_high"],
                "source_type": row.get("source_type"),
                "source_text": f"{row.get('region')}/{row.get('price_date') or row.get('price_year')}/{row.get('source_type')}",
                "source_region": row.get("region"),
                "source_date_or_year": str(row.get("price_date") or row.get("price_year") or ""),
                "match_score": row.get("match_score"),
                "confidence": _confidence(row.get("match_score", 0)),
                "match_reason": "名称/规格/单位/地区综合匹配",
                "risk_tags": ";".join(cand_suggestion["risk_tags"]),
                "risk_text": cand_suggestion["risk_text"],
            }
        )
    return candidates, suggestion


def explanation(item_name: str, suggestion: dict[str, Any]) -> str:
    if suggestion.get("suggested_price") is None:
        return f"{item_name} 未找到可靠参考价，建议补充近期询价后再采用。"
    return (
        f"{item_name} 建议价为 {suggestion['suggested_price']}，参考区间为 "
        f"{suggestion['reference_low']}-{suggestion['reference_high']}。主要参考："
        f"{suggestion.get('main_source_text') or '历史价格库'}。风险提示：{suggestion.get('risk_text') or '无明显风险'}。"
    )


def _weighted_median(rows: list[dict[str, Any]]) -> float:
    pairs = []
    for r in rows:
        weight = 1.0 + _region_points(r.get("region"), "深圳") / 20 + _source_points(r.get("source_type")) / 20
        pairs.append((float(r["unit_price"]), weight))
    pairs.sort()
    total = sum(w for _, w in pairs)
    acc = 0.0
    for price, weight in pairs:
        acc += weight
        if acc >= total / 2:
            return price
    return statistics.median([p for p, _ in pairs])


def _region_points(row_region: str | None, target: str) -> float:
    if not row_region:
        return 0
    if row_region == target:
        return 10
    if row_region in REGION_PRIORITY:
        return max(2, 8 - REGION_PRIORITY.index(row_region) * 2)
    return 0


def _source_points(source: str | None) -> float:
    return {"inquiry": 5, "contract": 4.5, "settlement": 4.5, "material_price": 3.5, "resource_price": 3, "unit_price_analysis": 3, "tender": 1}.get(source or "", 0)


def _confidence(score: float) -> str:
    if score >= 85:
        return "high"
    if score >= 70:
        return "medium"
    return "low"


def _risk_text(tags: list[str]) -> str:
    mapping = {
        "NO_SAMPLE": "无可靠样本",
        "SAMPLE_LOW": "有效样本较少，建议复核",
        "REGION_DIFF": "主要参考非目标地区数据，存在地区差异",
        "UNKNOWN_SCOPE": "部分价格口径不明，需确认含税、运费或安装边界",
        "TIGHT_SCHEDULE": "当前工期偏紧，建议取区间中高值",
        "NO_RECENT_INQUIRY": "缺少近期询价支撑",
    }
    return "；".join(mapping.get(t, t) for t in tags) or "无明显风险"


def _has_spec_conflict(query: str, candidate: str) -> bool:
    query_norm = normalize_name(query).upper()
    cand_norm = normalize_name(candidate).upper()
    for pattern in [r"C\d{2}", r"\b\d{1,3}T\b", r"HRB\d{3}E?", r"DN\d+", r"DE\d+"]:
        q = set(re.findall(pattern, query_norm))
        c = set(re.findall(pattern, cand_norm))
        if q and c and q.isdisjoint(c):
            return True
    return False
