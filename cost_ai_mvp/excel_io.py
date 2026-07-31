from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from .utils import infer_item_type, normalize_name, normalize_unit


HEADER_ALIASES = {
    "name": ["名称", "材料名称", "项目名称", "人工材料机械名称", "人材机名称", "设备名称", "材料设备名称"],
    "spec": ["规格", "规格型号", "型号", "项目特征", "特征"],
    "unit": ["单位", "计量单位"],
    "unit_price": ["单价", "市场价", "报价", "除税价", "含税价", "综合单价", "人材机单价"],
    "quantity": ["数量", "工程量", "消耗量", "含量"],
    "total_price": ["合价", "金额", "合计"],
    "type": ["类别", "类型", "人材机类别"],
    "brand": ["品牌", "厂家", "供应商"],
}


@dataclass
class ParsedRow:
    item_type: str
    original_name: str
    original_spec: str
    original_unit: str
    standard_name: str
    standard_spec: str
    standard_unit: str
    brand: str | None
    quantity: float | None
    unit_price: float
    total_price: float | None
    source_sheet: str
    source_row: int
    notes: str = ""


@dataclass
class ParseResult:
    rows: list[ParsedRow]
    errors: list[str]
    sheets_seen: list[str]


def read_excel_preview(path: Path, sheet_name: str | None = None, nrows: int = 20) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    target = sheet_name or xl.sheet_names[0]
    return pd.read_excel(path, sheet_name=target, header=None, nrows=nrows)


def parse_price_excel(path: Path, file_type: str) -> ParseResult:
    rows: list[ParsedRow] = []
    errors: list[str] = []
    sheets_seen: list[str] = []
    try:
        xl = pd.ExcelFile(path)
    except Exception as exc:
        return ParseResult([], [f"无法打开 Excel：{exc}"], [])

    for sheet in xl.sheet_names:
        sheets_seen.append(sheet)
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=None)
        except Exception as exc:
            errors.append(f"{sheet}: 读取失败：{exc}")
            continue
        if df.empty or df.dropna(how="all").empty:
            continue
        parsed, sheet_errors = _parse_sheet(df, sheet, file_type)
        rows.extend(parsed)
        errors.extend(sheet_errors)
    if not rows and not errors:
        errors.append("未在任何 sheet 中识别到价格数据")
    return ParseResult(rows, errors, sheets_seen)


def parse_new_project_items(path: Path, sheet_name: str | None = None) -> tuple[list[dict[str, Any]], list[str]]:
    errors: list[str] = []
    try:
        xl = pd.ExcelFile(path)
    except Exception as exc:
        return [], [f"无法打开 Excel：{exc}"]
    target = sheet_name or xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=target, header=None)
    if df.empty:
        return [], ["新项目表为空"]
    header_idx, mapping = _find_header(df)
    if header_idx is None:
        return [], ["未识别到名称/单位表头"]
    items: list[dict[str, Any]] = []
    for i in range(header_idx + 1, len(df)):
        raw_name = _cell(df.iat[i, mapping["name"]])
        raw_unit = _cell(df.iat[i, mapping["unit"]])
        if not raw_name or not raw_unit:
            continue
        raw_spec = _cell(df.iat[i, mapping.get("spec")]) if "spec" in mapping else ""
        raw_type = _cell(df.iat[i, mapping.get("type")]) if "type" in mapping else ""
        unit_price = _to_float(df.iat[i, mapping.get("unit_price")]) if "unit_price" in mapping else None
        item_type = _map_type(raw_type) or infer_item_type(raw_name, raw_unit)
        items.append(
            {
                "source_row": i + 1,
                "item_type": item_type,
                "original_name": raw_name,
                "original_spec": raw_spec,
                "original_unit": raw_unit,
                "standard_name": normalize_name(raw_name),
                "standard_spec": raw_spec,
                "standard_unit": normalize_unit(raw_unit),
                "original_unit_price": unit_price,
            }
        )
    if not items:
        errors.append("识别到表头，但未读取到有效人材机行")
    return items, errors


def export_pricing_result(input_path: Path, output_path: Path, items: list[dict[str, Any]], candidates: dict[str, list[dict[str, Any]]]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        try:
            xl = pd.ExcelFile(input_path)
            for sheet in xl.sheet_names:
                df = pd.read_excel(input_path, sheet_name=sheet)
                df.to_excel(writer, sheet_name=sheet[:31], index=False)
        except Exception:
            pass
        result_df = pd.DataFrame(
            [
                {
                    "原始行号": r.get("source_row"),
                    "系统类别": _type_cn(r.get("item_type")),
                    "原始名称": r.get("original_name"),
                    "原始规格": r.get("original_spec"),
                    "原始单位": r.get("original_unit"),
                    "标准化名称": r.get("standard_name"),
                    "标准化规格": r.get("standard_spec"),
                    "标准化单位": r.get("standard_unit"),
                    "建议价": r.get("suggested_price"),
                    "参考低值": r.get("reference_low"),
                    "参考高值": r.get("reference_high"),
                    "主要来源": r.get("main_source_text"),
                    "来源类型": r.get("source_type"),
                    "可信度": r.get("confidence"),
                    "风险提示": r.get("risk_text"),
                    "AI说明": r.get("ai_explanation"),
                    "是否采用": r.get("is_accepted"),
                    "人工确认价": r.get("confirmed_price"),
                    "补价状态": r.get("status"),
                    "备注": r.get("notes"),
                }
                for r in items
            ]
        )
        result_df.to_excel(writer, sheet_name="补价说明汇总", index=False)
        cand_rows = []
        for item_id, cand_list in candidates.items():
            for c in cand_list:
                cand_rows.append(c)
        pd.DataFrame(cand_rows).to_excel(writer, sheet_name="匹配候选记录", index=False)


def _parse_sheet(df: pd.DataFrame, sheet: str, file_type: str) -> tuple[list[ParsedRow], list[str]]:
    header_idx, mapping = _find_header(df)
    if header_idx is None:
        return [], []
    rows: list[ParsedRow] = []
    errors: list[str] = []
    for i in range(header_idx + 1, len(df)):
        try:
            raw_name = _cell(df.iat[i, mapping["name"]])
            raw_unit = _cell(df.iat[i, mapping["unit"]])
            price = _to_float(df.iat[i, mapping["unit_price"]]) if "unit_price" in mapping else None
            if not raw_name and not raw_unit and price is None:
                continue
            if not raw_name or not raw_unit or price is None:
                errors.append(f"{sheet} 第{i+1}行缺少名称/单位/单价")
                continue
            raw_spec = _cell(df.iat[i, mapping.get("spec")]) if "spec" in mapping else ""
            raw_type = _cell(df.iat[i, mapping.get("type")]) if "type" in mapping else ""
            brand = _cell(df.iat[i, mapping.get("brand")]) if "brand" in mapping else ""
            quantity = _to_float(df.iat[i, mapping.get("quantity")]) if "quantity" in mapping else None
            total = _to_float(df.iat[i, mapping.get("total_price")]) if "total_price" in mapping else None
            item_type = _map_type(raw_type) or infer_item_type(raw_name, raw_unit)
            rows.append(
                ParsedRow(
                    item_type=item_type,
                    original_name=raw_name,
                    original_spec=raw_spec,
                    original_unit=raw_unit,
                    standard_name=normalize_name(raw_name),
                    standard_spec=raw_spec,
                    standard_unit=normalize_unit(raw_unit),
                    brand=brand or None,
                    quantity=quantity,
                    unit_price=price,
                    total_price=total,
                    source_sheet=sheet,
                    source_row=i + 1,
                    notes=file_type,
                )
            )
        except Exception as exc:
            errors.append(f"{sheet} 第{i+1}行解析失败：{exc}")
    return rows, errors


def _find_header(df: pd.DataFrame) -> tuple[int | None, dict[str, int]]:
    max_scan = min(30, len(df))
    best: tuple[int | None, dict[str, int], int] = (None, {}, 0)
    for i in range(max_scan):
        row = [_cell(v) for v in df.iloc[i].tolist()]
        mapping: dict[str, int] = {}
        for field, aliases in HEADER_ALIASES.items():
            for idx, value in enumerate(row):
                if value and any(alias in value for alias in aliases):
                    mapping[field] = idx
                    break
        score = sum(1 for k in ["name", "unit", "unit_price"] if k in mapping) + len(mapping) * 0.1
        if score > best[2]:
            best = (i, mapping, int(score * 10))
    if best[0] is not None and "name" in best[1] and "unit" in best[1]:
        return best[0], best[1]
    return None, {}


def _cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _map_type(raw: str) -> str | None:
    if not raw:
        return None
    if "人工" in raw or raw in {"人"}:
        return "labor"
    if "机械" in raw or raw in {"机"}:
        return "machinery"
    if "设备" in raw:
        return "equipment"
    if "材料" in raw or raw in {"材"}:
        return "material"
    return None


def _type_cn(item_type: str | None) -> str:
    return {"labor": "人工", "material": "材料", "machinery": "机械", "equipment": "设备"}.get(item_type or "", item_type or "")

