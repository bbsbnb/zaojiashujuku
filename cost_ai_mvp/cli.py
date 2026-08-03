from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

from .db import connect, init_library, insert_many
from .excel_io import export_pricing_result, parse_new_project_items, parse_price_excel
from .matching import build_price_suggestion, candidates_for_item, explanation, search_prices
from .shaanxi_import import import_shaanxi_price
from .utils import copy_file, file_hash, new_id, now, rel_to, safe_filename, today_compact


FILE_TYPE_BY_NAME = {
    "工程量清单": "boq",
    "综合单价分析表": "unit_price_analysis",
    "人材机表": "resource_price",
    "材料设备价格表": "material_price",
    "供应商报价单": "supplier_quote",
    "询价": "supplier_quote",
    "合同": "contract_price",
    "结算": "settlement",
    "变更": "variation_claim",
    "分包": "subcontract_quote",
}

SOURCE_TYPE_BY_FILE = {
    "supplier_quote": "inquiry",
    "material_price": "material_price",
    "resource_price": "resource_price",
    "unit_price_analysis": "unit_price_analysis",
    "contract_price": "contract",
    "settlement": "settlement",
    "tender_price": "tender",
    "subcontract_quote": "subcontract_quote",
}


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="AI 人材机价格助手 MVP")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("init", help="初始化本地资料库")
    p.add_argument("--library", required=True)

    p = sub.add_parser("import-samples", help="导入第一轮测试样本")
    p.add_argument("--library", required=True)
    p.add_argument("--samples", required=True)

    p = sub.add_parser("query", help="查询价格")
    p.add_argument("--library", required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--region", default="深圳")
    p.add_argument("--type", dest="item_type")

    p = sub.add_parser("price-file", help="上传新项目表并补价导出")
    p.add_argument("--library", required=True)
    p.add_argument("--input", required=True)
    p.add_argument("--project-name", required=True)
    p.add_argument("--region", default="深圳")
    p.add_argument("--project-type", default="住宅")
    p.add_argument("--price-scope", default="含税")
    p.add_argument("--payment-terms", default="")
    p.add_argument("--duration", default="")

    p = sub.add_parser("import-shaanxi", help="导入陕西造价信息网材料信息价")
    p.add_argument("--library", required=True)
    p.add_argument("--source-url", default="")
    p.add_argument("--pdf-url", default="")
    p.add_argument("--pdf-path", default="")

    args = parser.parse_args()
    library = Path(args.library)
    if args.cmd == "init":
        init_library(library)
        print(f"资料库已初始化：{library.resolve()}")
    elif args.cmd == "import-samples":
        init_library(library)
        import_samples(library, Path(args.samples))
    elif args.cmd == "query":
        query(library, args.text, args.region, args.item_type)
    elif args.cmd == "price-file":
        init_library(library)
        price_file(library, Path(args.input), args.project_name, args.region, args.project_type, args.price_scope, args.payment_terms, args.duration)
    elif args.cmd == "import-shaanxi":
        init_library(library)
        result = import_shaanxi_price(
            library=library,
            source_url=args.source_url,
            pdf_url=args.pdf_url,
            pdf_path=Path(args.pdf_path) if args.pdf_path else None,
        )
        print(result.summary)
        for line in result.log:
            print(line)


def import_samples(library: Path, samples: Path) -> dict:
    info_path = samples / "项目基础信息汇总.xlsx"
    project_infos = _read_project_infos(info_path)
    imported_projects = 0
    imported_prices = 0
    log: list[str] = []
    with connect(library) as conn:
        for folder in sorted([p for p in samples.iterdir() if p.is_dir() and not p.name.startswith("06_")]):
            info = _project_info_for(folder, project_infos)
            project_id = new_id("P")
            archive_folder = library / "projects" / f"{project_id}_{safe_filename(info['name'])}"
            archive_folder.mkdir(parents=True, exist_ok=True)
            t = now()
            conn.execute(
                """
                INSERT INTO projects(id,name,short_name,region,year,project_type,cost_stage,price_scope,payment_terms,duration,original_folder_path,archive_folder_path,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    project_id,
                    info["name"],
                    info["name"],
                    info["region"],
                    info["year"],
                    info["project_type"],
                    info["cost_stage"],
                    info["price_scope"],
                    info.get("payment_terms", ""),
                    info.get("duration", ""),
                    str(folder.resolve()),
                    str(archive_folder.resolve()),
                    t,
                    t,
                ),
            )
            imported_projects += 1
            for file in sorted(folder.glob("*.xls*")):
                file_type = _guess_file_type(file.name)
                file_id, count = _archive_and_parse(conn, library, project_id, archive_folder, file, file_type, info)
                imported_prices += count
                msg = f"导入 {folder.name}/{file.name}: {file_type}, 价格 {count} 条"
                print(msg)
                log.append(msg)
        conn.commit()
    summary = f"完成：项目 {imported_projects} 个，价格记录 {imported_prices} 条。"
    print(summary)
    return {"projects": imported_projects, "prices": imported_prices, "log": log, "summary": summary}


def query(library: Path, text: str, region: str, item_type: str | None) -> dict:
    with connect(library) as conn:
        rows = search_prices(conn, text, region=region, item_type=item_type, limit=20)
        suggestion = build_price_suggestion(rows, region=region)
        qid = new_id("QRY")
        conn.execute(
            "INSERT INTO query_logs(id,query_text,filter_region,suggested_price,reference_low,reference_high,confidence,risk_tags,ai_explanation,result_summary,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (
                qid,
                text,
                region,
                suggestion.get("suggested_price"),
                suggestion.get("reference_low"),
                suggestion.get("reference_high"),
                suggestion.get("confidence"),
                ";".join(suggestion.get("risk_tags", [])),
                explanation(text, suggestion),
                suggestion.get("main_source_text"),
                now(),
            ),
        )
        conn.commit()
    result = {"query": text, "suggestion": suggestion, "explanation": explanation(text, suggestion), "top_rows": _public_rows(rows[:20])}
    _safe_print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def price_file(library: Path, input_path: Path, project_name: str, region: str, project_type: str, price_scope: str, payment_terms: str, duration: str) -> dict:
    items, errors = parse_new_project_items(input_path)
    if errors:
        print("解析提示：")
        for e in errors:
            print("-", e)
    task_id = new_id("N")
    task_dir = library / "new_projects" / f"{task_id}_{safe_filename(project_name)}"
    archive_input = task_dir / "input" / f"{task_id}_input_{safe_filename(input_path.name)}"
    copy_file(input_path, archive_input)
    output = task_dir / "output" / f"{task_id}_{safe_filename(project_name)}_人材机补价结果_{today_compact()}.xlsx"
    priced_items: list[dict] = []
    candidate_map: dict[str, list[dict]] = {}
    with connect(library) as conn:
        t = now()
        conn.execute(
            """
            INSERT INTO pricing_tasks(id,project_name,region,project_type,quote_date,payment_terms,duration,price_scope,input_file_path,archive_input_path,archive_input_relative_path,status,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (task_id, project_name, region, project_type, t[:10], payment_terms, duration, price_scope, str(input_path.resolve()), str(archive_input.resolve()), rel_to(archive_input, library), "matching", t, t),
        )
        for item in items:
            item_id = new_id("TI")
            candidates, suggestion = candidates_for_item(conn, item, region=region)
            status = "matched" if candidates else "need_inquiry"
            item_row = {
                **item,
                "id": item_id,
                "task_id": task_id,
                "suggested_price": suggestion.get("suggested_price"),
                "reference_low": suggestion.get("reference_low"),
                "reference_high": suggestion.get("reference_high"),
                "main_source_text": suggestion.get("main_source_text"),
                "main_source_price_id": suggestion.get("main_source_price_id"),
                "source_type": suggestion.get("source_type"),
                "confidence": suggestion.get("confidence"),
                "risk_tags": ";".join(suggestion.get("risk_tags", [])),
                "risk_text": suggestion.get("risk_text"),
                "ai_explanation": explanation(item.get("original_name", ""), suggestion),
                "status": status,
                "is_accepted": 1 if candidates and (suggestion.get("confidence") == "high") else 0,
                "confirmed_price": None,
                "notes": "",
            }
            conn.execute(
                """
                INSERT INTO pricing_task_items(id,task_id,source_row,item_type,original_name,original_spec,original_unit,standard_name,standard_spec,standard_unit,original_unit_price,suggested_price,reference_low,reference_high,main_source_text,main_source_price_id,source_type,confidence,risk_tags,risk_text,ai_explanation,is_accepted,confirmed_price,status,notes,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    item_id,
                    task_id,
                    item_row["source_row"],
                    item_row["item_type"],
                    item_row["original_name"],
                    item_row["original_spec"],
                    item_row["original_unit"],
                    item_row["standard_name"],
                    item_row["standard_spec"],
                    item_row["standard_unit"],
                    item_row.get("original_unit_price"),
                    item_row.get("suggested_price"),
                    item_row.get("reference_low"),
                    item_row.get("reference_high"),
                    item_row.get("main_source_text"),
                    item_row.get("main_source_price_id"),
                    item_row.get("source_type"),
                    item_row.get("confidence"),
                    item_row.get("risk_tags"),
                    item_row.get("risk_text"),
                    item_row.get("ai_explanation"),
                    item_row.get("is_accepted"),
                    item_row.get("confirmed_price"),
                    item_row.get("status"),
                    item_row.get("notes"),
                    t,
                    t,
                ),
            )
            out_cands = []
            for c in candidates:
                cid = new_id("MC")
                conn.execute(
                    """
                    INSERT INTO match_candidates(id,task_item_id,rank_no,source_price_id,source_quote_item_id,standard_item_id,candidate_name,candidate_spec,candidate_unit,candidate_price,reference_low,reference_high,source_type,source_text,source_region,source_date_or_year,match_score,confidence,match_reason,risk_tags,risk_text,is_selected,created_at)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        cid,
                        item_id,
                        c["rank_no"],
                        c.get("source_price_id"),
                        c.get("source_quote_item_id"),
                        c.get("standard_item_id"),
                        c["candidate_name"],
                        c.get("candidate_spec"),
                        c["candidate_unit"],
                        c["candidate_price"],
                        c.get("reference_low"),
                        c.get("reference_high"),
                        c["source_type"],
                        c["source_text"],
                        c.get("source_region"),
                        c.get("source_date_or_year"),
                        c.get("match_score"),
                        c["confidence"],
                        c.get("match_reason"),
                        c.get("risk_tags"),
                        c.get("risk_text"),
                        1 if c["rank_no"] == 1 else 0,
                        t,
                    ),
                )
                out_cands.append({"补价明细ID": item_id, **c})
            candidate_map[item_id] = out_cands
            priced_items.append(item_row)
        conn.execute("UPDATE pricing_tasks SET status=?, output_file_path=?, output_relative_path=?, updated_at=? WHERE id=?", ("exported", str(output.resolve()), rel_to(output, library), now(), task_id))
        conn.commit()
    export_pricing_result(input_path, output, priced_items, candidate_map)
    summary = f"补价完成：{len(priced_items)} 项，导出：{output.resolve()}"
    print(summary)
    return {"task_id": task_id, "items": len(priced_items), "output": str(output.resolve()), "summary": summary, "errors": errors}


def _archive_and_parse(conn, library: Path, project_id: str, archive_folder: Path, file: Path, file_type: str, info: dict) -> tuple[str, int]:
    file_id = new_id("F")
    sub = _folder_for_file_type(file_type)
    archive_path = archive_folder / "original_files" / sub / f"{file_id}_{file_type}_{safe_filename(file.name)}"
    copy_file(file, archive_path)
    t = now()
    h = file_hash(archive_path)
    result = parse_price_excel(archive_path, file_type) if file_type != "archive_only" else None
    rows = result.rows if result else []
    message = "; ".join(result.errors[:10]) if result and result.errors else ""
    conn.execute(
        """
        INSERT INTO files(id,project_id,original_name,original_path,archive_path,archive_relative_path,file_type,extension,file_size,file_hash,import_status,parsed_rows,error_rows,parse_message,is_archived_only,imported_at,created_at,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        (file_id, project_id, file.name, str(file.resolve()), str(archive_path.resolve()), rel_to(archive_path, library), file_type, file.suffix.lower(), archive_path.stat().st_size, h, "已入库" if rows else "已解析", len(rows), len(result.errors) if result else 0, message, 1 if file_type == "archive_only" else 0, t, t, t),
    )
    price_rows = []
    for r in rows:
        price_rows.append(
            {
                "id": new_id("PR"),
                "project_id": project_id,
                "file_id": file_id,
                "quote_item_id": None,
                "standard_item_id": None,
                "item_type": r.item_type,
                "original_name": r.original_name,
                "original_spec": r.original_spec,
                "original_unit": r.original_unit,
                "standard_name": r.standard_name,
                "standard_spec": r.standard_spec,
                "standard_unit": r.standard_unit,
                "brand": r.brand,
                "quantity": r.quantity,
                "unit_price": r.unit_price,
                "total_price": r.total_price,
                "source_type": SOURCE_TYPE_BY_FILE.get(file_type, file_type),
                "cost_stage": info["cost_stage"],
                "region": info["region"],
                "project_type": info["project_type"],
                "price_year": info["year"],
                "price_date": None,
                "price_scope": info["price_scope"],
                "is_tax_included": "yes" if "含税" in info["price_scope"] else "unknown",
                "tax_rate": None,
                "is_freight_included": "unknown",
                "is_installation_included": "unknown",
                "payment_terms": info.get("payment_terms", ""),
                "duration": info.get("duration", ""),
                "source_sheet": r.source_sheet,
                "source_row": r.source_row,
                "is_outlier": 0,
                "outlier_reason": None,
                "confidence_note": None,
                "notes": r.notes,
                "created_at": t,
                "updated_at": t,
            }
        )
    insert_many(conn, "prices", price_rows)
    return file_id, len(price_rows)


def _read_project_infos(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        import pandas as pd

        df = pd.read_excel(path)
    except Exception:
        return []
    return [dict(row.dropna()) for _, row in df.iterrows()]


def _project_info_for(folder: Path, infos: list[dict]) -> dict:
    text = folder.name
    region = "深圳" if "深圳" in text else "广州" if "广州" in text else "东莞" if "东莞" in text else "佛山" if "佛山" in text else "深圳"
    ptype = "住宅" if "住宅" in text else "商业办公" if "商业" in text or "办公" in text else "综合"
    for info in infos:
        joined = " ".join(str(v) for v in info.values())
        if folder.name[:3] in joined or region in joined:
            pass
    return {
        "name": folder.name.split("_", 1)[-1],
        "region": region,
        "year": 2026,
        "project_type": ptype,
        "cost_stage": "合同",
        "price_scope": "含税",
        "payment_terms": "",
        "duration": "",
    }


def _guess_file_type(filename: str) -> str:
    for key, value in FILE_TYPE_BY_NAME.items():
        if key in filename:
            return value
    return "archive_only"


def _folder_for_file_type(file_type: str) -> str:
    return {
        "boq": "boq",
        "unit_price_analysis": "unit_price_analysis",
        "resource_price": "resource_price",
        "material_price": "material_price",
        "supplier_quote": "supplier_quote",
        "contract_price": "contract_settlement",
        "settlement": "contract_settlement",
        "variation_claim": "variation_claim",
        "subcontract_quote": "subcontract_quote",
    }.get(file_type, "archive_only")


def _public_rows(rows: list[dict]) -> list[dict]:
    return [
        {
            "原始类型": r.get("item_type"),
            "名称": r.get("standard_name") or r.get("original_name"),
            "规格": r.get("standard_spec") or r.get("original_spec"),
            "单位": r.get("standard_unit") or r.get("original_unit"),
            "单价": r.get("unit_price"),
            "地区": r.get("region"),
            "年份": r.get("price_year") or r.get("price_date"),
            "来源": r.get("source_type"),
            "匹配分": r.get("match_score"),
        }
        for r in rows
    ]


def _safe_print(text: str) -> None:
    try:
        print(text)
    except UnicodeEncodeError:
        sys.stdout.buffer.write((text + "\n").encode("utf-8", errors="replace"))


if __name__ == "__main__":
    main()
