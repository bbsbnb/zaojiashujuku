from __future__ import annotations

import html
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import fitz
import requests

from .db import connect, insert_many
from .utils import copy_file, file_hash, new_id, now, normalize_name, normalize_unit, safe_filename, today_compact


DEFAULT_SOURCE_URL = "https://js.shaanxi.gov.cn/sy/yw/zjglfw/zjxx/"
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept": "text/html,application/pdf,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.7",
}

TABLE_HEADER_MARKERS = ("材料编码", "材料名称", "规格型号", "单位", "除税价格", "含税价格")
UNIT_HINTS = {
    "t", "kg", "g", "m", "m2", "m3", "㎡", "m²", "个", "套", "台", "只", "根", "张",
    "块", "袋", "箱", "卷", "樘", "辆", "组", "米", "瓶", "罐", "盘", "件", "桶",
    "副", "片", "盒", "门", "工日", "工时", "小时", "班", "台班", "kW", "kva", "kWh", "L", "l",
}


@dataclass
class ShaanxiImportResult:
    summary: str
    log: list[str]
    price_count: int
    file_id: str
    source_title: str
    source_url: str
    pdf_url: str
    stored_pdf: str


def import_shaanxi_price(library: Path, source_url: str = "", pdf_url: str = "", pdf_path: Path | None = None) -> ShaanxiImportResult:
    page_url = source_url.strip() or DEFAULT_SOURCE_URL
    pdf_link = pdf_url.strip()
    article_title = "陕西信息价"
    price_date = _date_from_url(page_url) or datetime.now()

    if pdf_path is not None:
        pdf_file = Path(pdf_path)
        if not pdf_file.exists():
            raise FileNotFoundError(f"PDF 不存在: {pdf_file}")
        stored_pdf = pdf_file
        if not source_url:
            article_title = pdf_file.stem
    else:
        html_text = ""
        if page_url.lower().endswith(".pdf"):
            pdf_link = page_url
        else:
            html_text = _fetch_text(page_url)
            if not pdf_link:
                pdf_link = _extract_pdf_url(html_text, page_url)
            article_title = _extract_title(html_text) or article_title
            parsed_date = _extract_date(html_text) or _date_from_url(page_url)
            if parsed_date:
                price_date = parsed_date
        if not pdf_link:
            raise ValueError("未能从来源页面找到 PDF 链接")
        referer = page_url if not page_url.lower().endswith(".pdf") else ""
        stored_pdf = _download_pdf(library, pdf_link, referer=referer, title=article_title)

    items, parse_log = _parse_pdf(stored_pdf, price_date, article_title)
    if not items:
        raise ValueError("未解析到陕西信息价明细")

    file_id = new_id("F")
    t = now()
    archive_dir = library / "temp" / "imports" / "shaanxi"
    archive_dir.mkdir(parents=True, exist_ok=True)
    archive_path = archive_dir / f"{file_id}_{safe_filename(stored_pdf.name)}"
    if stored_pdf != archive_path:
        copy_file(stored_pdf, archive_path)

    with connect(library) as conn:
        conn.execute(
            """
            INSERT INTO files(id,project_id,original_name,original_path,archive_path,archive_relative_path,file_type,extension,file_size,file_hash,import_status,parsed_rows,error_rows,parse_message,is_archived_only,imported_at,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                file_id, None, stored_pdf.name, str(stored_pdf.resolve()), str(archive_path.resolve()),
                f"temp/imports/shaanxi/{archive_path.name}", "material_price", stored_pdf.suffix.lower(),
                archive_path.stat().st_size, file_hash(archive_path), "parsed", len(items), 0, "", 1, t, t, t,
            ),
        )
        price_rows = []
        for index, item in enumerate(items, 1):
            price_rows.append(
                {
                    "id": new_id("PR"),
                    "project_id": None,
                    "file_id": file_id,
                    "quote_item_id": None,
                    "standard_item_id": None,
                    "item_type": "material",
                    "original_name": item["original_name"],
                    "original_spec": item["original_spec"],
                    "original_unit": item["original_unit"],
                    "standard_name": item["standard_name"],
                    "standard_spec": item["standard_spec"],
                    "standard_unit": item["standard_unit"],
                    "brand": None,
                    "quantity": None,
                    "unit_price": item["unit_price"],
                    "total_price": None,
                    "source_type": "shaanxi_material_price",
                    "cost_stage": "材料信息价",
                    "region": "陕西",
                    "project_type": None,
                    "price_year": price_date.year,
                    "price_date": price_date.strftime("%Y-%m-%d"),
                    "price_scope": "陕西省住房和城乡建设厅材料信息价（除税）",
                    "is_tax_included": "no",
                    "tax_rate": item["tax_rate"],
                    "is_freight_included": "unknown",
                    "is_installation_included": "unknown",
                    "payment_terms": "",
                    "duration": "",
                    "source_sheet": item["source_sheet"],
                    "source_row": index,
                    "is_outlier": 0,
                    "outlier_reason": None,
                    "confidence_note": f"含税价:{item['tax_included_price']:.2f}",
                    "notes": item["notes"],
                    "created_at": t,
                    "updated_at": t,
                }
            )
        insert_many(conn, "prices", price_rows)
        conn.commit()

    summary = f"完成：导入 {len(price_rows)} 条陕西材料信息价。"
    log = [
        f"来源页面：{page_url}",
        f"PDF 链接：{pdf_link}",
        f"保存文件：{archive_path.resolve()}",
        *parse_log[:20],
    ]
    return ShaanxiImportResult(summary, log, len(price_rows), file_id, article_title, page_url, pdf_link, str(archive_path.resolve()))


def _parse_pdf(pdf_path: Path, price_date: datetime, title: str) -> tuple[list[dict], list[str]]:
    doc = fitz.open(pdf_path)
    start_page = _find_table_start_page(doc)
    items: list[dict] = []
    current_section = "材料信息价"
    for page_no in range(start_page, len(doc)):
        lines = [_clean_line(line) for line in doc[page_no].get_text("text").splitlines()]
        lines = [line for line in lines if line]
        idx = 0
        while idx < len(lines):
            line = lines[idx]
            if line.isdigit() or _is_header_line(line):
                idx += 1
                continue
            if _looks_like_category(line):
                current_section = line
                idx += 1
                continue
            if _is_code_line(line):
                item, next_idx = _parse_record(lines, idx, current_section, page_no + 1, price_date, title)
                if item:
                    items.append(item)
                idx = next_idx
                continue
            idx += 1
    return items, [f"识别起始页：第 {start_page + 1} 页", f"解析记录数：{len(items)}"]


def _parse_record(lines: list[str], start: int, section: str, page_no: int, price_date: datetime, title: str) -> tuple[dict | None, int]:
    match = re.match(r"^(\d{9})\s+(.+)$", lines[start])
    if not match:
        return None, start + 1
    code = match.group(1)
    name = match.group(2).strip()
    extras: list[str] = []
    prices: list[float] = []
    idx = start + 1
    while idx < len(lines):
        candidate = lines[idx]
        if _is_code_line(candidate) or _is_header_line(candidate):
            break
        if _is_price(candidate):
            prices.append(float(candidate.replace(",", "")))
            idx += 1
            if len(prices) >= 2:
                break
            continue
        extras.append(candidate)
        idx += 1
    if len(prices) < 2:
        return None, idx

    unit = ""
    spec_parts = extras[:]
    for pos in range(len(extras) - 1, -1, -1):
        if _looks_like_unit(extras[pos]):
            unit = extras[pos]
            spec_parts = extras[:pos]
            break
    if not unit and extras:
        unit = extras[-1]
        spec_parts = extras[:-1]
    if not unit:
        unit = "个"

    spec = " ".join(spec_parts).strip()
    tax_rate = round((prices[1] / prices[0] - 1) * 100, 2) if prices[0] else None
    return (
        {
            "code": code,
            "original_name": name,
            "original_spec": spec,
            "original_unit": unit,
            "standard_name": normalize_name(name),
            "standard_spec": spec,
            "standard_unit": normalize_unit(unit),
            "unit_price": prices[0],
            "tax_included_price": prices[1],
            "tax_rate": tax_rate,
            "source_sheet": f"{section} / 第{page_no}页",
            "notes": f"材料编码={code};标题={title};日期={price_date.strftime('%Y-%m-%d')}",
        },
        idx,
    )


def _find_table_start_page(doc: fitz.Document) -> int:
    for page_no in range(len(doc)):
        text = doc[page_no].get_text("text")
        if all(marker in text for marker in TABLE_HEADER_MARKERS):
            return page_no
    return 0


def _download_pdf(library: Path, pdf_url: str, referer: str, title: str) -> Path:
    temp_dir = library / "temp" / "imports" / "shaanxi"
    temp_dir.mkdir(parents=True, exist_ok=True)
    target = temp_dir / f"{today_compact()}_{safe_filename(title or 'shaanxi_material_price')}.pdf"
    headers = dict(REQUEST_HEADERS)
    if referer:
        headers["Referer"] = referer
    resp = requests.get(pdf_url, headers=headers, timeout=60)
    resp.raise_for_status()
    target.write_bytes(resp.content)
    return target


def _fetch_text(url: str) -> str:
    resp = requests.get(url, headers=REQUEST_HEADERS, timeout=30)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or resp.encoding or "utf-8"
    return resp.text


def _extract_pdf_url(html_text: str, base_url: str) -> str:
    matches = re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', html_text, flags=re.I)
    if not matches:
        return ""
    return urljoin(base_url, html.unescape(matches[0]))


def _extract_title(html_text: str) -> str:
    match = re.search(r"<title>(.*?)</title>", html_text, flags=re.I | re.S)
    return re.sub(r"\s+", " ", html.unescape(match.group(1))).strip() if match else ""


def _extract_date(html_text: str) -> datetime | None:
    for pattern in (r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", r"(20\d{2})(\d{2})(\d{2})"):
        match = re.search(pattern, html_text)
        if match:
            try:
                return datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            except ValueError:
                pass
    return None


def _date_from_url(url: str) -> datetime | None:
    match = re.search(r"(20\d{2})(\d{2})(\d{2})", url)
    if not match:
        return None
    try:
        return datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


def _clean_line(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def _is_header_line(line: str) -> bool:
    return any(marker in line for marker in TABLE_HEADER_MARKERS)


def _is_code_line(line: str) -> bool:
    return bool(re.match(r"^\d{9}\s+", line))


def _is_price(line: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}(?:,?\d{3})*(?:\.\d+)?", line))


def _looks_like_unit(line: str) -> bool:
    return line in UNIT_HINTS or (bool(re.fullmatch(r"[A-Za-z0-9㎡²\-/]+", line)) and len(line) <= 8)


def _looks_like_category(line: str) -> bool:
    return bool(line) and not _is_code_line(line) and not _is_price(line) and not _is_header_line(line) and not any(ch.isdigit() for ch in line) and 2 <= len(line) <= 20

