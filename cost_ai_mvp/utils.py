from __future__ import annotations

import hashlib
import re
import shutil
import uuid
from datetime import datetime
from pathlib import Path


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_compact() -> str:
    return datetime.now().strftime("%Y%m%d")


def new_id(prefix: str) -> str:
    return f"{prefix}{uuid.uuid4().hex[:12].upper()}"


def safe_filename(name: str, max_len: int = 120) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:max_len] or "unnamed"


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def copy_file(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def rel_to(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


def normalize_unit(unit: str | None) -> str:
    if unit is None:
        return ""
    text = str(unit).strip()
    mapping = {
        "立方米": "m³",
        "m3": "m³",
        "M3": "m³",
        "m^3": "m³",
        "平方米": "m²",
        "m2": "m²",
        "M2": "m²",
        "m^2": "m²",
        "吨": "t",
        "T": "t",
        "千克": "kg",
        "公斤": "kg",
        "米": "m",
        "延米": "m",
        "工日": "工日",
        "台班": "台班",
    }
    return mapping.get(text, text)


def normalize_name(name: str | None) -> str:
    if name is None:
        return ""
    text = str(name).strip().replace(" ", "")
    replacements = {
        "商品砼": "商品混凝土",
        "预拌砼": "预拌混凝土",
        "砼": "混凝土",
        "普 工": "普工",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def infer_item_type(name: str, unit: str = "") -> str:
    text = normalize_name(name)
    unit = normalize_unit(unit)
    if unit in {"工日", "工时"} or any(k in text for k in ["普工", "技工", "钢筋工", "木工", "电工", "焊工", "人工"]):
        return "labor"
    if unit in {"台班", "台时"} or any(k in text for k in ["汽车吊", "塔吊", "挖掘机", "泵车", "机械", "吊车"]):
        return "machinery"
    if unit in {"台", "套", "组"} and any(k in text for k in ["风机", "水泵", "配电柜", "配电箱", "机组", "设备"]):
        return "equipment"
    return "material"

