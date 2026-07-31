from __future__ import annotations

from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "第一轮测试样本"


PROJECTS = [
    ("01_深圳住宅项目", "深圳", "住宅", 1.00),
    ("02_深圳商业办公项目", "深圳", "商业办公", 1.04),
    ("03_广州项目", "广州", "商业", 0.97),
    ("04_东莞或佛山项目", "东莞", "住宅", 0.94),
    ("05_询价记录完整项目", "深圳", "住宅", 1.02),
]


BASE_RESOURCE_ROWS = [
    ["材料", "C30商品混凝土", "泵送", "m³", 438],
    ["材料", "C35商品混凝土", "泵送", "m³", 455],
    ["材料", "HRB400E钢筋", "", "t", 4150],
    ["材料", "水泥", "42.5", "t", 520],
    ["材料", "中砂", "", "m³", 165],
    ["材料", "防水卷材", "3mm", "m²", 42],
    ["材料", "电缆", "WDZ-YJY-4x95+1x50", "m", 285],
    ["材料", "PVC管", "DN100", "m", 28],
    ["人工", "普工", "", "工日", 320],
    ["人工", "技工", "", "工日", 380],
    ["人工", "钢筋工", "", "工日", 410],
    ["人工", "电工", "", "工日", 430],
    ["机械", "25t汽车吊", "", "台班", 1500],
    ["机械", "50t汽车吊", "", "台班", 2300],
    ["机械", "挖掘机", "1m³", "台班", 1350],
    ["设备", "风机", "风量10000m³/h", "台", 5200],
    ["设备", "水泵", "流量50m³/h 扬程32m", "台", 4300],
]


QUOTE_ROWS = [
    ["C30商品混凝土", "泵送", "", "m³", 445, "含税到场"],
    ["C35商品混凝土", "泵送", "", "m³", 462, "含税到场"],
    ["HRB400E钢筋", "", "", "t", 4200, "含税到场"],
    ["电缆", "WDZ-YJY-4x95+1x50", "珠江", "m", 292, "含税"],
    ["25t汽车吊", "", "", "台班", 1550, "含司机燃油"],
    ["风机", "风量10000m³/h", "国产一线", "台", 5480, "含税不含安装"],
]


def scaled(value: float, factor: float) -> float:
    return round(value * factor, 2)


def make_resource_table(factor: float) -> pd.DataFrame:
    rows = []
    for item_type, name, spec, unit, price in BASE_RESOURCE_ROWS:
        rows.append([item_type, name, spec, unit, scaled(price, factor)])
    return pd.DataFrame(rows, columns=["类别", "名称", "规格型号", "单位", "单价"])


def make_quote_table(factor: float) -> pd.DataFrame:
    rows = []
    for name, spec, brand, unit, price, notes in QUOTE_ROWS:
        rows.append([name, spec, brand, unit, scaled(price, factor), notes])
    return pd.DataFrame(rows, columns=["名称", "规格型号", "品牌", "单位", "单价", "备注"])


def make_analysis_table(factor: float) -> pd.DataFrame:
    rows = []
    for idx, (item_type, name, spec, unit, price) in enumerate(BASE_RESOURCE_ROWS[:12], start=1):
        rows.append([f"010{idx:03d}", f"清单项{idx}", item_type, name, spec, unit, 1, scaled(price, factor), scaled(price, factor)])
    return pd.DataFrame(rows, columns=["清单编码", "清单名称", "类别", "名称", "规格型号", "单位", "消耗量", "单价", "合价"])


def make_material_table(factor: float) -> pd.DataFrame:
    rows = []
    for item_type, name, spec, unit, price in BASE_RESOURCE_ROWS:
        if item_type in {"材料", "设备"}:
            rows.append([name, spec, "", unit, scaled(price, factor), "含税"])
    return pd.DataFrame(rows, columns=["名称", "规格型号", "品牌", "单位", "单价", "备注"])


def make_boq_table(factor: float) -> pd.DataFrame:
    rows = []
    for idx, (_, name, spec, unit, price) in enumerate(BASE_RESOURCE_ROWS[:8], start=1):
        qty = 100 + idx * 10
        rows.append([f"010{idx:03d}", name, spec, unit, qty, scaled(price, factor), scaled(price * qty, factor)])
    return pd.DataFrame(rows, columns=["清单编码", "名称", "规格型号", "单位", "工程量", "综合单价", "合价"])


def main() -> None:
    if not SAMPLES.exists():
        raise SystemExit(f"样本目录不存在：{SAMPLES}")

    created = []
    for folder_name, region, project_type, factor in PROJECTS:
        folder = SAMPLES / folder_name
        folder.mkdir(parents=True, exist_ok=True)
        files = {
            "人材机表_模拟补全.xlsx": make_resource_table(factor),
            "供应商报价单_模拟补全.xlsx": make_quote_table(factor),
            "综合单价分析表_模拟补全.xlsx": make_analysis_table(factor),
            "材料设备价格表_模拟补全.xlsx": make_material_table(factor),
            "工程量清单_模拟补全.xlsx": make_boq_table(factor),
        }
        for filename, df in files.items():
            path = folder / filename
            df.to_excel(path, index=False)
            created.append(path)

    new_folder = SAMPLES / "06_新项目补价测试表"
    new_folder.mkdir(parents=True, exist_ok=True)
    new_items = pd.DataFrame(
        [
            ["材料", "商品砼", "C30泵送", "m³", ""],
            ["材料", "商品混凝土", "C35泵送", "m³", ""],
            ["材料", "钢筋", "HRB400E", "t", ""],
            ["材料", "电缆", "WDZ-YJY-4x95+1x50", "m", ""],
            ["人工", "普工", "", "工日", ""],
            ["人工", "技工", "", "工日", ""],
            ["机械", "汽车吊", "25t", "台班", ""],
            ["机械", "挖掘机", "1m³", "台班", ""],
            ["设备", "风机", "风量10000m³/h", "台", ""],
            ["材料", "未知材料X", "", "kg", ""],
        ],
        columns=["类别", "名称", "规格型号", "单位", "原单价"],
    )
    new_path = new_folder / "新项目补价测试表_模拟补全.xlsx"
    new_items.to_excel(new_path, index=False)
    created.append(new_path)

    print(f"已生成模拟补全样本 {len(created)} 个文件：")
    for path in created:
        print(path)


if __name__ == "__main__":
    main()

