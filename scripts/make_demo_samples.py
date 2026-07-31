from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DEMO = ROOT / "demo_samples" / "01_深圳住宅演示项目"


def main() -> None:
    DEMO.mkdir(parents=True, exist_ok=True)
    resource = pd.DataFrame(
        [
            ["材料", "C30商品混凝土", "泵送", "m³", 438],
            ["材料", "C35商品混凝土", "泵送", "m³", 455],
            ["材料", "HRB400E钢筋", "", "t", 4150],
            ["人工", "普工", "", "工日", 320],
            ["人工", "技工", "", "工日", 380],
            ["机械", "25t汽车吊", "", "台班", 1500],
            ["机械", "挖掘机", "1m³", "台班", 1350],
            ["材料", "电缆", "WDZ-YJY-4x95+1x50", "m", 285],
            ["设备", "风机", "风量10000m³/h", "台", 5200],
        ],
        columns=["类别", "名称", "规格型号", "单位", "单价"],
    )
    quote = pd.DataFrame(
        [
            ["C30商品混凝土", "泵送", "m³", 445, "含税到场"],
            ["C35商品混凝土", "泵送", "m³", 462, "含税到场"],
            ["HRB400E钢筋", "", "t", 4200, "含税到场"],
            ["25t汽车吊", "", "台班", 1550, "含司机燃油"],
            ["电缆", "WDZ-YJY-4x95+1x50", "m", 292, "含税"],
        ],
        columns=["名称", "规格型号", "单位", "单价", "备注"],
    )
    new_project = pd.DataFrame(
        [
            ["材料", "商品砼", "C30泵送", "m³", ""],
            ["材料", "商品混凝土", "C35泵送", "m³", ""],
            ["材料", "钢筋", "HRB400E", "t", ""],
            ["人工", "普工", "", "工日", ""],
            ["机械", "汽车吊", "25t", "台班", ""],
            ["材料", "电缆", "WDZ-YJY-4x95+1x50", "m", ""],
            ["设备", "风机", "风量10000m³/h", "台", ""],
            ["材料", "未知材料X", "", "kg", ""],
        ],
        columns=["类别", "名称", "规格型号", "单位", "原单价"],
    )
    resource.to_excel(DEMO / "人材机表.xlsx", index=False)
    quote.to_excel(DEMO / "供应商报价单.xlsx", index=False)
    (ROOT / "demo_samples" / "06_新项目补价演示表").mkdir(parents=True, exist_ok=True)
    new_project.to_excel(ROOT / "demo_samples" / "06_新项目补价演示表" / "新项目补价演示表.xlsx", index=False)
    print(DEMO)


if __name__ == "__main__":
    main()

