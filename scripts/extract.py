"""
读取 raw_data/ 下所有库存快照xlsx，输出统一长表 CSV。

支持两种格式：
1. 标准周报格式：sheet名为 慕咖/STTOKE/食品，含"商家编码""货品名称""合计"列，
   部分文件还有"周转月数"列。
2. ERP原始导出格式（如 副本库存6_1_1_.xlsx）：sheet名为 库存6.1/蜂蜜和苹果醋/sttoke，
   用"正常库存"列，需按商家编码groupby求和，并过滤"次品标记"==1的行。

输出长表字段：date, category, sku, name, qty, turnover_months(可能为NaN)
"""
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

RAW_DIR = Path(__file__).resolve().parent.parent / "raw_data"
OUT_PATH = Path(__file__).resolve().parent.parent / "output" / "snapshots_long.csv"

STANDARD_SHEETS = {"慕咖": "慕咖", "STTOKE": "STTOKE", "食品": "食品"}

# ERP原始导出sheet名 -> 处理函数在 _parse_erp_file 内分发
ERP_SHEET_NAMES = {"库存6.1", "蜂蜜和苹果醋", "sttoke", "蜂蜜", "慕咖sttoke"}

# sttoke sheet里按SKU前缀判断品类，发现新前缀时请补充此表
SKU_PREFIX_CATEGORY = {
    "STTOKE": "STTOKE",
    "MK": "慕咖",
    "MC": "慕咖",
}


def _extract_date_from_filename(path: Path) -> str:
    """从文件名中提取日期，支持三种格式：
    - 8位 YYYYMMDD，例如 副本20260622库存.xlsx -> 2026-06-22
    - 6位 YYMMDD（无世纪前缀），例如 260624原始数据.xlsx -> 2026-06-24
    - 月.日格式，例如 副本库存6.1(1).xlsx / 副本库存6.29.xlsx -> 2026-06-01 / 2026-06-29
      （年份取文件修改时间的年份）
    """
    m = re.search(r"(20\d{6})", path.stem)
    if m:
        d = m.group(1)
        return f"{d[0:4]}-{d[4:6]}-{d[6:8]}"
    m = re.search(r"(?<!\d)(\d{6})(?!\d)", path.stem)
    if m:
        d = m.group(1)
        return f"20{d[0:2]}-{d[2:4]}-{d[4:6]}"
    # 月.日格式：库存6.29、库存6.1(1) 等
    m = re.search(r"(\d{1,2})\.(\d{1,2})", path.stem)
    if m:
        year = datetime.fromtimestamp(path.stat().st_mtime).year
        month = int(m.group(1))
        day = int(m.group(2))
        return f"{year}-{month:02d}-{day:02d}"
    raise ValueError(f"无法从文件名解析日期: {path.name}")


def _category_from_sku(sku: str) -> str:
    for prefix, cat in SKU_PREFIX_CATEGORY.items():
        if sku.upper().startswith(prefix.upper()):
            return cat
    return "未知"


def _parse_standard_file(path: Path, date: str) -> pd.DataFrame:
    rows = []
    xls = pd.ExcelFile(path)
    for sheet, category in STANDARD_SHEETS.items():
        if sheet not in xls.sheet_names:
            continue
        df = pd.read_excel(path, sheet_name=sheet)
        df = df.dropna(subset=["商家编码"])
        for _, r in df.iterrows():
            rows.append(
                {
                    "date": date,
                    "category": category,
                    "sku": str(r["商家编码"]).strip(),
                    "name": r.get("货品名称", ""),
                    "qty": r.get("合计", 0) or 0,
                    "turnover_months": r.get("周转月数", pd.NA),
                }
            )
    return pd.DataFrame(rows)


def _parse_erp_sheet(path: Path, sheet: str, default_category: str | None) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.dropna(subset=["商家编码"])
    if "次品标记" in df.columns:
        df = df[df["次品标记"] != 1]
    grouped = df.groupby(["商家编码", "货品名称"], as_index=False)["正常库存"].sum()
    rows = []
    for _, r in grouped.iterrows():
        sku = str(r["商家编码"]).strip()
        rows.append(
            {
                "category": default_category or _category_from_sku(sku),
                "sku": sku,
                "name": r["货品名称"],
                "qty": r["正常库存"],
                "turnover_months": pd.NA,
            }
        )
    return pd.DataFrame(rows)


def _parse_erp_file(path: Path, date: str) -> pd.DataFrame:
    xls = pd.ExcelFile(path)

    # "库存6.1"如果存在，是包含全部品类的总表，"蜂蜜和苹果醋"/"sttoke"只是从总表
    # 里又单独拆出来的子集——三个sheet都读会重复计算，所以"库存6.1"存在时只读它，
    # 不存在时才退回去读拆分出来的"蜂蜜和苹果醋"+"sttoke"（比如6/29之后的格式只有
    # 拆分sheet，没有汇总sheet）
    if "库存6.1" in xls.sheet_names:
        df = _parse_erp_sheet(path, "库存6.1", default_category=None)
    else:
        frames = []
        for food_sheet in ["蜂蜜和苹果醋", "蜂蜜"]:
            if food_sheet in xls.sheet_names:
                frames.append(_parse_erp_sheet(path, food_sheet, default_category="食品"))
        for sttoke_sheet in ["sttoke", "慕咖sttoke"]:
            if sttoke_sheet in xls.sheet_names:
                frames.append(_parse_erp_sheet(path, sttoke_sheet, default_category=None))
        df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    if df.empty:
        return df
    df["date"] = date
    return df[["date", "category", "sku", "name", "qty", "turnover_months"]]


def _is_erp_file(path: Path) -> bool:
    xls = pd.ExcelFile(path)
    return bool(ERP_SHEET_NAMES & set(xls.sheet_names))


def parse_file(path: Path) -> pd.DataFrame:
    date = _extract_date_from_filename(path)
    if _is_erp_file(path):
        return _parse_erp_file(path, date)
    return _parse_standard_file(path, date)


def extract_all(raw_dir: Path = RAW_DIR) -> pd.DataFrame:
    files = sorted(raw_dir.glob("*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]
    frames = [parse_file(f) for f in files]
    if not frames:
        return pd.DataFrame(
            columns=["date", "category", "sku", "name", "qty", "turnover_months"]
        )
    return pd.concat(frames, ignore_index=True)


if __name__ == "__main__":
    long_df = extract_all()
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    long_df.to_csv(OUT_PATH, index=False, encoding="utf-8-sig")
    print(f"提取完成，共 {len(long_df)} 行，已写入 {OUT_PATH}")
