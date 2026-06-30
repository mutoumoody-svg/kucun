"""
计算"当月累计出库数量" = 本月第一个快照的库存量 - as_of_date当天快照的库存量。
直接拿月初库存减最新库存，不是跟上一期快照比，也不是逐段下降量累加。

当月已知的快照（按时间顺序），文件格式各不相同，这里各自写了一个小的加载函数，
统一转成 {sku: qty} 字典再做对比：
- 2026-06-01：副本库存6.1(1).xlsx（ERP格式，库存6.1/蜂蜜和苹果醋/sttoke三个sheet）
- 2026-06-08：副本20260608库存.xlsx（标准周报格式，慕咖/STTOKE/食品三个sheet）
- 2026-06-22：副本20260622库存.xlsx（标准周报格式）
- 2026-06-24：260624原始数据.xlsx（当前库存盘点导出，按仓库逐行展开）
- 2026-06-29：副本库存6.29.xlsx（ERP格式，蜂蜜/sttoke两个sheet）

以后每多一期新快照，只要按格式加进 SNAPSHOT_SOURCES 列表即可，不需要改计算逻辑。
"""
from pathlib import Path

import pandas as pd

from categorize import SKU_ALIAS

RAW_DIR = Path(__file__).resolve().parent.parent / "raw_data"


def _load_erp_qty(path: Path, sheets: list[str]) -> dict:
    qty = {}
    for sheet in sheets:
        df = pd.read_excel(path, sheet_name=sheet)
        if "次品标记" in df.columns:
            df = df[df["次品标记"] != 1]
        df = df.dropna(subset=["商家编码"])
        g = df.groupby("商家编码")["正常库存"].sum()
        for sku, q in g.items():
            sku = SKU_ALIAS.get(sku, sku)
            qty[sku] = qty.get(sku, 0) + q
    return qty


def _load_standard_qty(path: Path, sheets: list[str]) -> dict:
    qty = {}
    for sheet in sheets:
        df = pd.read_excel(path, sheet_name=sheet)
        df = df.dropna(subset=["商家编码"])
        g = df.groupby("商家编码")["合计"].sum()
        for sku, q in g.items():
            sku = SKU_ALIAS.get(sku, sku)
            qty[sku] = qty.get(sku, 0) + q
    return qty


def _load_260624_qty(path: Path) -> dict:
    df = pd.read_excel(path, sheet_name="Sheet1")
    g = df.groupby("货品编号")["库存量"].sum()
    return {SKU_ALIAS.get(sku, sku): q for sku, q in g.items()}


SNAPSHOT_SOURCES = [
    # 注意：副本库存6.1(1).xlsx里"库存6.1"sheet本身就是总表（352条），
    # "蜂蜜和苹果醋"(46条)和"sttoke"(306条)只是从总表里又单独拆出来的子集，
    # 全部都包含在"库存6.1"里——只能读"库存6.1"一个sheet，不然会重复计算。
    ("2026-06-01", lambda: _load_erp_qty(RAW_DIR / "副本库存6.1(1).xlsx", ["库存6.1"])),
    ("2026-06-08", lambda: _load_standard_qty(RAW_DIR / "副本20260608库存.xlsx", ["慕咖", "STTOKE", "食品"])),
    ("2026-06-22", lambda: _load_standard_qty(RAW_DIR / "副本20260622库存.xlsx", ["慕咖", "STTOKE", "食品"])),
    ("2026-06-24", lambda: _load_260624_qty(RAW_DIR / "260624原始数据.xlsx")),
    ("2026-06-29", lambda: _load_erp_qty(RAW_DIR / "副本库存6.29.xlsx", ["蜂蜜", "sttoke"])),
    ("2026-06-30", lambda: _load_erp_qty(RAW_DIR / "副本库存6.30.xlsx", ["蜂蜜", "慕咖sttoke"])),
]


def cumulative_month_outbound(as_of_date: str) -> pd.DataFrame:
    """返回 sku, 累计出库数量, 出库数据周期 三列。
    累计出库数量 = 本月第一个快照的库存量 - as_of_date当天快照的库存量
    （正数=本月净出库，负数=本月净补货增加）。
    """
    month = as_of_date[:7]
    sources = [(d, loader) for d, loader in SNAPSHOT_SOURCES if d[:7] == month and d <= as_of_date]
    sources = sorted(sources, key=lambda x: x[0])

    available = [(d, loader()) for d, loader in sources]
    available = [(d, q) for d, q in available if q]

    if len(available) < 2:
        return pd.DataFrame(columns=["sku", "累计出库数量", "出库数据周期"])

    start_date, start_qty = available[0]
    end_date, end_qty = available[-1]

    # 月初快照里压根没有这个SKU的，说明这个SKU是月中才新增/才开始有记录的，
    # 没有真实的月初基准库存，不能拿0去硬算（那样会显示出一个虚假的"补货"负数），
    # 这种情况直接不给出累计出库数量，留空
    common_skus = set(start_qty.keys()) & set(end_qty.keys())
    outbound = {sku: start_qty[sku] - end_qty[sku] for sku in common_skus}

    period = f"{start_date} ~ {end_date}"
    return pd.DataFrame(
        {"sku": list(outbound.keys()), "累计出库数量": list(outbound.values())}
    ).assign(出库数据周期=period)
