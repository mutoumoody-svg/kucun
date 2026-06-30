"""
从 analyze.py 产出的 output/data_sku_monthly_turnover.csv 里取最新月份的周转月数/状态，
给 clean_*.py 的库存整理表附加"最近一次算出的周转率"作参考。

注意：周转率是按月度快照（每周库存文件）算出来的历史指标，跟某一天的库存整理表
不是同一时间维度，只能反映"最近趋势"，不代表当天库存的精确周转——所以列名标注
"最新月周转月数"，并带上数据来自哪个月份，避免被误读成跟当天库存量精确对应。
"""
from pathlib import Path

import pandas as pd

TURNOVER_PATH = Path(__file__).resolve().parent.parent / "output" / "data_sku_monthly_turnover.csv"


def load_latest_turnover() -> pd.DataFrame:
    """返回 sku, 最新月周转月数, 库存状态, 周转数据月份 四列。文件不存在则返回空表。"""
    if not TURNOVER_PATH.exists():
        return pd.DataFrame(columns=["sku", "最新月周转月数", "库存状态", "周转数据月份"])

    df = pd.read_csv(TURNOVER_PATH)
    latest_month = df["month"].max()
    latest = df[df["month"] == latest_month][["sku", "turnover_months", "status"]].copy()
    latest = latest.rename(columns={"turnover_months": "最新月周转月数", "status": "库存状态"})
    latest["周转数据月份"] = latest_month
    return latest.drop_duplicates(subset=["sku"])
