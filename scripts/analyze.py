"""
从 extract.py 输出的长表，计算月度库存均值、月度周转月数、状态分级、品类汇总。

对应交接说明第4节的计算逻辑：
- 4.1 月度平均库存量
- 4.2 月度周转月数（优先用系统字段，缺失时用相邻快照差值估算）
- 4.3 库存状态分级
- 4.4 品类月度汇总
"""
from pathlib import Path

import numpy as np
import pandas as pd

IN_PATH = Path(__file__).resolve().parent.parent / "output" / "snapshots_long.csv"
OUT_DIR = Path(__file__).resolve().parent.parent / "output"

WEEKS_PER_MONTH = 4.33


def classify_status(turnover_months: float) -> str:
    """按交接说明4.3的四级标准分级，NaN返回'未知'"""
    if pd.isna(turnover_months):
        return "未知"
    if turnover_months <= 1:
        return "快速"
    if turnover_months <= 3:
        return "正常"
    if turnover_months <= 6:
        return "偏慢"
    return "积压"


def estimate_turnover_from_snapshots(sku_df: pd.DataFrame, month: str) -> float:
    """
    sku_df: 单个SKU的全部快照，列至少含 date(datetime), qty
    month: 'YYYY-MM'，估算该月的周转月数

    逐对相邻快照计算销售速度（库存下降量/间隔周数），取该月相关区间的
    平均周销售速度 * 4.33 得月销售量，再用月均库存 / 月销售量 得周转月数。
    库存上升（补货）记为0销售，不计负值。
    """
    s = sku_df.sort_values("date").reset_index(drop=True)
    weekly_rates = []
    for i in range(len(s) - 1):
        d0, d1 = s.loc[i, "date"], s.loc[i + 1, "date"]
        # 只统计与目标月份有交集的区间
        if d1.strftime("%Y-%m") != month and d0.strftime("%Y-%m") != month:
            continue
        days = (d1 - d0).days
        if days <= 0:
            continue
        sales = max(0.0, s.loc[i, "qty"] - s.loc[i + 1, "qty"])
        weekly_rates.append(sales / (days / 7))

    if not weekly_rates:
        return np.nan

    avg_weekly_rate = sum(weekly_rates) / len(weekly_rates)
    monthly_sales = avg_weekly_rate * WEEKS_PER_MONTH
    if monthly_sales <= 0:
        return np.nan

    avg_qty = s[s["date"].dt.strftime("%Y-%m") == month]["qty"].mean()
    return avg_qty / monthly_sales


def build_monthly_sku_table(long_df: pd.DataFrame) -> pd.DataFrame:
    """
    输入长表(date, category, sku, name, qty, turnover_months)，
    输出按 sku+month 聚合的月度表：avg_qty, turnover_months(优先系统字段,
    否则用estimate_turnover_from_snapshots估算), status
    """
    df = long_df.copy()
    df["date"] = pd.to_datetime(df["date"])
    df["month"] = df["date"].dt.strftime("%Y-%m")

    rows = []
    for (sku, month), g in df.groupby(["sku", "month"]):
        category = g["category"].iloc[0]
        name = g["name"].iloc[0]
        avg_qty = g["qty"].mean()

        if g["turnover_months"].notna().any():
            turnover = g["turnover_months"].dropna().mean()
        else:
            full_sku_df = df[df["sku"] == sku][["date", "qty"]]
            turnover = estimate_turnover_from_snapshots(full_sku_df, month)

        rows.append(
            {
                "sku": sku,
                "name": name,
                "category": category,
                "month": month,
                "avg_qty": avg_qty,
                "turnover_months": turnover,
                "status": classify_status(turnover),
            }
        )
    return pd.DataFrame(rows)


def build_category_monthly_summary(monthly_sku_df: pd.DataFrame) -> pd.DataFrame:
    """按品类+月份汇总：总库存量、中位周转月数、各状态SKU数"""
    rows = []
    for (category, month), g in monthly_sku_df.groupby(["category", "month"]):
        row = {
            "category": category,
            "month": month,
            "total_qty": g["avg_qty"].sum(),
            "median_turnover_months": g["turnover_months"].median(),
        }
        for status in ["快速", "正常", "偏慢", "积压", "未知"]:
            row[f"count_{status}"] = (g["status"] == status).sum()
        rows.append(row)
    return pd.DataFrame(rows).sort_values(["category", "month"])


if __name__ == "__main__":
    long_df = pd.read_csv(IN_PATH)
    monthly_sku = build_monthly_sku_table(long_df)
    category_summary = build_category_monthly_summary(monthly_sku)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    monthly_sku.to_csv(OUT_DIR / "data_sku_monthly_turnover.csv", index=False, encoding="utf-8-sig")
    category_summary.to_csv(
        OUT_DIR / "data_category_monthly_summary.csv", index=False, encoding="utf-8-sig"
    )
    print(f"SKU月度表 {len(monthly_sku)} 行，品类汇总表 {len(category_summary)} 行，已写入 output/")
