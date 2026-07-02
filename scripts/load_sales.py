"""
从 sales_agent 项目的月度 results_YYYYMM.json 读取各平台店铺的 SKU 销量，
按 SKU + 月份聚合（汇总所有店铺），供库存整理系统的销量页面使用。

只取 "正品" 类型的行（排除退货、赠品）的 销量/赠出次数 字段。
"""
import json
from pathlib import Path

# 优先读 raw_data/sales/（用户通过网页上传的文件存这里）
# 其次找本地同级的 sales_agent/data 目录（本地开发时自动读）
_RAW_SALES_DIR = Path(__file__).resolve().parent.parent / "raw_data" / "sales"
_LOCAL_SALES_DIR = Path(__file__).resolve().parent.parent.parent / "sales_agent" / "data"


def _get_sales_dir() -> Path:
    if _RAW_SALES_DIR.exists() and any(_RAW_SALES_DIR.glob("results_??????.json")):
        return _RAW_SALES_DIR
    return _LOCAL_SALES_DIR


def load_monthly_sales() -> dict[str, dict[str, int]]:
    """
    返回 {sku: {month: qty, ...}, ...}
    month 格式 'YYYY-MM'，qty 为正品销量之和（跨所有店铺）
    """
    sales_dir = _get_sales_dir()
    if not sales_dir.exists():
        return {}

    result: dict[str, dict[str, int]] = {}
    name_map: dict[str, str] = {}

    for f in sorted(sales_dir.glob("results_??????.json")):
        stem = f.stem  # results_202606
        ym = stem.split("_")[-1]  # 202606
        if len(ym) != 6:
            continue
        month = f"{ym[:4]}-{ym[4:6]}"

        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue

        for store_val in data.values():
            sku_rows = store_val.get("sheets", {}).get("SKU汇总", [])
            for row in sku_rows:
                if not isinstance(row, dict):
                    continue
                sku = str(row.get("SKU编码", "")).strip()
                if not sku:
                    continue
                item_type = str(row.get("类型", ""))
                if item_type not in ("正品", ""):
                    continue
                qty = row.get("销量/赠出次数", 0)
                try:
                    qty = int(qty) if qty else 0
                except (TypeError, ValueError):
                    qty = 0

                if sku not in result:
                    result[sku] = {}
                result[sku][month] = result[sku].get(month, 0) + qty

                name = str(row.get("商品名称", "")).strip()
                if name and sku not in name_map:
                    name_map[sku] = name

    return result, name_map


def load_monthly_sales_by_store() -> tuple[dict, dict]:
    """
    返回 (store_data, name_map)
    store_data: {store_name: {sku: {month: qty}}}
    name_map: {sku: name}
    """
    sales_dir = _get_sales_dir()
    if not sales_dir.exists():
        return {}, {}

    store_data: dict[str, dict[str, dict[str, int]]] = {}
    name_map: dict[str, str] = {}

    for f in sorted(sales_dir.glob("results_??????.json")):
        ym = f.stem.split("_")[-1]
        if len(ym) != 6:
            continue
        month = f"{ym[:4]}-{ym[4:6]}"

        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue

        for store_name, store_val in data.items():
            sku_rows = store_val.get("sheets", {}).get("SKU汇总", [])
            for row in sku_rows:
                if not isinstance(row, dict):
                    continue
                sku = str(row.get("SKU编码", "")).strip()
                if not sku:
                    continue
                item_type = str(row.get("类型", ""))
                if item_type not in ("正品", ""):
                    continue
                qty = row.get("销量/赠出次数", 0)
                try:
                    qty = int(qty) if qty else 0
                except (TypeError, ValueError):
                    qty = 0

                if store_name not in store_data:
                    store_data[store_name] = {}
                if sku not in store_data[store_name]:
                    store_data[store_name][sku] = {}
                store_data[store_name][sku][month] = store_data[store_name][sku].get(month, 0) + qty

                name = str(row.get("商品名称", "")).strip()
                if name and sku not in name_map:
                    name_map[sku] = name

    return store_data, name_map


def get_sales_months() -> list[str]:
    """返回有数据的月份列表，升序"""
    sales_dir = _get_sales_dir()
    months = []
    for f in sales_dir.glob("results_??????.json"):
        ym = f.stem.split("_")[-1]
        if len(ym) == 6:
            months.append(f"{ym[:4]}-{ym[4:6]}")
    return sorted(set(months))
