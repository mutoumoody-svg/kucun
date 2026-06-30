"""
库存整理工具的网页后端。

单页应用：上传一个新的库存快照xlsx，后端自动识别格式（260624_style/erp_style）、
提取快照日期、登记进 raw_data/snapshot_registry.json，然后跑一遍
scripts/clean_latest.py 的整理流程，把结果xlsx存到 output/，页面上直接展示分析
摘要（品牌数量、库存状态分布、临期预警、低库存预警），并提供完整结果下载链接。

每次成功整理的结果会按快照日期归档到 output/history/{date}.html 和
output/history/{date}.xlsx，左侧栏列出所有归档过的日期，点击可以查看/下载
历史上某一期的整理结果，不会因为后面又传了新快照就丢掉旧记录。

启动方式（本地测试）：
    cd Inventory/webapp
    uvicorn main:app --host 0.0.0.0 --port 8000
"""
import html
import shutil
import sys
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
RAW_DIR = PROJECT_ROOT / "raw_data"
OUTPUT_DIR = PROJECT_ROOT / "output"
HISTORY_DIR = OUTPUT_DIR / "history"

sys.path.insert(0, str(SCRIPTS_DIR))

import clean_latest  # noqa: E402
from categorize import STATUS_COLORS  # noqa: E402

app = FastAPI(title="库存整理工具")

PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>库存整理工具</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: -apple-system, "Microsoft YaHei", sans-serif; margin: 0; color: #1f2937; background: #fafafa; }}
  .layout {{ display: flex; min-height: 100vh; }}
  .sidebar {{ width: 220px; flex-shrink: 0; background: #1f2937; color: #d1d5db; padding: 20px 0; }}
  .sidebar h3 {{ font-size: 13px; color: #9ca3af; padding: 0 16px; margin: 0 0 8px; text-transform: uppercase; }}
  .sidebar a {{ display: block; padding: 8px 16px; color: #d1d5db; text-decoration: none; font-size: 14px; }}
  .sidebar a:hover {{ background: #374151; }}
  .sidebar a.active {{ background: #2563eb; color: #fff; font-weight: 600; }}
  .sidebar .empty {{ padding: 0 16px; font-size: 13px; color: #6b7280; }}
  .main {{ flex: 1; max-width: 980px; margin: 32px auto; padding: 0 16px; }}
  h1 {{ font-size: 22px; }}
  h2 {{ font-size: 16px; margin: 28px 0 10px; }}
  .card {{ border: 1px solid #e5e7eb; border-radius: 8px; padding: 24px; margin-top: 16px; background: #fff; }}
  input[type=file] {{ margin: 12px 0; }}
  button {{ background: #1f2937; color: #fff; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-size: 14px; }}
  button:hover {{ background: #374151; }}
  .error {{ color: #b91c1c; background: #fef2f2; border-radius: 6px; padding: 12px; }}
  a.download {{ display: inline-block; margin-top: 4px; background: #2563eb; color: #fff; padding: 10px 20px; border-radius: 6px; text-decoration: none; }}
  label {{ display: block; margin-top: 8px; font-size: 14px; color: #4b5563; }}
  .meta {{ color: #6b7280; font-size: 13px; margin-bottom: 8px; }}
  .stat-grid {{ display: flex; flex-wrap: wrap; gap: 12px; }}
  .stat-box {{ flex: 1; min-width: 140px; border: 1px solid #e5e7eb; border-radius: 8px; padding: 14px 16px; background: #f9fafb; }}
  .stat-box .num {{ font-size: 24px; font-weight: 700; }}
  .stat-box .label {{ font-size: 13px; color: #6b7280; margin-top: 2px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 8px; }}
  th, td {{ border: 1px solid #e5e7eb; padding: 6px 10px; text-align: left; white-space: nowrap; }}
  th {{ background: #1f2937; color: #fff; }}
  .empty-note {{ color: #9ca3af; font-size: 13px; padding: 8px 0; }}
  .more-note {{ color: #9ca3af; font-size: 12px; margin-top: 6px; }}
</style>
</head>
<body>
  <div class="layout">
    <div class="sidebar">
      <h3>历史整理记录</h3>
      {sidebar_links}
    </div>
    <div class="main">
      <h1>库存整理工具</h1>
      <div class="card">
        <form action="/upload" method="post" enctype="multipart/form-data">
          <label>选这次的库存快照xlsx文件：</label>
          <input type="file" name="file" accept=".xlsx" required>
          <label>如果文件本身不带库存日期（比如260624原始数据.xlsx这种单Sheet1格式），手动填一下日期（可不填，默认用文件修改日期）：</label>
          <input type="date" name="snapshot_date">
          <br><br>
          <button type="submit">上传并整理</button>
        </form>
      </div>
      {result_block}
    </div>
  </div>
</body>
</html>
"""


def list_history_dates() -> list[str]:
    if not HISTORY_DIR.exists():
        return []
    dates = [p.stem for p in HISTORY_DIR.glob("*.html")]
    return sorted(dates, reverse=True)


def build_sidebar(active_date: str | None = None) -> str:
    dates = list_history_dates()
    if not dates:
        return '<div class="empty">还没有整理记录</div>'
    links = []
    for d in dates:
        cls = ' class="active"' if d == active_date else ""
        links.append(f'<a href="/history/{d}"{cls}>{d}</a>')
    return "".join(links)


def _render(result_block: str = "", active_date: str | None = None) -> HTMLResponse:
    return HTMLResponse(
        PAGE_TEMPLATE.format(result_block=result_block, sidebar_links=build_sidebar(active_date))
    )


def _esc(v) -> str:
    if pd.isna(v):
        return ""
    return html.escape(str(v))


def _table_html(df: pd.DataFrame, columns: list[str], max_rows: int = 12, status_col: str | None = None) -> str:
    if df.empty:
        return '<div class="empty-note">无数据</div>'
    shown = df.head(max_rows)
    head = "".join(f"<th>{_esc(c)}</th>" for c in columns)
    rows = []
    for _, r in shown.iterrows():
        cells = []
        for c in columns:
            val = r.get(c, "")
            style = ""
            if status_col and c == status_col:
                color = STATUS_COLORS.get(val)
                if color:
                    style = f' style="background:#{color}"'
            cells.append(f"<td{style}>{_esc(val)}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    more = ""
    if len(df) > max_rows:
        more = f'<div class="more-note">共 {len(df)} 条，仅展示前 {max_rows} 条，完整结果请下载Excel</div>'
    return f"<table><thead><tr>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table>{more}"


QTY_BRANDS = [("STTOKE", "STTOKE"), ("慕咖", "慕咖"), ("食品", "巴恩天然")]


def build_result_html(tables: dict, entry: tuple | None) -> str:
    goods = tables["goods"]
    packaging = tables["packaging"]
    snapshot_date = tables["snapshot_date"]

    qty_boxes = ""
    if "汇总" in goods.columns:
        for brand_key, label in QTY_BRANDS:
            total_qty = int(goods.loc[goods["品牌"] == brand_key, "汇总"].sum())
            qty_boxes += (
                f'<div class="stat-box"><div class="num">{total_qty:,}</div>'
                f'<div class="label">{_esc(label)} 总库存数量</div></div>'
            )

    goods_counts = goods["品牌"].value_counts()
    pkg_counts = packaging["品牌"].value_counts()

    stat_boxes = "".join(
        f'<div class="stat-box"><div class="num">{v}</div><div class="label">{_esc(k)} 商品</div></div>'
        for k, v in goods_counts.items()
    )
    stat_boxes += "".join(
        f'<div class="stat-box"><div class="num">{v}</div><div class="label">{_esc(k)} 包装物料</div></div>'
        for k, v in pkg_counts.items()
    )

    status_pivot = (
        goods[goods["库存状态"].notna()]
        .groupby(["品牌", "库存状态"])
        .size()
        .unstack(fill_value=0)
        if "库存状态" in goods.columns and goods["库存状态"].notna().any()
        else pd.DataFrame()
    )
    status_cols = [c for c in ["快速", "正常", "偏慢", "积压", "未知"] if c in status_pivot.columns]
    status_table = ""
    if not status_pivot.empty:
        rows = []
        head = "<th>品牌</th>" + "".join(f"<th>{c}</th>" for c in status_cols)
        for brand, row in status_pivot.iterrows():
            cells = f"<td>{_esc(brand)}</td>"
            for c in status_cols:
                bg = STATUS_COLORS.get(c, "")
                cells += f'<td style="background:#{bg}">{int(row[c])}</td>'
            rows.append(f"<tr>{cells}</tr>")
        status_table = f"<table><thead><tr>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table>"
    else:
        status_table = '<div class="empty-note">无周转状态数据</div>'

    expiring = tables.get("expiring_soon", pd.DataFrame())
    expiring_cols = [c for c in ["货品名称", "货品编号", "汇总", "批次号", "过期日期", "距离到期天数_最近批次"] if c in expiring.columns]
    expiring_html = _table_html(expiring, expiring_cols, max_rows=12)

    low_stock = tables.get("low_stock", pd.DataFrame())
    low_stock_html = ""
    if low_stock is not None and len(low_stock):
        low_cols = [c for c in low_stock.columns if c in ("货品名称", "货品编号", "品牌", "汇总")]
        low_stock_html = f"""
        <h2>低库存预警（汇总≤10）</h2>
        {_table_html(low_stock, low_cols, max_rows=12)}
        """

    meta_line = f"快照日期：{snapshot_date}（格式：{tables['format']}）"
    if entry:
        meta_line = f"刚登记的快照：{_esc(entry[2])}（日期 {entry[0]}） · " + meta_line

    download_name = f"库存整理_{snapshot_date}.xlsx"

    return f"""
    <div class="card">
      <div class="meta">{meta_line}</div>
      <a class="download" href="/download/{snapshot_date}?name={download_name}">下载完整整理结果（含所有品牌sheet）</a>

      <h2>总库存数量统计</h2>
      <div class="stat-grid">{qty_boxes}</div>

      <h2>品牌数量概览</h2>
      <div class="stat-grid">{stat_boxes}</div>

      <h2>库存状态分布（按品牌）</h2>
      {status_table}

      <h2>临期预警（1年内到期，按紧迫度排序）</h2>
      {expiring_html}

      {low_stock_html}
    </div>
    """


def save_history(snapshot_date: str, result_html: str) -> None:
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    (HISTORY_DIR / f"{snapshot_date}.html").write_text(result_html, encoding="utf-8")
    shutil.copy(clean_latest.OUT_PATH, HISTORY_DIR / f"{snapshot_date}.xlsx")


@app.get("/", response_class=HTMLResponse)
def index():
    dates = list_history_dates()
    if dates:
        latest = dates[0]
        result_html = (HISTORY_DIR / f"{latest}.html").read_text(encoding="utf-8")
        return _render(result_html, active_date=latest)
    return _render()


@app.get("/history/{date}", response_class=HTMLResponse)
def history(date: str):
    path = HISTORY_DIR / f"{date}.html"
    if not path.exists():
        raise HTTPException(404, "没有这一期的记录")
    return _render(path.read_text(encoding="utf-8"), active_date=date)


@app.post("/upload", response_class=HTMLResponse)
async def upload(file: UploadFile = File(...), snapshot_date: str = Form(default="")):
    if not file.filename.lower().endswith(".xlsx"):
        return _render('<div class="error">只支持 .xlsx 文件</div>')

    RAW_DIR.mkdir(parents=True, exist_ok=True)
    dest_path = RAW_DIR / file.filename
    with open(dest_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        entry = clean_latest.register_snapshot(dest_path, date_override=snapshot_date or None)
        tables = clean_latest.run_pipeline()
    except Exception as e:
        dest_path.unlink(missing_ok=True)
        return _render(f'<div class="error">处理失败：{html.escape(str(e))}</div>')

    result_html = build_result_html(tables, entry)
    save_history(tables["snapshot_date"], result_html)

    return _render(result_html, active_date=tables["snapshot_date"])


@app.get("/download/{date}")
def download(date: str, name: str = "库存整理.xlsx"):
    path = HISTORY_DIR / f"{date}.xlsx"
    if not path.exists():
        raise HTTPException(404, "文件不存在")
    return FileResponse(path, filename=name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
