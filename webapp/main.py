"""
库存整理工具的网页后端。

单页应用：上传一个新的库存快照xlsx，后端自动识别格式（260624_style/erp_style）、
提取快照日期、登记进 raw_data/snapshot_registry.json，然后跑一遍
scripts/clean_latest.py 的整理流程，把结果xlsx存到 output/，页面上提供下载链接和
一个简单的整理摘要（各品牌商品/包装物料数量）。

启动方式（本地测试）：
    cd Inventory/webapp
    uvicorn main:app --host 0.0.0.0 --port 8000

部署到 kucun.riverline.com.cn 还需要：
- 一台能跑Python的服务器/容器，把这个目录部署上去并装好 requirements.txt 里的依赖
- 用 gunicorn/uvicorn + systemd（或类似的进程管理）常驻运行，前面挂nginx反向代理
- 域名DNS解析到这台服务器，nginx配置该域名转发到本服务监听的端口
这几步涉及服务器/DNS的实际操作权限，需要你那边配合提供环境或者授权访问。
"""
import shutil
import sys
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
RAW_DIR = PROJECT_ROOT / "raw_data"
OUTPUT_DIR = PROJECT_ROOT / "output"

sys.path.insert(0, str(SCRIPTS_DIR))

import clean_latest  # noqa: E402

app = FastAPI(title="库存整理工具")

PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>库存整理工具</title>
<style>
  body {{ font-family: -apple-system, "Microsoft YaHei", sans-serif; max-width: 720px; margin: 40px auto; padding: 0 16px; color: #1f2937; }}
  h1 {{ font-size: 22px; }}
  .card {{ border: 1px solid #e5e7eb; border-radius: 8px; padding: 24px; margin-top: 16px; }}
  input[type=file] {{ margin: 12px 0; }}
  button {{ background: #1f2937; color: #fff; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-size: 14px; }}
  button:hover {{ background: #374151; }}
  .summary {{ white-space: pre-wrap; background: #f9fafb; border-radius: 6px; padding: 16px; font-size: 14px; line-height: 1.6; }}
  .error {{ color: #b91c1c; background: #fef2f2; border-radius: 6px; padding: 12px; }}
  a.download {{ display: inline-block; margin-top: 12px; background: #2563eb; color: #fff; padding: 10px 20px; border-radius: 6px; text-decoration: none; }}
  label {{ display: block; margin-top: 8px; font-size: 14px; color: #4b5563; }}
</style>
</head>
<body>
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
</body>
</html>
"""


def _render(result_block: str = "") -> HTMLResponse:
    return HTMLResponse(PAGE_TEMPLATE.format(result_block=result_block))


@app.get("/", response_class=HTMLResponse)
def index():
    return _render()


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
        return _render(f'<div class="error">处理失败：{e}</div>')

    goods_counts = tables["goods"]["品牌"].value_counts().to_dict()
    pkg_counts = tables["packaging"]["品牌"].value_counts().to_dict()

    summary_lines = [
        f"已登记快照：{entry[2]}（日期 {entry[0]}，格式 {entry[1]}）",
        f"本次使用的最新快照：{tables['snapshot_date']}（格式 {tables['format']}）",
        "",
        "商品数量（按品牌）：",
        *[f"  {k}: {v}" for k, v in goods_counts.items()],
        "",
        "包装物料数量（按品牌）：",
        *[f"  {k}: {v}" for k, v in pkg_counts.items()],
    ]
    summary = "\n".join(summary_lines)

    # 用一个带时间戳的token让下载链接每次都指向这次刚生成的文件
    token = uuid.uuid4().hex[:8]
    download_name = f"库存整理_{tables['snapshot_date']}.xlsx"
    cached_path = OUTPUT_DIR / f"_dl_{token}.xlsx"
    shutil.copy(clean_latest.OUT_PATH, cached_path)

    block = f"""
    <div class="card">
      <div class="summary">{summary}</div>
      <a class="download" href="/download/{token}?name={download_name}">下载整理结果</a>
    </div>
    """
    return _render(block)


@app.get("/download/{token}")
def download(token: str, name: str = "库存整理.xlsx"):
    path = OUTPUT_DIR / f"_dl_{token}.xlsx"
    if not path.exists():
        raise HTTPException(404, "文件不存在或已过期")
    return FileResponse(path, filename=name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
