#!/usr/bin/env python3
import json
import mimetypes
import os
import re
import socket
import subprocess
import sys
import tempfile
import time
import uuid
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
WORKBENCH_DIR = ROOT / "workbench"
NODE_BIN = Path("/Users/cc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node")
NODE_MODULES = Path("/Users/cc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules")
CONFIG_PATH = ROOT / "config" / "products" / "blister_pad_bundle.json"
PRICING_CONFIG_PATH = ROOT / "config" / "pricing_config.json"
WORKBENCH_STATE_PATH = ROOT / "config" / "workbench_state.json"
HISTORY_ORDER_PATH = ROOT / "config" / "pricing_history_order.json"
OUTPUT_DIR = ROOT / "output"
INPUT_DIR = ROOT / "input"
UPLOAD_TARGETS = {
    "purchase_order": INPUT_DIR / "purchase_orders",
    "product_links": INPUT_DIR / "product_links",
    "dimensions": INPUT_DIR / "dimensions",
    "competitors": INPUT_DIR / "competitors",
}
ALLOWED_SUFFIXES = {
    "purchase_order": {".pdf", ".xlsx", ".xls", ".html", ".htm", ".txt", ".csv", ".png", ".jpg", ".jpeg", ".webp"},
    "product_links": {".txt", ".csv", ".tsv", ".xlsx", ".xls", ".html", ".htm"},
    "dimensions": {".xlsx", ".xls"},
    "competitors": {".html", ".htm"},
}


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path, data):
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def clean(value):
    if value is None:
        return ""
    return str(value).strip()


def default_workbench_state():
    return {
        "product_links": [""],
        "manual_dimensions": [
            {
                "sku": "",
                "title": "",
                "length_cm": "",
                "width_cm": "",
                "height_cm": "",
                "weight_g": "",
            }
        ],
    }


def read_workbench_state():
    if not WORKBENCH_STATE_PATH.exists():
        return default_workbench_state()
    state = read_json(WORKBENCH_STATE_PATH)
    fallback = default_workbench_state()
    return {
        "product_links": state.get("product_links") or fallback["product_links"],
        "manual_dimensions": state.get("manual_dimensions") or fallback["manual_dimensions"],
    }


def write_workbench_state(data):
    state = default_workbench_state()
    if isinstance(data.get("product_links"), list):
        state["product_links"] = [str(item).strip() for item in data["product_links"]]
    if isinstance(data.get("manual_dimensions"), list):
        rows = []
        for item in data["manual_dimensions"]:
            if isinstance(item, dict):
                rows.append(
                    {
                        "sku": str(item.get("sku", "")).strip(),
                        "title": str(item.get("title", "")).strip(),
                        "length_cm": item.get("length_cm", ""),
                        "width_cm": item.get("width_cm", ""),
                        "height_cm": item.get("height_cm", ""),
                        "weight_g": item.get("weight_g", ""),
                    }
                )
        state["manual_dimensions"] = rows or state["manual_dimensions"]
    WORKBENCH_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    write_json(WORKBENCH_STATE_PATH, state)
    return state


def workbook_rows(path, sheet_name, max_rows=50):
    if not path.exists():
        return []
    workbook = load_workbook(path, data_only=True)
    if sheet_name not in workbook.sheetnames:
        return []
    sheet = workbook[sheet_name]
    headers = [cell.value for cell in sheet[1]]
    rows = []
    for values in sheet.iter_rows(min_row=2, max_row=min(sheet.max_row, max_rows + 1), values_only=True):
        if not any(value is not None and value != "" for value in values):
            continue
        rows.append({headers[index]: value for index, value in enumerate(values) if index < len(headers)})
    return rows


def read_report(path):
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")


def output_state(product_batch):
    draft = OUTPUT_DIR / f"{product_batch}_统一输入草稿.xlsx"
    result = OUTPUT_DIR / f"{product_batch}_正式定价结果.xlsx"
    report = OUTPUT_DIR / f"{product_batch}_正式定价报告.md"
    upload = OUTPUT_DIR / f"{product_batch}_上品系统导入.xlsx"
    draft_rows = workbook_rows(draft, "统一输入草稿", 20)
    result_rows = workbook_rows(result, "建议售价", 20)
    draft_by_sku = {row.get("SKU"): row for row in draft_rows if row.get("SKU")}
    for row in result_rows:
        source = draft_by_sku.get(row.get("SKU"), {})
        for key in ["包装长cm", "包装宽cm", "包装高cm", "包装重量g"]:
            if key in source:
                row[key] = source[key]
    return {
        "draft": {
            "path": str(draft),
            "exists": draft.exists(),
            "rows": draft_rows,
            "competitors": workbook_rows(draft, "竞品解析", 20),
        },
        "result": {
            "path": str(result),
            "exists": result.exists(),
            "rows": result_rows,
            "scenarios": workbook_rows(result, "价格档位", 20),
        },
        "report": {
            "path": str(report),
            "exists": report.exists(),
            "content": read_report(report),
        },
        "upload": {
            "path": str(upload),
            "exists": upload.exists(),
            "rows": workbook_rows(upload, "上品系统导入", 20),
        },
    }


def get_value(row, *keys):
    for key in keys:
        if key in row and row.get(key) not in (None, ""):
            return row.get(key)
    return ""


def output_file(batch, *suffixes):
    for suffix in suffixes:
        path = OUTPUT_DIR / f"{batch}_{suffix}"
        if path.exists():
            return path
    return None


def read_history_order():
    if not HISTORY_ORDER_PATH.exists():
        return {"version": 1, "products": {}}
    try:
        data = read_json(HISTORY_ORDER_PATH)
    except (OSError, json.JSONDecodeError):
        return {"version": 1, "products": {}}
    products = data.get("products") if isinstance(data, dict) else {}
    return {"version": 1, "products": products if isinstance(products, dict) else {}}


def write_history_order(data):
    HISTORY_ORDER_PATH.parent.mkdir(parents=True, exist_ok=True)
    write_json(HISTORY_ORDER_PATH, data)


def history_batches():
    batches = {}
    for path in OUTPUT_DIR.glob("*.xlsx"):
        if path.name.startswith(("~$", ".~")):
            continue
        for suffix in ("正式定价结果.xlsx", "自动定价结果.xlsx"):
            marker = f"_{suffix}"
            if path.name.endswith(marker):
                batches[path.name[: -len(marker)]] = path
    return batches


def history_record(batch, result_path, order_info=None):
    draft_path = output_file(batch, "统一输入草稿.xlsx")
    upload_path = output_file(batch, "上品系统导入.xlsx")
    report_path = output_file(batch, "正式定价报告.md", "自动定价报告.md")
    draft_rows = workbook_rows(draft_path, "统一输入草稿", 200) if draft_path else []
    result_rows = workbook_rows(result_path, "建议售价", 200)
    upload_rows = workbook_rows(upload_path, "上品系统导入", 200) if upload_path else []
    competitor_rows = workbook_rows(result_path, "竞品", 200)
    result_by_sku = {row.get("SKU"): row for row in result_rows if row.get("SKU")}
    upload_by_sku = {row.get("SKU"): row for row in upload_rows if row.get("SKU")}
    source_rows = draft_rows or result_rows
    rows = []
    for index, row in enumerate(source_rows):
        sku = get_value(row, "SKU")
        result = result_by_sku.get(sku, row)
        title = get_value(row, "中文品名", "英文品名") or batch
        spec = get_value(row, "变体/规格", "规格") or title or f"采购款式-{index + 1}"
        rows.append(
            {
                "sku": sku,
                "itemCode": "",
                "title": title,
                "spec": spec,
                "quantity": get_value(row, "采购数量"),
                "cost": get_value(row, "采购单价RMB/包", "单件成本RMB") or get_value(result, "单件成本RMB"),
            }
        )
    final_rows = []
    for row in result_rows:
        sku = get_value(row, "SKU")
        upload = upload_by_sku.get(sku, {})
        final_rows.append(
            {
                "sku": sku,
                "title": get_value(row, "规格"),
                "price": get_value(row, "建议售价USD", "建议售价", "建议整包售价"),
                "weightLb": get_value(row, "计费重量lb"),
                "fbaFee": get_value(row, "FBA费"),
                "profitRate": get_value(row, "期望利润率"),
                "sizeCm": get_value(upload, "包装尺寸cm"),
                "weightGText": get_value(upload, "包装重量g"),
            }
        )
    first_draft = draft_rows[0] if draft_rows else {}
    first_upload = upload_rows[0] if upload_rows else {}
    dims = {
        "lengthCm": get_value(first_draft, "包装长cm"),
        "widthCm": get_value(first_draft, "包装宽cm"),
        "heightCm": get_value(first_draft, "包装高cm"),
        "weightG": get_value(first_draft, "包装重量g"),
    }
    if not any(dims.values()) and first_upload.get("包装尺寸cm"):
        parts = [part.strip() for part in str(first_upload.get("包装尺寸cm")).replace("×", "*").split("*")]
        if len(parts) >= 3:
            dims["lengthCm"], dims["widthCm"], dims["heightCm"] = parts[:3]
        dims["weightG"] = str(first_upload.get("包装重量g") or "").replace("g", "").strip()
    competitors = []
    for index, row in enumerate(competitor_rows):
        title = get_value(row, "标题", "文件") or f"竞品{index + 1}"
        competitors.append(
            {
                "label": get_value(row, "ASIN") or str(title)[:28],
                "title": title,
                "price": get_value(row, "页面主售价USD", "页面主售价"),
                "packCount": get_value(row, "包数") or 1,
                "rating": "评分待确认",
                "sales": "销量待确认",
            }
        )
    prices = [row.get("price") for row in final_rows if row.get("price") not in (None, "")]
    fees = [row.get("fbaFee") for row in final_rows if row.get("fbaFee") not in (None, "")]
    costs = [row.get("cost") for row in rows if row.get("cost") not in (None, "")]
    return {
        "name": batch,
        "source": "output",
        "firstCalculatedAt": (order_info or {}).get("firstCalculatedAt", result_path.stat().st_ctime),
        "orderSequence": (order_info or {}).get("sequence", 0),
        "updatedAt": result_path.stat().st_mtime,
        "purchaseCostRmb": costs[0] if len(set(map(str, costs))) == 1 and costs else " / ".join(str(item) for item in costs[:3]) or "待补",
        "shippingFee": fees[0] if len(set(map(str, fees))) == 1 and fees else " / ".join(str(item) for item in fees[:3]) or "待补",
        "finalPrice": prices[0] if len(set(map(str, prices))) == 1 and prices else " / ".join(str(item) for item in prices[:3]) or "待补",
        "rows": rows,
        "finalRows": final_rows,
        "dimensions": dims,
        "salePackQty": get_value(first_draft, "销售包数") or 1,
        "comparisonQty": get_value(first_draft, "对比单位数量") or 1,
        "targetMargin": get_value(first_draft, "目标最低利润率") or "",
        "competitors": competitors,
        "report": read_report(report_path) if report_path else "",
        "files": {
            "draft": str(draft_path) if draft_path else "",
            "result": str(result_path),
            "report": str(report_path) if report_path else "",
            "upload": str(upload_path) if upload_path else "",
        },
    }


def output_history():
    records = []
    batches = sorted(history_batches().items())
    order_data = read_history_order()
    products = order_data["products"]
    changed = False
    next_sequence = max(
        [int(item.get("sequence", 0)) for item in products.values() if isinstance(item, dict)] or [0]
    )
    for batch, result_path in batches:
        if not isinstance(products.get(batch), dict):
            next_sequence += 1
            products[batch] = {
                "firstCalculatedAt": result_path.stat().st_ctime or time.time(),
                "sequence": next_sequence,
            }
            changed = True
    if changed:
        write_history_order(order_data)
    for batch, result_path in batches:
        try:
            records.append(history_record(batch, result_path, products.get(batch)))
        except Exception as exc:
            records.append({"name": batch, "source": "output", "error": str(exc)})
    return sorted(
        records,
        key=lambda item: (item.get("firstCalculatedAt", 0), item.get("orderSequence", 0)),
        reverse=True,
    )


def safe_filename(name):
    cleaned = Path(name or "upload").name.replace("\x00", "").strip()
    return cleaned or f"upload-{uuid.uuid4().hex[:8]}"


def input_files():
    state = {}
    for key, directory in UPLOAD_TARGETS.items():
        directory.mkdir(parents=True, exist_ok=True)
        files = []
        for path in directory.iterdir():
            if not path.is_file() or path.name.startswith(("~$", ".~", ".DS_Store")):
                continue
            try:
                stat = path.stat()
            except OSError:
                continue
            files.append(
                {
                    "name": path.name,
                    "path": str(path),
                    "size": stat.st_size,
                    "mtime": stat.st_mtime,
                }
            )
        state[key] = [
            {key_name: item[key_name] for key_name in ("name", "path", "size")}
            for item in sorted(files, key=lambda item: item["mtime"], reverse=True)
        ]
    return state


def parse_multipart(headers, body):
    content_type = headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        raise ValueError("请使用表单上传文件")
    parser_headers = (
        f"Content-Type: {content_type}\r\n"
        f"MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8")
    message = BytesParser(policy=default).parsebytes(parser_headers + body)
    fields = {}
    files = []
    for part in message.iter_parts():
        disposition = part.get_content_disposition()
        if disposition != "form-data":
            continue
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        if filename:
            files.append({"field": name, "filename": filename, "payload": payload})
        elif name:
            fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
    return fields, files


def save_uploads(headers, body):
    fields, files = parse_multipart(headers, body)
    file_type = fields.get("type", "")
    if file_type not in UPLOAD_TARGETS:
        raise ValueError("未知资料类型")
    target_dir = UPLOAD_TARGETS[file_type]
    target_dir.mkdir(parents=True, exist_ok=True)
    allowed = ALLOWED_SUFFIXES[file_type]
    saved = []
    for item in files:
        original_name = safe_filename(item["filename"])
        suffix = Path(original_name).suffix.lower()
        if suffix not in allowed:
            raise ValueError(f"{original_name} 的格式不支持")
        target = target_dir / original_name
        if target.exists():
            target = target_dir / f"{target.stem}_{uuid.uuid4().hex[:6]}{target.suffix}"
        target.write_bytes(item["payload"])
        saved.append({"name": target.name, "path": str(target), "size": target.stat().st_size})
    if not saved:
        raise ValueError("没有收到文件")
    return saved


def run_image_ocr(headers, body):
    fields, files = parse_multipart(headers, body)
    if not files:
        raise ValueError("没有收到图片")
    item = files[0]
    original_name = safe_filename(item["filename"])
    suffix = Path(original_name).suffix.lower()
    if suffix not in {".png", ".jpg", ".jpeg", ".webp"}:
        raise ValueError(f"{original_name} 不是支持的图片格式")
    with tempfile.NamedTemporaryFile(prefix="price-ocr-", suffix=suffix, delete=False) as tmp:
        tmp.write(item["payload"])
        tmp_path = Path(tmp.name)
    try:
        node_bin = NODE_BIN if NODE_BIN.exists() else Path("node")
        env = os.environ.copy()
        if NODE_MODULES.exists():
            env["NODE_PATH"] = str(NODE_MODULES)
        completed = subprocess.run(
            [str(node_bin), str(WORKBENCH_DIR / "ocr_node.js"), str(tmp_path)],
            cwd=ROOT,
            text=True,
            capture_output=True,
            timeout=180,
            env=env,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or "OCR 识别失败")
        payload = json.loads(completed.stdout or "{}")
        return payload.get("text", "")
    finally:
        try:
            tmp_path.unlink()
        except FileNotFoundError:
            pass


def extract_links_from_file(path):
    links = []
    if path.suffix.lower() in {".xlsx", ".xls"}:
        workbook = load_workbook(path, data_only=True)
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=True):
                for value in row:
                    links.extend(re.findall(r"https?://\S+", str(value or "")))
    else:
        text = path.read_text(encoding="utf-8", errors="ignore")
        links.extend(re.findall(r"https?://\S+", text))
    cleaned = []
    for link in links:
        link = link.strip().strip('",;，；')
        if link and link not in cleaned:
            cleaned.append(link)
    return cleaned


def merge_product_links(saved):
    if not saved:
        return
    state = read_workbench_state()
    links = [link for link in state.get("product_links", []) if clean(link)]
    for item in saved:
        for link in extract_links_from_file(Path(item["path"])):
            if link not in links:
                links.append(link)
    if not links:
        links = [""]
    state["product_links"] = links
    write_workbench_state(state)


def run_script(script):
    completed = subprocess.run(
        [sys.executable, str(ROOT / script)],
        cwd=ROOT,
        text=True,
        capture_output=True,
        timeout=120,
    )
    return {
        "ok": completed.returncode == 0,
        "returncode": completed.returncode,
        "stdout": completed.stdout.strip(),
        "stderr": completed.stderr.strip(),
    }


def draft_script_for_mode(mode):
    if mode == "legacy_bundle":
        return "scripts/build_blister_pad_draft.py"
    if mode == "legacy_generic":
        return "scripts/generate_input_draft.py"
    return "scripts/build_workbench_draft.py"


def api_state():
    product = read_json(CONFIG_PATH)
    pricing = read_json(PRICING_CONFIG_PATH)
    workbench = read_workbench_state()
    purchase_order = ROOT / product["purchase_order"]
    competitor_dir = ROOT / product["competitor_dir"]
    competitor_files = []
    if competitor_dir.exists():
        competitor_files = [path.name for path in sorted(competitor_dir.glob("*.html"))]
    return {
        "product": product,
        "pricing": pricing,
        "sourceFiles": {
            "purchaseOrder": {
                "path": str(purchase_order),
                "exists": purchase_order.exists(),
            },
            "competitorDir": {
                "path": str(competitor_dir),
                "exists": competitor_dir.exists(),
                "files": competitor_files,
            },
        },
        "inputFiles": input_files(),
        "workbench": workbench,
        "outputs": output_state(product["product_batch"]),
        "history": output_history(),
    }


class WorkbenchHandler(BaseHTTPRequestHandler):
    server_version = "PriceWorkbench/1.0"

    def log_message(self, format, *args):
        return

    def send_json(self, data, status=200):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def send_file(self, path):
        if not path.exists() or not path.is_file():
            self.send_error(404)
            return
        content = path.read_bytes()
        mime_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", mime_type)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        parsed = urlparse(self.path)
        route = parsed.path
        if route == "/api/state":
            self.send_json(api_state())
            return
        if route == "/api/history":
            self.send_json({"records": output_history()})
            return
        if route.startswith("/files/output/"):
            name = unquote(route.removeprefix("/files/output/"))
            path = (OUTPUT_DIR / name).resolve()
            if OUTPUT_DIR.resolve() not in path.parents and path != OUTPUT_DIR.resolve():
                self.send_error(403)
                return
            self.send_file(path)
            return
        if route in {"/", "/index.html"}:
            self.send_file(WORKBENCH_DIR / "index.html")
            return
        if route in {"/styles.css", "/app.js"}:
            self.send_file(WORKBENCH_DIR / route.removeprefix("/"))
            return
        self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length) if length else b"{}"
        if parsed.path == "/api/config":
            try:
                body = raw_body.decode("utf-8")
                payload = json.loads(body)
                write_json(CONFIG_PATH, payload["product"])
                self.send_json({"ok": True, "state": api_state()})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
            return
        if parsed.path == "/api/upload":
            try:
                saved = save_uploads(self.headers, raw_body)
                if saved and saved[0]["path"].startswith(str(UPLOAD_TARGETS["product_links"])):
                    merge_product_links(saved)
                self.send_json({"ok": True, "saved": saved, "state": api_state()})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
            return
        if parsed.path == "/api/ocr":
            try:
                text = run_image_ocr(self.headers, raw_body)
                self.send_json({"ok": True, "text": text})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
            return
        if parsed.path == "/api/workbench":
            try:
                body = raw_body.decode("utf-8")
                payload = json.loads(body)
                write_workbench_state(payload.get("workbench", {}))
                self.send_json({"ok": True, "state": api_state()})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
            return
        if parsed.path == "/api/generate-draft":
            try:
                body = raw_body.decode("utf-8")
                payload = json.loads(body) if body.strip() else {}
            except json.JSONDecodeError:
                payload = {}
            result = run_script(draft_script_for_mode(payload.get("mode")))
            self.send_json({"ok": result["ok"], "command": result, "state": api_state()}, status=200 if result["ok"] else 500)
            return
        if parsed.path == "/api/run-pricing":
            try:
                body = raw_body.decode("utf-8")
                payload = json.loads(body) if body.strip() else {}
            except json.JSONDecodeError:
                payload = {}
            draft = run_script(draft_script_for_mode(payload.get("mode")))
            pricing = run_script("scripts/generate_final_pricing.py") if draft["ok"] else None
            ok = draft["ok"] and pricing and pricing["ok"]
            self.send_json(
                {"ok": ok, "draft": draft, "pricing": pricing, "state": api_state()},
                status=200 if ok else 500,
            )
            return
        self.send_error(404)


def main():
    host = "127.0.0.1"
    preferred_port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    port = preferred_port
    if len(sys.argv) == 1:
        for candidate in range(preferred_port, preferred_port + 25):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
                if probe.connect_ex((host, candidate)) != 0:
                    port = candidate
                    break
    server = ThreadingHTTPServer((host, port), WorkbenchHandler)
    print(f"工作台已启动：http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
