from __future__ import annotations

import os
import base64
import re
import shutil
import tempfile
import traceback
import time
from datetime import datetime
from typing import Any

import fitz
from flask import Flask, request, jsonify, send_file, send_from_directory

import annotate_pdf

app = Flask(__name__, static_folder=".", static_url_path="")

# --- UPGRADE: Local File Caching to replace Base64 ---
STATIC_CACHE_DIR = os.path.join(tempfile.gettempdir(), "pdf_studio_cache")
os.makedirs(STATIC_CACHE_DIR, exist_ok=True)

@app.route("/cache/<path:filename>")
def serve_cache(filename):
    return send_from_directory(STATIC_CACHE_DIR, filename)
# -----------------------------------------------------

_uploaded_files = {
    "excel": None,
    "master_pdf": None,
    "slave_pdf": None,
    "master_pdf_name": None,
    "slave_pdf_name": None,
}

_HEX_RGB_RE = re.compile(r"^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")
_HEX_RGBA_RE = re.compile(r"^#(?:[0-9a-fA-F]{8})$")

_FLOAT_CONFIG_KEYS = {
    "CIRCLE_SCALE",
    "CIRCLE_MAX_DIAMETER",
    "CIRCLE_MAX_DIAMETER_PAX",
    "CIRCLE_LINE_WIDTH",
    "CIRCLE_OFFSET_X_PAX",
    "CIRCLE_OFFSET_Y_PAX",
    "CIRCLE_OFFSET_X_TP",
    "CIRCLE_OFFSET_Y_TP",
    "LABEL_FONT_SIZE",
    "LABEL_LINE_SPACING",
    "LABEL_OFFSET_X_PAX_FIRST_HALF",
    "LABEL_OFFSET_Y_PAX_FIRST_HALF",
    "LABEL_OFFSET_X_PAX_SECOND_HALF",
    "LABEL_OFFSET_Y_PAX_SECOND_HALF",
    "LABEL_OFFSET_X_TP",
    "LABEL_OFFSET_Y_TP",
    "LABEL_BG_ALPHA",
    "LABEL_BG_PADDING_X",
    "LABEL_BG_PADDING_Y",
}
_INT_CONFIG_KEYS = {"LABEL_WRAP_CHARS"}
_BOOL_CONFIG_KEYS = {"LABEL_BG_ENABLED"}
_STR_ENUM_CONFIG_KEYS = {"LABEL_ATTACH_TP": {"top", "bottom", "left", "right"}}
_COLOR_CONFIG_KEYS = {"RED", "LABEL_BG_COLOR_HEX"}


def _to_float(value: Any, default: float = 0.0) -> float:
    try:
        f = float(value)
        if f != f or f in (float("inf"), float("-inf")):
            return default
        return f
    except (TypeError, ValueError):
        return default


def _to_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def _is_valid_color(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    s = value.strip()
    return bool(_HEX_RGB_RE.fullmatch(s) or _HEX_RGBA_RE.fullmatch(s))


def _sanitize_config(raw_cfg: Any) -> dict[str, Any]:
    if not isinstance(raw_cfg, dict):
        return {}

    cleaned: dict[str, Any] = {}
    for key, value in raw_cfg.items():
        if key in _FLOAT_CONFIG_KEYS:
            cleaned[key] = _to_float(value)
        elif key in _INT_CONFIG_KEYS:
            cleaned[key] = max(0, _to_int(value))
        elif key in _BOOL_CONFIG_KEYS:
            cleaned[key] = bool(value)
        elif key in _COLOR_CONFIG_KEYS and _is_valid_color(value):
            cleaned[key] = str(value).strip()
        elif key in _STR_ENUM_CONFIG_KEYS:
            normalized = str(value).strip().lower()
            if normalized in _STR_ENUM_CONFIG_KEYS[key]:
                cleaned[key] = normalized

    return cleaned


def _sanitize_manual_offsets(raw_offsets: Any) -> dict[str, dict[str, float]]:
    if not isinstance(raw_offsets, dict):
        return {}

    cleaned: dict[str, dict[str, float]] = {}
    for label, offset in raw_offsets.items():
        if not isinstance(offset, dict):
            continue
        dx = _to_float(offset.get("dx"), 0.0)
        dy = _to_float(offset.get("dy"), 0.0)
        cleaned[str(label)] = {"dx": dx, "dy": dy}
    return cleaned


def _parse_page_param(value: Any, default: int = 1) -> int:
    page = _to_int(value, default)
    return max(0, page)

def _save_upload(file_storage, suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    file_storage.save(path)
    return path

def _safe_name(name: str) -> str:
    bad = '<>:"/\\|?*'
    return "".join("_" if ch in bad else ch for ch in str(name)).strip()


def _stats_status(stats: dict) -> str:
    total = int(stats.get("total", 0) or 0)
    circled = int(stats.get("circled", 0) or 0)
    missing = int(stats.get("not_found", 0) or 0)

    if total == 0:
        return "NO_DATA"
    if missing == 0:
        return "SUCCESS"
    if circled > 0:
        return "PARTIAL_SUCCESS"
    return "FAILURE"

def _merge_status(overall: str, status: str) -> str:
    if status == "FAILURE":
        return "FAILURE"
    if status == "PARTIAL_SUCCESS" and overall != "FAILURE":
        return "PARTIAL_SUCCESS"
    return overall

# --- UPGRADE: Save to local disk instead of Base64 ---
def _render_page_to_url(page: fitz.Page, zoom: float = 2.0) -> str:
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    filename = f"render_{int(time.time() * 1000)}.png"
    filepath = os.path.join(STATIC_CACHE_DIR, filename)
    pix.save(filepath)
    return f"/cache/{filename}"

def _get_pdf_path(pdf_type: str) -> str | None:
    return _uploaded_files["slave_pdf"] if pdf_type == "slave" else _uploaded_files["master_pdf"]

def _err(e: Exception):
    return jsonify({"status": "ERROR", "error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/")
def index():
    return send_file("index.html", mimetype="text/html")


@app.route("/api/parse-excel", methods=["POST"])
def parse_excel():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No Excel file provided"}), 400

        file = request.files["file"]
        if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return jsonify({"error": "Unsupported Excel format"}), 400

        uploaded_path = _save_upload(file, suffix=os.path.splitext(file.filename)[1] or ".xlsx")
        _uploaded_files["excel"] = uploaded_path
        projects = annotate_pdf.detect_projects(uploaded_path, sheet_name="overview DSUB")

        return jsonify({
            "status": "SUCCESS",
            "projects": projects,
            "file": file.filename,
            "excel_name": file.filename,
            "current_path": uploaded_path,
        })
    except Exception as e:
        return _err(e)


@app.route("/api/upload-pdf", methods=["POST"])
def upload_pdf():
    try:
        if "file" not in request.files or "type" not in request.form:
            return jsonify({"error": "Missing file or type"}), 400

        file = request.files["file"]
        pdf_type = request.form["type"].strip().lower()

        if pdf_type not in {"master", "slave"}:
            return jsonify({"error": "type must be master or slave"}), 400
        if not file.filename or not file.filename.lower().endswith(".pdf"):
            return jsonify({"error": "Only PDF files are supported"}), 400

        path = _save_upload(file, suffix=".pdf")

        if pdf_type == "master":
            _uploaded_files["master_pdf"] = path
            _uploaded_files["master_pdf_name"] = file.filename
        else:
            _uploaded_files["slave_pdf"] = path
            _uploaded_files["slave_pdf_name"] = file.filename

        doc = fitz.open(path)
        page_count = len(doc)
        doc.close()

        return jsonify({
            "status": "SUCCESS", "type": pdf_type, "file": file.filename, "page_count": page_count,
        })
    except Exception as e:
        return _err(e)


@app.route("/api/categorize", methods=["POST"])
def categorize():
    try:
        if not _uploaded_files["excel"]:
            return jsonify({"error": "No Excel file uploaded"}), 400

        payload = request.get_json(force=True) or {}
        project_name = payload.get("project_name") or payload.get("project")
        if not project_name:
            return jsonify({"error": "No project selected"}), 400

        pairs, skipped = annotate_pdf.read_excel_data(
            xlsx_path=_uploaded_files["excel"], sheet_name="overview DSUB", project_name=project_name,
        )
        return jsonify(annotate_pdf.categorize_pairs(pairs, skipped))
    except Exception as e:
        return _err(e)


@app.route("/api/preview", methods=["GET"])
def preview():
    try:
        pdf_type = request.args.get("pdf", "master").strip().lower()
        pdf_path = _get_pdf_path(pdf_type)
        if not pdf_path:
            return jsonify({"error": f"No {pdf_type} PDF uploaded"}), 400

        doc = fitz.open(pdf_path)
        page_num = _parse_page_param(request.args.get("page", "1"), default=1)

        if page_num < 0 or page_num >= len(doc):
            doc.close()
            return jsonify({"error": "Invalid page number"}), 400

        page = doc[page_num]
        response = {
            "status": "SUCCESS",
            "image": _render_page_to_url(page),
            "page_count": len(doc),
            "pdf_width": page.rect.width,
            "pdf_height": page.rect.height,
        }
        doc.close()
        return jsonify(response)
    except Exception as e:
        return _err(e)


@app.route("/api/preview-annotated", methods=["POST"])
def preview_annotated():
    try:
        pdf_type = request.args.get("pdf", "master").strip().lower()
        page_num = _parse_page_param(request.args.get("page", "1"), default=1)

        payload = request.get_json(force=True) or {}
        config = _sanitize_config(payload.get("config", {}))
        pairs = payload.get("pairs", [])
        manual_offsets = _sanitize_manual_offsets(payload.get("manual_offsets", {}))

        pdf_path = _get_pdf_path(pdf_type)
        if not pdf_path:
            return jsonify({"error": f"No {pdf_type} PDF uploaded"}), 400

        annotate_pdf.apply_config(config)

        doc = fitz.open(pdf_path)

        if page_num < 0 or page_num >= len(doc):
            doc.close()
            return jsonify({"error": "Invalid page number"}), 400

        page = doc[page_num]

        stats = annotate_pdf.circle_page(page, pairs, manual_offsets=manual_offsets)
        stats["status"] = _stats_status(stats)

        response = {
            "status": stats["status"],
            "image": _render_page_to_url(page),
            "page_count": len(doc),
            "pdf_width": page.rect.width,
            "pdf_height": page.rect.height,
            "locations": stats.get("locations", {}),
            "not_found_labels": stats.get("not_found_labels", []),
            "resolved_aliases": stats.get("resolved_aliases", {}),
            "debug_logs": stats.get("debug_logs", []),
            "stats": stats,
        }

        doc.close()
        return jsonify(response)

    except Exception as e:
        return _err(e)


@app.route("/api/generate", methods=["POST"])
def generate():
    try:
        payload = request.get_json(force=True) or {}
        
        master_pdf = _uploaded_files.get("master_pdf")
        slave_pdf = _uploaded_files.get("slave_pdf")
        
        if not master_pdf and not slave_pdf:
            return jsonify({"error": "No PDFs uploaded for generation."}), 400

        config = _sanitize_config(payload.get("config", {}))
        config_master = _sanitize_config(payload.get("config_master", config))
        config_slave = _sanitize_config(payload.get("config_slave", config))
        pairs_master = payload.get("pairs_master", [])
        pairs_slave = payload.get("pairs_slave", [])
        project_name = payload.get("project_name") or "project"
        page_idx = _parse_page_param(payload.get("page_idx", 1), default=1)

        manual_offsets_master = _sanitize_manual_offsets(payload.get("manual_offsets_master", {}))
        manual_offsets_slave = _sanitize_manual_offsets(payload.get("manual_offsets_slave", {}))

        requested_dir = (payload.get("output_dir") or "").strip()
        save_server_side = bool(requested_dir)
        work_dir = requested_dir if save_server_side else tempfile.mkdtemp(prefix="pdfann_")
        os.makedirs(work_dir, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_project = _safe_name(project_name)

        output_paths = {}
        files = {}
        role_stats = {"master": None, "slave": None}
        overall_status = "SUCCESS"

        targets = [
            ("master", master_pdf, config_master, _uploaded_files.get("master_pdf_name") or "master.pdf", pairs_master, manual_offsets_master),
            ("slave", slave_pdf, config_slave, _uploaded_files.get("slave_pdf_name") or "slave.pdf", pairs_slave, manual_offsets_slave),
        ]

        try:
            for role, pdf_path, cfg, raw_name, pairs, offsets in targets:
                if not pdf_path:
                    continue
                annotate_pdf.apply_config(cfg)
                filename = f"annotated_{safe_project}_{ts}_{_safe_name(raw_name)}"
                out_path = os.path.join(work_dir, filename)
                stats = annotate_pdf.annotate_pdf_file(
                    pdf_path, out_path, pairs, page_idx=page_idx, manual_offsets=offsets
                )
                stats["status"] = _stats_status(stats)
                role_stats[role] = stats
                overall_status = _merge_status(overall_status, stats["status"])

                with open(out_path, "rb") as fh:
                    files[role] = {"filename": filename, "b64": base64.b64encode(fh.read()).decode("utf-8")}
                if save_server_side:
                    output_paths[role] = out_path
        finally:
            if not save_server_side:
                shutil.rmtree(work_dir, ignore_errors=True)

        return jsonify({
            "status": overall_status,
            "master_stats": role_stats["master"] or {},
            "slave_stats": role_stats["slave"] or {},
            "output_paths": output_paths,
            "files": files,
        })

    except Exception as e:
        return _err(e)


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)