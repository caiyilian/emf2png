#!/usr/bin/env python3
"""
server.py — 本地 Web 服务，供 Office 加载项调用。

启动:
    uv run python server.py

然后在 PowerPoint 中侧载 addin/manifest.xml，即可使用。
"""

import sys
import os
import tempfile
import subprocess
import base64
from pathlib import Path
from flask import Flask, request, jsonify, send_from_directory

ROOT = Path(__file__).parent.resolve()
ADDIN_DIR = ROOT / "addin"

app = Flask(__name__, static_folder=None)


# ── 提供加载项的静态文件 ──

@app.route("/")
def index():
    return send_from_directory(ADDIN_DIR, "taskpane.html")


@app.route("/<path:filename>")
def static_files(filename):
    return send_from_directory(ADDIN_DIR, filename)


# ── 转换 API ──

@app.route("/convert", methods=["POST"])
def convert():
    """接收 PPT 文件，返回裁剪白边后的 PDF（Base64）。"""
    data = request.get_json()
    if not data or "file_base64" not in data:
        return jsonify({"error": "missing file_base64"}), 400

    ppt_bytes = base64.b64decode(data["file_base64"])
    filename = data.get("filename", "presentation.pptx")

    tmp_dir = Path(tempfile.mkdtemp())
    ppt_path = tmp_dir / filename
    ppt_path.write_bytes(ppt_bytes)

    pdf_path = tmp_dir / f"{ppt_path.stem}.pdf"

    try:
        result = subprocess.run(
            [sys.executable, str(ROOT / "ppt_to_pdf.py"), str(ppt_path)],
            capture_output=True, text=True, timeout=120,
            cwd=ROOT,
        )

        if result.returncode != 0:
            return jsonify({"error": result.stderr or result.stdout}), 500

        if not pdf_path.exists():
            return jsonify({"error": "PDF not generated"}), 500

        pdf_base64_str = base64.b64encode(pdf_path.read_bytes()).decode()

        import shutil
        shutil.rmtree(tmp_dir, ignore_errors=True)

        return jsonify({
            "pdf_base64": pdf_base64_str,
            "filename": f"{ppt_path.stem}.pdf",
        })

    except subprocess.TimeoutExpired:
        return jsonify({"error": "timeout"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "1.0.2"})


if __name__ == "__main__":
    print("=" * 50)
    print("  emf2png Server started")
    print(f"  http://localhost:5678")
    print("=" * 50)
    print("\nIn PowerPoint: Insert -> My Add-ins -> Upload My Add-in")
    print(f"  Select: {ADDIN_DIR / 'manifest.xml'}")
    print()
    app.run(host="localhost", port=5678, debug=False)
