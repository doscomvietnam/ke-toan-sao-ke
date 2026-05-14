# -*- coding: utf-8 -*-
"""Serverless function cho Vercel — nhận file sao kê (POST multipart), trả file kết quả."""

import os
import sys
from datetime import datetime
from io import BytesIO

from flask import Flask, request, send_file, jsonify

# Cho phép import kt.py ở thư mục gốc
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT_DIR)

from kt import process, print_log  # noqa: E402

app = Flask(__name__)

NGAN_HANG_PATH = os.path.join(ROOT_DIR, "ngan-hang.xlsx")
MA_THONG_KE_PATH = os.path.join(ROOT_DIR, "ma-thong-ke.xlsx")
TEMPLATE_PATH = os.path.join(ROOT_DIR, "Template-dau-ra.xlsx")


@app.route("/api/process", methods=["POST"])
def handle_process():
    if "file" not in request.files:
        return jsonify({"error": "Thiếu trường file"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Tên file rỗng"}), 400

    input_bytes = f.read()
    if not input_bytes:
        return jsonify({"error": "File rỗng"}), 400

    try:
        output_bytes, log_rows = process(
            input_bytes,
            ngan_hang_path=NGAN_HANG_PATH,
            ma_thong_ke_path=MA_THONG_KE_PATH,
            template_path=TEMPLATE_PATH,
        )
    except Exception as e:
        return jsonify({"error": f"Lỗi xử lý: {e}"}), 500

    # In nhật ký xử lý ra terminal (server stderr)
    print(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] Xử lý: {f.filename}", file=sys.stderr)
    print_log(log_rows)

    orig_name = os.path.splitext(os.path.basename(f.filename))[0]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"Ket-qua_{orig_name}_{stamp}.xlsx"

    response = send_file(
        BytesIO(output_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["X-Log-Count"] = str(len(log_rows))
    return response


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "have_ngan_hang": os.path.exists(NGAN_HANG_PATH),
        "have_ma_thong_ke": os.path.exists(MA_THONG_KE_PATH),
        "have_template": os.path.exists(TEMPLATE_PATH),
    })


# Phục vụ index.html khi chạy local (Vercel sẽ tự serve static, không qua route này)
@app.route("/", methods=["GET"])
def index():
    index_path = os.path.join(ROOT_DIR, "index.html")
    if os.path.exists(index_path):
        return send_file(index_path)
    return "index.html không tồn tại", 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
