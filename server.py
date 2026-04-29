import os, subprocess, tempfile, shutil, time
from flask import Flask, request, send_file, jsonify, Response

app = Flask(__name__)

ALLOWED_ORIGINS = "*"

@app.after_request
def cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGINS
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Requested-With"
    return resp

@app.route("/", methods=["GET", "OPTIONS"])
def health():
    return jsonify({"status": "ok", "service": "libreoffice-converter"})

@app.route("/convert", methods=["POST", "OPTIONS"])
def convert():
    if request.method == "OPTIONS":
        return Response(status=204)
    # Accept multipart "file" or raw body
    if "file" in request.files:
        data = request.files["file"].read()
    else:
        data = request.get_data()
    if not data:
        return jsonify({"error": "no_file"}), 400
    if len(data) > 250 * 1024 * 1024:
        return jsonify({"error": "too_large", "max_mb": 250}), 413

    tmpdir = tempfile.mkdtemp(prefix="conv_")
    t0 = time.time()
    try:
        in_path = os.path.join(tmpdir, "deck.pptx")
        with open(in_path, "wb") as fp:
            fp.write(data)
        out_dir = os.path.join(tmpdir, "out")
        os.makedirs(out_dir)
        result = subprocess.run(
            ["soffice", "--headless", "--norestore", "--nologo", "--convert-to", "pdf", "--outdir", out_dir, in_path],
            capture_output=True, timeout=210
        )
        elapsed = time.time() - t0
        if result.returncode != 0:
            return jsonify({
                "error": "conversion_failed",
                "stderr": result.stderr.decode("utf-8", "ignore")[:500],
                "elapsed_s": round(elapsed, 1)
            }), 500
        out_files = [x for x in os.listdir(out_dir) if x.lower().endswith(".pdf")]
        if not out_files:
            return jsonify({"error": "no_output", "elapsed_s": round(elapsed, 1)}), 500
        out_path = os.path.join(out_dir, out_files[0])
        return send_file(out_path, mimetype="application/pdf", as_attachment=False, download_name="converted.pdf")
    except subprocess.TimeoutExpired:
        return jsonify({"error": "timeout"}), 504
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
