import os, subprocess, tempfile, shutil, time, zipfile, io
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
    return jsonify({"status": "ok", "service": "libreoffice-converter", "version": "2"})

def shrink_pptx(in_bytes, max_image_kb=300):
    """Pre-compress images inside a .pptx to keep memory usage manageable on free tier.
    Uses Pillow only when available; otherwise leaves the file as-is."""
    try:
        from PIL import Image
    except Exception:
        return in_bytes
    try:
        in_zip = zipfile.ZipFile(io.BytesIO(in_bytes))
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for item in in_zip.infolist():
                data = in_zip.read(item.filename)
                lname = item.filename.lower()
                if lname.startswith("ppt/media/") and (lname.endswith(".png") or lname.endswith(".jpg") or lname.endswith(".jpeg")):
                    if len(data) > max_image_kb * 1024:
                        try:
                            img = Image.open(io.BytesIO(data))
                            img = img.convert("RGB") if img.mode in ("RGBA", "P") and lname.endswith((".jpg", ".jpeg")) else img
                            # Cap dimensions to 1600px on the long side
                            max_dim = 1600
                            if max(img.size) > max_dim:
                                scale = max_dim / max(img.size)
                                img = img.resize((int(img.size[0]*scale), int(img.size[1]*scale)), Image.LANCZOS)
                            buf = io.BytesIO()
                            if lname.endswith(".png"):
                                # Convert PNG to JPEG when possible to save lots of space
                                try:
                                    img.convert("RGB").save(buf, "JPEG", quality=78, optimize=True)
                                    data = buf.getvalue()
                                except Exception:
                                    img.save(buf, "PNG", optimize=True)
                                    data = buf.getvalue()
                            else:
                                img.save(buf, "JPEG", quality=78, optimize=True)
                                data = buf.getvalue()
                        except Exception:
                            pass
                out_zip.writestr(item, data)
        return out_buf.getvalue()
    except Exception:
        return in_bytes

@app.route("/convert", methods=["POST", "OPTIONS"])
def convert():
    if request.method == "OPTIONS":
        return Response(status=204)
    data = request.files["file"].read() if "file" in request.files else request.get_data()
    if not data:
        return jsonify({"error": "no_file"}), 400
    if len(data) > 250 * 1024 * 1024:
        return jsonify({"error": "too_large", "max_mb": 250}), 413

    in_size_mb = round(len(data) / (1024*1024), 1)
    # Pre-compress if file is huge (free tier RAM is tight)
    if len(data) > 30 * 1024 * 1024:
        data = shrink_pptx(data)

    tmpdir = tempfile.mkdtemp(prefix="conv_")
    t0 = time.time()
    try:
        in_path = os.path.join(tmpdir, "deck.pptx")
        with open(in_path, "wb") as fp:
            fp.write(data)
        out_dir = os.path.join(tmpdir, "out")
        os.makedirs(out_dir)
        result = subprocess.run(
            ["soffice", "--headless", "--norestore", "--nologo",
             "--convert-to", "pdf:impress_pdf_Export:UseLosslessCompression=false;ReduceImageResolution=true;MaxImageResolution=150;Quality=80",
             "--outdir", out_dir, in_path],
            capture_output=True, timeout=210
        )
        elapsed = round(time.time() - t0, 1)
        if result.returncode != 0:
            return jsonify({"error": "conversion_failed",
                            "stderr": result.stderr.decode("utf-8", "ignore")[:500],
                            "elapsed_s": elapsed,
                            "in_size_mb": in_size_mb}), 500
        out_files = [x for x in os.listdir(out_dir) if x.lower().endswith(".pdf")]
        if not out_files:
            return jsonify({"error": "no_output", "elapsed_s": elapsed}), 500
        out_path = os.path.join(out_dir, out_files[0])
        return send_file(out_path, mimetype="application/pdf", as_attachment=False, download_name="converted.pdf")
    except subprocess.TimeoutExpired:
        return jsonify({"error": "timeout"}), 504
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
