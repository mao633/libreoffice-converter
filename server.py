import os, gc, subprocess, tempfile, shutil, time, zipfile, io, sys
from flask import Flask, request, send_file, jsonify, make_response

app = Flask(__name__)
ALLOWED_ORIGIN = "*"


def _add_cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGIN
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Requested-With"
    resp.headers["Access-Control-Max-Age"] = "86400"
    resp.headers["Vary"] = "Origin"
    return resp


@app.before_request
def handle_preflight():
    if request.method == "OPTIONS":
        return _add_cors(make_response("", 204))


@app.after_request
def cors(resp):
    return _add_cors(resp)


@app.errorhandler(Exception)
def on_error(e):
    sys.stderr.write(f"[handler] unhandled: {type(e).__name__}: {e}\n")
    sys.stderr.flush()
    return _add_cors(make_response(jsonify({"error": "internal", "type": type(e).__name__, "msg": str(e)[:300]}), 500))


@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "libreoffice-converter", "version": "5"})


def shrink_pptx(in_bytes, max_dim=1280, quality=70, log=None):
    try:
        from PIL import Image
    except Exception as e:
        if log is not None:
            log.append(f"Pillow_missing:{e}")
        return in_bytes

    try:
        in_zip = zipfile.ZipFile(io.BytesIO(in_bytes))
        out_buf = io.BytesIO()
        media_count = 0
        media_saved = 0
        media_orig_total = 0
        media_new_total = 0

        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as out_zip:
            for item in in_zip.infolist():
                data = in_zip.read(item.filename)
                lname = item.filename.lower()
                is_media_image = lname.startswith("ppt/media/") and (
                    lname.endswith(".png") or lname.endswith(".jpg") or lname.endswith(".jpeg")
                    or lname.endswith(".bmp") or lname.endswith(".tiff") or lname.endswith(".tif")
                )
                if is_media_image:
                    media_count += 1
                    media_orig_total += len(data)
                    try:
                        img = Image.open(io.BytesIO(data))
                        if max(img.size) > max_dim:
                            scale = max_dim / max(img.size)
                            new_size = (max(1, int(img.size[0] * scale)), max(1, int(img.size[1] * scale)))
                            img = img.resize(new_size, Image.LANCZOS)
                        has_alpha = (img.mode in ("RGBA", "LA")) or ("transparency" in img.info)
                        buf = io.BytesIO()
                        if has_alpha and lname.endswith(".png"):
                            img.save(buf, "PNG", optimize=True)
                        else:
                            img.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
                        new_data = buf.getvalue()
                        if len(new_data) < len(data):
                            data = new_data
                            media_saved += 1
                        media_new_total += len(data)
                        del img, buf
                    except Exception as ie:
                        if log is not None and len(log) < 20:
                            log.append(f"img_skip:{item.filename}:{type(ie).__name__}")
                out_zip.writestr(item, data)
                if media_count and media_count % 30 == 0:
                    gc.collect()

        if log is not None:
            log.append(f"media={media_count} reduced={media_saved} {round(media_orig_total/1024/1024,1)}MB->{round(media_new_total/1024/1024,1)}MB")
        result = out_buf.getvalue()
        del out_buf, in_zip
        gc.collect()
        return result
    except Exception as e:
        if log is not None:
            log.append(f"shrink_failed:{type(e).__name__}:{e}")
        return in_bytes


@app.route("/convert", methods=["POST"])
def convert():
    log = []
    t_start = time.time()
    try:
        data = request.files["file"].read() if "file" in request.files else request.get_data()
    except Exception as e:
        return jsonify({"error": "read_failed", "msg": str(e)[:200]}), 400

    if not data:
        return jsonify({"error": "no_file"}), 400
    if len(data) > 250 * 1024 * 1024:
        return jsonify({"error": "too_large", "max_mb": 250}), 413

    in_size_mb = round(len(data) / (1024 * 1024), 2)
    log.append(f"in={in_size_mb}MB")

    if len(data) > 5 * 1024 * 1024:
        if len(data) > 40 * 1024 * 1024:
            data = shrink_pptx(data, max_dim=1024, quality=65, log=log)
        else:
            data = shrink_pptx(data, max_dim=1280, quality=70, log=log)
        gc.collect()
        log.append(f"shrunk={round(len(data)/1024/1024,2)}MB")

    tmpdir = tempfile.mkdtemp(prefix="conv_")
    try:
        in_path = os.path.join(tmpdir, "deck.pptx")
        with open(in_path, "wb") as fp:
            fp.write(data)
        del data
        gc.collect()

        out_dir = os.path.join(tmpdir, "out")
        os.makedirs(out_dir)

        env = os.environ.copy()
        env["HOME"] = tmpdir

        result = subprocess.run(
            [
                "soffice", "--headless", "--norestore", "--nologo",
                "--nodefault", "--nofirststartwizard",
                "-env:UserInstallation=file://" + tmpdir + "/lo_profile",
                "--convert-to",
                "pdf:impress_pdf_Export:UseLosslessCompression=false;ReduceImageResolution=true;MaxImageResolution=150;Quality=78;ExportNotes=false;ExportBookmarks=false",
                "--outdir", out_dir, in_path,
            ],
            capture_output=True, timeout=200, env=env,
        )
        elapsed = round(time.time() - t_start, 1)
        log.append(f"soffice_rc={result.returncode} t={elapsed}s")

        if result.returncode != 0:
            return jsonify({
                "error": "conversion_failed",
                "stderr": result.stderr.decode("utf-8", "ignore")[:500],
                "elapsed_s": elapsed, "in_size_mb": in_size_mb, "log": log,
            }), 500

        out_files = [x for x in os.listdir(out_dir) if x.lower().endswith(".pdf")]
        if not out_files:
            return jsonify({"error": "no_output", "elapsed_s": elapsed, "log": log}), 500

        out_path = os.path.join(out_dir, out_files[0])
        out_size_mb = round(os.path.getsize(out_path) / (1024 * 1024), 2)
        log.append(f"pdf={out_size_mb}MB")
        sys.stderr.write(f"[convert] {' '.join(log)}\n")
        sys.stderr.flush()
        return send_file(out_path, mimetype="application/pdf", as_attachment=False, download_name="converted.pdf")
    except subprocess.TimeoutExpired:
        return jsonify({"error": "timeout", "log": log}), 504
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        gc.collect()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
