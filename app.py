import os
import tempfile
from flask import Flask, request, jsonify, send_file
from tl_report_engine import generate_report

app = Flask(__name__)

API_TOKEN = os.environ.get("API_TOKEN", "tropiclook-secret-change-me")


def _check_token():
    token = request.headers.get("X-API-Token", "")
    return token == API_TOKEN


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/generate", methods=["POST"])
def generate():
    if not _check_token():
        return jsonify({"error": "Unauthorized"}), 401

    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    uploaded = request.files["file"]
    if not uploaded.filename:
        return jsonify({"error": "Empty filename"}), 400

    # Save input to temp file
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
        uploaded.save(tmp_in.name)
        input_path = tmp_in.name

    output_path = input_path.replace(".xlsx", "_report.xlsx")

    try:
        warnings = generate_report(input_path, output_path)
    except ValueError as e:
        os.unlink(input_path)
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        os.unlink(input_path)
        return jsonify({"error": f"Internal error: {str(e)}"}), 500
    finally:
        try:
            os.unlink(input_path)
        except Exception:
            pass

    return send_file(
        output_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="owner_report.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
