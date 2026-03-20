"""
TropicLook Report Server
Принимает Excel INPUT TEMPLATE, возвращает готовый Owner Report Excel.

Make.com отправляет файл сюда → сервер генерирует отчёт → возвращает файл
"""

import os
import tempfile
import traceback
from flask import Flask, request, jsonify, send_file

# Импортируем наш движок
from tl_report_engine import InputData, ReportBuilder

app = Flask(__name__)

# Простая защита: Make.com должен передавать этот токен в заголовке
API_TOKEN = os.environ.get("API_TOKEN", "tropiclook-secret-change-me")


@app.route("/health", methods=["GET"])
def health():
    """Проверка что сервер работает. Make.com может пинговать этот URL."""
    return jsonify({"status": "ok", "service": "TropicLook Report Engine v1.0"})


@app.route("/generate", methods=["POST"])
def generate():
    """
    Принимает INPUT TEMPLATE Excel, генерирует Owner Report Excel.

    Make.com отправляет запрос:
      - Header: X-API-Token: <токен>
      - Body: multipart/form-data с полем "file" = Excel файл
    Возвращает: готовый Excel файл для скачивания
    """

    # Проверяем токен
    token = request.headers.get("X-API-Token", "")
    if token != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401

    # Проверяем что файл передан
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded. Send Excel as 'file' field."}), 400

    uploaded_file = request.files["file"]
    if not uploaded_file.filename.endswith(".xlsx"):
        return jsonify({"error": "File must be .xlsx"}), 400

    # Сохраняем входящий файл во временную папку
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.xlsx")
        output_dir = tmpdir
        uploaded_file.save(input_path)

        try:
            # Читаем данные
            data = InputData(input_path)

            # Валидация
            ok, errors, warnings = data.validate()

            if not ok:
                # Возвращаем список ошибок — бухгалтер получит уведомление
                return jsonify({
                    "status": "validation_failed",
                    "property": data.property_name,
                    "period": data.report_period,
                    "errors": errors,
                    "warnings": warnings
                }), 422

            # Генерируем отчёт
            builder = ReportBuilder(data)
            out_path = builder.build(output_dir)

            # Возвращаем готовый Excel файл
            return send_file(
                out_path,
                as_attachment=True,
                download_name=os.path.basename(out_path),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            # Любая другая ошибка — возвращаем текст для диагностики
            error_detail = traceback.format_exc()
            print(f"ERROR processing {uploaded_file.filename}:\n{error_detail}")
            return jsonify({
                "status": "error",
                "message": str(e),
                "detail": error_detail
            }), 500


@app.route("/validate", methods=["POST"])
def validate_only():
    """
    Только проверяет файл без генерации отчёта.
    Удобно для быстрой проверки перед отправкой.
    """
    token = request.headers.get("X-API-Token", "")
    if token != API_TOKEN:
        return jsonify({"error": "Unauthorized"}), 401

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files["file"]

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, "input.xlsx")
        uploaded_file.save(input_path)

        try:
            data = InputData(input_path)
            ok, errors, warnings = data.validate()

            return jsonify({
                "status": "ok" if ok else "validation_failed",
                "property": data.property_name,
                "property_code": data.property_code,
                "period": data.report_period,
                "owner": data.owner_name,
                "bookings": len(data.reservations),
                "expenses": len(data.expenses),
                "payouts": len(data.payouts),
                "gross_revenue": data.total_gross,
                "net_income": data.net_income,
                "closing_balance": data.closing_balance,
                "errors": errors,
                "warnings": warnings
            })
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
