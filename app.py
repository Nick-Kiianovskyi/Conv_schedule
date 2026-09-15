# -*- coding: utf-8 -*-
# web_app.py — Веб-інтерфейс конвертера розкладу (Flask)
# Дозволяє завантажити Excel-файл, конвертувати у .ics та завантажити результат.

import io
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import openpyxl
from excel_to_ics import extract_events, save_ics

app = Flask(__name__)
app.secret_key = "schedule-converter-secret-key"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        flash("Файл не обрано", "error")
        return redirect(url_for("index"))

    file = request.files["file"]

    if file.filename == "":
        flash("Файл не обрано", "error")
        return redirect(url_for("index"))

    if not file.filename.endswith((".xlsx", ".xls")):
        flash("Підтримуються лише файли .xlsx та .xls", "error")
        return redirect(url_for("index"))

    try:
        file_bytes = file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active

        events = extract_events(ws)

        if not events:
            flash("Не вдалося знайти жодних подій у файлі. Перевірте формат розкладу.", "error")
            return redirect(url_for("index"))

        output_buffer = io.BytesIO()
        ics_content = generate_ics_string(events)
        output_buffer.write(ics_content.encode("utf-8"))
        output_buffer.seek(0)

        output_name = file.filename.rsplit(".", 1)[0] + ".ics"

        return send_file(
            output_buffer,
            as_attachment=True,
            download_name=output_name,
            mimetype="text/calendar",
        )
    except Exception as e:
        flash(f"Помилка конвертації: {e}", "error")
        return redirect(url_for("index"))


def generate_ics_string(events):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//ExcelToICS//UA//",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-TIMEZONE:Europe/Kyiv",
    ]

    from datetime import datetime
    now = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    for event in events:
        from excel_to_ics import escape_ics
        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{event['uid']}",
            f"DTSTAMP:{now}",
            f"DTSTART:{event['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{event['end'].strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{escape_ics(event['summary'])}",
            f"LOCATION:{escape_ics(event['location'])}",
            f"DESCRIPTION:{escape_ics(event['description'])}",
            "END:VEVENT",
        ])

    lines.append("END:VCALENDAR")
    return "\n".join(lines)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
