#!/usr/bin/env python3
"""
เว็บแอปแปลงไฟล์งบทดลอง (.xls) 2 ไฟล์ (เดือนก่อนหน้า + เดือนปัจจุบัน)
เป็นไฟล์ตาม template_ยอดยกมา-ยอดยกไป-.xlsx
รันบนเครื่องตัวเอง ไม่ผูกกับระบบ intranet และไม่ใช้ฐานข้อมูลอีกต่อไป
(ยอดยกมาได้มาจากไฟล์เดือนก่อนหน้าที่อัปโหลดโดยตรง ผ่านสูตร Excel ในเทมเพลต)

การใช้งาน:
    pip install -r requirements.txt
    python app.py            # แล้วเปิด http://127.0.0.1:5000
"""

import io
import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

from flask import Flask, render_template, request, send_file

import trial_balance as tb

app = Flask(__name__)

# เครื่องมือ single-user รันเครื่องเดียว - เก็บผลลัพธ์ล่าสุดไว้ในหน่วยความจำพอ ไม่ต้องใช้ session store
LAST_RESULT = None


def _read_upload(field_name, label):
    """อ่านไฟล์อัปโหลดจาก field ที่กำหนด คืน (parsed_info, error_message)"""
    upload = request.files.get(field_name)
    if not upload or not upload.filename:
        return None, "กรุณาเลือกไฟล์{0}ก่อนครับ".format(label)

    fname = upload.filename.lower()
    if not (fname.endswith(".xls") or fname.endswith(".xlsx")):
        return None, "ไฟล์{0}ไม่ใช่ไฟล์ Excel ครับ".format(label)

    try:
        return tb.parse_source_file(upload), None
    except Exception as e:
        return None, "ไม่สามารถอ่านไฟล์{0}ได้: {1}".format(label, e)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", result=None, error=None)


@app.route("/process", methods=["POST"])
def process():
    global LAST_RESULT

    previous_info, err = _read_upload("previous_file", "เดือนก่อนหน้า")
    if err:
        return render_template("index.html", result=None, error=err)

    current_info, err = _read_upload("current_file", "เดือนปัจจุบัน")
    if err:
        return render_template("index.html", result=None, error=err)

    previous_parsed = previous_info["parsed"]
    current_parsed = current_info["parsed"]

    only_in_previous, only_in_current = tb.diff_codes(previous_parsed, current_parsed)

    total_debit = sum(v["debit"] for v in current_parsed.values())
    total_credit = sum(v["credit"] for v in current_parsed.values())

    LAST_RESULT = {
        "previous_parsed": previous_parsed,
        "current_parsed": current_parsed,
        "be_month": current_info["be_month"],
        "be_year": current_info["be_year"],
        "filename": tb.output_filename(current_info["be_month"], current_info["be_year"]),
    }

    summary = {
        "previous_month_name": tb.thai_month_name(previous_info["be_month"]),
        "previous_year": previous_info["be_year"],
        "month_name": tb.thai_month_name(current_info["be_month"]),
        "year": current_info["be_year"],
        "previous_found_codes": len(previous_parsed),
        "found_codes": len(current_parsed),
        "total_debit": total_debit,
        "total_credit": total_credit,
        "only_in_previous": only_in_previous,
        "only_in_current": only_in_current,
        "malformed_previous": previous_info["malformed"],
        "malformed_current": current_info["malformed"],
        "filename": LAST_RESULT["filename"],
    }

    return render_template("index.html", result=summary, error=None)


@app.route("/download", methods=["POST"])
def download():
    global LAST_RESULT

    if LAST_RESULT is None:
        return render_template("index.html", result=None,
                                error="กรุณาอัปโหลดและกดประมวลผลไฟล์ก่อนครับ")

    wb = tb.build_output_workbook(LAST_RESULT["previous_parsed"], LAST_RESULT["current_parsed"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = LAST_RESULT["filename"]
    LAST_RESULT = None

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
