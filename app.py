from flask import Flask, render_template, request, redirect
import pandas as pd
import os
from datetime import date
import jdatetime

app = Flask(__name__)

FILE = "data.xlsx"

CLINIC_NAME = "درمانکده سنتی بانوان"
DOCTOR_NAME = "هاله رستمی"


# =========================
# ساخت فایل اکسل
# =========================
def init_file():
    if not os.path.exists(FILE):
        df = pd.DataFrame(columns=[
            "id","name","age","phone",
            "disease","visit_date",
            "visit_note","fee"
        ])
        df.to_excel(FILE, index=False)

init_file()


def read():
    return pd.read_excel(FILE)


def save(df):
    temp = "temp.xlsx"
    df.to_excel(temp, index=False)
    os.replace(temp, FILE)


# =========================
# تبدیل به شمسی
# =========================
def to_jalali(g_date):
    try:
        dt = pd.to_datetime(g_date)
        return str(jdatetime.date.fromgregorian(date=dt.date()))
    except:
        return g_date


# =========================
# صفحه اصلی
# =========================
@app.route("/")
def index():
    df = read()

    if not df.empty:
        df["visit_date"] = df["visit_date"].apply(to_jalali)

    income = df["fee"].sum() if not df.empty else 0

    return render_template(
        "index.html",
        patients=df.to_dict(orient="records"),
        clinic=CLINIC_NAME,
        doctor=DOCTOR_NAME,
        income=income
    )


# =========================
# ثبت بیمار (با جلوگیری از تکرار)
# =========================
@app.route("/add", methods=["POST"])
def add():
    df = read()

    phone = request.form["phone"]

    # 🚫 جلوگیری از بیمار تکراری
    if not df.empty and phone in df["phone"].astype(str).values:
        return "<h2>❌ این بیمار قبلاً ثبت شده است (شماره تکراری)</h2><a href='/'>بازگشت</a>"

    new_id = 1 if df.empty else int(df["id"].max()) + 1

    new_row = {
        "id": new_id,
        "name": request.form["name"],
        "age": request.form["age"],
        "phone": phone,
        "disease": request.form["disease"],
        "visit_date": str(date.today()),
        "visit_note": request.form["visit_note"],
        "fee": float(request.form["fee"])
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save(df)

    return redirect("/")


# =========================
# پروفایل بیمار
# =========================
@app.route("/profile/<phone>")
def profile(phone):
    df = read()

    patient = df[df["phone"].astype(str) == str(phone)]

    if patient.empty:
        return "<h2>بیمار یافت نشد</h2><a href='/'>بازگشت</a>"

    patient = patient.iloc[0].to_dict()

    patient["visit_date"] = to_jalali(patient["visit_date"])

    return render_template("profile.html", p=patient)


# =========================
# حذف بیمار
# =========================
@app.route("/delete/<int:id>")
def delete(id):
    df = read()
    df = df[df["id"] != id]
    save(df)
    return redirect("/")


# =========================
# گزارش درآمد بازه‌ای (شمسی)
# =========================
@app.route("/report_income")
def report_income():
    start = request.args.get("start")
    end = request.args.get("end")

    df = read()

    df["visit_date"] = pd.to_datetime(df["visit_date"], errors="coerce")

    if start and end:
        start_g = jdatetime.date(*map(int, start.split("-"))).togregorian()
        end_g = jdatetime.date(*map(int, end.split("-"))).togregorian()

        df = df[
            (df["visit_date"] >= start_g) &
            (df["visit_date"] <= end_g)
        ]

    income = df["fee"].sum() if not df.empty else 0

    return f"""
    <div style="font-family:sans-serif;text-align:center;margin-top:50px">
        <h2>📊 گزارش درآمد</h2>
        <p>از {start} تا {end}</p>
        <h1>{income} تومان</h1>
        <a href='/'>بازگشت</a>
    </div>
    """


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
