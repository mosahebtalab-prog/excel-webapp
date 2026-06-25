from flask import Flask, render_template, request, redirect, jsonify
import pandas as pd
import os

app = Flask(__name__)

FILE = "data.xlsx"

# ساخت فایل اگر وجود نداشت
if not os.path.exists(FILE):
    df = pd.DataFrame(columns=[
        "id","name","age","phone","disease",
        "visit_date","doctor","fee"
    ])
    df.to_excel(FILE, index=False)


@app.route("/")
def index():
    df = pd.read_excel(FILE)
    return render_template("index.html", patients=df.to_dict(orient="records"))


# ➕ ثبت بیمار
@app.route("/add", methods=["POST"])
def add():
    name = request.form["name"]
    age = request.form["age"]
    phone = request.form["phone"]
    disease = request.form["disease"]
    visit_date = request.form["visit_date"]
    doctor = request.form["doctor"]
    fee = float(request.form["fee"])

    df = pd.read_excel(FILE)

    new_id = 1 if df.empty else int(df["id"].max()) + 1

    new_row = pd.DataFrame([{
        "id": new_id,
        "name": name,
        "age": age,
        "phone": phone,
        "disease": disease,
        "visit_date": visit_date,
        "doctor": doctor,
        "fee": fee
    }])

    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(FILE, index=False)

    return redirect("/")


# 🗑 حذف بیمار
@app.route("/delete/<int:id>")
def delete(id):
    df = pd.read_excel(FILE)
    df = df[df["id"] != id]
    df.to_excel(FILE, index=False)
    return redirect("/")


# 📊 گزارش مالی
@app.route("/report")
def report():
    df = pd.read_excel(FILE)

    total_income = float(df["fee"].sum()) if not df.empty else 0
    total_patients = len(df)

    today = df[df["visit_date"] == pd.Timestamp.today().strftime("%Y-%m-%d")]
    today_income = float(today["fee"].sum()) if not today.empty else 0

    return jsonify({
        "total_income": total_income,
        "today_income": today_income,
        "total_patients": total_patients
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
