from flask import Flask, render_template, request, redirect
import pandas as pd
import os

app = Flask(__name__)

FILE = "data.xlsx"

# اگر فایل اکسل وجود ندارد بساز
if not os.path.exists(FILE):
    df = pd.DataFrame(columns=["id", "name", "phone"])
    df.to_excel(FILE, index=False)

@app.route("/")
def index():
    df = pd.read_excel(FILE)
    data = df.to_dict(orient="records")
    return render_template("index.html", data=data)

@app.route("/add", methods=["POST"])
def add():
    name = request.form["name"]
    phone = request.form["phone"]

    df = pd.read_excel(FILE)

    new_id = 1 if df.empty else df["id"].max() + 1

    new_row = pd.DataFrame([[new_id, name, phone]], columns=["id", "name", "phone"])
    df = pd.concat([df, new_row], ignore_index=True)

    df.to_excel(FILE, index=False)

    return redirect("/")

if __name__ == "__main__":
    app.run()
