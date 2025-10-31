from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
import pandas as pd
import os
import glob

app = Flask(__name__)
app.secret_key = "secret"

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

@app.route("/")
def index():
    files = [os.path.basename(f) for f in glob.glob(os.path.join(UPLOAD_FOLDER, "*.xlsx"))]
    return render_template("index.html", files=files)

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        flash("ファイルが選択されていません。")
        return redirect(url_for("index"))
    file = request.files["file"]
    if not file:
        flash("ファイルを選択してください。")
        return redirect(url_for("index"))
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(filepath)
    flash(f"{file.filename} をアップロードしました。")
    return redirect(url_for("index"))

@app.route("/delete", methods=["POST"])
def delete():
    filename = request.form.get("filename")
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    if os.path.exists(filepath):
        os.remove(filepath)
        flash(f"{filename} を削除しました。")
    else:
        flash(f"{filename} が見つかりませんでした。")
    return redirect(url_for("index"))

@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename, as_attachment=True)

@app.route("/search", methods=["POST"])
def search():
    name = request.form.get("name") or ""
    name = name.strip()
    store_query = request.form.get("store") or ""
    store_query = store_query.strip()
    month_name = request.form.get("month_name") or ""
    month_store = request.form.get("month_store") or ""
    # 検索対象ファイル（全xlsx）
    files = [os.path.join(app.config["UPLOAD_FOLDER"], f) for f in os.listdir(app.config["UPLOAD_FOLDER"]) if f.endswith(".xlsx")]
    if not files:
        flash("アップロード済みファイルがありません。")
        return redirect(url_for("index"))
    def parse_date(date_str):
        import re
        import pandas as pd
        m = re.match(r"^(\d{1,2})月(\d{1,2})日$", str(date_str).strip())
        if m:
            month, day = m.groups()
            return f"2025-{int(month):02d}-{int(day):02d}"
        try:
            val = float(date_str)
            return (pd.Timestamp('1899-12-30') + pd.Timedelta(days=val)).strftime("%Y-%m-%d")
        except Exception:
            pass
        return str(date_str)
    def month_match(date_str, month_query):
        if not month_query:
            return True
        try:
            dt = pd.to_datetime(date_str, errors="coerce")
            return dt.month == int(month_query)
        except Exception:
            return False
    results = []
    store_results = {}
    for fpath in files:
        try:
            df_raw = pd.read_excel(fpath, header=None)
            header_row = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("勤務時間", na=False)).any(axis=1)].index[0]
            df = pd.read_excel(fpath, header=header_row)
            df = df[[col for col in df.columns if not str(col).startswith("Unnamed")]]
            columns = df.columns.tolist()
        except Exception as e:
            flash(f"{os.path.basename(fpath)} の読み込みに失敗しました: {e}")
            continue
        if name and not store_query:
            for i, row in df.iterrows():
                store = str(row.iloc[0]).strip()
                time = row.iloc[1]
                for col in columns[2:]:
                    value = str(row[col]).strip()
                    date_val = parse_date(col)
                    if name in value and month_match(date_val, month_name):
                        results.append({
                            "日付": date_val,
                            "店舗": store,
                            "勤務時間": time
                        })
        elif store_query and not name:
            for i, row in df.iterrows():
                store = str(row.iloc[0]).strip()
                if store_query in store and store:
                    for col in columns[2:]:
                        value = str(row[col]).strip()
                        date_val = parse_date(col)
                        if value and value.lower() != "nan" and month_match(date_val, month_store):
                            if date_val not in store_results:
                                store_results[date_val] = []
                            store_results[date_val].append(value)
        elif name and store_query:
            for i, row in df.iterrows():
                store = str(row.iloc[0]).strip()
                time = row.iloc[1]
                for col in columns[2:]:
                    value = str(row[col]).strip()
                    date_val = parse_date(col)
                    if name in value and store_query in store and month_match(date_val, month_name):
                        results.append({
                            "日付": date_val,
                            "店舗": store,
                            "勤務時間": time
                        })
    if name and not store_query:
        if not results:
            flash(f"「{name}」さんのシフトは見つかりませんでした。")
            return redirect(url_for("index"))
        result_df = pd.DataFrame(results)
        result_df["日付_dt"] = pd.to_datetime(result_df["日付"], errors="coerce")
        result_df = result_df.sort_values(by="日付_dt")
        calendar = []
        for _, row in result_df.iterrows():
            calendar.append({
                "date": row["日付"],
                "store": row["店舗"],
                "time": row["勤務時間"],
                "work": True
            })
        while len(calendar) % 7 != 0:
            calendar.append({"date": "-", "work": False})
        return render_template("result.html", name=name, calendar=calendar, mode="name")
    elif store_query and not name:
        if not store_results:
            flash(f"「{store_query}」のシフトは見つかりませんでした。")
            return redirect(url_for("index"))
        sorted_dates = sorted(store_results.keys(), key=lambda x: pd.to_datetime(x, errors="coerce"))
        store_calendar = [{"date": d, "names": store_results[d]} for d in sorted_dates]
        return render_template("result.html", store=store_query, store_calendar=store_calendar, mode="store")
    elif name and store_query:
        if not results:
            flash(f"「{name}」さんの「{store_query}」でのシフトは見つかりませんでした。")
            return redirect(url_for("index"))
        result_df = pd.DataFrame(results)
        result_df["日付_dt"] = pd.to_datetime(result_df["日付"], errors="coerce")
        result_df = result_df.sort_values(by="日付_dt")
        calendar = []
        for _, row in result_df.iterrows():
            calendar.append({
                "date": row["日付"],
                "store": row["店舗"],
                "time": row["勤務時間"],
                "work": True
            })
        while len(calendar) % 7 != 0:
            calendar.append({"date": "-", "work": False})
        return render_template("result.html", name=name, store=store_query, calendar=calendar, mode="both")
    else:
        flash("検索条件を入力してください。")
        return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)