from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash, session
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from collections import defaultdict
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key'

UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
SAMPLE_FOLDER = 'static/samples'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
os.makedirs(SAMPLE_FOLDER, exist_ok=True)


def create_sample_files():
    df = pd.DataFrame({
        'Roll. No': ['1001', '1002'],
        'Student Name': ['Alice', 'Bob'],
        '12-07': ['P', 'X'],
        '13-07': ['X', 'P']
    })
    df.to_excel(os.path.join(SAMPLE_FOLDER, 'sample_lecture.xlsx'), index=False)
    df.to_excel(os.path.join(SAMPLE_FOLDER, 'sample_lab.xlsx'), index=False)


def sort_key(date):
    if '.' in date:
        base, suffix = date.split('.')
        suffix = int(suffix[0])
    else:
        base = date
        suffix = 0
    day, month = map(int, base[:5].split('-'))
    return (month, day, suffix)


def setnumber(df, output_path):
    id_cols = df.columns[:2]
    data_cols = df.columns[2:]

    def renumber_row(row):
        counter = 1
        for col in data_cols:
            if str(row[col]).strip().upper() != 'X':
                row[col] = counter
                counter += 1
            else:
                row[col] = 'X'
        return row

    df.loc[1:] = df.loc[1:].apply(renumber_row, axis=1)
    df.to_excel(output_path, index=False)


def setX(excel_path, styled_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    header = [cell.value for cell in ws[1]]
    lab_columns = [i for i, col in enumerate(header) if isinstance(col, str) and col.endswith('_lab')]
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for col_idx in lab_columns:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            row[col_idx].fill = yellow_fill
    for row in ws.iter_rows(min_row=2):
        for cell in row[2:]:
            if cell.value == 'X':
                cell.font = Font(bold=True, size=14, color="FF0000")
    wb.save(styled_path)


@app.route("/", methods=["GET", "POST"])
def index():
    #create_sample_files()
    download_file = session.pop("download_file", None)

    if request.method == "POST":
        try:
            lecture_file = request.files['lecture']
            lab_file = request.files['lab']
            if not lecture_file or not lab_file:
                flash("Both files are required.", "error")
                return redirect(url_for("index"))

            lecture_path = os.path.join(UPLOAD_FOLDER, "lecture.xlsx")
            lab_path = os.path.join(UPLOAD_FOLDER, "lab.xlsx")
            lecture_file.save(lecture_path)
            lab_file.save(lab_path)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"styled_final_{timestamp}.xlsx"
            final_path = os.path.join(DOWNLOAD_FOLDER, f"final_{timestamp}.xlsx")
            styled_path = os.path.join(DOWNLOAD_FOLDER, file_name)

            key_columns = ["Roll. No", "Student Name"]
            df1 = pd.read_excel(lecture_path)
            df2 = pd.read_excel(lab_path)
            df2_renamed = df2.rename(columns={col: col + '_lab' for col in df2.columns if col not in key_columns})
            merged_df = pd.merge(df1, df2_renamed, on=key_columns, how="inner")
            merged_df.columns = merged_df.columns.str.strip()
            date_cols = [col for col in merged_df.columns if col not in key_columns]

            grouped = defaultdict(list)
            for col in date_cols:
                base = col.split("_")[0]
                try:
                    pd.to_datetime(base, format="%d-%m")
                    grouped[base].append(col)
                except ValueError:
                    pass

            sort_cols = sorted(date_cols, key=sort_key)
            final_df = merged_df.loc[:, key_columns + sort_cols]

            setnumber(final_df, final_path)
            setX(final_path, styled_path)

            session["download_file"] = file_name
            flash("File processed successfully!", "success")
            return redirect(url_for("index"))

        except Exception as e:
            flash(f"Error: {str(e)}", "error")
            return redirect(url_for("index"))

    return render_template("index.html", download_file=download_file)


@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(DOWNLOAD_FOLDER, filename, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True,host="0.0.0.0")
