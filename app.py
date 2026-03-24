from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, redirect, url_for, flash, session
import pandas as pd
import os
from werkzeug.utils import secure_filename
import uuid
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from collections import defaultdict

app = Flask(__name__)
app.secret_key = 'your-secret-key-here-change-in-production'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'
app.config['SAMPLE_FOLDER'] = 'static/samples'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Create necessary directories
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['SAMPLE_FOLDER'], exist_ok=True)


# ==================== LANDING PAGE ====================
@app.route('/')
def landing():
    return render_template('landing.html')


def clean_dataframe(df):
    """Replace NaN values with 'A' (Absent) or empty string"""
    df = df.fillna('A')
    return df


# ==================== PREPROCESS RAW ATTENDANCE FILE ====================
def preprocess_attendance_file(input_path, output_path):
    """
    Cleans the raw attendance Excel file:
    - Removes rows 1-7 (metadata/title rows)
    - Removes column 1 (#), Total column, and %age column
    - Creates a proper 2-row header:
        Row 1: Roll. No | Student Name | Dates (e.g. 21-01)
        Row 2: (merged)  (merged)      | Period numbers (7 or 8)
    - Roll. No and Student Name are merged vertically across both header rows
    """
    wb_src = load_workbook(input_path)
    ws_src = wb_src.active

    # Detect which source columns to keep by scanning header row (row 9)
    # Skip col 1 (#), and any col whose header is 'Total' or '%age'
    keep_cols = []
    for col_idx in range(1, ws_src.max_column + 1):
        header_val = ws_src.cell(9, col_idx).value
        if header_val in ('#', 'Total', '%age', None) and col_idx == 1:
            continue  # skip # column
        if str(header_val) in ('Total', '%age'):
            continue  # skip Total and %age columns
        keep_cols.append(col_idx)

    wb_new = Workbook()
    ws_new = wb_new.active

    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # --- ROW 1: Roll.No | Student Name | Dates from source row 8 ---
    for c_idx, src_col in enumerate(keep_cols, start=1):
        cell = ws_new.cell(1, c_idx)
        src_val = ws_src.cell(9, src_col).value  # Roll. No / Student Name / period
        if src_val == 'Roll. No':
            cell.value = 'Roll. No'
        elif src_val == 'Student Name':
            cell.value = 'Student Name'
        else:
            cell.value = ws_src.cell(8, src_col).value  # date e.g. '21-01'
        cell.font = Font(bold=True, name='Arial', size=10)
        cell.alignment = Alignment(horizontal='center', vertical='bottom')
        cell.border = border

    # --- ROW 2: blank for Roll.No/Student Name | Period numbers from source row 9 ---
    for c_idx, src_col in enumerate(keep_cols, start=1):
        cell = ws_new.cell(2, c_idx)
        src_val = ws_src.cell(9, src_col).value
        if src_val not in ('Roll. No', 'Student Name'):
            cell.value = src_val  # period number: 7 or 8
        cell.font = Font(bold=True, name='Arial', size=10)
        cell.alignment = Alignment(horizontal='center', vertical='top')
        cell.border = border

    # --- Merge Roll.No and Student Name across rows 1-2 ---
    ws_new.merge_cells('A1:A2')
    ws_new.merge_cells('B1:B2')
    ws_new.cell(1, 1).alignment = Alignment(horizontal='center', vertical='center')
    ws_new.cell(1, 2).alignment = Alignment(horizontal='center', vertical='center')

    ws_new.row_dimensions[1].height = 15
    ws_new.row_dimensions[2].height = 15

    # --- DATA ROWS: source rows 10 onwards → destination rows 3 onwards ---
    for src_row in range(10, ws_src.max_row + 1):
        dest_row = src_row - 7  # row 10 → row 3
        for dest_col, src_col in enumerate(keep_cols, start=1):
            dest_cell = ws_new.cell(dest_row, dest_col)
            dest_cell.value = ws_src.cell(src_row, src_col).value
            dest_cell.font = Font(name='Arial', size=10)
            dest_cell.alignment = Alignment(horizontal='center', vertical='center')
            dest_cell.border = border

    # --- Column widths ---
    ws_new.column_dimensions['A'].width = 14
    ws_new.column_dimensions['B'].width = 18
    for i in range(3, len(keep_cols) + 1):
        ws_new.column_dimensions[get_column_letter(i)].width = 7

    wb_new.save(output_path)


# ==================== ATTENDANCE MERGER ====================
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

    # --- Insert one row after header for Lecture row ---
    ws.insert_rows(2, amount=1)
    total_cols = ws.max_column

    # Row 2: merged A2:B2 = "Lecture", counting from C
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=2)
    ws.cell(row=2, column=1).value = "Lecture"
    ws.cell(row=2, column=1).font = Font(size=14, bold=True)
    ws.cell(row=2, column=1).alignment = Alignment(horizontal='center', vertical='center')
    for col_idx in range(3, total_cols + 1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = col_idx - 2
        cell.font = Font(size=14)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Row 3: merge A3:B3, write "Period Number"
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
    light_green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    ws.cell(row=3, column=1).value = "Period Number"
    ws.cell(row=3, column=1).font = Font(size=14, bold=True)
    ws.cell(row=3, column=1).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row=3, column=1).fill = light_green_fill
    for col_idx in range(3, total_cols + 1):
        ws.cell(row=3, column=col_idx).fill = light_green_fill

    # --- Style lab columns with yellow fill ---
    header = [cell.value for cell in ws[1]]
    lab_columns = [i for i, col in enumerate(header) if isinstance(col, str) and col.endswith('_lab')]
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for col_idx in lab_columns:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            row[col_idx].fill = yellow_fill

    # --- Style X cells: Red background, bold 14pt white font, center aligned ---
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    for row in ws.iter_rows(min_row=4):
        for cell in row[2:]:
            if cell.value == 'X':
                cell.font = Font(bold=True, size=14, color="FFFFFF")
                cell.fill = red_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # --- Border + center alignment ---
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border
            if isinstance(cell.value, (int, float)) or (cell.row == 2 and cell.column >= 3):
                cell.alignment = center

    wb.save(styled_path)


@app.route('/attendance')
def attendance():
    download_file = session.pop("download_file", None)
    return render_template('attendance.html', download_file=download_file)


@app.route('/attendance/upload', methods=['POST'])
def attendance_upload():
    try:
        lecture_file = request.files['lecture']
        lab_file = request.files['lab']
        if not lecture_file or not lab_file:
            flash("Both files are required.", "error")
            return redirect(url_for("attendance"))

        # Save raw uploaded files
        lecture_raw_path = os.path.join(app.config['UPLOAD_FOLDER'], "lecture_raw.xlsx")
        lab_raw_path = os.path.join(app.config['UPLOAD_FOLDER'], "lab_raw.xlsx")
        lecture_file.save(lecture_raw_path)
        lab_file.save(lab_raw_path)

        # Cleaned paths after preprocessing
        lecture_path = os.path.join(app.config['UPLOAD_FOLDER'], "lecture.xlsx")
        lab_path = os.path.join(app.config['UPLOAD_FOLDER'], "lab.xlsx")

        # ---- PREPROCESS: Remove extra rows/columns from raw files ----
        preprocess_attendance_file(lecture_raw_path, lecture_path)
        preprocess_attendance_file(lab_raw_path, lab_path)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"styled_final_{timestamp}.xlsx"
        final_path = os.path.join(app.config['DOWNLOAD_FOLDER'], f"final_{timestamp}.xlsx")
        styled_path = os.path.join(app.config['DOWNLOAD_FOLDER'], file_name)

        key_columns = ["Roll. No", "Student Name"]

        # Read cleaned Excel files
        # Skip the 2-row header (merged), read from row 3 onward
        df1 = pd.read_excel(lecture_path, header=[0, 1])
        df2 = pd.read_excel(lab_path, header=[0, 1])

        # Flatten multi-level header: combine date + period e.g. ('21-01', '7') → '21-01_7'
        def flatten_header(df):
            new_cols = []
            for col in df.columns:
                top, bot = col
                top = str(top).strip()
                bot = str(bot).strip()
                if top in ('Roll. No', 'Student Name'):
                    new_cols.append(top)
                elif bot in ('nan', ''):
                    new_cols.append(top)
                else:
                    new_cols.append(f"{top}_{bot}")
            df.columns = new_cols
            return df

        df1 = flatten_header(df1)
        df2 = flatten_header(df2)

        # Clean NaN → 'A'
        df1 = df1.fillna('A')
        df2 = df2.fillna('A')

        df2_renamed = df2.rename(columns={col: col + '_lab' for col in df2.columns if col not in key_columns})
        merged_df = pd.merge(df1, df2_renamed, on=key_columns, how="inner")
        merged_df.columns = merged_df.columns.str.strip()
        date_cols = [col for col in merged_df.columns if col not in key_columns]

        sort_cols = sorted(date_cols, key=sort_key)
        final_df = merged_df.loc[:, key_columns + sort_cols]

        setnumber(final_df, final_path)
        setX(final_path, styled_path)

        session["download_file"] = file_name
        flash("File processed successfully! Thanks to Mohit", "success")
        return redirect(url_for("attendance"))

    except Exception as e:
        flash(f"Error: {str(e)}", "error")
        return redirect(url_for("attendance"))


@app.route('/attendance/download/<filename>')
def attendance_download(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)
# ==================== EXCEL JOINER ====================
@app.route('/joiner')
def joiner():
    return render_template('joiner.html')

@app.route('/joiner/upload', methods=['POST'])
def joiner_upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    file_type = request.form.get('file_type')
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    try:
        file_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_id}_{filename}")
        file.save(filepath)
        
        if filename.endswith('.csv'):
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)
        
        if 'files' not in session:
            session['files'] = {}
        
        session['files'][file_type] = {
            'filepath': filepath,
            'filename': filename,
            'file_id': file_id
        }
        session.modified = True
        
        # Convert NaN to None for proper JSON serialization
        preview_df = df.head(3).where(pd.notna(df.head(3)), None)
        
        return jsonify({
            'success': True,
            'filename': filename,
            'columns': list(df.columns),
            'rows': len(df),
            'preview': preview_df.to_dict('records')
        })
    
    except Exception as e:
        return jsonify({'error': f'Failed to read file: {str(e)}'}), 400

@app.route('/joiner/get_columns', methods=['GET'])
def joiner_get_columns():
    if 'files' not in session or 'left' not in session['files'] or 'right' not in session['files']:
        return jsonify({'error': 'Please upload both files first'}), 400
    
    try:
        left_file = session['files']['left']['filepath']
        right_file = session['files']['right']['filepath']
        
        if left_file.endswith('.csv'):
            df_left = pd.read_csv(left_file)
        else:
            df_left = pd.read_excel(left_file)
        
        if right_file.endswith('.csv'):
            df_right = pd.read_csv(right_file)
        else:
            df_right = pd.read_excel(right_file)
        
        common_cols = list(set(df_left.columns) & set(df_right.columns))
        
        return jsonify({
            'success': True,
            'left_columns': list(df_left.columns),
            'right_columns': list(df_right.columns),
            'common_columns': common_cols
        })
    
    except Exception as e:
        return jsonify({'error': f'Failed to read columns: {str(e)}'}), 400

@app.route('/joiner/join', methods=['POST'])
def joiner_join():
    if 'files' not in session or 'left' not in session['files'] or 'right' not in session['files']:
        return jsonify({'error': 'Please upload both files first'}), 400
    
    try:
        data = request.json
        left_columns = data.get('left_columns', [])
        right_columns = data.get('right_columns', [])
        join_type = data.get('join_type', 'inner')
        
        if not left_columns or not right_columns:
            return jsonify({'error': 'Please select columns to join on'}), 400
        
        if len(left_columns) != len(right_columns):
            return jsonify({'error': 'Number of selected columns must match'}), 400
        
        left_file = session['files']['left']['filepath']
        right_file = session['files']['right']['filepath']
        
        if left_file.endswith('.csv'):
            df_left = pd.read_csv(left_file)
        else:
            df_left = pd.read_excel(left_file)
        
        if right_file.endswith('.csv'):
            df_right = pd.read_csv(right_file)
        else:
            df_right = pd.read_excel(right_file)
        
        # Handle the case where join columns have different names
        if left_columns == right_columns:
            df_joined = pd.merge(
                df_left, 
                df_right, 
                on=left_columns,
                how=join_type,
                suffixes=('_left', '_right')
            )
        else:
            df_joined = pd.merge(
                df_left, 
                df_right, 
                left_on=left_columns, 
                right_on=right_columns, 
                how=join_type,
                suffixes=('_left', '_right')
            )
            
            cols_to_drop = [col for col in right_columns if col in df_joined.columns and col not in left_columns]
            if cols_to_drop:
                df_joined = df_joined.drop(columns=cols_to_drop)
        
        if df_joined.empty:
            return jsonify({'warning': 'Join returned no rows', 'columns': [], 'rows': 0})
        
        joined_id = str(uuid.uuid4())
        joined_filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{joined_id}_joined.xlsx")
        df_joined.to_excel(joined_filepath, index=False)
        
        session['joined_file'] = {
            'filepath': joined_filepath,
            'file_id': joined_id
        }
        session.modified = True
        
        # Convert NaN to None for proper JSON serialization
        preview_df = df_joined.head(5).where(pd.notna(df_joined.head(5)), None)
        
        return jsonify({
            'success': True,
            'columns': list(df_joined.columns),
            'rows': len(df_joined),
            'preview': preview_df.to_dict('records')
        })
    
    except Exception as e:
        return jsonify({'error': f'Failed to join files: {str(e)}'}), 400

@app.route('/joiner/download', methods=['POST'])
def joiner_download():
    if 'joined_file' not in session:
        return jsonify({'error': 'No joined file available'}), 400
    
    try:
        data = request.json
        selected_columns = data.get('selected_columns', [])
        
        if not selected_columns:
            return jsonify({'error': 'Please select at least one column'}), 400
        
        joined_filepath = session['joined_file']['filepath']
        df_joined = pd.read_excel(joined_filepath)
        df_final = df_joined[selected_columns]
        
        final_id = str(uuid.uuid4())
        final_filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{final_id}_final.xlsx")
        df_final.to_excel(final_filepath, index=False)
        
        return send_file(
            final_filepath,
            as_attachment=True,
            download_name='joined_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        return jsonify({'error': f'Failed to create download: {str(e)}'}), 400

@app.route('/joiner/reset', methods=['POST'])
def joiner_reset():
    if 'files' in session:
        for file_type in session['files']:
            filepath = session['files'][file_type].get('filepath')
            if filepath and os.path.exists(filepath):
                os.remove(filepath)
    
    if 'joined_file' in session:
        filepath = session['joined_file'].get('filepath')
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
    
    session.clear()
    return jsonify({'success': True})

if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0")