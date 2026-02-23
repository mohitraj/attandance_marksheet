from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, redirect, url_for, flash, session
import pandas as pd
import os
from werkzeug.utils import secure_filename
import uuid
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
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

    # --- Insert two rows after header ---
    ws.insert_rows(2, amount=2)
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

    # Row 3: merged A3:B3 = "Period Number", clear C onwards
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
    ws.cell(row=3, column=1).value = "Period Number"
    ws.cell(row=3, column=1).font = Font(size=14, bold=True)
    ws.cell(row=3, column=1).alignment = Alignment(horizontal='center', vertical='center')
    for col_idx in range(3, total_cols + 1):
        ws.cell(row=3, column=col_idx).value = None

    # --- Style lab columns with yellow fill ---
    header = [cell.value for cell in ws[1]]
    lab_columns = [i for i, col in enumerate(header) if isinstance(col, str) and col.endswith('_lab')]
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for col_idx in lab_columns:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            row[col_idx].fill = yellow_fill

    # --- Style X cells: Red background, bold 14pt white font, center aligned ---
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    for row in ws.iter_rows(min_row=4):  # data starts at row 4 now
        for cell in row[2:]:
            if cell.value == 'X':
                cell.font = Font(bold=True, size=14, color="FFFFFF")
                cell.fill = red_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # --- Border for entire sheet + center alignment for numbers ---
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border
            # Center align numbers and the counting row
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

        lecture_path = os.path.join(app.config['UPLOAD_FOLDER'], "lecture.xlsx")
        lab_path = os.path.join(app.config['UPLOAD_FOLDER'], "lab.xlsx")
        lecture_file.save(lecture_path)
        lab_file.save(lab_path)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"styled_final_{timestamp}.xlsx"
        final_path = os.path.join(app.config['DOWNLOAD_FOLDER'], f"final_{timestamp}.xlsx")
        styled_path = os.path.join(app.config['DOWNLOAD_FOLDER'], file_name)

        key_columns = ["Roll. No", "Student Name"]

        # Read Excel files
        df1 = pd.read_excel(lecture_path)
        df2 = pd.read_excel(lab_path)

        # Clean NaN values - fill with 'A' for absent
        df1 = df1.fillna('A')
        df2 = df2.fillna('A')

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