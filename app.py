#!/usr/bin/env python3

"""
勤務表自動生成 Webアプリ
Author: Taiki Watanabe
"""

import os
import io
import random
from pathlib import Path
from functools import wraps

import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for
from openpyxl import load_workbook
from openpyxl.styles import Alignment

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

# パスワード設定（環境変数から取得、デフォルトは 'tnc2026'）
APP_PASSWORD = os.environ.get('APP_PASSWORD', 'tnc2026')


def login_required(f):
    """ログイン必須デコレータ"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('authenticated'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


# ===== 勤務表生成ロジック（generate_shift.pyから移植） =====

def is_weekend(day_of_week: str) -> bool:
    """土日かどうか判定"""
    return day_of_week in ['土', '日']


def is_off(value: str) -> bool:
    """休みかどうか判定（公，年，または空でない休み表記）"""
    if pd.isna(value) or value == '':
        return False
    return value in ['公', '年']


def get_off_value(value: str) -> str:
    """休みの値を取得（公 or 年）"""
    if pd.isna(value) or value == '':
        return ''
    if value in ['公', '年']:
        return value
    return ''


def select_y_candidate(candidates: list, y_counts: dict, prev_day_off: dict) -> str:
    """呼び出しを割り当てる候補を選ぶ"""
    if not candidates:
        raise ValueError("候補者がいません")
    
    if len(candidates) == 1:
        return candidates[0]
    
    not_after_off = [c for c in candidates if not prev_day_off.get(c, False)]
    
    if not_after_off:
        target_list = not_after_off
    else:
        target_list = candidates
    
    min_y = min(y_counts[c] for c in target_list)
    min_y_candidates = [c for c in target_list if y_counts[c] == min_y]
    
    return min_y_candidates[0]


def assign_shifts(df: pd.DataFrame, staff_columns: list) -> pd.DataFrame:
    """勤務表を自動生成する"""
    result_df = df.copy()
    y_counts = {staff: 0 for staff in staff_columns}
    prev_day_off = {staff: False for staff in staff_columns}
    
    for idx, row in df.iterrows():
        day_of_week = row['曜日']
        weekend = is_weekend(day_of_week)
        
        staff_off = {}
        staff_working = []
        
        for staff in staff_columns:
            off_value = get_off_value(row[staff])
            if off_value:
                staff_off[staff] = off_value
            else:
                staff_working.append(staff)
        
        for staff, off_value in staff_off.items():
            result_df.at[idx, staff] = off_value
        
        if not staff_working:
            for staff in staff_columns:
                prev_day_off[staff] = is_off(row[staff])
            continue
        
        if len(staff_working) == 1:
            only_staff = staff_working[0]
            result_df.at[idx, only_staff] = 'Y9'
            y_counts[only_staff] += 1
        elif weekend:
            y_candidate = select_y_candidate(staff_working, y_counts, prev_day_off)
            for staff in staff_working:
                if staff == y_candidate:
                    result_df.at[idx, staff] = 'Y9'
                    y_counts[staff] += 1
                else:
                    result_df.at[idx, staff] = '9'
        else:
            nine_am_staff = random.choice(staff_working)
            result_df.at[idx, nine_am_staff] = '9'
            
            ten_am_staff = [s for s in staff_working if s != nine_am_staff]
            y_candidate = select_y_candidate(ten_am_staff, y_counts, prev_day_off)
            
            for staff in ten_am_staff:
                if staff == y_candidate:
                    result_df.at[idx, staff] = 'Y10'
                    y_counts[staff] += 1
                else:
                    result_df.at[idx, staff] = '10'
        
        for staff in staff_columns:
            prev_day_off[staff] = is_off(row[staff])
    
    return result_df


def get_y_counts(df: pd.DataFrame, staff_columns: list) -> dict:
    """各スタッフの呼び出し回数を取得"""
    return {
        staff: df[staff].str.contains('Y', na=False).sum()
        for staff in staff_columns
    }


def is_balanced(y_counts: dict) -> bool:
    """呼び出し回数が±1以内かチェック"""
    counts = list(y_counts.values())
    return max(counts) - min(counts) <= 1


def parse_shift_value(value: str) -> tuple:
    """シフト値を（Y有無，時間/休み）に分解する"""
    if pd.isna(value) or value == '':
        return ('', '')
    
    value = str(value)
    
    if value.startswith('Y'):
        return ('Y', value[1:])
    elif value in ['公', '年']:
        return ('', value)
    else:
        return ('', value)


def convert_to_excel_format(df: pd.DataFrame, staff_columns: list) -> pd.DataFrame:
    """結果DataFrameをExcel出力用の形式に変換する"""
    excel_data = []
    
    header = ['', '']
    for staff in staff_columns:
        short_name = staff.replace('スタッフ', '')
        header.extend([short_name, ''])
    
    excel_data.append(header)
    
    for _, row in df.iterrows():
        excel_row = [row['日付'], row['曜日']]
        for staff in staff_columns:
            y_mark, time_value = parse_shift_value(row[staff])
            excel_row.extend([y_mark, time_value])
        excel_data.append(excel_row)
    
    return pd.DataFrame(excel_data)


def generate_shift_schedule(csv_content: str) -> tuple:
    """
    CSVコンテンツから勤務表を生成
    Returns: (excel_bytes, y_counts, attempts)
    """
    df = pd.read_csv(io.StringIO(csv_content), dtype=str)
    staff_columns = [col for col in df.columns if col not in ['日付', '曜日']]
    
    max_attempts = 1000
    best_result = None
    best_diff = float('inf')
    best_y_counts = {}
    attempts = 0
    
    for attempt in range(max_attempts):
        attempts = attempt + 1
        result_df = assign_shifts(df, staff_columns)
        y_counts = get_y_counts(result_df, staff_columns)
        counts = list(y_counts.values())
        diff = max(counts) - min(counts)
        
        if diff < best_diff:
            best_diff = diff
            best_result = result_df
            best_y_counts = y_counts
        
        if is_balanced(y_counts):
            break
    
    # Excel形式に変換
    excel_df = convert_to_excel_format(best_result, staff_columns)
    
    # Excelファイルをメモリに書き込み
    output = io.BytesIO()
    excel_df.to_excel(output, index=False, header=False, engine='openpyxl')
    output.seek(0)
    
    # セル結合処理
    wb = load_workbook(output)
    ws = wb.active
    
    for i, staff in enumerate(staff_columns):
        start_col = 3 + (i * 2)
        end_col = start_col + 1
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        ws.cell(row=1, column=start_col).alignment = Alignment(horizontal='center')
    
    # 最終的なExcelをメモリに保存
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    
    return final_output, best_y_counts, attempts


# ===== Flaskルート =====

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'データが送信されていません'}), 400
        
        year = data.get('year')
        month = data.get('month')
        staff_names = data.get('staffNames', ['A', 'B', 'C'])
        schedule = data.get('schedule', [])
        
        if not schedule:
            return jsonify({'error': 'スケジュールデータがありません'}), 400
        
        # DataFrameを構築
        rows = []
        for day_data in schedule:
            row = {
                '日付': day_data['day'],
                '曜日': day_data['dayOfWeek']
            }
            for staff_name in staff_names:
                row[staff_name] = day_data['staff'].get(staff_name, '')
            rows.append(row)
        
        df = pd.DataFrame(rows)
        
        # 勤務表生成
        staff_columns = staff_names
        
        max_attempts = 1000
        best_result = None
        best_diff = float('inf')
        best_y_counts = {}
        attempts = 0
        
        for attempt in range(max_attempts):
            attempts = attempt + 1
            result_df = assign_shifts(df.copy(), staff_columns)
            y_counts = get_y_counts(result_df, staff_columns)
            counts = list(y_counts.values())
            diff = max(counts) - min(counts)
            
            if diff < best_diff:
                best_diff = diff
                best_result = result_df.copy()
                best_y_counts = y_counts
            
            if is_balanced(y_counts):
                break
        
        # Excel形式に変換
        excel_df = convert_to_excel_format(best_result, staff_columns)
        
        # Excelファイルをメモリに書き込み
        output = io.BytesIO()
        excel_df.to_excel(output, index=False, header=False, engine='openpyxl')
        output.seek(0)
        
        # セル結合処理
        wb = load_workbook(output)
        ws = wb.active
        
        for i, staff in enumerate(staff_columns):
            start_col = 3 + (i * 2)
            end_col = start_col + 1
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
            ws.cell(row=1, column=start_col).alignment = Alignment(horizontal='center')
        
        # 最終的なExcelをメモリに保存
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        # セッションに保存
        app.config['LAST_EXCEL'] = final_output
        
        return jsonify({
            'success': True,
            'attempts': attempts,
            'y_counts': {k: int(v) for k, v in best_y_counts.items()}
        })
    except Exception as e:
        return jsonify({'error': f'生成エラー: {str(e)}'}), 500


@app.route('/download')
def download():
    if 'LAST_EXCEL' not in app.config or app.config['LAST_EXCEL'] is None:
        return jsonify({'error': 'まず勤務表を生成してください'}), 400
    
    excel_bytes = app.config['LAST_EXCEL']
    excel_bytes.seek(0)
    
    return send_file(
        excel_bytes,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='shift_schedule.xlsx'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('FLASK_DEBUG', 'true').lower() == 'true'
    app.run(debug=debug, host='0.0.0.0', port=port)
