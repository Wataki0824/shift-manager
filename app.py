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
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

# パスワード設定（環境変数から取得，デフォルトは 'kitakyushu'）
APP_PASSWORD = os.environ.get('APP_PASSWORD', 'kitakyushu')


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





def fill_missing_public_holidays(df: pd.DataFrame, staff_columns: list, target_holidays: dict) -> pd.DataFrame:
    """不足している公休を自動で埋める（全体最適と連勤考慮）"""
    result_df = df.copy()
    
    # 各スタッフの不足公休数を計算（管理用）
    missing_counts = {}
    for staff in staff_columns:
        current_holidays = result_df[staff].apply(lambda x: 1 if x == '公' else 0).sum()
        target = target_holidays.get(staff, 0)
        missing_counts[staff] = max(0, target - current_holidays)

    # ==========================================
    # Phase 0: 土日出勤の調整（全体最適）
    # 土日は必ず1人出勤になるように、余剰人員を休ませる。
    # 「土日休みが少ない人」を優先的に休ませる。
    # ==========================================
    
    # 土日のインデックスを取得
    weekend_indices = []
    for idx, row in result_df.iterrows():
        if is_weekend(row['曜日']):
            weekend_indices.append(idx)
    
    # 完全に解消されるか、公休ストックが尽きるまでループ
    # （無限ループ防止のため回数制限）
    for _ in range(100):
        adjusted_any = False
        
        # 土日をシャッフルして順不同にチェック
        random.shuffle(weekend_indices)
        
        for idx in weekend_indices:
            # この日の出勤者を確認
            working_staff = []
            for s in staff_columns:
                if not is_off(result_df.at[idx, s]):
                    working_staff.append(s)
            
            # 2人以上なら調整が必要
            if len(working_staff) > 1:
                # 候補者の中で、実際に公休ストックが残っている人を抽出
                candidates = [s for s in working_staff if missing_counts[s] > 0]
                
                if not candidates:
                    continue # 誰もこれ以上休めない
                
                # --- 優先度判定 ---
                # 現在の「土日休み回数」が少ない人を優先的に休ませたい
                # （＝土日出勤が多い人を休ませたい）
                
                weekend_off_counts = {}
                for s in candidates:
                    # その人の現在の土日休み数
                    count = 0
                    for _idx, _row in result_df.iterrows():
                        if is_weekend(_row['曜日']) and is_off(_row[s]):
                            count += 1
                    weekend_off_counts[s] = count
                
                # 土日休みが一番少ない人を探す
                min_off = min(weekend_off_counts.values())
                target_candidates = [s for s in candidates if weekend_off_counts[s] == min_off]
                
                # 同率ならランダムに選ぶ
                target_staff = random.choice(target_candidates)
                
                # 実行
                result_df.at[idx, target_staff] = '公'
                missing_counts[target_staff] -= 1
                adjusted_any = True
        
        if not adjusted_any:
            break

    # ==========================================
    # Phase 1 & 2: 連勤解消 & 残り埋め（スタッフごと）
    # ==========================================
    
    for staff in staff_columns:
        missing = missing_counts[staff]
        if missing <= 0:
            continue
            
        added_count = 0
        
        # --- 戦略1: 連勤を解消する箇所を優先的に埋める ---
        limits_to_check = [6, 5] 
        
        for limit in limits_to_check:
            if added_count >= missing:
                break
                
            while added_count < missing:
                # 連勤検知
                consecutive_indices = []
                current_streak = []
                for idx, row in result_df.iterrows():
                    if not is_off(row[staff]):
                        current_streak.append(idx)
                    else:
                        if len(current_streak) > limit:
                            consecutive_indices.append(current_streak.copy())
                        current_streak = []
                if len(current_streak) > limit:
                     consecutive_indices.append(current_streak.copy())
                
                if not consecutive_indices:
                    break
                
                filled_something = False
                for streak in consecutive_indices:
                    if added_count >= missing:
                         break
                    candidates = streak.copy()
                    random.shuffle(candidates)
                    for idx in candidates:
                        if can_take_off(result_df, idx, staff, staff_columns):
                            result_df.at[idx, staff] = '公'
                            added_count += 1
                            filled_something = True
                            break 
                if not filled_something:
                    break

        # --- 戦略2: まだ足りなければランダムに埋める ---
        if added_count < missing:
            candidate_indices = []
            for idx, row in result_df.iterrows():
                if not is_off(row[staff]):
                    candidate_indices.append(idx)
            
            random.shuffle(candidate_indices)
            
            for idx in candidate_indices:
                if added_count >= missing:
                    break
                if can_take_off(result_df, idx, staff, staff_columns):
                    result_df.at[idx, staff] = '公'
                    added_count += 1
            
    return result_df


def can_take_off(df: pd.DataFrame, idx: int, staff: str, staff_columns: list) -> bool:
    """指定した日にそのスタッフが休んでも大丈夫かチェック"""
    day_of_week = df.at[idx, '曜日']
    is_wknd = is_weekend(day_of_week)
    
    # 他の出勤スタッフを確認
    working_staff = []
    for s in staff_columns:
        if s == staff:
            continue
        if not is_off(df.at[idx, s]):
            working_staff.append(s)
    
    working_count = len(working_staff)
    
    if is_wknd:
        # 土日は最低1人必要
        return working_count >= 1
    else:
        # 平日は最低1人必要
        return working_count >= 1



def optimize_shifts(df: pd.DataFrame, staff_columns: list) -> pd.DataFrame:
    """生成後のシフトを微調整して最適化する"""
    # 主に「休み明けのY」を回避するためのスワップを行う
    result_df = df.copy()
    
    for idx, row in result_df.iterrows():
        day_of_week = row['曜日']
        if is_weekend(day_of_week):
            continue # 土日は1人出勤が基本なのでスワップ余地なし
            
        # この日の出勤者とシフト情報を取得
        workers = []
        for staff in staff_columns:
            val = result_df.at[idx, staff]
            if not is_off(val):
                workers.append({'staff': staff, 'val': val})
        
        # 2人以上出勤していないとスワップできない
        if len(workers) < 2:
            continue
            
        # Yがついている人を探す
        y_staff_info = None
        normal_staff_info = None
        
        # 平日で「Y付きの人」と「Yなしの人」のペアを探す
        for w in workers:
            if 'Y' in w['val']:
                y_staff_info = w
            else:
                normal_staff_info = w # 複数人いる場合、最初に見つかった人を対象にする（簡易実装）
        
        if not y_staff_info or not normal_staff_info:
            continue
            
        y_staff = y_staff_info['staff']
        norm_staff = normal_staff_info['staff']
        
        # 前日の状態を確認
        if idx == 0:
            continue
            
        prev_idx = idx - 1
        y_staff_prev_off = is_off(result_df.at[prev_idx, y_staff])
        norm_staff_prev_off = is_off(result_df.at[prev_idx, norm_staff])
        
        # 「Y担当が前日休み」かつ「通常担当が前日出勤」の場合、入れ替えるべき
        if y_staff_prev_off and not norm_staff_prev_off:
            # スワップ実行
            val_y = y_staff_info['val']
            val_norm = normal_staff_info['val']
            
            result_df.at[idx, y_staff] = val_norm
            result_df.at[idx, norm_staff] = val_y

    return result_df


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
    
    # 最適化（スワップ）実行
    result_df = optimize_shifts(result_df, staff_columns)
    
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

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        password = request.form.get('password', '')
        if password == APP_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        else:
            error = 'パスワードが正しくありません'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.pop('authenticated', None)
    return redirect(url_for('login'))


@app.route('/')
@login_required
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
@login_required
def generate():
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'データが送信されていません'}), 400
        
        year = data.get('year')
        month = data.get('month')
        staff_names = data.get('staffNames', ['A', 'B', 'C'])
        schedule = data.get('schedule', [])
        target_holidays = data.get('targetHolidays', {})
        
        if not schedule:
            return jsonify({'error': 'スケジュールデータがありません'}), 400
        
        # DataFrameを構築 & 希望休マスクを作成
        rows = []
        request_mask = {} # (day_index, staff_name) -> bool
        
        for day_idx, day_data in enumerate(schedule):
            row = {
                '日付': day_data['day'],
                '曜日': day_data['dayOfWeek']
            }
            for staff_name in staff_names:
                val = day_data['staff'].get(staff_name, '')
                row[staff_name] = val
                
                # 公 または 年 が入力されている場合は希望休としてマーク
                if val in ['公', '年']:
                    request_mask[(day_idx, staff_name)] = True
                else:
                    request_mask[(day_idx, staff_name)] = False
                    
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
            
            # 公休の自動充填
            df_with_holidays = fill_missing_public_holidays(df.copy(), staff_columns, target_holidays)
            
            result_df = assign_shifts(df_with_holidays, staff_columns)
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

        # セル結合処理 & レイアウト調整
        wb = load_workbook(output)
        ws = wb.active
        
        # フォント・スタイル設定
        large_font = Font(size=14)
        bold_font = Font(size=14, bold=True)
        gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        
        # 全体の行高さを設定
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 40
            for cell in row:
                cell.font = large_font
            
        # 列幅の設定
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 6
        
        for i, staff in enumerate(staff_columns):
            start_col_idx = 3 + (i * 2)
            end_col_idx = start_col_idx + 1
            
            y_col_letter = get_column_letter(start_col_idx)     # C, E, G...
            val_col_letter = get_column_letter(end_col_idx)     # D, F, H...
            
            # Yマーク列を狭く
            ws.column_dimensions[y_col_letter].width = 5
            
            # 時間列は少し広めに
            ws.column_dimensions[val_col_letter].width = 10
            
            # ヘッダー結合
            ws.merge_cells(start_row=1, start_column=start_col_idx, end_row=1, end_column=end_col_idx)
            
            # スタイル設定
            # ヘッダー中央揃え
            header_cell = ws.cell(row=1, column=start_col_idx)
            header_cell.alignment = Alignment(horizontal='center', vertical='center')
            header_cell.font = large_font
            
            # データ領域のスタイル（中央揃え & 希望休ハイライト）
            for row_idx, row in enumerate(range(2, ws.max_row + 1)):
                # row_idx は 0始まり (DataFrameの行インデックスと一致)
                
                # Yマーク列
                cell_y = ws.cell(row=row, column=start_col_idx)
                cell_y.alignment = Alignment(horizontal='center', vertical='center')
                cell_y.font = large_font
                
                # 時間列 / 休み列（ここに公/年が入る）
                cell_val = ws.cell(row=row, column=end_col_idx)
                cell_val.alignment = Alignment(horizontal='center', vertical='center')
                cell_val.font = large_font
                
                # 希望休かどうかのチェック
                # 公や年の場合、セルに値が入っているはず
                # request_mask[(row_idx, staff)] が True ならスタイル適用
                if request_mask.get((row_idx, staff), False):
                     # 時間列（休み文字が入っている方）を強調
                     cell_val.font = bold_font
                     cell_val.fill = gray_fill
        
        # A, B列も中央揃え
        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=1).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=row, column=2).alignment = Alignment(horizontal='center', vertical='center')
        
        # 最終的なExcelをメモリに保存
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        # セッションに保存
        app.config['LAST_EXCEL'] = final_output
        app.config['LAST_FILENAME'] = f'勤務表{year}{int(month):02d}.xlsx'
        
        # プレビュー用データを作成
        preview_data = []
        for idx, (_, row) in enumerate(best_result.iterrows()):
            row_data = {
                'day': row['日付'],
                'dayOfWeek': row['曜日'],
                'staff': {}
            }
            for staff in staff_columns:
                val = row[staff] if pd.notna(row[staff]) else ''
                is_request = request_mask.get((idx, staff), False)
                row_data['staff'][staff] = {
                    'value': val,
                    'isRequest': is_request
                }
            preview_data.append(row_data)
        
        return jsonify({
            'success': True,
            'attempts': attempts,
            'y_counts': {k: int(v) for k, v in best_y_counts.items()},
            'staffNames': staff_columns,
            'preview': preview_data
        })
    except Exception as e:
        return jsonify({'error': f'生成エラー: {str(e)}'}), 500


@app.route('/download')
def download():
    if 'LAST_EXCEL' not in app.config or app.config['LAST_EXCEL'] is None:
        return jsonify({'error': 'まず勤務表を生成してください'}), 400
    
    excel_bytes = app.config['LAST_EXCEL']
    excel_bytes.seek(0)
    
    filename = app.config.get('LAST_FILENAME', 'shift_schedule.xlsx')
    
    return send_file(
        excel_bytes,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('FLASK_DEBUG', 'true').lower() == 'true'
    app.run(debug=debug, host='0.0.0.0', port=port)
