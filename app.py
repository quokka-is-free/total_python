from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
import csv
from datetime import datetime
import requests
import os
from dotenv import load_dotenv
import win32com.client
import pythoncom
import pandas as pd
from werkzeug.utils import secure_filename
import shutil

load_dotenv()

app = Flask(__name__)

if not os.getenv('FLASK_SECRET_KEY'):
    app.secret_key = 'temporary-secret-key-for-dev'
    print("Warning: FLASK_SECRET_KEY not set. Using temporary key.")
else:
    app.secret_key = os.getenv('FLASK_SECRET_KEY')

api_key = os.getenv('KAKAO_API_KEY')

app.config['PERMANENT_SESSION_LIFETIME'] = 600
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

IC_COORDINATES = {
    "논산ic": {"x": "127.0896", "y": "36.2041"},
    "서울ic": {"x": "127.1045", "y": "37.5997"},
}

# --- Helper Functions ---

def get_coordinates(address):
    if not address or not address.strip():
        return None, None
    
    address_lower = address.lower().replace(" ", "")
    if address_lower in IC_COORDINATES:
        return IC_COORDINATES[address_lower]["x"], IC_COORDINATES[address_lower]["y"]
    
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {api_key}"}
    params = {"query": address}
    
    try:
        response = requests.get(url, headers=headers, params=params, timeout=5)
        if response.status_code == 200:
            data = response.json()
            if data["documents"]:
                return data["documents"][0]["x"], data["documents"][0]["y"]
    except Exception as e:
        print(f"API Error: {e}")
    return None, None

def get_toll_distance(origin, destination):
    origin_x, origin_y = get_coordinates(origin)
    dest_x, dest_y = get_coordinates(destination)
    if not origin_x or not dest_x:
        return "주소 변환 실패"
    
    url = "https://apis-navi.kakaomobility.com/v1/directions"
    headers = {"Authorization": f"KakaoAK {api_key}"}
    params = {
        "origin": f"{origin_x},{origin_y}",
        "destination": f"{dest_x},{dest_y}",
        "priority": "DISTANCE"
    }
    
    try:
        response = requests.get(url, headers=headers, params=params, timeout=5)
        if response.status_code == 200:
            data = response.json()
            distance = data["routes"][0]["summary"]["distance"] / 1000
            return f"{distance:.2f} km"
    except Exception:
        pass
    return "거리 계산 실패"

def get_username_by_id(user_id):
    try:
        with open('users.csv', 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if row[0] == user_id:
                    return row[1]
    except FileNotFoundError:
        pass
    return user_id

def get_department_by_id(user_id):
    try:
        with open('users.csv', 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if row[0] == user_id:
                    return row[3]
    except FileNotFoundError:
        pass
    return "미등록"

def get_workplace_by_id(user_id):
    try:
        with open('users.csv', 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if row[0] == user_id:
                    return row[4] if row[4] in ['논산', '대전', '수원'] else '논산'
    except FileNotFoundError:
        pass
    return '논산'

# --- Routes ---

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user_id = request.form.get('user_id')
        password = request.form.get('password')
        try:
            with open('users.csv', 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    # [보안 Note] 실제 운영 시에는 hash 비교(check_password_hash) 사용 필수
                    if row[0] == user_id and row[2] == password:
                        session['logged_in'] = True
                        session['username'] = user_id
                        session['realname'] = row[1]
                        session.permanent = False
                        if user_id == 'admin':
                            return redirect(url_for('admin_dashboard'))
                        else:
                            return redirect(url_for('index'))
        except FileNotFoundError:
            pass
        return render_template('login.html', error="로그인 실패")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/local_trip', methods=['GET', 'POST'])
def local_trip():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    error_message = None
    user_id = session.get('username')
    if request.method == 'POST':
        trip_date = request.form.get('trip_date')
        departure_time = request.form.get('departure_time')
        origin = request.form.get('origin')
        car_number = request.form.get('car_number')
        purpose = request.form.get('purpose')
        destination = request.form.get('destination')
        submit_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        if not origin or not destination:
            error_message = "출발지와 목적지를 모두 입력해주세요."
        else:
            distance = get_toll_distance(origin, destination)
            if "실패" in distance:
                error_message = distance
            else:
                with open('local_trips.csv', 'a', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    # [보안 4] CSV Injection 방지 (입력값 검증)
                    safe_origin = "'" + origin if origin.startswith(('=', '+', '-', '@')) else origin
                    safe_dest = "'" + destination if destination.startswith(('=', '+', '-', '@')) else destination
                    writer.writerow([user_id, submit_time, trip_date, departure_time, safe_origin, car_number, purpose, safe_dest, distance])
                return redirect(url_for('local_trip'))

    try:
        with open('local_trips.csv', 'r', encoding='utf-8') as f:
            local_trips = [row for row in csv.reader(f) if row[0] == user_id]
    except FileNotFoundError:
        local_trips = []

    filter_date = request.form.get('filter_date')
    if filter_date:
        local_trips = [trip for trip in local_trips if trip[2] == filter_date]

    return render_template('local_trip.html', trips=local_trips, error_message=error_message)

@app.route('/outdoor_trip', methods=['GET', 'POST'])
def outdoor_trip():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    error_message = None
    user_id = session.get('username')
    if request.method == 'POST':
        trip_date = request.form.get('trip_date')
        departure_time = request.form.get('departure_time')
        origin = request.form.get('origin')
        car_number = request.form.get('car_number')
        purpose = request.form.get('purpose')
        destination = request.form.get('destination')
        submit_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        if not origin or not destination:
            error_message = "출발지와 목적지를 모두 입력해주세요."
        else:
            distance = get_toll_distance(origin, destination)
            if "실패" in distance:
                error_message = distance
            else:
                with open('outdoor_trips.csv', 'a', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow([user_id, submit_time, trip_date, departure_time, origin, car_number, purpose, destination, distance])
                return redirect(url_for('outdoor_trip'))

    try:
        with open('outdoor_trips.csv', 'r', encoding='utf-8') as f:
            outdoor_trips = [row for row in csv.reader(f) if row[0] == user_id]
    except FileNotFoundError:
        outdoor_trips = []

    filter_date = request.form.get('filter_date')
    if filter_date:
        outdoor_trips = [trip for trip in outdoor_trips if trip[2] == filter_date]

    return render_template('outdoor_trip.html', trips=outdoor_trips, error_message=error_message)

@app.route('/admin_dashboard')
def admin_dashboard():
    if not session.get('logged_in') or session.get('username') != 'admin':
        return redirect(url_for('login'))
    return render_template('admin_dashboard.html')

@app.route('/admin_trips', methods=['GET', 'POST'])
def admin_trips():
    if not session.get('logged_in') or session.get('username') != 'admin':
        return redirect(url_for('admin_dashboard'))

    if request.method == 'POST':
        user_id = request.form.get('user_id')
        username = request.form.get('username')
        password = request.form.get('password')
        department = request.form.get('department')
        workplace = request.form.get('workplace')
        position = request.form.get('position')
        email = request.form.get('email')
        register_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        users = []
        user_found = False
        try:
            with open('users.csv', 'r', encoding='utf-8') as f:
                users = list(csv.reader(f))
            for i, user in enumerate(users):
                if user[0] == user_id:
                    users[i] = [user_id, username, password, department, workplace, position, email, register_date]
                    user_found = True
                    break
            if not user_found:
                users.append([user_id, username, password, department, workplace, position, email, register_date])
        except FileNotFoundError:
            users = [[user_id, username, password, department, workplace, position, email, register_date]]

        with open('users.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(users)

    try:
        with open('users.csv', 'r', encoding='utf-8') as f:
            users = list(csv.reader(f))
    except FileNotFoundError:
        users = []

    try:
        with open('local_trips.csv', 'r', encoding='utf-8') as f:
            local_trips = list(csv.reader(f))
    except FileNotFoundError:
        local_trips = []

    try:
        with open('outdoor_trips.csv', 'r', encoding='utf-8') as f:
            outdoor_trips = list(csv.reader(f))
    except FileNotFoundError:
        outdoor_trips = []

    local_trips_display = []
    for trip in local_trips:
        username_val = get_username_by_id(trip[0])
        local_trips_display.append([username_val, trip[1], trip[2], trip[3], trip[4], trip[7], trip[6], trip[5], trip[8]])

    outdoor_trips_display = []
    for trip in outdoor_trips:
        username_val = get_username_by_id(trip[0])
        outdoor_trips_display.append([username_val, trip[1], trip[2], trip[3], trip[4], trip[7], trip[6], trip[5], trip[8]])

    return render_template('admin.html', users=users, local_trips=local_trips_display, outdoor_trips=outdoor_trips_display)

@app.route('/delete_user', methods=['POST'])
def delete_user():
    if not session.get('logged_in') or session.get('username') != 'admin':
        return redirect(url_for('admin_dashboard'))
    user_id_to_delete = request.form.get('user_id')
    try:
        with open('users.csv', 'r', encoding='utf-8') as f:
            users = list(csv.reader(f))
        users = [user for user in users if user[0] != user_id_to_delete]
        with open('users.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(users)
    except FileNotFoundError:
        pass
    return redirect(url_for('admin_trips'))

@app.route('/delete_local_trip', methods=['POST'])
def delete_local_trip():
    if not session.get('logged_in') or session.get('username') != 'admin':
        return redirect(url_for('admin_dashboard'))

    submit_time = request.form.get('submit_time')
    try:
        with open('local_trips.csv', 'r', encoding='utf-8') as f:
            trips = list(csv.reader(f))
        trips = [trip for trip in trips if trip[1] != submit_time]
        with open('local_trips.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(trips)
    except FileNotFoundError:
        pass
    return redirect(url_for('admin_trips'))

@app.route('/delete_outdoor_trip', methods=['POST'])
def delete_outdoor_trip():
    if not session.get('logged_in') or session.get('username') != 'admin':
        return redirect(url_for('admin_dashboard'))

    submit_time = request.form.get('submit_time')
    try:
        with open('outdoor_trips.csv', 'r', encoding='utf-8') as f:
            trips = list(csv.reader(f))
        trips = [trip for trip in trips if trip[1] != submit_time]
        with open('outdoor_trips.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(trips)
    except FileNotFoundError:
        pass
    return redirect(url_for('admin_trips'))

@app.route('/admin_attendance', methods=['GET', 'POST'])
def admin_attendance():
    if not session.get('logged_in') or session.get('username') != 'admin':
        return redirect(url_for('admin_dashboard'))

    attendance_data = []
    try:
        with open('attendance.csv', 'r', encoding='utf-8') as f:
            attendance_data = list(csv.reader(f))
    except FileNotFoundError:
        pass

    approvals = {}
    try:
        with open('approvals.csv', 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader) 
            for row in reader:
                approvals[(row[0], str(row[1]))] = row[2]
    except FileNotFoundError:
        pass
    except StopIteration:
        pass

    if request.method == 'POST' and 'file' in request.files:
        file = request.files['file']
        if file and file.filename.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
            expected_columns = ['발생일자', '발생시각', '일시', '사원번호', '이름', '모드']
            # [보안 5] 컬럼 유효성 검사 강화
            if not all(col in df.columns for col in expected_columns):
                return jsonify({'error': '엑셀 형식이 올바르지 않습니다. 필요한 컬럼: ' + ', '.join(expected_columns)})
            
            existing_df = pd.DataFrame(attendance_data[1:], columns=attendance_data[0]) if attendance_data else pd.DataFrame(columns=['사원번호', '이름', '부서', '출근시간', '퇴근시간', '날짜', '결재상태', '근무지', '비고'])
            existing_keys = set(zip(existing_df['사원번호'], existing_df['날짜'].astype(str)))
            
            df['날짜'] = df['발생일자'].astype(str) + ' ' + df['발생시각'].astype(str)
            attendance_by_date = {}
            for date in df['발생일자'].unique():
                date_df = df[df['발생일자'] == date]
                attendance = date_df[date_df['모드'].isin(['출근', '퇴근'])].copy()
                attendance['부서'] = attendance['사원번호'].apply(get_department_by_id)
                attendance_grouped = attendance.groupby('사원번호').agg({
                    '이름': 'first', '부서': 'first', '날짜': list, '모드': 'first'
                }).reset_index()
                
                attendance_processed = []
                for _, row in attendance_grouped.iterrows():
                    key = (row['사원번호'], str(row['날짜'][0]) if row['날짜'] else None)
                    if key not in existing_keys:
                        dates = row['날짜']
                        출근시간 = next((d for d in dates if '출근' in attendance[attendance['날짜'] == d]['모드'].values), None)
                        퇴근시간 = next((d for d in dates if '퇴근' in attendance[attendance['날짜'] == d]['모드'].values), None)
                        workplace = get_workplace_by_id(row['사원번호'])
                        
                        mode = row['모드'].lower()
                        remark = '정상'
                        if '출장' in mode:
                            if '시내' in mode: remark = '출(시내)'
                            elif '시외' in mode: remark = '출(시외)'
                        elif '연차' in mode: remark = '연차'
                        elif '반차' in mode: remark = '반차'
                        elif '휴직' in mode: remark = '휴직'
                        elif '육아' in mode: remark = '육아'
                        
                        attendance_processed.append({
                            '사원번호': row['사원번호'],
                            '이름': row['이름'],
                            '부서': row['부서'],
                            '출근시간': 출근시간,
                            '퇴근시간': 퇴근시간,
                            '날짜': row['날짜'][0] if row['날짜'] else None,
                            '결재상태': '대기',
                            '근무지': workplace,
                            '비고': remark
                        })
                if attendance_processed:
                    attendance_by_date[date] = pd.DataFrame(attendance_processed)
            
            if attendance_by_date:
                df_processed = pd.concat(attendance_by_date.values(), ignore_index=True)
                if not existing_df.empty:
                    df_processed = pd.concat([existing_df, df_processed], ignore_index=True)
                df_processed.to_csv('attendance.csv', index=False, encoding='utf-8')
            
            if not os.path.exists('approvals.csv'):
                with open('approvals.csv', 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(['사원번호', '날짜', '상태'])
            return jsonify({'success': True})
        return jsonify({'error': '유효한 엑셀 파일을 업로드해주세요.'})

    if request.method == 'POST' and request.form.get('action') == 'delete_data':
        employee_id = request.form.get('employee_id')
        date = request.form.get('date')
        try:
            with open('attendance.csv', 'r', encoding='utf-8') as f:
                attendance_data = list(csv.reader(f))
            if attendance_data:
                headers = attendance_data[0]
                data = attendance_data[1:]
                updated_data = [headers]
                for row in data:
                    if not (row[headers.index('사원번호')] == employee_id and row[headers.index('날짜')] == date):
                        updated_data.append(row)
                with open('attendance.csv', 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerows(updated_data)
                
                with open('approvals.csv', 'r', encoding='utf-8') as f:
                    approvals_data = list(csv.reader(f))
                if approvals_data:
                    headers = approvals_data[0]
                    data = approvals_data[1:]
                    updated_approvals = [headers]
                    for row in data:
                        if not (row[headers.index('사원번호')] == employee_id and row[headers.index('날짜')] == date):
                            updated_approvals.append(row)
                    with open('approvals.csv', 'w', newline='', encoding='utf-8') as f:
                        writer = csv.writer(f)
                        writer.writerows(updated_approvals)
            return redirect(url_for('admin_attendance'))
        except Exception as e:
            return jsonify({'error': '데이터 삭제 실패'}), 500

    df = pd.DataFrame(attendance_data[1:], columns=attendance_data[0]) if attendance_data else pd.DataFrame()
    
    if not df.empty and '사원번호' in df.columns:
        if '근무지' not in df.columns: 
            df['근무지'] = df['사원번호'].apply(get_workplace_by_id)
            
    if '결재상태' not in df.columns: df['결재상태'] = '대기'
    if '비고' not in df.columns: df['비고'] = '정상'
    
    locations = ['논산', '대전', '수원']
    attendance_by_loc = {}
    for loc in locations:
        loc_df = df[df['근무지'] == loc] if not df.empty else pd.DataFrame()
        departments = loc_df['부서'].unique().tolist() if '부서' in loc_df.columns else []
        attendance_by_loc[loc] = {}
        for dept in departments:
            dept_df = loc_df[loc_df['부서'] == dept]
            records = dept_df.to_dict('records')
            for record in records:
                key = (record['사원번호'], str(record['날짜']))
                record['결재상태'] = approvals.get(key, record['결재상태'])
            attendance_by_loc[loc][dept] = records

    if request.method == 'POST' and request.form.get('action') == 'approve_all':
        loc = request.form.get('loc')
        dept = request.form.get('dept')
        with open('attendance.csv', 'r', encoding='utf-8') as f:
            attendance_data = list(csv.reader(f))
        df = pd.DataFrame(attendance_data[1:], columns=attendance_data[0])
        if loc and dept and '부서' in df.columns:
            df = df[df['부서'] == dept]
        for index, row in df.iterrows():
            key = (row['사원번호'], str(row['날짜']))
            if key not in approvals or approvals.get(key) == '대기':
                with open('approvals.csv', 'a', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow([row['사원번호'], row['날짜'], '승인'])
                approvals[key] = '승인'
        return redirect(url_for('admin_attendance'))

    if request.method == 'POST' and 'action' in request.form and request.form['action'] == 'approve':
        employee_id = request.form.get('employee_id')
        date = request.form.get('date')
        key = (employee_id, str(date))
        if key not in approvals or approvals.get(key) == '대기':
            with open('approvals.csv', 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([employee_id, date, '승인'])
            approvals[key] = '승인'
        return redirect(url_for('admin_attendance'))

    if request.method == 'POST' and request.form.get('action') == 'delete_all':
       
        return redirect(url_for('admin_attendance'))

    return render_template('admin_attendance.html', attendance_by_loc=attendance_by_loc, locations=locations)

@app.route('/generate_attendance_excel', methods=['GET'])
def generate_attendance_excel():
    if not session.get('logged_in') or session.get('username') != 'admin':
        return redirect(url_for('admin_dashboard'))
    
    try:

        with open('attendance.csv', 'r', encoding='utf-8') as f:
            attendance_data = list(csv.reader(f))
        
        df = pd.DataFrame(attendance_data[1:], columns=attendance_data[0])
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        save_folder = os.path.join(base_dir, 'downloads')
        if not os.path.exists(save_folder): os.makedirs(save_folder)
        
        filename = f'approved_attendance_{datetime.now().strftime("%Y%m%d")}.xlsx'
        output_path = os.path.join(save_folder, filename)
        
        df.to_excel(output_path, index=False)
        
        return send_file(output_path, as_attachment=True, download_name=filename)
    except Exception as e:
        return render_template('admin_attendance.html', error=f"엑셀 생성 오류: {str(e)}")

@app.route('/expense_claim')
def expense_claim():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    trip_date = request.args.get('trip_date')
    origin = request.args.get('origin')
    destination = request.args.get('destination')
    car_number = request.args.get('car_number')
    purpose = request.args.get('purpose')

    return render_template('expense_claim.html', 
                          trip_date=trip_date, 
                          origin=origin, 
                          destination=destination, 
                          car_number=car_number, 
                          purpose=purpose, 
                          location="")

@app.route('/generate_expense_excel', methods=['POST'])
def generate_expense_excel():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    try:
        pythoncom.CoInitialize()
        data = request.form
        
        safe_username = secure_filename(session.get('username'))
        trip_date = secure_filename(data.get('trip_date', ''))
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, 'travel.xlsx')
        save_folder = os.path.join(base_dir, 'downloads')
        
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)
            
        filename = f"expense_report_{safe_username}_{trip_date}.xlsx"
        save_path = os.path.join(save_folder, filename)

        if not os.path.exists(template_path):
            return "Error: 'travel.xlsx' template not found in project folder.", 500

        shutil.copy(template_path, save_path)

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb = excel.Workbooks.Open(save_path)
            ws = wb.ActiveSheet
            
            ws.Range('C10').Value = data.get('trip_date')
            ws.Range('C11').Value = data.get('location')
            ws.Range('C12').Value = "차량"
            ws.Range('C13').Value = data.get('car_number')
            ws.Range('B17').Value = data.get('purpose')
            ws.Range('C28').Value = data.get('origin')
            ws.Range('E28').Value = data.get('destination')
            
            try: toll = float(data.get('toll_fee', '0').replace(',', ''))
            except: toll = 0
            ws.Range('G28').Value = toll

            distance_str = get_toll_distance(data.get('origin'), data.get('destination'))
            try: dist = float(distance_str.replace(' km', ''))
            except: dist = 0
            ws.Range('I28').Value = dist

            wb.Save()
            wb.Close()
        except Exception as e:
            raise e
        finally:
            excel.Quit()

        return send_file(save_path, as_attachment=True, download_name=filename)

    except Exception as e:
        return f"Error: {str(e)}", 500
    finally:
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    
    app.run(host='0.0.0.0', port=8000, debug=False)
