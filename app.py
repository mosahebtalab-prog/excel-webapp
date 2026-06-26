import os
import openpyxl
from openpyxl import load_workbook
from flask import Flask, render_template, request, redirect, url_for
import jdatetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'

# ========== تنظیم مسیر فایل اکسل ==========
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'data.xlsx')

# ========== توابع مدیریت فایل اکسل ==========
def ensure_excel_file():
    """اگر فایل اکسل وجود نداشته باشد، با ساختار کامل می‌سازد"""
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        
        # برگه بیماران
        ws_patients = wb.active
        ws_patients.title = 'patients'
        ws_patients.append(['id', 'name', 'phone', 'disease', 'created_at'])
        
        # برگه ویزیت‌ها
        ws_visits = wb.create_sheet('visits')
        ws_visits.append(['id', 'patient_id', 'visit_date', 'reason', 'treatment', 'cost', 'created_at'])
        
        wb.save(EXCEL_FILE)
        print(f"✅ فایل اکسل در مسیر {EXCEL_FILE} ساخته شد.")

def get_all_patients():
    """دریافت همه بیماران از اکسل"""
    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['patients']
    patients = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            patients.append({
                'id': row[0],
                'name': row[1] or '',
                'phone': row[2] or '',
                'disease': row[3] or '',
                'created_at': row[4] or ''
            })
    wb.close()
    return patients

def get_patient_by_id(patient_id):
    """دریافت یک بیمار با ID"""
    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['patients']
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == int(patient_id):
            wb.close()
            return {
                'id': row[0],
                'name': row[1] or '',
                'phone': row[2] or '',
                'disease': row[3] or '',
                'created_at': row[4] or ''
            }
    wb.close()
    return None

def get_patient_visits(patient_id):
    """دریافت ویزیت‌های یک بیمار"""
    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['visits']
    visits = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[1] == int(patient_id):
            visits.append({
                'id': row[0],
                'patient_id': row[1],
                'visit_date': row[2] or '',
                'reason': row[3] or '',
                'treatment': row[4] or '',
                'cost': row[5] or 0,
                'created_at': row[6] or ''
            })
    wb.close()
    return visits

def get_total_income():
    """محاسبه درآمد کل از همه ویزیت‌ها"""
    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['visits']
    total = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[5] is not None:
            try:
                total += int(row[5])
            except:
                pass
    wb.close()
    return total

def get_income_by_date_range(start_date, end_date):
    """محاسبه درآمد در بازه زمانی شمسی"""
    try:
        start_jd = jdatetime.datetime.strptime(start_date, '%Y/%m/%d')
        end_jd = jdatetime.datetime.strptime(end_date, '%Y/%m/%d')
    except:
        return 0, []

    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['visits']
    total_income = 0
    visits_in_range = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[2] is not None:
            try:
                visit_date = row[2]
                if '/' in visit_date:
                    visit_jd = jdatetime.datetime.strptime(visit_date, '%Y/%m/%d')
                    if start_jd <= visit_jd <= end_jd:
                        cost = int(row[5]) if row[5] else 0
                        total_income += cost
                        visits_in_range.append({
                            'id': row[0],
                            'patient_id': row[1],
                            'visit_date': visit_date,
                            'reason': row[3] or '',
                            'treatment': row[4] or '',
                            'cost': cost
                        })
            except:
                pass
    
    wb.close()
    return total_income, visits_in_range

# ========== مسیرها (Routes) ==========
@app.route('/')
def index():
    patients = get_all_patients()
    total_income = get_total_income()
    return render_template('index.html', patients=patients, total_income=total_income)

@app.route('/add_patient', methods=['POST'])
def add_patient():
    name = request.form.get('name', '').strip()
    phone = request.form.get('phone', '').strip()
    disease = request.form.get('disease', '').strip()

    if not name:
        return redirect(url_for('index'))

    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['patients']

    # پیدا کردن آخرین ID
    max_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[0] > max_id:
            max_id = row[0]

    new_id = max_id + 1
    now = jdatetime.datetime.now().strftime('%Y/%m/%d %H:%M')
    ws.append([new_id, name, phone, disease, now])
    wb.save(EXCEL_FILE)
    wb.close()
    return redirect(url_for('index'))

@app.route('/delete_patient/<int:patient_id>')
def delete_patient(patient_id):
    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['patients']
    row_to_delete = None
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == patient_id:
            row_to_delete = idx
            break
    if row_to_delete:
        ws.delete_rows(row_to_delete)
        wb.save(EXCEL_FILE)
    wb.close()
    return redirect(url_for('index'))

@app.route('/profile/<int:patient_id>')
def profile(patient_id):
    patient = get_patient_by_id(patient_id)
    if not patient:
        return redirect(url_for('index'))
    visits = get_patient_visits(patient_id)
    return render_template('profile.html', patient=patient, visits=visits)

@app.route('/add_visit', methods=['POST'])
def add_visit():
    patient_id = request.form.get('patient_id', '').strip()
    visit_date = request.form.get('visit_date', '').strip()
    reason = request.form.get('reason', '').strip()
    treatment = request.form.get('treatment', '').strip()
    cost = request.form.get('cost', '').strip()

    if not patient_id or not visit_date:
        return redirect(url_for('index'))

    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['visits']

    # پیدا کردن آخرین ID برای ویزیت
    max_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[0] > max_id:
            max_id = row[0]

    new_id = max_id + 1
    now = jdatetime.datetime.now().strftime('%Y/%m/%d %H:%M')
    ws.append([new_id, int(patient_id), visit_date, reason, treatment, int(cost) if cost else 0, now])
    wb.save(EXCEL_FILE)
    wb.close()
    return redirect(url_for('profile', patient_id=patient_id))

@app.route('/delete_visit/<int:visit_id>/<int:patient_id>')
def delete_visit(visit_id, patient_id):
    ensure_excel_file()
    wb = load_workbook(EXCEL_FILE)
    ws = wb['visits']
    row_to_delete = None
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == visit_id:
            row_to_delete = idx
            break
    if row_to_delete:
        ws.delete_rows(row_to_delete)
        wb.save(EXCEL_FILE)
    wb.close()
    return redirect(url_for('profile', patient_id=patient_id))

@app.route('/income_report', methods=['GET', 'POST'])
def income_report():
    total_income = 0
    visits = []
    start_date = ''
    end_date = ''

    if request.method == 'POST':
        start_date = request.form.get('start_date', '').strip()
        end_date = request.form.get('end_date', '').strip()

        if start_date and end_date:
            total_income, visits = get_income_by_date_range(start_date, end_date)

    return render_template('income_report.html',
                         total_income=total_income,
                         visits=visits,
                         start_date=start_date,
                         end_date=end_date)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
