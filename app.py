from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import json, io, os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
CORS(
    app,
    origins=[
        "https://student-attendance-tracking-system-3un4pgeat.vercel.app",
    ],
)

STUDENTS = {
    "A1": {"label":"المجموعة A1","time":"10:00 - 12:00","students":["احمد سعيد شرقاوي نصر الرقيق","اساف عادل كامل فرج الله","اسلام اشرف عبد المجيد عبد الفتاح","حبيبه محمد كمال محمد عامر","حال حاتم محمد الحسيني عبد العظيم","خلود محمد مصطفى بشرى","رانيا يوسف راشد عبد المجيد","رزق محمد رزق محمد","روان شريف عطيه عبدالفتاح","روان مصطفى مصطفى السيد على","زياد احمد مرعى عبد اللطيف قاسم","سما حسام ممدوح شوقي","عبد الرحمن ابراهيم سيد خليل","عبد الرحمن احمد سعيد كامل سالم","عبد الرحمن طارق عبد الحميد والى","علاء أحمد عبدالله محمد","عمار ياسر فتحي احمد"]},
    "B1": {"label":"المجموعة B1","time":"8:00 - 10:00","students":["فتحي محمد فتحي محمد الشاذلي","لوجين ابراهيم مصطفى محمد","ليلى شعبان احمد امام","محمد أشرف يحيى عبدالله","محمد محمد محمد محمد العراقي","محمد هيثم محمد فضل على مسعود","محمد ياسر على الجميل بدر","محمود احمد محمود عبدالجواد كريم","مصطفى محمد عبد الفتاح محمد","ملك محمد امام ابراهيم","منه جمال محمد دمرداش","ميار حاتم عبده سعد عثمان","ندى محمد ابراهيم محمد التميمي","هادى شريف محمد الهادي محمد على","هشام سعيد صالح عثمان","والء على سيد محمد عبد الجواد","يمنى إسماعيل إسماعيل إمام حسن","يوسف احمد جلال محمد احمد عامر","يوسف اشرف ابراهيم ذكى","يوسف طارق محمد امام محمود","يوسف محمد جمال احمد شربيني فروح"]},
    "A2": {"label":"المجموعة A2","time":"14:00 - 16:00","students":["احمد تميم فيصل تميم","احمد محمد محمد احمد امين","جلال هاني جلال فتحي","حبيبه اسامه عبد المنعم عبد الرحيم","رؤى هشام فهمى ريحان","رحمه فؤاد سيد احمد مهدى","زياد محمد فتحي عباس السيد","ساره صالح محمد محمدين","شهد ابراهيم محمد ابراهيم"]},
    "B2": {"label":"المجموعة B2","time":"12:00 - 14:00","students":["عبد الله نجاح حامد احمد عزازي","عمرو علاء الدين سيد إبراهيم خليل","مؤمن على فتحي عبد العاطي مبروك","ماريا فيكتور عياد يوسف ابراهيم","مازن محمد رائف حافظ","مريم محمد حافظ عبد العال","مصطفى عبد المحسن مصطفى محمد","منه محمد ابو الحجاج محمد","نور الدين ادهم نور الدين مصطفى","رسمية مسعود سعد محمد ياسين"]},
    "A3": {"label":"المجموعة A3","time":"16:00 - 18:00","students":["ابراهيم مسعود سيد ابراهيم","احمد سعيد فتحي السيد محمد","اسراء سيد عبد العزيز جلال","بافلي حشمت بولص حنا","جنه هاني صالح عبد المنعم احمد","جوزيف عادل عزيز عياد","جومانا هشام سيد بدوي محمد","حازم حسنى فهيم محمد","شهد محمد رمضان إبراهيم","عبد الله حسن السيد احمد موسى","عبد الله طارق مصباح عبد الحميد","عمر اسماعيل حلمى أبو ضيف"]},
    "B3": {"label":"المجموعة B3","time":"18:00 - 20:00","students":["عمر ايمن حنفى محمود","عمر كمال عبد النبي طلبة زايد","كنزى حمدى سالم أمين سالم","محمد إبراهيم طويل محمد سالم","محمد سيد محمد حسن","مصطفى محمد محمود مصطفى سويدان","ملك طارق عواد عيد","منة الله عبد العال على عبد العال"]},
    "A4": {"label":"المجموعة A4","time":"16:00 - 18:00","students":["احمد ايهاب بديع محمد الماظ","احمد سعيد صابر كامل احمد","احمد عرفه عمر احمد مكاوي","الاء كرم خليفه عبد ربه عبد الرحمن","اميرة صبحي عبد الشافي حسن الصغير","بسملة محمد عبد العزيز عبد الحافظ","بسمه هاني حفني محمد","تونى ايمن كمال يونان","جنى ياسر على حامد","حازم محمود عبد الحميد محمود محمد","روان رضا رمضان حسن وهدان","روان محمد سالم حسن","روان وليد احمد محمد","زياد عادل امين عبد الصمد","زينب حسام الدين محمد حسين","سارة حسين عبد الحليم الجمال","سلفانا عاطف نصرى عزيز","شمس الدين شوقي عبد الله عبد القادر","شهد حسنى حامد احمد حسن","صبحى عصام صبحى عبد القوى","صفية يحيى احمد على","عمار عماد باهى منصور"]},
    "B4": {"label":"المجموعة B4","time":"18:00 - 20:00","students":["عمر أحمد شلقامى ذكي","عمر خالد عبد الله احمد حسن","عهد عبد العليم محمود محمود قديرة","فريدة حاتم عبد القادر عبد العزيز محمد","محمد حسين عبد التواب علي","محمد رمضان عبد التواب السيد","محمد محمود عبد الحميد محمود","مريم بخيت فكرى نصير","مريم ياسر حسن قطب حسن","ملك محمد محمد صالح النبالوى","منه الله نبيل سعيد محمود","منه محمد السيد حسن خميس","ناصر حسن ثابت حسين","نرمين حنا سعيد حنا","نور احمد سعيد محمد احمد ابو العز","نور الهدى عمار على محمد عبدالمجيد","نور محمد حمدى حسن","نورهان رفاعي محجوب خليفة محمد","هاجر محمد احمد عبد الهادي","هاله السيد احمد سالم","هبة الله مدحت حسن صقر مبارك","ياسمين عبد الفتاح محمد عبد الفتاح","يمنى يوسف احمد حامد رضوان","يوسف ابراهيم سيد ابراهيم","يوسف عادل بديع نخلة حنين","يوسف محمد سيد احمد احمد المدني"]},
}

attendance_store = {}

@app.route('/')
def index():
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), 'index.html')

@app.route('/api/groups', methods=['GET'])
def get_groups():
    return jsonify([{"id":k,"label":v["label"],"time":v["time"],"count":len(v["students"])} for k,v in STUDENTS.items()])

@app.route('/api/students/<group_id>', methods=['GET'])
def get_students(group_id):
    if group_id not in STUDENTS:
        return jsonify({"error":"not found"}), 404
    g = STUDENTS[group_id]
    return jsonify({
        "label": g["label"],
        "time": g["time"],
        "students": g["students"],
        "attendance": attendance_store.get(group_id, {}),
    })

@app.route('/api/attendance/<group_id>', methods=['GET'])
def get_attendance_history(group_id):
    if group_id not in STUDENTS:
        return jsonify({"error":"not found"}), 404
    return jsonify(attendance_store.get(group_id, {}))

@app.route('/api/attendance', methods=['POST'])
def save_attendance():
    data = request.json
    gid = data.get('group_id')
    records = data.get('records', {})
    date = data.get('date', datetime.now().strftime('%Y-%m-%d'))
    if gid not in attendance_store:
        attendance_store[gid] = {}
    attendance_store[gid][date] = records
    return jsonify({"status":"saved"})

@app.route('/api/export/<group_id>', methods=['GET'])
def export_attendance(group_id):
    date = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))
    if group_id not in STUDENTS:
        return jsonify({"error":"not found"}), 404
    group = STUDENTS[group_id]
    records = attendance_store.get(group_id, {}).get(date, {})
    wb = Workbook()
    ws = wb.active
    ws.title = "كشف الحضور"
    ws.sheet_view.rightToLeft = True
    hf = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    hfill = PatternFill('solid', start_color='1a3a5c')
    pfill = PatternFill('solid', start_color='d4edda')
    afill = PatternFill('solid', start_color='fee2e2')
    ctr = Alignment(horizontal='center', vertical='center')
    rgt = Alignment(horizontal='right', vertical='center')
    thin = Side(style='thin', color='cccccc')
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells('A1:D1')
    ws['A1'] = f"كشف حضور - {group['label']} | {group['time']} | تاريخ: {date}"
    ws['A1'].font = Font(name='Arial', bold=True, size=14, color='1a3a5c')
    ws['A1'].alignment = ctr
    ws['A1'].fill = PatternFill('solid', start_color='e8f0fe')
    ws.row_dimensions[1].height = 30
    for col, h in enumerate(['م','اسم الطالب','الحضور','ملاحظات'], 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = hf; c.fill = hfill; c.alignment = ctr; c.border = brd
    ws.row_dimensions[2].height = 25
    def is_present(raw):
        return raw in ('present', 'حاضر', 'متاخر')

    present_count = 0
    for i, name in enumerate(group['students'], 1):
        row = i + 2
        raw = records.get(name, 'absent')
        status = 'present' if is_present(raw) else 'absent'
        if status == 'present': present_count += 1
        fill = pfill if status == 'present' else afill
        label = 'متأخر ⏰' if raw == 'متاخر' else ('حاضر ✔' if status == 'present' else 'غائب ✘')
        status_text = label
        for col, (val, aln) in enumerate(zip([i, name, status_text, ''], [ctr,rgt,ctr,ctr]), 1):
            c = ws.cell(row=row, column=col, value=val)
            c.alignment = aln; c.border = brd; c.font = Font(name='Arial', size=11)
            if col in (1, 3): c.fill = fill
        ws.row_dimensions[row].height = 22
    sr = len(group['students']) + 3
    ws.merge_cells(f'A{sr}:B{sr}')
    ws[f'A{sr}'] = f'إجمالي الحاضرين: {present_count} / {len(group["students"])}'
    ws[f'A{sr}'].font = Font(name='Arial', bold=True, size=11, color='155724')
    ws[f'A{sr}'].fill = pfill; ws[f'A{sr}'].alignment = ctr; ws[f'A{sr}'].border = brd
    ws.merge_cells(f'C{sr}:D{sr}')
    ws[f'C{sr}'] = f'إجمالي الغائبين: {len(group["students"]) - present_count}'
    ws[f'C{sr}'].font = Font(name='Arial', bold=True, size=11, color='721c24')
    ws[f'C{sr}'].fill = afill; ws[f'C{sr}'].alignment = ctr; ws[f'C{sr}'].border = brd
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 20
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"attendance_{group_id}_{date}.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=False, port=5050)
