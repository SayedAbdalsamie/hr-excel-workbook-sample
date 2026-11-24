"""Builds the sample HR & operations workbook for a security company (all data is fictional sample data)."""
import random, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter as L
random.seed(7)
wb = Workbook()
HF = PatternFill("solid", fgColor="14285A"); HFont = Font(bold=True, color="FFFFFF", name="Arial", size=11)
thin = Side(style="thin", color="BBBBBB"); B = Border(left=thin, right=thin, top=thin, bottom=thin)
def sheet(name, headers, widths, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name; ws.sheet_view.rightToLeft = True
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, i, h); c.fill = HF; c.font = HFont; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = B
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[1].height = 30; ws.freeze_panes = "A2"; return ws
def put(ws, rows):
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            x = ws.cell(r, c, v); x.border = B; x.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(v, (dt.date, dt.datetime)): x.number_format = "yyyy-mm-dd"
def dv(ws, col, formula, n=300, first=2):
    d = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True, errorTitle="قيمة غير صحيحة", error="اختر من القائمة المنسدلة")
    ws.add_data_validation(d); d.add(f"{col}{first}:{col}{n}")
# ---- Lists
lists = {"المسمى الوظيفي": ["حارس أمن", "مشرف أمن", "قائد وردية", "مراقب كاميرات", "سائق", "موظف إداري"],
         "حالة الموظف": ["على رأس العمل", "إجازة", "موقوف", "منتهي الخدمة"],
         "الموقع": ["المقر الرئيسي", "مجمع تجاري", "مصنع", "فعالية"],
         "حالة الحضور": ["حاضر", "غائب", "متأخر", "إجازة", "إجازة مرضية"],
         "حالة الاعتماد": ["معتمد", "قيد المراجعة", "مرفوض"],
         "نوع العهدة": ["زي رسمي", "جهاز لاسلكي", "كشاف", "هاتف", "مفتاح", "سترة واقية"],
         "حالة العهدة": ["مع الموظف", "تم الاستلام", "تالفة", "مفقودة"],
         "حالة المهمة": ["مجدولة", "جارية", "منتهية", "ملغاة"],
         "نوع الإجراء": ["إنذار شفهي", "إنذار كتابي", "خصم", "إيقاف"]}
wl = sheet("القوائم", list(lists.keys()), [20] * len(lists))
for c, (k, v) in enumerate(lists.items(), 1):
    for r, x in enumerate(v, 2): wl.cell(r, c, x).alignment = Alignment(horizontal="center")
def LR(key):
    c = L(list(lists).index(key) + 1); return f"=القوائم!${c}$2:${c}${len(lists[key]) + 1}"
# ---- Employees
names = ["أحمد محمود", "محمد علي", "خالد حسن", "يوسف إبراهيم", "عمر سعيد", "مصطفى كمال", "إسلام فتحي", "طارق عادل", "حسام نبيل", "كريم رضا", "ماجد سمير", "هشام جمال", "ياسر عبد الله", "وليد صلاح", "رامي فؤاد", "سامح ماهر", "بلال أنور", "نادر شريف", "إيهاب لطفي", "فادي منصور"]
emps = []
for i, n in enumerate(names):
    emps.append([f"EMP{i+1:03d}", n, random.choice(lists["المسمى الوظيفي"][:5]), f"01{random.choice([0,1,2,5])}{random.randint(10000000,99999999)}", random.choice(lists["الموقع"]), dt.date(2022, 1, 1) + dt.timedelta(days=random.randint(0, 1000)), "على رأس العمل" if i != 17 else "إجازة"])
we = sheet("الموظفون", ["الرقم الوظيفي", "الاسم", "المسمى الوظيفي", "الهاتف", "موقع العمل", "تاريخ التعيين", "حالة الموظف"], [14, 22, 18, 16, 18, 16, 16])
put(we, emps); dv(we, "C", LR("المسمى الوظيفي")); dv(we, "E", LR("الموقع")); dv(we, "G", LR("حالة الموظف"))
for st, col in [("موقوف", "F8CBAD"), ("منتهي الخدمة", "D9D9D9"), ("إجازة", "FFF2CC")]:
    we.conditional_formatting.add("G2:G300", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", bgColor=col)))
EMPR = "=الموظفون!$A$2:$A$300"
def empdv(ws, col="A", n=500):
    d = DataValidation(type="list", formula1=EMPR, allow_blank=True, showErrorMessage=True, error="الرقم الوظيفي غير موجود في ورقة الموظفين"); ws.add_data_validation(d); d.add(f"{col}2:{col}{n}")
def lookup(r): return f'=IFERROR(VLOOKUP(A{r},الموظفون!$A$2:$B$300,2,FALSE),"")'
# ---- Attendance
wa = sheet("الحضور", ["الرقم الوظيفي", "الاسم", "التاريخ", "الحضور", "الانصراف", "الحالة", "ساعات التأخير", "السبب"], [14, 22, 14, 12, 12, 14, 14, 26])
rows = []
for d in range(10):
    day = dt.date(2026, 11, 1) + dt.timedelta(days=d)
    for e in emps[:12]:
        s = random.choices(["حاضر", "متأخر", "غائب", "إجازة"], [75, 12, 8, 5])[0]
        tin = dt.time(8, 0) if s == "حاضر" else dt.time(8, random.choice([20, 35, 50])) if s == "متأخر" else None
        rows.append([e[0], None, day, tin, dt.time(16, 0) if tin else None, s, None, "ظرف طارئ" if s == "غائب" else "إجازة اعتيادية" if s == "إجازة" else ""])
put(wa, rows)
for r in range(2, 501):
    wa.cell(r, 2, lookup(r)); wa.cell(r, 7, f'=IF(AND(D{r}<>"",D{r}>TIME(8,0,0)),ROUND((D{r}-TIME(8,0,0))*24,2),0)')
    wa.cell(r, 4).number_format = wa.cell(r, 5).number_format = "hh:mm"; wa.cell(r, 3).number_format = "yyyy-mm-dd"
    for c in (2, 7): wa.cell(r, c).border = B; wa.cell(r, c).alignment = Alignment(horizontal="center")
empdv(wa); dv(wa, "F", LR("حالة الحضور"), 500)
wa.conditional_formatting.add("F2:F500", CellIsRule(operator="equal", formula=['"غائب"'], fill=PatternFill("solid", bgColor="F8CBAD")))
wa.conditional_formatting.add("F2:F500", CellIsRule(operator="equal", formula=['"متأخر"'], fill=PatternFill("solid", bgColor="FFF2CC")))
# ---- Declarations
wd = sheet("الإقرارات", ["الرقم الوظيفي", "الاسم", "تاريخ الإقرار", "الموضوع", "حالة الاعتماد", "مكان حفظ المستند"], [14, 22, 14, 30, 16, 28])
subj = ["إقرار استلام لائحة العمل", "إقرار السرية", "إقرار استلام زي", "إقرار سلامة مهنية"]
put(wd, [[e[0], None, dt.date(2026, 10, random.randint(1, 28)), random.choice(subj), random.choice(lists["حالة الاعتماد"][:2] + ["معتمد"]), f"الأرشيف / دولاب 2 / ملف {i+1}"] for i, e in enumerate(emps[:15])])
for r in range(2, 301): wd.cell(r, 2, lookup(r)).border = B
empdv(wd, n=300); dv(wd, "E", LR("حالة الاعتماد"), 300)
wd.conditional_formatting.add("E2:E300", CellIsRule(operator="equal", formula=['"قيد المراجعة"'], fill=PatternFill("solid", bgColor="FFF2CC")))
# ---- Custody
wc = sheet("العهد", ["الرقم الوظيفي", "الاسم", "نوع العهدة", "الكمية", "تاريخ التسليم", "تاريخ الاستلام", "الحالة", "ملاحظات"], [14, 22, 16, 10, 14, 14, 14, 24])
rows = []
for e in emps[:14]:
    for t in random.sample(lists["نوع العهدة"], 2):
        rows.append([e[0], None, t, 1, dt.date(2026, random.randint(6, 10), random.randint(1, 28)), None, "مع الموظف", ""])
rows[3][6] = "تالفة"; rows[3][7] = "تحتاج استبدال"; rows[5][5] = dt.date(2026, 11, 10); rows[5][6] = "تم الاستلام"
put(wc, rows)
for r in range(2, 401): wc.cell(r, 2, lookup(r)).border = B
empdv(wc, n=400); dv(wc, "C", LR("نوع العهدة"), 400); dv(wc, "G", LR("حالة العهدة"), 400)
for st, col in [("تالفة", "FFF2CC"), ("مفقودة", "F8CBAD")]:
    wc.conditional_formatting.add("G2:G400", CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", bgColor=col)))
# ---- Tasks
wt = sheet("المهام", ["رقم المهمة", "المهمة / الفعالية", "الموقع", "الجهة / العميل", "التاريخ", "الموظف المكلف", "الاسم", "الحالة"], [12, 28, 18, 20, 14, 14, 22, 14])
tk = ["تأمين فعالية معرض", "حراسة مجمع تجاري", "نقل شحنة", "دورية ليلية", "تأمين حفل"]
put(wt, [[f"T{i+1:03d}", random.choice(tk), random.choice(lists["الموقع"]), random.choice(["شركة النور", "مجمع الأمل", "مصنع الدلتا"]), dt.date(2026, 11, random.randint(1, 28)), emps[i][0], None, random.choice(lists["حالة المهمة"])] for i in range(12)])
for r in range(2, 301): wt.cell(r, 7, f'=IFERROR(VLOOKUP(F{r},الموظفون!$A$2:$B$300,2,FALSE),"")').border = B
empdv(wt, "F", 300); dv(wt, "H", LR("حالة المهمة"), 300)
# ---- Discipline (protected)
wp = sheet("الانضباط", ["الرقم الوظيفي", "الاسم", "التاريخ", "الواقعة", "نوع الإجراء", "الجهة المعتمدة"], [14, 22, 14, 34, 16, 18])
put(wp, [[emps[2][0], None, dt.date(2026, 10, 12), "غياب بدون إذن", "إنذار شفهي", "الموارد البشرية"], [emps[7][0], None, dt.date(2026, 11, 3), "تأخر متكرر", "إنذار كتابي", "مدير العمليات"]])
for r in range(2, 201): wp.cell(r, 2, lookup(r)).border = B
empdv(wp, n=200); dv(wp, "E", LR("نوع الإجراء"), 200)
# ---- Emergency (protected)
wm = sheet("الطوارئ", ["الرقم الوظيفي", "الاسم", "اسم جهة الاتصال", "الصلة", "هاتف الطوارئ"], [14, 22, 22, 14, 18])
put(wm, [[e[0], None, "(بيانات تجريبية)", random.choice(["أب", "أخ", "زوجة", "أم"]), f"01{random.choice([0,1,2])}{random.randint(10000000,99999999)}"] for e in emps])
for r in range(2, 301): wm.cell(r, 2, lookup(r)).border = B
empdv(wm, n=300)
# ---- Reports
wr = sheet("التقارير", ["الرقم الوظيفي", "الاسم", "أيام الحضور", "أيام الغياب", "مرات التأخير", "إجمالي ساعات التأخير", "عدد العهد المسلمة", "عهد غير مستلمة تالفة/مفقودة", "إقرارات قيد المراجعة"], [14, 22, 12, 12, 12, 16, 14, 20, 16])
for i in range(2, 22):
    wr.cell(i, 1, f"=IF(الموظفون!A{i}=\"\",\"\",الموظفون!A{i})"); wr.cell(i, 2, f'=IFERROR(VLOOKUP(A{i},الموظفون!$A$2:$B$300,2,FALSE),"")')
    wr.cell(i, 3, f'=COUNTIFS(الحضور!$A:$A,$A{i},الحضور!$F:$F,"حاضر")+COUNTIFS(الحضور!$A:$A,$A{i},الحضور!$F:$F,"متأخر")')
    wr.cell(i, 4, f'=COUNTIFS(الحضور!$A:$A,$A{i},الحضور!$F:$F,"غائب")'); wr.cell(i, 5, f'=COUNTIFS(الحضور!$A:$A,$A{i},الحضور!$F:$F,"متأخر")')
    wr.cell(i, 6, f'=SUMIFS(الحضور!$G:$G,الحضور!$A:$A,$A{i})'); wr.cell(i, 7, f'=COUNTIFS(العهد!$A:$A,$A{i},العهد!$G:$G,"مع الموظف")')
    wr.cell(i, 8, f'=COUNTIFS(العهد!$A:$A,$A{i},العهد!$G:$G,"تالفة")+COUNTIFS(العهد!$A:$A,$A{i},العهد!$G:$G,"مفقودة")'); wr.cell(i, 9, f'=COUNTIFS(الإقرارات!$A:$A,$A{i},الإقرارات!$E:$E,"قيد المراجعة")')
    for c in range(1, 10): wr.cell(i, c).border = B; wr.cell(i, c).alignment = Alignment(horizontal="center")
wr.cell(23, 2, "الإجمالي").font = Font(bold=True)
for c in range(3, 10): wr.cell(23, c, f"=SUM({L(c)}2:{L(c)}22)").font = Font(bold=True)
wr.conditional_formatting.add("D2:D22", CellIsRule(operator="greaterThanOrEqual", formula=["2"], fill=PatternFill("solid", bgColor="F8CBAD")))
wr.conditional_formatting.add("H2:H22", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", bgColor="FFF2CC")))
wr.page_setup.orientation = "landscape"; wr.page_setup.fitToWidth = 1; wr.sheet_properties.pageSetUpPr.fitToPage = True; wr.page_setup.fitToHeight = 0
for w in (wa, wc): w.page_setup.orientation = "landscape"
# ---- Instructions (first)
wi = wb.create_sheet("تعليمات", 0); wi.sheet_view.rightToLeft = True; wi.column_dimensions["A"].width = 100
txt = ["نموذج ملف Excel لإدارة شؤون الموظفين والعمليات الأمنية (نسخة تجريبية)",
 "كل البيانات في هذا الملف بيانات وهمية للتجربة فقط. احذف الصفوف التجريبية وأدخل بياناتكم.", "",
 "طريقة الاستخدام:",
 "1) ابدأ بورقة الموظفين: أدخل كل موظف مرة واحدة برقم وظيفي فريد. باقي الأوراق تقرأ الاسم تلقائيا من الرقم الوظيفي.",
 "2) الأعمدة التي فيها قائمة منسدلة (المسمى، الحالة، الموقع...) تقبل الاختيار من القائمة فقط. لتعديل القوائم عدل ورقة القوائم.",
 "3) الحضور: أدخل التاريخ ووقت الحضور والانصراف والحالة. ساعات التأخير تحسب تلقائيا بالنسبة لبداية الدوام 08:00.",
 "4) الإقرارات والعهد والمهام: سجل صفا لكل إقرار أو عهدة أو مهمة. الخلايا الصفراء والحمراء تنبه للحالات التي تحتاج متابعة.",
 "5) التقارير: ملخص لكل موظف (حضور، غياب، تأخير، عهد، إقرارات قيد المراجعة) جاهز للطباعة والتصفية. أضف الموظفين الجدد بنسخ الصيغ لأسفل.",
 "6) ورقتا الانضباط والطوارئ محميتان للتعديل بكلمة مرور (الكلمة في هذا النموذج: 1234، غيرها قبل الاستخدام الفعلي). مراجعة > إلغاء حماية الورقة.", "",
 "الأوراق: الموظفون، الحضور، الإقرارات، العهد، المهام، الانضباط، الطوارئ، التقارير، القوائم."]
for i, t in enumerate(txt, 1):
    c = wi.cell(i, 1, t); c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="right"); c.font = Font(name="Arial", size=14 if i == 1 else 11, bold=i in (1, 4))
for n in ("الانضباط", "الطوارئ"):
    ws = wb[n]; ws.protection.sheet = True; ws.protection.password = "1234"
    for row in ws.iter_rows(min_row=2, max_row=200):
        for c in row: c.protection = Protection(locked=False) if False else Protection(locked=True)
wb.save("HR_Security_Company_Sample_Workbook.xlsx")
