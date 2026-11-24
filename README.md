# HR and Operations Workbook for a Security Company (Excel sample)

A ready-to-adapt Excel workbook for managing employees and security operations. **All names, numbers and records in the file are fictional sample data**, included only to show how it works.

## Sheets
| Sheet | Purpose |
|---|---|
| الموظفون (Employees) | One row per employee: ID, name, job title, phone, site, hire date, status |
| الحضور (Attendance) | Daily check-in/out, status, late hours (calculated), reason |
| الإقرارات (Declarations) | Signed declarations, date, subject, approval status, where the paper is filed |
| العهد (Custody) | Uniforms, radios, devices: delivery and return dates, condition |
| المهام (Tasks) | Tasks and events: site, client, date, assigned employee, status |
| الانضباط (Discipline) | Incidents and actions (sheet protected) |
| الطوارئ (Emergency) | Emergency contacts (sheet protected) |
| التقارير (Reports) | Per-employee summary: attendance, absences, late hours, custody, pending declarations. Print-ready |
| القوائم (Lists) | Drop-down lists used across the workbook |

## Design
- Employee ID is the key. Other sheets pick it from a drop-down and fetch the name automatically.
- Data validation on status, job, site and custody fields, so entries stay consistent.
- Colour alerts for absences, delays, damaged or lost custody and declarations under review.
- Right-to-left layout, frozen headers, landscape print setup on reports.
- Discipline and Emergency sheets are protected (sample password `1234`; change it before real use).

## Files
- `HR_Security_Company_Sample_Workbook.xlsx` the workbook
- `build.py` the Python (openpyxl) script that generates it
- `report_preview.png` preview of the Reports sheet

Language: Arabic UI. Free to use and adapt (MIT).
