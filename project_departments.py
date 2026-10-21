import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

data = [
    ("C19602",  "ID - ADA CTY VVS T",         "Ada County",                    "Proactive.USPS.Del.SLED.Northstar"),
    ("C10576",  "IL-T-Mig Support",             "IL Toll Highway Auth",          "Proactive.USPS.Del.SLED.Northstar"),
    ("C16507",  "MN Dept of Corr VVS",          "MN Dept of Corrections",        "Proactive.USPS.Del.SLED.Northstar"),
    ("E55085",  "TX-T-COH-Next Gen",            "City of Houston",               "Proactive.USPS.Del.SLED.Central"),
    ("F18212",  "IL-S-Univ Chic SAS 2Yr",       "University of Chicago",         "Proactive.USPS.Del.SLED.Northstar"),
    ("F22532",  "IL-T-Tollway Mig SME",         "IL Toll Highway Auth",          "Proactive.USPS.Del.SLED.Northstar"),
    ("F27174",  "CAP01873418 JPS Health UC",    "JPS Health",                    "Proactive.USPS.Del.SLED.Central"),
    ("F29426",  "FL-JSO VVS SAS",              "Jacksonville Sheriff's Office", "Proactive.USPS.Del.SLED.SE"),
    ("F31176",  "NA-SAS-Neb Med Infra",         "Nebraska Medical Center",       "Proactive.USPS.Del.SLED.Northstar"),
    ("F31542",  "CityofVictorville-Webexcalling","City of Victorville",           "Proactive.USPS.Del.SLED.Northstar"),
    ("F33347",  "NYC-T-NYC Flex SME - 1 Year",  "City of New York",              "Proactive.USPS.Del.SLED.Northeast"),
    ("F33607",  "NYC-T-NYC Flex SME - 2 Years", "City of New York",              "Proactive.USPS.Del.SLED.Northeast"),
    ("F40696",  "NY-S NYS SENATE SAS",          "NYS Senate",                    "Proactive.USPS.Del.SLED.Northeast"),
    ("F43673",  "WI-Hennepin County CC",        "Hennepin County",               "CX Amer.Proactive.USPS.Del.PMO-SLED West"),
    ("F46100",  "WI-St of WI VVS SAS",          "State of Wisconsin",            "Proactive.USPS.Del.SLED.Northstar"),
    ("F47009",  "Green County VVS SAS",         "Greene County MO",              "Proactive.USPS.Del.SLED.Northstar"),
    ("F48852",  "MD-S-MDDHS ISE SAS",           "MD Dept Human Resources",       "CX.Customer Success.AMER.USPS - USA"),
    ("F48861",  "MD-S-MDDHS CatCntr SAS",       "MD Dept Human Resources",       "Proactive.USPS.Del.SLED.CE"),
    ("F51119",  "TX-S-COH LCS YR4",             "City of Houston",               "Proactive.USPS.Del.SLED.Central"),
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Project Departments"

header_fill = PatternFill("solid", fgColor="1565A0")
header_font = Font(bold=True, color="FFFFFF", name="Aptos")
even_fill = PatternFill("solid", fgColor="EBF2FA")
odd_fill = PatternFill("solid", fgColor="FFFFFF")
body_font = Font(name="Aptos", size=11)
thin = Side(style="thin", color="CCCCCC")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(vertical="center")

headers = ["PID", "Project Name", "Customer", "Department Name"]
ws.append(headers)
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border

for i, row in enumerate(data, start=2):
    fill = even_fill if i % 2 == 0 else odd_fill
    for col, val in enumerate(row, 1):
        cell = ws.cell(row=i, column=col, value=val)
        cell.fill = fill
        cell.font = body_font
        cell.alignment = center
        cell.border = border

ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 36
ws.column_dimensions["C"].width = 32
ws.column_dimensions["D"].width = 46
ws.row_dimensions[1].height = 20

out = os.path.join(os.path.expanduser("~"), "Desktop", "Project_Departments.xlsx")
wb.save(out)
print(f"Saved: {out}")
