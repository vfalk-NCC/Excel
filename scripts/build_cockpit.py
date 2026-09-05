"""
Bygger NSV PRODUCTION COCKPIT ovanpå det befintliga produktionsunderlaget.

Källbladen (741/742/744/745/775, LISTOR, AKT-SAMT, DP1, RESURS*) rörs INTE -
de kopieras oförändrade in i den nya arbetsboken. Ovanpå dem läggs:

  6-WEEK PLAN        - konsoliderad, strukturerad datatabell (tblPlan)
  PRODUCTION STATUS   - faktisk status per aktivitet (tblStatus)
  RISKS & BLOCKERS    - hinderregister (tblRisks)
  MATERIAL            - materialleveranser (tblMaterial)
  DOCUMENTS           - handlingar (tblDocuments)
  _ENGINE             - dold beräkningsmotor (rangordning/filter)
  _HISTORIK           - veckovis trendlogg (tblHistorik)
  PRODUCTION COCKPIT  - huvuddashboard
  AI ASSISTANT        - förberedd för AI/API-integration

Körs med: python3 scripts/build_cockpit.py
"""
import json
import datetime
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

SRC = "source/NSV_Bygg_4veckorsplanering_ORIGINAL.xlsm"
OUT = "NSV_Production_Cockpit.xlsm"
FONT = "Arial"

# ---------------------------------------------------------------- palette --
DARK_BG = "0E1420"
PANEL_BG = "141B2A"
CARD_BG = "1B2436"
CARD_BG2 = "202B42"
HEADER_BAR = "0A0E16"
BORDER_DARK = "2A3448"

ACCENT = "3DA9FC"
TEXT_LIGHT = "EAF0F8"
TEXT_MUTED = "93A0B8"

GREEN = "2ECC71"; GREEN_BG = "153B27"
AMBER = "F5A623"; AMBER_BG = "4A3A12"
RED = "E5493C"; RED_BG = "4A1E19"
GREY = "9AA7BC"; GREY_BG = "26314A"

LIGHT_BG = "F4F6FA"
LIGHT_PANEL = "FFFFFF"
LIGHT_HEADER = "13233B"
LIGHT_BORDER = "D6DCE5"
LIGHT_TEXT = "1B2436"

STATUS_KLAR = "✅ Klar"
STATUS_PLAN = "\U0001F7E2 Enligt plan"
STATUS_RISK = "\U0001F7E1 Risk"
STATUS_FORSENAD = "\U0001F534 Försenad"
STATUS_EJ = "⚪ Ej startad"

AREA_SHEETS = [
    ("741 - SEKTIONSFICKOR", "SEK"),
    ("742 - SIKTHALL", "SIK"),
    ("744 - Fläkthus", "FLA"),
    ("745 - FÖRTJOCKARHUS", "FOR"),
    ("775 - VATTENRESERVOAR", "VAT"),
]

with open("scripts/activities.json", encoding="utf-8") as f:
    ACTIVITIES = json.load(f)

N_ACT = len(ACTIVITIES)
PLAN_FIRST_ROW = 3          # first data row in 6-WEEK PLAN / PRODUCTION STATUS
PLAN_LAST_ROW = PLAN_FIRST_ROW + N_ACT - 1

# ================================================================== utils ==

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def style_header_cell(cell, bg=HEADER_BAR, fg=TEXT_LIGHT, size=10, bold=True, align="left"):
    cell.font = Font(name=FONT, bold=bold, color=fg, size=size)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)


def banner(ws, cell_range, text, bg=HEADER_BAR, fg=TEXT_LIGHT, size=16, height=28):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    c = ws[top_left]
    c.value = text
    c.font = Font(name=FONT, bold=True, size=size, color=fg)
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    row = int("".join(ch for ch in top_left if ch.isdigit()))
    ws.row_dimensions[row].height = height


THIN = Side(style="thin", color=BORDER_DARK)
THIN_LIGHT = Side(style="thin", color=LIGHT_BORDER)


def all_border(color=BORDER_DARK):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def add_table(ws, ref, name, style="TableStyleMedium2"):
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tbl)
    return tbl


print(f"Laddar källfil ({N_ACT} aktiviteter extraherade)...")
wb = openpyxl.load_workbook(SRC, data_only=False, keep_vba=True)
print("Klart.")

# Namngivna intervall för datavalidering i hela arbetsboken
def add_name(name, ref):
    wb.defined_names[name] = openpyxl.workbook.defined_name.DefinedName(name, attr_text=ref)

# ============================================================== LISTOR ====
# Utökar befintliga listor utan att röra kolumn A-E.
listor = wb["LISTOR"]
listor["F4"] = "Risktyp"
for i, v in enumerate(["Handling", "Material", "Produktion", "Projektering",
                        "Beslut", "UE", "Myndighet", "Logistik", "Annat"], start=5):
    listor.cell(row=i, column=6, value=v)

listor["G4"] = "Blockerstatus"
for i, v in enumerate(["Öppen", "Pågående", "Löst"], start=5):
    listor.cell(row=i, column=7, value=v)

listor["H4"] = "Materialstatus"
for i, v in enumerate(["Beställt", "Bekräftat", "Levererat", "Försenat"], start=5):
    listor.cell(row=i, column=8, value=v)

listor["I4"] = "Dokumentstatus"
for i, v in enumerate(["Beställd", "Under arbete", "Levererad", "Försenad"], start=5):
    listor.cell(row=i, column=9, value=v)

# Statuslistan (kolumn D) kompletteras med "Försenad" för statusflödet i cockpiten.
if listor["D9"].value is None:
    listor["D9"] = "Försenad"

listor["J4"] = "Disciplin (filter)"
listor["J5"] = "(Alla)"
disc_vals = [listor.cell(row=i, column=2).value for i in range(5, 13) if listor.cell(row=i, column=2).value]
for i, v in enumerate(disc_vals, start=6):
    listor.cell(row=i, column=10, value=v)

listor["K4"] = "Ansvarig/UE (filter)"
listor["K5"] = "(Alla)"
ansv_vals = [listor.cell(row=i, column=3).value for i in range(5, 9) if listor.cell(row=i, column=3).value]
for i, v in enumerate(ansv_vals, start=6):
    listor.cell(row=i, column=11, value=v)

add_name("List_Disciplin_F", f"LISTOR!$J$5:$J${5+len(disc_vals)}")
add_name("List_Ansvarig_F", f"LISTOR!$K$5:$K${5+len(ansv_vals)}")

for col, hdr_cell in (("F", "F4"), ("G", "G4"), ("H", "H4"), ("I", "I4"), ("J", "J4"), ("K", "K4")):
    listor[hdr_cell].font = Font(name=FONT, bold=True)
    listor.column_dimensions[col].width = 18

add_name("List_Disciplin", "LISTOR!$B$5:$B$12")
add_name("List_Ansvarig", "LISTOR!$C$5:$C$8")
add_name("List_Status5", "LISTOR!$D$5:$D$9")
add_name("List_Entreprenad", "LISTOR!$E$5:$E$8")
add_name("List_Risktyp", "LISTOR!$F$5:$F$13")
add_name("List_Blockerstatus", "LISTOR!$G$5:$G$7")
add_name("List_Materialstatus", "LISTOR!$H$5:$H$8")
add_name("List_Dokumentstatus", "LISTOR!$I$5:$I$8")
add_name("List_AktivitetID", f"'6-WEEK PLAN'!$A${PLAN_FIRST_ROW}:$A${PLAN_LAST_ROW}")
print("LISTOR utökad.")

# ========================================================= 6-WEEK PLAN ====
ws = wb.create_sheet("6-WEEK PLAN")
ws.sheet_view.showGridLines = False
banner(ws, "A1:P1", "6-WEEK PLAN  —  Konsoliderad aktivitetsdata (primär datakälla för cockpiten)",
       bg=LIGHT_HEADER, fg="FFFFFF", size=14)

PLAN_HEADERS = [
    "ID", "Område", "Del/Linje", "Aktivitet", "Typ", "DP", "Disciplin", "ÄTA",
    "Planerad start", "Antal dagar", "Planerad slut", "Vecka",
    "Framdrift % (planerad vikt)", "Helgarbete", "Ansvarig/UE", "Kommentar",
]
hdr_row = 2
for i, h in enumerate(PLAN_HEADERS, start=1):
    c = ws.cell(row=hdr_row, column=i, value=h)
    style_header_cell(c, bg=LIGHT_HEADER, fg="FFFFFF")

for offset, act in enumerate(ACTIVITIES):
    r = PLAN_FIRST_ROW + offset
    ws.cell(row=r, column=1, value=act["id"])
    ws.cell(row=r, column=2, value=act["omrade"])
    ws.cell(row=r, column=3, value=act["del_linje"])
    ws.cell(row=r, column=4, value=act["aktivitet"])
    ws.cell(row=r, column=5, value=act["typ"])
    ws.cell(row=r, column=6, value=act["dp"])
    ws.cell(row=r, column=7, value=act["disciplin"])
    ws.cell(row=r, column=8, value=act["ata"])
    ws.cell(row=r, column=9, value=datetime.datetime.strptime(act["datum_start"], "%Y-%m-%d"))
    ws.cell(row=r, column=10, value=act["antal_dagar"])
    ws.cell(row=r, column=11, value=datetime.datetime.strptime(act["datum_slut"], "%Y-%m-%d"))
    ws.cell(row=r, column=12, value=f"=IFERROR(WEEKNUM(I{r},21),\"\")")
    ws.cell(row=r, column=13, value=act["framdrift_kalla"])
    ws.cell(row=r, column=14, value=act["helgarbete"])
    ws.cell(row=r, column=15, value=None)  # Ansvarig/UE - fylls i av användaren
    ws.cell(row=r, column=16, value=None)  # Kommentar - fylls i av användaren

    for col in (9, 11):
        ws.cell(row=r, column=col).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=13).number_format = "0%"
    ws.cell(row=r, column=15).fill = PatternFill("solid", fgColor="FFF9DB")
    ws.cell(row=r, column=16).fill = PatternFill("solid", fgColor="FFF9DB")

add_table(ws, f"A{hdr_row}:P{PLAN_LAST_ROW}", "tblPlan", style="TableStyleMedium2")

set_col_widths(ws, {
    "A": 10, "B": 18, "C": 22, "D": 46, "E": 10, "F": 7, "G": 12, "H": 8,
    "I": 13, "J": 10, "K": 13, "L": 7, "M": 12, "N": 10, "O": 18, "P": 28,
})
ws.freeze_panes = "D3"

dv_disc = DataValidation(type="list", formula1="=List_Disciplin", allow_blank=True)
dv_ansv = DataValidation(type="list", formula1="=List_Ansvarig", allow_blank=True, showErrorMessage=False)
ws.add_data_validation(dv_disc); dv_disc.add(f"G{PLAN_FIRST_ROW}:G{PLAN_LAST_ROW}")
ws.add_data_validation(dv_ansv); dv_ansv.add(f"O{PLAN_FIRST_ROW}:O{PLAN_LAST_ROW}")

note = ws.cell(row=PLAN_LAST_ROW + 2, column=1,
    value=("Källa: automatiskt extraherad från de fem områdesbladen (741/742/744/745/775) vid byggtillfället. "
           "Gula fält (Ansvarig/UE, Kommentar) fylls i här - inga dubbla inmatningar. "
           "Lägg till nya rader längst ner i tabellen (tabellen expanderar automatiskt) när nya aktiviteter tillkommer i planen."))
note.font = Font(name=FONT, italic=True, size=9, color="666666")
ws.row_dimensions[PLAN_LAST_ROW + 2].height = 28
print("6-WEEK PLAN klart.")

# Statusfärger återanvänder Excels inbyggda Bra/Neutral/Dålig-toner,
# samma som redan används i de befintliga områdesbladen (Klar/Pågående/Risk).
STATUS_COLORS = {
    "Klar":       ("C6EFCE", "006100"),
    "Pågående":   ("FFEB9C", "9C6500"),
    "Risk":       ("FFEB9C", "9C6500"),
    "Försenad":   ("FFC7CE", "9C0006"),
    "Ej startad": ("F2F2F2", "7F7F7F"),
}

# ====================================================== PRODUCTION STATUS ==
ws = wb.create_sheet("PRODUCTION STATUS")
ws.sheet_view.showGridLines = False
banner(ws, "A1:O1", "PRODUCTION STATUS  —  Faktisk status per aktivitet (uppdateras löpande av planerare/platsledning)",
       bg=LIGHT_HEADER, fg="FFFFFF", size=14)

STATUS_HEADERS = [
    "ID", "Aktivitet", "Område", "Ansvarig/UE", "Planerad start", "Planerad slut",
    "% planerad idag", "Faktisk start", "Faktiskt slut", "% faktisk",
    "Avvikelse %", "Status", "Kritikpoäng", "Kommentar", "Uppdaterad",
]
hdr_row = 2
for i, h in enumerate(STATUS_HEADERS, start=1):
    c = ws.cell(row=hdr_row, column=i, value=h)
    style_header_cell(c, bg=LIGHT_HEADER, fg="FFFFFF")

for offset, act in enumerate(ACTIVITIES):
    r = PLAN_FIRST_ROW + offset
    ws.cell(row=r, column=1, value=f"='6-WEEK PLAN'!A{r}")
    ws.cell(row=r, column=2, value=f"='6-WEEK PLAN'!D{r}")
    ws.cell(row=r, column=3, value=f"='6-WEEK PLAN'!B{r}")
    ws.cell(row=r, column=4, value=f"='6-WEEK PLAN'!O{r}")
    ws.cell(row=r, column=5, value=f"='6-WEEK PLAN'!I{r}")
    ws.cell(row=r, column=6, value=f"='6-WEEK PLAN'!K{r}")
    ws.cell(row=r, column=7,
        value=f"=IFERROR(MAX(0,MIN(1,(TODAY()-E{r}+1)/(F{r}-E{r}+1))),0)")
    ws.cell(row=r, column=8, value=None)   # Faktisk start - fylls i
    ws.cell(row=r, column=9, value=None)   # Faktiskt slut - fylls i
    ws.cell(row=r, column=10, value=None)  # % faktisk - fylls i
    ws.cell(row=r, column=11, value=f'=IF(J{r}="","",J{r}-G{r})')
    ws.cell(row=r, column=12, value=(
        f'=IF(I{r}<>"","Klar",'
        f'IF(F{r}<TODAY(),"Försenad",'
        f'IF(AND(E{r}<=TODAY()+7,E{r}>=TODAY(),OR(_ENGINE!B{r}=1,'
        f'COUNTIFS(tblRisks[Aktivitet-ID],A{r},tblRisks[Status],"<>Löst")>0,'
        f'COUNTIFS(tblMaterial[Aktivitet-ID],A{r},tblMaterial[Risk],"Ja")>0,'
        f'COUNTIFS(tblDocuments[Aktivitet-ID],A{r},tblDocuments[Risk],"Ja")>0)),"Risk",'
        f'IF(H{r}<>"","Pågående",'
        f'IF(E{r}>TODAY(),"Ej startad","Pågående")))))'
    ))
    ws.cell(row=r, column=13, value=(
        f'=IF(L{r}="Klar",0,'
        f'(IF(L{r}="Försenad",3,0))'
        f'+(IF(L{r}="Risk",2,0))'
        f'+(IF(_ENGINE!B{r}=1,1,0))'
        f'+COUNTIFS(tblRisks[Aktivitet-ID],A{r},tblRisks[Status],"<>Löst")*2'
        f'+COUNTIFS(tblMaterial[Aktivitet-ID],A{r},tblMaterial[Risk],"Ja")*2'
        f'+COUNTIFS(tblDocuments[Aktivitet-ID],A{r},tblDocuments[Risk],"Ja")*2)'
    ))
    ws.cell(row=r, column=14, value=None)  # Kommentar
    ws.cell(row=r, column=15, value=None)  # Uppdaterad

    for col in (5, 6, 8, 9, 15):
        ws.cell(row=r, column=col).number_format = "yyyy-mm-dd"
    for col in (7, 10, 11):
        ws.cell(row=r, column=col).number_format = "0%"
    for col in (8, 9, 10, 14, 15):
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor="FFF9DB")

add_table(ws, f"A{hdr_row}:O{PLAN_LAST_ROW}", "tblStatus", style="TableStyleMedium2")

for status_val, (bg, fg) in STATUS_COLORS.items():
    ws.conditional_formatting.add(
        f"L{PLAN_FIRST_ROW}:L{PLAN_LAST_ROW}",
        CellIsRule(operator="equal", formula=[f'"{status_val}"'],
                   fill=PatternFill("solid", fgColor=bg), font=Font(color=fg, bold=True)))

ws.conditional_formatting.add(
    f"M{PLAN_FIRST_ROW}:M{PLAN_LAST_ROW}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=10,
                color="E5493C", showValue=True, minLength=None, maxLength=None))

set_col_widths(ws, {
    "A": 10, "B": 46, "C": 18, "D": 18, "E": 13, "F": 13, "G": 12, "H": 13,
    "I": 13, "J": 10, "K": 11, "L": 12, "M": 11, "N": 28, "O": 12,
})
ws.freeze_panes = "B3"
print("PRODUCTION STATUS klart.")

# Exempel-ankare: riktiga aktiviteter ur planen används för att visa hur
# registren hänger ihop med 6-WEEK PLAN. Tydligt märkta som exempel i UI:t.
anchor_idx = [10, 45, 90, 150, 220, 300, 380, 450]
anchors = [ACTIVITIES[i] for i in anchor_idx if i < len(ACTIVITIES)]

def fillable(ws, cell, val=None):
    c = ws[cell]
    if val is not None:
        c.value = val
    c.fill = PatternFill("solid", fgColor="FFF9DB")
    return c

# ===================================================== RISKS & BLOCKERS ====
ws = wb.create_sheet("RISKS & BLOCKERS")
ws.sheet_view.showGridLines = False
banner(ws, "A1:M1", "RISKS & BLOCKERS  —  Register över produktionshinder",
       bg=LIGHT_HEADER, fg="FFFFFF", size=14)

RISK_HEADERS = ["ID", "Datum", "Aktivitet-ID", "Aktivitet", "Problem", "Typ",
                "Ansvarig", "Förfallodatum", "Status", "Konsekvens", "Åtgärd",
                "Kommentar", "Försenad åtgärd"]
hdr_row = 2
for i, h in enumerate(RISK_HEADERS, start=1):
    c = ws.cell(row=hdr_row, column=i, value=h)
    style_header_cell(c, bg=LIGHT_HEADER, fg="FFFFFF")

RISK_EXAMPLES = [
    ("Ritning saknas för fundament", "Handling", "Platschef mark (NCC)", 6, "Öppen",
     "Kan inte starta formsättning i tid", "Efterlyst hos projektering 2026-08-20"),
    ("Material (kamstål) ej bekräftat av leverantör", "Material", "Platschef (Procyon)", 3, "Pågående",
     "Risk för stopp i bergförankring", "Uppföljningssamtal bokat"),
    ("Föregående aktivitet i linjen ej avslutad", "Produktion", "Montageledare", 2, "Öppen",
     "Startdatum kan behöva skjutas fram", "Avstämning med UE pågår"),
    ("Beslut om ÄTA-hantering saknas", "Beslut", "Platschef (Procyon)", 5, "Öppen",
     "Osäkert vilket omfång som ska utföras", "Eskalerat till projektledning"),
    ("UE har inte återkopplat bemanningsplan", "UE", "Installationsledare", 4, "Pågående",
     "Risk för underbemanning vecka 38-39", "Påminnelse skickad"),
]
for i, (problem, typ, ansv, due_offset, status, konsekvens, atgard) in enumerate(RISK_EXAMPLES):
    r = 3 + i
    act = anchors[i % len(anchors)]
    ws.cell(row=r, column=1, value=f'="RB-"&TEXT(ROW()-2,"000")')
    fillable(ws, f"B{r}", datetime.datetime(2026, 8, 20 + i))
    fillable(ws, f"C{r}", act["id"])
    ws.cell(row=r, column=4, value=f'=IFERROR(INDEX(tblPlan[Aktivitet],MATCH(C{r},tblPlan[ID],0)),"")')
    fillable(ws, f"E{r}", problem)
    fillable(ws, f"F{r}", typ)
    fillable(ws, f"G{r}", ansv)
    fillable(ws, f"H{r}", datetime.datetime(2026, 9, 5 + due_offset))
    fillable(ws, f"I{r}", status)
    fillable(ws, f"J{r}", konsekvens)
    fillable(ws, f"K{r}", atgard)
    fillable(ws, f"L{r}", "EXEMPEL - ersätt eller radera")
    ws.cell(row=r, column=13, value=f'=IF(AND(I{r}<>"Löst",H{r}<>"",H{r}<TODAY()),"Ja","Nej")')
    ws.cell(row=r, column=2).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=8).number_format = "yyyy-mm-dd"

RISK_LAST_ROW = 2 + len(RISK_EXAMPLES)
ws.cell(row=hdr_row, column=14, value="Sortordning")
for i in range(len(RISK_EXAMPLES)):
    r = 3 + i
    ws.cell(row=r, column=14, value=f'=IF(I{r}<>"Löst",H{r}*1000+ROW(),9999999999+ROW())')
add_table(ws, f"A{hdr_row}:N{RISK_LAST_ROW}", "tblRisks", style="TableStyleMedium3")

for status_val, (bg, fg) in {"Öppen": ("FFC7CE", "9C0006"), "Pågående": ("FFEB9C", "9C6500"),
                              "Löst": ("C6EFCE", "006100")}.items():
    ws.conditional_formatting.add(
        f"I3:I{RISK_LAST_ROW}",
        CellIsRule(operator="equal", formula=[f'"{status_val}"'],
                   fill=PatternFill("solid", fgColor=bg), font=Font(color=fg, bold=True)))

dv = DataValidation(type="list", formula1="=List_AktivitetID", allow_blank=True)
ws.add_data_validation(dv); dv.add(f"C3:C{RISK_LAST_ROW+200}")
dv2 = DataValidation(type="list", formula1="=List_Risktyp", allow_blank=True)
ws.add_data_validation(dv2); dv2.add(f"F3:F{RISK_LAST_ROW+200}")
dv3 = DataValidation(type="list", formula1="=List_Blockerstatus", allow_blank=True)
ws.add_data_validation(dv3); dv3.add(f"I3:I{RISK_LAST_ROW+200}")
dv4 = DataValidation(type="list", formula1="=List_Ansvarig", allow_blank=True, showErrorMessage=False)
ws.add_data_validation(dv4); dv4.add(f"G3:G{RISK_LAST_ROW+200}")

set_col_widths(ws, {"A": 10, "B": 12, "C": 12, "D": 40, "E": 34, "F": 13,
                     "G": 20, "H": 13, "I": 11, "J": 30, "K": 30, "L": 34, "M": 12, "N": 10})
ws.column_dimensions["N"].hidden = True
ws.freeze_panes = "E3"
print("RISKS & BLOCKERS klart.")

# ================================================================ MATERIAL ==
ws = wb.create_sheet("MATERIAL")
ws.sheet_view.showGridLines = False
banner(ws, "A1:M1", "MATERIAL  —  Materialleveranser kopplade till produktionsplanen",
       bg=LIGHT_HEADER, fg="FFFFFF", size=14)

MAT_HEADERS = ["ID", "Material", "Aktivitet-ID", "Aktivitet", "Produktionsstart (aktivitet)",
               "Leverantör", "Planerad leverans", "Faktisk leverans", "Status",
               "Ansvarig", "Kommentar", "Risk", "Riskbeskrivning"]
hdr_row = 2
for i, h in enumerate(MAT_HEADERS, start=1):
    c = ws.cell(row=hdr_row, column=i, value=h)
    style_header_cell(c, bg=LIGHT_HEADER, fg="FFFFFF")

MAT_EXAMPLES = [
    ("Kamstål φ25 bergförankring", "Ståldepån AB", -2, 0, "Beställt", "Platschef (Procyon)"),
    ("Formvirke grundsula", "Byggmax Projekt", 3, 3, "Bekräftat", "Platschef mark (NCC)"),
    ("Prefab fundamentselement M-linje", "Betongelement Nord", 10, None, "Försenat", "Montageledare"),
    ("Ventilationsaggregat Fläkthus", "Klimatpartner", 25, None, "Beställt", "Installationsledare"),
]
for i, (material, lev, delivery_offset, actual_offset, status, ansv) in enumerate(MAT_EXAMPLES):
    r = 3 + i
    act = anchors[(i + 2) % len(anchors)]
    ws.cell(row=r, column=1, value=f'="MAT-"&TEXT(ROW()-2,"000")')
    fillable(ws, f"B{r}", material)
    fillable(ws, f"C{r}", act["id"])
    ws.cell(row=r, column=4, value=f'=IFERROR(INDEX(tblPlan[Aktivitet],MATCH(C{r},tblPlan[ID],0)),"")')
    ws.cell(row=r, column=5, value=f'=IFERROR(INDEX(tblPlan[Planerad start],MATCH(C{r},tblPlan[ID],0)),"")')
    fillable(ws, f"F{r}", lev)
    fillable(ws, f"G{r}", datetime.datetime(2026, 9, 5) + datetime.timedelta(days=delivery_offset))
    if actual_offset is not None:
        fillable(ws, f"H{r}", datetime.datetime(2026, 9, 5) + datetime.timedelta(days=actual_offset))
    else:
        fillable(ws, f"H{r}", None)
    fillable(ws, f"I{r}", status)
    fillable(ws, f"J{r}", ansv)
    fillable(ws, f"K{r}", "EXEMPEL - ersätt eller radera")
    ws.cell(row=r, column=12, value=(
        f'=IF(OR(AND(H{r}="",G{r}<>"",E{r}<>"",G{r}>E{r}),AND(H{r}<>"",E{r}<>"",H{r}>E{r})),"Ja","Nej")'))
    ws.cell(row=r, column=13, value=(
        f'=IF(L{r}="Ja","Risk: leverans riskerar komma efter produktionsstart ("&TEXT(E{r},"yyyy-mm-dd")&")","")'))
    ws.cell(row=r, column=5).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=7).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=8).number_format = "yyyy-mm-dd"

MAT_LAST_ROW = 2 + len(MAT_EXAMPLES)
ws.cell(row=hdr_row, column=14, value="Sortordning")
for i in range(len(MAT_EXAMPLES)):
    r = 3 + i
    ws.cell(row=r, column=14, value=(
        f'=IF(AND(G{r}<>"",G{r}<=TODAY()+30),G{r}*1000+ROW(),9999999999+ROW())'))
add_table(ws, f"A{hdr_row}:N{MAT_LAST_ROW}", "tblMaterial", style="TableStyleMedium3")

ws.conditional_formatting.add(f"L3:L{MAT_LAST_ROW}",
    CellIsRule(operator="equal", formula=['"Ja"'],
               fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
ws.conditional_formatting.add(f"L3:L{MAT_LAST_ROW}",
    CellIsRule(operator="equal", formula=['"Nej"'],
               fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100")))

dv = DataValidation(type="list", formula1="=List_AktivitetID", allow_blank=True)
ws.add_data_validation(dv); dv.add(f"C3:C{MAT_LAST_ROW+200}")
dv2 = DataValidation(type="list", formula1="=List_Materialstatus", allow_blank=True)
ws.add_data_validation(dv2); dv2.add(f"I3:I{MAT_LAST_ROW+200}")
dv3 = DataValidation(type="list", formula1="=List_Ansvarig", allow_blank=True, showErrorMessage=False)
ws.add_data_validation(dv3); dv3.add(f"J3:J{MAT_LAST_ROW+200}")

set_col_widths(ws, {"A": 10, "B": 30, "C": 12, "D": 40, "E": 15, "F": 20,
                     "G": 14, "H": 14, "I": 12, "J": 20, "K": 30, "L": 8, "M": 40, "N": 10})
ws.column_dimensions["N"].hidden = True
ws.freeze_panes = "D3"
print("MATERIAL klart.")

# ============================================================== DOCUMENTS ==
ws = wb.create_sheet("DOCUMENTS")
ws.sheet_view.showGridLines = False
banner(ws, "A1:N1", "DOCUMENTS  —  Handlingar/ritningar kopplade till produktionsplanen",
       bg=LIGHT_HEADER, fg="FFFFFF", size=14)

DOC_HEADERS = ["ID", "Handling", "Dokumentnummer", "Revision", "Aktivitet-ID", "Aktivitet",
               "Produktionsstart (aktivitet)", "Ansvarig", "Behövs senast",
               "Förväntad leverans", "Faktisk leverans", "Status", "Risk", "Riskbeskrivning"]
hdr_row = 2
for i, h in enumerate(DOC_HEADERS, start=1):
    c = ws.cell(row=hdr_row, column=i, value=h)
    style_header_cell(c, bg=LIGHT_HEADER, fg="FFFFFF")

DOC_EXAMPLES = [
    ("Formritning grundsula M-linje", "K-101-M30", "B", 5, "Under arbete", "Platschef mark (NCC)"),
    ("Armeringsritning fundament", "K-210-SIK", "A", -1, "Beställd", "Platschef (Procyon)"),
    ("Bygglov komplettering Fläkthus", "BL-744-02", "-", 20, "Beställd", "Montageledare"),
]
for i, (handling, docnr, rev, need_offset, status, ansv) in enumerate(DOC_EXAMPLES):
    r = 3 + i
    act = anchors[(i + 4) % len(anchors)]
    ws.cell(row=r, column=1, value=f'="DOC-"&TEXT(ROW()-2,"000")')
    fillable(ws, f"B{r}", handling)
    fillable(ws, f"C{r}", docnr)
    fillable(ws, f"D{r}", rev)
    fillable(ws, f"E{r}", act["id"])
    ws.cell(row=r, column=6, value=f'=IFERROR(INDEX(tblPlan[Aktivitet],MATCH(E{r},tblPlan[ID],0)),"")')
    ws.cell(row=r, column=7, value=f'=IFERROR(INDEX(tblPlan[Planerad start],MATCH(E{r},tblPlan[ID],0)),"")')
    fillable(ws, f"H{r}", ansv)
    fillable(ws, f"I{r}", datetime.datetime(2026, 9, 5) + datetime.timedelta(days=need_offset))
    fillable(ws, f"J{r}", datetime.datetime(2026, 9, 5) + datetime.timedelta(days=need_offset + 3))
    fillable(ws, f"K{r}", None)
    fillable(ws, f"L{r}", status)
    ws.cell(row=r, column=13, value=f'=IF(AND(K{r}="",I{r}<>"",I{r}<=TODAY()+7),"Ja","Nej")')
    ws.cell(row=r, column=14, value=(
        f'=IF(M{r}="Ja","Risk: handling krävs inom 7 dagar (senast "&TEXT(I{r},"yyyy-mm-dd")&") men saknas","")'))
    for col in (7, 9, 10, 11):
        ws.cell(row=r, column=col).number_format = "yyyy-mm-dd"

DOC_LAST_ROW = 2 + len(DOC_EXAMPLES)
ws.cell(row=hdr_row, column=15, value="Sortordning")
for i in range(len(DOC_EXAMPLES)):
    r = 3 + i
    ws.cell(row=r, column=15, value=(
        f'=IF(AND(I{r}<>"",I{r}<=TODAY()+30),I{r}*1000+ROW(),9999999999+ROW())'))
add_table(ws, f"A{hdr_row}:O{DOC_LAST_ROW}", "tblDocuments", style="TableStyleMedium3")

ws.conditional_formatting.add(f"M3:M{DOC_LAST_ROW}",
    CellIsRule(operator="equal", formula=['"Ja"'],
               fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
ws.conditional_formatting.add(f"M3:M{DOC_LAST_ROW}",
    CellIsRule(operator="equal", formula=['"Nej"'],
               fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100")))

dv = DataValidation(type="list", formula1="=List_AktivitetID", allow_blank=True)
ws.add_data_validation(dv); dv.add(f"E3:E{DOC_LAST_ROW+200}")
dv2 = DataValidation(type="list", formula1="=List_Dokumentstatus", allow_blank=True)
ws.add_data_validation(dv2); dv2.add(f"L3:L{DOC_LAST_ROW+200}")
dv3 = DataValidation(type="list", formula1="=List_Ansvarig", allow_blank=True, showErrorMessage=False)
ws.add_data_validation(dv3); dv3.add(f"H3:H{DOC_LAST_ROW+200}")

set_col_widths(ws, {"A": 10, "B": 30, "C": 15, "D": 9, "E": 12, "F": 38, "G": 15,
                     "H": 20, "I": 13, "J": 15, "K": 14, "L": 13, "M": 8, "N": 42, "O": 10})
ws.column_dimensions["O"].hidden = True
ws.freeze_panes = "F3"
print("DOCUMENTS klart.")

# ============================================================== filters ===
# Fasta celladresser för filterraden på PRODUCTION COCKPIT (definieras här
# eftersom _ENGINE-bladet måste referera dem). Se cockpit-sektionen nedan.
F_VECKA, F_OMRADE, F_DISC, F_TYP, F_ANSVARIG, F_STATUS = (
    "'PRODUCTION COCKPIT'!$C$4", "'PRODUCTION COCKPIT'!$E$4", "'PRODUCTION COCKPIT'!$G$4",
    "'PRODUCTION COCKPIT'!$I$4", "'PRODUCTION COCKPIT'!$K$4", "'PRODUCTION COCKPIT'!$M$4",
)

iso_weeks = sorted({datetime.datetime.strptime(a["datum_start"], "%Y-%m-%d").isocalendar()[1] for a in ACTIVITIES})
week_lo, week_hi = max(1, iso_weeks[0] - 1), min(53, iso_weeks[-1] + 1)
WEEK_LIST = "(Alla)," + ",".join(str(w) for w in range(week_lo, week_hi + 1))
TYP_LIST = "(Alla)," + ",".join(sorted({a["typ"] for a in ACTIVITIES if a["typ"]}))
OMRADE_LIST = "(Alla)," + ",".join(s for s, _ in AREA_SHEETS)
STATUS_LIST = "(Alla),Ej startad,Pågående,Risk,Försenad,Klar"

# ================================================================ _ENGINE ==
ws = wb.create_sheet("_ENGINE")
ws.sheet_state = "hidden"
ws["A1"] = "Beräkningsmotor - dold. Rör inte manuellt. Rad-för-rad samma ordning som 6-WEEK PLAN / PRODUCTION STATUS."
ENGINE_HEADERS = ["ID", "FöregåendeEjKlar", "StatusIkon", "PassesFilter", "Kritikpoäng",
                   "SortKritisk", "SortStart7", "SortStart30", "SortWindow6v", "SortMilstolpe30"]
for i, h in enumerate(ENGINE_HEADERS, start=1):
    ws.cell(row=2, column=i, value=h)

for offset in range(N_ACT):
    r = PLAN_FIRST_ROW + offset
    ws.cell(row=r, column=1, value=f"='PRODUCTION STATUS'!A{r}")
    ws.cell(row=r, column=2, value=(
        f"=--(SUMPRODUCT((tblPlan[Del/Linje]='6-WEEK PLAN'!C{r})*"
        f"(tblPlan[Område]='6-WEEK PLAN'!B{r})*"
        f"(tblPlan[Planerad start]<'6-WEEK PLAN'!I{r})*"
        f"(tblStatus[Faktiskt slut]=\"\"))>0)"
    ))
    ws.cell(row=r, column=3, value=(
        f"=IF('PRODUCTION STATUS'!L{r}=\"Klar\",\"✅ Klar\","
        f"IF('PRODUCTION STATUS'!L{r}=\"Försenad\",\"🔴 Försenad\","
        f"IF('PRODUCTION STATUS'!L{r}=\"Risk\",\"🟡 Risk\","
        f"IF('PRODUCTION STATUS'!L{r}=\"Pågående\",\"🟢 Pågående\",\"⚪ Ej startad\"))))"
    ))
    ws.cell(row=r, column=4, value=(
        f"=IF(AND("
        f"OR({F_VECKA}=\"(Alla)\",'6-WEEK PLAN'!N{r}={F_VECKA}),"
        f"OR({F_OMRADE}=\"(Alla)\",'6-WEEK PLAN'!B{r}={F_OMRADE}),"
        f"OR({F_DISC}=\"(Alla)\",'6-WEEK PLAN'!G{r}={F_DISC}),"
        f"OR({F_TYP}=\"(Alla)\",'6-WEEK PLAN'!E{r}={F_TYP}),"
        f"OR({F_ANSVARIG}=\"(Alla)\",'6-WEEK PLAN'!O{r}={F_ANSVARIG}),"
        f"OR({F_STATUS}=\"(Alla)\",'PRODUCTION STATUS'!L{r}={F_STATUS})"
        f"),1,0)"
    ))
    ws.cell(row=r, column=5, value=f"='PRODUCTION STATUS'!M{r}")
    ws.cell(row=r, column=6, value=f"=IF(D{r}=1,E{r}*100000-{r},-999999)")
    ws.cell(row=r, column=7, value=(
        f"=IF(AND(D{r}=1,'6-WEEK PLAN'!I{r}>=TODAY(),'6-WEEK PLAN'!I{r}<=TODAY()+7),"
        f"'6-WEEK PLAN'!I{r}*1000+{r},9999999999+{r})"
    ))
    ws.cell(row=r, column=8, value=(
        f"=IF(AND(D{r}=1,'6-WEEK PLAN'!I{r}>=TODAY(),'6-WEEK PLAN'!I{r}<=TODAY()+30),"
        f"'6-WEEK PLAN'!I{r}*1000+{r},9999999999+{r})"
    ))
    ws.cell(row=r, column=9, value=(
        f"=IF(AND(D{r}=1,'6-WEEK PLAN'!I{r}>=TODAY()-7,'6-WEEK PLAN'!I{r}<=TODAY()+42),"
        f"'6-WEEK PLAN'!I{r}*1000+{r},9999999999+{r})"
    ))
    ws.cell(row=r, column=10, value=(
        f"=IF(AND(D{r}=1,'6-WEEK PLAN'!J{r}=0,'6-WEEK PLAN'!I{r}>=TODAY(),'6-WEEK PLAN'!I{r}<=TODAY()+30),"
        f"'6-WEEK PLAN'!I{r}*1000+{r},9999999999+{r})"
    ))
    # Kolumn K beror på användarens val på AI ASSISTANT-bladet (fylls i formler
    # nedan efter att det bladet skapats, se AI_SEL_* celladresser).
    ws.cell(row=r, column=11, value=(
        f"=IF(AND('6-WEEK PLAN'!C{r}=AI_SEL_LINJE,'6-WEEK PLAN'!B{r}=AI_SEL_OMRADE,"
        f"'6-WEEK PLAN'!I{r}>AI_SEL_START),'6-WEEK PLAN'!I{r}*1000+{r},9999999999+{r})"
    ))

print("_ENGINE klart.")

# ========================================================= helper: ranking =
VAL_COL, IDX_COL = "V", "W"   # dolda scratch-kolumner på COCKPIT, återanvänds radvis
ENGINE_LAST = PLAN_LAST_ROW

def rank_ref(ws, row, k, sort_range, mode):
    """Skriver rankningsformler i dolda hjälpkolumner och returnerar cellreferensen
    med radnumret i 6-WEEK PLAN / PRODUCTION STATUS / _ENGINE (samma radnr överallt)."""
    if mode == "desc":
        ws[f"{VAL_COL}{row}"] = f"=IFERROR(LARGE({sort_range},{k}),-999999)"
        guard = f"{VAL_COL}{row}<=-999999"
    else:
        ws[f"{VAL_COL}{row}"] = f"=IFERROR(SMALL({sort_range},{k}),9999999999)"
        guard = f"{VAL_COL}{row}>=9999999999"
    ws[f"{IDX_COL}{row}"] = f'=IF({guard},"",MATCH({VAL_COL}{row},{sort_range},0)+{PLAN_FIRST_ROW-1})'
    return f"{IDX_COL}{row}"


def raw_index(idx_cell, sheet, col_letter):
    rng = f"'{sheet}'!${col_letter}${PLAN_FIRST_ROW}:${col_letter}${PLAN_LAST_ROW}"
    return f"INDEX({rng},{idx_cell}-{PLAN_FIRST_ROW-1})"


def field(idx_cell, sheet, col_letter):
    ri = raw_index(idx_cell, sheet, col_letter)
    return f'=IF({idx_cell}="","",IF({ri}="","",{ri}))'


def field_engine(idx_cell, col_letter):
    return field(idx_cell, "_ENGINE", col_letter)


def rank_ref_table(ws, row, k, table_range, mode, first_row_of_table, val_col=VAL_COL, idx_col=IDX_COL):
    if mode == "desc":
        ws[f"{val_col}{row}"] = f"=IFERROR(LARGE({table_range},{k}),-999999)"
        guard = f"{val_col}{row}<=-999999"
    else:
        ws[f"{val_col}{row}"] = f"=IFERROR(SMALL({table_range},{k}),9999999999)"
        guard = f"{val_col}{row}>=9999999999"
    ws[f"{idx_col}{row}"] = f'=IF({guard},"",MATCH({val_col}{row},{table_range},0)+{first_row_of_table-1})'
    return f"{idx_col}{row}"


def field_table(idx_cell, table_name, col_name, first_row_of_table, last_row_of_table):
    rng = f"{table_name}[{col_name}]"
    ri = f"INDEX({rng},{idx_cell}-{first_row_of_table-1})"
    return f'=IF({idx_cell}="","",IF({ri}="","",{ri}))'


# ========================================================= PRODUCTION COCKPIT
ws = wb.create_sheet("PRODUCTION COCKPIT")
wb.move_sheet("PRODUCTION COCKPIT", offset=-(len(wb.sheetnames) - 1))
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = ACCENT
ws.sheet_view.zoomScale = 85
for col in "ABCDEFGHIJKLMNOPQR":
    ws.column_dimensions[col].width = 3 if col in ("A", "R") else 14.5
ws.column_dimensions[VAL_COL].hidden = True
ws.column_dimensions[IDX_COL].hidden = True

COCKPIT_MAX_ROW = 420  # generöst tilltaget - täcker alla sektioner A-F
for r in range(1, COCKPIT_MAX_ROW):
    ws.row_dimensions[r].height = 16
    for col in "ABCDEFGHIJKLMNOPQR":
        ws[f"{col}{r}"].fill = PatternFill("solid", fgColor=DARK_BG)

# ---- Titelbanner ----------------------------------------------------------
ws.merge_cells("B1:Q2")
t = ws["B1"]
t.value = "NSV PRODUCTION COCKPIT"
t.font = Font(name=FONT, bold=True, size=26, color="FFFFFF")
t.fill = PatternFill("solid", fgColor=HEADER_BAR)
t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[1].height = 22
ws.row_dimensions[2].height = 22
for col in "ABCDEFGHIJKLMNOPQR":
    ws[f"{col}1"].fill = PatternFill("solid", fgColor=HEADER_BAR)
    ws[f"{col}2"].fill = PatternFill("solid", fgColor=HEADER_BAR)

ws.merge_cells("B3:Q3")
sub = ws["B3"]
sub.value = ('="Nytt sovringsverk 2025-2028   |   Vecka "&WEEKNUM(TODAY(),21)&", "&TEXT(TODAY(),"yyyy-mm-dd")'
             '&"   |   Källa: 6-WEEK PLAN (uppdateras automatiskt)"')
sub.font = Font(name=FONT, size=11, color=TEXT_MUTED)
sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[3].height = 18

# ---- Filterrad (rad 4 - adresser matchar F_VECKA m.fl. i _ENGINE) ---------
FROW = 4
ws.row_dimensions[FROW].height = 20
filter_defs = [
    ("B", "Vecka:", "C", WEEK_LIST),
    ("D", "Område:", "E", OMRADE_LIST),
    ("F", "Disciplin:", "G", None),
    ("H", "Aktivitetstyp:", "I", TYP_LIST),
    ("J", "Ansvarig/UE:", "K", None),
    ("L", "Status:", "M", STATUS_LIST),
]
for lab_col, lab_text, val_col, _lst in filter_defs:
    lc = ws[f"{lab_col}{FROW}"]
    lc.value = lab_text
    lc.font = Font(name=FONT, bold=True, size=9, color=TEXT_MUTED)
    lc.alignment = Alignment(horizontal="right", vertical="center")
    vc = ws[f"{val_col}{FROW}"]
    vc.value = "(Alla)"
    vc.font = Font(name=FONT, bold=True, size=10, color=DARK_BG)
    vc.fill = PatternFill("solid", fgColor="FFFFFF")
    vc.alignment = Alignment(horizontal="center", vertical="center")
    vc.border = all_border(ACCENT)

def literal_dv(literal_list, allow_blank=True):
    return DataValidation(type="list", formula1=f'"{literal_list}"', allow_blank=allow_blank)

def named_dv(named_ref, allow_blank=True):
    return DataValidation(type="list", formula1=named_ref, allow_blank=allow_blank, showErrorMessage=False)

dv_vecka = literal_dv(WEEK_LIST); ws.add_data_validation(dv_vecka); dv_vecka.add(f"C{FROW}")
dv_omrade = literal_dv(OMRADE_LIST); ws.add_data_validation(dv_omrade); dv_omrade.add(f"E{FROW}")
dv_disc = named_dv("=List_Disciplin_F"); ws.add_data_validation(dv_disc); dv_disc.add(f"G{FROW}")
dv_typ = literal_dv(TYP_LIST); ws.add_data_validation(dv_typ); dv_typ.add(f"I{FROW}")
dv_ansv = named_dv("=List_Ansvarig_F"); ws.add_data_validation(dv_ansv); dv_ansv.add(f"K{FROW}")
dv_status = literal_dv(STATUS_LIST); ws.add_data_validation(dv_status); dv_status.add(f"M{FROW}")

print("COCKPIT: bas & filterrad klara.")

# ---- KPI-kort (rad 7-10) ---------------------------------------------------
ENG_RANGE = lambda col: f"_ENGINE!${col}${PLAN_FIRST_ROW}:${col}${PLAN_LAST_ROW}"
STAT_RANGE = lambda col: f"'PRODUCTION STATUS'!${col}${PLAN_FIRST_ROW}:${col}${PLAN_LAST_ROW}"

TOTAL_F = f"SUMPRODUCT({ENG_RANGE('D')})"

def count_status(status_val):
    return f"SUMPRODUCT({ENG_RANGE('D')}*({STAT_RANGE('L')}=\"{status_val}\"))"

KPI_ROW = 7
tiles = [
    ("AKTIVITETER\n(FILTER)", f"={TOTAL_F}", "0", ACCENT),
    ("ENLIGT PLAN /\nPÅGÅENDE", f'=IFERROR(({count_status("Pågående")}+{count_status("Klar")})/{TOTAL_F},0)', "0%", GREEN),
    ("RISK", f'=IFERROR({count_status("Risk")}/{TOTAL_F},0)', "0%", AMBER),
    ("FÖRSENADE", f'=IFERROR({count_status("Försenad")}/{TOTAL_F},0)', "0%", RED),
    ("EJ STARTADE", f'={count_status("Ej startad")}', "0", GREY),
    ("ÖPPNA\nBLOCKERS", '=COUNTIFS(tblRisks[Status],"<>Löst")', "0", RED),
    ("MATERIAL-\nRISKER", '=COUNTIFS(tblMaterial[Risk],"Ja")', "0", AMBER),
    ("HANDLINGS-\nRISKER", '=COUNTIFS(tblDocuments[Risk],"Ja")', "0", AMBER),
]
tile_cols = ["B", "D", "F", "H", "J", "L", "N", "P"]

def draw_tile(row, col, title, formula, fmt, color, size=24):
    ws.row_dimensions[row].height = 14
    ws.row_dimensions[row + 1].height = 34
    ws.row_dimensions[row + 2].height = 14
    col2 = chr(ord(col) + 1)
    ws.merge_cells(f"{col}{row}:{col2}{row}")
    lc = ws[f"{col}{row}"]
    lc.value = title.replace("\n", " ")
    lc.font = Font(name=FONT, bold=True, size=8, color=TEXT_MUTED)
    lc.alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=True, indent=1)
    ws.merge_cells(f"{col}{row+1}:{col2}{row+1}")
    vc = ws[f"{col}{row+1}"]
    vc.value = formula
    vc.number_format = fmt
    vc.font = Font(name=FONT, bold=True, size=size, color=color)
    vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(f"{col}{row+2}:{col2}{row+2}")
    bc = ws[f"{col}{row+2}"]
    bc.border = Border(bottom=Side(style="medium", color=color))

for (title, formula, fmt, color), col in zip(tiles, tile_cols):
    draw_tile(KPI_ROW, col, title, formula, fmt, color)
print("COCKPIT: KPI-kort klara.")

def section_header(row, text):
    ws.merge_cells(f"B{row}:Q{row}")
    c = ws[f"B{row}"]
    c.value = text
    c.font = Font(name=FONT, bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=PANEL_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 22
    for col in "ABCDEFGHIJKLMNOPQR":
        ws[f"{col}{row}"].fill = PatternFill("solid", fgColor=PANEL_BG)

def sub_header(row, headers_cols):
    for col, text in headers_cols:
        c = ws[f"{col}{row}"]
        c.value = text
        c.font = Font(name=FONT, bold=True, size=9, color=TEXT_MUTED)
        c.fill = PatternFill("solid", fgColor=CARD_BG)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 16

def data_row_bg(row, last_col="Q"):
    for col in "BCDEFGHIJKLMNOPQ":
        ws[f"{col}{row}"].fill = PatternFill("solid", fgColor=CARD_BG2 if row % 2 else CARD_BG)
        ws[f"{col}{row}"].font = Font(name=FONT, size=9.5, color=TEXT_LIGHT)

# ==================================================================== A ====
ROW = 11
section_header(ROW, "A.  6-VECKORS PRODUKTIONSPLAN  —  filtrerad vy (vecka -1 till +6 från idag)")
ROW += 1
cap = ws[f"B{ROW}"]
cap.value = (f'="Visar de "&MIN(60,SUMPRODUCT(({ENG_RANGE("I")}<9999999999)*1))&" närmaste av totalt "'
             f'&SUMPRODUCT(({ENG_RANGE("I")}<9999999999)*1)&" aktiviteter i fönstret (enligt valt filter)"')
cap.font = Font(name=FONT, italic=True, size=9, color=TEXT_MUTED)
ws.merge_cells(f"B{ROW}:Q{ROW}")
ROW += 1
sub_header(ROW, [("B", "Vecka"), ("C", "Start"), ("D", "Slut"), ("E", "Aktivitet"),
                  ("I", "Ansvarig/UE"), ("J", "Status"), ("K", "Framdrift (faktisk)"), ("L", "Avvikelse")])
ROW += 1
A_FIRST = ROW
A_COUNT = 60
for k in range(1, A_COUNT + 1):
    r = ROW
    idx = rank_ref(ws, r, k, ENG_RANGE("I"), "asc")
    ws[f"B{r}"] = field(idx, "6-WEEK PLAN", "N")
    ws[f"C{r}"] = field(idx, "6-WEEK PLAN", "I"); ws[f"C{r}"].number_format = "yyyy-mm-dd"
    ws[f"D{r}"] = field(idx, "6-WEEK PLAN", "K"); ws[f"D{r}"].number_format = "yyyy-mm-dd"
    ws.merge_cells(f"E{r}:H{r}")
    ws[f"E{r}"] = field(idx, "6-WEEK PLAN", "D")
    ws[f"I{r}"] = field(idx, "6-WEEK PLAN", "O")
    ws[f"J{r}"] = field_engine(idx, "C")
    ws[f"K{r}"] = (f'=IF({idx}="","",REPT("█",ROUND(IFERROR({raw_index(idx,"PRODUCTION STATUS","J")},0)*20,0))'
                   f'&"  "&TEXT(IFERROR({raw_index(idx,"PRODUCTION STATUS","J")},0),"0%"))')
    ws[f"L{r}"] = field(idx, "PRODUCTION STATUS", "K"); ws[f"L{r}"].number_format = "0%;-0%;\"\""
    data_row_bg(r)
    ROW += 1
A_LAST = ROW - 1
print("COCKPIT: sektion A klar.")

# ==================================================================== B ====
ROW = A_LAST + 2
section_header(ROW, "B.  PLANERAT VS FAKTISKT  —  produktionsframdrift (enligt valt filter)")
ROW += 1
B_ROW = ROW
PLAN_PCT_CELL = f"S{B_ROW+1}"
FAKT_PCT_CELL = f"T{B_ROW+1}"
ws[PLAN_PCT_CELL] = f"=IFERROR(SUMPRODUCT({ENG_RANGE('D')}*{STAT_RANGE('G')})/{TOTAL_F},0)"
ws[FAKT_PCT_CELL] = (f"=IFERROR(SUMPRODUCT({ENG_RANGE('D')}*({STAT_RANGE('J')}<>\"\")*{STAT_RANGE('J')})"
                      f"/SUMPRODUCT({ENG_RANGE('D')}*({STAT_RANGE('J')}<>\"\")),0)")
ws.column_dimensions["S"].hidden = True
ws.column_dimensions["T"].hidden = True

draw_tile(B_ROW, "B", "PLANERAD\nFRAMDRIFT (SNITT)", f"={PLAN_PCT_CELL}", "0%", ACCENT)
draw_tile(B_ROW, "D", "FAKTISK\nFRAMDRIFT (SNITT)", f"={FAKT_PCT_CELL}", "0%", GREEN)
draw_tile(B_ROW, "F", "AVVIKELSE", f"={FAKT_PCT_CELL}-{PLAN_PCT_CELL}", "0%;-0%;\"±0%\"", AMBER)
trend_formula = (f'=IF({FAKT_PCT_CELL}-{PLAN_PCT_CELL}>=-0.03,"→ Stabil",'
                  f'IF({FAKT_PCT_CELL}-{PLAN_PCT_CELL}>=-0.1,"↘ Något efter","↓ Tydligt efter"))')
draw_tile(B_ROW, "H", "TREND", trend_formula, "@", TEXT_LIGHT, size=16)
note = ws[f"J{B_ROW+1}"]
ws.merge_cells(f"J{B_ROW+1}:Q{B_ROW+1}")
note.value = ("Baserat på hittills rapporterad faktisk framdrift i PRODUCTION STATUS. "
              "Logga veckans snittvärden i _HISTORIK för att bygga upp en trendkurva över tid.")
note.font = Font(name=FONT, italic=True, size=8.5, color=TEXT_MUTED)
note.alignment = Alignment(wrap_text=True, vertical="center")
B_LAST = B_ROW + 2
print("COCKPIT: sektion B klar.")

def orsak_formula(idx_cell):
    status = raw_index(idx_cell, "PRODUCTION STATUS", "L")
    foreg = raw_index(idx_cell, "_ENGINE", "B")
    actid = raw_index(idx_cell, "6-WEEK PLAN", "A")
    blockers = f'COUNTIFS(tblRisks[Aktivitet-ID],{actid},tblRisks[Status],"<>Löst")'
    matrisk = f'COUNTIFS(tblMaterial[Aktivitet-ID],{actid},tblMaterial[Risk],"Ja")'
    docrisk = f'COUNTIFS(tblDocuments[Aktivitet-ID],{actid},tblDocuments[Risk],"Ja")'
    expr = (f'IF({status}="Försenad","Försenad; ","")'
            f'&IF({status}="Risk","Riskstart; ","")'
            f'&IF({foreg}=1,"Föregående ej klar; ","")'
            f'&IF({blockers}>0,"Öppet hinder; ","")'
            f'&IF({matrisk}>0,"Materialrisk; ","")'
            f'&IF({docrisk}>0,"Handlingsrisk; ","")')
    return f'=IF({idx_cell}="","",IF({expr}="","OK",{expr}))'

# ==================================================================== C ====
ROW = B_LAST + 2
section_header(ROW, "C.  KRITISKA AKTIVITETER  —  topp 10 som kräver mest uppmärksamhet just nu")
ROW += 1
sub_header(ROW, [("B", "#"), ("C", "Aktivitet"), ("G", "Område"), ("I", "Ansvarig/UE"),
                  ("J", "Status"), ("K", "Poäng"), ("L", "Orsak")])
ROW += 1
C_FIRST = ROW
for k in range(1, 11):
    r = ROW
    idx = rank_ref(ws, r, k, ENG_RANGE("F"), "desc")
    ws[f"B{r}"] = f'=IF({idx}="","",{k})'
    ws.merge_cells(f"C{r}:F{r}")
    ws[f"C{r}"] = field(idx, "6-WEEK PLAN", "D")
    ws[f"G{r}"] = field(idx, "6-WEEK PLAN", "B")
    ws.merge_cells(f"G{r}:H{r}")
    ws[f"I{r}"] = field(idx, "6-WEEK PLAN", "O")
    ws[f"J{r}"] = field_engine(idx, "C")
    ws[f"K{r}"] = field(idx, "PRODUCTION STATUS", "M")
    ws.merge_cells(f"L{r}:Q{r}")
    ws[f"L{r}"] = orsak_formula(idx)
    data_row_bg(r)
    ROW += 1
C_LAST = ROW - 1
print("COCKPIT: sektion C klar.")

# ==================================================================== D ====
ROW = C_LAST + 2
section_header(ROW, "D.  KOMMANDE 7 DAGAR  —  aktiviteter som startar inom en vecka")
ROW += 1
cap = ws[f"B{ROW}"]
cap.value = (f'="Visar de "&MIN(30,SUMPRODUCT(({ENG_RANGE("G")}<9999999999)*1))&" närmaste av totalt "'
             f'&SUMPRODUCT(({ENG_RANGE("G")}<9999999999)*1)&" som startar inom 7 dagar (enligt valt filter)"')
cap.font = Font(name=FONT, italic=True, size=9, color=TEXT_MUTED)
ws.merge_cells(f"B{ROW}:Q{ROW}")
ROW += 1
sub_header(ROW, [("B", "Start"), ("C", "Aktivitet"), ("G", "Ansvarig/UE"), ("I", "Status"),
                  ("J", "Förutsättning"), ("L", "Risk")])
ROW += 1
D_FIRST = ROW
for k in range(1, 31):
    r = ROW
    idx = rank_ref(ws, r, k, ENG_RANGE("G"), "asc")
    ws[f"B{r}"] = field(idx, "6-WEEK PLAN", "I"); ws[f"B{r}"].number_format = "yyyy-mm-dd"
    ws.merge_cells(f"C{r}:F{r}")
    ws[f"C{r}"] = field(idx, "6-WEEK PLAN", "D")
    ws.merge_cells(f"G{r}:H{r}")
    ws[f"G{r}"] = field(idx, "6-WEEK PLAN", "O")
    ws[f"I{r}"] = field_engine(idx, "C")
    ws.merge_cells(f"J{r}:K{r}")
    ws[f"J{r}"] = (f'=IF({idx}="","",IF({raw_index(idx,"_ENGINE","B")}=1,'
                   f'"Föregående aktivitet ej klar","OK"))')
    ws.merge_cells(f"L{r}:Q{r}")
    ws[f"L{r}"] = orsak_formula(idx)
    data_row_bg(r)
    ROW += 1
D_LAST = ROW - 1
print("COCKPIT: sektion D klar.")

# ==================================================================== E ====
ROW = D_LAST + 2
section_header(ROW, "E.  KOMMANDE 30 DAGAR  —  aktiviteter, materialleveranser, handlingar och milstolpar")
ROW += 1

# E1: Aktiviteter
ws[f"B{ROW}"] = "Aktiviteter"
ws[f"B{ROW}"].font = Font(name=FONT, bold=True, size=10, color=ACCENT)
ROW += 1
sub_header(ROW, [("B", "Start"), ("C", "Aktivitet"), ("G", "Område"), ("I", "Ansvarig/UE"), ("J", "Status")])
ROW += 1
for k in range(1, 21):
    r = ROW
    idx = rank_ref(ws, r, k, ENG_RANGE("H"), "asc")
    ws[f"B{r}"] = field(idx, "6-WEEK PLAN", "I"); ws[f"B{r}"].number_format = "yyyy-mm-dd"
    ws.merge_cells(f"C{r}:F{r}")
    ws[f"C{r}"] = field(idx, "6-WEEK PLAN", "D")
    ws[f"G{r}"] = field(idx, "6-WEEK PLAN", "B")
    ws.merge_cells(f"G{r}:H{r}")
    ws[f"I{r}"] = field(idx, "6-WEEK PLAN", "O")
    ws.merge_cells(f"J{r}:K{r}")
    ws[f"J{r}"] = field_engine(idx, "C")
    data_row_bg(r)
    ROW += 1
ROW += 1

# E2: Materialleveranser
ws[f"B{ROW}"] = "Materialleveranser"
ws[f"B{ROW}"].font = Font(name=FONT, bold=True, size=10, color=ACCENT)
ROW += 1
sub_header(ROW, [("B", "Planerad leverans"), ("D", "Material"), ("G", "Aktivitet"), ("K", "Status"), ("M", "Risk")])
ROW += 1
MAT_SORT_RANGE = "tblMaterial[Sortordning]"
for k in range(1, 11):
    r = ROW
    idx = rank_ref_table(ws, r, k, MAT_SORT_RANGE, "asc", 3)
    ws[f"B{r}"] = field_table(idx, "tblMaterial", "Planerad leverans", 3, MAT_LAST_ROW)
    ws[f"B{r}"].number_format = "yyyy-mm-dd"
    ws.merge_cells(f"D{r}:F{r}")
    ws[f"D{r}"] = field_table(idx, "tblMaterial", "Material", 3, MAT_LAST_ROW)
    ws.merge_cells(f"G{r}:J{r}")
    ws[f"G{r}"] = field_table(idx, "tblMaterial", "Aktivitet", 3, MAT_LAST_ROW)
    ws.merge_cells(f"K{r}:L{r}")
    ws[f"K{r}"] = field_table(idx, "tblMaterial", "Status", 3, MAT_LAST_ROW)
    ws.merge_cells(f"M{r}:Q{r}")
    ws[f"M{r}"] = field_table(idx, "tblMaterial", "Risk", 3, MAT_LAST_ROW)
    data_row_bg(r)
    ROW += 1
ROW += 1

# E3: Handlingar
ws[f"B{ROW}"] = "Handlingar"
ws[f"B{ROW}"].font = Font(name=FONT, bold=True, size=10, color=ACCENT)
ROW += 1
sub_header(ROW, [("B", "Behövs senast"), ("D", "Handling"), ("G", "Aktivitet"), ("K", "Status"), ("M", "Risk")])
ROW += 1
DOC_SORT_RANGE = "tblDocuments[Sortordning]"
for k in range(1, 11):
    r = ROW
    idx = rank_ref_table(ws, r, k, DOC_SORT_RANGE, "asc", 3)
    ws[f"B{r}"] = field_table(idx, "tblDocuments", "Behövs senast", 3, DOC_LAST_ROW)
    ws[f"B{r}"].number_format = "yyyy-mm-dd"
    ws.merge_cells(f"D{r}:F{r}")
    ws[f"D{r}"] = field_table(idx, "tblDocuments", "Handling", 3, DOC_LAST_ROW)
    ws.merge_cells(f"G{r}:J{r}")
    ws[f"G{r}"] = field_table(idx, "tblDocuments", "Aktivitet", 3, DOC_LAST_ROW)
    ws.merge_cells(f"K{r}:L{r}")
    ws[f"K{r}"] = field_table(idx, "tblDocuments", "Status", 3, DOC_LAST_ROW)
    ws.merge_cells(f"M{r}:Q{r}")
    ws[f"M{r}"] = field_table(idx, "tblDocuments", "Risk", 3, DOC_LAST_ROW)
    data_row_bg(r)
    ROW += 1
ROW += 1

# E4: Milstolpar
ws[f"B{ROW}"] = "Milstolpar (0-dagars aktiviteter / överlämningar)"
ws[f"B{ROW}"].font = Font(name=FONT, bold=True, size=10, color=ACCENT)
ROW += 1
sub_header(ROW, [("B", "Datum"), ("C", "Milstolpe"), ("G", "Område"), ("I", "DP")])
ROW += 1
for k in range(1, 16):
    r = ROW
    idx = rank_ref(ws, r, k, ENG_RANGE("J"), "asc")
    ws[f"B{r}"] = field(idx, "6-WEEK PLAN", "I"); ws[f"B{r}"].number_format = "yyyy-mm-dd"
    ws.merge_cells(f"C{r}:F{r}")
    ws[f"C{r}"] = field(idx, "6-WEEK PLAN", "D")
    ws[f"G{r}"] = field(idx, "6-WEEK PLAN", "B")
    ws.merge_cells(f"G{r}:H{r}")
    ws.merge_cells(f"I{r}:J{r}")
    ws[f"I{r}"] = field(idx, "6-WEEK PLAN", "F")
    data_row_bg(r)
    ROW += 1
E_LAST = ROW - 1
print("COCKPIT: sektion E klar.")

# ==================================================================== F ====
ROW = E_LAST + 2
section_header(ROW, "F.  BLOCKERS / HINDER  —  öppna produktionshinder kopplade till aktiviteter")
ROW += 1
sub_header(ROW, [("B", "Förfaller"), ("C", "Aktivitet"), ("G", "Problem"), ("K", "Typ"),
                  ("L", "Ansvarig"), ("M", "Status")])
ROW += 1
RISK_SORT_RANGE = "tblRisks[Sortordning]"
for k in range(1, 16):
    r = ROW
    idx = rank_ref_table(ws, r, k, RISK_SORT_RANGE, "asc", 3)
    ws[f"B{r}"] = field_table(idx, "tblRisks", "Förfallodatum", 3, RISK_LAST_ROW)
    ws[f"B{r}"].number_format = "yyyy-mm-dd"
    ws.merge_cells(f"C{r}:F{r}")
    ws[f"C{r}"] = field_table(idx, "tblRisks", "Aktivitet", 3, RISK_LAST_ROW)
    ws.merge_cells(f"G{r}:J{r}")
    ws[f"G{r}"] = field_table(idx, "tblRisks", "Problem", 3, RISK_LAST_ROW)
    ws[f"K{r}"] = field_table(idx, "tblRisks", "Typ", 3, RISK_LAST_ROW)
    ws[f"L{r}"] = field_table(idx, "tblRisks", "Ansvarig", 3, RISK_LAST_ROW)
    ws.merge_cells(f"M{r}:Q{r}")
    ws[f"M{r}"] = field_table(idx, "tblRisks", "Status", 3, RISK_LAST_ROW)
    data_row_bg(r)
    ROW += 1
F_LAST = ROW - 1
print("COCKPIT: sektion F klar.")

COCKPIT_LAST_ROW = F_LAST
ws.freeze_panes = "B11"

# ============================================================ AI ASSISTANT =
ai = wb.create_sheet("AI ASSISTANT")
ai.sheet_view.showGridLines = False
for col in "ABCDEFGHIJKLMNOPQ":
    ai.column_dimensions[col].width = 3 if col in ("A",) else 14.5
ai.column_dimensions["V"].hidden = True

banner(ai, "A1:Q2", "AI ASSISTANT  —  Förberedd för AI-/API-integration (Trimble Connect, BIM, LLM)", bg=LIGHT_HEADER, fg="FFFFFF", size=16)
intro = ai["B3"]
ai.merge_cells("B3:Q4")
intro.value = ("Detta blad innehåller fem beslutsstöd. Nyckeltalen och listorna nedan beräknas redan idag "
               "med vanliga Excel-formler direkt ur 6-WEEK PLAN / PRODUCTION STATUS / RISKS & BLOCKERS / "
               "MATERIAL / DOCUMENTS. Det som saknas för att göra samma sak i fritext (en skriven analys "
               "eller ett mötesunderlag i löptext) är en koppling till ett AI/LLM-API (t.ex. Claude) som kan "
               "läsa dessa tabeller och skriva en sammanfattning - markerat med 🔌 nedan.")
intro.font = Font(name=FONT, italic=True, size=10, color="444444")
intro.alignment = Alignment(wrap_text=True, vertical="top")
ai.row_dimensions[3].height = 15
ai.row_dimensions[4].height = 30

def ai_section(row, title, desc):
    ai.merge_cells(f"B{row}:Q{row}")
    c = ai[f"B{row}"]
    c.value = title
    c.font = Font(name=FONT, bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=LIGHT_HEADER)
    c.alignment = Alignment(vertical="center", indent=1)
    ai.row_dimensions[row].height = 22
    ai.merge_cells(f"B{row+1}:Q{row+1}")
    d = ai[f"B{row+1}"]
    d.value = desc
    d.font = Font(name=FONT, italic=True, size=9.5, color="555555")
    d.alignment = Alignment(wrap_text=True, vertical="top")
    ai.row_dimensions[row+1].height = 28
    return row + 2

ROW = 6
ROW = ai_section(ROW, "1. ANALYSERA PLAN — kommande 2 veckor",
    "🔌 Framtida AI-funktion: fri textanalys av förseningar, risker, saknade förutsättningar, "
    "kritiska aktiviteter, material- och dokumentproblem. Nyckeltalen nedan är redan liveberäknade.")
ai_metrics_1 = [
    ("Aktiviteter som startar kommande 14 dagar", f'=SUMPRODUCT(({STAT_RANGE("E")}>=TODAY())*({STAT_RANGE("E")}<=TODAY()+14))'),
    ("...varav Risk", f'=SUMPRODUCT(({STAT_RANGE("E")}>=TODAY())*({STAT_RANGE("E")}<=TODAY()+14)*({STAT_RANGE("L")}="Risk"))'),
    ("Aktiviteter försenade just nu (totalt)", f'=SUMPRODUCT(({STAT_RANGE("L")}="Försenad")*1)'),
    ("Kritiska aktiviteter (poäng > 0)", f'=SUMPRODUCT(({STAT_RANGE("M")}>0)*1)'),
    ("Öppna materialrisker", '=COUNTIFS(tblMaterial[Risk],"Ja")'),
    ("Öppna handlingsrisker (dokument)", '=COUNTIFS(tblDocuments[Risk],"Ja")'),
]
for label, f in ai_metrics_1:
    ai[f"B{ROW}"] = label; ai.merge_cells(f"B{ROW}:M{ROW}")
    ai[f"N{ROW}"] = f; ai.merge_cells(f"N{ROW}:Q{ROW}")
    ai[f"N{ROW}"].font = Font(name=FONT, bold=True, size=11, color=LIGHT_HEADER)
    ai[f"N{ROW}"].alignment = Alignment(horizontal="center")
    ROW += 1
ROW += 1

ROW = ai_section(ROW, "2. FÖRBERED PRODUKTIONSMÖTE — underlag",
    "🔌 Framtida AI-funktion: sammanhängande mötesunderlag i löptext (7 punkter enligt begäran). "
    "Byggstenarna/talen nedan kan användas direkt eller klistras in i mötesanteckningar redan idag.")
ai_meeting = [
    ("1. Enligt plan / klart", f'=({count_status("Klar")}+{count_status("Pågående")})&" st av {N_ACT}"'),
    ("2. Ej enligt plan (risk)", f'={count_status("Risk")}&" st"'),
    ("3. Försenat", f'={count_status("Försenad")}&" st"'),
    ("4. Kritiska aktiviteter (se Cockpit sektion C)", f'=SUMPRODUCT(({STAT_RANGE("M")}>0)*1)&" st med kritikpoäng > 0"'),
    ("5. Blockers (se RISKS & BLOCKERS)", '=COUNTIFS(tblRisks[Status],"<>Löst")&" öppna hinder"'),
    ("6. Beslut som behövs (Typ=Beslut, ej löst)", '=COUNTIFS(tblRisks[Typ],"Beslut",tblRisks[Status],"<>Löst")&" st"'),
    ("7. Uppföljning nästa vecka", '="Se sektion D (Kommande 7 dagar) på PRODUCTION COCKPIT"'),
]
for label, f in ai_meeting:
    ai[f"B{ROW}"] = label; ai.merge_cells(f"B{ROW}:H{ROW}")
    ai[f"I{ROW}"] = f; ai.merge_cells(f"I{ROW}:Q{ROW}")
    ai[f"I{ROW}"].font = Font(name=FONT, bold=True, size=10, color=LIGHT_HEADER)
    ROW += 1
ROW += 1

ROW = ai_section(ROW, "3. VAD SKA JAG FÖLJA UPP? — topp 5 just nu",
    "Automatiskt beräknat: de aktiviteter med högst kritikpoäng (försenad + risk + öppna hinder/material-/"
    "handlingsrisk + föregående aktivitet ej klar). Samma logik som Cockpit sektion C, kortad till 5 rader.")
sub_header_ai_row = ROW
for col, txt in [("B", "#"), ("C", "Aktivitet"), ("J", "Ansvarig/UE"), ("L", "Status"), ("N", "Poäng")]:
    ai[f"{col}{ROW}"] = txt
    ai[f"{col}{ROW}"].font = Font(name=FONT, bold=True, size=9, color="666666")
ROW += 1
for k in range(1, 6):
    r = ROW
    idx = rank_ref(ai, r, k, ENG_RANGE("F"), "desc")
    ai[f"B{r}"] = f'=IF({idx}="","",{k})'
    ai.merge_cells(f"C{r}:I{r}")
    ai[f"C{r}"] = field(idx, "6-WEEK PLAN", "D")
    ai.merge_cells(f"J{r}:K{r}")
    ai[f"J{r}"] = field(idx, "6-WEEK PLAN", "O")
    ai.merge_cells(f"L{r}:M{r}")
    ai[f"L{r}"] = field_engine(idx, "C")
    ai[f"N{r}"] = field(idx, "PRODUCTION STATUS", "M")
    ROW += 1
ROW += 1

ROW = ai_section(ROW, "4. RISKER KOMMANDE 2 VECKOR",
    "Automatiskt beräknat: aktiviteter med status Risk som startar inom 14 dagar, samt öppna hinder "
    "med förfallodatum inom 14 dagar.")
for col, txt in [("B", "Start"), ("C", "Aktivitet"), ("J", "Ansvarig/UE"), ("L", "Orsak")]:
    ai[f"{col}{ROW}"] = txt
    ai[f"{col}{ROW}"].font = Font(name=FONT, bold=True, size=9, color="666666")
ROW += 1
for k in range(1, 8):
    r = ROW
    idx = rank_ref(ai, r, k, ENG_RANGE("G"), "asc")
    ai[f"B{r}"] = field(idx, "6-WEEK PLAN", "I"); ai[f"B{r}"].number_format = "yyyy-mm-dd"
    ai.merge_cells(f"C{r}:I{r}")
    ai[f"C{r}"] = field(idx, "6-WEEK PLAN", "D")
    ai.merge_cells(f"J{r}:K{r}")
    ai[f"J{r}"] = field(idx, "6-WEEK PLAN", "O")
    ai.merge_cells(f"L{r}:Q{r}")
    ai[f"L{r}"] = orsak_formula(idx)
    ROW += 1
ROW += 1

ROW = ai_section(ROW, "5. VAD PÅVERKAS? — konsekvensanalys vid försening",
    "Välj en aktivitet i listan nedan. Systemet visar aktiviteter i samma del/linje och område med senare "
    "planerad start - dvs de som troligen påverkas om den valda aktiviteten blir försenad. Baseras på "
    "ordningsföljden i planen, inte en fullständig länkad beroendekedja (kräver BIM/Trimble Connect-koppling).")
ai[f"B{ROW}"] = "Välj aktivitet (ID):"
ai[f"B{ROW}"].font = Font(name=FONT, bold=True, size=10)
ai.merge_cells(f"B{ROW}:C{ROW}")
SEL_CELL = f"D{ROW}"
ai[SEL_CELL] = ACTIVITIES[10]["id"]
ai[SEL_CELL].fill = PatternFill("solid", fgColor="FFF9DB")
ai.merge_cells(f"D{ROW}:F{ROW}")
dv = DataValidation(type="list", formula1="=List_AktivitetID", allow_blank=True)
ai.add_data_validation(dv); dv.add(SEL_CELL)
ai[f"H{ROW}"] = f'="Vald aktivitet: "&IFERROR(INDEX(tblPlan[Aktivitet],MATCH({SEL_CELL},tblPlan[ID],0)),"okänt ID")'
ai.merge_cells(f"H{ROW}:Q{ROW}")
ai[f"H{ROW}"].font = Font(name=FONT, italic=True, size=9.5, color="555555")
ROW += 1
ai[f"V{ROW}"] = f'=IFERROR(INDEX(tblPlan[Del/Linje],MATCH({SEL_CELL},tblPlan[ID],0)),"")'
ai[f"V{ROW+1}"] = f'=IFERROR(INDEX(tblPlan[Område],MATCH({SEL_CELL},tblPlan[ID],0)),"")'
ai[f"V{ROW+2}"] = f'=IFERROR(INDEX(tblPlan[Planerad start],MATCH({SEL_CELL},tblPlan[ID],0)),0)'
add_name("AI_SEL_LINJE", f"'AI ASSISTANT'!$V${ROW}")
add_name("AI_SEL_OMRADE", f"'AI ASSISTANT'!$V${ROW+1}")
add_name("AI_SEL_START", f"'AI ASSISTANT'!$V${ROW+2}")
ROW += 4
for col, txt in [("B", "Start"), ("C", "Möjligen påverkad aktivitet"), ("J", "Ansvarig/UE"), ("L", "Status")]:
    ai[f"{col}{ROW}"] = txt
    ai[f"{col}{ROW}"].font = Font(name=FONT, bold=True, size=9, color="666666")
ROW += 1
for k in range(1, 8):
    r = ROW
    idx = rank_ref(ai, r, k, ENG_RANGE("K"), "asc")
    ai[f"B{r}"] = field(idx, "6-WEEK PLAN", "I"); ai[f"B{r}"].number_format = "yyyy-mm-dd"
    ai.merge_cells(f"C{r}:I{r}")
    ai[f"C{r}"] = field(idx, "6-WEEK PLAN", "D")
    ai.merge_cells(f"J{r}:K{r}")
    ai[f"J{r}"] = field(idx, "6-WEEK PLAN", "O")
    ai.merge_cells(f"L{r}:M{r}")
    ai[f"L{r}"] = field_engine(idx, "C")
    ROW += 1
ai[f"V{ROW}"] = ""  # scratch column terminator

ai.column_dimensions[VAL_COL].hidden = True
ai.column_dimensions[IDX_COL].hidden = True
print("AI ASSISTANT klart.")

# ================================================================ _HISTORIK
hs = wb.create_sheet("_HISTORIK")
hs.sheet_view.showGridLines = False
banner(hs, "A1:G1", "_HISTORIK  —  Veckovis trendlogg (loggas manuellt eller via Office Script)", bg=LIGHT_HEADER, fg="FFFFFF", size=13)
HIST_HEADERS = ["Datum", "Planerad framdrift (snitt)", "Faktisk framdrift (snitt)", "Avvikelse",
                "Antal försenade", "Antal risk", "Antal klara"]
for i, h in enumerate(HIST_HEADERS, start=1):
    c = hs.cell(row=2, column=i, value=h)
    style_header_cell(c, bg=LIGHT_HEADER, fg="FFFFFF")

hs["A3"] = datetime.datetime(2026, 9, 5)
hs["A3"].number_format = "yyyy-mm-dd"
hs["B3"] = f"='PRODUCTION COCKPIT'!S{B_ROW+1}"
hs["C3"] = f"='PRODUCTION COCKPIT'!T{B_ROW+1}"
hs["D3"] = "=C3-B3"
hs["E3"] = f'=SUMPRODUCT((\'PRODUCTION STATUS\'!$L${PLAN_FIRST_ROW}:$L${PLAN_LAST_ROW}="Försenad")*1)'
hs["F3"] = f'=SUMPRODUCT((\'PRODUCTION STATUS\'!$L${PLAN_FIRST_ROW}:$L${PLAN_LAST_ROW}="Risk")*1)'
hs["G3"] = f'=SUMPRODUCT((\'PRODUCTION STATUS\'!$L${PLAN_FIRST_ROW}:$L${PLAN_LAST_ROW}="Klar")*1)'
for col in ("B", "C", "D"):
    hs[f"{col}3"].number_format = "0%"

note = hs["A5"]
hs.merge_cells("A5:G6")
note.value = ("Rad 3 ovan uppdateras automatiskt med DAGENS värden varje gång filen öppnas (den fryser INTE "
              "historiken). För en riktig trendkurva: kopiera rad 3 -> Klistra in som VÄRDEN på en ny rad "
              "varje vecka (eller vid varje statusavstämning). Tabellen expanderar automatiskt och grafen "
              "nedan uppdateras då automatiskt.")
note.font = Font(name=FONT, italic=True, size=9, color="666666")
note.alignment = Alignment(wrap_text=True, vertical="top")

add_table(hs, "A2:G3", "tblHistorik", style="TableStyleLight9")
set_col_widths(hs, {"A": 13, "B": 22, "C": 22, "D": 12, "E": 14, "F": 12, "G": 12})

chart = LineChart()
chart.title = "Planerad vs faktisk framdrift över tid"
chart.style = 2
chart.y_axis.title = "Framdrift"
chart.x_axis.title = "Datum"
chart.height = 8
chart.width = 20
data = Reference(hs, min_col=2, max_col=3, min_row=2, max_row=3)
cats = Reference(hs, min_col=1, min_row=3, max_row=3)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
hs.add_chart(chart, "A9")
print("_HISTORIK klart.")

# ============================================================ finalisering =
_ENGINE = wb["_ENGINE"]
_ENGINE.sheet_state = "hidden"

TAB_COLORS = {
    "PRODUCTION COCKPIT": "3DA9FC",
    "6-WEEK PLAN": "13233B",
    "PRODUCTION STATUS": "13233B",
    "RISKS & BLOCKERS": "C0504D",
    "MATERIAL": "C0504D",
    "DOCUMENTS": "C0504D",
    "AI ASSISTANT": "7030A0",
    "LISTOR": "808080",
    "_HISTORIK": "808080",
}
for name, color in TAB_COLORS.items():
    wb[name].sheet_properties.tabColor = color

DESIRED_ORDER = [
    "PRODUCTION COCKPIT", "6-WEEK PLAN", "PRODUCTION STATUS",
    "RISKS & BLOCKERS", "MATERIAL", "DOCUMENTS", "AI ASSISTANT",
    "741 - SEKTIONSFICKOR", "741 - RESURS EXP", "741 - RESURSÖVERSIKT",
    "742 - SIKTHALL", "744 - Fläkthus", "745 - FÖRTJOCKARHUS", "775 - VATTENRESERVOAR",
    "DP1 - Kvarstående arbeten", "AKT – SAMT WBS-OMRÅDEN", "LISTOR",
    "_HISTORIK", "_ENGINE",
]
assert set(DESIRED_ORDER) == set(wb.sheetnames), (
    set(DESIRED_ORDER) ^ set(wb.sheetnames))
wb._sheets = [wb[name] for name in DESIRED_ORDER]
wb.active = 0

wb.save(OUT)
print(f"SPARAD: {OUT}")
