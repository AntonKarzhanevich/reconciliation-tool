"""
Add two sheets to the report:
  15. Δ1 Банк vs РКО  — which bank payments have no/partial write-off from kassa Vlad
  16. Δ2 WB vs ПКО   — which WB weeks have no/partial PKO in 1C (by month)
"""
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
from collections import defaultdict

ROOT   = r"D:\Cursor\kassa"
REPORT = ROOT + r"\СВОДНЫЙ ОТЧЁТ ВБ кассы и дебиторка.xlsx"
RATE   = 3.5

THIN   = Side(border_style="thin",   color="BBBBBB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD   = Font(bold=True)
TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
H2_FONT    = Font(bold=True, size=11, color="FFFFFF")
NUM_FORMAT = '#,##0.00;[Red]-#,##0.00;""'
WRAP   = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left", vertical="top")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
H2_FILL    = PatternFill("solid", fgColor="2E75B6")
SUB_FILL   = PatternFill("solid", fgColor="D9E1F2")
OK_FILL    = PatternFill("solid", fgColor="C6E0B4")
WARN_FILL  = PatternFill("solid", fgColor="FFEB9C")
ERR_FILL   = PatternFill("solid", fgColor="FFC7CE")

def st(cell, *, font=None, fill=None, align=None, fmt=None, border=BORDER):
    if font:   cell.font       = font
    if fill:   cell.fill       = fill
    if align:  cell.alignment  = align
    if fmt:    cell.number_format = fmt
    if border: cell.border     = border

def parse_date(val):
    if val is None: return None
    if isinstance(val, datetime): return val.date()
    s = str(val).strip()[:10]
    for fmt in ('%d.%m.%Y', '%Y-%m-%d'):
        try: return datetime.strptime(s, fmt).date()
        except: pass
    return None

def fmt_date(d):
    return d.strftime('%d.%m.%Y') if d else '—'


# ───────────────────────────────────────────────────────────────
# 1. Parse Doc9 "Сверка" — groups of (bank payment → RKO list)
# Structure: col B = RKO doc, col C = RKO amt,
#            col H = bank amount, col I = bank date
# ───────────────────────────────────────────────────────────────
wb9 = openpyxl.load_workbook(ROOT + r"\ИМВРБ главный с аудитом.xlsx", data_only=True)
ws9 = wb9['Сверка']

groups = []
cur = None

for r in range(3, ws9.max_row + 1):
    rko_dt  = ws9.cell(r, 1).value
    rko_doc = ws9.cell(r, 2).value
    rko_amt = ws9.cell(r, 3).value
    note    = ws9.cell(r, 4).value
    bank_h  = ws9.cell(r, 8).value
    bank_dt = ws9.cell(r, 9).value

    is_bank = bank_h is not None
    try:
        is_bank = is_bank and float(bank_h) > 0
    except (TypeError, ValueError):
        is_bank = False

    if is_bank:
        if cur is not None:
            groups.append(cur)
        cur = {
            'bank_amt': round(float(bank_h), 2),
            'bank_dt':  parse_date(bank_dt),
            'rkos':     [],
            'note':     str(note).strip() if note else ''
        }
        # same row may also have an RKO
        if rko_doc and rko_amt:
            try:
                parts = str(rko_doc).strip().split()
                num   = parts[3] if len(parts) > 3 else str(rko_doc)[:25]
                cur['rkos'].append({'doc': num, 'amt': round(float(rko_amt), 2), 'dt': parse_date(rko_dt)})
            except (TypeError, ValueError):
                pass
    elif cur is not None and rko_doc and rko_amt:
        try:
            parts = str(rko_doc).strip().split()
            num   = parts[3] if len(parts) > 3 else str(rko_doc)[:25]
            cur['rkos'].append({'doc': num, 'amt': round(float(rko_amt), 2), 'dt': parse_date(rko_dt)})
            if note:
                cur['note'] = (cur['note'] + ' ' + str(note)).strip()
        except (TypeError, ValueError):
            pass

if cur is not None:
    groups.append(cur)

# Remove the aggregate row (bank_dt=None, bank_amt=18197.36 total)
groups = [g for g in groups if g['bank_dt'] is not None]
groups.sort(key=lambda g: g['bank_dt'])
print(f"Groups: {len(groups)}, bank_total={sum(g['bank_amt'] for g in groups):.2f}")
print(f"Groups no RKO: {sum(1 for g in groups if not g['rkos'])}")


# ───────────────────────────────────────────────────────────────
# 2. WB weekly totals by month (from totals.json + Doc8 sheet24)
# ───────────────────────────────────────────────────────────────
with open(ROOT + r"\_scripts\totals.json", encoding="utf-8") as f:
    totals = json.load(f)

# Doc7 by_month "total" = "Итого к оплате" per month (both types combined)
wb_by_month = defaultdict(float)

# From Doc7 (2025): read ALL rows (Основной + По выкупам)
wb7 = openpyxl.load_workbook(ROOT + r"\Еженедельный отчет 01.01.2025 - 31.12.2025_4037299.xlsx", data_only=True)
ws7 = wb7.active
for r in range(2, ws7.max_row + 1):
    d_start = ws7.cell(r, 1).value
    total   = ws7.cell(r, 10).value
    if d_start and total:
        d = parse_date(str(d_start)[:10])
        if d:
            wb_by_month[d.strftime('%Y-%m')] += round(float(total), 2)

# From Doc8 sheet24 (2024): ALL rows, cols shifted by 1
wb8 = openpyxl.load_workbook(ROOT + r"\Сверка реализаций по wb 24-25год.xlsx", data_only=True)
ws8_24 = wb8.worksheets[0]
for r in range(2, ws8_24.max_row + 1):
    d_start = ws8_24.cell(r, 2).value   # col 2 (shifted)
    total   = ws8_24.cell(r, 11).value  # col 11 = Итого к оплате
    if d_start and total:
        d = parse_date(str(d_start)[:10])
        if d:
            wb_by_month[d.strftime('%Y-%m')] += round(float(total), 2)

print(f"\nWB by month total: {sum(wb_by_month.values()):.2f}")


# ───────────────────────────────────────────────────────────────
# 3. PKO by month from 1С
#    Old base (Doc5): col 10 = Уменьшение (USD) → × 3.5 → BYN
#    New base (Doc4): col 10 = Уменьшение (BYN)
# ───────────────────────────────────────────────────────────────
pko_by_month = defaultdict(float)

# Old base
wb5 = openpyxl.load_workbook(ROOT + r"\client_ведомость с кассами старая_2.xlsx", data_only=True)
ws5 = wb5.active
for r in range(12, ws5.max_row + 1):
    val = ws5.cell(r, 1).value
    if val and 'Приходный кассовый ордер' in str(val) and 'Касса Влад' in str(val):
        dec = ws5.cell(r, 10).value
        if dec:
            s = str(val)
            dt = None
            if 'от' in s:
                dt = parse_date(s.split('от')[1].strip()[:10])
            if dt:
                pko_by_month[dt.strftime('%Y-%m')] += round(float(dec) * RATE, 2)

# New base
wb4 = openpyxl.load_workbook(ROOT + r"\client_ведомость полная по всем движениям с кассами новая база.xlsx", data_only=True)
ws4 = wb4.active
for r in range(12, ws4.max_row + 1):
    val = ws4.cell(r, 1).value
    if val and 'Приходный кассовый ордер' in str(val) and 'Касса Влад' in str(val):
        dec = ws4.cell(r, 10).value
        if dec:
            s = str(val)
            dt = None
            if 'от' in s:
                dt = parse_date(s.split('от')[1].strip()[:10])
            if dt:
                pko_by_month[dt.strftime('%Y-%m')] += round(float(dec), 2)

print(f"PKO by month total: {sum(pko_by_month.values()):.2f}")

months_all = sorted(set(wb_by_month) | set(pko_by_month))
print("\nMonth    WB total    PKO total   Delta")
for m in months_all:
    wb_t  = round(wb_by_month.get(m, 0), 2)
    pko_t = round(pko_by_month.get(m, 0), 2)
    print(f"  {m}  {wb_t:>10.2f}  {pko_t:>10.2f}  {wb_t-pko_t:>10.2f}")


# ───────────────────────────────────────────────────────────────
# 4. Build sheets in the report
# ───────────────────────────────────────────────────────────────
wb_rep = openpyxl.load_workbook(REPORT)

# ── SHEET 15: Δ1 Банк vs РКО ──────────────────────────────────
for name in ["15. Δ1 Банк vs РКО"]:
    if name in wb_rep.sheetnames:
        del wb_rep[name]
ws15 = wb_rep.create_sheet("15. Δ1 Банк vs РКО")
ws15.sheet_view.showGridLines = False

ws15.merge_cells("A1:H1")
c = ws15.cell(1, 1, "РАСХОЖДЕНИЕ №1 — Банк получил от WB, но с Кассы Влад не списано (РКО отсутствует или частичное)")
c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = CENTER; ws15.row_dimensions[1].height = 30

ws15.merge_cells("A2:H2")
c = ws15.cell(2, 1,
    "Банк всего: 129 410,20 BYN (81 платёж).  РКО Касса Влад: 117 781,69 BYN.  "
    "Δ недосписано = 11 628,51 BYN.  Источник: Dok6 + Dok9 «Сверка».")
c.font = Font(italic=True); c.alignment = WRAP; ws15.row_dimensions[2].height = 28

hdr = ["№", "Дата поступления\nв банк", "Сумма банк,\nBYN",
       "РКО (Касса Влад)\n№ документа", "Дата РКО", "Сумма РКО,\nBYN",
       "Δ (банк−РКО),\nBYN", "Статус / Примечание"]
for j, h in enumerate(hdr):
    c = ws15.cell(3, 1+j, h)
    st(c, font=BOLD, fill=SUB_FILL, align=CENTER)
ws15.row_dimensions[3].height = 36

r = 4
total_bank_sum = 0
total_rko_sum  = 0

for idx, g in enumerate(groups, 1):
    rko_sum  = round(sum(x['amt'] for x in g['rkos']), 2)
    delta    = round(g['bank_amt'] - rko_sum, 2)
    total_bank_sum += g['bank_amt']
    total_rko_sum  += rko_sum

    if not g['rkos']:
        status = "✗ нет РКО"
        fill   = ERR_FILL
    elif abs(delta) < 1.0:
        status = "✓ закрыт"
        fill   = OK_FILL
    elif delta > 0:
        status = f"⚠ недосписано {delta:.2f} BYN"
        fill   = WARN_FILL
    else:
        status = f"⚠ пересписано {abs(delta):.2f} BYN"
        fill   = WARN_FILL

    if g['note']:
        status += "  |  " + g['note'][:60]

    if g['rkos']:
        for k, rko in enumerate(g['rkos']):
            rr = r + k
            if k == 0:
                ws15.cell(rr, 1, idx)
                ws15.cell(rr, 2, fmt_date(g['bank_dt']))
                ws15.cell(rr, 3, g['bank_amt']).number_format = NUM_FORMAT
                ws15.cell(rr, 7, delta).number_format = NUM_FORMAT
                ws15.cell(rr, 8, status).alignment = WRAP
                ws15.cell(rr, 7).fill = fill
            ws15.cell(rr, 4, rko['doc']).alignment = LEFT
            ws15.cell(rr, 5, fmt_date(rko['dt'])).alignment = LEFT
            ws15.cell(rr, 6, rko['amt']).number_format = NUM_FORMAT
            for j in range(1, 9):
                ws15.cell(rr, j).border = BORDER
        r += len(g['rkos'])
    else:
        ws15.cell(r, 1, idx)
        ws15.cell(r, 2, fmt_date(g['bank_dt']))
        ws15.cell(r, 3, g['bank_amt']).number_format = NUM_FORMAT
        ws15.cell(r, 4, "—")
        ws15.cell(r, 5, "—")
        ws15.cell(r, 6, None)
        ws15.cell(r, 7, delta).number_format = NUM_FORMAT
        ws15.cell(r, 7).fill = fill
        ws15.cell(r, 8, status).alignment = WRAP
        for j in range(1, 9):
            ws15.cell(r, j).border = BORDER
        r += 1

# Totals row
ws15.cell(r, 2, "ИТОГО").font = BOLD
ws15.cell(r, 3, round(total_bank_sum, 2)).number_format = NUM_FORMAT
ws15.cell(r, 3).font = BOLD
ws15.cell(r, 6, round(total_rko_sum, 2)).number_format = NUM_FORMAT
ws15.cell(r, 6).font = BOLD
delta_total = round(total_bank_sum - total_rko_sum, 2)
ws15.cell(r, 7, delta_total).number_format = NUM_FORMAT
ws15.cell(r, 7).font = BOLD
ws15.cell(r, 7).fill = ERR_FILL
for j in range(1, 9):
    ws15.cell(r, j).border = BORDER

ws15.column_dimensions["A"].width = 4
ws15.column_dimensions["B"].width = 14
ws15.column_dimensions["C"].width = 14
ws15.column_dimensions["D"].width = 24
ws15.column_dimensions["E"].width = 14
ws15.column_dimensions["F"].width = 14
ws15.column_dimensions["G"].width = 18
ws15.column_dimensions["H"].width = 44
ws15.freeze_panes = "A4"

print(f"\nSheet 15 rows written: {r-4}")


# ── SHEET 16: Δ2 WB vs ПКО помесячно ─────────────────────────
for name in ["16. Δ2 WB vs ПКО"]:
    if name in wb_rep.sheetnames:
        del wb_rep[name]
ws16 = wb_rep.create_sheet("16. Δ2 WB vs ПКО")
ws16.sheet_view.showGridLines = False

ws16.merge_cells("A1:F1")
c = ws16.cell(1, 1, "РАСХОЖДЕНИЕ №2 — WB начислил к оплате, но ПКО в Касса Влад не введены (сравнение помесячно)")
c.font = TITLE_FONT; c.fill = TITLE_FILL; c.alignment = CENTER; ws16.row_dimensions[1].height = 30

ws16.merge_cells("A2:F2")
c = ws16.cell(2, 1,
    "WB всего (Dok7+Dok8): 134 781,68 BYN.  ПКО Касса Влад (Dok4+Dok5×3,5): 115 673,80 BYN.  "
    "Δ = 19 107,88 BYN ПКО не добиты.  ⚠ ПКО могут вводиться с задержкой 1–4 нед. после отчётной недели WB.")
c.font = Font(italic=True); c.alignment = WRAP; ws16.row_dimensions[2].height = 40

hdr16 = ["Месяц", "WB «к оплате»,\nBYN", "ПКО в 1С\n(Касса Влад), BYN",
         "Δ (WB−ПКО),\nBYN", "Статус", "Источник / Примечание"]
for j, h in enumerate(hdr16):
    c = ws16.cell(3, 1+j, h)
    st(c, font=BOLD, fill=SUB_FILL, align=CENTER)
ws16.row_dimensions[3].height = 36

r16 = 4
total_wb_sum  = 0
total_pko_sum = 0

for m in months_all:
    wb_t  = round(wb_by_month.get(m,  0), 2)
    pko_t = round(pko_by_month.get(m, 0), 2)
    delta = round(wb_t - pko_t, 2)
    if wb_t == 0 and pko_t == 0:
        continue
    total_wb_sum  += wb_t
    total_pko_sum += pko_t

    if wb_t == 0:
        status = "нет данных WB"
        fill   = None
    elif pko_t == 0:
        status = "✗ ПКО не введены"
        fill   = ERR_FILL
    elif abs(delta) < 100:
        status = "✓ совпадает (±100)"
        fill   = OK_FILL
    elif delta > 0:
        status = f"⚠ недобито {delta:.2f}"
        fill   = WARN_FILL
    else:
        status = f"ПКО > WB на {abs(delta):.2f}"
        fill   = WARN_FILL

    yr = int(m[:4])
    mo = int(m[5:])
    if m < '2025-10':
        src = "WB: Dok8 (24 г.) / Dok7 (25 г.).  1С: Dok5 старая база (USD×3,5)"
    elif m < '2026-01':
        src = "WB: Dok7.  1С: Dok4 новая база (BYN)"
    else:
        src = "WB: нет данных за 26 г.  1С: Dok4 новая база (BYN)"

    ws16.cell(r16, 1, m)
    ws16.cell(r16, 2, wb_t  if wb_t  > 0 else None)
    ws16.cell(r16, 2).number_format = NUM_FORMAT
    ws16.cell(r16, 3, pko_t if pko_t > 0 else None)
    ws16.cell(r16, 3).number_format = NUM_FORMAT
    ws16.cell(r16, 4, delta)
    ws16.cell(r16, 4).number_format = NUM_FORMAT
    ws16.cell(r16, 5, status)
    ws16.cell(r16, 6, src)

    for j in range(1, 7):
        ws16.cell(r16, j).border = BORDER
        ws16.cell(r16, j).alignment = LEFT
    if fill:
        ws16.cell(r16, 4).fill = fill
        ws16.cell(r16, 5).fill = fill
    r16 += 1

# Totals
ws16.cell(r16, 1, "ИТОГО").font = BOLD
ws16.cell(r16, 2, round(total_wb_sum,  2)).number_format = NUM_FORMAT
ws16.cell(r16, 2).font = BOLD
ws16.cell(r16, 3, round(total_pko_sum, 2)).number_format = NUM_FORMAT
ws16.cell(r16, 3).font = BOLD
ws16.cell(r16, 4, round(total_wb_sum - total_pko_sum, 2)).number_format = NUM_FORMAT
ws16.cell(r16, 4).font = BOLD
ws16.cell(r16, 4).fill = ERR_FILL
for j in range(1, 7):
    ws16.cell(r16, j).border = BORDER

ws16.merge_cells(f"A{r16+2}:F{r16+2}")
note_cell = ws16.cell(r16+2, 1,
    "Примечание: сравнение по дате ввода ПКО vs дате отчётной недели WB — возможно смещение до 3–4 недель. "
    "Месяцы с «ПКО > WB» означают, что ПКО за прошлые периоды введены позже. "
    "Для точной сверки нужно сопоставлять по номеру отчёта WB (4037299/XX) с номером ПКО.")
note_cell.font = Font(italic=True, color="993300")
note_cell.alignment = WRAP
ws16.row_dimensions[r16+2].height = 56

ws16.column_dimensions["A"].width = 12
ws16.column_dimensions["B"].width = 18
ws16.column_dimensions["C"].width = 20
ws16.column_dimensions["D"].width = 18
ws16.column_dimensions["E"].width = 24
ws16.column_dimensions["F"].width = 50
ws16.freeze_panes = "A4"

print(f"Sheet 16 rows written: {r16-4}")

wb_rep.save(REPORT)
print("\nSaved:", REPORT)
print("Sheets:", [s.title for s in wb_rep.worksheets])
