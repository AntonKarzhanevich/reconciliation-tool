"""
Match bank payments (Doc6) to RKO (Doc9 kasса vlad) and WB weekly (Doc7/8) to PKO (Doc4/5).
Output: adds two sheets to the report showing specific unmatched documents.
"""
import openpyxl
from datetime import datetime, timedelta

ROOT = r"D:\Cursor\kassa"

def parse_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    s = str(val).strip()[:10]
    for fmt in ('%d.%m.%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except:
            pass
    return None


# ========== DOC 6: bank payments ==========
wb6 = openpyxl.load_workbook(ROOT + r"\client_приходы банк за всё время.xlsx", data_only=True)
ws6 = wb6.active
bank = []
for r in range(2, ws6.max_row + 1):
    id_ = ws6.cell(r, 1).value
    amt = ws6.cell(r, 2).value
    dt  = ws6.cell(r, 4).value
    if id_ and amt:
        bank.append({'id': str(id_), 'amt': round(float(amt), 2), 'date': parse_date(dt)})
bank.sort(key=lambda x: x['date'] or datetime(2000, 1, 1).date())


# ========== DOC 9 kasса vlad: RKO ==========
wb9 = openpyxl.load_workbook(ROOT + r"\ИМВРБ главный с аудитом.xlsx", data_only=True)
ws9 = wb9['касса влад']
rkos = []
for r in range(2, ws9.max_row + 1):
    dt_raw = ws9.cell(r, 1).value
    doc    = ws9.cell(r, 2).value
    amt    = ws9.cell(r, 3).value
    comm   = ws9.cell(r, 4).value
    if doc and amt:
        doc_str = str(doc).strip()
        # extract short number like 0000-В01199
        parts = doc_str.split()
        num = parts[3] if len(parts) > 3 else doc_str[:30]
        rkos.append({
            'dt': parse_date(dt_raw),
            'doc': num,
            'doc_full': doc_str,
            'amt': round(float(amt), 2),
            'comm': str(comm) if comm else ''
        })
rkos.sort(key=lambda x: x['dt'] or datetime(2000, 1, 1).date())

print(f"Bank: {len(bank)}, total={sum(x['amt'] for x in bank):.2f}")
print(f"RKO: {len(rkos)}, total={sum(x['amt'] for x in rkos):.2f}")


# ========== MATCHING: bank → RKO ==========
WINDOW = 7   # days
TOL    = 1.0  # BYN tolerance

used_rkos = set()
matched = []
unmatched_bank = []

for b in bank:
    if b['date'] is None:
        unmatched_bank.append({**b, 'note': 'нет даты'})
        continue

    cands = [i for i, r in enumerate(rkos)
             if i not in used_rkos and r['dt'] and abs((r['dt'] - b['date']).days) <= WINDOW]

    # single exact match
    exact = [i for i in cands if abs(rkos[i]['amt'] - b['amt']) < TOL]
    if exact:
        best = min(exact, key=lambda i: abs(rkos[i]['amt'] - b['amt']))
        used_rkos.add(best)
        matched.append({'bank': b, 'rko_docs': [rkos[best]['doc']], 'rko_sum': rkos[best]['amt'],
                        'delta': round(b['amt'] - rkos[best]['amt'], 2), 'rko_dates': [str(rkos[best]['dt'])]})
        continue

    # pair match
    found = False
    for ii in range(len(cands)):
        for jj in range(ii + 1, len(cands)):
            i, j = cands[ii], cands[jj]
            s = rkos[i]['amt'] + rkos[j]['amt']
            if abs(s - b['amt']) < TOL:
                used_rkos.add(i)
                used_rkos.add(j)
                matched.append({'bank': b, 'rko_docs': [rkos[i]['doc'], rkos[j]['doc']],
                                'rko_sum': round(s, 2), 'delta': round(b['amt'] - s, 2),
                                'rko_dates': [str(rkos[i]['dt']), str(rkos[j]['dt'])]})
                found = True
                break
        if found:
            break

    if not found:
        # find closest single RKO by date for context
        near = sorted(cands, key=lambda i: abs(rkos[i]['amt'] - b['amt']))
        note_doc = rkos[near[0]]['doc'] if near else '—'
        note_amt = rkos[near[0]]['amt'] if near else 0
        unmatched_bank.append({
            **b,
            'note': f'Ближайший РКО: {note_doc} ({note_amt:.2f} BYN)'
        })

unmatched_rko = [r for i, r in enumerate(rkos) if i not in used_rkos]

print(f"\nMatched: {len(matched)}, unmatched bank: {len(unmatched_bank)}, unmatched RKO: {len(unmatched_rko)}")
print(f"Unmatched bank sum: {sum(x['amt'] for x in unmatched_bank):.2f}")
print(f"Unmatched RKO sum: {sum(x['amt'] for x in unmatched_rko):.2f}")

print("\nUnmatched bank payments (нет РКО):")
for x in unmatched_bank:
    print(f"  {x['date']}  {x['id']}  {x['amt']} BYN  | {x['note']}")

print("\nUnmatched RKOs (нет банка):")
for x in unmatched_rko[:20]:
    print(f"  {x['dt']}  {x['doc']}  {x['amt']} BYN")


# ========== DOC 7: WB weekly + Doc 4 PKO ==========
wb7 = openpyxl.load_workbook(ROOT + r"\Еженедельный отчет 01.01.2025 - 31.12.2025_4037299.xlsx", data_only=True)
ws7 = wb7.active
wb_weeks_25 = []
for r in range(2, ws7.max_row + 1):
    d_start = ws7.cell(r, 1).value
    d_end   = ws7.cell(r, 2).value
    type_   = ws7.cell(r, 4).value
    total   = ws7.cell(r, 10).value
    if d_start and total and type_ == 'Основной':
        wb_weeks_25.append({
            'date_start': parse_date(str(d_start)[:10]),
            'date_end':   parse_date(str(d_end)[:10]),
            'total': round(float(total), 2)
        })

# Doc 8 sheet 24
wb8 = openpyxl.load_workbook(ROOT + r"\Сверка реализаций по wb 24-25год.xlsx", data_only=True)
ws8_24 = wb8.worksheets[0]
wb_weeks_24 = []
for r in range(2, ws8_24.max_row + 1):
    d_start = ws8_24.cell(r, 1).value
    d_end   = ws8_24.cell(r, 2).value
    type_   = ws8_24.cell(r, 4).value
    total   = ws8_24.cell(r, 10).value
    if d_start and total and str(type_) == 'Основной':
        wb_weeks_24.append({
            'date_start': parse_date(str(d_start)[:10]),
            'date_end':   parse_date(str(d_end)[:10]),
            'total': round(float(total), 2)
        })

all_wb_weeks = wb_weeks_24 + wb_weeks_25
all_wb_weeks.sort(key=lambda x: x['date_start'] or datetime(2000,1,1).date())
print(f"\nWB weeks: {len(all_wb_weeks)}, total={sum(x['total'] for x in all_wb_weeks):.2f}")

# PKO from Doc 4 (new base) - from the data we know format is row 12+ with "Приходный кассовый ордер"
pkos_new = []
wb4 = openpyxl.load_workbook(ROOT + r"\client_ведомость полная по всем движениям с кассами новая база.xlsx", data_only=True)
ws4 = wb4.active
for r in range(12, ws4.max_row + 1):
    val = ws4.cell(r, 1).value
    if val and 'Приходный кассовый ордер' in str(val) and 'Влад' in str(val):
        dec = ws4.cell(r, 10).value  # decrease = payment received
        dt_str = str(val).split('от')[1].strip().split(',')[0].strip() if 'от' in str(val) else ''
        dt = None
        if dt_str:
            try:
                dt = datetime.strptime(dt_str[:10], '%d.%m.%Y').date()
            except:
                try:
                    dt = datetime.strptime(dt_str[:10], '%Y-%m-%d').date()
                except:
                    pass
        if dec:
            pkos_new.append({'dt': dt, 'doc': str(val).strip()[:80], 'amt': round(float(dec), 2)})

print(f"PKO new base (Влад): {len(pkos_new)}, total={sum(x['amt'] for x in pkos_new):.2f}")
print("First 3:", pkos_new[:3])
