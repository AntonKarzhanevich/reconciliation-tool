"""Extract precise totals & key figures from all 10 source files."""
import os
import json
import openpyxl
from datetime import datetime, date

ROOT = r"D:\Cursor\kassa"
EXCEL_REVIEWED = os.path.join(ROOT, "data", "excel-reviewed")


def xlsx_path(filename: str) -> str:
    """Исходные .xlsx: приоритет data/excel-reviewed/, иначе корень проекта."""
    p = os.path.join(EXCEL_REVIEWED, filename)
    return p if os.path.isfile(p) else os.path.join(ROOT, filename)


def parse_dt(v):
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    if isinstance(v, str):
        try: return datetime.fromisoformat(v.replace("Z","+00:00")).date()
        except Exception:
            for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                try: return datetime.strptime(v[:10], fmt).date()
                except Exception: continue
    return None

def numlike(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    if isinstance(v, str):
        try: return float(v.replace(" ", "").replace(",", "."))
        except Exception: return None
    return None

result = {}


# ============= DOC 5 (1С Старая_2, USD) =============
def doc5():
    fp = xlsx_path( "client_ведомость с кассами старая_2.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb["Лист_1"]
    info = {"path": fp, "sheet": "Лист_1", "period_header": ws.cell(4,3).value,
            "currency": "USD"}
    # KAZAP summary row 10: Cols I=Увеличение, J=Уменьшение, K=Кон.долг клиента, L=Кон.наш долг, M=Сальдо
    info["KAZAP"] = {
        "Увеличение_USD": ws.cell(10, 9).value,
        "Уменьшение_USD": ws.cell(10,10).value,
        "Кон_долг_клиента_USD": ws.cell(10,11).value,
        "Кон_наш_долг_USD": ws.cell(10,12).value,
        "Сальдо_USD": ws.cell(10,13).value,
    }
    # Иmvbrb row 11
    info["ИМВБРБ"] = {
        "Увеличение_USD": ws.cell(11, 9).value,
        "Уменьшение_USD": ws.cell(11,10).value,
        "Сальдо_USD": ws.cell(11,13).value,
    }
    # Count cassa types and doc types in rows 12..end
    by_cassa = {}
    by_doctype = {}
    by_year = {}
    real_min, real_max = None, None
    inc_usd = 0; dec_usd = 0
    for r in range(12, ws.max_row+1):
        txt = ws.cell(r,1).value
        if not isinstance(txt, str): continue
        inc = numlike(ws.cell(r,9).value) or 0
        dec = numlike(ws.cell(r,10).value) or 0
        inc_usd += inc; dec_usd += dec
        # extract date from "от dd.mm.yyyy"
        idx = txt.find(" от ")
        d = None
        if idx >= 0:
            try: d = datetime.strptime(txt[idx+4:idx+14], "%d.%m.%Y").date()
            except Exception: pass
        if d:
            if real_min is None or d < real_min: real_min = d
            if real_max is None or d > real_max: real_max = d
            y = d.year
            by_year[y] = by_year.get(y, {"inc":0, "dec":0, "rows":0})
            by_year[y]["inc"] += inc; by_year[y]["dec"] += dec; by_year[y]["rows"] += 1
        # doc type
        dtype = txt.split(" 0000-")[0].split(" от ")[0].strip() if (" 0000-" in txt or " от " in txt) else txt.strip()
        by_doctype[dtype] = by_doctype.get(dtype, 0) + 1
        # cassa (after last comma)
        if "," in txt:
            cash = txt.rsplit(",", 1)[-1].strip()
            if cash:
                by_cassa[cash] = by_cassa.get(cash, 0) + 1
    info["row_data_range"] = {"min": str(real_min), "max": str(real_max)}
    info["sum_increase_USD"] = inc_usd
    info["sum_decrease_USD"] = dec_usd
    info["by_year"] = {y: {**v, "inc":v["inc"], "dec":v["dec"]} for y,v in by_year.items()}
    info["by_cassa"] = by_cassa
    info["by_doctype"] = by_doctype
    # KAZAP totals × 3.5
    s = info["KAZAP"]
    info["KAZAP_BYN_x3.5"] = {k.replace("_USD","_BYN"): (v or 0)*3.5 for k,v in s.items()}
    return info

result["Док 5"] = doc5()


# ============= DOC 4 (1С новая полная, BYN) =============
def doc4():
    fp = xlsx_path( "client_ведомость полная по всем движениям с кассами новая база.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb["Лист_1"]
    info = {"path": fp, "sheet": "Лист_1",
            "period_header": ws.cell(4,3).value, "currency": "BYN"}
    # row 10 KAZAP, cols 9..13
    info["KAZAP"] = {
        "Увеличение_BYN": ws.cell(10, 9).value,
        "Уменьшение_BYN": ws.cell(10,10).value,
        "Кон_долг_клиента_BYN": ws.cell(10,11).value,
        "Кон_наш_долг_BYN": ws.cell(10,12).value,
        "Сальдо_BYN": ws.cell(10,13).value,
    }
    # find actual end date from registrar text
    by_cassa = {}
    by_doctype = {}
    by_year = {}
    real_min, real_max = None, None
    inc_b = 0; dec_b = 0
    for r in range(12, ws.max_row+1):
        txt = ws.cell(r,1).value
        if not isinstance(txt, str): continue
        inc = numlike(ws.cell(r,9).value) or 0
        dec = numlike(ws.cell(r,10).value) or 0
        inc_b += inc; dec_b += dec
        idx = txt.find(" от ")
        d = None
        if idx >= 0:
            try: d = datetime.strptime(txt[idx+4:idx+14], "%d.%m.%Y").date()
            except Exception: pass
        if d:
            if real_min is None or d < real_min: real_min = d
            if real_max is None or d > real_max: real_max = d
            y = d.year
            by_year[y] = by_year.get(y, {"inc":0, "dec":0, "rows":0})
            by_year[y]["inc"] += inc; by_year[y]["dec"] += dec; by_year[y]["rows"] += 1
        dtype = txt.split(" 0000-")[0].split(" от ")[0].strip() if (" 0000-" in txt or " от " in txt) else txt.strip()
        by_doctype[dtype] = by_doctype.get(dtype, 0) + 1
        if "," in txt:
            cash = txt.rsplit(",", 1)[-1].strip()
            if cash:
                by_cassa[cash] = by_cassa.get(cash, 0) + 1
    info["row_data_range"] = {"min": str(real_min), "max": str(real_max)}
    info["sum_increase_BYN"] = inc_b
    info["sum_decrease_BYN"] = dec_b
    info["by_year"] = by_year
    info["by_cassa"] = by_cassa
    info["by_doctype"] = by_doctype
    return info

result["Док 4"] = doc4()


# ============= DOC 1 (client_ведомость с кассами новая база) =============
def doc1():
    fp = xlsx_path( "client_ведомость с кассами новая база.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    sheets = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        # try to find row 10 KAZAP totals
        kazap = {}
        # Need to detect column structure. Try cols 4..8 for cash report, 9..13 for settlement
        cols_settlement = [9,10,11,12,13]
        cols_cash = [2,3,4,5]
        # Detect from row 8
        row10 = [ws.cell(10,c).value for c in range(1, ws.max_column+1)]
        if ws.max_column >= 13:
            kazap["model"] = "settlement"
            kazap["Увеличение"] = ws.cell(10, 9).value
            kazap["Уменьшение"] = ws.cell(10,10).value
            kazap["Кон_долг_клиента"] = ws.cell(10,11).value
            kazap["Кон_наш_долг"] = ws.cell(10,12).value
            kazap["Сальдо"] = ws.cell(10,13).value
        else:
            kazap["model"] = "cash"
            # row 10 in such sheet often: KAZAP | None | inc | dec | bal
            # Need to find numeric cells; we'll grab first 4-5 numeric trail
            vals_num = [v for v in row10[1:] if isinstance(v,(int,float))]
            kazap["Увеличение"] = vals_num[0] if len(vals_num)>0 else None
            kazap["Уменьшение"] = vals_num[1] if len(vals_num)>1 else None
            kazap["Кон_долг_клиента"] = vals_num[2] if len(vals_num)>2 else None
            kazap["Кон_наш_долг"] = vals_num[3] if len(vals_num)>3 else None
        # collect cassa & date stats
        by_cassa = {}
        real_min, real_max = None, None
        for r in range(12, ws.max_row+1):
            txt = ws.cell(r,1).value
            if not isinstance(txt, str): continue
            idx = txt.find(" от ")
            if idx >= 0:
                try:
                    d = datetime.strptime(txt[idx+4:idx+14], "%d.%m.%Y").date()
                    if real_min is None or d < real_min: real_min = d
                    if real_max is None or d > real_max: real_max = d
                except Exception: pass
            if "," in txt:
                cash = txt.rsplit(",", 1)[-1].strip()
                if cash:
                    by_cassa[cash] = by_cassa.get(cash, 0) + 1
        # final cell (last cell value)
        last_row_vals = None
        for r in range(ws.max_row, 0, -1):
            vals = [ws.cell(r,c).value for c in range(1, ws.max_column+1)]
            if any(v not in (None,"") for v in vals):
                last_row_vals = vals
                break
        sheets[sn] = {
            "rows": ws.max_row, "cols": ws.max_column,
            "row4_param": ws.cell(4,3).value, "row5_param": ws.cell(5,3).value,
            "KAZAP_row10": kazap,
            "data_date_min": str(real_min), "data_date_max": str(real_max),
            "by_cassa": by_cassa,
            "last_nonempty_row": last_row_vals[:8] if last_row_vals else None,
        }
    return {"path": fp, "sheets": sheets}

result["Док 1"] = doc1()


# ============= DOC 3 (Старая Влад и ФО приходники/расходники) =============
def doc3():
    fp = xlsx_path( "client_ведомость с кассами старая влад и фо расходники и приходники.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb["Лист_1"]
    info = {"path": fp, "sheet": "Лист_1",
            "param_period": ws.cell(2,2).value,
            "param_otbor": ws.cell(4,1).value,
            "header_row6": [ws.cell(6,c).value for c in range(1,ws.max_column+1)],
            "header_row7": [ws.cell(7,c).value for c in range(1,ws.max_column+1)],
            "kazap_row8": [ws.cell(8,c).value for c in range(1,ws.max_column+1)],
            "currency": None,
            }
    # last row totals
    last_real = None
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(r,1).value
        if isinstance(v, str) and "Итого" in v:
            last_real = r; break
    if last_real:
        info["row_itogo"] = last_real
        info["row_itogo_vals"] = [ws.cell(last_real,c).value for c in range(1, ws.max_column+1)]
    # Count касса by tails in column 1 (where касса is sometimes mentioned)
    by_cassa = {}
    real_min, real_max = None, None
    pos_count = 0; neg_count = 0
    for r in range(7, ws.max_row+1):
        txt = ws.cell(r,1).value
        if not isinstance(txt, str): continue
        idx = txt.find(" от ")
        d = None
        if idx >= 0:
            try:
                d = datetime.strptime(txt[idx+4:idx+14], "%d.%m.%Y").date()
                if real_min is None or d < real_min: real_min = d
                if real_max is None or d > real_max: real_max = d
            except Exception: pass
        bal = numlike(ws.cell(r,5).value)
        if bal is not None:
            if bal < 0: neg_count += 1
            else: pos_count += 1
    info["data_date_min"] = str(real_min)
    info["data_date_max"] = str(real_max)
    info["balance_neg_rows"] = neg_count
    info["balance_pos_rows"] = pos_count
    return info

result["Док 3"] = doc3()


# ============= DOC 6 (client_приходы банк) =============
def doc6():
    fp = xlsx_path( "client_приходы банк за всё время.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb["Лист1"]
    sum_b = 0; cnt = 0
    by_year = {}
    dates = []
    for r in range(2, ws.max_row+1):
        sm = numlike(ws.cell(r,2).value)
        dt = parse_dt(ws.cell(r,4).value)
        if sm is None or dt is None: continue
        sum_b += sm; cnt += 1
        dates.append(dt)
        y = dt.year
        by_year[y] = by_year.get(y, {"sum":0, "cnt":0})
        by_year[y]["sum"] += sm; by_year[y]["cnt"] += 1
    return {"path": fp, "sheet": "Лист1",
            "rows_total": ws.max_row, "rows_with_data": cnt,
            "sum_BYN": sum_b, "date_min": str(min(dates)) if dates else None,
            "date_max": str(max(dates)) if dates else None,
            "by_year": by_year}

result["Док 6"] = doc6()


# ============= DOC 7 (Еженедельный 2025) =============
def doc7():
    fp = xlsx_path( "Еженедельный отчет 01.01.2025 - 31.12.2025_4037299.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb["Sheet1"]
    sums = {"sale":0,"pay":0,"log":0,"stor":0,"fines":0,"total":0}
    cnt = 0
    by_type = {}
    by_month = {}
    dates = []
    for r in range(2, ws.max_row+1):
        d1 = parse_dt(ws.cell(r,1).value)
        t = ws.cell(r,4).value
        sale  = numlike(ws.cell(r,5).value) or 0
        pay   = numlike(ws.cell(r,6).value) or 0
        log_  = numlike(ws.cell(r,7).value) or 0
        stor  = numlike(ws.cell(r,8).value) or 0
        fines = numlike(ws.cell(r,9).value) or 0
        total = numlike(ws.cell(r,10).value) or 0
        if not any((sale, pay, total)): continue
        cnt += 1
        sums["sale"]+=sale; sums["pay"]+=pay; sums["log"]+=log_
        sums["stor"]+=stor; sums["fines"]+=fines; sums["total"]+=total
        by_type[str(t)] = by_type.get(str(t),0)+1
        if d1:
            dates.append(d1)
            ym = f"{d1.year}-{d1.month:02d}"
            by_month[ym] = by_month.get(ym, {"sale":0,"pay":0,"total":0,"cnt":0})
            for k,v in (("sale",sale),("pay",pay),("total",total)):
                by_month[ym][k] += v
            by_month[ym]["cnt"] += 1
    return {"path": fp, "sheet": "Sheet1",
            "rows_total": ws.max_row, "rows_with_data": cnt,
            "sums_BYN": sums, "types": by_type,
            "date_min": str(min(dates)) if dates else None,
            "date_max": str(max(dates)) if dates else None,
            "by_month": by_month}

result["Док 7"] = doc7()


# ============= DOC 8 (Сверка реализаций 24-25) =============
def doc8():
    fp = xlsx_path( "Сверка реализаций по wb 24-25год.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    out = {"path": fp, "sheets": {}}
    for sn in wb.sheetnames:
        ws = wb[sn]
        is_24 = "24" in sn
        # col offsets
        if is_24:
            cs,ce,ct,cv1,cv2,cv3,cv4,cv5,cv6 = 2,3,5,6,7,8,9,10,11
        else:
            cs,ce,ct,cv1,cv2,cv3,cv4,cv5,cv6 = 1,2,4,5,6,7,8,9,10
        sums = {"sale":0,"pay":0,"log":0,"stor":0,"fines":0,"total":0}
        cnt=0; by_type={}; by_month={}; dates=[]
        for r in range(2, ws.max_row+1):
            d1 = parse_dt(ws.cell(r,cs).value)
            t  = ws.cell(r,ct).value
            sale  = numlike(ws.cell(r,cv1).value) or 0
            pay   = numlike(ws.cell(r,cv2).value) or 0
            log_  = numlike(ws.cell(r,cv3).value) or 0
            stor  = numlike(ws.cell(r,cv4).value) or 0
            fines = numlike(ws.cell(r,cv5).value) or 0
            total = numlike(ws.cell(r,cv6).value) or 0
            if not any((sale, pay, total)): continue
            cnt += 1
            sums["sale"]+=sale; sums["pay"]+=pay; sums["log"]+=log_
            sums["stor"]+=stor; sums["fines"]+=fines; sums["total"]+=total
            by_type[str(t)] = by_type.get(str(t),0)+1
            if d1:
                dates.append(d1)
                ym = f"{d1.year}-{d1.month:02d}"
                by_month[ym] = by_month.get(ym, {"sale":0,"pay":0,"total":0,"cnt":0})
                for k,v in (("sale",sale),("pay",pay),("total",total)):
                    by_month[ym][k] += v
                by_month[ym]["cnt"] += 1
        # also slice 25 sheet by 2025-10-03 boundary
        boundary = date(2025,10,3)
        out["sheets"][sn] = {
            "rows_with_data": cnt, "sums_BYN": sums, "types": by_type,
            "date_min": str(min(dates)) if dates else None,
            "date_max": str(max(dates)) if dates else None,
            "by_month": by_month,
        }
    return out

result["Док 8"] = doc8()


# ============= DOC 8 (sliced phase A/B) =============
def doc8_slice():
    fp = xlsx_path( "Сверка реализаций по wb 24-25год.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    boundary = date(2025,10,3)
    out = {"phase_A":{"sale":0,"pay":0,"log":0,"stor":0,"fines":0,"total":0,"rows":0},
           "phase_B":{"sale":0,"pay":0,"log":0,"stor":0,"fines":0,"total":0,"rows":0}}
    for sn in wb.sheetnames:
        ws = wb[sn]
        is_24 = "24" in sn
        if is_24:
            cs,ce,ct,cv1,cv2,cv3,cv4,cv5,cv6 = 2,3,5,6,7,8,9,10,11
        else:
            cs,ce,ct,cv1,cv2,cv3,cv4,cv5,cv6 = 1,2,4,5,6,7,8,9,10
        for r in range(2, ws.max_row+1):
            d1 = parse_dt(ws.cell(r,cs).value)
            if d1 is None: continue
            phase = "phase_A" if d1 < boundary else "phase_B"
            for k,c in (("sale",cv1),("pay",cv2),("log",cv3),("stor",cv4),("fines",cv5),("total",cv6)):
                v = numlike(ws.cell(r,c).value) or 0
                out[phase][k] += v
            out[phase]["rows"] += 1
    return out

result["Док 8 slice"] = doc8_slice()


# ============= DOC 6 sliced =============
def doc6_slice():
    fp = xlsx_path( "client_приходы банк за всё время.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb["Лист1"]
    boundary = date(2025,10,3)
    out = {"phase_A":{"sum":0,"cnt":0}, "phase_B":{"sum":0,"cnt":0}}
    for r in range(2, ws.max_row+1):
        sm = numlike(ws.cell(r,2).value)
        dt = parse_dt(ws.cell(r,4).value)
        if sm is None or dt is None: continue
        phase = "phase_A" if dt < boundary else "phase_B"
        out[phase]["sum"] += sm
        out[phase]["cnt"] += 1
    return out

result["Док 6 slice"] = doc6_slice()


# ============= DOC 9 (главный с аудитом) =============
def doc9():
    fp = xlsx_path( "ИМВРБ главный с аудитом.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    info = {"path": fp, "sheets": {}}
    for sn in wb.sheetnames:
        ws = wb[sn]
        # collect any numeric cells in last 5 rows
        tails = []
        for r in range(max(1, ws.max_row-5), ws.max_row+1):
            row_vals = [ws.cell(r,c).value for c in range(1, ws.max_column+1)]
            if any(v not in (None,"") for v in row_vals):
                tails.append((r, row_vals))
        info["sheets"][sn] = {
            "rows": ws.max_row, "cols": ws.max_column,
            "tail": tails[-5:]
        }
    return info

result["Док 9"] = doc9()


# ============= DOC 10 (итог аудит) =============
def doc10():
    fp = xlsx_path( "Итог аудит.xlsx")
    wb = openpyxl.load_workbook(fp, data_only=True)
    info = {"path": fp, "sheets": {}}
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = []
        for r in range(1, ws.max_row+1):
            row_vals = [ws.cell(r,c).value for c in range(1, ws.max_column+1)]
            if any(v not in (None,"") for v in row_vals):
                rows.append((r, row_vals))
        info["sheets"][sn] = {
            "rows_total": ws.max_row, "cols": ws.max_column,
            "non_empty_rows": len(rows),
            "all_data": rows,
        }
    return info

result["Док 10"] = doc10()


# Dump
out_fp = os.path.join(ROOT, "_scripts", "totals.json")
with open(out_fp, "w", encoding="utf-8") as f:
    def default(o):
        if isinstance(o, (date, datetime)): return str(o)
        if isinstance(o, bytes): return o.decode("utf-8", "replace")
        return str(o)
    json.dump(result, f, ensure_ascii=False, indent=2, default=default)
print(f"Saved to {out_fp}")
print(f"Size: {os.path.getsize(out_fp):,} bytes")
