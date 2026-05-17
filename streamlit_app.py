import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from collections import defaultdict
import datetime, io

st.set_page_config(page_title="QC Monitor", page_icon="🔍", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
.block-container{padding-top:1.5rem}
div[data-testid="metric-container"]{background:#0f1117;border:1px solid #1c2030;border-radius:8px;padding:12px 16px}
div[data-testid="metric-container"] label{font-size:11px!important;color:#7c879e!important;text-transform:uppercase;letter-spacing:.5px}
</style>""", unsafe_allow_html=True)

# ─── HELPERS ─────────────────────────────────────────────────────────────────

SKIP_ROWS = {'Total','Total Panel Dédupliqué','Total Panel Dédupliqué - Top 5 Sites',
    'Total Panel Dédupliqué  - Top 11 Sites','Total Panel Dédupliqué Marché',
    'Immobilier Notaire','Immonot','Site','Département','Totaux','Total Panel Dedup'}

SITES = ['AvendreAlouer',"Bien'ici",'Figaro Immo','Green-Acres','Leboncoin',
         'LogicImmo','MeilleursAgents','OuestFrance','PAP','ParuVendu','SeLoger','SuperImmo']

def excel_date_str(v):
    if isinstance(v, datetime.datetime): return v.strftime('%b-%y')
    if isinstance(v, str): return v.strip()
    if isinstance(v,(int,float)) and 40000<v<50000:
        return (datetime.datetime(1899,12,30)+datetime.timedelta(days=int(v))).strftime('%b-%y')
    return str(v)

def get_sheet_data(ws, site_col=2, start_col=3):
    """Extract site data + months from a sheet."""
    header_row = None
    for r in range(1, min(20, ws.max_row+1)):
        v = ws.cell(r, site_col).value
        if v in ('Site','Département','Région'):
            header_row = r; break
    if not header_row: return None, None, None

    month_cols, months = [], []
    for c in range(start_col, ws.max_column+1):
        v = ws.cell(header_row, c).value
        if v is None: continue
        if (isinstance(v, datetime.datetime) or
            (isinstance(v,(int,float)) and 40000<v<50000) or
            (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
            month_cols.append(c); months.append(excel_date_str(v))
    if len(months)<2: return None, None, None

    data = {}
    for r in range(header_row+1, ws.max_row+1):
        site = ws.cell(r, site_col).value
        if not site or not isinstance(site,str): continue
        site = site.strip()
        if not site: continue
        vm = ws.cell(r, month_cols[-1]).value
        vm1 = ws.cell(r, month_cols[-2]).value
        if isinstance(vm,(int,float)) and isinstance(vm1,(int,float)):
            data[site] = {'m':float(vm),'m1':float(vm1)}
    return data, months[-1], months[-2]

def check_ok(name, condition, detail, severity="error", sheet=""):
    return {'check':name,'status':'ok' if condition else 'fail',
            'severity':severity if not condition else 'ok','detail':detail,'sheet':sheet}

# ─── QC GOLD CHECKS ──────────────────────────────────────────────────────────

def run_sheet_checks(wb, file_label):
    checks = []
    for sn in wb.sheetnames:
        if sn == 'Intro': continue
        ws = wb[sn]
        data, month_m, month_m1 = get_sheet_data(ws)
        if not data or not month_m: continue
        label = f"{file_label} › {sn}"

        total = data.get('Total')
        dedup = next((data[k] for k in data if 'Dédupliqué' in k or 'Dedup' in k), None)
        sites = {k:v for k,v in data.items() if k not in SKIP_ROWS and 'Dédupliqué' not in k and 'Dedup' not in k}

        # Check 1: Total = sum of sites
        if total and sites:
            s = sum(v['m'] for v in sites.values())
            if s > 1000:
                diff_pct = abs(total['m']-s)/s*100
                checks.append(check_ok(
                    f"Total = sum of sites ({month_m})",
                    diff_pct < 1,
                    f"Total={total['m']:,.0f} | Sum={s:,.0f} | Diff={diff_pct:.2f}%",
                    "error", label))

        # Check 2: Dedup <= Total
        if dedup and total and total['m'] > 1000:
            checks.append(check_ok(
                f"Total Dédupliqué ≤ Total ({month_m})",
                dedup['m'] <= total['m'],
                f"Dedup={dedup['m']:,.0f} | Total={total['m']:,.0f}",
                "error", label))

        # Check 3: Evol% per site
        for site, vals in sites.items():
            vm, vm1 = vals['m'], vals['m1']
            if vm1 > 10 and vm > 10:
                evol = (vm/vm1-1)*100
                if abs(evol) > 30:
                    checks.append({'check':f"Evol% {site}",
                        'status':'warning','severity':'warning',
                        'detail':f"{evol:+.1f}% ({month_m1}={vm1:,.0f} → {month_m}={vm:,.0f})",
                        'sheet':label})
    return checks

def run_cross_checks(files_bytes):
    checks = []
    def get_dedup(key):
        if key not in files_bytes: return None, None
        wb = load_workbook(io.BytesIO(files_bytes[key]), data_only=True)
        sn = [s for s in wb.sheetnames if s != 'Intro']
        if not sn: return None, None
        ws = wb[sn[0]]
        data, m, _ = get_sheet_data(ws)
        if not data: return None, None
        dedup = next((data[k] for k in data if 'Dédupliqué' in k or 'Dedup' in k), None)
        return (dedup['m'] if dedup else None), m

    pairs = [
        ('file1','file3_1','File 1 Total Dedup = File 3.1 Total Dedup'),
        ('file1','file3_2','File 1 Total Dedup = File 3.2 Total Dedup'),
        ('file3_2','file4_1','File 3.2 Total Dedup = File 4.1 Total Dedup'),
    ]
    for k1,k2,name in pairs:
        v1,m1 = get_dedup(k1)
        v2,m2 = get_dedup(k2)
        if v1 and v2:
            diff_pct = abs(v1-v2)/v1*100 if v1>0 else 0
            checks.append(check_ok(
                f"{name} ({m1})",
                diff_pct < 1,
                f"{k1}={v1:,.0f} | {k2}={v2:,.0f} | Diff={diff_pct:.2f}%",
                "error", "Cross-file check"))
    return checks

"""
QC Gold V2 — Exact replicas of Panel Checker controls applied to HAM files.
"""
import io, datetime
from openpyxl import load_workbook

def excel_date_str(v):
    if isinstance(v, datetime.datetime): return v.strftime('%b-%y')
    if isinstance(v, str): return v.strip()
    if isinstance(v,(int,float)) and 40000<v<50000:
        return (datetime.datetime(1899,12,30)+datetime.timedelta(days=int(v))).strftime('%b-%y')
    return str(v)

def get_val_by_label(ws, label, col_offset=0, search_cols=range(1,5)):
    """Find row by label in any of search_cols, return value from last data col + offset."""
    for r in range(1, ws.max_row+1):
        for c in search_cols:
            cell_val = ws.cell(r,c).value
            if cell_val and str(cell_val).strip() == label:
                # Find last non-None column
                last_c = ws.max_column
                while last_c > c and ws.cell(r, last_c).value is None:
                    last_c -= 1
                target_c = last_c + col_offset
                if target_c > 0:
                    return ws.cell(r, target_c).value
    return None

def get_val_contains(ws, label_part, col_offset=0, search_cols=range(1,5)):
    """Find row where label contains label_part."""
    for r in range(1, ws.max_row+1):
        for c in search_cols:
            cell_val = ws.cell(r,c).value
            if cell_val and label_part.lower() in str(cell_val).lower():
                last_c = ws.max_column
                while last_c > c and ws.cell(r, last_c).value is None:
                    last_c -= 1
                target_c = last_c + col_offset
                if target_c > 0:
                    return ws.cell(r, target_c).value, str(cell_val).strip()
    return None, None

def get_last_two_months(ws, site_col=2):
    """Get last and second-to-last month columns and their labels."""
    header_row = None
    for r in range(1, min(20, ws.max_row+1)):
        v = ws.cell(r, site_col).value
        if v in ('Site','Département','Région'):
            header_row = r; break
    if not header_row: return None, None, None, None

    month_cols, months = [], []
    for c in range(site_col+1, ws.max_column+1):
        v = ws.cell(header_row, c).value
        if v is None: continue
        if (isinstance(v, datetime.datetime) or
            (isinstance(v,(int,float)) and 40000<v<50000) or
            (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
            month_cols.append(c); months.append(excel_date_str(v))
    if len(months)<2: return None, None, None, None
    return month_cols[-1], months[-1], month_cols[-2], months[-2]

def get_site_val(ws, site_name, site_col=2):
    """Get last month value for a site."""
    lc, lm, pc, pm = get_last_two_months(ws, site_col)
    if not lc: return None, None, None, None
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, site_col).value
        if v and str(v).strip() == site_name:
            return ws.cell(r,lc).value, ws.cell(r,pc).value, lm, pm
    return None, None, lm, pm

def check(name, ok, detail, severity="error", sheet=""):
    return {'check':name, 'status':'ok' if ok else 'fail',
            'severity':severity if not ok else 'ok', 'detail':detail, 'sheet':sheet}

def is_close(a, b, tol=0.01):
    if a is None or b is None: return False
    if a == 0 and b == 0: return True
    if a == 0 or b == 0: return False
    return abs(a-b)/max(abs(a),abs(b)) < tol

# ─── HELPERS FOR SOURCE-FILE CHECKS ─────────────────────────────────────────

def get_lc(ws):
    """Son ay kolonunu (lc) ve etiketini (lm) döndür — header aranmadan."""
    for r in range(1, 15):
        mc = []
        for c in range(1, min(ws.max_column+1, 200)):
            v = ws.cell(r, c).value
            if v is None: continue
            if (isinstance(v, datetime.datetime) or
                (isinstance(v,(int,float)) and 40000<v<50000) or
                (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
                mc.append((c, excel_date_str(v)))
        if len(mc) >= 3:
            return mc[-1][0], mc[-1][1]
    return None, None

def row_val(ws, label, lc, col=2, skip=0):
    """col kolonunda tam eşleşen etiketi bul, lc kolonundaki değeri döndür. skip=N → Nth eşleşme."""
    found = 0
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, col).value
        if v and isinstance(v, str) and v.strip() == label:
            if found == skip:
                return ws.cell(r, lc).value
            found += 1
    return None

def row_val_contains(ws, parts, lc, col=2):
    """Etikette 'parts' kelimelerini içeren ilk satırın lc değerini döndür."""
    parts = parts if isinstance(parts, list) else [parts]
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, col).value
        if v and isinstance(v, str):
            if all(p.lower() in v.lower() for p in parts):
                return ws.cell(r, lc).value
    return None

def fmt(v):
    if v is None: return 'N/A'
    try: return f"{float(v):,.0f}"
    except: return str(v)

# ─── TAB 3.1 CHECKS ──────────────────────────────────────────────────────────

def check_tab31(wb31_bytes, wb1_bytes):
    """
    Panel Checker tab 3.1 kontrolleri:
    - contrôle type     : Agences+Intermédiaires+Notaires+Autres = Total (3.1.4)
    - contrôle tab 1    : 3.1.4 Total Dédupliqué = 1.2 Total Panel Dédupliqué Marché (Pros)
    - contrôle segments : sections sum = grand total (3.1.4)
    """
    checks = []
    wb31 = load_workbook(io.BytesIO(wb31_bytes), data_only=True)
    wb1  = load_workbook(io.BytesIO(wb1_bytes), data_only=True)

    ws314 = wb31['3.1.4 Evolution Pros par type']
    ws12  = wb1['1.2 Pro_Part']

    lc314, lm = get_lc(ws314)
    lc12,  _  = get_lc(ws12)
    if not lc314 or not lc12:
        return checks

    # ── CHECK A: contrôle type — Agences+Intermédiaires+Notaires+Autres = Total général
    # Panel Checker tab 3.1 row 21: C21=C8+C9+C10+C11=C12
    # Dans 3.1.4 : section 1 (rows 3-22) = all pros
    #              section 2 (rows 25-45) = Agences
    #              section 3 (rows 48-65) = Intermédiaires
    #              section 4 (rows 70-89) = Notaires
    #              section 5 (rows 92-111) = Autres
    total_all  = row_val(ws314, 'Total', lc314, skip=0)          # row 18
    agences    = row_val(ws314, 'Total', lc314, skip=1)          # row 40
    interm     = row_val(ws314, 'Total', lc314, skip=2)          # row 63
    notaires   = row_val(ws314, 'Total', lc314, skip=3)          # row 85
    autres     = row_val(ws314, 'Total', lc314, skip=4)          # row 107

    if all(v is not None for v in [total_all, agences, interm, notaires, autres]):
        soma = agences + interm + notaires + autres
        checks.append(check(
            f"3.1.4 Agences+Interméd.+Notaires+Autres = Total ({lm})",
            soma == total_all,
            f"Somme segments={fmt(soma)} | Total={fmt(total_all)} | diff={fmt(abs(soma-total_all))}",
            "error", "tab 3.1 — contrôle type"
        ))

    # ── CHECK B: contrôle tab 1 — 3.1.4 Total Dédupliqué = 1.2 Total Dédupliqué Marché (Pros)
    # Panel Checker tab 3.1 row 23 AQ23: AQ18='tab 1'!O130
    # 3.1.4 row 20 = "Total Panel Dédupliqué" (annonceurs pros dédupliqués)
    # 1.2 Pro_Part row 21 = "Total Panel Dédupliqué Marché" (annonces pros dédup section Ancien)
    dedup_314 = row_val(ws314, 'Total Panel Dédupliqué', lc314)
    dedup_12  = row_val_contains(ws12, 'Dédupliqué Marché', lc12)

    if dedup_314 is not None and dedup_12 is not None:
        checks.append(check(
            f"3.1.4 Total Dédupliqué = 1.2 Total Dédupliqué Marché ({lm})",
            is_close(float(dedup_314), float(dedup_12)),
            f"3.1.4={fmt(dedup_314)} | 1.2={fmt(dedup_12)} | diff={fmt(abs(float(dedup_314)-float(dedup_12)))}",
            "error", "tab 3.1 — contrôle tab 1"
        ))

    # ── CHECK C: contrôle identification — par section, Dédupliqué ≤ Total
    # Panel Checker tab 3.1 row 22: C22=C18=C16+C17
    section_labels = [
        ('Agences immobilières',  1),
        ('Intermédiaires',        2),
        ('Notaires',              3),
        ('Autres annonceurs',     4),
    ]
    skip_map = {'Agences immobilières': 1, 'Intermédiaires': 2, 'Notaires': 3, 'Autres annonceurs': 4}
    for sec_name, skip_n in section_labels:
        sec_total = row_val(ws314, 'Total', lc314, skip=skip_n)
        sec_dedup = row_val(ws314, 'Total Panel Dédupliqué', lc314, skip=skip_n-1)
        if sec_total and sec_dedup and sec_total > 0:
            checks.append(check(
                f"3.1.4 {sec_name}: Dédupliqué ≤ Total ({lm})",
                float(sec_dedup) <= float(sec_total),
                f"Dédup={fmt(sec_dedup)} | Total={fmt(sec_total)}",
                "warning", "tab 3.1 — contrôle identification"
            ))

    return checks

# ─── TAB 3.2 CHECKS ──────────────────────────────────────────────────────────

def check_tab32(wb32_bytes, wb31_bytes):
    """
    Panel Checker tab 3.2 kontrolleri:
    - contrôle tab 3.1.4 : 3.2.1 Total par site = 3.1.4 Total Dédupliqué par site
    - contrôle segment   : 3.2.1 Grand Total = Agences+Interméd.+Notaires+Autres
    - contrôle total     : chaque section TOTAL = sum des régions
    """
    checks = []
    wb32 = load_workbook(io.BytesIO(wb32_bytes), data_only=True)
    wb31 = load_workbook(io.BytesIO(wb31_bytes), data_only=True)

    ws321 = wb32['3.2.1 Pros par régions']
    ws314 = wb31['3.1.4 Evolution Pros par type']
    lc314, lm = get_lc(ws314)
    if not lc314: return checks

    # ── BUILD 3.2.1 SITE COLUMN MAP ──
    # Row 6 = site names (avec quotes dans certains fichiers), row 7 = Pros/Poids
    # Colonnes impaires = Pros, paires = Poids
    site_col_map = {}  # site_name -> col index (Pros column)
    for c in range(3, ws321.max_column+1):
        raw = ws321.cell(6, c).value
        if raw and isinstance(raw, str):
            name = raw.strip().strip("'\"")
            # Ne garder que les colonnes "Pros" (col 7 = Pros/Poids alternance)
            header7 = ws321.cell(7, c).value
            if header7 and 'Pros' in str(header7):
                site_col_map[name] = c

    # ── FIND SECTION TOTAL ROWS in 3.2.1 ──
    # Structure: section 1 (all pros) rows 8→106 with TOTAL at row 106
    #            section 2 (Agences) rows → with TOTAL
    #            etc.
    total_rows = []  # list of row indices with label 'TOTAL'
    for r in range(1, ws321.max_row+1):
        v = ws321.cell(r, 2).value
        if v and str(v).strip().upper() == 'TOTAL':
            total_rows.append(r)

    # ── CHECK A: contrôle tab 3.1.4
    # 3.2.1 Grand Total per site = 3.1.4 Total Dédupliqué per site
    # Panel Checker: K24: K22='tab 3.1'!$O194
    # K22 = 3.2.1 Leboncoin TOTAL (section 1, grand total row)
    # tab 3.1 O194 = 3.1.4 Leboncoin Total Dédupliqué
    #
    # 3.1.4 n'a PAS de colonnes par site — chaque site est une ROW, colonnes = mois
    # Donc: pour chaque site dans 3.1.4, chercher row_val(ws314, site, lc314)
    # et comparer avec 3.2.1 TOTAL col pour ce site
    SITES_314 = ["Bien'ici", 'Figaro Immo', 'Green-Acres', 'Leboncoin',
                 'LogicImmo', 'MeilleursAgents', 'OuestFrance', 'PAP',
                 'ParuVendu', 'SeLoger', 'Superimmo']
    SITE_ALIAS = {'GreenAcres': 'Green-Acres', 'SuperImmo': 'Superimmo',
                  'Green-Acres': 'Green-Acres'}

    if total_rows:
        grand_total_row = total_rows[0]  # première section = all pros
        for site in SITES_314:
            # 3.1.4 value for this site
            v314 = row_val(ws314, site, lc314)
            if v314 is None:
                # try alias
                for alias, canonical in SITE_ALIAS.items():
                    if canonical == site:
                        v314 = row_val(ws314, alias, lc314)
                        if v314: break

            # 3.2.1 TOTAL for this site
            v321 = None
            for name, col in site_col_map.items():
                canon = SITE_ALIAS.get(name, name)
                if canon == site or name == site:
                    v321 = ws321.cell(grand_total_row, col).value
                    break

            if v314 is not None and v321 is not None and float(v314) > 0:
                checks.append(check(
                    f"3.2.1 {site} TOTAL = 3.1.4 {site} ({lm})",
                    is_close(float(v314), float(v321)),
                    f"3.1.4={fmt(v314)} | 3.2.1={fmt(v321)} | diff={fmt(abs(float(v314)-float(v321)))}",
                    "error", "tab 3.2 — contrôle tab 3.1.4"
                ))

    # ── CHECK B: contrôle segment
    # 3.2.1 Grand Total = section Agences + Intermédiaires + Notaires + Autres
    # Panel Checker: K25: K22=K45+K69+K91+K113
    # total_rows[0]=all pros, total_rows[1]=Agences, total_rows[2]=Interm, ...
    if len(total_rows) >= 5 and site_col_map:
        # Pour chaque site, vérifier grand_total = sum des 4 sections
        for site in SITES_314:
            col = None
            for name, c in site_col_map.items():
                canon = SITE_ALIAS.get(name, name)
                if canon == site or name == site:
                    col = c; break
            if col is None: continue

            gt   = ws321.cell(total_rows[0], col).value  # grand total
            ag   = ws321.cell(total_rows[1], col).value  # agences
            intr = ws321.cell(total_rows[2], col).value  # intermédiaires
            nt   = ws321.cell(total_rows[3], col).value  # notaires
            aut  = ws321.cell(total_rows[4], col).value  # autres

            vals = [gt, ag, intr, nt, aut]
            if all(isinstance(v, (int,float)) for v in vals) and float(gt) > 0:
                soma = float(ag) + float(intr) + float(nt) + float(aut)
                checks.append(check(
                    f"3.2.1 {site} Grand Total = Agences+Interméd.+Notaires+Autres ({lm})",
                    is_close(soma, float(gt)),
                    f"Somme={fmt(soma)} | Grand Total={fmt(gt)} | diff={fmt(abs(soma-float(gt)))}",
                    "error", "tab 3.2 — contrôle segment"
                ))

    return checks

# ─── ALL TABS CHECKS (tab 1, 2, 4.1, 4.2, 5, 5-2) ───────────────────────────

def check_generic_tab(wb_bytes, file_label):
    """Generic checks for all files: total consistency, dedup logic, evol% anomalies."""
    checks = []
    wb = load_workbook(io.BytesIO(wb_bytes), data_only=True)
    SKIP = {'Total Panel Dédupliqué','Total Panel Dédupliqué - Top 5 Sites',
            'Total Panel Dédupliqué  - Top 11 Sites','Total Panel Dédupliqué Marché',
            'Immobilier Notaire','Immonot','Site','Département','Région','TOTAL'}

    for sn in wb.sheetnames:
        if sn == 'Intro': continue
        ws = wb[sn]
        lc, lm, pc, pm = get_last_two_months(ws)
        if not lc: continue
        label = f"{file_label} › {sn}"

        # Collect data
        total = dedup = None
        sites = {}
        for r in range(1, ws.max_row+1):
            b = ws.cell(r,2).value
            if not b or not isinstance(b,str): continue
            b = b.strip()
            vm = ws.cell(r,lc).value
            vm1 = ws.cell(r,pc).value
            if not isinstance(vm,(int,float)): continue
            if b == 'Total': total = float(vm)
            elif 'Dédupliqué Marché' in b or (b=='Total Panel Dédupliqué' and dedup is None):
                dedup = float(vm)
            elif b not in SKIP and 'Dédupliqué' not in b and 'Panel' not in b:
                if float(vm) > 0:
                    sites[b] = {'m':float(vm),'m1':float(vm1) if isinstance(vm1,(int,float)) else 0}

        # Check: Total = sum of sites
        if total and sites:
            s = sum(v['m'] for v in sites.values())
            if s > 1000 and total > 1000:
                diff_pct = abs(total-s)/s*100
                checks.append(check(
                    f"Total = sum of sites ({lm})",
                    diff_pct < 1,
                    f"Total={total:,.0f} | Sum sites={s:,.0f} | Diff={diff_pct:.2f}%",
                    "error", label
                ))

        # Check: Dedup <= Total
        if dedup and total and total > 1000:
            checks.append(check(
                f"Total Dédupliqué ≤ Total ({lm})",
                dedup <= total,
                f"Dedup={dedup:,.0f} | Total={total:,.0f}",
                "error", label
            ))

        # Check: evol% per site
        for site, vals in sites.items():
            if vals['m1'] > 100 and vals['m'] > 10:
                evol = (vals['m']/vals['m1']-1)*100
                if abs(evol) > 30:
                    checks.append({
                        'check': f"Evol% {site}",
                        'status': 'warning', 'severity': 'warning',
                        'detail': f"{evol:+.1f}% ({pm}={vals['m1']:,.0f} → {lm}={vals['m']:,.0f})",
                        'sheet': label
                    })

    return checks

# ─── CROSS-FILE CHECKS ───────────────────────────────────────────────────────

def check_cross_files(files_bytes):
    checks = []

    def get_dedup_marche(key, sheet_idx=0):
        if key not in files_bytes: return None, None
        wb = load_workbook(io.BytesIO(files_bytes[key]), data_only=True)
        sheets = [s for s in wb.sheetnames if s != 'Intro']
        if not sheets: return None, None
        ws = wb[sheets[sheet_idx] if sheet_idx < len(sheets) else sheets[0]]
        lc, lm, _, _ = get_last_two_months(ws)
        if not lc: return None, None
        for r in range(1, ws.max_row+1):
            b = ws.cell(r,2).value
            if b and 'Dédupliqué Marché' in str(b):
                v = ws.cell(r,lc).value
                if isinstance(v,(int,float)): return float(v), lm
        # Try first dedup
        for r in range(1, ws.max_row+1):
            b = ws.cell(r,2).value
            if b and 'Dédupliqué' in str(b):
                v = ws.cell(r,lc).value
                if isinstance(v,(int,float)): return float(v), lm
        return None, None

    # 1.1 Total Dedup = 3.1.4 Total Dedup
    if 'file1' in files_bytes and 'file3_1' in files_bytes:
        v1, m1 = get_dedup_marche('file1', 0)  # 1.1 Total
        wb31 = load_workbook(io.BytesIO(files_bytes['file3_1']), data_only=True)
        ws314 = wb31['3.1.4 Evolution Pros par type']
        lc, lm, _, _ = get_last_two_months(ws314)
        v31 = None
        if lc:
            for r in range(1, ws314.max_row+1):
                b = ws314.cell(r,2).value
                if b and 'général' in str(b).lower() and 'profes' in str(b).lower():
                    v31 = ws314.cell(r, lc).value; break
        if v1 and v31:
            checks.append(check(
                f"1.1 Total Panel Dedup Marché = 3.1.4 Total général pros ({m1})",
                is_close(v1, float(v31)),
                f"File1={v1:,.0f} | File3.1={v31:,.0f} | diff={abs(v1-float(v31)):,.0f}",
                "error", "Cross-file: 1 vs 3.1"
            ))

    # 3.1.4 Total = 3.2.1 Grand Total
    if 'file3_1' in files_bytes and 'file3_2' in files_bytes:
        wb31 = load_workbook(io.BytesIO(files_bytes['file3_1']), data_only=True)
        wb32 = load_workbook(io.BytesIO(files_bytes['file3_2']), data_only=True)
        ws314 = wb31['3.1.4 Evolution Pros par type']
        ws321 = wb32['3.2.1 Pros par régions']

        lc31, lm31, _, _ = get_last_two_months(ws314)
        lc32, lm32, _, _ = get_last_two_months(ws321, site_col=2)

        # 3.1.4 Total Panel Dédupliqué (all sites combined)
        v314_total = None
        if lc31:
            for r in range(1, ws314.max_row+1):
                b = ws314.cell(r,2).value
                if b and str(b).strip() == 'Total':
                    v314_total = ws314.cell(r,lc31).value; break

        # 3.2.1 Grand TOTAL (sum of all regions, combined sites)
        v321_total = None
        if lc32:
            for r in range(ws321.max_row, 0, -1):
                b = ws321.cell(r,2).value
                if b and str(b).strip() == 'TOTAL':
                    # Sum all site columns
                    total_sum = 0
                    for c in range(3, ws321.max_column+1):
                        v = ws321.cell(r,c).value
                        if isinstance(v,(int,float)): total_sum += v
                    v321_total = total_sum; break

        if v314_total and v321_total and float(v314_total)>0:
            checks.append(check(
                f"3.1.4 Total pros = 3.2.1 TOTAL ({lm31})",
                is_close(float(v314_total), v321_total, tol=0.02),
                f"3.1.4={v314_total:,.0f} | 3.2.1={v321_total:,.0f}",
                "warning", "Cross-file: 3.1 vs 3.2"
            ))

    # 3.2 Total = 4.1 Total
    if 'file3_2' in files_bytes and 'file4_1' in files_bytes:
        v32, m32 = get_dedup_marche('file3_2', 0)
        v41, m41 = get_dedup_marche('file4_1', 0)
        if v32 and v41:
            checks.append(check(
                f"3.2 Total Dedup = 4.1 Total Dedup ({m32})",
                is_close(v32, v41),
                f"File3.2={v32:,.0f} | File4.1={v41:,.0f} | diff={abs(v32-v41):,.0f}",
                "error", "Cross-file: 3.2 vs 4.1"
            ))

    return checks

# ─── MAIN RUNNER ─────────────────────────────────────────────────────────────

def classify_files(uploaded_files):
    file_labels = {
        'file1':'1 — Evolution panel','file2':'2 — Performance qualité',
        'file3_1':'3.1 — Analyse Pros','file3_2':'3.2 — Géographique Pros',
        'file4_1':'4.1 — Stats géographiques','file4_2':'4.2 — Exclusivité/Partage',
        'file5':'5 — Focus IDF','file5_2':'5.2 — Grand Ouest',
    }
    files_bytes = {}
    for f in uploaded_files:
        nl = f.name.lower().replace(' ','_').replace('é','e').replace('è','e').replace('ô','o')
        fb = f.read()
        if '5_2' in nl or ('5' in nl and ('grand' in nl or 'ouest' in nl)):
            key = 'file5_2'
        elif '5' in nl and ('idf' in nl or 'ile' in nl or 'alpes' in nl or 'focus' in nl) and '5_2' not in nl:
            key = 'file5'
        elif '4_2' in nl or ('4' in nl and ('exclus' in nl or 'partag' in nl)):
            key = 'file4_2'
        elif '4_1' in nl or ('4' in nl and 'stat' in nl and '4_2' not in nl):
            key = 'file4_1'
        elif '3_2' in nl or ('3' in nl and 'geo' in nl):
            key = 'file3_2'
        elif '3_1' in nl or ('3' in nl and 'pros' in nl and '3_2' not in nl):
            key = 'file3_1'
        elif '2' in nl and ('perform' in nl or 'qualit' in nl):
            key = 'file2'
        elif '1' in nl and ('evolution' in nl or 'annonce' in nl) and '3_1' not in nl and '4_1' not in nl:
            key = 'file1'
        else:
            continue
        files_bytes[key] = fb
    return files_bytes, file_labels

def run_all_qc_gold(uploaded_files):
    files_bytes, file_labels = classify_files(uploaded_files)
    all_checks = []

    # 3.1 × 1 — contrôle type + contrôle tab 1
    if 'file3_1' in files_bytes and 'file1' in files_bytes:
        all_checks += check_tab31(files_bytes['file3_1'], files_bytes['file1'])
    # 3.2 × 3.1 — contrôle tab 3.1.4 + contrôle segment
    if 'file3_2' in files_bytes and 'file3_1' in files_bytes:
        all_checks += check_tab32(files_bytes['file3_2'], files_bytes['file3_1'])
    # Generic checks for all files
    for key, label in file_labels.items():
        if key in files_bytes:
            all_checks += check_generic_tab(files_bytes[key], label)
    # Cross-file checks
    all_checks += check_cross_files(files_bytes)

    return all_checks, files_bytes


def parse_ham_sections(uploaded_files):
    result = {}
    for f in uploaded_files:
        fb = f.read()
        wb = load_workbook(io.BytesIO(fb), data_only=True)
        for sn in wb.sheetnames:
            if sn == 'Intro': continue
            ws = wb[sn]
            for r in range(1, ws.max_row+1):
                b = ws.cell(r,2).value
                if b not in ('Site','Département'): continue
                month_cols, months = [], []
                for c in range(3, ws.max_column+1):
                    v = ws.cell(r,c).value
                    if v is None: continue
                    if (isinstance(v,datetime.datetime) or
                        (isinstance(v,(int,float)) and 40000<v<50000) or
                        (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23']))):
                        month_cols.append(c); months.append(excel_date_str(v))
                if len(months)<2: continue
                label = None
                for tr in range(r-1,max(0,r-5),-1):
                    v = ws.cell(tr,2).value
                    if v and isinstance(v,str) and len(v.strip())>3 and v.strip()!='Site':
                        label=v.strip(); break
                if not label: label=sn
                sites={}; totals={}
                for dr in range(r+1, min(r+35,ws.max_row+1)):
                    site=ws.cell(dr,2).value
                    if not site or not isinstance(site,str): break
                    site=site.strip()
                    if site=='Site': break
                    vals=[float(ws.cell(dr,c).value) if isinstance(ws.cell(dr,c).value,(int,float)) else None for c in month_cols]
                    real=[v for v in vals if v is not None and v>0]
                    if 'Total Panel' in site or site=='Total':
                        if len(real)>=2: totals[site]=vals
                    elif site not in SKIP_ROWS and site!='' and len(real)>=2:
                        sites[site]=vals
                if sites:
                    result[f"{sn}_{r}"] = {'sheet':sn,'label':label,'months':months,
                                           'sites':sites,'totals':totals,'file':f.name}
    return result

def analyze_trends(sections):
    rows = []
    for key, sec in sections.items():
        months = sec['months']
        total_dedup = {}
        for k,v in sec['totals'].items():
            if 'Dédupliqué' in k or 'Dedup' in k:
                total_dedup={i:x for i,x in enumerate(v) if x is not None and x>0}; break
        for site, vals in sec['sites'].items():
            real=[(i,v) for i,v in enumerate(vals) if v is not None and v>5]
            if len(real)<2: continue
            li,lv=real[-1]; pi,pv=real[-2]
            evol=(lv/pv-1)*100 if pv>0 else None
            pr_m=(lv/total_dedup[li]*100) if li in total_dedup and total_dedup[li]>0 else None
            pr_m1=(pv/total_dedup[pi]*100) if pi in total_dedup and total_dedup[pi]>0 else None
            evol_pr=(pr_m-pr_m1) if pr_m and pr_m1 else None
            status='ok'; flags=[]
            if evol is not None:
                if evol<=-20: status='critical'; flags.append(f"Monthly change: {evol:+.1f}% (vs prev month)")
                elif evol<=-10: status='warning'; flags.append(f"Monthly change: {evol:+.1f}% (vs prev month)")
                elif evol>=30: status='warning'; flags.append(f"Monthly surge: {evol:+.1f}%")
            max_hist=max((v for v in vals[:-1] if v and v>5),default=None)
            if max_hist and lv/max_hist<0.6:
                status='critical'; flags.append(f"Crash vs historical max: {(lv/max_hist-1)*100:+.1f}%")
            last3=[v for v in vals[-3:] if v and v>0]
            if len(last3)==3 and last3[0]>last3[1]>last3[2]:
                drop=(last3[2]-last3[0])/last3[0]*100
                if drop<-5:
                    if status=='ok': status='warning'
                    flags.append(f"3-month downtrend: {drop:+.1f}%")
            rows.append({'file':sec['file'],'sheet':sec['sheet'],'section':sec['label'],
                'site':site,'month_m':months[li] if li<len(months) else '?',
                'month_m1':months[pi] if pi<len(months) else '?',
                'val_m':lv,'val_m1':pv,'evol_pct':evol,'pr_m':pr_m,'evol_pr':evol_pr,
                'status':status,'flags':flags,'values':vals,'months':months})
    return rows

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## QC Monitor")
    st.divider()
    st.markdown("**Market**")
    sector = st.radio("", ["🏠 Real Estate","🚗 Auto"], label_visibility="collapsed", key="sector")
    if sector == "🏠 Real Estate":
        market = st.radio("", ["REFR"], label_visibility="collapsed", key="re_sub")
    else:
        st.caption("Coming soon"); market = None
    st.divider()
    st.markdown("**Page**")
    page = st.radio("", ["📊 Panel Tables Monitor","✅ Panel Checker (QC Gold)"],
        label_visibility="collapsed", key="page_sel")
    st.divider()
    st.markdown("**Upload files**")
    st.caption("Drop all files here — system sorts automatically")
    uploaded_files = st.file_uploader("", type=['xlsx'], accept_multiple_files=True,
        label_visibility="collapsed", key="global_upload")
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) loaded")

if market is None:
    st.markdown("# 🚗 Auto"); st.info("Coming soon"); st.stop()

# ─── PAGE: PANEL TABLES MONITOR ──────────────────────────────────────────────

if page == "📊 Panel Tables Monitor":
    st.markdown(f"# 📊 Panel Tables Monitor — {market}")
    st.divider()
    if not uploaded_files:
        st.info("👆 Upload your files in the sidebar to get started"); st.stop()

    with st.spinner("Loading..."):
        sections = parse_ham_sections(uploaded_files)
        rows = analyze_trends(sections)

    if not rows:
        st.warning("No data found."); st.stop()

    n_crit=sum(1 for r in rows if r['status']=='critical')
    n_warn=sum(1 for r in rows if r['status']=='warning')

    if n_crit>0: st.error(f"❌ REFUSED — {n_crit} critical · {n_warn} warnings")
    elif n_warn>3: st.warning(f"⚠️ TO REVIEW — {n_warn} warnings")
    else:
        st.success(f"✅ VALIDATED — {n_warn} minor warning(s)")
        if n_warn>0: st.info(f"{n_warn} warning(s) to monitor below")

    c1,c2 = st.columns([2,3])
    with c1:
        sev=st.radio("",["All","🔴 Critical","🟡 Warnings"],horizontal=True,key="sev_ptm")
    with c2:
        search=st.text_input("",placeholder="🔍 Search...",key="search_ptm",label_visibility="collapsed")

    filtered=rows
    if sev=="🔴 Critical": filtered=[r for r in rows if r['status']=='critical']
    elif sev=="🟡 Warnings": filtered=[r for r in rows if r['status'] in ('warning','critical')]
    if search:
        q=search.lower()
        filtered=[r for r in filtered if q in r['site'].lower() or q in r['section'].lower()]
    filtered=sorted(filtered,key=lambda r:{'critical':3,'warning':2,'ok':1}.get(r['status'],0),reverse=True)

    month_m=rows[0]['month_m'] if rows else ''
    month_m1=rows[0]['month_m1'] if rows else ''
    st.caption(f"M = {month_m}  ·  M-1 = {month_m1}  ·  {len(filtered)} series")

    table=[]
    for r in filtered:
        icon="🔴" if r['status']=='critical' else "🟡" if r['status']=='warning' else "✅"
        evol=f"{r['evol_pct']:+.1f}%" if r['evol_pct'] is not None else "—"
        pr=f"{r['pr_m']:.1f}%" if r['pr_m'] is not None else "—"
        evol_pr=f"{r['evol_pr']:+.2f}pp" if r['evol_pr'] is not None else "—"
        table.append({'':icon,'Section':r['section'][:35],'Site':r['site'],
            f"M ({month_m})":f"{r['val_m']:,.0f}",f"M-1 ({month_m1})":f"{r['val_m1']:,.0f}",
            'Evol %':evol,'Market share':pr,'MS evol':evol_pr,
            'Note':' | '.join(r['flags']) if r['flags'] else '—'})

    df=pd.DataFrame(table)
    def color_rows(row):
        if '🔴' in str(row.iloc[0]): return ['background-color:rgba(255,68,68,0.08)']*len(row)
        if '🟡' in str(row.iloc[0]): return ['background-color:rgba(245,166,35,0.05)']*len(row)
        return ['']*len(row)
    st.dataframe(df.style.apply(color_rows,axis=1),use_container_width=True,hide_index=True,
        height=min(600,40+35*len(table)))

    issue_rows=[r for r in filtered if r['flags']]
    if issue_rows:
        st.markdown("### Sites with anomalies")
        cols_n=3
        for i in range(0,min(len(issue_rows),12),cols_n):
            cols=st.columns(cols_n)
            for j,row in enumerate(issue_rows[i:i+cols_n]):
                with cols[j]:
                    clr='#ff4444' if row['status']=='critical' else '#f5a623'
                    evol_s=f"{row['evol_pct']:+.1f}%" if row['evol_pct'] else ''
                    icon2='🔴' if row['status']=='critical' else '🟡'
                    st.markdown(f"**{icon2} {row['site']}** — <span style='color:{clr}'>{evol_s}</span>",unsafe_allow_html=True)
                    st.caption(row['section'][:40])
                    vals=[v if v else 0 for v in row['values']]
                    fig=go.Figure(go.Scatter(x=row['months'][:len(vals)],y=vals,mode='lines+markers',
                        line=dict(color=clr,width=2),
                        marker=dict(size=[6 if k==len(vals)-1 else 0 for k in range(len(vals))]),
                        connectgaps=True,hovertemplate='%{x}: %{y:,.0f}<extra></extra>'))
                    fig.update_layout(height=120,margin=dict(l=0,r=0,t=0,b=0),
                        paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(showgrid=False,tickfont=dict(size=8),nticks=4),
                        yaxis=dict(showgrid=True,gridcolor='rgba(128,128,128,0.15)',
                            tickfont=dict(size=8),tickformat='.2s'),showlegend=False)
                    st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False},
                        key=f"ptm_{i}_{j}_{row['site'][:8]}")
                    for flag in row['flags']:
                        st.caption(f"⚠️ {flag}")

    st.divider()
    export=[{'File':r['file'],'Sheet':r['sheet'],'Section':r['section'],'Site':r['site'],
        'Status':r['status'],'Month M':r['month_m'],'Value M':r['val_m'],
        'Month M-1':r['month_m1'],'Value M-1':r['val_m1'],
        'Evol %':f"{r['evol_pct']:+.1f}%" if r['evol_pct'] else '',
        'Market share':f"{r['pr_m']:.1f}%" if r['pr_m'] else '',
        'Flags':' | '.join(r['flags'])} for r in rows]
    csv=pd.DataFrame(export).to_csv(index=False).encode('utf-8-sig')
    st.download_button("⬇️ Download report (CSV)",data=csv,
        file_name=f"ptm_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",mime='text/csv')

# ─── PAGE: PANEL CHECKER (QC GOLD) ───────────────────────────────────────────

elif page == "✅ Panel Checker (QC Gold)":
    st.markdown(f"# ✅ Panel Checker (QC Gold) — {market}")
    st.caption("Upload the 10 source files — same checks as the Panel Checker, no Panel Checker file needed.")
    st.divider()

    if not uploaded_files:
        st.info("👆 Upload your source files in the sidebar to get started"); st.stop()

    with st.spinner("Running QC Gold checks..."):
        all_checks, files_bytes = run_all_qc_gold(uploaded_files)

    # ── Fichiers détectés ──
    FILE_LABEL_MAP = {
        'file1':'1 — Evolution panel','file2':'2 — Performance qualité',
        'file3_1':'3.1 — Analyse Pros','file3_2':'3.2 — Géographique Pros',
        'file4_1':'4.1 — Stats géographiques','file4_2':'4.2 — Exclusivité/Partage',
        'file5':'5 — Focus IDF','file5_2':'5.2 — Grand Ouest',
    }
    EXPECTED = list(FILE_LABEL_MAP.keys())
    detected_keys = list(files_bytes.keys())
    missing = [FILE_LABEL_MAP[k] for k in EXPECTED if k not in detected_keys]

    if missing:
        st.warning(f"⚠️ {len(missing)} file(s) not detected: {', '.join(missing)}")
    else:
        st.success(f"✅ All {len(detected_keys)} source files detected")

    with st.expander("📂 Detected files", expanded=False):
        for k in EXPECTED:
            icon = "✅" if k in detected_keys else "❌"
            st.markdown(f"{icon} {FILE_LABEL_MAP[k]}")

    if not all_checks:
        st.warning("No checks could be run. Make sure the correct files are uploaded."); st.stop()

    st.divider()

    errors   = [c for c in all_checks if c['status']=='fail' and c['severity']=='error']
    warnings = [c for c in all_checks if c['severity']=='warning']
    ok_checks= [c for c in all_checks if c['status']=='ok']

    # ── Global verdict ──
    if errors:
        st.error(f"❌ **{len(errors)} error(s)** — Files need correction before sending to the team")
    elif warnings:
        st.warning(f"⚠️ **{len(warnings)} warning(s)** — Review before validating")
    else:
        st.success(f"✅ **All {len(ok_checks)} checks passed** — Files are clean")

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total checks", len(all_checks))
    c2.metric("✅ Passed",    len(ok_checks))
    c3.metric("❌ Errors",    len(errors))
    c4.metric("🟡 Warnings",  len(warnings))

    st.divider()

    # ── Résultats par section (groupés par sheet) ──
    by_sheet = defaultdict(list)
    for c in all_checks:
        by_sheet[c.get('sheet','Other')].append(c)

    def sheet_priority(items):
        if any(i['status']=='fail' and i['severity']=='error' for i in items): return 0
        if any(i['severity']=='warning' for i in items): return 1
        return 2

    for sheet, items in sorted(by_sheet.items(), key=lambda x: sheet_priority(x[1])):
        n_err  = sum(1 for i in items if i['status']=='fail' and i['severity']=='error')
        n_warn = sum(1 for i in items if i['severity']=='warning')
        n_ok   = sum(1 for i in items if i['status']=='ok')
        icon   = "❌" if n_err else "🟡" if n_warn else "✅"
        badge  = f"{n_err} error(s)" if n_err else f"{n_warn} warning(s)" if n_warn else f"{n_ok} OK"

        with st.expander(f"{icon} **{sheet}** — {badge}", expanded=(n_err > 0)):
            for item in items:
                if item['status'] == 'ok':
                    st.markdown(f"✅ {item['check']}")
                elif item['severity'] == 'error':
                    st.markdown(f"❌ **{item['check']}**")
                    st.caption(f"   {item['detail']}")
                else:
                    st.markdown(f"🟡 {item['check']}")
                    st.caption(f"   {item['detail']}")

    st.divider()

    # ── Message taslağı (sadece hata varsa) ──
    if errors or warnings:
        st.markdown("### 📧 Message draft for the team")
        month_guess = ""
        for f in uploaded_files:
            n = f.name.lower()
            for m in ['janvier','février','mars','avril','mai','juin',
                      'juillet','août','septembre','octobre','novembre','décembre']:
                if m in n:
                    month_guess = m.capitalize() + " 2026"; break
            if month_guess: break
        if not month_guess: month_guess = "mois courant"

        lines = [f"Bonjour,\n",
                 f"En vérifiant les fichiers sources pour {month_guess}, "
                 f"j'ai détecté les points suivants à corriger :\n"]
        for item in errors:
            lines.append(f"❌ [{item.get('sheet','')}] {item['check']}\n   → {item['detail']}")
        for item in warnings:
            lines.append(f"⚠️ [{item.get('sheet','')}] {item['check']}\n   → {item['detail']}")
        lines.append("\nMerci de vérifier et corriger avant validation.\n\nCordialement")
        msg = "\n".join(lines)
        st.text_area("", value=msg, height=320, key="msg_qc_gold")
        st.download_button("⬇️ Download message (.txt)", data=msg.encode("utf-8"),
            file_name=f"qc_errors_{datetime.datetime.now().strftime('%Y%m%d')}.txt",
            mime="text/plain")

    # ── CSV Export ──
    rows_exp = [{'Sheet': c.get('sheet',''), 'Check': c['check'],
        'Status': '❌ Error' if c['status']=='fail' and c['severity']=='error'
                  else '🟡 Warning' if c['severity']=='warning' else '✅ OK',
        'Detail': c.get('detail','')} for c in all_checks]
    csv = pd.DataFrame(rows_exp).to_csv(index=False).encode('utf-8-sig')
    st.download_button("⬇️ Download full report (CSV)", data=csv,
        file_name=f"qc_gold_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime='text/csv')
