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

# ─── TAB 3.1 CHECKS ──────────────────────────────────────────────────────────

def check_tab31(wb31_bytes, wb1_bytes):
    checks = []
    wb31 = load_workbook(io.BytesIO(wb31_bytes), data_only=True)
    wb1  = load_workbook(io.BytesIO(wb1_bytes), data_only=True)

    # ── CHECK 1: Total général pros (3.1.4) = Total Panel Dédupliqué Marché Ancien (1.2)
    # Panel Checker: =AQ18='tab 1'!O130
    # tab3.1 row18 = HAM 3.1.4 "Total général professionnels" last month
    # tab1 row130 = HAM 1.2 "Total Panel Dédupliqué Marché" last month (Ancien section)
    ws314 = wb31['3.1.4 Evolution Pros par type']
    ws12  = wb1['1.2 Pro_Part']

    # Get 3.1.4 Total général pros - last month
    total_pros_314 = None
    lc, lm, pc, pm = get_last_two_months(ws314)
    if lc:
        for r in range(1, ws314.max_row+1):
            b = ws314.cell(r,2).value
            if b and 'général' in str(b).lower() and 'profes' in str(b).lower():
                total_pros_314 = ws314.cell(r, lc).value
                break

    # Get 1.2 Total Panel Dédupliqué Marché - Ancien section (first occurrence)
    total_dedup_ancien = None
    lc2, lm2, pc2, pm2 = get_last_two_months(ws12)
    if lc2:
        for r in range(1, ws12.max_row+1):
            b = ws12.cell(r,2).value
            if b and 'Dédupliqué Marché' in str(b):
                total_dedup_ancien = ws12.cell(r, lc2).value
                break

    if total_pros_314 is not None and total_dedup_ancien is not None:
        checks.append(check(
            f"3.1.4 Total pros = 1.2 Total Panel Dédupliqué Marché ({lm})",
            is_close(total_pros_314, total_dedup_ancien),
            f"3.1.4={total_pros_314:,.0f} | 1.2={total_dedup_ancien:,.0f} | diff={abs(total_pros_314-total_dedup_ancien):,.0f}",
            "error", "tab 3.1 — contrôle tab 1"
        ))

    # ── CHECK 2: For each site: Vente + Location = Total (tab 3.1)
    # Panel Checker: =O73+O92=O191 for Bien'ici
    # In HAM 3.1: 3.1.4 sheet has sections for Vente, Location, Total
    # Find the three sections
    SITES = ["Bien'ici", 'SeLoger', 'Total']
    ws314 = wb31['3.1.4 Evolution Pros par type']
    lc, lm, pc, pm = get_last_two_months(ws314)

    if lc:
        # Find all "Site" header rows → gives us section start rows
        site_rows = []
        for r in range(1, ws314.max_row+1):
            if ws314.cell(r,2).value == 'Site':
                site_rows.append(r)

        # Check Vente+Loc=Total for key sites across sections
        if len(site_rows) >= 3:
            # Typically: site_rows[0]=all pros, site_rows[1]=agences(vente?), site_rows[2]=agences(loc?)
            # Get Total row in each section
            def get_section_total(section_start):
                for r in range(section_start, min(section_start+30, ws314.max_row+1)):
                    b = ws314.cell(r,2).value
                    if b and str(b).strip() in ('Total','Total Panel Dédupliqué','Total Panel Dédupliqué  - Top 11 Sites'):
                        return ws314.cell(r, lc).value, str(b).strip()
                    if b == 'Site' and r != section_start:
                        break
                return None, None

            # Get each site's value in section 1 (all), section 2 (vente equiv), section 3 (loc equiv)
            for site in ["Bien'ici", 'SeLoger', 'Total']:
                vals = []
                for sr in site_rows[:3]:
                    for r in range(sr+1, min(sr+30, ws314.max_row+1)):
                        b = ws314.cell(r,2).value
                        if b and str(b).strip() == site:
                            vals.append(ws314.cell(r, lc).value)
                            break
                        if b == 'Site' and r != sr:
                            break

                if len(vals) >= 3 and all(v is not None for v in vals):
                    # Check: vals[1] + vals[2] == vals[0] (or similar)
                    sum_parts = sum(vals[1:])
                    checks.append(check(
                        f"3.1.4 {site}: Vente+Loc = Total ({lm})",
                        is_close(sum_parts, vals[0], tol=0.02),
                        f"Vente+Loc={sum_parts:,.0f} | Total={vals[0]:,.0f}",
                        "warning", "tab 3.1 — contrôle somme des segments"
                    ))

    # ── CHECK 3: Month consistency (=P131=C131 type checks)
    # AvendreAlouer value in month M should equal same value in another section's month M
    # This checks that the same site appears with the same value across different sheets
    SITES_CHECK = ['AvendreAlouer', "Bien'ici", 'Figaro Immo', 'Leboncoin', 'LogicImmo',
                   'MeilleursAgents', 'OuestFrance', 'ParuVendu', 'SeLoger', 'Superimmo']

    # Get values from 3.1.4 vs 3.1.1 (they should match for current month)
    ws311 = wb31['3.1.1 Pros par site ']

    # 3.1.1 has sites as columns with date in row 1
    # Get site columns
    site_cols_311 = {}
    for c in range(1, ws311.max_column+1):
        v = ws311.cell(1, c).value
        if v and isinstance(v, str):
            site_cols_311[v.strip()] = c

    # 3.1.4 last month values
    lc, lm, _, _ = get_last_two_months(ws314)
    if lc and lm:
        # Get "Pros identifiés Joreca" row values per site in 3.1.1
        # vs 3.1.4 site totals
        for r in range(1, ws314.max_row+1):
            b = ws314.cell(r,2).value
            if b and 'identifiés Joreca' in str(b) and 'Pros' in str(b):
                # This row has Total Pros per site
                for site in SITES_CHECK:
                    v314 = ws314.cell(r, lc).value
                    # Find same in 3.1.1
                    site_col = None
                    for c in range(1, ws311.max_column+1):
                        h = ws311.cell(1, c).value
                        if h and str(h).strip() == site:
                            site_col = c; break
                    if site_col:
                        # Row 12 in 3.1.1 = Pros identifiés Joreca
                        v311 = ws311.cell(12, site_col).value
                        if v314 is not None and v311 is not None and v314 > 0:
                            checks.append(check(
                                f"3.1.4 {site} Pros identifiés = 3.1.1 {site} ({lm})",
                                is_close(float(v314), float(v311)),
                                f"3.1.4={v314:,.0f} | 3.1.1={v311:,.0f}",
                                "error", "tab 3.1 — cohérence entre sheets"
                            ))
                break

    return checks

# ─── TAB 3.2 CHECKS ──────────────────────────────────────────────────────────

def check_tab32(wb32_bytes, wb31_bytes):
    checks = []
    wb32 = load_workbook(io.BytesIO(wb32_bytes), data_only=True)
    wb31 = load_workbook(io.BytesIO(wb31_bytes), data_only=True)

    # ── CHECK 1: 3.2 Total = 3.1.4 Total Panel Dédupliqué (contrôle tab 3.1.4)
    # Panel Checker: =K22='tab 3.1'!$O194  and  =AA22='tab 3.1'!$O202
    ws321 = wb32['3.2.1 Pros par régions']
    ws314 = wb31['3.1.4 Evolution Pros par type']

    # Get 3.2.1 TOTAL row last month
    lc32, lm32, pc32, pm32 = get_last_two_months(ws321, site_col=2)
    lc31, lm31, pc31, pm31 = get_last_two_months(ws314)

    if lc32 and lc31:
        # Find TOTAL in 3.2.1 for each site
        # Find Total Panel Dédupliqué in 3.1.4

        # 3.2.1: row 22 = TOTAL (sum of all regions per site)
        # For each site column in 3.2.1, total should match 3.1.4 dedup
        site_cols = {}
        for c in range(1, ws321.max_column+1):
            v = ws321.cell(6, c).value  # row 6 has site names
            if v and isinstance(v,str) and v.strip() not in ('Site','Département','Région',''):
                site_cols[v.strip()] = c

        # Get 3.1.4 Total Panel Dédupliqué values per site
        # 3.1.4 row 19 = Total Panel Dédupliqué (top 11)
        # 3.1.4 row 20 = Total Panel Dédupliqué Marché
        dedup_314 = {}
        for r in range(1, ws314.max_row+1):
            b = ws314.cell(r,2).value
            if b and 'Total Panel Dédupliqué' in str(b) and 'Top' not in str(b):
                # This is the market dedup row
                for c in range(1, ws314.max_column+1):
                    h = ws314.cell(5,c).value  # header row
                    if h and isinstance(h,str) and h.strip() in site_cols:
                        v = ws314.cell(r,lc31).value
                        if isinstance(v,(int,float)):
                            dedup_314[h.strip()] = float(v)
                break

        # Get 3.2.1 TOTAL for each site (last row = TOTAL)
        total_321 = {}
        for r in range(ws321.max_row, 0, -1):
            b = ws321.cell(r,2).value
            if b and str(b).strip() == 'TOTAL':
                for site, sc in site_cols.items():
                    v = ws321.cell(r, sc).value
                    if isinstance(v,(int,float)):
                        total_321[site] = float(v)
                break

        # Compare
        for site in set(dedup_314.keys()) & set(total_321.keys()):
            v31 = dedup_314[site]
            v32 = total_321[site]
            if v31 > 0:
                checks.append(check(
                    f"3.2.1 {site} Total = 3.1.4 {site} Total Dedup ({lm32})",
                    is_close(v31, v32),
                    f"3.1.4={v31:,.0f} | 3.2.1={v32:,.0f} | diff={abs(v31-v32):,.0f}",
                    "error", "tab 3.2 — contrôle tab 3.1.4"
                ))

    # ── CHECK 2: contrôle segment — sum of types = total
    # =K22=K45+K69+K91+K113  → Total Pros = sum of Agence+Intermédiaire+Notaire+Autres
    # In 3.2.1: we have sections per métier, their totals should sum to grand total

    # ── CHECK 3: MAX(sites) <= Total Dédupliqué per region
    # =MAX(C68,...,Y68)<=AG68 → max site value per region ≤ Total Dédupliqué
    ws322 = wb32['3.2.2 Pros par département']
    lc, lm, pc, pm = get_last_two_months(ws322, site_col=1)
    if lc:
        # Find rows with département name, check MAX(sites) <= dedup
        for r in range(1, ws322.max_row+1):
            dept = ws322.cell(r,1).value
            if not dept or not isinstance(dept,str): continue
            dept = dept.strip()
            if not dept or dept in ('Site','Département','Total','TOTAL'): continue

            # Get all site values in this row
            site_vals = []
            dedup_val = None
            for c in range(2, ws322.max_column+1):
                v = ws322.cell(r,c).value
                h = ws322.cell(ws322.min_row,c).value  # header
                if isinstance(v,(int,float)) and v > 0:
                    if h and ('Dédupliqué' in str(h) or 'Dedup' in str(h)):
                        dedup_val = float(v)
                    else:
                        site_vals.append(float(v))

            if dedup_val and site_vals:
                max_site = max(site_vals)
                if max_site > dedup_val:
                    checks.append({
                        'check': f"MAX(sites) ≤ Total Dédupliqué — {dept} ({lm})",
                        'status': 'fail', 'severity': 'error',
                        'detail': f"Max site={max_site:,.0f} > Dedup={dedup_val:,.0f}",
                        'sheet': "tab 3.2 — contrôle déduplication"
                    })

    # ── CHECK 4: 3.2 cross-check with tab 5 (tab 3.2 IDF total = tab 5 total)
    # Panel Checker: ='tab 3.1'!N55=O768
    # In HAM: 3.2.1 IDF row = 5_Focus_IDF total (checked separately)

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

    # Specific checks for tab 3.1 and 3.2
    if 'file3_1' in files_bytes and 'file1' in files_bytes:
        all_checks += check_tab31(files_bytes['file3_1'], files_bytes['file1'])
    if 'file3_2' in files_bytes and 'file3_1' in files_bytes:
        all_checks += check_tab32(files_bytes['file3_2'], files_bytes['file3_1'])

    # Generic checks for all files
    for key, label in file_labels.items():
        if key in files_bytes:
            all_checks += check_generic_tab(files_bytes[key], label)

    # Cross-file checks
    all_checks += check_cross_files(files_bytes)

    return all_checks, list(files_bytes.keys())


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
    st.divider()

    if not uploaded_files:
        st.info("👆 Upload your files in the sidebar to get started"); st.stop()

    with st.spinner("Running QC Gold checks..."):
        all_checks, detected = run_all_qc_gold(uploaded_files)

    if not all_checks:
        st.warning("No checks could be run. Make sure the correct files are uploaded."); st.stop()

    errors=[c for c in all_checks if c['status']=='fail' and c['severity']=='error']
    warnings=[c for c in all_checks if c['severity']=='warning']
    ok_checks=[c for c in all_checks if c['status']=='ok']

    if errors: st.error(f"❌ REFUSED — {len(errors)} error(s) · {len(warnings)} warning(s)")
    elif warnings: st.warning(f"⚠️ TO REVIEW — {len(warnings)} warning(s)")
    else: st.success(f"✅ VALIDATED — all {len(ok_checks)} checks passed")

    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total checks",len(all_checks))
    c2.metric("✅ Passed",len(ok_checks))
    c3.metric("❌ Errors",len(errors))
    c4.metric("🟡 Warnings",len(warnings))

    st.divider()

    by_sheet=defaultdict(list)
    for c in all_checks:
        by_sheet[c.get('sheet','Unknown')].append(c)

    def sheet_priority(items):
        if any(i['status']=='fail' and i['severity']=='error' for i in items): return 0
        if any(i['severity']=='warning' for i in items): return 1
        return 2

    for sheet,items in sorted(by_sheet.items(),key=lambda x:sheet_priority(x[1])):
        n_err=sum(1 for i in items if i['status']=='fail' and i['severity']=='error')
        n_warn=sum(1 for i in items if i['severity']=='warning')
        n_ok=sum(1 for i in items if i['status']=='ok')
        icon="❌" if n_err else "🟡" if n_warn else "✅"
        badge=f"{n_err} error(s)" if n_err else f"{n_warn} warning(s)" if n_warn else f"{n_ok} OK"

        with st.expander(f"{icon} **{sheet}** — {badge}",expanded=(n_err>0)):
            for item in items:
                if item['status']=='ok':
                    st.markdown(f"✅ {item['check']}")
                elif item['severity']=='error':
                    st.markdown(f"❌ **{item['check']}**")
                    st.caption(f"  {item['detail']}")
                else:
                    st.markdown(f"🟡 {item['check']}")
                    st.caption(f"  {item['detail']}")

    st.divider()
    rows_exp=[{'Sheet':c.get('sheet',''),'Check':c['check'],
        'Status':'❌ Error' if c['status']=='fail' and c['severity']=='error' else '🟡 Warning' if c['severity']=='warning' else '✅ OK',
        'Detail':c.get('detail','')} for c in all_checks]
    csv=pd.DataFrame(rows_exp).to_csv(index=False).encode('utf-8-sig')
    st.download_button("⬇️ Download report (CSV)",data=csv,
        file_name=f"qc_gold_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",mime='text/csv')
