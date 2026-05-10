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

def classify_files(uploaded_files):
    file_map = {
        'file1': ['1_analyse_evolution','1_analyze_evolution'],
        'file2': ['2_analyse_performance','2_analyze_performance'],
        'file3_1': ['3_1_analyse','3_1_analyze'],
        'file3_2': ['3_2_analyse','3_2_analyze'],
        'file4_1': ['4_1_statistiques','4_1_statistics'],
        'file4_2': ['4_2_statistiques','4_2_statistics'],
        'file5': ['5_focus_idf','5_focus_ile','5_focus_alpes'],
        'file5_2': ['5_2_focus','5_2_grand'],
    }
    file_labels = {
        'file1':'1 — Evolution panel','file2':'2 — Performance qualité',
        'file3_1':'3.1 — Analyse Pros','file3_2':'3.2 — Géographique Pros',
        'file4_1':'4.1 — Stats géographiques','file4_2':'4.2 — Exclusivité/Partage',
        'file5':'5 — Focus IDF','file5_2':'5.2 — Grand Ouest',
    }
    files_bytes = {}
    for f in uploaded_files:
        nl = f.name.lower().replace(' ','_')
        for key, patterns in file_map.items():
            if any(p in nl for p in patterns):
                files_bytes[key] = f.read()
                break
    return files_bytes, file_labels

def run_all_qc_gold(uploaded_files):
    files_bytes, file_labels = classify_files(uploaded_files)
    all_checks = []
    for key, label in file_labels.items():
        if key not in files_bytes: continue
        wb = load_workbook(io.BytesIO(files_bytes[key]), data_only=True)
        all_checks += run_sheet_checks(wb, label)
    all_checks += run_cross_checks(files_bytes)
    return all_checks, list(files_bytes.keys())

# ─── PANEL TABLES MONITOR (TREND ANALYSIS) ───────────────────────────────────

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
