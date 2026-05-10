import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import datetime, io
from pathlib import Path

st.set_page_config(page_title="QC Monitor", page_icon="🔍", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.block-container{padding-top:1.5rem}
div[data-testid="metric-container"]{background:#0f1117;border:1px solid #1c2030;border-radius:8px;padding:12px 16px}
div[data-testid="metric-container"] label{font-size:11px!important;color:#7c879e!important;text-transform:uppercase;letter-spacing:.5px}
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ──────────────────────────────────────────────────────────────────

SKIP_ROWS = {'Total','Total Panel Dédupliqué','Total Panel Dédupliqué - Top 5 Sites',
    'Total Panel Dédupliqué  - Top 11 Sites','Total Panel Dédupliqué Marché',
    'Immobilier Notaire','Immonot','Site','Département','Totaux','Total Panel Dedup'}

def excel_date_str(v):
    if isinstance(v, datetime.datetime): return v.strftime('%b-%y')
    if isinstance(v, str): return v.strip()
    if isinstance(v,(int,float)) and 40000<v<50000:
        return (datetime.datetime(1899,12,30)+datetime.timedelta(days=int(v))).strftime('%b-%y')
    return str(v)

# ─── PANEL CHECKER (QC GOLD) LOGIC ───────────────────────────────────────────

def validate_panel_checker(file_bytes):
    wb_v = load_workbook(io.BytesIO(file_bytes), data_only=True)
    wb_f = load_workbook(io.BytesIO(file_bytes), data_only=False)
    results = {}
    for sn in [s for s in wb_v.sheetnames if s.lower().startswith('tab')]:
        ws_v = wb_v[sn]; ws_f = wb_f[sn]
        trues = 0; falses = []
        for r in range(1, ws_v.max_row+1):
            for c in range(1, ws_v.max_column+1):
                v = ws_v.cell(r,c).value
                if v is True:
                    trues += 1
                elif v is False:
                    from openpyxl.utils import get_column_letter
                    label = ws_v.cell(r,2).value or ws_v.cell(r,1).value or ws_v.cell(r-1,2).value or ''
                    formula = str(wb_f[sn].cell(r,c).value or '')
                    col_header = ''
                    for tr in range(r-1, max(0,r-15),-1):
                        hv = ws_v.cell(tr,c).value
                        if hv and isinstance(hv,str) and len(hv.strip())>2:
                            col_header=hv.strip(); break
                    falses.append({'row':r,'col':get_column_letter(c),
                        'label':str(label).strip(),'col_header':col_header,'formula':formula[:100]})
        results[sn] = {'trues':trues,'falses':falses}
    return results

# ─── PANEL TABLES MONITOR (QC RAW) LOGIC ─────────────────────────────────────

def parse_ham_file(file_bytes, filename):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    result = {}
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
                    (isinstance(v,str) and any(m in v.lower() for m in ['-25','-26','-24','-23','avr','mars','mai','juin','juil','sep','oct','nov','dec','jan','fev']))):
                    month_cols.append(c); months.append(excel_date_str(v))
            if len(months)<2: continue
            label = None
            for tr in range(r-1, max(0,r-5),-1):
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
                result[f"{sn}_{r}"] = {'sheet':sn,'label':label,'months':months,'sites':sites,'totals':totals,'file':filename}
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

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## QC Monitor")
    st.divider()

    st.markdown("**Market**")
    sector = st.radio("", ["🏠 Real Estate", "🚗 Auto"], label_visibility="collapsed", key="sector")
    if sector == "🏠 Real Estate":
        market = st.radio("", ["REFR"], label_visibility="collapsed", key="re_sub")
    else:
        st.caption("Coming soon")
        market = None

    st.divider()

    st.markdown("**Page**")
    page = st.radio("", ["📊 Panel Tables Monitor", "✅ Panel Checker (QC Gold)"],
        label_visibility="collapsed", key="page_sel")

    st.divider()

    st.markdown("**Upload files**")
    st.caption("Drop all files — system sorts them automatically")
    uploaded_files = st.file_uploader("", type=['xlsx'],
        accept_multiple_files=True, label_visibility="collapsed", key="global_upload")

    # Auto-classify by filename
    panel_checker_file = None
    ham_files = []
    if uploaded_files:
        for f in uploaded_files:
            nl = f.name.lower()
            if 'panel_checker' in nl or 'panel checker' in nl:
                panel_checker_file = f
            else:
                ham_files.append(f)
        if panel_checker_file:
            st.success(f"✅ Panel Checker detected")
        if ham_files:
            st.success(f"✅ {len(ham_files)} source file(s)")

# ─── COMING SOON ──────────────────────────────────────────────────────────────

if market is None:
    st.markdown("# 🚗 Auto")
    st.info("Coming soon")
    st.stop()

# ─── PAGE: PANEL TABLES MONITOR ───────────────────────────────────────────────

if page == "📊 Panel Tables Monitor":
    st.markdown(f"# 📊 Panel Tables Monitor — {market}")
    st.divider()

    if not ham_files:
        st.info("👆 Upload your files in the sidebar to get started")
        st.stop()

    with st.spinner("Loading..."):
        all_sections = {}
        for f in ham_files:
            all_sections.update(parse_ham_file(f.read(), f.name))
        rows = analyze_trends(all_sections)

    if not rows:
        st.warning("No data found."); st.stop()

    n_crit = sum(1 for r in rows if r['status']=='critical')
    n_warn = sum(1 for r in rows if r['status']=='warning')

    if n_crit > 0: st.error(f"❌ REFUSED — {n_crit} critical · {n_warn} warnings")
    elif n_warn > 3: st.warning(f"⚠️ TO REVIEW — {n_warn} warnings")
    else:
        st.success(f"✅ VALIDATED — {n_warn} minor warning(s)")
        if n_warn > 0: st.info(f"{n_warn} warning(s) to monitor below")

    c1, c2 = st.columns([2,3])
    with c1:
        sev = st.radio("", ["All","🔴 Critical","🟡 Warnings"], horizontal=True, key="sev_ptm")
    with c2:
        search = st.text_input("", placeholder="🔍 Search...", key="search_ptm", label_visibility="collapsed")

    filtered = rows
    if sev == "🔴 Critical": filtered = [r for r in rows if r['status']=='critical']
    elif sev == "🟡 Warnings": filtered = [r for r in rows if r['status'] in ('warning','critical')]
    if search:
        q = search.lower()
        filtered = [r for r in filtered if q in r['site'].lower() or q in r['section'].lower()]
    filtered = sorted(filtered, key=lambda r: {'critical':3,'warning':2,'ok':1}.get(r['status'],0), reverse=True)

    month_m = rows[0]['month_m'] if rows else ''
    month_m1 = rows[0]['month_m1'] if rows else ''
    st.caption(f"M = {month_m}  ·  M-1 = {month_m1}  ·  {len(filtered)} series")

    table = []
    for r in filtered:
        icon = "🔴" if r['status']=='critical' else "🟡" if r['status']=='warning' else "✅"
        evol = f"{r['evol_pct']:+.1f}%" if r['evol_pct'] is not None else "—"
        pr = f"{r['pr_m']:.1f}%" if r['pr_m'] is not None else "—"
        evol_pr = f"{r['evol_pr']:+.2f}pp" if r['evol_pr'] is not None else "—"
        table.append({'':icon,'Section':r['section'][:35],'Site':r['site'],
            f"M ({month_m})":f"{r['val_m']:,.0f}",f"M-1 ({month_m1})":f"{r['val_m1']:,.0f}",
            'Evol %':evol,'Market share':pr,'MS evol':evol_pr,
            'Note':' | '.join(r['flags']) if r['flags'] else '—'})

    df = pd.DataFrame(table)
    def color_rows(row):
        if '🔴' in str(row.iloc[0]): return ['background-color:rgba(255,68,68,0.08)']*len(row)
        if '🟡' in str(row.iloc[0]): return ['background-color:rgba(245,166,35,0.05)']*len(row)
        return ['']*len(row)
    st.dataframe(df.style.apply(color_rows,axis=1), use_container_width=True, hide_index=True,
        height=min(600,40+35*len(table)))

    issue_rows = [r for r in filtered if r['flags']]
    if issue_rows:
        st.markdown("### Sites with anomalies")
        cols_n = 3
        for i in range(0, min(len(issue_rows),12), cols_n):
            cols = st.columns(cols_n)
            for j, row in enumerate(issue_rows[i:i+cols_n]):
                with cols[j]:
                    clr = '#ff4444' if row['status']=='critical' else '#f5a623'
                    evol_s = f"{row['evol_pct']:+.1f}%" if row['evol_pct'] else ''
                    icon2 = '🔴' if row['status']=='critical' else '🟡'
                    st.markdown(f"**{icon2} {row['site']}** — <span style='color:{clr}'>{evol_s}</span>", unsafe_allow_html=True)
                    st.caption(row['section'][:40])
                    vals = [v if v else 0 for v in row['values']]
                    fig = go.Figure(go.Scatter(
                        x=row['months'][:len(vals)], y=vals, mode='lines+markers',
                        line=dict(color=clr, width=2),
                        marker=dict(size=[6 if k==len(vals)-1 else 0 for k in range(len(vals))]),
                        connectgaps=True, hovertemplate='%{x}: %{y:,.0f}<extra></extra>',
                    ))
                    fig.update_layout(height=120, margin=dict(l=0,r=0,t=0,b=0),
                        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(showgrid=False,tickfont=dict(size=8),nticks=4),
                        yaxis=dict(showgrid=True,gridcolor='rgba(128,128,128,0.15)',
                            tickfont=dict(size=8),tickformat='.2s'),showlegend=False)
                    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar':False},
                        key=f"ptm_{i}_{j}_{row['site'][:8]}")
                    for flag in row['flags']:
                        st.caption(f"⚠️ {flag}")

    st.divider()
    export = [{'File':r['file'],'Sheet':r['sheet'],'Section':r['section'],'Site':r['site'],
        'Status':r['status'],'Month M':r['month_m'],'Value M':r['val_m'],
        'Month M-1':r['month_m1'],'Value M-1':r['val_m1'],
        'Evol %':f"{r['evol_pct']:+.1f}%" if r['evol_pct'] else '',
        'Market share':f"{r['pr_m']:.1f}%" if r['pr_m'] else '',
        'Flags':' | '.join(r['flags'])} for r in rows]
    csv = pd.DataFrame(export).to_csv(index=False).encode('utf-8-sig')
    st.download_button("⬇️ Download report (CSV)", data=csv,
        file_name=f"ptm_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime='text/csv')

# ─── PAGE: PANEL CHECKER (QC GOLD) ───────────────────────────────────────────

elif page == "✅ Panel Checker (QC Gold)":
    st.markdown(f"# ✅ Panel Checker (QC Gold) — {market}")
    st.divider()

    if not panel_checker_file:
        st.info("👆 Upload your Panel Checker file in the sidebar to get started")
        st.caption("File name must contain 'Panel_checker' or 'Panel checker'")
        st.stop()

    with st.spinner("Checking controls..."):
        validation = validate_panel_checker(panel_checker_file.read())

    total_false = sum(len(v['falses']) for v in validation.values())

    if total_false == 0:
        st.success("✅ VALIDATED — all controls passed")
    else:
        st.error(f"❌ REFUSED — {total_false} error(s) found")

    for sn, data in validation.items():
        fc = len(data['falses'])
        if fc == 0:
            st.markdown(f"✅ **{sn}** — OK ({data['trues']} checks)")
        else:
            with st.expander(f"❌ **{sn}** — {fc} error(s)"):
                for f in data['falses']:
                    ch = f" · *{f['col_header']}*" if f['col_header'] else ""
                    formula = f"`{f['formula']}`" if f['formula'] else ""
                    st.markdown(f"- **Row {f['row']} · Col {f['col']}** · `{f['label']}`{ch} {formula}")
