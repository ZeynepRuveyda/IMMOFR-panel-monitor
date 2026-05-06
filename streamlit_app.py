import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import datetime, json, io
from pathlib import Path

"""
Panel Checker V2 — Analyse directe des fichiers HAM
Parse les fichiers source et reproduit les calculs du Panel Checker Excel.
"""

import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import numpy as np
import datetime
import io
from openpyxl import load_workbook

# ─── CONSTANTES ───────────────────────────────────────────────────────────────

SITES_ORDER = [
    'AvendreAlouer', "Bien'ici", 'Figaro Immo', 'Green-Acres',
    'Leboncoin', 'LogicImmo', 'MeilleursAgents', 'OuestFrance',
    'PAP', 'ParuVendu', 'SeLoger', 'SuperImmo'
]

SITES_COLORS = {
    'AvendreAlouer': '#27CCC3',
    "Bien'ici": '#FFC000',
    'Figaro Immo': '#E80536',
    'Green-Acres': '#84A824',
    'Leboncoin': '#FC6E2B',
    'LogicImmo': '#000000',
    'MeilleursAgents': '#1E91FF',
    'OuestFrance': '#E1000F',
    'PAP': '#1A48DC',
    'ParuVendu': '#242021',
    'SeLoger': '#E6103E',
    'SuperImmo': '#FB2375',
}

SKIP_ROWS = {
    'Total', 'Total Panel Dédupliqué', 'Total Panel Dédupliqué - Top 5 Sites',
    'Total Panel Dédupliqué  - Top 11 Sites', 'Total Panel Dédupliqué Marché',
    'Total Panel Dédupliq', 'Immobilier Notaire', 'Immonot', 'Site',
    'Totaux', 'Intro', 'Total Panel Dedup',
}

# ─── PARSING ──────────────────────────────────────────────────────────────────

def excel_date_str(v):
    if isinstance(v, datetime.datetime):
        return v.strftime('%b-%y')
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (int, float)) and 40000 < v < 50000:
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime('%b-%y')
    return str(v)

def parse_ham_file(file_bytes, filename):
    """Parse un fichier HAM et retourne les sections avec données mensuelles."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    result = {}
    
    for sn in wb.sheetnames:
        if sn == 'Intro':
            continue
        ws = wb[sn]
        
        # Find Site header rows
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if b != 'Site' and b != 'Département':
                continue
            
            # Get month columns
            month_cols = []
            months = []
            for c in range(3, ws.max_column + 1):
                v = ws.cell(r, c).value
                if v is None:
                    continue
                s = excel_date_str(v)
                if (isinstance(v, datetime.datetime) or
                    (isinstance(v, (int, float)) and 40000 < v < 50000) or
                    (isinstance(v, str) and any(m in v.lower() for m in
                     ['-25', '-26', '-24', 'avr', 'mars', 'mai', 'juin',
                      'juil', 'août', 'sep', 'oct', 'nov', 'dec', 'jan', 'fev']))):
                    month_cols.append(c)
                    months.append(s)
            
            if len(months) < 2:
                continue
            
            # Get section label
            label = None
            for tr in range(r - 1, max(0, r - 5), -1):
                v = ws.cell(tr, 2).value
                if v and isinstance(v, str) and len(v.strip()) > 3 and v.strip() != 'Site':
                    label = v.strip()
                    break
            if not label:
                label = sn
            
            # Extract site data
            sites = {}
            for dr in range(r + 1, min(r + 30, ws.max_row + 1)):
                site = ws.cell(dr, 2).value
                if not site or not isinstance(site, str):
                    break
                site = site.strip()
                if site in SKIP_ROWS or site == '' or site == 'Site':
                    if site == 'Site':
                        break
                    continue
                
                vals = []
                for c in month_cols:
                    v = ws.cell(dr, c).value
                    if isinstance(v, (int, float)):
                        vals.append(float(v))
                    else:
                        vals.append(None)
                
                real = [v for v in vals if v is not None and v > 0]
                if len(real) >= 2:
                    sites[site] = vals
            
            # Also get totals
            total_data = {}
            for dr in range(r + 1, min(r + 40, ws.max_row + 1)):
                site = ws.cell(dr, 2).value
                if not site or not isinstance(site, str):
                    break
                site = site.strip()
                if 'Total Panel' in site or site == 'Total':
                    vals = []
                    for c in month_cols:
                        v = ws.cell(dr, c).value
                        vals.append(float(v) if isinstance(v, (int, float)) else None)
                    total_data[site] = vals
            
            if sites:
                key = f"{sn}__r{r}"
                result[key] = {
                    'sheet': sn,
                    'label': label,
                    'months': months,
                    'sites': sites,
                    'totals': total_data,
                    'file': filename,
                }
    
    return result

# ─── CALCULS PANEL CHECKER ───────────────────────────────────────────────────

def compute_metrics(sites_data, totals_data, months):
    """
    Calcule pour chaque site:
    - Valeur M (dernier mois)
    - Valeur M-1 
    - Evol% = (M - M-1) / M-1
    - Part de Réseau M = site / total_dedup
    - Part de Réseau M-1
    - Market Share Evol (pp) = PR_M - PR_M-1
    """
    results = []
    
    # Find total dédupliqué
    total_dedup_m = None
    total_dedup_m1 = None
    for k, v in totals_data.items():
        if 'Dédupliqué' in k or 'Dedup' in k:
            real = [(i, x) for i, x in enumerate(v) if x is not None and x > 0]
            if len(real) >= 2:
                total_dedup_m = real[-1][1]
                total_dedup_m1 = real[-2][1]
                break
    
    # Total simple
    total_m = total_m1 = None
    if 'Total' in totals_data:
        real = [(i, x) for i, x in enumerate(totals_data['Total']) if x is not None and x > 0]
        if len(real) >= 2:
            total_m = real[-1][1]
            total_m1 = real[-2][1]
    
    for site_name, vals in sites_data.items():
        real_pairs = [(i, v) for i, v in enumerate(vals) if v is not None and v > 0]
        if len(real_pairs) < 2:
            continue
        
        last_i, last_v = real_pairs[-1]
        prev_i, prev_v = real_pairs[-2]
        
        month_m = months[last_i] if last_i < len(months) else '?'
        month_m1 = months[prev_i] if prev_i < len(months) else '?'
        
        evol_pct = (last_v / prev_v - 1) * 100 if prev_v > 0 else None
        
        pr_m = (last_v / total_dedup_m * 100) if total_dedup_m and total_dedup_m > 0 else None
        pr_m1 = (prev_v / total_dedup_m1 * 100) if total_dedup_m1 and total_dedup_m1 > 0 else None
        evol_pr = (pr_m - pr_m1) if pr_m is not None and pr_m1 is not None else None
        
        # Anomaly detection
        anomalies = []
        max_hist = max((v for v in vals[:-1] if v is not None and v > 5), default=None)
        
        if evol_pct is not None:
            if evol_pct <= -20:
                anomalies.append(('critical', f'Monthly change: {evol_pct:+.1f}%'))
            elif evol_pct <= -10:
                anomalies.append(('warning', f'Monthly change: {evol_pct:+.1f}%'))
        
        if max_hist and last_v / max_hist < 0.6:
            pct = (last_v / max_hist - 1) * 100
            anomalies.append(('critical', f'Crash vs max: {pct:+.1f}%'))
        
        if evol_pr is not None and evol_pr < -2:
            anomalies.append(('warning', f'Market share loss: {evol_pr:+.2f}pp'))
        
        # Sustained decline
        last3 = [v for v in vals[-3:] if v is not None and v > 0]
        if len(last3) == 3 and last3[0] > last3[1] > last3[2]:
            drop = (last3[2] - last3[0]) / last3[0] * 100
            if drop < -5:
                anomalies.append(('warning', f'3M downtrend: {drop:+.1f}%'))
        
        # Zero drop
        if last_v == 0 and prev_v > 0:
            anomalies.append(('critical', f'Dropped to zero (préc: {prev_v:,.0f})'))
        
        worst = 'critical' if any(a[0] == 'critical' for a in anomalies) else \
                'warning' if anomalies else 'ok'
        
        results.append({
            'site': site_name,
            'values': vals,
            'months': months,
            'month_m': month_m,
            'month_m1': month_m1,
            'val_m': last_v,
            'val_m1': prev_v,
            'evol_pct': evol_pct,
            'pr_m': pr_m,
            'pr_m1': pr_m1,
            'evol_pr': evol_pr,
            'anomalies': anomalies,
            'status': worst,
        })
    
    return results, total_dedup_m, total_m

# ─── SPARKLINE ────────────────────────────────────────────────────────────────

def sparkline(vals, months, status, height=100):
    color = {'critical': '#ff4444', 'warning': '#f5a623', 'ok': '#22d3a0'}.get(status, '#4b5568')
    fill = {'critical': 'rgba(255,68,68,0.08)', 'warning': 'rgba(245,166,35,0.08)',
            'ok': 'rgba(34,211,160,0.06)'}.get(status, 'rgba(75,85,99,0.05)')
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=months[:len(vals)], y=vals, mode='lines+markers',
        line=dict(color=color, width=2),
        marker=dict(size=[7 if i == len(vals)-1 else 0 for i in range(len(vals))], color=color),
        fill='tozeroy', fillcolor=fill,
        hovertemplate='%{x}: <b>%{y:,.0f}</b><extra></extra>',
        connectgaps=True,
    ))
    fig.update_layout(
        height=height, margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(showgrid=False, tickfont=dict(size=8, color='#4b5568'), nticks=4),
        yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.04)',
                   tickfont=dict(size=8, color='#4b5568'), tickformat='.2s'),
        showlegend=False,
    )
    return fig

# ─── MAIN PAGE ────────────────────────────────────────────────────────────────

def render_panel_checker_v2():
    st.markdown("""
    <style>
    .verdict-ok { background: linear-gradient(135deg,#0d2818,#0a3320); border:2px solid #22d3a0;
        border-radius:16px; padding:28px 36px; text-align:center; margin-bottom:24px; }
    .verdict-warn { background: linear-gradient(135deg,#2d1f00,#3d2a00); border:2px solid #f5a623;
        border-radius:16px; padding:28px 36px; text-align:center; margin-bottom:24px; }
    .verdict-fail { background: linear-gradient(135deg,#2d0000,#3d0808); border:2px solid #ff4444;
        border-radius:16px; padding:28px 36px; text-align:center; margin-bottom:24px; }
    .anom-crit { background:rgba(255,68,68,0.07); border-left:3px solid #ff4444;
        border-radius:6px; padding:8px 12px; margin:4px 0; font-size:12px; }
    .anom-warn { background:rgba(245,166,35,0.07); border-left:3px solid #f5a623;
        border-radius:6px; padding:8px 12px; margin:4px 0; font-size:12px; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("# 🔍 Panel Checker — Source Files Analysis")
    st.markdown("*Upload HAM files for automatic trend analysis*")
    st.divider()
    
    # ── UPLOAD ──
    uploaded = st.file_uploader(
        "Upload HAM files (.xlsx)",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Files: 1_analyse_evolution_panel, 2_analyse_performance, 3_1, 3_2, 4_1, 4_2, 5_2, 6_..."
    )
    
    if not uploaded:
        st.info("👆 Upload one or more HAM files to start the analysis")
        cols = st.columns(4)
        cols[0].info("📁 1_ Panel evolution")
        cols[1].info("📁 2_ Quality performance")
        cols[2].info("📁 5_2_ Grand Ouest")
        cols[3].info("📁 6_ New listings IDF")
        return
    
    # ── PARSE ──
    with st.spinner("⚙️ Analysing..."):
        all_sections = {}
        for f in uploaded:
            data = parse_ham_file(f.read(), f.name)
            all_sections.update(data)
    
    if not all_sections:
        st.error("No data found in the uploaded files.")
        return
    
    # ── COMPUTE ALL METRICS ──
    all_results = {}
    total_critical = 0
    total_warning = 0
    
    for key, section in all_sections.items():
        metrics, total_dedup, total = compute_metrics(
            section['sites'], section['totals'], section['months']
        )
        if metrics:
            crits = sum(1 for m in metrics if m['status'] == 'critical')
            warns = sum(1 for m in metrics if m['status'] == 'warning')
            total_critical += crits
            total_warning += warns
            all_results[key] = {
                **section,
                'metrics': metrics,
                'total_dedup_m': total_dedup,
                'total_m': total,
                'critical_count': crits,
                'warning_count': warns,
            }
    
    # ── VERDICT ──
    if total_critical > 0:
        verdict = 'REFUSED'
        st.markdown(f"""<div class="verdict-fail">
            <div style="font-size:2.5rem;font-weight:900;letter-spacing:4px;color:#ff4444">❌ REFUSED</div>
            <div style="color:#fca5a5;margin-top:8px">{total_critical} critical anomaly(ies) · {total_warning} warning(s)</div>
        </div>""", unsafe_allow_html=True)
    elif total_warning > 3:
        verdict = 'TO REVIEW'
        st.markdown(f"""<div class="verdict-warn">
            <div style="font-size:2.5rem;font-weight:900;letter-spacing:4px;color:#f5a623">⚠️ TO REVIEW</div>
            <div style="color:#fcd48a;margin-top:8px">{total_warning} warning(s) to review</div>
        </div>""", unsafe_allow_html=True)
    else:
        verdict = 'VALIDATED'
        st.markdown(f"""<div class="verdict-ok">
            <div style="font-size:2.5rem;font-weight:900;letter-spacing:4px;color:#22d3a0">✅ VALIDATED</div>
            <div style="color:#a0e9d4;margin-top:8px">No critical anomaly · {total_warning} minor warning(s)</div>
        </div>""", unsafe_allow_html=True)
    
    # ── STATS ──
    sections_with_issues = len([r for r in all_results.values() if r['critical_count'] + r['warning_count'] > 0])
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Sections analysed", len(all_results))
    m2.metric("Files", len(uploaded))
    m3.metric("Sections with anomalies", sections_with_issues)
    m4.metric("🔴 Criticals", total_critical)
    m5.metric("🟡 Warnings", total_warning)
    
    st.divider()
    
    # ── FILTERS ──
    col_f1, col_f2, col_f3 = st.columns([2, 2, 3])
    with col_f1:
        sev = st.radio("", ["Tout", "🔴 Critical", "🟡 Warnings"], horizontal=True)
    with col_f2:
        file_opts = ["All"] + sorted(set(r['file'] for r in all_results.values()))
        file_filter = st.selectbox("File", file_opts, label_visibility="collapsed")
    with col_f3:
        search = st.text_input("", placeholder="🔍 Search...", label_visibility="collapsed")
    
    # Apply filters
    filtered = dict(all_results)
    if sev == "🔴 Critical":
        filtered = {k: v for k, v in filtered.items() if v['critical_count'] > 0}
    elif sev == "🟡 Warnings":
        filtered = {k: v for k, v in filtered.items() if v['warning_count'] > 0}
    if file_filter != "All":
        filtered = {k: v for k, v in filtered.items() if v['file'] == file_filter}
    if search:
        q = search.lower()
        filtered = {k: v for k, v in filtered.items() if
                    q in v['label'].lower() or
                    any(q in m['site'].lower() for m in v['metrics'])}
    
    # Sort: critical first
    sorted_sections = sorted(filtered.items(), key=lambda x: (-x[1]['critical_count'], -x[1]['warning_count']))
    
    # Info banner when validated
    if verdict == 'VALIDATED' and total_warning > 0:
        st.info(f"ℹ️ Panel validated — {total_warning} warning(s) to monitor below")
    
    st.markdown(f"### 📋 {len(filtered)} section(s)")
    
    # ── SECTION CARDS ──
    for key, section in sorted_sections:
        has_crit = section['critical_count'] > 0
        icon = "🔴" if has_crit else "🟡" if section['warning_count'] > 0 else "✅"
        badge = f"{section['critical_count']}c · {section['warning_count']}a"
        
        with st.expander(
            f"{icon} **{section['sheet']}** — {section['label']}   `{badge}`",
            expanded=(has_crit and len(filtered) <= 8)
        ):
            # Month info
            months = section['months']
            if months:
                st.caption(f"📅 Period: {months[0]} → {months[-1]}  |  M = {months[-1]}  |  M-1 = {months[-2] if len(months) >= 2 else '?'}")
            
            # Summary table
            metrics = sorted(section['metrics'], key=lambda m: {'critical':3,'warning':2,'ok':1}.get(m['status'],0), reverse=True)
            
            # Group: anomalies first, then OK
            anom_metrics = [m for m in metrics if m['anomalies']]
            ok_metrics = [m for m in metrics if not m['anomalies']]
            
            # Show anomaly sites with chart
            if anom_metrics:
                cols = st.columns(min(3, len(anom_metrics)))
                for i, m in enumerate(anom_metrics[:6]):
                    with cols[i % 3]:
                        st_color = '#ff4444' if m['status'] == 'critical' else '#f5a623'
                        evol_str = f"{m['evol_pct']:+.1f}%" if m['evol_pct'] is not None else "—"
                        pr_str = f"PR: {m['pr_m']:.1f}%" if m['pr_m'] is not None else ""
                        
                        with st.container(border=True):
                            hc1, hc2 = st.columns([3, 1])
                            hc1.markdown(f"**{'🔴' if m['status']=='critical' else '🟡'} {m['site']}**")
                            hc2.markdown(f"<span style='color:{st_color};font-weight:700'>{evol_str}</span>", unsafe_allow_html=True)
                            
                            if m['values']:
                                fig = sparkline(m['values'], m['months'], m['status'])
                                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
                            
                            # Metrics row
                            mc1, mc2, mc3 = st.columns(3)
                            mc1.caption(f"M: **{m['val_m']:,.0f}**")
                            mc2.caption(f"M-1: **{m['val_m1']:,.0f}**")
                            mc3.caption(pr_str)
                            
                            # Anomaly tags
                            for anom_type, anom_detail in m['anomalies']:
                                css = 'anom-crit' if anom_type == 'critical' else 'anom-warn'
                                ico = '🔴' if anom_type == 'critical' else '🟡'
                                st.markdown(f'<div class="{css}">{ico} {anom_detail}</div>', unsafe_allow_html=True)
            
            # Summary table for all sites
            if metrics:
                st.markdown("**Summary table**")
                table_data = []
                for m in metrics:
                    evol = f"{m['evol_pct']:+.1f}%" if m['evol_pct'] is not None else "—"
                    pr = f"{m['pr_m']:.1f}%" if m['pr_m'] is not None else "—"
                    evol_pr = f"{m['evol_pr']:+.2f}pp" if m['evol_pr'] is not None else "—"
                    status_icon = "🔴" if m['status']=='critical' else "🟡" if m['status']=='warning' else "✅"
                    table_data.append({
                        '': status_icon,
                        'Site': m['site'],
                        f"M ({m['month_m']})": f"{m['val_m']:,.0f}",
                        f"M-1 ({m['month_m1']})": f"{m['val_m1']:,.0f}",
                        'Evol %': evol,
                        'PR M': pr,
                        'Market Share Evol': evol_pr,
                        'Anomalies': ' | '.join(a[1] for a in m['anomalies']) if m['anomalies'] else '—',
                    })
                df = pd.DataFrame(table_data)
                st.dataframe(df, use_container_width=True, hide_index=True)
    
    # ── EXPORT ──
    if all_results:
        st.divider()
        st.markdown("### 📥 Export")
        rows = []
        for key, section in all_results.items():
            for m in section['metrics']:
                if m['anomalies']:
                    for anom_type, anom_detail in m['anomalies']:
                        rows.append({
                            'File': section['file'],
                            'Sheet': section['sheet'],
                            'Section': section['label'],
                            'Site': m['site'],
                            'Sévérité': '🔴 Critical' if anom_type=='critical' else '🟡 Avertissement',
                            'Anomaly': anom_detail,
                            'Month M': m['month_m'],
                            'Value M': m['val_m'],
                            'Value M-1': m['val_m1'],
                            'Evol %': f"{m['evol_pct']:+.1f}%" if m['evol_pct'] else '',
                            'PR M': f"{m['pr_m']:.1f}%" if m['pr_m'] else '',
                            'Market Share Evol': f"{m['evol_pr']:+.2f}pp" if m['evol_pr'] else '',
                        })
        if rows:
            df = pd.DataFrame(rows)
            csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("⬇️ Download CSV report", data=csv,
                             file_name=f"panel_anomalies_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                             mime='text/csv')




"""
Panel Checker Module — à ajouter dans app.py
Analyse tendances du Panel Checker Excel et détecte les anomalies.
"""

import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
import datetime
import io
from openpyxl import load_workbook

# ─── PARSING ──────────────────────────────────────────────────────────────────

SKIP_SITES = {
    'Total','Total Panel','Total Panel Dédupliqué','Total Panel Dédupliqué - Top 5 Sites',
    'Total Panel Dédupliqué  - Top 11 Sites','Total Panel Dédupliqué Marché',
    'Immobilier Notaire','Immonot','Site','Département','Région',
    'Total Panel Dedup','Total Dédupliqué',
}

def excel_date_str(v):
    if isinstance(v, datetime.datetime):
        return v.strftime('%b-%y')
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (int, float)) and 40000 < v < 50000:
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime('%b-%y')
    return str(v)

def parse_panel_checker(file_bytes):
    """Parse le fichier Panel Checker Excel et retourne toutes les tables."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    
    all_tables = []  # list of dicts
    tab_sheets = [s for s in wb.sheetnames if s.lower().startswith('tab')]
    
    for sheet_name in tab_sheets:
        ws = wb[sheet_name]
        
        # Find all Site/Département header rows (tables)
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if b not in ('Site', 'Département', 'Région'):
                continue
            
            # Check if followed by month-like values in cols 3+
            month_cols = []
            months = []
            for c in range(3, min(30, ws.max_column + 1)):
                v = ws.cell(r, c).value
                if v is None:
                    continue
                s = excel_date_str(v)
                # Accept if it looks like a month label
                if (isinstance(v, datetime.datetime) or
                    (isinstance(v, (int, float)) and 40000 < v < 50000) or
                    (isinstance(v, str) and any(m in v.lower() for m in
                     ['-25', '-26', '-24', '-23', 'avr', 'mars', 'mai', 'juin',
                      'juil', 'août', 'sep', 'oct', 'nov', 'dec', 'jan', 'fev']))):
                    month_cols.append(c)
                    months.append(s)
            
            if len(months) < 3:
                continue
            
            # Get table label (look above for title)
            table_label = None
            for tr in range(r - 1, max(0, r - 5), -1):
                v = ws.cell(tr, 2).value
                if v and isinstance(v, str) and len(v.strip()) > 3 and v.strip() != 'Site':
                    table_label = v.strip()
                    break
            if not table_label:
                table_label = f"Table r{r}"
            
            # Get section label (look further up for x.x headers)
            section_label = None
            for tr in range(r - 1, max(0, r - 20), -1):
                v = ws.cell(tr, 2).value
                if v and isinstance(v, str) and len(v) > 2:
                    # Check if it's a section header like "1.1 Total", "2.3 ..."
                    if v[0].isdigit() and '.' in v[:4]:
                        section_label = v.strip()
                        break
            
            # Extract site data rows
            sites = {}
            for dr in range(r + 1, min(r + 50, ws.max_row + 1)):
                site_name = ws.cell(dr, 2).value
                if not site_name or not isinstance(site_name, str):
                    break
                if site_name.strip() in SKIP_SITES:
                    continue
                if site_name.strip() == '' or site_name.strip() == 'Site':
                    break
                
                vals = []
                for c in month_cols:
                    v = ws.cell(dr, c).value
                    if isinstance(v, (int, float)):
                        vals.append(float(v))
                    elif isinstance(v, str):
                        try:
                            vals.append(float(v.replace(',', '.').replace(' ', '')))
                        except:
                            vals.append(None)
                    else:
                        vals.append(None)
                
                real = [v for v in vals if v is not None and v > 0]
                if len(real) >= 2:
                    sites[site_name.strip()] = vals
            
            if sites:
                all_tables.append({
                    'sheet': sheet_name,
                    'section': section_label or sheet_name,
                    'label': table_label,
                    'months': months,
                    'sites': sites,
                    'header_row': r,
                })
    
    return all_tables

# ─── ANOMALY DETECTION ────────────────────────────────────────────────────────

def detect_anomalies(vals, months, site_name, table_label):
    """
    Détecte les anomalies de tendance pour une série temporelle.
    Retourne une liste d'anomalies avec niveau de sévérité.
    """
    anomalies = []
    MIN_VAL = 5  # ignore placeholder values
    real_pairs = [(i, v) for i, v in enumerate(vals) if v is not None and v > MIN_VAL]
    if len(real_pairs) < 3:
        return anomalies
    
    indices, values = zip(*real_pairs)
    values = list(values)
    last_idx = indices[-1]
    last_val = values[-1]
    prev_val = values[-2] if len(values) >= 2 else None
    
    # Skip placeholder values
    if last_val <= 5:
        return anomalies
    
    # 1. MoM variation (M vs M-1)
    if prev_val is not None and prev_val > 0:
        mom = (last_val - prev_val) / prev_val * 100
        if mom <= -30:
            anomalies.append({
                'type': 'Monthly change',
                'severity': 'critical',
                'detail': f"{mom:+.1f}% vs mois précédent",
                'month': months[last_idx] if last_idx < len(months) else '?',
                'value': last_val,
                'prev_value': prev_val,
            })
        elif mom <= -15:
            anomalies.append({
                'type': 'Monthly change',
                'severity': 'warning',
                'detail': f"{mom:+.1f}% vs mois précédent",
                'month': months[last_idx] if last_idx < len(months) else '?',
                'value': last_val,
                'prev_value': prev_val,
            })
        elif mom >= 50:
            anomalies.append({
                'type': 'Hausse anormale MoM',
                'severity': 'warning',
                'detail': f"{mom:+.1f}% vs mois précédent",
                'month': months[last_idx] if last_idx < len(months) else '?',
                'value': last_val,
                'prev_value': prev_val,
            })
    
    # 2. Crash vs historical max
    if len(values) >= 4:
        hist_max = max(values[:-1])
        if hist_max > 0:
            pct_from_max = (last_val - hist_max) / hist_max * 100
            if pct_from_max <= -40:
                anomalies.append({
                    'type': 'Crash vs historique',
                    'severity': 'critical',
                    'detail': f"{pct_from_max:+.1f}% vs max historique ({hist_max:,.0f})",
                    'month': months[last_idx] if last_idx < len(months) else '?',
                    'value': last_val,
                    'prev_value': hist_max,
                })
    
    # 3. Z-score anomaly (statistical outlier)
    if len(values) >= 4:
        hist = values[:-1]
        mean = np.mean(hist)
        std = np.std(hist)
        if std > 0:
            z = (last_val - mean) / std
            if z < -2.5:
                anomalies.append({
                    'type': 'Anomalie statistique (Z-score)',
                    'severity': 'warning' if z > -3 else 'critical',
                    'detail': f"Z-score={z:.2f} (moyenne hist.={mean:,.0f})",
                    'month': months[last_idx] if last_idx < len(months) else '?',
                    'value': last_val,
                    'prev_value': mean,
                })
    
    # 4. Sustained downtrend (3+ months declining)
    if len(values) >= 4:
        last_3 = values[-3:]
        if all(last_3[i] > last_3[i+1] for i in range(len(last_3)-1)):
            total_drop = (last_3[-1] - last_3[0]) / last_3[0] * 100 if last_3[0] > 0 else 0
            if total_drop < -10:
                anomalies.append({
                    'type': 'Sustained downtrend',
                    'severity': 'warning',
                    'detail': f"3 mois consécutifs en baisse ({total_drop:+.1f}% sur 3M)",
                    'month': months[last_idx] if last_idx < len(months) else '?',
                    'value': last_val,
                    'prev_value': last_3[0],
                })
    
    # 5. Value = 0 when it wasn't before
    if last_val == 0 and len(values) >= 2:
        prev_nonzero = [v for v in values[:-1] if v is not None and v > 0]
        if len(prev_nonzero) >= 2:
            anomalies.append({
                'type': 'Valeur tombée à zéro',
                'severity': 'critical',
                'detail': f"Valeur précédente: {prev_nonzero[-1]:,.0f}",
                'month': months[last_idx] if last_idx < len(months) else '?',
                'value': 0,
                'prev_value': prev_nonzero[-1],
            })
    
    return anomalies

def analyze_all_tables(tables):
    """Analyse toutes les tables et retourne un résumé des anomalies."""
    results = []
    
    for table in tables:
        table_anomalies = []
        months = table['months']
        
        for site_name, vals in table['sites'].items():
            site_anomalies = detect_anomalies(vals, months, site_name, table['label'])
            if site_anomalies:
                for a in site_anomalies:
                    table_anomalies.append({
                        'site': site_name,
                        **a
                    })
        
        if table_anomalies:
            results.append({
                'sheet': table['sheet'],
                'section': table['section'],
                'label': table['label'],
                'months': months,
                'sites': table['sites'],
                'anomalies': table_anomalies,
                'critical_count': sum(1 for a in table_anomalies if a['severity'] == 'critical'),
                'warning_count': sum(1 for a in table_anomalies if a['severity'] == 'warning'),
            })
    
    return results

# ─── VERDICT ──────────────────────────────────────────────────────────────────

def compute_verdict(analysis_results):
    total_critical = sum(r['critical_count'] for r in analysis_results)
    total_warning = sum(r['warning_count'] for r in analysis_results)
    
    if total_critical > 0:
        return 'REFUSED', total_critical, total_warning
    elif total_warning > 3:
        return 'TO REVIEW', total_critical, total_warning
    else:
        return 'VALIDATED', total_critical, total_warning

# ─── SPARKLINE ────────────────────────────────────────────────────────────────

def mini_chart(vals, months, has_critical=False, has_warning=False, height=100):
    color = '#ff4444' if has_critical else '#f5a623' if has_warning else '#22d3a0'
    fill = 'rgba(255,68,68,0.1)' if has_critical else 'rgba(245,166,35,0.08)' if has_warning else 'rgba(34,211,160,0.06)'
    
    # Mark last point
    marker_sizes = [0] * len(vals)
    marker_colors = [color] * len(vals)
    if vals:
        marker_sizes[-1] = 8
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=months[:len(vals)], y=vals,
        mode='lines+markers',
        line=dict(color=color, width=2),
        marker=dict(size=marker_sizes, color=marker_colors),
        fill='tozeroy', fillcolor=fill,
        hovertemplate='%{x}: <b>%{y:,.0f}</b><extra></extra>',
        connectgaps=True,
    ))
    fig.update_layout(
        height=height,
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(showgrid=False, tickfont=dict(size=8, color='#4b5568'), nticks=4, showline=False),
        yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.04)',
                   tickfont=dict(size=8, color='#4b5568'), tickformat='.2s', showline=False),
        showlegend=False,
    )
    return fig

# ─── MAIN PAGE ────────────────────────────────────────────────────────────────

def render_panel_checker_page():
    st.markdown("""
    <style>
    .verdict-ok {
        background: linear-gradient(135deg, #0d2818 0%, #0a3320 100%);
        border: 2px solid #22d3a0;
        border-radius: 16px;
        padding: 28px 36px;
        text-align: center;
        margin-bottom: 24px;
    }
    .verdict-warn {
        background: linear-gradient(135deg, #2d1f00 0%, #3d2a00 100%);
        border: 2px solid #f5a623;
        border-radius: 16px;
        padding: 28px 36px;
        text-align: center;
        margin-bottom: 24px;
    }
    .verdict-fail {
        background: linear-gradient(135deg, #2d0000 0%, #3d0808 100%);
        border: 2px solid #ff4444;
        border-radius: 16px;
        padding: 28px 36px;
        text-align: center;
        margin-bottom: 24px;
    }
    .verdict-title {
        font-size: 2.5rem;
        font-weight: 900;
        letter-spacing: 4px;
        margin-bottom: 8px;
    }
    .anomaly-card {
        background: rgba(255,68,68,0.06);
        border-left: 3px solid #ff4444;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 6px 0;
        font-size: 13px;
    }
    .warning-card {
        background: rgba(245,166,35,0.06);
        border-left: 3px solid #f5a623;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 6px 0;
        font-size: 13px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("# 🔍 Panel Checker — Trend Analysis")
    st.markdown("*Détection automatique des anomalies et validation du panel*")
    st.divider()

    # Upload
    uploaded = st.file_uploader(
        "Charger le fichier Panel Checker (.xlsx)",
        type=['xlsx'],
        help="Fichier Panel_checker_-_LBC_-_IMMO_FR_-_Resales_*.xlsx"
    )

    if not uploaded:
        st.info("👆 Chargez votre fichier Panel Checker pour lancer l'analyse")
        
        # Show example of what we detect
        st.markdown("### Ce que nous détectons :")
        c1, c2, c3, c4 = st.columns(4)
        c1.error("🔴 Monthly change\n>30% drop")
        c2.warning("🟡 Monthly change\n>15% drop")
        c3.error("🔴 Crash historique\n>40% vs max")
        c4.warning("🟡 Downtrend\n3 consecutive months")
        return

    with st.spinner("⚙️ Analysing... Parsing des tableaux et détection des anomalies..."):
        file_bytes = uploaded.read()
        tables = parse_panel_checker(file_bytes)
        analysis = analyze_all_tables(tables)
        verdict, n_critical, n_warning = compute_verdict(analysis)

    # VERDICT
    if verdict == 'VALIDATED':
        st.markdown(f"""
        <div class="verdict-ok">
            <div class="verdict-title" style="color:#22d3a0">✅ VALIDATED</div>
            <div style="color:#a0e9d4;font-size:1rem">Aucune anomalie critique détectée</div>
            <div style="color:#6ee7c9;font-size:0.85rem;margin-top:8px">{n_warning} minor warningineurs</div>
        </div>
        """, unsafe_allow_html=True)
    elif verdict == 'TO REVIEW':
        st.markdown(f"""
        <div class="verdict-warn">
            <div class="verdict-title" style="color:#f5a623">⚠️ TO REVIEW</div>
            <div style="color:#fcd48a;font-size:1rem">Several warnings detected</div>
            <div style="color:#fbbf24;font-size:0.85rem;margin-top:8px">{n_warning} warnings · {n_critical} critiques</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="verdict-fail">
            <div class="verdict-title" style="color:#ff4444">❌ REFUSED</div>
            <div style="color:#fca5a5;font-size:1rem">Critical anomalies detected — review required</div>
            <div style="color:#f87171;font-size:0.85rem;margin-top:8px">{n_critical} critiques · {n_warning} avertissements</div>
        </div>
        """, unsafe_allow_html=True)

    # Summary stats
    total_tables = len(tables)
    tables_with_issues = len(analysis)
    all_sites_checked = sum(len(t['sites']) for t in tables)

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Tableaux analysés", total_tables)
    m2.metric("Sites / séries", all_sites_checked)
    m3.metric("Tableaux avec anomalies", tables_with_issues)
    m4.metric("🔴 Anomalies critiques", n_critical)
    m5.metric("🟡 Warnings", n_warning)

    if not analysis:
        st.success("✅ Aucune anomalie détectée dans l'ensemble des tableaux.")
        return

    st.divider()

    # ── INFO: always show warnings even if VALIDÉ ──
    if verdict == 'VALIDATED' and n_warning > 0:
        st.info(f"ℹ️ Pas d'anomalie critique — mais {n_warning} avertissement(s) détecté(s) ci-dessous. À surveiller.")
    elif verdict == 'TO REVIEW':
        st.warning(f"⚠️ {n_warning} avertissements détectés — vérifiez les tableaux ci-dessous avant validation.")
    else:
        st.error(f"❌ {n_critical} anomalie(s) critique(s) à corriger — le panel ne peut pas être validé en l'état.")

    # ── FILTERS ──
    col_f1, col_f2, col_f3 = st.columns([2, 2, 3])
    with col_f1:
        sev_filter = st.radio(
            "Sévérité",
            ["Tout", "🔴 Critical seulement", "🟡 Warnings seulement"],
            index=0,
            horizontal=True,
            label_visibility="collapsed"
        )
    with col_f2:
        sheet_filter = st.selectbox(
            "Sheet",
            ["Tous les sheets"] + sorted(set(r['sheet'] for r in analysis)),
            label_visibility="collapsed"
        )
    with col_f3:
        search_q = st.text_input("Rechercher site ou tableau...", label_visibility="collapsed",
                                  placeholder="🔍 Rechercher site ou tableau...")

    # Apply filters
    filtered = analysis
    if sheet_filter != "Tous les sheets":
        filtered = [r for r in filtered if r['sheet'] == sheet_filter]
    if sev_filter == "🔴 Critical seulement":
        filtered = [r for r in filtered if r['critical_count'] > 0]
    elif sev_filter == "🟡 Warnings seulement":
        filtered = [r for r in filtered if r['warning_count'] > 0]
    # "Tout" shows everything — warnings always visible even if VALIDÉ
    if search_q:
        q = search_q.lower()
        filtered = [r for r in filtered if
                    q in r['label'].lower() or
                    any(q in a['site'].lower() for a in r['anomalies'])]

    # Sort: critical first
    filtered = sorted(filtered, key=lambda r: (-r['critical_count'], -r['warning_count']))

    st.markdown(f"### 📋 {len(filtered)} tableau(x) avec anomalies")

    # ── ANOMALY TABLE CARDS ──
    for result in filtered:
        has_crit = result['critical_count'] > 0
        icon = "🔴" if has_crit else "🟡"
        badge = f"{result['critical_count']} critique(s)" if has_crit else f"{result['warning_count']} avertissement(s)"

        with st.expander(f"{icon} **{result['sheet']}** — {result['label']}   `{badge}`",
                         expanded=(len(filtered) <= 5)):

            # Group anomalies by site
            by_site = {}
            for a in result['anomalies']:
                by_site.setdefault(a['site'], []).append(a)

            # For each affected site: show chart + anomalies
            for site, anoms in sorted(by_site.items(),
                                       key=lambda x: -sum(1 for a in x[1] if a['severity']=='critical')):
                site_vals = result['sites'].get(site, [])
                site_crit = any(a['severity'] == 'critical' for a in anoms)
                site_warn = any(a['severity'] == 'warning' for a in anoms)

                col_chart, col_info = st.columns([3, 2])
                with col_chart:
                    st.markdown(f"**{'🔴' if site_crit else '🟡'} {site}**")
                    if site_vals and any(v is not None and v > 0 for v in site_vals):
                        fig = mini_chart(site_vals, result['months'], site_crit, site_warn, height=110)
                        st.plotly_chart(fig, use_container_width=True,
                                        config={'displayModeBar': False})

                with col_info:
                    st.markdown("")
                    for a in anoms:
                        if a['severity'] == 'critical':
                            st.markdown(
                                f'<div class="anomaly-card">🔴 <b>{a["type"]}</b><br>'
                                f'<span style="color:#aaa">{a["detail"]}</span><br>'
                                f'<span style="color:#888;font-size:11px">{a["month"]} · val={a["value"]:,.0f}</span></div>',
                                unsafe_allow_html=True
                            )
                        else:
                            st.markdown(
                                f'<div class="warning-card">🟡 <b>{a["type"]}</b><br>'
                                f'<span style="color:#aaa">{a["detail"]}</span><br>'
                                f'<span style="color:#888;font-size:11px">{a["month"]} · val={a["value"]:,.0f}</span></div>',
                                unsafe_allow_html=True
                            )

                st.markdown("---")

    # ── FULL ANOMALY EXPORT TABLE ──
    if analysis:
        st.divider()
        st.markdown("### 📥 Export — Toutes les anomalies")

        rows = []
        for result in analysis:
            for a in result['anomalies']:
                rows.append({
                    'Sheet': result['sheet'],
                    'Section': result['section'],
                    'Tableau': result['label'],
                    'Site': a['site'],
                    'Sévérité': '🔴 Critical' if a['severity'] == 'critical' else '🟡 Avertissement',
                    'Type': a['type'],
                    'Détail': a['detail'],
                    'Mois': a['month'],
                    'Valeur': a.get('value', ''),
                })

        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, height=300)

        # Download
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "⬇️ Download CSV report",
            data=csv,
            file_name=f"panel_anomalies_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime='text/csv'
        )



st.set_page_config(
    page_title="IMMO FR — Panel Intelligence",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.block-container{padding-top:1.5rem}
div[data-testid="metric-container"]{background:#0f1117;border:1px solid #1c2030;border-radius:8px;padding:12px 16px}
div[data-testid="metric-container"] label{font-size:11px!important;color:#7c879e!important;text-transform:uppercase;letter-spacing:.5px}
.stPlotlyChart{margin-bottom:0}
</style>
""", unsafe_allow_html=True)

# ── DATA DIRECTORY ────────────────────────────────────────────────────────
DATA_DIR = Path("panel_data")
DATA_DIR.mkdir(exist_ok=True)

# ── HELPERS ───────────────────────────────────────────────────────────────
def excel_date(v):
    if isinstance(v, datetime.datetime): return v.strftime('%b-%y')
    if isinstance(v, (int,float)) and 40000 < v < 50000:
        return (datetime.datetime(1899,12,30)+datetime.timedelta(days=int(v))).strftime('%b-%y')
    if isinstance(v, str): return v[:7]
    return str(v)

SKIP = ['Total','Totaux','total','TOTAL','Total Panel','Total Dédupliqué',
        'Immobilier','Immonot','Site','Région','Département','CONTRÔLE',
        'contrôle','somme','Par métier','Statistiques','Pros identifiés',
        'Pros à identifier','Total général','Annonces incomplètes']

def is_data_row(b):
    if not b or not isinstance(b, str) or len(str(b).strip()) < 2: return False
    return not any(str(b).startswith(s) for s in SKIP)

def compute_metrics_crawling(vals):
    valid = [(i,x) for i,x in enumerate(vals) if x is not None and isinstance(x,(int,float)) and x > 0]
    if len(valid) < 2: return None, None, False
    last_v, prev_v = valid[-1][1], valid[-2][1]
    max_v = max(x for _,x in valid)
    pct_var      = round((last_v/prev_v-1)*100, 1) if prev_v > 0 else None
    pct_from_max = round((last_v/max_v -1)*100, 1) if max_v  > 0 else None
    crashed      = (last_v/max_v) < 0.4 if max_v > 0 else False
    return pct_var, pct_from_max, crashed

def get_status(p, crashed, pfm):
    if crashed: return 'err'
    if pfm is not None and pfm < -50: return 'err'
    if p is None: return 'na'
    if p < -20: return 'err'
    if p < -10: return 'warn'
    return 'ok'

def fmt_pct(p, crashed, pfm):
    if crashed and pfm is not None: return f"MAX: {pfm:+.0f}%"
    if p is None: return "—"
    return f"{p:+.1f}%"

def fmt_num(n):
    if n is None: return "—"
    if n >= 1e6: return f"{n/1e6:.2f}M"
    if n >= 1e3: return f"{n/1e3:.0f}k"
    return f"{int(n):,}".replace(",", " ")

def extract_sheet(ws, sheet_name):
    sections = []
    site_rows = []
    for r in range(1, ws.max_row+1):
        for c in range(1, 4):
            if ws.cell(r, c).value == 'Site':
                has_m = False
                for cc in range(c+1, min(c+15, ws.max_column+1)):
                    v = ws.cell(r, cc).value
                    if isinstance(v, datetime.datetime): has_m=True; break
                    if isinstance(v, str) and any(m in v.lower() for m in ['mars','avr','mai','juin']): has_m=True; break
                    if isinstance(v, (int,float)) and 40000 < v < 50000: has_m=True; break
                if has_m: site_rows.append((r, c))
                break
    if not site_rows: return sections

    hr0, hc0 = site_rows[0]
    months, month_cols = [], []
    for c in range(hc0+1, ws.max_column+1):
        v = ws.cell(hr0, c).value
        if v is None: continue
        lbl = excel_date(v)
        if any(m in lbl.lower() for m in ['25','26','24','23']) or isinstance(v, datetime.datetime):
            months.append(lbl); month_cols.append(c)
        if len(months) >= 13: break
    if not months: return sections

    for i, (hdr_r, hdr_c) in enumerate(site_rows):
        next_hdr = site_rows[i+1][0] if i+1 < len(site_rows) else ws.max_row+1
        title = None
        for tr in range(hdr_r-1, max(0, hdr_r-6), -1):
            v = ws.cell(tr, hdr_c).value
            if v and isinstance(v, str) and len(v.strip()) > 3 and v != 'Site':
                title = v.strip(); break
        full_title = f"{sheet_name} — {title}" if title else f"{sheet_name} — Section {i+1}"
        site_data = []
        for r in range(hdr_r+1, next_hdr):
            b = ws.cell(r, hdr_c).value
            if not is_data_row(b): continue
            vals = []
            for c in month_cols:
                v = ws.cell(r, c).value
                if isinstance(v, (int,float)): vals.append(float(v))
                elif isinstance(v, str) and v.strip() == '-': vals.append(0.0)
                else: vals.append(None)
            real = [v for v in vals if v is not None and v > 0]
            if not real or max(real) < 1 or all(0 < v < 2 for v in real): continue
            pv, pfm, cr = compute_metrics_crawling(vals)
            site_data.append({
                'name': str(b).strip(), 'values': vals,
                'pct_var': pv, 'pct_from_max': pfm,
                'crashed': cr, 'status': get_status(pv, cr, pfm)
            })
        if site_data:
            sections.append({'title': full_title, 'months': months, 'site_data': site_data})
    return sections

SKIP_SHEETS = {'Intro', '2.6 DPE-GES'}
FILE_LABELS = {
    '1_': '1 — Stock Annonces',
    '2_': '2 — Fraîcheur & Qualité',
    '5_2': '5-2 — Grand Ouest',
    '6_': '6 — Nouvelles IDF',
    '3_1': '3.1 — Annonceurs Pro',
    '3_2': '3.2 — Géo Pros',
    '4_1': '4.1 — Stats Géo',
    '4_2': '4.2 — Exclusivité',
}

def get_file_label(filename):
    fn = filename.lower()
    for prefix, label in FILE_LABELS.items():
        if fn.startswith(prefix): return label
    return filename[:25]

@st.cache_data(show_spinner=False)
def process_uploaded_files(file_contents_dict):
    """Process Excel files and return structured data. Cached by content hash."""
    import io
    result = {}
    for fname, content in file_contents_dict.items():
        try:
            wb = load_workbook(io.BytesIO(content), data_only=True)
            label = get_file_label(fname)
            sections = []
            for sheet_name in wb.sheetnames:
                if sheet_name in SKIP_SHEETS: continue
                secs = extract_sheet(wb[sheet_name], sheet_name)
                sections.extend(secs)
            if sections:
                key = fname.split('_')[0] + '_' + fname.split('_')[1] if '_' in fname else fname[:8]
                key = key.replace('.xlsx','').replace(' ','_').lower()
                result[key] = {
                    'label': label,
                    'months': sections[0]['months'],
                    'sections': sections
                }
        except Exception as e:
            st.warning(f"⚠️ {fname}: {e}")
    return result

def group_by_site(sections):
    RK = {'err':3,'warn':2,'ok':1,'na':0}
    sites = {}
    for sec in sections:
        for si in sec['site_data']:
            name = si['name']
            if name not in sites:
                sites[name] = {'name': name, 'values': None, 'sections': [],
                               'worst_status': 'na', 'worst_pct': None, 'worst_pfm': None}
            s = sites[name]
            if s['values'] is None:
                real = [v for v in si['values'] if v is not None and v > 0]
                if len(real) >= 6: s['values'] = si['values']
            st = si['status']
            s['sections'].append({
                'title': sec['title'], 'pct_var': si['pct_var'],
                'pct_from_max': si['pct_from_max'], 'crashed': si['crashed'],
                'status': st, 'values': si['values']
            })
            if RK.get(st,0) > RK.get(s['worst_status'],0):
                s['worst_status'] = st
                s['worst_pct'] = si['pct_var']
                s['worst_pfm'] = si['pct_from_max']
    return sorted(sites.values(), key=lambda s: -{'err':3,'warn':2,'ok':1,'na':0}.get(s['worst_status'],0))

def make_sparkline(vals, months, status, height=130):
    colors = {'err':'#ff4444','warn':'#f5a623','ok':'#4f8cff','na':'#4b5568'}
    c = colors.get(status, '#4b5568')
    fills = {'err':'rgba(255,68,68,0.08)','warn':'rgba(245,166,35,0.08)','ok':'rgba(79,140,255,0.06)','na':'rgba(75,85,99,0.05)'}
    f = fills.get(status,'rgba(75,85,99,0.05)')
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=months, y=vals, mode='lines+markers',
        line=dict(color=c, width=2),
        marker=dict(size=[7 if i==len(vals)-1 else 0 for i in range(len(vals))], color=c),
        fill='tozeroy', fillcolor=f,
        hovertemplate='%{x}: <b>%{y:,.0f}</b><extra></extra>',
        connectgaps=True,
    ))
    fig.update_layout(
        height=height, margin=dict(l=0,r=0,t=4,b=0),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(showgrid=False, tickfont=dict(size=9,color='#4b5568'), nticks=5, showline=False),
        yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)',
                  tickfont=dict(size=9,color='#4b5568'), tickformat='.2s', showline=False),
        showlegend=False,
    )
    return fig

def save_month(month_key, data):
    path = DATA_DIR / f"{month_key}.json"
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

def load_all_months():
    months = {}
    for p in sorted(DATA_DIR.glob("*.json")):
        try:
            with open(p, encoding='utf-8') as f:
                months[p.stem] = json.load(f)
        except: pass
    return months

# ══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════

# ── PAGE NAVIGATION ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏠 IMMO FR Monitor")
    st.markdown("*Panel Intelligence Platform*")
    st.divider()
    page = st.radio(
        "Navigation",
        ["📊 Crawling Monitor", "✅ Panel Checker"],
        label_visibility="collapsed"
    )
    st.divider()

if page == "✅ Panel Checker":
    render_panel_checker_v2()
    st.stop()

# ── CRAWLING MONITOR ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Panel Monitor")
    st.markdown("*Crawling Trend Tracker*")
    st.divider()

    # ── UPLOAD NEW MONTH ──
    st.markdown("### ➕ Nouveau mois")
    uploaded = st.file_uploader(
        "Files Excel",
        type=['xlsx'],
        accept_multiple_files=True,
        label_visibility="collapsed",
        help="Déposez les fichiers 1_, 2_, 3_, 4_, 5_, 6_..."
    )

    month_name = st.text_input(
        "Nom du mois",
        placeholder="ex: avr-26",
        label_visibility="collapsed"
    )

    if st.button("💾 Enregistrer", type="primary",
                 disabled=not (uploaded and month_name.strip()),
                 use_container_width=True):
        with st.spinner(f"Traitement de {len(uploaded)} fichier(s)..."):
            file_dict = {f.name: f.read() for f in uploaded}
            data = process_uploaded_files(file_dict)
            if data:
                save_month(month_name.strip(), data)
                st.success(f"✅ {month_name} sauvegardé ({len(data)} fichiers)")
                st.rerun()
            else:
                st.error("Aucune donnée extraite.")

    st.divider()

    # ── SELECT MONTHS ──
    all_months = load_all_months()

    if not all_months:
        st.info("Aucun mois encore.\nChargez vos fichiers ci-dessus.")
        st.stop()

    st.markdown("### 📅 Sélectionner mois")
    month_keys = sorted(all_months.keys())

    selected_months = st.multiselect(
        "Mois",
        options=month_keys,
        default=[month_keys[-1]],
        label_visibility="collapsed",
        help="Sélectionnez un ou plusieurs mois pour comparer"
    )

    if not selected_months:
        st.warning("Sélectionnez au moins un mois.")
        st.stop()

    # Delete month
    with st.expander("🗑 Supprimer un mois"):
        del_month = st.selectbox("Mois à supprimer", month_keys, label_visibility="collapsed")
        if st.button("Supprimer", type="secondary"):
            (DATA_DIR / f"{del_month}.json").unlink(missing_ok=True)
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════
st.markdown("# 📊 Crawling Trend Monitor")

# Merge available file keys from selected months
all_file_keys = {}
for m in selected_months:
    for fk, fv in all_months[m].items():
        if fk not in all_file_keys:
            all_file_keys[fk] = fv['label']

if not all_file_keys:
    st.warning("Aucune donnée pour les mois sélectionnés.")
    st.stop()

# ── CONTROLS ──────────────────────────────────────────────────────────────
col_view, col_file, col_filt, col_srch = st.columns([2, 3, 3, 2])

with col_view:
    view = st.radio("Vue", ["Par Site", "Par Tableau"], horizontal=True, label_visibility="collapsed")

with col_file:
    selected_file = st.selectbox("File", list(all_file_keys.keys()),
                                 format_func=lambda k: all_file_keys[k],
                                 label_visibility="collapsed")

with col_filt:
    filt = st.radio("Filtre", ["Tous", "🔴 Erreurs", "🟡 Attention", "✅ OK"],
                    horizontal=True, label_visibility="collapsed")

with col_srch:
    search = st.text_input("Recherche", placeholder="Site...", label_visibility="collapsed")

st.divider()

# ── GET DATA ──────────────────────────────────────────────────────────────
ref_month = selected_months[-1]
tab_data  = all_months[ref_month].get(selected_file)

if not tab_data:
    st.warning(f"Fichier '{all_file_keys.get(selected_file)}' non disponible pour {ref_month}.")
    st.stop()

sections = tab_data['sections']
months   = tab_data['months']

# ── SINGLE MONTH ──────────────────────────────────────────────────────────
if len(selected_months) == 1:

    all_sites = group_by_site(sections)
    ec = sum(1 for s in all_sites if s['worst_status']=='err')
    wc = sum(1 for s in all_sites if s['worst_status']=='warn')
    oc = sum(1 for s in all_sites if s['worst_status'] in ('ok','na'))

    # Stats row
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Sites analysés", len(all_sites))
    m2.metric("🔴 Erreurs",    ec, delta=None)
    m3.metric("🟡 Attention",  wc, delta=None)
    m4.metric("✅ OK",         oc, delta=None)

    # Anomaly banner
    err_sites = [s for s in all_sites if s['worst_status']=='err']
    if err_sites:
        names = "  ·  ".join(
            f"**{s['name']}** ({fmt_pct(s['worst_pct'], any(x['crashed'] for x in s['sections']), s['worst_pfm'])})"
            for s in err_sites
        )
        st.error(f"⚠️ Anomalies détectées :  {names}")

    filt_map = {'Tous': None, '🔴 Erreurs': 'err', '🟡 Attention': 'warn', '✅ OK': 'ok'}
    filt_val = filt_map[filt]
    q = search.lower().strip()

    # ── PAR SITE ──
    if view == "Par Site":
        display = [s for s in all_sites if
                   (not q or q in s['name'].lower()) and
                   (not filt_val or s['worst_status']==filt_val or
                    (filt_val=='ok' and s['worst_status'] in ('ok','na')))]

        if not display:
            st.info("Aucun site pour ce filtre.")
        else:
            cols = st.columns(3)
            RK = {'err':3,'warn':2,'ok':1,'na':0}
            for i, site in enumerate(display):
                with cols[i % 3]:
                    st_ = site['worst_status']
                    cr  = any(s['crashed'] for s in site['sections'])
                    ps  = fmt_pct(site['worst_pct'], cr, site['worst_pfm'])
                    col_map = {'err':'🔴','warn':'🟡','ok':'🟢','na':'⚪'}
                    icon = col_map.get(st_,'⚪')

                    with st.container(border=True):
                        hc1, hc2 = st.columns([3,1])
                        hc1.markdown(f"**{icon} {site['name']}**")
                        clr = 'red' if st_=='err' else 'orange' if st_=='warn' else 'green'
                        hc2.markdown(f"<span style='color:{clr};font-weight:700'>{ps}</span>",
                                     unsafe_allow_html=True)

                        # Chart using worst section
                        sorted_secs = sorted(site['sections'], key=lambda s: -RK.get(s['status'],0))
                        chart_vals = sorted_secs[0]['values'] if sorted_secs else site['values']
                        if chart_vals:
                            st.plotly_chart(make_sparkline(chart_vals, months, st_),
                                           use_container_width=True, config={'displayModeBar':False})

                        # Section list
                        for sec in sorted_secs[:4]:
                            sst = sec['status']
                            sv  = fmt_pct(sec['pct_var'], sec['crashed'], sec['pct_from_max'])
                            ico = '🔴' if sst=='err' else '🟡' if sst=='warn' else '🟢'
                            short = sec['title'].split(' — ')[-1][:42]
                            st.caption(f"{ico} {short} — **{sv}**")

                        extra = sorted_secs[4:]
                        if extra:
                            with st.expander(f"+{len(extra)} autres tableaux"):
                                for sec in extra:
                                    sv = fmt_pct(sec['pct_var'], sec['crashed'], sec['pct_from_max'])
                                    ico = '🔴' if sec['status']=='err' else '🟡' if sec['status']=='warn' else '🟢'
                                    short = sec['title'].split(' — ')[-1][:42]
                                    st.caption(f"{ico} {short} — **{sv}**")

    # ── PAR TABLEAU ──
    else:
        sec_options = {i: s['title'] for i, s in enumerate(sections)}
        # Pre-select first section with errors
        default_idx = next((i for i,s in enumerate(sections)
                           if any(si['status']=='err' for si in s['site_data'])), 0)

        sel_idx = st.selectbox(
            "Tableau",
            options=list(sec_options.keys()),
            index=default_idx,
            format_func=lambda i: sec_options[i]
        )

        sec = sections[sel_idx]
        items = sec['site_data']

        filt_map = {'Tous': None, '🔴 Erreurs': 'err', '🟡 Attention': 'warn', '✅ OK': 'ok'}
        filt_val = filt_map[filt]
        q = search.lower().strip()

        RK = {'err':3,'warn':2,'ok':1,'na':0}
        display = sorted(
            [si for si in items if
             (not q or q in si['name'].lower()) and
             (not filt_val or si['status']==filt_val or (filt_val=='ok' and si['status'] in ('ok','na')))],
            key=lambda si: -RK.get(si['status'],0)
        )

        if not display:
            st.info("Aucun site pour ce filtre.")
        else:
            cols = st.columns(3)
            for i, si in enumerate(display):
                with cols[i % 3]:
                    st_ = si['status']
                    ps  = fmt_pct(si['pct_var'], si['crashed'], si['pct_from_max'])
                    icon= '🔴' if st_=='err' else '🟡' if st_=='warn' else '🟢'

                    with st.container(border=True):
                        hc1, hc2 = st.columns([3,1])
                        hc1.markdown(f"**{icon} {si['name']}**")
                        clr = 'red' if st_=='err' else 'orange' if st_=='warn' else 'green'
                        hc2.markdown(f"<span style='color:{clr};font-weight:700'>{ps}</span>",
                                     unsafe_allow_html=True)

                        if si['values']:
                            st.plotly_chart(
                                make_sparkline(si['values'], sec['months'], st_),
                                use_container_width=True, config={'displayModeBar':False}
                            )

                        real = [v for v in si['values'] if v is not None and v > 0]
                        if real:
                            sc1,sc2,sc3 = st.columns(3)
                            sc1.caption(f"Moy: {fmt_num(sum(real)/len(real))}")
                            sc2.caption(f"Min: {fmt_num(min(real))}")
                            sc3.caption(f"Max: {fmt_num(max(real))}")

# ── MULTI-MONTH COMPARISON ────────────────────────────────────────────────
else:
    st.markdown(f"### Comparaison — {' vs '.join(selected_months)}")

    # Site selector
    all_site_names = sorted({
        si['name']
        for m in selected_months
        for fdata in [all_months[m].get(selected_file)]
        if fdata
        for sec in fdata.get('sections', [])
        for si in sec['site_data']
    })

    if not all_site_names:
        st.warning("Aucun site commun entre les mois sélectionnés.")
        st.stop()

    sel_site = st.selectbox("Site à comparer", all_site_names)

    # Overlay chart
    fig = go.Figure()
    palette = ['#4f8cff','#ff4444','#22d3a0','#f5a623','#a78bfa','#34d399']

    for mi, month in enumerate(selected_months):
        mdata = all_months[month].get(selected_file)
        if not mdata: continue
        for sec in mdata['sections']:
            for si in sec['site_data']:
                if si['name'] == sel_site and si['values']:
                    real = [v for v in si['values'] if v is not None and v > 0]
                    if len(real) >= 4:
                        fig.add_trace(go.Scatter(
                            x=sec['months'], y=si['values'],
                            mode='lines+markers', name=month,
                            line=dict(color=palette[mi%len(palette)], width=2.5),
                            marker=dict(size=4),
                            hovertemplate=f'{month} · %{{x}}: <b>%{{y:,.0f}}</b><extra></extra>',
                            connectgaps=True,
                        ))
                        break

    fig.update_layout(
        height=380,
        title=dict(text=f"Tendance — {sel_site}", font=dict(size=14)),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(showgrid=False, tickfont=dict(size=10,color='#7c879e')),
        yaxis=dict(showgrid=True, gridcolor='rgba(128,128,128,0.1)',
                  tickfont=dict(size=10,color='#7c879e'), tickformat='.2s'),
        legend=dict(font=dict(color='#dde1ec'), bgcolor='rgba(0,0,0,0)',
                   orientation='h', yanchor='bottom', y=1.02),
    )
    st.plotly_chart(fig, use_container_width=True)

    # Summary table
    st.markdown("#### Tableau comparatif")
    rows = []
    for m in selected_months:
        mdata = all_months[m].get(selected_file)
        if not mdata: continue
        row = {'Mois': m}
        sites_in_month = {}
        for sec in mdata['sections']:
            for si in sec['site_data']:
                if si['name'] not in sites_in_month:
                    sites_in_month[si['name']] = si
        for sname in all_site_names[:12]:
            si = sites_in_month.get(sname)
            if si:
                row[sname] = fmt_pct(si['pct_var'], si['crashed'], si['pct_from_max'])
            else:
                row[sname] = '—'
        rows.append(row)

    if rows:
        df = pd.DataFrame(rows).set_index('Mois')
        st.dataframe(df, use_container_width=True)
