import streamlit as st
import plotly.graph_objects as go
import pandas as pd
from openpyxl import load_workbook
import datetime, json
from pathlib import Path

st.set_page_config(
    page_title="Panel Checker — Crawling Monitor",
    page_icon="📊",
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

def compute_metrics(vals):
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
            pv, pfm, cr = compute_metrics(vals)
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
with st.sidebar:
    st.markdown("## 📊 Panel Monitor")
    st.markdown("*Crawling Trend Tracker*")
    st.divider()

    # ── UPLOAD NEW MONTH ──
    st.markdown("### ➕ Nouveau mois")
    uploaded = st.file_uploader(
        "Fichiers Excel",
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
    selected_file = st.selectbox("Fichier", list(all_file_keys.keys()),
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
