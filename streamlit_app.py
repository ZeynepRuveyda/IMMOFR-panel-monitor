"""
Panel Checker QC — Sade kontrol uygulaması
Amaç: Panel Checker Excel dosyasındaki contrôle satırlarını oku,
      FALSE olanları listele ve mesaj taslağı oluştur.
"""

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import datetime

st.set_page_config(
    page_title="Panel Checker QC",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
  .block-container { padding-top: 1.5rem; max-width: 1100px; }
  .stAlert { border-radius: 6px; }
  table { font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# ─── CORE LOGIC ──────────────────────────────────────────────────────────────

TABS_TO_CHECK = [
    'tab 1', 'tab 2',
    'tab 3.1', 'tab 3.2', 'tab 3.2 Y-1',
    'tab 4.1.1 & 4.1.2', 'tab 4.1.3', 'tab 4.1.4',
    'tab 4.1.5&4.1.6', 'tab 4.1.7&4.1.8',
    'tab 4.2.1&4.2.3', 'tab 4.2.2&4.2.4&4.2.5',
    'tab 5', 'tab 5-2', 'tab 5-2 Y-1'
]

def find_row_context(ws, row_num, search_range=8):
    """Satırın bağlamını bul: önceki veri satırlarına bakarak ne kontrolü olduğunu anla."""
    # Önce aynı satırda col B'ye bak
    col_b = ws.cell(row_num, 2).value
    if col_b and isinstance(col_b, str) and len(col_b.strip()) > 2:
        return col_b.strip()
    # Önceki satırlarda başlık ara
    for r in range(row_num - 1, max(0, row_num - search_range), -1):
        v = ws.cell(r, 2).value
        if v and isinstance(v, str) and len(v.strip()) > 3:
            return v.strip()
    return ""

def get_false_column_context(ws, row_num, false_col):
    """FALSE olan kolonun ne olduğunu anla (başlık satırlarına bak)."""
    for header_row in range(1, min(row_num, 10)):
        v = ws.cell(header_row, false_col).value
        if v and isinstance(v, str) and len(v.strip()) > 1:
            return v.strip()
    # Birkaç satır önce aynı kolonda değer ara
    for r in range(row_num - 1, max(0, row_num - 15), -1):
        v = ws.cell(r, false_col).value
        if v and isinstance(v, str) and len(v.strip()) > 1:
            return v.strip()
    return f"col {false_col}"

def scan_panel_checker(wb):
    """Panel Checker'daki tüm contrôle satırlarını tara, FALSE olanları döndür."""
    results = []  # {'tab', 'row', 'label', 'context', 'false_cols', 'status'}
    
    for tab_name in TABS_TO_CHECK:
        if tab_name not in wb.sheetnames:
            continue
        ws = wb[tab_name]
        
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if not val or not isinstance(val, str):
                    continue
                if 'contr' not in val.lower():
                    continue
                
                # Bu satırdaki boolean False değerleri bul
                false_cols = []
                true_count = 0
                for c in ws.iter_cols(min_row=cell.row, max_row=cell.row,
                                       min_col=1, max_col=ws.max_column):
                    for cell2 in c:
                        if cell2.value is True:
                            true_count += 1
                        elif cell2.value is False:
                            col_ctx = get_false_column_context(ws, cell.row, cell2.column)
                            false_cols.append({
                                'coord': cell2.coordinate,
                                'col': cell2.column,
                                'context': col_ctx
                            })
                
                # Kontrol satırı bulundu - kaydet
                row_context = find_row_context(ws, cell.row)
                
                results.append({
                    'tab': tab_name,
                    'row': cell.row,
                    'label': val.strip(),
                    'context': row_context,
                    'false_cols': false_cols,
                    'true_count': true_count,
                    'status': 'FALSE' if false_cols else 'TRUE'
                })
    
    return results

def generate_message(false_items, month_label="Avril 2026"):
    """FALSE olan kontroller için mesaj taslağı oluştur."""
    if not false_items:
        return None
    
    lines = [
        f"Bonjour,",
        f"",
        f"En contrôlant le Panel Checker pour {month_label}, j'ai détecté les erreurs suivantes :",
        f"",
    ]
    
    # Tab bazında grupla
    by_tab = {}
    for item in false_items:
        tab = item['tab']
        if tab not in by_tab:
            by_tab[tab] = []
        by_tab[tab].append(item)
    
    for tab, items in by_tab.items():
        lines.append(f"📋 **{tab}**")
        for item in items:
            false_details = []
            for fc in item['false_cols']:
                ctx = fc['context']
                if ctx and ctx not in ('None', ''):
                    false_details.append(ctx)
                else:
                    false_details.append(fc['coord'])
            
            detail_str = ", ".join(false_details[:5])
            if len(false_details) > 5:
                detail_str += f" (+{len(false_details)-5} autres)"
            
            lines.append(f"  • {item['label']} (ligne {item['row']}) → FALSE : {detail_str}")
        lines.append("")
    
    lines += [
        "Merci de vérifier et corriger ces points avant validation.",
        "",
        "Cordialement"
    ]
    
    return "\n".join(lines)

# ─── UI ──────────────────────────────────────────────────────────────────────

st.markdown("# 🔍 Panel Checker QC")
st.markdown("Télécharge le **Panel Checker** pour voir instantanément quels contrôles sont en **FALSE**.")
st.divider()

uploaded = st.file_uploader(
    "📂 Panel Checker (.xlsx)",
    type=["xlsx"],
    help="Glisse le fichier Panel_checker_-_LBC_-_IMMO_FR_-_Resales_-_*.xlsx ici"
)

month_label = st.text_input("Mois (pour le message)", value="Avril 2026")

if not uploaded:
    st.info("👆 Upload the Panel Checker file to start.")
    st.stop()

# ─── ANALYSE ─────────────────────────────────────────────────────────────────

with st.spinner("Scanning contrôle rows..."):
    wb = load_workbook(io.BytesIO(uploaded.read()), data_only=True)
    results = scan_panel_checker(wb)

if not results:
    st.error("No contrôle rows found. Check that this is the correct file.")
    st.stop()

all_controls   = [r for r in results]
false_controls = [r for r in results if r['status'] == 'FALSE']
true_controls  = [r for r in results if r['status'] == 'TRUE']

# ─── SUMMARY METRICS ─────────────────────────────────────────────────────────

col1, col2, col3 = st.columns(3)
col1.metric("Total contrôles", len(all_controls))
col2.metric("✅ TRUE (OK)", len(true_controls))
col3.metric("❌ FALSE (Errors)", len(false_controls),
            delta=f"-{len(false_controls)}" if false_controls else "0",
            delta_color="inverse")

st.divider()

# ─── GLOBAL STATUS ───────────────────────────────────────────────────────────

if not false_controls:
    st.success("✅ **VALIDATED** — All contrôle checks passed. The Panel Checker is clean.")
else:
    st.error(f"❌ **{len(false_controls)} error(s) found** — The following contrôle rows returned FALSE:")

# ─── FALSE CONTROLS TABLE ────────────────────────────────────────────────────

if false_controls:
    st.markdown("### ❌ Errors to fix")
    
    rows = []
    for item in false_controls:
        false_details = []
        for fc in item['false_cols']:
            ctx = fc['context']
            if ctx and ctx not in ('None', ''):
                false_details.append(f"{fc['coord']} ({ctx})")
            else:
                false_details.append(fc['coord'])
        
        rows.append({
            "Tab": item['tab'],
            "Row": item['row'],
            "Contrôle label": item['label'],
            "Context": item['context'][:60] if item['context'] else "",
            "FALSE cells": ", ".join(false_details[:6]) + (f" +{len(false_details)-6}" if len(false_details) > 6 else ""),
            "# FALSE": len(item['false_cols'])
        })
    
    df = pd.DataFrame(rows)
    
    def highlight_rows(row):
        return ['background-color: rgba(255, 68, 68, 0.08)'] * len(row)
    
    st.dataframe(
        df.style.apply(highlight_rows, axis=1),
        use_container_width=True,
        hide_index=True,
        height=min(500, 40 + 38 * len(rows))
    )
    
    # ─── MESSAGE DRAFT ───────────────────────────────────────────────────────
    
    st.divider()
    st.markdown("### 📧 Message draft for the team")
    
    msg = generate_message(false_controls, month_label)
    st.text_area("Copy and send this to your team:", value=msg, height=350, key="msg_draft")
    
    st.download_button(
        "⬇️ Download message (.txt)",
        data=msg.encode("utf-8"),
        file_name=f"panel_checker_errors_{datetime.datetime.now().strftime('%Y%m%d')}.txt",
        mime="text/plain"
    )

# ─── ALL CONTROLS (collapsible) ──────────────────────────────────────────────

st.divider()
with st.expander("📋 View all contrôle rows (including TRUE)", expanded=False):
    # Group by tab
    by_tab = {}
    for r in all_controls:
        if r['tab'] not in by_tab:
            by_tab[r['tab']] = []
        by_tab[r['tab']].append(r)
    
    for tab_name, items in by_tab.items():
        n_false = sum(1 for i in items if i['status'] == 'FALSE')
        icon = "❌" if n_false else "✅"
        badge = f"{n_false} error(s)" if n_false else "all OK"
        
        with st.expander(f"{icon} **{tab_name}** — {badge}", expanded=(n_false > 0)):
            for item in items:
                if item['status'] == 'TRUE':
                    st.markdown(f"✅ row {item['row']} — *{item['label']}*")
                else:
                    false_str = ", ".join(fc['coord'] for fc in item['false_cols'])
                    st.markdown(f"❌ **row {item['row']} — {item['label']}** → FALSE: `{false_str}`")
                    if item['context']:
                        st.caption(f"  Context: {item['context']}")

# ─── CSV EXPORT ──────────────────────────────────────────────────────────────

if false_controls:
    export = [{
        'Tab': r['tab'],
        'Row': r['row'],
        'Label': r['label'],
        'Context': r['context'],
        'FALSE_cells': " | ".join(fc['coord'] for fc in r['false_cols']),
        'FALSE_count': len(r['false_cols'])
    } for r in false_controls]
    
    csv = pd.DataFrame(export).to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "⬇️ Download errors (CSV)",
        data=csv,
        file_name=f"panel_checker_errors_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )
