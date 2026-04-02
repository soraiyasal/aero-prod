"""
Production Dashboard — Cap & Aerosol Lines
February 2026

Run with:
    streamlit run production_dashboard.py

Expects the two xlsx files either via the sidebar file uploaders
or at the default paths below.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Production Dashboard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Styling ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: #f8f9fa;
        border-left: 4px solid #0066cc;
        padding: 16px 20px;
        border-radius: 6px;
        margin-bottom: 8px;
    }
    .metric-card.warn { border-left-color: #ff6b35; }
    .metric-card.good { border-left-color: #28a745; }
    .block-title {
        font-size: 13px;
        font-weight: 600;
        color: #555;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .block-value {
        font-size: 28px;
        font-weight: 700;
        color: #1a1a2e;
    }
    .block-sub {
        font-size: 12px;
        color: #888;
        margin-top: 2px;
    }
    div[data-testid="stTabs"] button {
        font-size: 15px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ─── Default file paths (when running with uploaded files in place) ──────────
DEFAULT_CAP  = "отчет цеха производства колпаков февраль 2026 eng.xlsx"
DEFAULT_AERO = "отчет фасовочные линии февраль 2026 eng.xlsx"

# ─── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("🏭 Production Dashboard")
    st.caption("February 2026")
    st.divider()
    st.subheader("📂 Data Files")
    cap_file  = st.file_uploader("Cap Production (.xlsx)",  type="xlsx", key="cap")
    aero_file = st.file_uploader("Aerosol Production (.xlsx)", type="xlsx", key="aero")
    st.divider()
    st.caption("Upload files above, or place them in the same folder as this script.")

# ─── Data loaders ───────────────────────────────────────────────────────────
@st.cache_data
def load_caps(src):
    cap_setup = pd.read_excel(src, sheet_name='Setup', header=None, engine='openpyxl')
    cap_setup.columns = ['type_code', 'type_name']

    frames = []
    for m in ['1', '2', '3', '4', '5']:
        df = pd.read_excel(src, sheet_name=m, engine='openpyxl')
        # Keep only real production rows: col B must be a non-empty string
        df = df[df.iloc[:, 1].apply(lambda x: isinstance(x, str) and len(x) > 0)].copy()
        df['machine'] = f"Machine {m}"
        frames.append(df)

    caps = pd.concat(frames, ignore_index=True)

    # ── Raw columns (manually entered by operators) ──
    RAW = {
        caps.columns[1]:  'type_code',        # cap type abbreviated code
        caps.columns[3]:  'color',             # cap colour
        caps.columns[4]:  'cap_code',          # product cap code
        caps.columns[5]:  'batch_num',         # batch number (often empty)
        caps.columns[6]:  'planned_pcs',       # planned production (pcs)
        caps.columns[7]:  'actual_pcs',        # actual production (pcs)
        caps.columns[8]:  'cycle_time_sec',    # target cycle time (sec)
        caps.columns[9]:  'weight_g',          # target cap weight (g)
        caps.columns[10]: 'cavities',          # number of mould cavities
        caps.columns[12]: 'changeover_min',    # changeover / mould replacement time (min)
        caps.columns[13]: 'daily_maint_min',   # daily mould maintenance (min)
        caps.columns[14]: 'other_stops_min',   # uncontrollable stops (min)
        caps.columns[15]: 'stop_desc',         # free-text stop description
        caps.columns[16]: 'total_time_min',    # total batch time (min)
        caps.columns[17]: 'defective_kg',      # defective material (kg)
        # ── Calculated columns (kept for reference / cross-check) ──
        caps.columns[11]: 'target_ppm',        # = (60/cycle) × cavities
        caps.columns[18]: 'defect_pct',        # = defective_kg / (planned × weight/1000)
        caps.columns[19]: 'OEE1',              # availability × performance
        caps.columns[20]: 'OEE2',              # quality factor
        caps.columns[21]: 'OEE',               # OEE1 × OEE2
    }
    caps.rename(columns=RAW, inplace=True)
    caps['machine'] = caps['machine'].astype(str)

    # Add full name from setup lookup
    lkp = dict(zip(cap_setup['type_code'], cap_setup['type_name']))
    caps['type_name'] = caps['type_code'].map(lkp)

    # Numeric coercion
    for c in ['planned_pcs','actual_pcs','cycle_time_sec','weight_g','cavities',
              'changeover_min','daily_maint_min','other_stops_min',
              'total_time_min','defective_kg','OEE1','OEE2','OEE']:
        caps[c] = pd.to_numeric(caps[c], errors='coerce')

    # Derived columns (recomputed here, not taken from Excel formulas)
    caps['productive_min'] = (
        caps['total_time_min']
        - caps['changeover_min'].fillna(0)
        - caps['daily_maint_min'].fillna(0)
        - caps['other_stops_min'].fillna(0)
    )
    caps['gap_pcs'] = caps['actual_pcs'] - caps['planned_pcs']
    caps['attainment_pct'] = (caps['actual_pcs'] / caps['planned_pcs'] * 100).round(1)

    return caps


@st.cache_data
def load_aero(src):
    frames = []
    for l in ['1', '2', '3', '4']:
        df = pd.read_excel(src, sheet_name=l, engine='openpyxl')
        df = df[df.iloc[:, 0].apply(lambda x: isinstance(x, str) and len(x) > 0)].copy()
        df['line'] = f"Line {l}"
        frames.append(df)

    aero = pd.concat(frames, ignore_index=True)

    # ── Raw columns ──
    RAW = {
        aero.columns[0]:  'line_id',
        aero.columns[1]:  'name_ru',           # product name (Russian)
        aero.columns[2]:  'code',              # product code
        aero.columns[7]:  'batch_num',         # batch number
        aero.columns[8]:  'planned_pcs',       # planned production (pcs)
        aero.columns[9]:  'actual_pcs',        # actual production (pcs)
        aero.columns[10]: 'target_ppm',        # target line speed (pcs/min)
        aero.columns[11]: 'setup_min',         # line setup / changeover (min)
        aero.columns[12]: 'food_stops_min',    # meal-break stops (min)
        aero.columns[13]: 'other_stops_min',   # uncontrollable stops (min)
        aero.columns[14]: 'stop_desc',         # free-text stop description
        aero.columns[15]: 'total_time_min',    # total batch time (min)
        aero.columns[16]: 'initial_defects_pcs',  # incoming defects (pre-line)
        aero.columns[17]: 'mfg_defects_pcs',   # manufacturing defects (on-line)
        aero.columns[18]: 'shortage_pcs',      # can shortage
        aero.columns[19]: 'surplus_pcs',       # can surplus
        aero.columns[20]: 'lab_samples_pcs',   # lab retention samples
        # ── Calculated columns (kept for cross-check) ──
        aero.columns[3]:  'name_en',           # English name (VLOOKUP)
        aero.columns[4]:  'brand',             # brand (VLOOKUP)
        aero.columns[5]:  'nom_group',         # nomenclature group (VLOOKUP)
        aero.columns[6]:  'nom_type',          # nomenclature type (VLOOKUP)
        aero.columns[21]: 'initial_defect_pct',
        aero.columns[22]: 'mfg_defect_pct',
        aero.columns[23]: 'lab_pct',
        aero.columns[24]: 'OEE1',
        aero.columns[25]: 'OEE2',
        aero.columns[26]: 'OEE',
        aero.columns[27]: 'comment',
    }
    aero.rename(columns=RAW, inplace=True)
    aero['line'] = aero['line'].astype(str)

    for c in ['planned_pcs','actual_pcs','target_ppm','setup_min','food_stops_min',
              'other_stops_min','total_time_min','initial_defects_pcs','mfg_defects_pcs',
              'shortage_pcs','surplus_pcs','lab_samples_pcs','OEE1','OEE2','OEE']:
        aero[c] = pd.to_numeric(aero[c], errors='coerce')

    aero['productive_min'] = (
        aero['total_time_min']
        - aero['setup_min'].fillna(0)
        - aero['food_stops_min'].fillna(0)
        - aero['other_stops_min'].fillna(0)
    )
    aero['gap_pcs'] = aero['actual_pcs'] - aero['planned_pcs']
    aero['attainment_pct'] = (aero['actual_pcs'] / aero['planned_pcs'] * 100).round(1)

    return aero


# ─── Load data ──────────────────────────────────────────────────────────────
cap_src  = cap_file  if cap_file  else DEFAULT_CAP
aero_src = aero_file if aero_file else DEFAULT_AERO

try:
    caps = load_caps(cap_src)
    aero = load_aero(aero_src)
    data_loaded = True
except Exception as e:
    data_loaded = False
    st.error(f"Could not load files: {e}")
    st.info("Upload both .xlsx files using the sidebar, or place them in the same directory as this script.")
    st.stop()


# ─── Colour helpers ──────────────────────────────────────────────────────────
def oee_color(v):
    if pd.isna(v):      return "#aaa"
    if v >= 0.85:       return "#28a745"
    if v >= 0.70:       return "#ffc107"
    return "#dc3545"

BLUE_SCALE = px.colors.sequential.Blues[3:]
MACHINE_COLORS = {
    "Machine 1": "#0066cc", "Machine 2": "#3399ff",
    "Machine 3": "#66b2ff", "Machine 4": "#ff6b35",
    "Machine 5": "#004d99",
}
LINE_COLORS = {
    "Line 1": "#7b2d8b", "Line 2": "#b05eb4",
    "Line 3": "#d4a0d8", "Line 4": "#4a0e5a",
}


# ─── KPI card helper ─────────────────────────────────────────────────────────
def kpi(label, value, sub="", css_class=""):
    st.markdown(f"""
    <div class="metric-card {css_class}">
        <div class="block-title">{label}</div>
        <div class="block-value">{value}</div>
        <div class="block-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════════════
tab_overview, tab_caps, tab_aero = st.tabs([
    "📊  Overview", "🔩  Cap Production", "🧴  Aerosol Lines"
])


# ══════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════
with tab_overview:
    st.header("February 2026 — Combined Production Overview")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi("Total Cap Batches", f"{len(caps):,}", "across 5 machines")
    with c2:
        kpi("Total Cap Output", f"{caps['actual_pcs'].sum()/1e6:.2f}M pcs",
            f"of {caps['planned_pcs'].sum()/1e6:.2f}M planned")
    with c3:
        kpi("Total Aerosol Batches", f"{len(aero):,}", "across 4 lines")
    with c4:
        kpi("Total Aerosol Output", f"{aero['actual_pcs'].sum()/1e6:.2f}M pcs",
            f"of {aero['planned_pcs'].sum()/1e6:.2f}M planned")

    st.divider()

    col_a, col_b = st.columns(2)

    # OEE comparison caps
    with col_a:
        st.subheader("🔩 Cap Machine OEE")
        cap_oee = caps.groupby('machine').agg(
            avg_OEE=('OEE','mean'), avg_OEE1=('OEE1','mean'), avg_OEE2=('OEE2','mean')
        ).reset_index().round(3)
        cap_oee['OEE_pct'] = (cap_oee['avg_OEE'] * 100).round(1)

        fig = go.Figure()
        for _, row in cap_oee.iterrows():
            color = MACHINE_COLORS.get(row['machine'], '#0066cc')
            fig.add_trace(go.Bar(
                x=[row['machine']], y=[row['OEE_pct']],
                marker_color=color, name=row['machine'],
                text=[f"{row['OEE_pct']}%"], textposition='outside',
            ))
        fig.add_hline(y=85, line_dash='dash', line_color='#28a745',
                      annotation_text='World-class 85%', annotation_position='top right')
        fig.update_layout(
            yaxis=dict(range=[0, 110], ticksuffix='%', title='OEE'),
            showlegend=False, height=320,
            margin=dict(t=20, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

    # OEE comparison aerosol
    with col_b:
        st.subheader("🧴 Aerosol Line OEE")
        aero_oee = aero.groupby('line').agg(
            avg_OEE=('OEE','mean'), avg_OEE1=('OEE1','mean'), avg_OEE2=('OEE2','mean')
        ).reset_index().round(3)
        aero_oee['OEE_pct'] = (aero_oee['avg_OEE'] * 100).round(1)

        fig2 = go.Figure()
        for _, row in aero_oee.iterrows():
            color = LINE_COLORS.get(row['line'], '#7b2d8b')
            fig2.add_trace(go.Bar(
                x=[row['line']], y=[row['OEE_pct']],
                marker_color=color, name=row['line'],
                text=[f"{row['OEE_pct']}%"], textposition='outside',
            ))
        fig2.add_hline(y=85, line_dash='dash', line_color='#28a745',
                       annotation_text='World-class 85%', annotation_position='top right')
        fig2.update_layout(
            yaxis=dict(range=[0, 110], ticksuffix='%', title='OEE'),
            showlegend=False, height=320,
            margin=dict(t=20, b=20),
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    # OEE component breakdown — both operations side by side
    st.subheader("OEE Component Breakdown (Availability × Speed vs Quality)")
    combined = pd.concat([
        cap_oee.assign(operation='Cap').rename(columns={'machine':'entity'}),
        aero_oee.assign(operation='Aerosol').rename(columns={'line':'entity'}),
    ])

    fig3 = go.Figure()
    fig3.add_trace(go.Bar(name='OEE1 (Time & Speed)',
                          x=combined['entity'], y=(combined['avg_OEE1']*100).round(1),
                          marker_color='#0066cc'))
    fig3.add_trace(go.Bar(name='OEE2 (Quality)',
                          x=combined['entity'], y=(combined['avg_OEE2']*100).round(1),
                          marker_color='#66b2ff'))
    fig3.add_trace(go.Scatter(name='OEE (combined)',
                              x=combined['entity'], y=(combined['avg_OEE']*100).round(1),
                              mode='markers+text', text=(combined['avg_OEE']*100).round(1),
                              textposition='top center',
                              marker=dict(size=12, color='#ff6b35', symbol='diamond')))
    fig3.update_layout(
        barmode='group', yaxis=dict(range=[0,110], ticksuffix='%'),
        height=380, legend=dict(orientation='h', y=-0.15),
        margin=dict(t=20, b=20),
    )
    st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════
# TAB 2 — CAP PRODUCTION
# ══════════════════════════════════════════════════
with tab_caps:
    st.header("🔩 Cap Production — Machine Analysis")

    # Filter
    machines = sorted(caps['machine'].unique())
    sel_machines = st.multiselect("Filter machines", machines, default=machines, key='cap_m')
    df_c = caps[caps['machine'].isin(sel_machines)]

    # KPIs
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: kpi("Batches", f"{len(df_c):,}")
    with k2: kpi("Planned", f"{df_c['planned_pcs'].sum()/1e6:.2f}M pcs")
    with k3: kpi("Actual",  f"{df_c['actual_pcs'].sum()/1e6:.2f}M pcs")
    with k4: kpi("Avg OEE", f"{df_c['OEE'].mean()*100:.1f}%",
                 css_class='good' if df_c['OEE'].mean()>=0.85 else 'warn')
    with k5: kpi("Defective Material",
                 f"{df_c['defective_kg'].sum():.0f} kg",
                 css_class='warn')

    st.divider()
    row1a, row1b = st.columns(2)

    # OEE by machine
    with row1a:
        st.subheader("OEE by Machine")
        c_oee = df_c.groupby('machine').agg(
            OEE=('OEE','mean'), OEE1=('OEE1','mean'), OEE2=('OEE2','mean')
        ).reset_index().sort_values('machine')

        fig = make_subplots(specs=[[{"secondary_y": False}]])
        fig.add_trace(go.Bar(name='OEE1 (Time/Speed)',
                             x=c_oee['machine'], y=(c_oee['OEE1']*100).round(1),
                             marker_color='#0066cc'))
        fig.add_trace(go.Bar(name='OEE2 (Quality)',
                             x=c_oee['machine'], y=(c_oee['OEE2']*100).round(1),
                             marker_color='#66b2ff'))
        fig.add_trace(go.Scatter(name='Combined OEE',
                                 x=c_oee['machine'], y=(c_oee['OEE']*100).round(1),
                                 mode='markers+text',
                                 text=(c_oee['OEE']*100).round(1),
                                 textposition='top center',
                                 marker=dict(size=14, color='#ff6b35', symbol='diamond')))
        fig.add_hline(y=85, line_dash='dash', line_color='green')
        fig.update_layout(barmode='group', yaxis=dict(range=[0,110], ticksuffix='%'),
                          height=320, legend=dict(orientation='h',y=-0.2),
                          margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # Time breakdown
    with row1b:
        st.subheader("Time Allocation (minutes)")
        t_data = df_c.groupby('machine').agg(
            productive=('productive_min','sum'),
            changeover=('changeover_min','sum'),
            maintenance=('daily_maint_min','sum'),
            other_stops=('other_stops_min','sum'),
        ).reset_index()

        fig2 = go.Figure()
        for col, name, color in [
            ('productive', 'Productive Time', '#28a745'),
            ('changeover', 'Changeover', '#ffc107'),
            ('maintenance', 'Maintenance', '#17a2b8'),
            ('other_stops', 'Other Stops', '#dc3545'),
        ]:
            fig2.add_trace(go.Bar(name=name, x=t_data['machine'],
                                  y=t_data[col].fillna(0), marker_color=color))
        fig2.update_layout(barmode='stack', yaxis_title='Minutes',
                           height=320, legend=dict(orientation='h', y=-0.2),
                           margin=dict(t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    row2a, row2b = st.columns(2)

    # Production attainment
    with row2a:
        st.subheader("Planned vs Actual by Machine")
        pv = df_c.groupby('machine')[['planned_pcs','actual_pcs']].sum().reset_index()
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name='Planned', x=pv['machine'],
                              y=pv['planned_pcs'], marker_color='#cce5ff'))
        fig3.add_trace(go.Bar(name='Actual', x=pv['machine'],
                              y=pv['actual_pcs'], marker_color='#0066cc'))
        fig3.update_layout(barmode='overlay', yaxis_title='Pieces',
                           height=320, legend=dict(orientation='h', y=-0.2),
                           margin=dict(t=10, b=10))
        st.plotly_chart(fig3, use_container_width=True)

    # Defective material
    with row2b:
        st.subheader("Defective Material by Cap Type (kg)")
        def_by_type = (df_c.groupby('type_code')['defective_kg']
                       .sum().reset_index()
                       .sort_values('defective_kg', ascending=True)
                       .tail(12))
        fig4 = px.bar(def_by_type, x='defective_kg', y='type_code',
                      orientation='h', color='defective_kg',
                      color_continuous_scale='Reds',
                      labels={'defective_kg':'Defective (kg)', 'type_code':'Cap Type'})
        fig4.update_layout(height=320, margin=dict(t=10, b=10),
                           coloraxis_showscale=False)
        st.plotly_chart(fig4, use_container_width=True)

    # OEE distribution
    st.subheader("OEE Distribution Across All Batches")
    fig5 = px.histogram(df_c.dropna(subset=['OEE']), x='OEE', color='machine',
                        nbins=30, barmode='overlay', opacity=0.7,
                        color_discrete_map={k: v for k, v in MACHINE_COLORS.items()},
                        labels={'OEE': 'OEE Score', 'count': 'Batches'})
    fig5.add_vline(x=0.85, line_dash='dash', line_color='green',
                   annotation_text='85% benchmark')
    fig5.update_layout(height=300, margin=dict(t=10, b=10))
    st.plotly_chart(fig5, use_container_width=True)

    # Batch-level table
    with st.expander("📋 Batch Detail Table"):
        show_cols = ['machine','type_code','type_name','color','cap_code',
                     'planned_pcs','actual_pcs','attainment_pct',
                     'changeover_min','daily_maint_min','other_stops_min',
                     'total_time_min','defective_kg','OEE1','OEE2','OEE','stop_desc']
        st.dataframe(
            df_c[show_cols].rename(columns={
                'machine':'Machine','type_code':'Type','type_name':'Full Name',
                'color':'Colour','cap_code':'Cap Code',
                'planned_pcs':'Planned','actual_pcs':'Actual',
                'attainment_pct':'Attainment %',
                'changeover_min':'Changeover (min)','daily_maint_min':'Maint (min)',
                'other_stops_min':'Other Stops (min)','total_time_min':'Total (min)',
                'defective_kg':'Defective (kg)',
                'OEE1':'OEE1','OEE2':'OEE2','OEE':'OEE',
                'stop_desc':'Stop Description'
            }).style.format({
                'OEE1':'{:.1%}','OEE2':'{:.1%}','OEE':'{:.1%}',
                'Attainment %':'{:.1f}',
            }),
            use_container_width=True, height=400,
        )


# ══════════════════════════════════════════════════
# TAB 3 — AEROSOL LINES
# ══════════════════════════════════════════════════
with tab_aero:
    st.header("🧴 Aerosol Filling Lines — Analysis")

    lines = sorted(aero['line'].unique())
    sel_lines = st.multiselect("Filter lines", lines, default=lines, key='aero_l')
    df_a = aero[aero['line'].isin(sel_lines)]

    # KPIs
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: kpi("Batches", f"{len(df_a):,}")
    with k2: kpi("Planned", f"{df_a['planned_pcs'].sum()/1e6:.2f}M pcs")
    with k3: kpi("Actual",  f"{df_a['actual_pcs'].sum()/1e6:.2f}M pcs")
    with k4: kpi("Avg OEE", f"{df_a['OEE'].mean()*100:.1f}%",
                 css_class='good' if df_a['OEE'].mean()>=0.85 else 'warn')
    with k5: kpi("Mfg Defects",
                 f"{df_a['mfg_defects_pcs'].sum():,.0f} pcs",
                 css_class='warn')

    st.divider()
    r1a, r1b = st.columns(2)

    # OEE by line
    with r1a:
        st.subheader("OEE by Line")
        a_oee = df_a.groupby('line').agg(
            OEE=('OEE','mean'), OEE1=('OEE1','mean'), OEE2=('OEE2','mean')
        ).reset_index()

        fig = go.Figure()
        fig.add_trace(go.Bar(name='OEE1 (Time/Speed)',
                             x=a_oee['line'], y=(a_oee['OEE1']*100).round(1),
                             marker_color='#7b2d8b'))
        fig.add_trace(go.Bar(name='OEE2 (Quality)',
                             x=a_oee['line'], y=(a_oee['OEE2']*100).round(1),
                             marker_color='#d4a0d8'))
        fig.add_trace(go.Scatter(name='Combined OEE',
                                 x=a_oee['line'], y=(a_oee['OEE']*100).round(1),
                                 mode='markers+text',
                                 text=(a_oee['OEE']*100).round(1),
                                 textposition='top center',
                                 marker=dict(size=14, color='#ff6b35', symbol='diamond')))
        fig.add_hline(y=85, line_dash='dash', line_color='green')
        fig.update_layout(barmode='group', yaxis=dict(range=[0,115], ticksuffix='%'),
                          height=320, legend=dict(orientation='h',y=-0.2),
                          margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # Time breakdown
    with r1b:
        st.subheader("Time Allocation (minutes)")
        t_data = df_a.groupby('line').agg(
            productive=('productive_min','sum'),
            setup=('setup_min','sum'),
            food=('food_stops_min','sum'),
            other=('other_stops_min','sum'),
        ).reset_index()

        fig2 = go.Figure()
        for col, name, color in [
            ('productive','Productive Time','#28a745'),
            ('setup','Setup / Changeover','#ffc107'),
            ('food','Food Breaks','#17a2b8'),
            ('other','Other Stops','#dc3545'),
        ]:
            fig2.add_trace(go.Bar(name=name, x=t_data['line'],
                                  y=t_data[col].fillna(0), marker_color=color))
        fig2.update_layout(barmode='stack', yaxis_title='Minutes',
                           height=320, legend=dict(orientation='h',y=-0.2),
                           margin=dict(t=10,b=10))
        st.plotly_chart(fig2, use_container_width=True)

    r2a, r2b = st.columns(2)

    # Brand analysis
    with r2a:
        st.subheader("Output by Brand (Top 12)")
        brand_df = (df_a.groupby('brand')[['planned_pcs','actual_pcs','mfg_defects_pcs']]
                    .sum().reset_index()
                    .sort_values('actual_pcs', ascending=True)
                    .tail(12))
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name='Planned', y=brand_df['brand'],
                              x=brand_df['planned_pcs'],
                              orientation='h', marker_color='#e6d0ee'))
        fig3.add_trace(go.Bar(name='Actual', y=brand_df['brand'],
                              x=brand_df['actual_pcs'],
                              orientation='h', marker_color='#7b2d8b'))
        fig3.update_layout(barmode='overlay', xaxis_title='Pieces',
                           height=360, legend=dict(orientation='h',y=-0.15),
                           margin=dict(t=10,b=10))
        st.plotly_chart(fig3, use_container_width=True)

    # Defect split
    with r2b:
        st.subheader("Defect Type by Line")
        def_data = df_a.groupby('line').agg(
            initial=('initial_defects_pcs','sum'),
            mfg=('mfg_defects_pcs','sum'),
        ).reset_index()

        fig4 = go.Figure()
        fig4.add_trace(go.Bar(name='Initial (incoming) Defects',
                              x=def_data['line'], y=def_data['initial'].fillna(0),
                              marker_color='#ffc107'))
        fig4.add_trace(go.Bar(name='Manufacturing Defects (on-line)',
                              x=def_data['line'], y=def_data['mfg'].fillna(0),
                              marker_color='#dc3545'))
        fig4.update_layout(barmode='group', yaxis_title='Pieces',
                           height=320, legend=dict(orientation='h',y=-0.2),
                           margin=dict(t=10,b=10))
        st.plotly_chart(fig4, use_container_width=True)

    # OEE scatter: speed vs quality
    st.subheader("OEE1 vs OEE2 — Speed/Availability vs Quality (each dot = one batch)")
    scatter_df = df_a.dropna(subset=['OEE1','OEE2','brand'])
    fig5 = px.scatter(scatter_df, x='OEE1', y='OEE2',
                      color='line', size='actual_pcs',
                      size_max=20, opacity=0.6,
                      color_discrete_map={k: v for k,v in LINE_COLORS.items()},
                      hover_data=['name_en','brand','batch_num','actual_pcs'],
                      labels={'OEE1':'OEE1 (Time & Speed)','OEE2':'OEE2 (Quality)'})
    fig5.add_vline(x=0.85, line_dash='dot', line_color='#aaa')
    fig5.add_hline(y=0.85, line_dash='dot', line_color='#aaa')
    fig5.update_layout(height=420, margin=dict(t=10,b=10))
    st.plotly_chart(fig5, use_container_width=True)

    # Batch detail table
    with st.expander("📋 Batch Detail Table"):
        show_cols = ['line','line_id','name_en','brand','batch_num',
                     'planned_pcs','actual_pcs','attainment_pct',
                     'setup_min','food_stops_min','other_stops_min',
                     'total_time_min','initial_defects_pcs','mfg_defects_pcs',
                     'shortage_pcs','surplus_pcs','OEE1','OEE2','OEE',
                     'stop_desc','comment']
        available = [c for c in show_cols if c in df_a.columns]
        st.dataframe(
            df_a[available].rename(columns={
                'line':'Line','line_id':'Line ID','name_en':'Product',
                'brand':'Brand','batch_num':'Batch',
                'planned_pcs':'Planned','actual_pcs':'Actual',
                'attainment_pct':'Attainment %',
                'setup_min':'Setup (min)','food_stops_min':'Food Stops (min)',
                'other_stops_min':'Other Stops (min)','total_time_min':'Total (min)',
                'initial_defects_pcs':'Initial Defects','mfg_defects_pcs':'Mfg Defects',
                'shortage_pcs':'Shortage','surplus_pcs':'Surplus',
                'OEE1':'OEE1','OEE2':'OEE2','OEE':'OEE',
                'stop_desc':'Stop Description','comment':'Comment'
            }).style.format({
                'OEE1':'{:.1%}','OEE2':'{:.1%}','OEE':'{:.1%}',
                'Attainment %':'{:.1f}',
            }),
            use_container_width=True, height=400,
        )
