"""
Production Dashboard — Cap & Aerosol Lines
Multi-month · Auto-insights · Trend analysis

Run with:
    streamlit run production_dashboard.py

HOW TO ADD A NEW MONTH:
    Upload the new month's cap file AND aerosol file using the sidebar uploaders.
    Files are auto-detected from their filenames (Russian month names + year).
    You can upload all months at once — the dashboard combines them automatically.
    If the filename doesn't match the standard pattern, you'll be prompted to
    enter the month label (e.g. "Mar 2026") manually.
"""

import os, re
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(page_title="Production Dashboard", page_icon="🏭",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
  .metric-card{background:#f8f9fa;border-left:4px solid #0066cc;
    padding:14px 18px;border-radius:6px;margin-bottom:8px;}
  .metric-card.warn{border-left-color:#ff6b35;}
  .metric-card.good{border-left-color:#28a745;}
  .metric-card.info{border-left-color:#17a2b8;}
  .block-title{font-size:12px;font-weight:600;color:#555;
    text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px;}
  .block-value{font-size:26px;font-weight:700;color:#1a1a2e;}
  .block-sub{font-size:11px;color:#888;margin-top:2px;}
  .insight-box{padding:12px 16px;border-radius:8px;margin-bottom:8px;
    border-left:4px solid;font-size:14px;line-height:1.5;}
  .insight-box.red{background:#fff5f5;border-color:#dc3545;color:#721c24;}
  .insight-box.amber{background:#fffbf0;border-color:#ffc107;color:#664d03;}
  .insight-box.green{background:#f0fff4;border-color:#28a745;color:#155724;}
  .insight-box.blue{background:#f0f7ff;border-color:#0066cc;color:#004085;}
  .insight-badge{font-size:11px;font-weight:700;text-transform:uppercase;
    letter-spacing:.5px;margin-right:8px;opacity:.7;}
</style>
""", unsafe_allow_html=True)

# ─── Month detection helpers ─────────────────────────────────────────────────
RU_MONTHS = {
    'январь':'Jan','февраль':'Feb','март':'Mar','апрель':'Apr',
    'май':'May','июнь':'Jun','июль':'Jul','август':'Aug',
    'сентябрь':'Sep','октябрь':'Oct','ноябрь':'Nov','декабрь':'Dec',
}
MONTH_ORDER = ['Jan','Feb','Mar','Apr','May','Jun',
               'Jul','Aug','Sep','Oct','Nov','Dec']

def detect_period(filename):
    fn = filename.lower()
    for ru, en in RU_MONTHS.items():
        if ru in fn:
            yr = re.search(r'(202\d|201\d)', fn)
            return f"{en} {yr.group()}" if yr else en
    return None

def sort_periods(periods):
    def key(p):
        parts = p.split()
        mon = MONTH_ORDER.index(parts[0]) if parts[0] in MONTH_ORDER else 99
        yr  = int(parts[1]) if len(parts) > 1 else 0
        return (yr, mon)
    return sorted(set(periods), key=key)

# ─── Data loaders ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_caps(src, period):
    cap_setup = pd.read_excel(src, sheet_name='Setup', header=None, engine='openpyxl')
    cap_setup.columns = ['type_code', 'type_name']
    frames = []
    for m in ['1','2','3','4','5']:
        try:
            df = pd.read_excel(src, sheet_name=m, engine='openpyxl')
        except Exception:
            continue
        df = df[df.iloc[:,1].apply(lambda x: isinstance(x, str) and len(x) > 0)].copy()
        df['machine'] = f"Machine {m}"
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    caps = pd.concat(frames, ignore_index=True)
    RAW = {
        caps.columns[1]:'type_code', caps.columns[3]:'color',
        caps.columns[4]:'cap_code',  caps.columns[5]:'batch_num',
        caps.columns[6]:'planned_pcs', caps.columns[7]:'actual_pcs',
        caps.columns[8]:'cycle_time_sec', caps.columns[9]:'weight_g',
        caps.columns[10]:'cavities',  caps.columns[12]:'changeover_min',
        caps.columns[13]:'daily_maint_min', caps.columns[14]:'other_stops_min',
        caps.columns[15]:'stop_desc', caps.columns[16]:'total_time_min',
        caps.columns[17]:'defective_kg', caps.columns[11]:'target_ppm',
        caps.columns[18]:'defect_pct', caps.columns[19]:'OEE1',
        caps.columns[20]:'OEE2', caps.columns[21]:'OEE',
    }
    caps.rename(columns=RAW, inplace=True)
    lkp = dict(zip(cap_setup['type_code'], cap_setup['type_name']))
    caps['type_name'] = caps['type_code'].map(lkp)
    for c in ['planned_pcs','actual_pcs','cycle_time_sec','weight_g','cavities',
              'changeover_min','daily_maint_min','other_stops_min',
              'total_time_min','defective_kg','OEE1','OEE2','OEE']:
        caps[c] = pd.to_numeric(caps[c], errors='coerce')
    caps['productive_min'] = (caps['total_time_min']
        - caps['changeover_min'].fillna(0)
        - caps['daily_maint_min'].fillna(0)
        - caps['other_stops_min'].fillna(0))
    caps['gap_pcs']         = caps['actual_pcs'] - caps['planned_pcs']
    caps['attainment_pct']  = (caps['actual_pcs'] / caps['planned_pcs'] * 100).round(1)
    caps['units_per_hour']  = (caps['actual_pcs'] / (caps['total_time_min'] / 60)).round(0)
    caps['changeover_pct']  = (caps['changeover_min'].fillna(0) / caps['total_time_min'] * 100).round(1)
    caps['maint_pct']       = (caps['daily_maint_min'].fillna(0) / caps['total_time_min'] * 100).round(1)
    caps['unplanned_pct']   = (caps['other_stops_min'].fillna(0) / caps['total_time_min'] * 100).round(1)
    caps['defect_rate_pct'] = (
        caps['defective_kg'] /
        (caps['planned_pcs'] * caps['weight_g'].fillna(0) / 1000) * 100
    ).round(2)
    caps['first_pass_yield'] = (caps['OEE2'] * 100).round(1)
    caps['period']           = period
    return caps


@st.cache_data(show_spinner=False)
def load_aero(src, period):
    frames = []
    for l in ['1','2','3','4']:
        try:
            df = pd.read_excel(src, sheet_name=l, engine='openpyxl')
        except Exception:
            continue
        df = df[df.iloc[:,0].apply(lambda x: isinstance(x, str) and len(x) > 0)].copy()
        df['line'] = f"Line {l}"
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    aero = pd.concat(frames, ignore_index=True)
    RAW = {
        aero.columns[0]:'line_id',  aero.columns[1]:'name_ru',
        aero.columns[2]:'code',     aero.columns[7]:'batch_num',
        aero.columns[8]:'planned_pcs', aero.columns[9]:'actual_pcs',
        aero.columns[10]:'target_ppm', aero.columns[11]:'setup_min',
        aero.columns[12]:'food_stops_min', aero.columns[13]:'other_stops_min',
        aero.columns[14]:'stop_desc', aero.columns[15]:'total_time_min',
        aero.columns[16]:'initial_defects_pcs', aero.columns[17]:'mfg_defects_pcs',
        aero.columns[18]:'shortage_pcs', aero.columns[19]:'surplus_pcs',
        aero.columns[20]:'lab_samples_pcs',
        aero.columns[3]:'name_en',  aero.columns[4]:'brand',
        aero.columns[5]:'nom_group', aero.columns[6]:'nom_type',
        aero.columns[21]:'initial_defect_pct', aero.columns[22]:'mfg_defect_pct',
        aero.columns[23]:'lab_pct', aero.columns[24]:'OEE1',
        aero.columns[25]:'OEE2',   aero.columns[26]:'OEE',
        aero.columns[27]:'comment',
    }
    aero.rename(columns=RAW, inplace=True)
    for c in ['planned_pcs','actual_pcs','target_ppm','setup_min','food_stops_min',
              'other_stops_min','total_time_min','initial_defects_pcs','mfg_defects_pcs',
              'shortage_pcs','surplus_pcs','lab_samples_pcs','OEE1','OEE2','OEE']:
        aero[c] = pd.to_numeric(aero[c], errors='coerce')
    aero['productive_min'] = (aero['total_time_min']
        - aero['setup_min'].fillna(0)
        - aero['food_stops_min'].fillna(0)
        - aero['other_stops_min'].fillna(0))
    aero['gap_pcs']             = aero['actual_pcs'] - aero['planned_pcs']
    aero['attainment_pct']      = (aero['actual_pcs'] / aero['planned_pcs'] * 100).round(1)
    aero['units_per_hour']      = (aero['actual_pcs'] / (aero['total_time_min'] / 60)).round(0)
    aero['changeover_pct']      = (aero['setup_min'].fillna(0) / aero['total_time_min'] * 100).round(1)
    aero['food_pct']            = (aero['food_stops_min'].fillna(0) / aero['total_time_min'] * 100).round(1)
    aero['unplanned_pct']       = (aero['other_stops_min'].fillna(0) / aero['total_time_min'] * 100).round(1)
    # Incoming defect rate: incoming defects as % of actual output
    aero['incoming_defect_pct'] = (
        aero['initial_defects_pcs'].fillna(0) /
        aero['actual_pcs'].replace(0, np.nan) * 100
    ).round(3)
    # Manufacturing defect rate: on-line defects as % of actual output
    aero['mfg_defect_rate_pct'] = (
        aero['mfg_defects_pcs'].fillna(0) /
        aero['actual_pcs'].replace(0, np.nan) * 100
    ).round(3)
    # First pass yield approximation via OEE2
    aero['first_pass_yield']    = (aero['OEE2'] * 100).round(1)
    # Flag batches where defect field is blank (data quality)
    aero['defect_recorded']     = aero['mfg_defects_pcs'].notna()
    aero['period']              = period
    return aero


# ─── Insights engine ─────────────────────────────────────────────────────────
def generate_insights(caps, aero, multi_month):
    insights = []
    def add(level, cat, title, body, entity=''):
        insights.append(dict(level=level, category=cat,
                             title=title, body=body, entity=entity))

    # OEE — cap machines
    if not caps.empty:
        cm = caps.groupby(['period','machine'])['OEE'].mean().reset_index()
        for _, r in cm.iterrows():
            if r['OEE'] < 0.70:
                add('red','OEE',
                    f"{r['machine']} OEE critical: {r['OEE']*100:.1f}% ({r['period']})",
                    "OEE is more than 15 points below the 85% world-class benchmark. "
                    "This machine lost significant productive capacity.",
                    r['machine'])
            elif r['OEE'] < 0.85:
                add('amber','OEE',
                    f"{r['machine']} OEE below benchmark: {r['OEE']*100:.1f}% ({r['period']})",
                    "Below the 85% industry benchmark. Review changeover and "
                    "maintenance logs for improvement opportunities.",
                    r['machine'])
        best = cm.loc[cm['OEE'].idxmax()]
        if best['OEE'] >= 0.90:
            add('green','OEE',
                f"Best cap machine: {best['machine']} at {best['OEE']*100:.1f}% ({best['period']})",
                "At or above world-class. Identify what's working here — scheduling, "
                "mould condition, operator practice — and apply it to other machines.",
                best['machine'])

    # OEE — aerosol lines
    if not aero.empty:
        lm = aero.groupby(['period','line'])['OEE'].mean().reset_index()
        for _, r in lm.iterrows():
            if r['OEE'] < 0.70:
                add('red','OEE',
                    f"{r['line']} OEE critical: {r['OEE']*100:.1f}% ({r['period']})",
                    "Well below 85% benchmark. Gap is almost entirely OEE1 "
                    "(time & speed) — quality is near-perfect, so focus on "
                    "reducing changeover time first.",
                    r['line'])
            elif r['OEE'] < 0.85:
                add('amber','OEE',
                    f"{r['line']} OEE below benchmark: {r['OEE']*100:.1f}% ({r['period']})",
                    "Below 85%. Main driver is changeover time, not quality.",
                    r['line'])

    # Data quality: OEE1 > 1.0
    if not aero.empty:
        over = aero[aero['OEE1'] > 1.0]
        if len(over) > 0:
            add('amber','Data quality',
                f"{len(over)} aerosol batches with OEE1 > 100% (impossible value)",
                "OEE1 above 1.0 means actual output exceeded the theoretical maximum "
                "for the recorded time. Most likely the target speed (target_ppm) is set "
                "too low for some products, or total_time_min is under-recorded. "
                f"Affected lines: {', '.join(over['line'].unique())}.",
                'Aerosol lines')

    if not caps.empty:
        over_c = caps[caps['OEE1'] > 1.0]
        if len(over_c) > 0:
            add('amber','Data quality',
                f"{len(over_c)} cap batches with OEE1 > 100%",
                "Target cycle time may be set too conservatively for some cap types. "
                f"Affected machines: {', '.join(over_c['machine'].unique())}.",
                'Cap machines')

    # Data quality: blank defect recording
    if not aero.empty:
        blank_pct = aero['mfg_defects_pcs'].isna().sum() / len(aero) * 100
        if blank_pct > 30:
            add('amber','Data quality',
                f"{blank_pct:.0f}% of aerosol batches have no defect count recorded",
                f"{aero['mfg_defects_pcs'].isna().sum()} of {len(aero)} batches are blank. "
                "This makes OEE2 look like 100% even when it isn't. "
                "Operators should record 0 explicitly when there are no defects.",
                'Aerosol lines')

    # Machine breakdown detection
    if not caps.empty:
        for period in caps['period'].unique():
            for machine in caps['machine'].unique():
                sub = caps[(caps['period']==period) & (caps['machine']==machine)]
                if 0 < len(sub) <= 5:
                    add('red','Capacity',
                        f"{machine}: only {len(sub)} batches in {period} — possible breakdown",
                        "Very few batches suggest a mid-period machine failure or "
                        "extended unplanned downtime. Verify against maintenance logs.",
                        machine)

    # Changeover > productive time (data quality / very short runs)
    if not aero.empty:
        odd = aero[aero['changeover_pct'] > 100]
        if len(odd) > 0:
            add('amber','Data quality',
                f"{len(odd)} aerosol batches where changeover exceeds run time",
                f"Changeover % > 100% means the setup took longer than the actual "
                f"production run. These are very short runs — consider whether batches "
                f"this small are worth running at all, or whether they can be "
                f"consolidated with adjacent same-product batches.",
                'Aerosol lines')

    # Changeover — aerosol
    if not aero.empty:
        for period in aero['period'].unique():
            sub = aero[aero['period']==period]
            tot_setup = sub['setup_min'].sum()
            tot_time  = sub['total_time_min'].sum()
            pct = tot_setup / tot_time * 100 if tot_time > 0 else 0
            if pct > 15:
                recovery = tot_time / 100
                add('amber','Changeover',
                    f"Changeover uses {pct:.1f}% of aerosol line time in {period}",
                    f"Total setup time: {tot_setup:,.0f} min across {len(sub)} batches. "
                    f"Each 1% reduction recovers ~{recovery:,.0f} min of capacity. "
                    "Prioritise the highest-batch-count line for SMED analysis.",
                    'Aerosol lines')

    # Changeover — caps
    if not caps.empty:
        for period in caps['period'].unique():
            sub = caps[caps['period']==period]
            tot_co   = sub['changeover_min'].sum()
            tot_time = sub['total_time_min'].sum()
            pct = tot_co / tot_time * 100 if tot_time > 0 else 0
            if pct > 3:
                add('blue','Changeover',
                    f"Cap machine changeovers: {pct:.1f}% of total time in {period}",
                    f"Total changeover time: {tot_co:,.0f} min. "
                    "Consider batching same-mould runs consecutively to reduce resets.",
                    'Cap machines')

    # Month-on-month trends
    if multi_month:
        if not caps.empty:
            periods = sort_periods(caps['period'].unique().tolist())
            if len(periods) >= 2:
                p_prev, p_curr = periods[-2], periods[-1]
                prev_m = caps[caps['period']==p_prev].groupby('machine')['OEE'].mean()
                curr_m = caps[caps['period']==p_curr].groupby('machine')['OEE'].mean()
                for machine in curr_m.index:
                    if machine in prev_m.index:
                        delta = curr_m[machine] - prev_m[machine]
                        if delta < -0.03:
                            add('red','Trend',
                                f"{machine} OEE fell {abs(delta)*100:.1f}pp vs {p_prev}",
                                f"Dropped from {prev_m[machine]*100:.1f}% to "
                                f"{curr_m[machine]*100:.1f}%. Investigate scheduling, "
                                "mould condition, or product mix changes.",
                                machine)
                        elif delta > 0.03:
                            add('green','Trend',
                                f"{machine} OEE improved {delta*100:.1f}pp vs {p_prev}",
                                f"Rose from {prev_m[machine]*100:.1f}% to "
                                f"{curr_m[machine]*100:.1f}%. Identify what changed "
                                "and replicate across other machines.",
                                machine)

        if not aero.empty:
            periods = sort_periods(aero['period'].unique().tolist())
            if len(periods) >= 2:
                p_prev, p_curr = periods[-2], periods[-1]
                prev_l = aero[aero['period']==p_prev].groupby('line')['OEE'].mean()
                curr_l = aero[aero['period']==p_curr].groupby('line')['OEE'].mean()
                for line in curr_l.index:
                    if line in prev_l.index:
                        delta = curr_l[line] - prev_l[line]
                        if delta < -0.03:
                            add('red','Trend',
                                f"{line} OEE fell {abs(delta)*100:.1f}pp vs {p_prev}",
                                f"Dropped from {prev_l[line]*100:.1f}% to "
                                f"{curr_l[line]*100:.1f}%.", line)
                        elif delta > 0.03:
                            add('green','Trend',
                                f"{line} OEE improved {delta*100:.1f}pp vs {p_prev}",
                                f"Rose from {prev_l[line]*100:.1f}% to "
                                f"{curr_l[line]*100:.1f}%.", line)

    return insights


# ─── UI helpers ──────────────────────────────────────────────────────────────
def kpi(label, value, sub='', css_class=''):
    st.markdown(
        f'<div class="metric-card {css_class}">'
        f'<div class="block-title">{label}</div>'
        f'<div class="block-value">{value}</div>'
        f'<div class="block-sub">{sub}</div></div>',
        unsafe_allow_html=True)

def insight_card(ins):
    badge = {'red':'⚠ Alert','amber':'△ Warning',
             'green':'✓ Good','blue':'ℹ Info'}.get(ins['level'],'')
    st.markdown(
        f'<div class="insight-box {ins["level"]}">'
        f'<span class="insight-badge">{badge} · {ins["category"]}</span>'
        f'<strong>{ins["title"]}</strong><br>'
        f'<span style="opacity:.85">{ins["body"]}</span></div>',
        unsafe_allow_html=True)

MACHINE_COLORS = {
    'Machine 1':'#0066cc','Machine 2':'#3399ff',
    'Machine 3':'#66b2ff','Machine 4':'#ff6b35','Machine 5':'#004d99',
}
LINE_COLORS = {
    'Line 1':'#7b2d8b','Line 2':'#b05eb4',
    'Line 3':'#d4a0d8','Line 4':'#4a0e5a',
}

# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("🏭 Production Dashboard")
    st.divider()
    st.subheader("📂 Add data")
    st.caption(
        "February 2026 loads automatically as the baseline. "
        "Upload additional months here to add them — "
        "the Trends tab unlocks once two or more months are loaded."
    )
    cap_uploads  = st.file_uploader("Cap files (.xlsx)", type="xlsx",
                                    accept_multiple_files=True, key="cap_up")
    aero_uploads = st.file_uploader("Aerosol files (.xlsx)", type="xlsx",
                                    accept_multiple_files=True, key="aero_up")
    st.divider()
    st.caption("Or place files in the same folder as this script — "
               "they'll load automatically.")

# ─── Build file → period maps ─────────────────────────────────────────────────
_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CAP  = os.path.join(_SCRIPT_DIR, "отчет цеха производства колпаков февраль 2026 eng")
DEFAULT_AERO = os.path.join(_SCRIPT_DIR, "отчет фасовочные линии февраль 2026 eng.xlsx")

def build_map(uploads, default_path):
    fmap = {}
    # Always load the bundled default file first (Feb 2026 baseline).
    # This means the dashboard is never empty — it always has something to show.
    if os.path.exists(default_path):
        default_period = detect_period(default_path) or "Feb 2026"
        fmap[default_period] = default_path
    # Uploaded files are added on top. If an upload covers the same period
    # as a default (e.g. a corrected Feb file), it overwrites the default.
    for f in (uploads or []):
        period = detect_period(f.name)
        if not period:
            period = st.sidebar.text_input(
                f"Month for '{f.name}'",
                placeholder="e.g. Mar 2026", key=f"lbl_{f.name}")
        if period:
            fmap[period] = f
    return fmap

cap_map  = build_map(cap_uploads,  DEFAULT_CAP)
aero_map = build_map(aero_uploads, DEFAULT_AERO)

if not cap_map and not aero_map:
    st.warning(
        "No data files found. Place the two February xlsx files in the same "
        "folder as this script, or upload files via the sidebar."
    )
    st.stop()

# ─── Load all data ────────────────────────────────────────────────────────────
all_caps, all_aero = [], []
for period, src in cap_map.items():
    with st.spinner(f"Loading caps — {period}…"):
        df = load_caps(src, period)
        if not df.empty:
            all_caps.append(df)

for period, src in aero_map.items():
    with st.spinner(f"Loading aerosol — {period}…"):
        df = load_aero(src, period)
        if not df.empty:
            all_aero.append(df)

caps = pd.concat(all_caps, ignore_index=True) if all_caps else pd.DataFrame()
aero = pd.concat(all_aero, ignore_index=True) if all_aero else pd.DataFrame()

all_periods = sort_periods(
    (caps['period'].unique().tolist() if not caps.empty else []) +
    (aero['period'].unique().tolist() if not aero.empty else [])
)
multi_month = len(all_periods) > 1

# ─── Period filter in sidebar ─────────────────────────────────────────────────
with st.sidebar:
    if multi_month:
        st.subheader("📅 Period filter")
        sel_periods = st.multiselect("Show months", all_periods,
                                     default=all_periods, key="period_sel")
        if not sel_periods:
            sel_periods = all_periods
    else:
        sel_periods = all_periods

caps_f = caps[caps['period'].isin(sel_periods)] if not caps.empty else caps
aero_f = aero[aero['period'].isin(sel_periods)] if not aero.empty else aero

# ─── Insights ────────────────────────────────────────────────────────────────
insights   = generate_insights(caps_f, aero_f, multi_month)
n_red      = sum(1 for i in insights if i['level']=='red')
n_amber    = sum(1 for i in insights if i['level']=='amber')
ins_label  = f"⚡ Insights  ({n_red} alerts · {n_amber} warnings)"

# ═══════════════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════════════
tab_names = ["📊 Overview", ins_label, "🔩 Cap machines", "🧴 Aerosol lines"]
if multi_month:
    tab_names.append("📈 Trends")
tabs = st.tabs(tab_names)
t_ov, t_ins, t_cap, t_aero = tabs[0], tabs[1], tabs[2], tabs[3]
t_trend = tabs[4] if multi_month else None


# ══════════════════════════════════════════════════
# OVERVIEW
# ══════════════════════════════════════════════════
with t_ov:
    period_label = " · ".join(sel_periods)
    st.header(f"Production Overview — {period_label}")
    if multi_month:
        st.info(f"Showing {len(sel_periods)} month(s). Use the sidebar to filter.")

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("Cap batches",  f"{len(caps_f):,}" if not caps_f.empty else "—",
                 f"{caps_f['machine'].nunique()} machines" if not caps_f.empty else "")
    with c2: kpi("Cap output",
                 f"{caps_f['actual_pcs'].sum()/1e6:.2f}M pcs" if not caps_f.empty else "—",
                 f"of {caps_f['planned_pcs'].sum()/1e6:.2f}M planned" if not caps_f.empty else "")
    with c3: kpi("Aerosol batches", f"{len(aero_f):,}" if not aero_f.empty else "—",
                 f"{aero_f['line'].nunique()} lines" if not aero_f.empty else "")
    with c4: kpi("Aerosol output",
                 f"{aero_f['actual_pcs'].sum()/1e6:.2f}M pcs" if not aero_f.empty else "—",
                 f"of {aero_f['planned_pcs'].sum()/1e6:.2f}M planned" if not aero_f.empty else "")

    st.divider()
    col_a, col_b = st.columns(2)

    def oee_bar_chart(df, group_col, oee1_col, oee2_col, oee_col,
                      c1, c2, title):
        fig = go.Figure()
        grp = df.groupby(group_col).agg(
            OEE1=(oee1_col,'mean'), OEE2=(oee2_col,'mean'), OEE=(oee_col,'mean')
        ).reset_index()
        fig.add_trace(go.Bar(name='OEE1', x=grp[group_col],
                             y=(grp['OEE1']*100).round(1), marker_color=c1))
        fig.add_trace(go.Bar(name='OEE2', x=grp[group_col],
                             y=(grp['OEE2']*100).round(1), marker_color=c2))
        fig.add_trace(go.Scatter(name='OEE combined', x=grp[group_col],
                                 y=(grp['OEE']*100).round(1),
                                 mode='markers+text',
                                 text=(grp['OEE']*100).round(1),
                                 textposition='top center',
                                 marker=dict(size=12,color='#ff6b35',symbol='diamond')))
        fig.add_hline(y=85, line_dash='dash', line_color='#28a745',
                      annotation_text='85% benchmark')
        fig.update_layout(barmode='group', yaxis=dict(range=[0,115],ticksuffix='%'),
                          height=320, legend=dict(orientation='h',y=-0.2),
                          margin=dict(t=10,b=10))
        return fig

    with col_a:
        st.subheader("🔩 Cap machine OEE")
        if not caps_f.empty:
            st.plotly_chart(
                oee_bar_chart(caps_f,'machine','OEE1','OEE2','OEE',
                              '#0066cc','#66b2ff','Cap OEE'),
                use_container_width=True)

    with col_b:
        st.subheader("🧴 Aerosol line OEE")
        if not aero_f.empty:
            st.plotly_chart(
                oee_bar_chart(aero_f,'line','OEE1','OEE2','OEE',
                              '#7b2d8b','#d4a0d8','Aerosol OEE'),
                use_container_width=True)


# ══════════════════════════════════════════════════
# INSIGHTS
# ══════════════════════════════════════════════════
with t_ins:
    st.header("⚡ Auto-generated Insights")
    st.caption(
        "Generated automatically from the loaded data. "
        "Covers OEE gaps, data quality issues, downtime patterns, "
        "and — when multiple months are loaded — month-on-month trends."
    )

    if not insights:
        st.success("No issues detected. All machines and lines are performing within benchmark.")
    else:
        CATS = ['OEE','Capacity','Changeover','Trend','Data quality']
        LEVELS = {'red':'Alert','amber':'Warning','green':'Good news','blue':'Info'}

        fc, fl = st.columns(2)
        with fc:
            sel_cat = st.multiselect("Category", CATS, default=CATS, key='ic')
        with fl:
            sel_lvl = st.multiselect("Level", list(LEVELS.keys()),
                                     default=list(LEVELS.keys()),
                                     format_func=lambda x: LEVELS[x], key='il')

        filtered = [i for i in insights
                    if i['category'] in sel_cat and i['level'] in sel_lvl]

        if not filtered:
            st.info("No insights match the selected filters.")
        else:
            for cat in CATS:
                cat_ins = [i for i in filtered if i['category']==cat]
                if cat_ins:
                    st.subheader(cat)
                    for ins in cat_ins:
                        insight_card(ins)

        st.divider()
        s1,s2,s3,s4 = st.columns(4)
        with s1: kpi("Alerts",    str(n_red),   css_class='warn')
        with s2: kpi("Warnings",  str(n_amber), css_class='warn')
        with s3: kpi("Good news", str(sum(1 for i in insights if i['level']=='green')),
                     css_class='good')
        with s4: kpi("Info",      str(sum(1 for i in insights if i['level']=='blue')),
                     css_class='info')


# ══════════════════════════════════════════════════
# CAP MACHINES
# ══════════════════════════════════════════════════
with t_cap:
    st.header("🔩 Cap Production — Machine Analysis")
    if caps_f.empty:
        st.warning("No cap data loaded."); st.stop()

    machines = sorted(caps_f['machine'].unique())
    sel_m = st.multiselect("Filter machines", machines, default=machines, key='cap_m')
    df_c  = caps_f[caps_f['machine'].isin(sel_m)]

    k1,k2,k3,k4,k5 = st.columns(5)
    with k1: kpi("Batches",  f"{len(df_c):,}")
    with k2: kpi("Planned",  f"{df_c['planned_pcs'].sum()/1e6:.2f}M pcs")
    with k3: kpi("Actual",   f"{df_c['actual_pcs'].sum()/1e6:.2f}M pcs")
    with k4: kpi("Avg OEE",  f"{df_c['OEE'].mean()*100:.1f}%",
                 css_class='good' if df_c['OEE'].mean()>=0.85 else 'warn')
    with k5: kpi("Defective material",
                 f"{df_c['defective_kg'].sum():.0f} kg", css_class='warn')

    # ── New KPI row ──
    k6,k7,k8,k9,k10 = st.columns(5)
    attain_100 = (df_c['attainment_pct'] >= 99.5).sum() / len(df_c) * 100
    avg_uph    = df_c['units_per_hour'].mean()
    avg_co_pct = df_c['changeover_pct'].mean()
    avg_maint_pct = df_c['maint_pct'].mean()
    avg_defect_rate = df_c['defect_rate_pct'].dropna()
    avg_fpy    = df_c['first_pass_yield'].mean()
    with k6:  kpi("Attainment ≥100%",
                  f"{attain_100:.0f}%",
                  "batches hitting plan",
                  css_class='good' if attain_100>=95 else 'warn')
    with k7:  kpi("Avg units/hour",
                  f"{avg_uph:,.0f}",
                  "across all machines")
    with k8:  kpi("Avg changeover",
                  f"{avg_co_pct:.1f}%",
                  "of total batch time",
                  css_class='warn' if avg_co_pct>5 else 'good')
    with k9:  kpi("Avg defect rate",
                  f"{avg_defect_rate.mean():.2f}%",
                  "defective kg / planned kg",
                  css_class='warn')
    with k10: kpi("First pass yield",
                  f"{avg_fpy:.1f}%",
                  "via OEE2 (approx)",
                  css_class='good' if avg_fpy>=97 else 'warn')
    st.divider()

    r1a,r1b = st.columns(2)
    with r1a:
        st.subheader("OEE by machine")
        fig = px.bar(
            df_c.groupby(['machine','period']).agg(OEE=('OEE','mean')).reset_index(),
            x='machine', y='OEE',
            color='period' if multi_month else 'machine', barmode='group',
            color_discrete_sequence=list(MACHINE_COLORS.values()))
        fig.add_hline(y=0.85,line_dash='dash',line_color='#28a745',
                      annotation_text='85%')
        fig.update_layout(yaxis=dict(range=[0,1.1],tickformat='.0%'),
                          height=300,legend=dict(orientation='h',y=-0.2),
                          margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with r1b:
        st.subheader("Time allocation (minutes)")
        td = df_c.groupby('machine').agg(
            productive=('productive_min','sum'), changeover=('changeover_min','sum'),
            maint=('daily_maint_min','sum'), other=('other_stops_min','sum')
        ).reset_index()
        fig2 = go.Figure()
        for col,nm,clr in [('productive','Productive','#28a745'),
                            ('changeover','Changeover','#ffc107'),
                            ('maint','Maintenance','#17a2b8'),
                            ('other','Other stops','#dc3545')]:
            fig2.add_trace(go.Bar(name=nm, x=td['machine'],
                                  y=td[col].fillna(0), marker_color=clr))
        fig2.update_layout(barmode='stack', yaxis_title='Minutes',
                           height=300, legend=dict(orientation='h',y=-0.2),
                           margin=dict(t=10,b=10))
        st.plotly_chart(fig2, use_container_width=True)

    r2a,r2b = st.columns(2)
    with r2a:
        st.subheader("Planned vs actual")
        pv = df_c.groupby('machine')[['planned_pcs','actual_pcs']].sum().reset_index()
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name='Planned',x=pv['machine'],y=pv['planned_pcs'],
                              marker_color='#cce5ff'))
        fig3.add_trace(go.Bar(name='Actual',x=pv['machine'],y=pv['actual_pcs'],
                              marker_color='#0066cc'))
        fig3.update_layout(barmode='overlay',yaxis_title='Pieces',height=300,
                           legend=dict(orientation='h',y=-0.2),margin=dict(t=10,b=10))
        st.plotly_chart(fig3, use_container_width=True)

    with r2b:
        st.subheader("Defective material by cap type (kg)")
        def_t = (df_c.groupby('type_code')['defective_kg'].sum().reset_index()
                 .sort_values('defective_kg',ascending=True).tail(12))
        fig4 = px.bar(def_t,x='defective_kg',y='type_code',orientation='h',
                      color='defective_kg',color_continuous_scale='Reds')
        fig4.update_layout(height=300,margin=dict(t=10,b=10),coloraxis_showscale=False)
        st.plotly_chart(fig4, use_container_width=True)

    st.subheader("OEE distribution across all batches")
    fig5 = px.histogram(df_c.dropna(subset=['OEE']),x='OEE',color='machine',
                        nbins=30,barmode='overlay',opacity=0.7,
                        color_discrete_map=MACHINE_COLORS)
    fig5.add_vline(x=0.85,line_dash='dash',line_color='green',
                   annotation_text='85%')
    fig5.update_layout(height=280,margin=dict(t=10,b=10))
    st.plotly_chart(fig5, use_container_width=True)

    # ── New KPI section ──
    st.divider()
    st.subheader("📐 Additional KPIs")

    n1,n2 = st.columns(2)

    with n1:
        st.markdown("**Units per hour by machine**")
        uph = df_c.groupby('machine').agg(
            avg_uph=('units_per_hour','mean'),
            min_uph=('units_per_hour','min'),
            max_uph=('units_per_hour','max'),
        ).reset_index()
        fig_uph = go.Figure()
        fig_uph.add_trace(go.Bar(
            name='Avg units/hour', x=uph['machine'], y=uph['avg_uph'].round(0),
            marker_color='#0066cc',
            error_y=dict(type='data', symmetric=False,
                         array=(uph['max_uph']-uph['avg_uph']).round(0),
                         arrayminus=(uph['avg_uph']-uph['min_uph']).round(0),
                         visible=True, color='#888')
        ))
        fig_uph.update_layout(yaxis_title='Units per hour', height=280,
                              margin=dict(t=10,b=10),
                              legend=dict(orientation='h',y=-0.2))
        st.plotly_chart(fig_uph, use_container_width=True)

    with n2:
        st.markdown("**Changeover vs maintenance vs other stops (% of batch time)**")
        stop_pct = df_c.groupby('machine').agg(
            changeover=('changeover_pct','mean'),
            maintenance=('maint_pct','mean'),
            unplanned=('unplanned_pct','mean'),
        ).reset_index()
        fig_sp = go.Figure()
        for col,nm,clr in [('changeover','Changeover','#ffc107'),
                            ('maintenance','Maintenance','#17a2b8'),
                            ('unplanned','Unplanned stops','#dc3545')]:
            fig_sp.add_trace(go.Bar(name=nm, x=stop_pct['machine'],
                                    y=stop_pct[col].fillna(0), marker_color=clr))
        fig_sp.update_layout(barmode='group', yaxis_title='% of batch time',
                             height=280, legend=dict(orientation='h',y=-0.2),
                             margin=dict(t=10,b=10))
        st.plotly_chart(fig_sp, use_container_width=True)

    n3,n4 = st.columns(2)

    with n3:
        st.markdown("**Defect rate % by machine** (defective kg ÷ planned material kg)")
        dr = df_c.groupby('machine').agg(
            defect_rate=('defect_rate_pct','mean')
        ).reset_index().dropna()
        fig_dr = px.bar(dr, x='machine', y='defect_rate',
                        color='defect_rate', color_continuous_scale='Reds',
                        labels={'defect_rate':'Defect rate %'})
        fig_dr.update_layout(height=280, margin=dict(t=10,b=10),
                             coloraxis_showscale=False, yaxis_title='Defect rate %')
        st.plotly_chart(fig_dr, use_container_width=True)

    with n4:
        st.markdown("**First pass yield by machine** (via OEE2 — approximate)")
        fpy = df_c.groupby('machine').agg(
            fpy=('first_pass_yield','mean')
        ).reset_index()
        fig_fpy = go.Figure()
        fig_fpy.add_trace(go.Bar(x=fpy['machine'], y=fpy['fpy'].round(1),
                                 marker_color='#28a745',
                                 text=fpy['fpy'].round(1),
                                 textposition='outside'))
        fig_fpy.add_hline(y=98, line_dash='dash', line_color='#ffc107',
                          annotation_text='98% target')
        fig_fpy.update_layout(yaxis=dict(range=[90,101], title='First pass yield %'),
                              height=280, margin=dict(t=10,b=10), showlegend=False)
        st.plotly_chart(fig_fpy, use_container_width=True)

    # Downtime reason Pareto from stop_desc
    stop_texts = df_c['stop_desc'].dropna()
    if len(stop_texts) > 0:
        st.markdown("**Stop reason log** (free-text descriptions entered by operators)")
        stop_df = stop_texts.value_counts().reset_index()
        stop_df.columns = ['reason','count']
        fig_par = px.bar(stop_df.head(15), x='count', y='reason',
                         orientation='h', color='count',
                         color_continuous_scale='Oranges',
                         labels={'count':'Batches affected','reason':'Reason'})
        fig_par.update_layout(height=max(200, len(stop_df.head(15))*28),
                              margin=dict(t=10,b=10), coloraxis_showscale=False,
                              yaxis=dict(autorange='reversed'))
        st.plotly_chart(fig_par, use_container_width=True)
    else:
        st.info("No stop reason text has been entered for cap batches this period.")

    with st.expander("📋 Batch detail table"):
        avail = [c for c in ['period','machine','type_code','type_name','color',
                              'cap_code','planned_pcs','actual_pcs','attainment_pct',
                              'changeover_min','daily_maint_min','other_stops_min',
                              'total_time_min','defective_kg','OEE1','OEE2','OEE',
                              'stop_desc'] if c in df_c.columns]
        st.dataframe(df_c[avail].style.format({
            'OEE1':'{:.1%}','OEE2':'{:.1%}','OEE':'{:.1%}',
            'attainment_pct':'{:.1f}'}),
            use_container_width=True, height=400)


# ══════════════════════════════════════════════════
# AEROSOL LINES
# ══════════════════════════════════════════════════
with t_aero:
    st.header("🧴 Aerosol Filling Lines — Analysis")
    if aero_f.empty:
        st.warning("No aerosol data loaded."); st.stop()

    lines  = sorted(aero_f['line'].unique())
    sel_l  = st.multiselect("Filter lines", lines, default=lines, key='aero_l')
    df_a   = aero_f[aero_f['line'].isin(sel_l)]

    k1,k2,k3,k4,k5 = st.columns(5)
    with k1: kpi("Batches",  f"{len(df_a):,}")
    with k2: kpi("Planned",  f"{df_a['planned_pcs'].sum()/1e6:.2f}M pcs")
    with k3: kpi("Actual",   f"{df_a['actual_pcs'].sum()/1e6:.2f}M pcs")
    with k4: kpi("Avg OEE",  f"{df_a['OEE'].mean()*100:.1f}%",
                 css_class='good' if df_a['OEE'].mean()>=0.85 else 'warn')
    with k5: kpi("Mfg defects",
                 f"{df_a['mfg_defects_pcs'].sum():,.0f} pcs", css_class='warn')

    # ── New KPI row ──
    k6,k7,k8,k9,k10 = st.columns(5)
    a_attain_100    = (df_a['attainment_pct'] >= 99.5).sum() / len(df_a) * 100
    a_avg_uph       = df_a['units_per_hour'].mean()
    a_avg_co_pct    = df_a['changeover_pct'].mean()
    a_rec_rate      = df_a['defect_recorded'].mean() * 100
    a_incoming_rate = df_a['incoming_defect_pct'].mean()
    with k6:  kpi("Attainment ≥100%",
                  f"{a_attain_100:.0f}%",
                  "batches hitting plan",
                  css_class='good' if a_attain_100>=95 else 'warn')
    with k7:  kpi("Avg units/hour",
                  f"{a_avg_uph:,.0f}",
                  "across all lines")
    with k8:  kpi("Avg changeover",
                  f"{a_avg_co_pct:.1f}%",
                  "of total batch time",
                  css_class='warn' if a_avg_co_pct>10 else 'good')
    with k9:  kpi("Defect recording",
                  f"{a_rec_rate:.0f}%",
                  "batches with defect entered",
                  css_class='good' if a_rec_rate>=90 else 'warn')
    with k10: kpi("Incoming defect rate",
                  f"{a_incoming_rate:.3f}%",
                  "pre-line defects / output",
                  css_class='warn' if a_incoming_rate>0.05 else 'good')
    st.divider()

    r1a,r1b = st.columns(2)
    with r1a:
        st.subheader("OEE by line")
        fig = px.bar(
            df_a.groupby(['line','period']).agg(OEE=('OEE','mean')).reset_index(),
            x='line', y='OEE',
            color='period' if multi_month else 'line', barmode='group',
            color_discrete_sequence=list(LINE_COLORS.values()))
        fig.add_hline(y=0.85,line_dash='dash',line_color='#28a745',
                      annotation_text='85%')
        fig.update_layout(yaxis=dict(range=[0,1.1],tickformat='.0%'),
                          height=300,legend=dict(orientation='h',y=-0.2),
                          margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with r1b:
        st.subheader("Time allocation (minutes)")
        td = df_a.groupby('line').agg(
            productive=('productive_min','sum'), setup=('setup_min','sum'),
            food=('food_stops_min','sum'), other=('other_stops_min','sum')
        ).reset_index()
        fig2 = go.Figure()
        for col,nm,clr in [('productive','Productive','#28a745'),
                            ('setup','Setup / changeover','#ffc107'),
                            ('food','Food breaks','#17a2b8'),
                            ('other','Other stops','#dc3545')]:
            fig2.add_trace(go.Bar(name=nm, x=td['line'],
                                  y=td[col].fillna(0), marker_color=clr))
        fig2.update_layout(barmode='stack',yaxis_title='Minutes',height=300,
                           legend=dict(orientation='h',y=-0.2),margin=dict(t=10,b=10))
        st.plotly_chart(fig2, use_container_width=True)

    r2a,r2b = st.columns(2)
    with r2a:
        st.subheader("Output by brand (top 12)")
        bd = (df_a.groupby('brand')[['planned_pcs','actual_pcs']].sum()
              .reset_index().sort_values('actual_pcs',ascending=True).tail(12))
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name='Planned',y=bd['brand'],x=bd['planned_pcs'],
                              orientation='h',marker_color='#e6d0ee'))
        fig3.add_trace(go.Bar(name='Actual',y=bd['brand'],x=bd['actual_pcs'],
                              orientation='h',marker_color='#7b2d8b'))
        fig3.update_layout(barmode='overlay',xaxis_title='Pieces',height=360,
                           legend=dict(orientation='h',y=-0.15),margin=dict(t=10,b=10))
        st.plotly_chart(fig3, use_container_width=True)

    with r2b:
        st.subheader("Defect type by line")
        dd = df_a.groupby('line').agg(
            initial=('initial_defects_pcs','sum'),
            mfg=('mfg_defects_pcs','sum')).reset_index()
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(name='Initial (incoming)',x=dd['line'],
                              y=dd['initial'].fillna(0),marker_color='#ffc107'))
        fig4.add_trace(go.Bar(name='Manufacturing (on-line)',x=dd['line'],
                              y=dd['mfg'].fillna(0),marker_color='#dc3545'))
        fig4.update_layout(barmode='group',yaxis_title='Pieces',height=300,
                           legend=dict(orientation='h',y=-0.2),margin=dict(t=10,b=10))
        st.plotly_chart(fig4, use_container_width=True)

    st.subheader("OEE1 vs OEE2 — speed/availability vs quality (each dot = one batch)")
    sc = df_a.dropna(subset=['OEE1','OEE2','brand'])
    fig5 = px.scatter(sc, x='OEE1', y='OEE2', color='line',
                      size='actual_pcs', size_max=20, opacity=0.6,
                      color_discrete_map=LINE_COLORS,
                      hover_data=['name_en','brand','batch_num','actual_pcs','period'])
    fig5.add_vline(x=0.85,line_dash='dot',line_color='#aaa')
    fig5.add_hline(y=0.85,line_dash='dot',line_color='#aaa')
    fig5.update_layout(height=420,margin=dict(t=10,b=10))
    st.plotly_chart(fig5, use_container_width=True)

    # ── New KPI section ──
    st.divider()
    st.subheader("📐 Additional KPIs")

    a1,a2 = st.columns(2)

    with a1:
        st.markdown("**Units per hour by line**")
        uph_a = df_a.groupby('line').agg(
            avg_uph=('units_per_hour','mean'),
            min_uph=('units_per_hour','min'),
            max_uph=('units_per_hour','max'),
        ).reset_index()
        fig_uph = go.Figure()
        fig_uph.add_trace(go.Bar(
            name='Avg units/hour', x=uph_a['line'], y=uph_a['avg_uph'].round(0),
            marker_color='#7b2d8b',
            error_y=dict(type='data', symmetric=False,
                         array=(uph_a['max_uph']-uph_a['avg_uph']).round(0),
                         arrayminus=(uph_a['avg_uph']-uph_a['min_uph']).round(0),
                         visible=True, color='#888')
        ))
        fig_uph.update_layout(yaxis_title='Units per hour', height=280,
                              margin=dict(t=10,b=10), showlegend=False)
        st.plotly_chart(fig_uph, use_container_width=True)

    with a2:
        st.markdown("**Changeover, food breaks & unplanned stops (% of batch time)**")
        sp_a = df_a.groupby('line').agg(
            changeover=('changeover_pct','mean'),
            food=('food_pct','mean'),
            unplanned=('unplanned_pct','mean'),
        ).reset_index()
        fig_spa = go.Figure()
        for col,nm,clr in [('changeover','Changeover','#ffc107'),
                            ('food','Food breaks','#17a2b8'),
                            ('unplanned','Unplanned stops','#dc3545')]:
            fig_spa.add_trace(go.Bar(name=nm, x=sp_a['line'],
                                     y=sp_a[col].fillna(0), marker_color=clr))
        fig_spa.update_layout(barmode='group', yaxis_title='% of batch time',
                              height=280, legend=dict(orientation='h',y=-0.2),
                              margin=dict(t=10,b=10))
        st.plotly_chart(fig_spa, use_container_width=True)

    a3,a4 = st.columns(2)

    with a3:
        st.markdown("**Incoming defect rate by line** (pre-line defects ÷ actual output)")
        inc = df_a.groupby('line').agg(
            incoming_pct=('incoming_defect_pct','mean')
        ).reset_index()
        fig_inc = px.bar(inc, x='line', y='incoming_pct',
                         color='incoming_pct', color_continuous_scale='Oranges',
                         labels={'incoming_pct':'Incoming defect rate %'})
        fig_inc.update_layout(height=280, margin=dict(t=10,b=10),
                              coloraxis_showscale=False, yaxis_title='%')
        st.plotly_chart(fig_inc, use_container_width=True)

    with a4:
        st.markdown("**Defect recording completeness by line**")
        st.caption("32% of aerosol batches have no defect count — this inflates OEE2 to ~100%")
        rec = df_a.groupby('line').agg(
            recorded=('defect_recorded','mean')
        ).reset_index()
        rec['recorded_pct'] = (rec['recorded']*100).round(1)
        rec['missing_pct']  = 100 - rec['recorded_pct']
        fig_rec = go.Figure()
        fig_rec.add_trace(go.Bar(name='Recorded', x=rec['line'],
                                 y=rec['recorded_pct'], marker_color='#28a745'))
        fig_rec.add_trace(go.Bar(name='Blank / missing', x=rec['line'],
                                 y=rec['missing_pct'], marker_color='#dc3545'))
        fig_rec.add_hline(y=100, line_dash='dash', line_color='#28a745',
                          annotation_text='100% target')
        fig_rec.update_layout(barmode='stack', yaxis_title='% of batches',
                              height=280, legend=dict(orientation='h',y=-0.2),
                              margin=dict(t=10,b=10))
        st.plotly_chart(fig_rec, use_container_width=True)

    # Attainment distribution
    st.markdown("**Production attainment distribution** (actual ÷ planned per batch)")
    fig_att = px.histogram(
        df_a.dropna(subset=['attainment_pct']),
        x='attainment_pct', color='line', nbins=40,
        barmode='overlay', opacity=0.7,
        color_discrete_map=LINE_COLORS,
        labels={'attainment_pct':'Attainment %'}
    )
    fig_att.add_vline(x=100, line_dash='dash', line_color='#28a745',
                      annotation_text='100%')
    fig_att.update_layout(height=280, margin=dict(t=10,b=10))
    st.plotly_chart(fig_att, use_container_width=True)

    # Stop reason Pareto
    stop_texts_a = df_a['stop_desc'].dropna()
    if len(stop_texts_a) > 0:
        st.markdown("**Stop reason log** (free-text descriptions entered by operators)")
        sd = stop_texts_a.value_counts().reset_index()
        sd.columns = ['reason','count']
        fig_par = px.bar(sd.head(15), x='count', y='reason',
                         orientation='h', color='count',
                         color_continuous_scale='Purples',
                         labels={'count':'Batches affected','reason':'Reason'})
        fig_par.update_layout(height=max(200, len(sd.head(15))*28),
                              margin=dict(t=10,b=10), coloraxis_showscale=False,
                              yaxis=dict(autorange='reversed'))
        st.plotly_chart(fig_par, use_container_width=True)
    else:
        st.info("No stop reason text has been entered for aerosol batches this period.")

    with st.expander("📋 Batch detail table"):
        avail = [c for c in ['period','line','line_id','name_en','brand','batch_num',
                              'planned_pcs','actual_pcs','attainment_pct',
                              'setup_min','food_stops_min','other_stops_min',
                              'total_time_min','initial_defects_pcs','mfg_defects_pcs',
                              'shortage_pcs','surplus_pcs','OEE1','OEE2','OEE',
                              'stop_desc','comment'] if c in df_a.columns]
        st.dataframe(df_a[avail].style.format({
            'OEE1':'{:.1%}','OEE2':'{:.1%}','OEE':'{:.1%}',
            'attainment_pct':'{:.1f}'}),
            use_container_width=True, height=400)


# ══════════════════════════════════════════════════
# TRENDS (multi-month only)
# ══════════════════════════════════════════════════
if multi_month and t_trend is not None:
    with t_trend:
        st.header("📈 Month-on-month Trends")
        st.caption("Each point = monthly average for that machine or line.")

        def period_sort_key(p):
            parts = p.split()
            yr  = int(parts[1]) if len(parts)>1 else 0
            mon = MONTH_ORDER.index(parts[0]) if parts[0] in MONTH_ORDER else 99
            return yr*100 + mon

        t1,t2 = st.columns(2)
        with t1:
            st.subheader("Cap OEE trend")
            if not caps_f.empty:
                tc = (caps_f.groupby(['period','machine'])['OEE'].mean()
                      .reset_index().sort_values('period', key=lambda s: s.map(period_sort_key)))
                fig = px.line(tc, x='period', y='OEE', color='machine',
                              markers=True, color_discrete_map=MACHINE_COLORS)
                fig.add_hline(y=0.85,line_dash='dash',line_color='#28a745',
                              annotation_text='85%')
                fig.update_layout(yaxis=dict(tickformat='.0%'),height=300,
                                  legend=dict(orientation='h',y=-0.2),
                                  margin=dict(t=10,b=10))
                st.plotly_chart(fig, use_container_width=True)

        with t2:
            st.subheader("Aerosol OEE trend")
            if not aero_f.empty:
                ta = (aero_f.groupby(['period','line'])['OEE'].mean()
                      .reset_index().sort_values('period', key=lambda s: s.map(period_sort_key)))
                fig2 = px.line(ta, x='period', y='OEE', color='line',
                               markers=True, color_discrete_map=LINE_COLORS)
                fig2.add_hline(y=0.85,line_dash='dash',line_color='#28a745',
                               annotation_text='85%')
                fig2.update_layout(yaxis=dict(tickformat='.0%'),height=300,
                                   legend=dict(orientation='h',y=-0.2),
                                   margin=dict(t=10,b=10))
                st.plotly_chart(fig2, use_container_width=True)

        t3,t4 = st.columns(2)
        with t3:
            st.subheader("Cap production volume")
            if not caps_f.empty:
                vc = caps_f.groupby('period').agg(
                    planned=('planned_pcs','sum'),actual=('actual_pcs','sum')
                ).reset_index()
                fig3 = go.Figure()
                fig3.add_trace(go.Bar(name='Planned',x=vc['period'],
                                      y=vc['planned']/1e6,marker_color='#cce5ff'))
                fig3.add_trace(go.Bar(name='Actual',x=vc['period'],
                                      y=vc['actual']/1e6,marker_color='#0066cc'))
                fig3.update_layout(barmode='overlay',yaxis_title='Million pcs',
                                   height=280,legend=dict(orientation='h',y=-0.2),
                                   margin=dict(t=10,b=10))
                st.plotly_chart(fig3, use_container_width=True)

        with t4:
            st.subheader("Aerosol production volume")
            if not aero_f.empty:
                va = aero_f.groupby('period').agg(
                    planned=('planned_pcs','sum'),actual=('actual_pcs','sum')
                ).reset_index()
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(name='Planned',x=va['period'],
                                      y=va['planned']/1e6,marker_color='#e6d0ee'))
                fig4.add_trace(go.Bar(name='Actual',x=va['period'],
                                      y=va['actual']/1e6,marker_color='#7b2d8b'))
                fig4.update_layout(barmode='overlay',yaxis_title='Million pcs',
                                   height=280,legend=dict(orientation='h',y=-0.2),
                                   margin=dict(t=10,b=10))
                st.plotly_chart(fig4, use_container_width=True)

        st.subheader("Changeover % of total time — aerosol lines")
        if not aero_f.empty:
            def co_pct(x):
                t = x['total_time_min'].sum()
                return x['setup_min'].sum() / t * 100 if t > 0 else 0
            co = (aero_f.groupby(['period','line'])
                  .apply(co_pct).reset_index(name='changeover_pct')
                  .sort_values('period', key=lambda s: s.map(period_sort_key)))
            fig5 = px.line(co, x='period', y='changeover_pct', color='line',
                           markers=True, color_discrete_map=LINE_COLORS,
                           labels={'changeover_pct':'% of total time'})
            fig5.add_hline(y=15,line_dash='dash',line_color='#ffc107',
                           annotation_text='15% watch level')
            fig5.update_layout(height=300,legend=dict(orientation='h',y=-0.2),
                               margin=dict(t=10,b=10))
            st.plotly_chart(fig5, use_container_width=True)

        t5,t6 = st.columns(2)
        with t5:
            st.subheader("Cap units per hour trend")
            if not caps_f.empty:
                uph_t = (caps_f.groupby(['period','machine'])['units_per_hour']
                         .mean().reset_index()
                         .sort_values('period', key=lambda s: s.map(period_sort_key)))
                fig6 = px.line(uph_t, x='period', y='units_per_hour', color='machine',
                               markers=True, color_discrete_map=MACHINE_COLORS,
                               labels={'units_per_hour':'Avg units/hour'})
                fig6.update_layout(height=280, legend=dict(orientation='h',y=-0.2),
                                   margin=dict(t=10,b=10))
                st.plotly_chart(fig6, use_container_width=True)

        with t6:
            st.subheader("Aerosol units per hour trend")
            if not aero_f.empty:
                uph_ta = (aero_f.groupby(['period','line'])['units_per_hour']
                          .mean().reset_index()
                          .sort_values('period', key=lambda s: s.map(period_sort_key)))
                fig7 = px.line(uph_ta, x='period', y='units_per_hour', color='line',
                               markers=True, color_discrete_map=LINE_COLORS,
                               labels={'units_per_hour':'Avg units/hour'})
                fig7.update_layout(height=280, legend=dict(orientation='h',y=-0.2),
                                   margin=dict(t=10,b=10))
                st.plotly_chart(fig7, use_container_width=True)

        t7,t8 = st.columns(2)
        with t7:
            st.subheader("Cap defect rate trend (%)")
            if not caps_f.empty:
                dr_t = (caps_f.groupby(['period','machine'])['defect_rate_pct']
                        .mean().reset_index()
                        .sort_values('period', key=lambda s: s.map(period_sort_key)))
                fig8 = px.line(dr_t, x='period', y='defect_rate_pct', color='machine',
                               markers=True, color_discrete_map=MACHINE_COLORS,
                               labels={'defect_rate_pct':'Defect rate %'})
                fig8.update_layout(height=280, legend=dict(orientation='h',y=-0.2),
                                   margin=dict(t=10,b=10))
                st.plotly_chart(fig8, use_container_width=True)

        with t8:
            st.subheader("Aerosol defect recording completeness (%)")
            if not aero_f.empty:
                rc_t = (aero_f.groupby(['period','line'])['defect_recorded']
                        .mean().mul(100).reset_index()
                        .sort_values('period', key=lambda s: s.map(period_sort_key)))
                fig9 = px.line(rc_t, x='period', y='defect_recorded', color='line',
                               markers=True, color_discrete_map=LINE_COLORS,
                               labels={'defect_recorded':'Recording rate %'})
                fig9.add_hline(y=100, line_dash='dash', line_color='#28a745',
                               annotation_text='100% target')
                fig9.update_layout(yaxis=dict(range=[0,105]),
                                   height=280, legend=dict(orientation='h',y=-0.2),
                                   margin=dict(t=10,b=10))
                st.plotly_chart(fig9, use_container_width=True)