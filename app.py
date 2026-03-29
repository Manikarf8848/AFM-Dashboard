import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import report_builder
import history_db

st.set_page_config(page_title="LCY3 AFM Dashboard", layout="wide", page_icon="📊")

if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

DM = st.session_state.dark_mode

_bg      = "#0f1117" if DM else "#ffffff"
_bg2     = "#1a1d27" if DM else "#f0f4ff"
_bg3     = "#22263a" if DM else "#e8eaf6"
_card    = "#1e2235" if DM else "#ffffff"
_text    = "#e8eaf6" if DM else "#1a237e"
_sub     = "#8892b0" if DM else "#666666"
_border  = "#3949ab"
_accent  = "#7986cb" if DM else "#3949ab"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

*, .stApp, .block-container, [data-testid="stAppViewContainer"] {{
    font-family: 'Inter', sans-serif !important;
}}

.stApp {{ background: {_bg} !important; }}
.block-container {{ padding-top: 0.6rem; background: {_bg} !important; }}
[data-testid="stSidebar"] {{ background: {_bg2} !important; }}
[data-testid="stSidebar"] * {{ color: {_text} !important; }}

.dash-header {{
    background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #3949ab 100%);
    padding: 1.1rem 2rem; border-radius: 14px; margin-bottom: 1rem;
    color: white; display: flex; align-items: center; gap: 1.5rem;
    box-shadow: 0 4px 24px rgba(57,73,171,0.35);
    animation: slideDown 0.5s ease-out;
}}
@keyframes slideDown {{
    from {{ transform: translateY(-18px); opacity: 0; }}
    to   {{ transform: translateY(0);     opacity: 1; }}
}}
.dash-header h1 {{ margin: 0; font-size: 1.75rem; font-weight: 800; }}
.dash-header p  {{ margin: 0.2rem 0 0 0; opacity: 0.85; font-size: 0.82rem; }}

.kpi-box {{
    background: {_card}; border-radius: 12px; padding: 1rem 1.1rem;
    box-shadow: 0 2px 16px rgba(0,0,0,{"0.35" if DM else "0.08"});
    text-align: center; border-top: 4px solid {_border};
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    animation: fadeUp 0.4s ease-out both;
}}
.kpi-box:hover {{
    transform: translateY(-5px);
    box-shadow: 0 8px 28px rgba(57,73,171,{"0.45" if DM else "0.22"});
}}
@keyframes fadeUp {{
    from {{ transform: translateY(14px); opacity: 0; }}
    to   {{ transform: translateY(0);    opacity: 1; }}
}}
.kpi-label {{ font-size: 0.72rem; color: {_sub}; font-weight: 700; text-transform: uppercase; letter-spacing: 0.07em; }}
.kpi-value {{ font-size: 1.95rem; font-weight: 800; color: {_text}; line-height: 1.1; }}
.kpi-sub   {{ font-size: 0.72rem; color: {_sub}; margin-top: 0.2rem; }}

.sec-title {{
    font-size: 0.95rem; font-weight: 700; color: {_text};
    padding: 0.45rem 0; border-bottom: 2px solid {_accent};
    margin-bottom: 0.75rem;
}}

.profile-card {{
    background: {_card}; border-radius: 14px; padding: 1.2rem 1.5rem;
    border-left: 5px solid {_accent};
    box-shadow: 0 2px 16px rgba(0,0,0,{"0.35" if DM else "0.08"});
    animation: fadeUp 0.35s ease-out both;
}}
.profile-name {{ font-size: 1.5rem; font-weight: 800; color: {_text}; }}
.profile-sub  {{ font-size: 0.82rem; color: {_sub}; }}

.badge {{
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 0.72rem; font-weight: 700; margin: 2px 3px;
}}
.badge-gold   {{ background: #f59e0b22; color: #f59e0b; border: 1px solid #f59e0b; }}
.badge-red    {{ background: #ef535022; color: #ef5350; border: 1px solid #ef5350; }}
.badge-green  {{ background: #4caf5022; color: #4caf50; border: 1px solid #4caf50; }}
.badge-blue   {{ background: {_accent}22; color: {_accent}; border: 1px solid {_accent}; }}

div[data-testid="stTabs"] button {{
    font-weight: 600; font-size: 0.85rem;
    color: {_sub} !important;
    transition: color 0.2s;
}}
div[data-testid="stTabs"] button[aria-selected="true"] {{
    color: {_accent} !important;
    border-bottom-color: {_accent} !important;
}}

[data-testid="stDataFrame"] {{ border-radius: 10px; overflow: hidden; }}

[data-testid="stTextInput"] input, [data-testid="stSelectbox"] select {{
    background: {_bg2} !important; color: {_text} !important;
    border-color: {_border} !important;
}}

p, label, span, div {{ color: {_text}; }}

.rc-banner {{
    background: linear-gradient(135deg, {"#1e2235" if DM else "#fff3e0"} 0%,
                                        {"#22263a" if DM else "#fff8e1"} 100%);
    border-left: 5px solid #f59e0b; border-radius: 10px;
    padding: 1rem 1.4rem; margin-bottom: 1rem;
    animation: fadeUp 0.4s ease-out both;
}}
.rc-issue {{ font-size: 1.05rem; font-weight: 700; color: #f59e0b; }}
.rc-sub   {{ font-size: 0.82rem; color: {_sub}; margin-top: 4px; }}
</style>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"<div style='font-size:1.1rem;font-weight:800;color:{_text};'>⚙️ Settings</div>", unsafe_allow_html=True)
    dm_toggle = st.toggle("🌙 Dark Mode", value=st.session_state.dark_mode, key="dm_toggle_key")
    if dm_toggle != st.session_state.dark_mode:
        st.session_state.dark_mode = dm_toggle
        st.rerun()

    st.markdown("---")
    st.markdown(f"<div style='font-size:0.95rem;font-weight:700;color:{_text};'>📂 Upload History</div>", unsafe_allow_html=True)
    history_records = history_db.get_history(20)
    if history_records:
        for rec in history_records[:8]:
            weeks_str = ", ".join([f"Wk {w}" for w in rec["week_numbers"]]) or "—"
            st.markdown(f"""
            <div style="background:{_bg3}; border-radius:8px; padding:7px 10px;
                        margin-bottom:6px; border-left:3px solid {_accent}; font-size:0.75rem;">
                <div style="font-weight:700;color:{_text};white-space:nowrap;overflow:hidden;
                            text-overflow:ellipsis;" title="{rec['file_name']}">
                    📄 {rec['file_name'][:26]}{'…' if len(rec['file_name'])>26 else ''}
                </div>
                <div style="color:{_sub};">{rec['upload_ts'][:16]}</div>
                <div style="color:{_accent};">{rec['total_andons']:,} andons · {weeks_str}</div>
                <div style="color:{_sub};">{rec['date_min']} → {rec['date_max']}</div>
            </div>""", unsafe_allow_html=True)
        if st.button("🗑️ Clear History", use_container_width=True):
            history_db.clear_history()
            st.rerun()
    else:
        st.caption("No uploads recorded yet.")

    st.markdown("---")
    st.markdown(f"<div style='font-size:0.7rem;color:{_sub};text-align:center;'>LCY3 AFM Dashboard<br>Made by <b>Manish Karki</b></div>", unsafe_allow_html=True)

st.markdown(f"""
<div class="dash-header">
    <div>
        <h1>📊 LCY3 AFM Dashboard — Floor Health</h1>
        <p>Made by <strong>Manish Karki</strong></p>
    </div>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("Upload one or more Andon data files (CSV or JSON)", type=["json", "csv"], accept_multiple_files=True)

THRESHOLDS = {
    "Amnesty": 10,
    "Drive Lacking Capability": 10,
}
DEFAULT_THRESHOLD = 5
NO_THRESHOLD = ["Unreachable Charger"]


def get_threshold(andon_type):
    if andon_type in NO_THRESHOLD:
        return None
    return THRESHOLDS.get(andon_type, DEFAULT_THRESHOLD)


def threshold_color(val, andon_type):
    t = get_threshold(andon_type)
    if pd.isna(val) or t is None:
        return ""
    if val > t * 1.5:
        return "background-color: rgb(210,40,40); color: white; font-weight: 700"
    elif val > t:
        return "background-color: rgb(255,140,0); color: black; font-weight: 700"
    return "background-color: rgb(60,180,60); color: white; font-weight: 700"


CANONICAL_COLS = {
    "status": "Status",
    "resolver": "Resolver",
    "andon type": "Andon Type",
    "dwell time (hh:mm:ss)": "Dwell Time (hh:mm:ss)",
    "time created": "Time Created",
    "equipment type": "Equipment Type",
    "zone": "Zone",
    "shift": "Shift",
    "blocking": "Blocking",
    "equipment id": "Equipment ID",
    "id": "Equipment ID",
}


@st.cache_data
def load_data(file):
    if file.name.endswith(".json"):
        df = pd.read_json(file)
    else:
        df = pd.read_csv(file)
    df.columns = df.columns.str.strip()
    rename_map = {col: CANONICAL_COLS[col.lower()] for col in df.columns if col.lower() in CANONICAL_COLS}
    df = df.rename(columns=rename_map)
    return df


def dwell_color(val, series):
    if pd.isna(val):
        return ""
    valid = series.dropna()
    if len(valid) < 2:
        return ""
    mn, mx = valid.min(), valid.max()
    if mx == mn:
        return ""
    norm = (val - mn) / (mx - mn)
    if norm >= 0.85:
        r, g, b = 210, 40, 40
    elif norm >= 0.65:
        r, g, b = 255, 120, 0
    elif norm >= 0.45:
        r, g, b = 255, 210, 0
    elif norm >= 0.2:
        r, g, b = 130, 220, 130
    else:
        r, g, b = 50, 160, 50
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    fg = "white" if lum < 155 else "black"
    return f"background-color: rgb({r},{g},{b}); color: {fg}; font-weight: 600"


def build_group_pivot(df, group_col):
    cats = sorted(df[group_col].dropna().unique())
    df = df.copy()
    df["DateLabel"] = df["Time Created"].dt.strftime("%b %d, %Y")
    df["Date"] = df["Time Created"].dt.date

    date_count = df.groupby(["Date", "DateLabel", group_col])["Resolve_Min"].count().reset_index()
    date_avg   = df.groupby(["Date", "DateLabel", group_col])["Resolve_Min"].mean().reset_index()

    all_dates = df.groupby(["Date", "DateLabel"]).apply(lambda x: x).reset_index(drop=True)[["Date", "DateLabel"]].drop_duplicates().sort_values("Date", ascending=False)
    date_labels = all_dates.set_index("Date")["DateLabel"].to_dict()

    cols_dict = {}
    for cat in cats:
        c = date_count[date_count[group_col] == cat].set_index("Date")["Resolve_Min"]
        a = date_avg[date_avg[group_col] == cat].set_index("Date")["Resolve_Min"].round(2)
        cols_dict[(cat, "Andons")] = c
        cols_dict[(cat, "Avg Time")] = a

    total_c = df.groupby("Date")["Resolve_Min"].count()
    total_a = df.groupby("Date")["Resolve_Min"].mean().round(2)
    cols_dict[("Total", "Andons")]   = total_c
    cols_dict[("Total", "Avg Time")] = total_a

    tbl = pd.DataFrame(cols_dict).sort_index(ascending=False)
    tbl.index = [date_labels.get(d, str(d)) for d in tbl.index]
    tbl.index.name = "Rows"
    tbl.columns = pd.MultiIndex.from_tuples(tbl.columns)

    grand = {}
    for cat in cats:
        sub = df[df[group_col] == cat]
        grand[(cat, "Andons")]   = int(sub["Resolve_Min"].count())
        grand[(cat, "Avg Time")] = round(sub["Resolve_Min"].mean(), 2)
    grand[("Total", "Andons")]   = int(df["Resolve_Min"].count())
    grand[("Total", "Avg Time")] = round(df["Resolve_Min"].mean(), 2)

    grand_row = pd.DataFrame(grand, index=["Total"])
    grand_row.columns = pd.MultiIndex.from_tuples(grand_row.columns)
    return pd.concat([grand_row, tbl])


def apply_pivot_style(tbl):
    avg_cols = [c for c in tbl.columns if c[1] == "Avg Time"]

    def _style(data):
        s = pd.DataFrame("", index=data.index, columns=data.columns)
        data_rows = [i for i in data.index if i != "Total"]
        for col in avg_cols:
            if col in data.columns:
                series = data.loc[data_rows, col]
                for idx in data_rows:
                    s.loc[idx, col] = dwell_color(data.loc[idx, col], series)
        if "Total" in data.index:
            s.loc["Total"] = "font-weight: 700; background-color: #e8eaf6; color: #1a237e"
        return s

    fmt_avg    = {c: "{:.2f}" for c in tbl.columns if c[1] == "Avg Time"}
    fmt_counts = {c: "{:,.0f}" for c in tbl.columns if c[1] == "Andons"}
    return tbl.style.apply(_style, axis=None).format(fmt_avg, na_rep="—").format(fmt_counts, na_rep="—")


def donut_chart(df, col, title):
    counts = df.groupby(col)["Resolve_Min"].count().reset_index()
    counts.columns = [col, "Count"]
    total = counts["Count"].sum()
    fig = px.pie(counts, names=col, values="Count", hole=0.55,
                 color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_traces(textinfo="percent+label", pull=[0.03] * len(counts))
    fig.add_annotation(text=f"{total:,}", x=0.5, y=0.5, font_size=22,
                       font_color="#1a237e", showarrow=False, font=dict(weight=700))
    fig.update_layout(showlegend=True, title=title, height=340,
                      legend=dict(orientation="h", yanchor="bottom", y=-0.35),
                      margin=dict(t=40, b=10, l=0, r=0))
    return fig


def hbar_chart(df, col, title):
    avg = (df.groupby(col)["Resolve_Min"].mean()
             .reset_index().sort_values("Resolve_Min")
             .rename(columns={"Resolve_Min": "Avg Time (min)"}))
    fig = px.bar(avg, x="Avg Time (min)", y=col, orientation="h",
                 text=avg["Avg Time (min)"].round(2),
                 color="Avg Time (min)", color_continuous_scale="Blues")
    fig.update_traces(textposition="outside")
    fig.update_layout(coloraxis_showscale=False, title=title, height=340,
                      yaxis_title="", xaxis_title="Avg Dwell Time (min)",
                      margin=dict(t=40, b=10, l=0, r=0))
    return fig


if uploaded_files:
    parts = []
    for uf in uploaded_files:
        part = load_data(uf)
        required_cols = ["Status", "Resolver", "Andon Type", "Dwell Time (hh:mm:ss)", "Time Created"]
        missing = [c for c in required_cols if c not in part.columns]
        if missing:
            st.error(
                f"**{uf.name}** is missing required column(s): {', '.join(f'`{c}`' for c in missing)}\n\n"
                f"**Columns found:** {', '.join(f'`{c}`' for c in part.columns.tolist())}"
            )
            st.stop()
        part["_source_file"] = uf.name
        parts.append(part)
        try:
            history_db.record_upload(uf.name, part)
        except Exception:
            pass

    df = pd.concat(parts, ignore_index=True)

    df = df[df["Status"] == "Resolved"].copy()
    df = df[~df["Andon Type"].isin(["Product Problem", "Out of Work"])]
    df["Resolver"] = df["Resolver"].fillna("System").replace("", "System")
    df["Resolve_Min"] = pd.to_timedelta(df["Dwell Time (hh:mm:ss)"], errors="coerce").dt.total_seconds() / 60
    df = df[df["Resolve_Min"].notna()]
    df["Time Created"] = pd.to_datetime(df["Time Created"], errors="coerce")
    df = df[df["Time Created"].notna()]
    df["Date"]  = df["Time Created"].dt.date
    df["Hour"]  = df["Time Created"].dt.hour
    df["Week"]  = df["Time Created"].dt.isocalendar().week.astype(int)

    optional_cols = {
        "Equipment Type": "Equipment Type" in df.columns,
        "Zone":           "Zone"           in df.columns,
        "Shift":          "Shift"          in df.columns,
        "Blocking":       "Blocking"       in df.columns,
        "Equipment ID":   "Equipment ID"   in df.columns,
    }

    if len(uploaded_files) > 0:
        file_summary_rows = []
        for uf in uploaded_files:
            fpart = df[df["_source_file"] == uf.name]
            if fpart.empty:
                continue
            resolved_count = len(fpart)
            min_dt = fpart["Time Created"].min()
            max_dt = fpart["Time Created"].max()
            file_summary_rows.append((uf.name, resolved_count, min_dt, max_dt))

        if file_summary_rows:
            card_cols = st.columns(len(file_summary_rows))
            for col_obj, (fname, cnt, dt_min, dt_max) in zip(card_cols, file_summary_rows):
                date_label = (dt_min.strftime("%d %b %Y") if dt_min.date() == dt_max.date()
                              else f"{dt_min.strftime('%d %b')} – {dt_max.strftime('%d %b %Y')}")
                time_label = f"{dt_min.strftime('%H:%M')} → {dt_max.strftime('%H:%M')}"
                display_name = fname if len(fname) <= 28 else fname[:25] + "…"
                col_obj.markdown(f"""
                <div style="background:#f0f4ff; border:1px solid #c5cae9; border-left:4px solid #3949ab;
                            border-radius:8px; padding:10px 14px; margin-bottom:8px;">
                    <div style="font-size:0.75rem; color:#5c6bc0; font-weight:700;
                                white-space:nowrap; overflow:hidden; text-overflow:ellipsis;"
                         title="{fname}">📄 {display_name}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#1a237e; line-height:1.2;">{cnt:,}</div>
                    <div style="font-size:0.72rem; color:#555;">andons resolved</div>
                    <div style="font-size:0.72rem; color:#3949ab; margin-top:4px;">📅 {date_label}</div>
                    <div style="font-size:0.72rem; color:#777;">⏰ {time_label}</div>
                </div>""", unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns([2.5, 1.2, 1.2, 1.2])
    with f1:
        min_d, max_d = df["Date"].min(), df["Date"].max()
        date_range = st.date_input("Creation Date", value=(min_d, max_d),
                                   min_value=min_d, max_value=max_d)
    with f2:
        shift_opts = (["All"] + sorted(df["Shift"].dropna().unique().tolist())
                      if optional_cols["Shift"] else ["All"])
        sel_shift = st.selectbox("SHIFT equals", shift_opts)
    with f3:
        resolver_opts = ["All"] + sorted(df["Resolver"].unique().tolist())
        sel_resolver = st.selectbox("Resolver", resolver_opts)
    with f4:
        andon_opts = ["All"] + sorted(df["Andon Type"].dropna().unique().tolist())
        sel_andon = st.selectbox("Andon Type", andon_opts)

    with st.expander("➕ More filters"):
        all_resolvers_list = sorted(df["Resolver"].unique().tolist())
        excluded_resolvers = st.multiselect(
            "Hide resolvers from dashboard",
            options=all_resolvers_list,
            default=[],
            help="Select any logins you want to temporarily hide from all charts and tables."
        )
        search_resolver = st.text_input("🔎 Search resolver by login", "")

    fdf = df.copy()
    if len(date_range) == 2:
        fdf = fdf[(fdf["Date"] >= date_range[0]) & (fdf["Date"] <= date_range[1])]
    if sel_shift != "All" and optional_cols["Shift"]:
        fdf = fdf[fdf["Shift"] == sel_shift]
    if sel_resolver != "All":
        fdf = fdf[fdf["Resolver"] == sel_resolver]
    if sel_andon != "All":
        fdf = fdf[fdf["Andon Type"] == sel_andon]
    if excluded_resolvers:
        fdf = fdf[~fdf["Resolver"].isin(excluded_resolvers)]
    if search_resolver.strip():
        fdf = fdf[fdf["Resolver"].str.contains(search_resolver.strip(), case=False, na=False)]

    st.markdown("<br>", unsafe_allow_html=True)
    k1, k2, k3, k4, k5 = st.columns(5)

    def within_threshold(row):
        t = get_threshold(row["Andon Type"])
        if t is None:
            return True
        return row["Resolve_Min"] <= t

    within_pct = fdf.apply(within_threshold, axis=1).mean() * 100
    efficiency_avg = (fdf.groupby("Resolver").agg(n=("Resolve_Min","count"), avg=("Resolve_Min","mean"))
                     .apply(lambda r: r["n"]/r["avg"] if r["avg"] > 0 else 0, axis=1).mean())

    for col_obj, label, val, sub in [
        (k1, "Total Andons",         f"{len(fdf):,}",                          "Resolved records"),
        (k2, "Avg Resolve Time",     f"{fdf['Resolve_Min'].mean():.2f} min",    "Per andon"),
        (k3, "Median Resolve Time",  f"{fdf['Resolve_Min'].median():.2f} min",  "50th percentile"),
        (k4, "% Within Threshold",   f"{within_pct:.1f}%",                      "vs type targets"),
        (k5, "Avg Efficiency Score", f"{efficiency_avg:.1f}",                   "Andons ÷ Avg Time"),
    ]:
        col_obj.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{val}</div>
            <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    tab_names = ["🏆 Leaderboard", "👤 AFM Profile", "🔍 Root Cause", "AFM Performance", "By Andon Type", "Weekly Breakdown"]
    if optional_cols["Equipment Type"]: tab_names.append("By Equipment Type")
    if optional_cols["Zone"]:           tab_names.append("By Zone")
    if optional_cols["Shift"]:          tab_names.append("By Shift")
    if optional_cols["Blocking"]:       tab_names.append("Blocking Analysis")
    if optional_cols["Equipment ID"]:   tab_names.append("Equipment ID Analysis")
    tab_names += ["Hourly Trend", "Heatmap", "Raw Data", "📤 Export"]

    tabs = st.tabs(tab_names)
    tab = {n: t for n, t in zip(tab_names, tabs)}

    # ── Tab: Leaderboard ──────────────────────────────────────────────────────
    
# ── Tab: Leaderboard ──────────────────────────────────────────────────────
    # ── Tab: Leaderboard ──────────────────────────────────────────────────────
    with tab["🏆 Leaderboard"]:

        all_lb_resolvers  = sorted(fdf["Resolver"].unique().tolist())
        all_lb_andon_types = sorted(fdf["Andon Type"].dropna().unique().tolist())

        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            hidden_resolvers = st.multiselect(
                "🙈 Hide Resolvers",
                options=all_lb_resolvers,
                default=["System"] if "System" in all_lb_resolvers else [],
                key="lb_hide"
            )
        with fc2:
            selected_resolvers = st.multiselect(
                "👁️ Show Only These Resolvers",
                options=all_lb_resolvers,
                default=[],
                key="lb_show"
            )
        with fc3:
            selected_andon_types = st.multiselect(
                "☑️ Filter Andon Types",
                options=all_lb_andon_types,
                default=[],
                key="lb_andon"
            )

        lb_fdf = fdf.copy()
        if hidden_resolvers:
            lb_fdf = lb_fdf[~lb_fdf["Resolver"].isin(hidden_resolvers)]
        if selected_resolvers:
            lb_fdf = lb_fdf[lb_fdf["Resolver"].isin(selected_resolvers)]
        if selected_andon_types:
            lb_fdf = lb_fdf[lb_fdf["Andon Type"].isin(selected_andon_types)]

        if lb_fdf.empty:
            st.warning("No data matches your leaderboard filters. Please adjust your selection.")
            st.stop()

        st.markdown("<br>", unsafe_allow_html=True)

        lb = (lb_fdf.groupby("Resolver")
              .agg(Total_Andons=("Resolve_Min", "count"), Avg_Time=("Resolve_Min", "mean"))
              .reset_index())
        lb["Avg_Time"] = lb["Avg_Time"].round(2)
        lb["Efficiency"] = (lb["Total_Andons"] / lb["Avg_Time"]).round(2)

        within_map = (lb_fdf.groupby("Resolver")
                      .apply(lambda g: g.apply(within_threshold, axis=1).mean() * 100)
                      .round(1).reset_index())
        within_map.columns = ["Resolver", "Within_Threshold"]
        lb = lb.merge(within_map, on="Resolver", how="left")

        lb = lb.sort_values("Avg_Time").reset_index(drop=True)
        lb.index += 1
        lb.index.name = "Rank"

        fastest     = lb.iloc[0]["Resolver"]
        most_active = lb.nlargest(1, "Total_Andons").iloc[0]["Resolver"]
        most_eff    = lb.nlargest(1, "Efficiency").iloc[0]["Resolver"]

        def assign_badge(r):
            badges = []
            if r == fastest:     badges.append("⚡ Fastest")
            if r == most_active: badges.append("🔥 Most Active")
            if r == most_eff:    badges.append("🎯 Most Efficient")
            return "  ".join(badges) if badges else ""

        lb["Badge"] = lb["Resolver"].apply(assign_badge)

        def flag_slow(row):
            if row["Avg_Time"] > DEFAULT_THRESHOLD * 1.5:
                return "🚨 Slow"
            elif row["Avg_Time"] > DEFAULT_THRESHOLD:
                return "⚠️ Above target"
            return "✅ On target"

        lb["Status"] = lb.apply(flag_slow, axis=1)

        b1, b2, b3 = st.columns(3)
        for box, icon, title, name in [
            (b1, "⚡", "Fastest Resolver", fastest),
            (b2, "🔥", "Most Active",      most_active),
            (b3, "🎯", "Most Efficient",   most_eff),
        ]:
            stats = lb[lb["Resolver"] == name].iloc[0]
            box.markdown(f"""
            <div class="kpi-box" style="border-top-color:#f59e0b;">
                <div class="kpi-label">{icon} {title}</div>
                <div class="kpi-value" style="font-size:1.3rem;">{name}</div>
                <div class="kpi-sub">{stats['Total_Andons']:,} andons · {stats['Avg_Time']:.2f} min avg</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="sec-title">Full Rankings</div>', unsafe_allow_html=True)

        def style_lb(data):
            s = pd.DataFrame("", index=data.index, columns=data.columns)
            for idx in data.index:
                status = data.loc[idx, "Status"]
                if "🚨" in str(status):
                    s.loc[idx, "Avg Time (min)"] = "background-color: rgb(210,40,40); color:white; font-weight:700"
                elif "⚠️" in str(status):
                    s.loc[idx, "Avg Time (min)"] = "background-color: rgb(255,140,0); color:black; font-weight:700"
                else:
                    s.loc[idx, "Avg Time (min)"] = "background-color: rgb(60,180,60); color:white; font-weight:700"
                if data.loc[idx, "Badge"]:
                    s.loc[idx, "Badge"] = "background-color: #fef3c7; font-weight:700"
            return s

        display_lb = lb[["Resolver", "Total_Andons", "Avg_Time", "Efficiency", "Within_Threshold", "Badge", "Status"]].copy()
        display_lb.columns = ["Resolver", "Total Andons", "Avg Time (min)", "Efficiency Score", "% Within Threshold", "Badge", "Status"]
        st.dataframe(
            display_lb.style.apply(style_lb, axis=None)
                .format({"Avg Time (min)": "{:.2f}", "Efficiency Score": "{:.2f}", "% Within Threshold": "{:.1f}%"}),
            use_container_width=True, height=450
        )

        st.markdown('<div class="sec-title">Avg Resolve Time vs Target</div>', unsafe_allow_html=True)
        fig_lb = go.Figure()
        fig_lb.add_trace(go.Bar(
            x=lb["Resolver"], y=lb["Avg_Time"],
            marker_color=[
                "rgb(210,40,40)" if "🚨" in s
                else "rgb(255,140,0)" if "⚠️" in s
                else "rgb(60,180,60)"
                for s in lb["Status"]
            ],
            text=lb["Avg_Time"].round(2), textposition="outside", name="Avg Time"
        ))
        fig_lb.add_hline(y=DEFAULT_THRESHOLD, line_dash="dash", line_color="gray",
                         annotation_text=f"Default target ({DEFAULT_THRESHOLD} min)")
        fig_lb.update_layout(height=400, xaxis_title="", yaxis_title="Avg Dwell Time (min)",
                             margin=dict(t=30, b=40, l=0, r=0))
        st.plotly_chart(fig_lb, use_container_width=True)
    # ── Tab: AFM Profile ──────────────────────────────────────────────────────
    with tab["👤 AFM Profile"]:
        st.markdown('<div class="sec-title">AFM Resolver Profile — Drill Down</div>', unsafe_allow_html=True)

        resolver_list = sorted(fdf["Resolver"].unique().tolist())
        sel_profile = st.selectbox("Select Resolver", resolver_list, key="profile_sel")
        pdf = fdf[fdf["Resolver"] == sel_profile]

        p_total   = len(pdf)
        p_avg     = pdf["Resolve_Min"].mean()
        p_med     = pdf["Resolve_Min"].median()
        p_within  = pdf.apply(within_threshold, axis=1).mean() * 100
        p_eff     = p_total / p_avg if p_avg > 0 else 0

        daily_p  = pdf.groupby("Date")["Resolve_Min"].agg(Count="count", Avg="mean").reset_index()
        best_day  = daily_p.nsmallest(1, "Avg").iloc[0] if not daily_p.empty else None
        worst_day = daily_p.nlargest(1, "Avg").iloc[0] if not daily_p.empty else None

        is_fastest     = sel_profile == fdf.groupby("Resolver")["Resolve_Min"].mean().idxmin()
        is_most_active = sel_profile == fdf.groupby("Resolver")["Resolve_Min"].count().idxmax()
        badge_html = ""
        if is_fastest:     badge_html += '<span class="badge badge-gold">⚡ Fastest</span>'
        if is_most_active: badge_html += '<span class="badge badge-blue">🔥 Most Active</span>'
        if p_within >= 80: badge_html += '<span class="badge badge-green">✅ On Target</span>'
        if p_avg > DEFAULT_THRESHOLD * 1.5: badge_html += '<span class="badge badge-red">🚨 Slow</span>'

        st.markdown(f"""
        <div class="profile-card">
            <div class="profile-name">🧑‍💻 {sel_profile}</div>
            <div class="profile-sub">Resolver · {p_total:,} andons handled</div>
            <div style="margin-top:8px;">{badge_html}</div>
        </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        pk1, pk2, pk3, pk4, pk5 = st.columns(5)
        for c_obj, lbl, val, sub in [
            (pk1, "Total Andons",     f"{p_total:,}",         "handled"),
            (pk2, "Avg Time",         f"{p_avg:.2f} min",     "per andon"),
            (pk3, "Median Time",      f"{p_med:.2f} min",     "50th pct"),
            (pk4, "% Within Target",  f"{p_within:.1f}%",     "threshold"),
            (pk5, "Efficiency Score", f"{p_eff:.1f}",         "andons÷avg"),
        ]:
            c_obj.markdown(f"""
            <div class="kpi-box">
                <div class="kpi-label">{lbl}</div>
                <div class="kpi-value">{val}</div>
                <div class="kpi-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        pc1, pc2 = st.columns(2)

        with pc1:
            st.markdown('<div class="sec-title">📈 Daily Andon Count Trend</div>', unsafe_allow_html=True)
            fig_pt = go.Figure()
            fig_pt.add_trace(go.Scatter(
                x=[str(d) for d in daily_p["Date"]], y=daily_p["Count"],
                mode="lines+markers+text",
                text=daily_p["Count"], textposition="top center",
                line=dict(color=_accent, width=2.5),
                marker=dict(size=8, color=_accent),
                fill="tozeroy",
                fillcolor=f"rgba(57,73,171,{'0.2' if not DM else '0.15'})"
            ))
            fig_pt.update_layout(
                height=320, xaxis_title="", yaxis_title="Andons",
                margin=dict(t=20, b=40, l=0, r=0),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(showgrid=False), yaxis=dict(gridcolor="#333" if DM else "#eee")
            )
            st.plotly_chart(fig_pt, use_container_width=True)

        with pc2:
            st.markdown('<div class="sec-title">⏱️ Avg Resolve Time per Day</div>', unsafe_allow_html=True)
            colors_day = ["#ef5350" if v > DEFAULT_THRESHOLD * 1.5
                          else "#ffa726" if v > DEFAULT_THRESHOLD
                          else "#66bb6a" for v in daily_p["Avg"]]
            fig_pa = go.Figure(go.Bar(
                x=[str(d) for d in daily_p["Date"]],
                y=daily_p["Avg"].round(2),
                marker_color=colors_day,
                text=daily_p["Avg"].round(2), textposition="outside"
            ))
            fig_pa.add_hline(y=DEFAULT_THRESHOLD, line_dash="dash", line_color="gray",
                             annotation_text=f"Target ({DEFAULT_THRESHOLD} min)")
            fig_pa.update_layout(
                height=320, xaxis_title="", yaxis_title="Avg Time (min)",
                margin=dict(t=20, b=40, l=0, r=0),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(showgrid=False), yaxis=dict(gridcolor="#333" if DM else "#eee")
            )
            st.plotly_chart(fig_pa, use_container_width=True)

        pc3, pc4 = st.columns(2)
        with pc3:
            st.markdown('<div class="sec-title">🥧 Andon Type Breakdown</div>', unsafe_allow_html=True)
            type_p = pdf.groupby("Andon Type")["Resolve_Min"].count().reset_index()
            type_p.columns = ["Andon Type", "Count"]
            fig_pp = px.pie(type_p, names="Andon Type", values="Count", hole=0.5,
                            color_discrete_sequence=px.colors.qualitative.Set2)
            fig_pp.update_layout(height=280, margin=dict(t=20, b=10, l=0, r=0),
                                 paper_bgcolor="rgba(0,0,0,0)",
                                 legend=dict(font_size=9, orientation="h", y=-0.3))
            st.plotly_chart(fig_pp, use_container_width=True)

        with pc4:
            st.markdown('<div class="sec-title">🏅 Best & Worst Days</div>', unsafe_allow_html=True)
            if best_day is not None:
                st.markdown(f"""
                <div style="background:{'#1b2e1b' if DM else '#e8f5e9'};border-left:4px solid #4caf50;
                            border-radius:8px;padding:10px 14px;margin-bottom:8px;">
                    <div style="font-weight:700;color:#4caf50;">🏆 Best Day</div>
                    <div style="font-size:1.1rem;font-weight:800;color:{_text};">{str(best_day['Date'])}</div>
                    <div style="color:{_sub};">Avg: {best_day['Avg']:.2f} min · {int(best_day['Count'])} andons</div>
                </div>""", unsafe_allow_html=True)
            if worst_day is not None:
                st.markdown(f"""
                <div style="background:{'#2e1b1b' if DM else '#ffebee'};border-left:4px solid #ef5350;
                            border-radius:8px;padding:10px 14px;">
                    <div style="font-weight:700;color:#ef5350;">📉 Worst Day</div>
                    <div style="font-size:1.1rem;font-weight:800;color:{_text};">{str(worst_day['Date'])}</div>
                    <div style="color:{_sub};">Avg: {worst_day['Avg']:.2f} min · {int(worst_day['Count'])} andons</div>
                </div>""", unsafe_allow_html=True)

        st.markdown('<div class="sec-title">Full Record Table</div>', unsafe_allow_html=True)
        show_cols = ["Time Created", "Andon Type", "Resolve_Min", "Status"]
        if optional_cols["Zone"]:          show_cols.append("Zone")
        if optional_cols["Equipment ID"]:  show_cols.append("Equipment ID")
        st.dataframe(pdf[show_cols].sort_values("Time Created", ascending=False).rename(
            columns={"Resolve_Min": "Time (min)"}), use_container_width=True, height=340)

    # ── Tab: Root Cause Analysis ───────────────────────────────────────────────
    with tab["🔍 Root Cause"]:
        st.markdown('<div class="sec-title">Root Cause Analysis — Recurring Issues</div>', unsafe_allow_html=True)

        type_counts = fdf["Andon Type"].value_counts()
        top_issue   = type_counts.index[0]
        top_pct     = type_counts.iloc[0] / len(fdf) * 100

        st.markdown(f"""
        <div class="rc-banner">
            <div class="rc-issue">🚨 Top Recurring Issue: {top_issue} ({top_pct:.1f}%)</div>
            <div class="rc-sub">
                {int(type_counts.iloc[0]):,} of {len(fdf):,} andons · Avg resolve time:
                {fdf[fdf['Andon Type']==top_issue]['Resolve_Min'].mean():.2f} min
            </div>
        </div>""", unsafe_allow_html=True)

        rc1, rc2 = st.columns(2)

        with rc1:
            st.markdown('<div class="sec-title">Top Andon Types by Volume</div>', unsafe_allow_html=True)
            tc_df = type_counts.reset_index()
            tc_df.columns = ["Andon Type", "Count"]
            tc_df["% of Total"] = (tc_df["Count"] / tc_df["Count"].sum() * 100).round(1)
            tc_df["Avg Time (min)"] = tc_df["Andon Type"].map(
                fdf.groupby("Andon Type")["Resolve_Min"].mean().round(2))
            tc_df["Status"] = tc_df["Andon Type"].apply(
                lambda t: "🚨 Above target" if fdf[fdf["Andon Type"]==t]["Resolve_Min"].mean() > get_threshold(t) * 1.5
                else ("⚠️ Borderline" if fdf[fdf["Andon Type"]==t]["Resolve_Min"].mean() > (get_threshold(t) or DEFAULT_THRESHOLD)
                      else "✅ OK") if get_threshold(t) is not None else "—"
            )
            st.dataframe(tc_df.style.format({"Count": "{:,}", "% of Total": "{:.1f}%",
                                              "Avg Time (min)": "{:.2f}"}),
                         use_container_width=True, height=320)

        with rc2:
            st.markdown('<div class="sec-title">Andon Type Distribution</div>', unsafe_allow_html=True)
            fig_rc_pie = px.pie(tc_df.head(12), names="Andon Type", values="Count", hole=0.5,
                                color_discrete_sequence=px.colors.qualitative.Set3)
            fig_rc_pie.update_traces(textinfo="percent+label", textfont_size=9)
            fig_rc_pie.update_layout(height=340, margin=dict(t=20, b=10, l=0, r=0),
                                     paper_bgcolor="rgba(0,0,0,0)", showlegend=False)
            st.plotly_chart(fig_rc_pie, use_container_width=True)

        rc3, rc4 = st.columns(2)

        with rc3:
            st.markdown('<div class="sec-title">⏰ Peak Hours by Andon Type (Top 5 Types)</div>', unsafe_allow_html=True)
            top5_types = type_counts.head(5).index.tolist()
            hour_type = (fdf[fdf["Andon Type"].isin(top5_types)]
                         .groupby(["Hour", "Andon Type"])["Resolve_Min"].count().reset_index()
                         .rename(columns={"Resolve_Min": "Count"}))
            fig_rc_heat = px.bar(hour_type, x="Hour", y="Count", color="Andon Type",
                                 barmode="stack",
                                 color_discrete_sequence=px.colors.qualitative.Set2,
                                 labels={"Hour": "Hour of Day", "Count": "Andon Count"})
            fig_rc_heat.update_layout(height=320, margin=dict(t=20, b=40, l=0, r=0),
                                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                                      legend=dict(font_size=9, orientation="h", y=-0.35))
            st.plotly_chart(fig_rc_heat, use_container_width=True)

        with rc4:
            st.markdown('<div class="sec-title">📅 Week-over-Week Top Issue Trend</div>', unsafe_allow_html=True)
            wow = (fdf[fdf["Andon Type"].isin(top5_types)]
                   .groupby(["Week", "Andon Type"])["Resolve_Min"].count().reset_index()
                   .rename(columns={"Resolve_Min": "Count"}))
            wow["Week"] = "Wk " + wow["Week"].astype(str)
            fig_wow = px.line(wow, x="Week", y="Count", color="Andon Type",
                              markers=True,
                              color_discrete_sequence=px.colors.qualitative.Set2,
                              labels={"Count": "Andon Count", "Week": ""})
            fig_wow.update_layout(height=320, margin=dict(t=20, b=40, l=0, r=0),
                                  paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                                  legend=dict(font_size=9, orientation="h", y=-0.35))
            st.plotly_chart(fig_wow, use_container_width=True)

        if optional_cols["Zone"]:
            st.markdown('<div class="sec-title">📍 Top Issue per Zone</div>', unsafe_allow_html=True)
            zone_top = (fdf.groupby(["Zone", "Andon Type"])["Resolve_Min"].count()
                        .reset_index().rename(columns={"Resolve_Min": "Count"})
                        .sort_values("Count", ascending=False)
                        .groupby("Zone").first().reset_index())
            zone_top.columns = ["Zone", "Top Andon Type", "Count"]
            zone_top["% in Zone"] = zone_top.apply(
                lambda r: f"{r['Count'] / fdf[fdf['Zone']==r['Zone']]['Resolve_Min'].count() * 100:.1f}%", axis=1)
            st.dataframe(zone_top, use_container_width=True)

        st.markdown('<div class="sec-title">🧑‍💻 Slowest Resolvers — Flagged Above Threshold</div>', unsafe_allow_html=True)
        slow_df = (fdf.groupby("Resolver")["Resolve_Min"]
                   .agg(Count="count", Avg="mean").reset_index()
                   .sort_values("Avg", ascending=False))
        slow_df["Status"] = slow_df["Avg"].apply(
            lambda x: "🚨 Slow" if x > DEFAULT_THRESHOLD * 1.5
            else ("⚠️ Above target" if x > DEFAULT_THRESHOLD else "✅ OK"))
        slow_df["Avg"] = slow_df["Avg"].round(2)
        slow_flagged = slow_df[slow_df["Status"] != "✅ OK"]
        if not slow_flagged.empty:
            slow_flagged.columns = ["Resolver", "Total Andons", "Avg Time (min)", "Status"]
            st.dataframe(slow_flagged.style.applymap(
                lambda v: "color: #ef5350; font-weight:700" if "🚨" in str(v)
                else ("color: #ffa726; font-weight:700" if "⚠️" in str(v) else ""),
                subset=["Status"]
            ), use_container_width=True)
        else:
            st.success("✅ All resolvers are within target thresholds.")

    # ── Tab: AFM Performance ──────────────────────────────────────────────────
    with tab["AFM Performance"]:
        st.markdown('<div class="sec-title">Count and Average Dwell Time by Resolver × Andon Type</div>', unsafe_allow_html=True)

        andon_types   = sorted(fdf["Andon Type"].dropna().unique())
        all_resolvers = sorted(fdf["Resolver"].unique())

        cp = fdf.pivot_table(index="Resolver", columns="Andon Type",
                             values="Resolve_Min", aggfunc="count", fill_value=0).reindex(all_resolvers, fill_value=0)
        ap = fdf.pivot_table(index="Resolver", columns="Andon Type",
                             values="Resolve_Min", aggfunc="mean").round(2).reindex(all_resolvers)

        afm_cols = {}
        for cat in andon_types:
            afm_cols[(cat, "Andon Count")]    = cp[cat].astype(int)
            afm_cols[(cat, "Dwell Time Avg")] = ap[cat]
        afm_cols[("Total Andons", "Count")]    = cp[andon_types].sum(axis=1).astype(int)
        afm_cols[("Total Andons", "Avg Time")] = fdf.groupby("Resolver")["Resolve_Min"].mean().reindex(all_resolvers).round(2)

        afm_tbl = pd.DataFrame(afm_cols)
        afm_tbl.columns = pd.MultiIndex.from_tuples(afm_tbl.columns)
        afm_tbl.index.name = "AFM"

        grand = {}
        for cat in andon_types:
            sub = fdf[fdf["Andon Type"] == cat]
            grand[(cat, "Andon Count")]    = int(sub["Resolve_Min"].count())
            grand[(cat, "Dwell Time Avg")] = round(sub["Resolve_Min"].mean(), 2)
        grand[("Total Andons", "Count")]    = int(fdf["Resolve_Min"].count())
        grand[("Total Andons", "Avg Time")] = round(fdf["Resolve_Min"].mean(), 2)

        grand_row = pd.DataFrame(grand, index=["Grand Total"])
        grand_row.columns = pd.MultiIndex.from_tuples(grand_row.columns)
        afm_tbl = pd.concat([afm_tbl, grand_row])

        dwell_cs = [(c, "Dwell Time Avg") for c in andon_types] + [("Total Andons", "Avg Time")]

        def _style_afm(data):
            s = pd.DataFrame("", index=data.index, columns=data.columns)
            rows = [i for i in data.index if i != "Grand Total"]
            for col in dwell_cs:
                if col in data.columns:
                    ser = data.loc[rows, col]
                    for idx in rows:
                        s.loc[idx, col] = dwell_color(data.loc[idx, col], ser)
            s.loc["Grand Total"] = "font-weight:700; background-color:#e8eaf6; color:#1a237e"
            return s

        afm_styler = (afm_tbl.style.apply(_style_afm, axis=None)
            .format({c: "{:.2f}" for c in afm_tbl.columns if c[1] in ("Dwell Time Avg", "Avg Time")}, na_rep="—")
            .format({c: "{:.0f}" for c in afm_tbl.columns if c[1] in ("Andon Count", "Count")}, na_rep="—"))
        st.dataframe(afm_styler, use_container_width=True, height=430)

        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(donut_chart(fdf, "Resolver", "Andons by Resolver"), use_container_width=True)
        with c2:
            st.plotly_chart(hbar_chart(fdf, "Resolver", "Avg Resolve Time by Resolver"), use_container_width=True)

    # ── Tab: By Andon Type ────────────────────────────────────────────────────
    with tab["By Andon Type"]:
        st.markdown('<div class="sec-title">Number of Andons and Dwell Time by Date × Andon Type</div>', unsafe_allow_html=True)
        tbl_at = build_group_pivot(fdf, "Andon Type")
        st.dataframe(apply_pivot_style(tbl_at), use_container_width=True, height=400)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(donut_chart(fdf, "Andon Type", "Andons by Type"), use_container_width=True)
        with c2:
            st.plotly_chart(hbar_chart(fdf, "Andon Type", "Avg Resolve Time by Andon Type"), use_container_width=True)

    # ── Tab: Weekly Breakdown ─────────────────────────────────────────────────
    with tab["Weekly Breakdown"]:
        weeks_avail = sorted(fdf["Week"].dropna().unique(), reverse=True)

        st.markdown('<div class="sec-title">Andons by Type and Week</div>', unsafe_allow_html=True)
        week_count_p = fdf.pivot_table(index="Andon Type", columns="Week",
                                       values="Resolve_Min", aggfunc="count", fill_value=0)
        week_avg_p   = fdf.pivot_table(index="Andon Type", columns="Week",
                                       values="Resolve_Min", aggfunc="mean")

        week_cols = {}
        for w in weeks_avail:
            if w in week_count_p.columns:
                week_cols[(f"Wk {w}", "Andons")]   = week_count_p[w].astype(int)
                week_cols[(f"Wk {w}", "Avg Time")] = week_avg_p[w].round(2) if w in week_avg_p.columns else pd.Series(dtype=float)
        week_cols[("Total", "Andons")]   = week_count_p[weeks_avail].sum(axis=1).astype(int)
        week_cols[("Total", "Avg Time")] = fdf.groupby("Andon Type")["Resolve_Min"].mean().round(2)

        weekly_tbl = pd.DataFrame(week_cols)
        weekly_tbl.columns = pd.MultiIndex.from_tuples(weekly_tbl.columns)
        weekly_tbl.index.name = "Andon Type"
        weekly_tbl = weekly_tbl.sort_values(("Total", "Andons"), ascending=False)

        wk_grand = {}
        for w in weeks_avail:
            sub_w = fdf[fdf["Week"] == w]
            wk_grand[(f"Wk {w}", "Andons")]   = int(sub_w["Resolve_Min"].count())
            wk_grand[(f"Wk {w}", "Avg Time")] = round(sub_w["Resolve_Min"].mean(), 2)
        wk_grand[("Total", "Andons")]   = int(fdf["Resolve_Min"].count())
        wk_grand[("Total", "Avg Time")] = round(fdf["Resolve_Min"].mean(), 2)

        wk_grand_row = pd.DataFrame(wk_grand, index=["Grand Total"])
        wk_grand_row.columns = pd.MultiIndex.from_tuples(wk_grand_row.columns)
        weekly_tbl = pd.concat([wk_grand_row, weekly_tbl])

        wk_avg_cols = [c for c in weekly_tbl.columns if c[1] == "Avg Time"]

        def _style_weekly(data):
            s = pd.DataFrame("", index=data.index, columns=data.columns)
            data_rows = [i for i in data.index if i != "Grand Total"]
            for col in wk_avg_cols:
                if col in data.columns:
                    ser = data.loc[data_rows, col]
                    for idx in data_rows:
                        s.loc[idx, col] = dwell_color(data.loc[idx, col], ser)
            if "Grand Total" in data.index:
                s.loc["Grand Total"] = "font-weight:700; background-color:#e8eaf6; color:#1a237e"
            return s

        wk_styler = (weekly_tbl.style.apply(_style_weekly, axis=None)
            .format({c: "{:.2f}" for c in weekly_tbl.columns if c[1] == "Avg Time"}, na_rep="—")
            .format({c: "{:,.0f}" for c in weekly_tbl.columns if c[1] == "Andons"}, na_rep="—"))
        st.dataframe(wk_styler, use_container_width=True, height=420)

        st.markdown('<div class="sec-title">AFM Performance by Week (Resolver × Week)</div>', unsafe_allow_html=True)

        afm_wk_cnt = fdf.pivot_table(index="Resolver", columns="Week",
                                     values="Resolve_Min", aggfunc="count", fill_value=0)
        afm_wk_avg = fdf.pivot_table(index="Resolver", columns="Week",
                                     values="Resolve_Min", aggfunc="mean")

        afm_wk_cols = {}
        for w in weeks_avail:
            if w in afm_wk_cnt.columns:
                afm_wk_cols[(f"Wk {w}", "Andons")]   = afm_wk_cnt[w].astype(int)
                afm_wk_cols[(f"Wk {w}", "Avg Time")] = afm_wk_avg[w].round(2) if w in afm_wk_avg.columns else pd.Series(dtype=float)
        afm_wk_cols[("Total", "Andons")]   = afm_wk_cnt[weeks_avail].sum(axis=1).astype(int)
        afm_wk_cols[("Total", "Avg Time")] = fdf.groupby("Resolver")["Resolve_Min"].mean().round(2)

        afm_wk_tbl = pd.DataFrame(afm_wk_cols)
        afm_wk_tbl.columns = pd.MultiIndex.from_tuples(afm_wk_tbl.columns)
        afm_wk_tbl.index.name = "Resolver"
        afm_wk_tbl = afm_wk_tbl.sort_values(("Total", "Andons"), ascending=False)

        afm_wk_grand = {}
        for w in weeks_avail:
            sub_w = fdf[fdf["Week"] == w]
            afm_wk_grand[(f"Wk {w}", "Andons")]   = int(sub_w["Resolve_Min"].count())
            afm_wk_grand[(f"Wk {w}", "Avg Time")] = round(sub_w["Resolve_Min"].mean(), 2)
        afm_wk_grand[("Total", "Andons")]   = int(fdf["Resolve_Min"].count())
        afm_wk_grand[("Total", "Avg Time")] = round(fdf["Resolve_Min"].mean(), 2)

        afm_wk_grand_row = pd.DataFrame(afm_wk_grand, index=["Grand Total"])
        afm_wk_grand_row.columns = pd.MultiIndex.from_tuples(afm_wk_grand_row.columns)
        afm_wk_tbl = pd.concat([afm_wk_grand_row, afm_wk_tbl])

        afm_wk_avg_cols = [c for c in afm_wk_tbl.columns if c[1] == "Avg Time"]

        def _style_afm_wk(data):
            s = pd.DataFrame("", index=data.index, columns=data.columns)
            data_rows = [i for i in data.index if i != "Grand Total"]
            for col in afm_wk_avg_cols:
                if col in data.columns:
                    ser = data.loc[data_rows, col]
                    for idx in data_rows:
                        s.loc[idx, col] = dwell_color(data.loc[idx, col], ser)
            if "Grand Total" in data.index:
                s.loc["Grand Total"] = "font-weight:700; background-color:#e8eaf6; color:#1a237e"
            return s

        afm_wk_styler = (afm_wk_tbl.style.apply(_style_afm_wk, axis=None)
            .format({c: "{:.2f}" for c in afm_wk_tbl.columns if c[1] == "Avg Time"}, na_rep="—")
            .format({c: "{:,.0f}" for c in afm_wk_tbl.columns if c[1] == "Andons"}, na_rep="—"))
        st.dataframe(afm_wk_styler, use_container_width=True, height=420)

        st.markdown('<div class="sec-title">Avg Resolve Time per Resolver</div>', unsafe_allow_html=True)
        afm_bar = (fdf.groupby("Resolver")["Resolve_Min"]
                   .agg(Avg="mean", Count="count").reset_index()
                   .sort_values("Avg"))
        afm_bar["Color"] = afm_bar["Avg"].apply(
            lambda x: "rgb(60,180,60)" if x <= DEFAULT_THRESHOLD
            else ("rgb(255,140,0)" if x <= DEFAULT_THRESHOLD * 1.5 else "rgb(210,40,40)"))
        fig_afm_bar = go.Figure(go.Bar(
            x=afm_bar["Resolver"], y=afm_bar["Avg"],
            marker_color=afm_bar["Color"],
            text=afm_bar["Avg"].round(2), textposition="outside",
            customdata=afm_bar["Count"],
            hovertemplate="<b>%{x}</b><br>Avg: %{y:.2f} min<br>Andons: %{customdata}<extra></extra>"
        ))
        fig_afm_bar.add_hline(y=DEFAULT_THRESHOLD, line_dash="dash", line_color="gray",
                               annotation_text=f"Default target ({DEFAULT_THRESHOLD} min)")
        fig_afm_bar.update_layout(height=380, xaxis_title="", yaxis_title="Avg Dwell Time (min)",
                                  margin=dict(t=30, b=40, l=0, r=0))
        st.plotly_chart(fig_afm_bar, use_container_width=True)

        st.markdown('<div class="sec-title">System vs Non-System Andons by Week</div>', unsafe_allow_html=True)
        fdf["_resolver_type"] = fdf["Resolver"].apply(lambda r: "System" if r == "System" else "Non-System")
        sys_wk = (fdf.groupby(["Week", "_resolver_type"])["Resolve_Min"]
                  .agg(Andons="count", Avg_Time="mean").reset_index())
        sys_wk["Week_Label"] = "Wk " + sys_wk["Week"].astype(str)

        col_sys, col_nonsys = st.columns(2)
        with col_sys:
            fig_sys_bar = px.bar(sys_wk, x="Week_Label", y="Andons", color="_resolver_type",
                                 barmode="group",
                                 color_discrete_map={"System": "#ef5350", "Non-System": "#42a5f5"},
                                 title="Total Andons: System vs Non-System per Week",
                                 labels={"_resolver_type": "Resolver Type", "Week_Label": "Week", "Andons": "Count"})
            fig_sys_bar.update_layout(height=360, margin=dict(t=50, b=30, l=0, r=0),
                                      legend=dict(orientation="h", y=-0.25))
            st.plotly_chart(fig_sys_bar, use_container_width=True)

        with col_nonsys:
            fig_sys_avg = px.bar(sys_wk, x="Week_Label", y="Avg_Time", color="_resolver_type",
                                 barmode="group",
                                 color_discrete_map={"System": "#ef5350", "Non-System": "#42a5f5"},
                                 title="Avg Resolve Time: System vs Non-System per Week",
                                 labels={"_resolver_type": "Resolver Type", "Week_Label": "Week", "Avg_Time": "Avg Time (min)"})
            fig_sys_avg.add_hline(y=DEFAULT_THRESHOLD, line_dash="dash", line_color="gray",
                                   annotation_text=f"Target ({DEFAULT_THRESHOLD} min)")
            fig_sys_avg.update_layout(height=360, margin=dict(t=50, b=30, l=0, r=0),
                                       legend=dict(orientation="h", y=-0.25))
            st.plotly_chart(fig_sys_avg, use_container_width=True)

        sys_summary = (fdf.groupby("_resolver_type")["Resolve_Min"]
                       .agg(Total_Andons="count", Avg_Time="mean").reset_index()
                       .rename(columns={"_resolver_type": "Resolver Type",
                                        "Total_Andons": "Total Andons", "Avg_Time": "Avg Time (min)"}))
        sys_summary["Avg Time (min)"] = sys_summary["Avg Time (min)"].round(2)
        sys_summary["% of Total"] = (sys_summary["Total Andons"] / sys_summary["Total Andons"].sum() * 100).round(1).astype(str) + "%"
        st.dataframe(sys_summary.set_index("Resolver Type"), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            top10 = (fdf.groupby("Andon Type")["Resolve_Min"].count()
                     .nlargest(10).reset_index().rename(columns={"Resolve_Min": "Count"}))
            fig_wk_d = px.pie(top10, names="Andon Type", values="Count", hole=0.55,
                              color_discrete_sequence=px.colors.qualitative.Set2,
                              title="Count of Problem Id by Andon Type (Top 10)")
            fig_wk_d.update_traces(textinfo="percent+label")
            fig_wk_d.add_annotation(text=f"{fdf['Resolve_Min'].count():,}", x=0.5, y=0.5,
                                    font_size=22, font_color="#1a237e", showarrow=False)
            fig_wk_d.update_layout(height=380, legend=dict(orientation="h", y=-0.35),
                                   margin=dict(t=50, b=10, l=0, r=0))
            st.plotly_chart(fig_wk_d, use_container_width=True)
        with c2:
            st.plotly_chart(hbar_chart(fdf, "Andon Type", "Average of Dwell Time Minutes by Andon Type"),
                            use_container_width=True)

    # ── Optional tabs ─────────────────────────────────────────────────────────
    if optional_cols["Equipment Type"]:
        with tab["By Equipment Type"]:
            st.markdown('<div class="sec-title">Number of Andons and Resolution Times by Equipment Type</div>', unsafe_allow_html=True)
            tbl_et = build_group_pivot(fdf, "Equipment Type")
            st.dataframe(apply_pivot_style(tbl_et), use_container_width=True, height=400)
            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(donut_chart(fdf, "Equipment Type", "Andons by Equipment Type"), use_container_width=True)
            with c2:
                st.plotly_chart(hbar_chart(fdf, "Equipment Type", "Avg Resolve Time by Equipment Type"), use_container_width=True)

    if optional_cols["Zone"]:
        with tab["By Zone"]:
            st.markdown('<div class="sec-title">Count of Resolver and Avg Dwell Time by Creation Date and Zone</div>', unsafe_allow_html=True)
            tbl_z = build_group_pivot(fdf, "Zone")
            st.dataframe(apply_pivot_style(tbl_z), use_container_width=True, height=400)
            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(donut_chart(fdf, "Zone", "Count of Records by Zone"), use_container_width=True)
            with c2:
                st.plotly_chart(hbar_chart(fdf, "Zone", "Avg Dwell Time by Zone"), use_container_width=True)

    if optional_cols["Shift"]:
        with tab["By Shift"]:
            st.markdown('<div class="sec-title">Count of Resolver and Avg Dwell Time by Creation Date and Shift</div>', unsafe_allow_html=True)
            tbl_sh = build_group_pivot(fdf, "Shift")
            st.dataframe(apply_pivot_style(tbl_sh), use_container_width=True, height=400)
            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(donut_chart(fdf, "Shift", "Count of Resolver by Shift"), use_container_width=True)
            with c2:
                st.plotly_chart(hbar_chart(fdf, "Shift", "Avg Dwell Time by Shift"), use_container_width=True)

    if optional_cols["Blocking"]:
        with tab["Blocking Analysis"]:
            st.markdown('<div class="sec-title">Number of Andons and Resolution Times by Resolver and Blocking Status</div>', unsafe_allow_html=True)
            bl_vals = sorted(fdf["Blocking"].dropna().unique())
            bl_res_cols = {}
            for bv in bl_vals:
                sub_b = fdf[fdf["Blocking"] == bv]
                bl_res_cols[(str(bv), "Andons")]   = sub_b.groupby("Resolver")["Resolve_Min"].count()
                bl_res_cols[(str(bv), "Avg Time")] = sub_b.groupby("Resolver")["Resolve_Min"].mean().round(2)
            all_res = sorted(fdf["Resolver"].unique())
            bl_res_cols[("Total", "Andons")]   = fdf.groupby("Resolver")["Resolve_Min"].count().reindex(all_res, fill_value=0).astype(int)
            bl_res_cols[("Total", "Avg Time")] = fdf.groupby("Resolver")["Resolve_Min"].mean().round(2).reindex(all_res)
            bl_tbl = pd.DataFrame(bl_res_cols).reindex(all_res)
            bl_tbl.columns = pd.MultiIndex.from_tuples(bl_tbl.columns)
            bl_tbl.index.name = "Resolver"
            bl_grand = {}
            for bv in bl_vals:
                sub_b = fdf[fdf["Blocking"] == bv]
                bl_grand[(str(bv), "Andons")]   = int(sub_b["Resolve_Min"].count())
                bl_grand[(str(bv), "Avg Time")] = round(sub_b["Resolve_Min"].mean(), 2)
            bl_grand[("Total", "Andons")]   = int(fdf["Resolve_Min"].count())
            bl_grand[("Total", "Avg Time")] = round(fdf["Resolve_Min"].mean(), 2)
            bl_grand_row = pd.DataFrame(bl_grand, index=["Grand Total"])
            bl_grand_row.columns = pd.MultiIndex.from_tuples(bl_grand_row.columns)
            bl_tbl = pd.concat([bl_tbl, bl_grand_row])
            bl_avg_cols_list = [(str(bv), "Avg Time") for bv in bl_vals] + [("Total", "Avg Time")]

            def _style_bl(data):
                s = pd.DataFrame("", index=data.index, columns=data.columns)
                rows = [i for i in data.index if i != "Grand Total"]
                for col in bl_avg_cols_list:
                    if col in data.columns:
                        ser = data.loc[rows, col]
                        for idx in rows:
                            s.loc[idx, col] = dwell_color(data.loc[idx, col], ser)
                s.loc["Grand Total"] = "font-weight:700; background-color:#e8eaf6; color:#1a237e"
                return s

            bl_styler = (bl_tbl.style.apply(_style_bl, axis=None)
                .format({c: "{:.2f}" for c in bl_tbl.columns if c[1] == "Avg Time"}, na_rep="—")
                .format({c: "{:,.0f}" for c in bl_tbl.columns if c[1] == "Andons"}, na_rep="—"))
            st.dataframe(bl_styler, use_container_width=True, height=420)
            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(donut_chart(fdf, "Blocking", "Blocking vs Non-Blocking Distribution"), use_container_width=True)
            with c2:
                st.plotly_chart(hbar_chart(fdf, "Blocking", "Avg Resolve Time: Blocking vs Non-Blocking"), use_container_width=True)

    if optional_cols["Equipment ID"]:
        with tab["Equipment ID Analysis"]:
            st.markdown('<div class="sec-title">Count of Problem Id by Equipment ID and Week Number</div>', unsafe_allow_html=True)
            eid_weeks = sorted(fdf["Week"].dropna().unique(), reverse=True)
            eid_count_p = fdf.pivot_table(index="Equipment ID", columns="Week",
                                          values="Resolve_Min", aggfunc="count", fill_value=0)
            eid_cols = {}
            for w in eid_weeks:
                if w in eid_count_p.columns:
                    eid_cols[f"Wk {w}"] = eid_count_p[w].astype(int)
            eid_cols["Total"] = eid_count_p[eid_weeks].sum(axis=1).astype(int)
            eid_tbl = pd.DataFrame(eid_cols).sort_values("Total", ascending=False)
            eid_tbl.index.name = "Equipment ID"
            eid_grand = {f"Wk {w}": int(fdf[fdf["Week"] == w]["Resolve_Min"].count()) for w in eid_weeks if w in eid_count_p.columns}
            eid_grand["Total"] = int(fdf["Resolve_Min"].count())
            eid_grand_row = pd.DataFrame(eid_grand, index=["Grand Total"])
            eid_tbl = pd.concat([eid_grand_row, eid_tbl])

            def _style_eid(data):
                s = pd.DataFrame("", index=data.index, columns=data.columns)
                if "Grand Total" in data.index:
                    s.loc["Grand Total"] = "font-weight:700; background-color:#e8eaf6; color:#1a237e"
                return s

            st.dataframe(eid_tbl.style.apply(_style_eid, axis=None).format("{:,.0f}", na_rep="—"),
                         use_container_width=True, height=420)
            c1, c2 = st.columns(2)
            with c1:
                top20_eid = (fdf.groupby("Equipment ID")["Resolve_Min"].count()
                             .nlargest(20).reset_index().rename(columns={"Resolve_Min": "Count"}))
                fig_eid_pie = px.pie(top20_eid, names="Equipment ID", values="Count",
                                    title="Top 20 Equipment IDs by Andon Count",
                                    color_discrete_sequence=px.colors.qualitative.Alphabet)
                fig_eid_pie.update_traces(textinfo="percent+label", textposition="inside")
                fig_eid_pie.update_layout(height=450, showlegend=False, margin=dict(t=50, b=10, l=0, r=0))
                st.plotly_chart(fig_eid_pie, use_container_width=True)
            with c2:
                top20_avg = (fdf.groupby("Equipment ID")["Resolve_Min"].agg(["count", "mean"])
                             .reset_index().nlargest(20, "count")
                             .rename(columns={"count": "Andons", "mean": "Avg Time (min)"}))
                fig_eid_bar = px.bar(top20_avg.sort_values("Avg Time (min)", ascending=True),
                                     x="Avg Time (min)", y="Equipment ID", orientation="h",
                                     color="Avg Time (min)", color_continuous_scale="Blues",
                                     text=top20_avg.sort_values("Avg Time (min)")["Avg Time (min)"].round(2),
                                     title="Avg Dwell Time — Top 20 Equipment IDs")
                fig_eid_bar.update_traces(textposition="outside")
                fig_eid_bar.update_layout(coloraxis_showscale=False, height=450,
                                          yaxis_title="", xaxis_title="Avg Dwell Time (min)",
                                          margin=dict(t=50, b=10, l=0, r=0))
                st.plotly_chart(fig_eid_bar, use_container_width=True)

    # ── Tab: Hourly Trend ─────────────────────────────────────────────────────
    with tab["Hourly Trend"]:
        st.markdown('<div class="sec-title">Count of Andon Type and Avg Dwell Time by Hour of Day</div>', unsafe_allow_html=True)
        hourly_count = (fdf.groupby(["Hour", "Andon Type"])["Resolve_Min"]
                        .count().reset_index().rename(columns={"Resolve_Min": "Count"}))
        hourly_avg   = (fdf.groupby("Hour")["Resolve_Min"]
                        .mean().reset_index().rename(columns={"Resolve_Min": "Avg Time"}))
        fig_h = go.Figure()
        colors = px.colors.qualitative.Pastel + px.colors.qualitative.Set2
        for i, at in enumerate(sorted(fdf["Andon Type"].dropna().unique())):
            sub = hourly_count[hourly_count["Andon Type"] == at]
            fig_h.add_trace(go.Bar(x=sub["Hour"], y=sub["Count"], name=at,
                                   marker_color=colors[i % len(colors)]))
        fig_h.add_trace(go.Scatter(
            x=hourly_avg["Hour"], y=hourly_avg["Avg Time"],
            name="Avg Dwell Time", yaxis="y2", mode="lines+markers",
            line=dict(color="#b71c1c", width=2.5), marker=dict(size=6, color="#b71c1c")
        ))
        fig_h.update_layout(barmode="stack", height=500,
                            xaxis=dict(title="Hour of Day", dtick=1),
                            yaxis=dict(title="Andon Count"),
                            yaxis2=dict(title="Avg Dwell Time (min)", overlaying="y", side="right", showgrid=False),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font_size=10),
                            margin=dict(t=80, b=40, l=0, r=0))
        st.plotly_chart(fig_h, use_container_width=True)

        st.markdown('<div class="sec-title">Resolve Time Trend by Day</div>', unsafe_allow_html=True)
        daily = (fdf.groupby("Date").agg(Count=("Resolve_Min", "count"), Avg=("Resolve_Min", "mean"))
                 .reset_index().sort_values("Date"))
        fig_d = go.Figure()
        fig_d.add_trace(go.Bar(x=daily["Date"], y=daily["Count"], name="Andons",
                               marker_color="#90caf9", yaxis="y"))
        fig_d.add_trace(go.Scatter(x=daily["Date"], y=daily["Avg"].round(2),
                                   name="Avg Dwell Time", yaxis="y2", mode="lines+markers",
                                   line=dict(color="#1a237e", width=2.5)))
        fig_d.update_layout(barmode="group", height=380,
                            xaxis=dict(title="Date"), yaxis=dict(title="Count"),
                            yaxis2=dict(title="Avg Dwell Time (min)", overlaying="y", side="right", showgrid=False),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02),
                            margin=dict(t=60, b=40, l=0, r=0))
        st.plotly_chart(fig_d, use_container_width=True)

    # ── Tab: Heatmap ──────────────────────────────────────────────────────────
    with tab["Heatmap"]:
        st.markdown('<div class="sec-title">Avg Resolve Time Heatmap — Resolver × Andon Type</div>', unsafe_allow_html=True)
        pivot_hm = fdf.pivot_table(index="Resolver", columns="Andon Type",
                                   values="Resolve_Min", aggfunc="mean").round(2)
        fig_hm = px.imshow(pivot_hm, text_auto=True, aspect="auto",
                           color_continuous_scale="RdYlGn_r", labels=dict(color="Avg (min)"))
        fig_hm.update_layout(height=max(350, len(pivot_hm) * 40 + 100), margin=dict(t=30, b=40, l=0, r=0))
        st.plotly_chart(fig_hm, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            res_perf = (fdf.groupby("Resolver")
                        .agg(Count=("Resolve_Min", "count"), Avg=("Resolve_Min", "mean"))
                        .reset_index().sort_values("Avg", ascending=False))
            fig_r = px.bar(res_perf, x="Resolver", y="Avg", color="Avg",
                           color_continuous_scale="RdYlGn_r", text=res_perf["Avg"].round(2),
                           labels={"Avg": "Avg Time (min)"})
            fig_r.update_traces(textposition="outside")
            fig_r.update_layout(coloraxis_showscale=False, height=380,
                                xaxis_title="", yaxis_title="Avg Dwell Time (min)",
                                margin=dict(t=20, b=40, l=0, r=0))
            st.plotly_chart(fig_r, use_container_width=True)
        with c2:
            cat_perf = (fdf.groupby("Andon Type")
                        .agg(Count=("Resolve_Min", "count"), Avg=("Resolve_Min", "mean"))
                        .reset_index().sort_values("Avg"))
            fig_c = px.bar(cat_perf, x="Avg", y="Andon Type", orientation="h",
                           color="Avg", color_continuous_scale="Blues", text=cat_perf["Avg"].round(2),
                           labels={"Avg": "Avg Time (min)"})
            fig_c.update_traces(textposition="outside")
            fig_c.update_layout(coloraxis_showscale=False, height=380,
                                yaxis_title="", xaxis_title="Avg Dwell Time (min)",
                                margin=dict(t=20, b=40, l=0, r=0))
            st.plotly_chart(fig_c, use_container_width=True)

    # ── Tab: Raw Data ─────────────────────────────────────────────────────────
    with tab["Raw Data"]:
        st.markdown('<div class="sec-title">Raw Resolved Andon Records</div>', unsafe_allow_html=True)
        st.markdown(f"**{len(fdf):,}** records matching current filters · from **{len(uploaded_files)}** file(s)")
        st.dataframe(fdf, use_container_width=True, height=500)
        dl1, dl2 = st.columns(2)
        with dl1:
            csv = fdf.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download as CSV", csv, "andon_filtered.csv", "text/csv", use_container_width=True)
        with dl2:
            import io
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                fdf.to_excel(writer, index=False, sheet_name="Andon Data")
                lb_export = lb[["Resolver", "Total_Andons", "Avg_Time", "Efficiency", "Within_Threshold", "Badge", "Status"]].copy()
                lb_export.columns = ["Resolver", "Total Andons", "Avg Time (min)", "Efficiency Score", "% Within Threshold", "Badge", "Status"]
                lb_export.to_excel(writer, index=True, index_label="Rank", sheet_name="Leaderboard")
            buf.seek(0)
            st.download_button("⬇️ Download as Excel", buf.getvalue(), "andon_report.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

    # ── Tab: Export ───────────────────────────────────────────────────────────
    with tab["📤 Export"]:
        st.markdown('<div class="sec-title">Export Full Reports</div>', unsafe_allow_html=True)
        st.markdown(
            "Each report is a **multi-sheet Excel workbook** — summary KPIs, all pivot tables, "
            "leaderboard, and embedded charts on a dedicated Charts sheet."
        )
        st.markdown("<br>", unsafe_allow_html=True)
        e1, e2 = st.columns(2)

        with e1:
            st.markdown("""
            <div style="background:#e8f5e9; border-left:5px solid #388e3c; border-radius:8px;
                        padding:16px 20px; margin-bottom:12px;">
                <div style="font-size:1.1rem; font-weight:700; color:#1b5e20;">📅 Daily Report</div>
                <div style="font-size:0.85rem; color:#2e7d32; margin-top:6px; line-height:1.6;">
                    ✅ Summary KPIs · ✅ AFM Performance · ✅ Andons by Type<br>
                    ✅ Resolver Leaderboard · ✅ Charts sheet · ✅ Raw Data
                </div>
            </div>""", unsafe_allow_html=True)
            with st.spinner("Generating Daily Report…"):
                daily_bytes = report_builder.build_daily_report(fdf, uploaded_files, within_threshold)
            st.download_button("⬇️ Download Daily Report (.xlsx)", daily_bytes,
                               "LCY3_Daily_Report.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

        with e2:
            st.markdown("""
            <div style="background:#e3f2fd; border-left:5px solid #1976d2; border-radius:8px;
                        padding:16px 20px; margin-bottom:12px;">
                <div style="font-size:1.1rem; font-weight:700; color:#0d47a1;">📆 Weekly Report</div>
                <div style="font-size:0.85rem; color:#1565c0; margin-top:6px; line-height:1.6;">
                    ✅ Weekly KPIs · ✅ Andon Type × Week · ✅ AFM × Week<br>
                    ✅ System vs Non-System · ✅ Leaderboard · ✅ Charts
                </div>
            </div>""", unsafe_allow_html=True)
            with st.spinner("Generating Weekly Report…"):
                weekly_bytes = report_builder.build_weekly_report(fdf, uploaded_files, within_threshold)
            st.download_button("⬇️ Download Weekly Report (.xlsx)", weekly_bytes,
                               "LCY3_Weekly_Report.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

        st.info("💡 Reports reflect your current filter selections. Adjust filters on any tab, then download.")

else:
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; background:#f5f5f5;
                border-radius:12px; margin-top:1rem; border: 2px dashed #c5cae9;">
        <h2 style="color:#3949ab; margin-bottom:0.5rem;">👋 Welcome to the LCY3 AFM Dashboard</h2>
        <p style="color:#666; font-size:1.05rem;">Upload a JSON or CSV file above to explore your Andon data</p>
        <p style="color:#999; font-size:0.85rem; margin-top:0.5rem;">
            Required columns: <strong>Status, Resolver, Andon Type, Dwell Time (hh:mm:ss), Time Created</strong><br>
            Optional columns: <strong>Equipment Type, Zone, Shift, Blocking, Equipment ID</strong>
        </p>
    </div>
    """, unsafe_allow_html=True)
