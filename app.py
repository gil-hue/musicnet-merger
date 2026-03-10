import streamlit as st
import pandas as pd
import os
import io
import yaml
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit_authenticator as stauth
from yaml.loader import SafeLoader

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="MusicNet Excel Merger",
    page_icon="🎵",
    layout="wide"
)

# ── Authentication ────────────────────────────────────────────
with open("config.yaml") as f:
    config = yaml.load(f, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
)

authenticator.login(location="main")

if st.session_state.get("authentication_status") is False:
    st.error("שם משתמש או סיסמה שגויים")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.warning("אנא הכנס שם משתמש וסיסמה")
    st.stop()

# ── Logged in ─────────────────────────────────────────────────
col_user, col_logout = st.columns([8, 1])
with col_logout:
    authenticator.logout("התנתק")

# ── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fb; }
    .block-container { padding-top: 2rem; }
    .title-bar {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .title-bar h1 { color: white; margin: 0; font-size: 2rem; }
    .title-bar p  { color: #cce0f5; margin: 0.3rem 0 0 0; font-size: 1rem; }
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 4px solid #2E75B6;
    }
    .metric-card .num  { font-size: 2.2rem; font-weight: 700; color: #1F3864; }
    .metric-card .lbl  { font-size: 0.85rem; color: #666; margin-top: 0.2rem; }
    .file-row {
        background: white;
        border-radius: 8px;
        padding: 0.7rem 1rem;
        margin-bottom: 0.4rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .stButton > button {
        background: linear-gradient(135deg, #1F3864, #2E75B6);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        font-size: 1rem;
        font-weight: 600;
        width: 100%;
    }
    .stButton > button:hover { opacity: 0.9; }
    .success-box {
        background: #e8f5e9;
        border-left: 4px solid #43a047;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        color: #2e7d32;
        font-weight: 600;
    }
    .section-header {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1F3864;
        border-bottom: 2px solid #2E75B6;
        padding-bottom: 0.4rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────
st.markdown("""
<div class="title-bar">
    <h1>🎵 MusicNet Excel Merger</h1>
    <p>העלה קבצי Excel, בחר עמודות, וקבל קובץ מאוחד להורדה</p>
</div>
""", unsafe_allow_html=True)

# ── Helper: build Excel bytes ─────────────────────────────────
def build_excel(summary_data, merged_df, col_headers):
    wb = Workbook()

    header_fill = PatternFill("solid", start_color="1F3864")
    sub_fill    = PatternFill("solid", start_color="2E75B6")
    total_fill  = PatternFill("solid", start_color="D6E4F0")
    white_fill  = PatternFill("solid", start_color="FFFFFF")
    alt_fill    = PatternFill("solid", start_color="EBF2FA")

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=14)
    col_font    = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    data_font   = Font(name="Arial", size=11)
    total_font  = Font(name="Arial", bold=True, size=11, color="1F3864")

    thin   = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")

    # ── Summary sheet ─────────────────────────────────────────
    ws_s = wb.active
    ws_s.title = "סיכום"

    ws_s.merge_cells("A1:B1")
    ws_s["A1"] = "סיכום קבצי Excel שעובדו"
    ws_s["A1"].font = header_font
    ws_s["A1"].fill = header_fill
    ws_s["A1"].alignment = center
    ws_s.row_dimensions[1].height = 32

    ws_s.merge_cells("A2:B2")
    ws_s["A2"] = "תאריך עיבוד: " + date.today().strftime("%d/%m/%Y")
    ws_s["A2"].font = Font(name="Arial", italic=True, size=10, color="595959")
    ws_s["A2"].alignment = center
    ws_s.row_dimensions[2].height = 18

    for col, txt in [(1, "שם הקובץ"), (2, "מספר רשומות")]:
        c = ws_s.cell(3, col, txt)
        c.font = col_font; c.fill = sub_fill
        c.alignment = center; c.border = border
    ws_s.row_dimensions[3].height = 24

    for i, (fname, cnt) in enumerate(summary_data):
        r = i + 4
        fill = white_fill if i % 2 == 0 else alt_fill
        c1 = ws_s.cell(r, 1, fname)
        c2 = ws_s.cell(r, 2, cnt)
        for c in [c1, c2]:
            c.font = data_font; c.fill = fill; c.border = border
        c1.alignment = left; c2.alignment = center
        ws_s.row_dimensions[r].height = 22

    tr = len(summary_data) + 4
    total_label = 'סה"כ רשומות'
    c1 = ws_s.cell(tr, 1, total_label)
    c2 = ws_s.cell(tr, 2, "=SUM(B4:B" + str(tr - 1) + ")")
    for c in [c1, c2]:
        c.font = total_font; c.fill = total_fill; c.border = border
    c1.alignment = left; c2.alignment = center
    ws_s.row_dimensions[tr].height = 24
    ws_s.column_dimensions["A"].width = 35
    ws_s.column_dimensions["B"].width = 20

    # ── Data sheet ────────────────────────────────────────────
    ws_d = wb.create_sheet("נתונים מאוחדים")
    col_keys = list(merged_df.columns)

    for ci, col in enumerate(col_keys, 1):
        hdr = col_headers.get(col, col)
        c = ws_d.cell(1, ci, hdr)
        c.font = col_font; c.fill = sub_fill
        c.alignment = center; c.border = border
        ws_d.column_dimensions[get_column_letter(ci)].width = 32
    ws_d.row_dimensions[1].height = 26

    for ri, row in enumerate(merged_df.itertuples(index=False), 2):
        fill = white_fill if ri % 2 == 0 else alt_fill
        for ci, val in enumerate(row, 1):
            c = ws_d.cell(ri, ci, str(val) if pd.notna(val) else "")
            c.font = data_font; c.fill = fill
            c.alignment = left; c.border = border
        ws_d.row_dimensions[ri].height = 18

    ws_d.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════
#  MAIN UI
# ══════════════════════════════════════════════════════════════

# ── Step 1: Upload ────────────────────────────────────────────
st.markdown('<div class="section-header">📂 שלב 1 — העלאת קבצים</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "גרור קבצי Excel לכאן או לחץ לבחירה",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if not uploaded:
    st.info("⬆️ העלה לפחות קובץ Excel אחד כדי להמשיך.")
    st.stop()

# ── Read files ────────────────────────────────────────────────
file_data = {}
all_columns = set()

for f in uploaded:
    try:
        df = pd.read_excel(f)
        file_data[f.name] = df
        all_columns.update(df.columns.tolist())
    except Exception as e:
        st.error(f"שגיאה בקריאת {f.name}: {e}")

all_columns = sorted(all_columns)

# ── Metrics ───────────────────────────────────────────────────
total_recs = sum(len(d) for d in file_data.values())
c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f'<div class="metric-card"><div class="num">{len(file_data)}</div><div class="lbl">קבצים הועלו</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-card"><div class="num">{total_recs:,}</div><div class="lbl">רשומות סה"כ</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="metric-card"><div class="num">{len(all_columns)}</div><div class="lbl">עמודות נמצאו</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── File summary table ────────────────────────────────────────
with st.expander("📋 פירוט רשומות לפי קובץ", expanded=True):
    rows = [{"📄 שם קובץ": name, "📊 רשומות": f"{len(df):,}", "📑 עמודות": len(df.columns)}
            for name, df in file_data.items()]
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ── Step 2: Column selection ──────────────────────────────────
st.markdown('<div class="section-header">🔧 שלב 2 — בחירת עמודות לשמור</div>', unsafe_allow_html=True)

default_cols = [c for c in ['album_name', 'artist_name', 'track_name', 'track_uri', 'label'] if c in all_columns]

selected_cols = st.multiselect(
    "בחר את העמודות שברצונך לשמור בקובץ המאוחד:",
    options=all_columns,
    default=default_cols
)

if not selected_cols:
    st.warning("⚠️ יש לבחור לפחות עמודה אחת.")
    st.stop()

# ── Step 3: Process ───────────────────────────────────────────
st.markdown('<div class="section-header">⚙️ שלב 3 — עיבוד ויצוא</div>', unsafe_allow_html=True)

col_headers = {
    'album_name':  'שם אלבום',
    'artist_name': 'שם אמן',
    'track_name':  'שם שיר',
    'track_uri':   'Track URI',
    'label':       'לייבל'
}

if st.button("🚀 עבד וצור קובץ מאוחד"):
    with st.spinner("מעבד קבצים..."):
        summary_data = []
        dfs = []

        for fname, df in file_data.items():
            existing = [c for c in selected_cols if c in df.columns]
            missing  = [c for c in selected_cols if c not in df.columns]
            dfs.append(df[existing])
            summary_data.append((fname, len(df)))
            if missing:
                st.warning(f"⚠️ {fname}: עמודות חסרות — {missing}")

        merged = pd.concat(dfs, ignore_index=True)
        today  = date.today().strftime("%Y-%m-%d")
        out_name = f"MusicNet_Merged_{today}.xlsx"

        excel_bytes = build_excel(summary_data, merged, col_headers)

    st.markdown(f"""
    <div class="success-box">
        ✅ הקובץ המאוחד נוצר בהצלחה!<br>
        סה"כ רשומות: <strong>{len(merged):,}</strong> | עמודות: <strong>{len(merged.columns)}</strong>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.download_button(
        label="⬇️ הורד את הקובץ המאוחד",
        data=excel_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    with st.expander("📊 תצוגה מקדימה — 20 שורות ראשונות"):
        st.dataframe(merged.head(20), use_container_width=True, hide_index=True)
