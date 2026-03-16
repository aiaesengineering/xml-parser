# import streamlit as st
# import pandas as pd
# import xml.etree.ElementTree as ET
# import re
# import io
# import datetime
# from pathlib import Path
# import unicodedata

# st.set_page_config(page_title="Navisworks Clash Analyzer", layout="wide")

# # ── Styling ────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
#     .stApp { background: #f8f9fa; }
#     .block-container { max-width: 1100px; padding-top: 2rem; }
#     h1, h2, h3 { color: #1a1a1a; }
#     .metric { background: #f0f4ff; border-radius: 8px; padding: 16px; text-align: center; }
#     .metric-val { font-size: 2.1rem; font-weight: 700; color: #2563eb; }
#     .metric-label { font-size: 0.85rem; color: #555; text-transform: uppercase; letter-spacing: 0.6px; }
# </style>
# """, unsafe_allow_html=True)


# # ── File & name handling ────────────────────────────────────────────────────

# PROJECT_FOLDER = Path("project_reports")
# PROJECT_FOLDER.mkdir(exist_ok=True, parents=True)

# def normalize_project_name(raw: str) -> str:
#     if not raw or not str(raw).strip():
#         return "default_project"
#     s = str(raw).strip()
#     s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
#     s = re.sub(r'[^a-zA-Z0-9_-]+', '_', s.lower())
#     s = re.sub(r'_+', '_', s).strip('_')
#     return s if s else "default_project"


# def get_project_filepath(project_input: str) -> Path:
#     norm = normalize_project_name(project_input)
#     return PROJECT_FOLDER / f"{norm}_clash_report.xlsx"


# # ── Weekly Progress helpers (ALWAYS append new row – even same day for testing) ──

# def load_weekly_progress(filepath: Path) -> pd.DataFrame:
#     if not filepath.is_file():
#         return pd.DataFrame(columns=["Date", "Week", "Open", "Closed"])
#     try:
#         with pd.ExcelFile(filepath, engine="openpyxl") as xls:
#             if "Weekly_Progress" in xls.sheet_names:
#                 df = pd.read_excel(xls, sheet_name="Weekly_Progress")
#                 if "Date" in df.columns:
#                     df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
#                 return df.reset_index(drop=True)
#     except:
#         pass
#     return pd.DataFrame(columns=["Date", "Week", "Open", "Closed"])


# def append_progress(existing: pd.DataFrame, open_c: int, closed_c: int) -> pd.DataFrame:
#     today = datetime.date.today()
#     week = f"{today.year}-W{today.isocalendar()[1]:02d}"

#     new_row = pd.DataFrame([{
#         "Date": today,
#         "Week": week,
#         "Open": open_c,
#         "Closed": closed_c
#     }])

#     updated = pd.concat([existing, new_row], ignore_index=True)
#     return updated.sort_values("Date").reset_index(drop=True)


# # ── Prefix helpers ──────────────────────────────────────────────────────────

# def extract_prefix_regex(test_name):
#     if not test_name: return None
#     matches = re.findall(r'(?:^|_)([A-Z]{2,4})-', test_name)
#     return matches[0] if matches else None


# def extract_prefix_position(test_name, pos):
#     if not test_name: return None
#     parts = test_name.split("_")
#     if len(parts) >= pos:
#         return parts[pos-1].split("-")[0].strip() or None
#     return None


# # ── XML Parser ──────────────────────────────────────────────────────────────

# def parse_xml(file_bytes: bytes, use_manual: bool, position: int | None):
#     root = ET.parse(io.BytesIO(file_bytes)).getroot()
#     rows = []

#     for test in root.iter("clashtest"):
#         name = test.get("name")
#         prio = test.get("priority")
#         prefix = extract_prefix_position(name, position) if use_manual else extract_prefix_regex(name)

#         clashes = list(test.iter("clashresult")) or list(test.iter("clashgroup"))
#         total = len(clashes)

#         for clash in clashes:
#             row = {
#                 "Test Name": name,
#                 "Prefix": prefix,
#                 "Test Priority": prio,
#                 "Total Clashes In Test": total,
#             }

#             row.update({f"clash_{k}": v for k,v in clash.attrib.items()})

#             for el in clash.iter():
#                 if el.text and el.text.strip():
#                     row[el.tag] = el.text.strip()
#                 row.update({f"{el.tag}_{k}": v for k,v in el.attrib.items()})

#             for oa in clash.findall(".//objectattribute"):
#                 n = oa.findtext("name")
#                 v = oa.findtext("value")
#                 if n and v:
#                     row[n] = v

#             rows.append(row)

#     return rows


# # ── Main Excel update logic ────────────────────────────────────────────────

# def update_project_file(df_new, project_input, open_cnt, closed_cnt):
#     path = get_project_filepath(project_input)

#     # === Clash_Details & Status_Summary & Prefix_Summary = REPLACE with latest XML ===
#     # === Weekly_Progress = ALWAYS APPEND new row (even same day) ===

#     # 1. Prepare latest Clash_Details (replace)
#     combined_details = df_new.copy()

#     # 2. Priority mapping – updated as requested: 1 = H, 2 = M, 3 = L
#     if "Test Priority" in combined_details.columns:
#         prio_map = {
#             1: "H", "1": "H",
#             2: "M", "2": "M",
#             3: "L", "3": "L"
#         }
#         combined_details["Priority"] = combined_details["Test Priority"].map(prio_map).fillna("M")

#     # 3. Recalculate Status_Summary & Prefix_Summary from latest data
#     status_summary = None
#     prefix_summary = None

#     status_col = next((c for c in combined_details.columns if "status" in str(c).lower()), None)
#     if status_col:
#         lower = combined_details[status_col].astype(str).str.lower()
#         combined_details["Open"]   = lower.isin(["new", "active", "reviewed"]).astype(int)
#         combined_details["Closed"] = lower.isin(["approved", "resolved"]).astype(int)

#         status_summary = lower.value_counts().reset_index(name="Count").rename(columns={status_col: "Status"})

#     # if "Prefix" in combined_details.columns:
#     #     prefix_summary = combined_details["Prefix"].dropna().value_counts().reset_index(name="Clash Count")

#     if "Prefix" in combined_details.columns:

#         df_prefix = combined_details.copy()

#         # Total clashes per prefix
#         total_counts = df_prefix.groupby("Prefix").size().reset_index(name="Clash Count")

#         # Priority counts per prefix
#         priority_counts = (
#             df_prefix.groupby(["Prefix", "Priority"])
#             .size()
#             .unstack(fill_value=0)
#             .reset_index()
#         )

#         # Rename columns for clarity
#         priority_counts = priority_counts.rename(columns={
#             "H": "High",
#             "M": "Medium",
#             "L": "Low"
#         })

#         # Ensure columns exist even if missing in XML
#         for col in ["High", "Medium", "Low"]:
#             if col not in priority_counts.columns:
#                 priority_counts[col] = 0

#         # Merge totals + priority breakdown
#         prefix_summary = pd.merge(total_counts, priority_counts, on="Prefix", how="left")

#         # Final column order
#         prefix_summary = prefix_summary[["Prefix", "Clash Count", "High", "Medium", "Low"]]

#         # Sort by most clashes
#         prefix_summary = prefix_summary.sort_values("Clash Count", ascending=False)

#         # 4. Weekly_Progress – always append new row
#         existing_weekly = load_weekly_progress(path)
#         weekly_df = append_progress(existing_weekly, open_cnt, closed_cnt)

#         # 5. Write all 4 sheets back
#         output = io.BytesIO()
#         with pd.ExcelWriter(output, engine="openpyxl") as writer:
#             combined_details.to_excel(writer, sheet_name="Clash_Details", index=False)
#             if status_summary is not None:
#                 status_summary.to_excel(writer, sheet_name="Status_Summary", index=False)
#             if prefix_summary is not None:
#                 prefix_summary.to_excel(writer, sheet_name="Prefix_Summary", index=False)
#             weekly_df.to_excel(writer, sheet_name="Weekly_Progress", index=False)

#         bytes_data = output.getvalue()
#         path.write_bytes(bytes_data)

#     return bytes_data, weekly_df, str(path.absolute())


# # ── UI ──────────────────────────────────────────────────────────────────────

# st.title("🏗️ Navisworks Clash → Excel Analyzer")

# for k in ["df", "excel_bytes", "cached_bytes", "last_file"]:
#     if k not in st.session_state:
#         st.session_state[k] = None

# st.subheader("1. Upload Navisworks Clash XML")
# uploaded = st.file_uploader("Choose XML file", type=["xml"])

# if uploaded:
#     project_input = st.text_input(
#         "2. Project name (same name = same file)",
#         value="MyProject",
#         help="Use the **same name every time**"
#     )

#     computed_fn = get_project_filepath(project_input).name
#     full_path = str(get_project_filepath(project_input).absolute())

#     with st.expander("📂 File that will be updated", expanded=False):
#         st.markdown(f"**Filename:** `{computed_fn}`")
#         st.markdown(f"**Full path:** `{full_path}`")

#     if st.session_state.last_file != uploaded.name:
#         st.session_state.cached_bytes = uploaded.read()
#         st.session_state.last_file = uploaded.name
#         st.session_state.df = None
#         st.session_state.excel_bytes = None

#     st.caption(f"Loaded: **{uploaded.name}**  •  {len(st.session_state.cached_bytes)/1024/1024:.1f} MB")

#     col1, col2 = st.columns([3, 2])
#     with col1:
#         mode = st.radio("3. Prefix detection", ["Auto (Regex)", "Manual (position after _ split)"], horizontal=True)
#     use_manual = mode.startswith("Manual")
#     position = None
#     if use_manual:
#         with col2:
#             position = st.number_input("Position", min_value=1, max_value=15, value=2, step=1)

#     st.divider()

#     if st.button("Process & Append to Project File", type="primary", use_container_width=True):
#         with st.spinner("Parsing XML + appending to project file..."):
#             try:
#                 rows = parse_xml(st.session_state.cached_bytes, use_manual, position)
#                 if not rows:
#                     st.error("No clashes found in this XML.")
#                 else:
#                     df_new = pd.DataFrame(rows)
#                     st.session_state.df = df_new

#                     status_col = next((c for c in df_new.columns if "status" in str(c).lower()), None)
#                     open_c = closed_c = 0
#                     if status_col:
#                         lower = df_new[status_col].astype(str).str.lower()
#                         open_c  = lower.isin(["new","active","reviewed"]).sum()
#                         closed_c = lower.isin(["approved","resolved"]).sum()

#                     bytes_data, hist_df, path_str = update_project_file(df_new, project_input, open_c, closed_c)
#                     st.session_state.excel_bytes = bytes_data

#                     st.success(f"**File updated** ✓\n\n"
#                                f"This XML: **{len(df_new):,}** clashes • Open: **{open_c:,}** • Closed: **{closed_c:,}**\n"
#                                f"Saved to: `{path_str}`")

#             except Exception as e:
#                 st.error(f"Error: {str(e)}")

#     if st.session_state.df is not None:
#         df = st.session_state.df
#         st.divider()

#         cols = st.columns(4)
#         cols[0].metric("This XML – Clashes", f"{len(df):,}")
#         cols[1].metric("This XML – Tests", df["Test Name"].nunique())
#         cols[2].metric("This XML – Prefixes", df["Prefix"].nunique() if "Prefix" in df.columns else 0)
#         cols[3].metric("Columns", len(df.columns))

#         if "Prefix" in df.columns:
#             st.subheader("Prefix Breakdown (this XML only)")
#             cnt = df["Prefix"].value_counts().reset_index()
#             cnt.columns = ["Prefix", "Clash Count"]
#             t, c = st.columns([1, 3])
#             with t:
#                 st.dataframe(cnt.head(15), hide_index=True, use_container_width=True)
#             with c:
#                 st.bar_chart(cnt.head(15).set_index("Prefix")["Clash Count"], color="#4e79a7")

#         # ── Open Clashes by Priority (H / M / L) ───────────────────────
#         if "Priority" in df.columns:
#             st.subheader("Open Clashes by Priority")
#             prio_open = df.groupby("Priority")["Open"].sum().reset_index()
#             st.bar_chart(prio_open.set_index("Priority")["Open"], color="#2563eb")

#         with st.expander("Preview this XML (top 150 rows)"):
#             st.dataframe(df.head(150))

#         if st.session_state.excel_bytes:
#             p = get_project_filepath(project_input)
#             st.download_button(
#                 "⬇ Download Project Excel (all accumulated data)",
#                 data=st.session_state.excel_bytes,
#                 file_name=p.name,
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                 use_container_width=True
#             )

#         # Show accumulated weekly progress
#         try:
#             hist = load_weekly_progress(p)
#             if not hist.empty:
#                 st.subheader("Weekly Progress History (accumulated)")
#                 st.dataframe(hist, use_container_width=True)
#                 if len(hist) > 1:
#                     st.line_chart(hist.set_index("Date")[["Open", "Closed"]])
#         except:
#             pass

# else:
#     st.info("Upload a clash XML to start.")













import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import re
import io
import datetime
from pathlib import Path
import unicodedata

import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Navisworks Clash Analyzer", layout="wide")

page = st.sidebar.selectbox(
    "Navigation",
    ["XML Processor", "Clash Dashboard"]
)

# ── Styling ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background: #f8f9fa; }
    .block-container { max-width: 1100px; padding-top: 2rem; }
    h1, h2, h3 { color: #1a1a1a; }
    .metric { background: #f0f4ff; border-radius: 8px; padding: 16px; text-align: center; }
    .metric-val { font-size: 2.1rem; font-weight: 700; color: #2563eb; }
    .metric-label { font-size: 0.85rem; color: #555; text-transform: uppercase; letter-spacing: 0.6px; }
</style>
""", unsafe_allow_html=True)


# ── File & name handling ────────────────────────────────────────────────────

PROJECT_FOLDER = Path("project_reports")
PROJECT_FOLDER.mkdir(exist_ok=True, parents=True)

def normalize_project_name(raw: str) -> str:
    if not raw or not str(raw).strip():
        return "default_project"
    s = str(raw).strip()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    s = re.sub(r'[^a-zA-Z0-9_-]+', '_', s.lower())
    s = re.sub(r'_+', '_', s).strip('_')
    return s if s else "default_project"


def get_project_filepath(project_input: str) -> Path:
    norm = normalize_project_name(project_input)
    return PROJECT_FOLDER / f"{norm}_clash_report.xlsx"


# ── Weekly Progress helpers (ALWAYS append new row – even same day for testing) ──

def load_weekly_progress(filepath: Path) -> pd.DataFrame:
    if not filepath.is_file():
        return pd.DataFrame(columns=["Date", "Week", "Open", "Closed"])
    try:
        with pd.ExcelFile(filepath, engine="openpyxl") as xls:
            if "Weekly_Progress" in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name="Weekly_Progress")
                if "Date" in df.columns:
                    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
                return df.reset_index(drop=True)
    except:
        pass
    return pd.DataFrame(columns=["Date", "Week", "Open", "Closed"])


def append_progress(existing: pd.DataFrame, open_c: int, closed_c: int) -> pd.DataFrame:
    today = datetime.date.today()
    week = f"{today.year}-W{today.isocalendar()[1]:02d}"

    new_row = pd.DataFrame([{
        "Date": today,
        "Week": week,
        "Open": open_c,
        "Closed": closed_c
    }])

    updated = pd.concat([existing, new_row], ignore_index=True)
    return updated.sort_values("Date").reset_index(drop=True)


# ── Prefix helpers ──────────────────────────────────────────────────────────

def extract_prefix_regex(test_name):
    if not test_name: return None
    matches = re.findall(r'(?:^|_)([A-Z]{2,4})-', test_name)
    return matches[0] if matches else None


def extract_prefix_position(test_name, pos):
    if not test_name: return None
    parts = test_name.split("_")
    if len(parts) >= pos:
        return parts[pos-1].split("-")[0].strip() or None
    return None


# ── XML Parser ──────────────────────────────────────────────────────────────

def parse_xml(file_bytes: bytes, use_manual: bool, position: int | None):
    root = ET.parse(io.BytesIO(file_bytes)).getroot()
    rows = []

    for test in root.iter("clashtest"):
        name = test.get("name")
        prio = test.get("priority")
        prefix = extract_prefix_position(name, position) if use_manual else extract_prefix_regex(name)

        clashes = list(test.iter("clashresult")) or list(test.iter("clashgroup"))
        total = len(clashes)

        for clash in clashes:
            row = {
                "Test Name": name,
                "Prefix": prefix,
                "Test Priority": prio,
                "Total Clashes In Test": total,
            }

            row.update({f"clash_{k}": v for k,v in clash.attrib.items()})

            for el in clash.iter():
                if el.text and el.text.strip():
                    row[el.tag] = el.text.strip()
                row.update({f"{el.tag}_{k}": v for k,v in el.attrib.items()})

            for oa in clash.findall(".//objectattribute"):
                n = oa.findtext("name")
                v = oa.findtext("value")
                if n and v:
                    row[n] = v

            rows.append(row)

    return rows


# ── Main Excel update logic ────────────────────────────────────────────────

def update_project_file(df_new, project_input, open_cnt, closed_cnt):
    path = get_project_filepath(project_input)

    # === Clash_Details & Status_Summary & Prefix_Summary = REPLACE with latest XML ===
    # === Weekly_Progress = ALWAYS APPEND new row (even same day) ===

    # 1. Prepare latest Clash_Details (replace)
    combined_details = df_new.copy()

    # 2. Priority mapping – updated as requested: 1 = H, 2 = M, 3 = L
    if "Test Priority" in combined_details.columns:
        prio_map = {
            1: "H", "1": "H",
            2: "M", "2": "M",
            3: "L", "3": "L"
        }
        combined_details["Priority"] = combined_details["Test Priority"].map(prio_map).fillna("M")

    # 3. Recalculate Status_Summary & Prefix_Summary from latest data
    status_summary = None
    prefix_summary = None

    status_col = next((c for c in combined_details.columns if "status" in str(c).lower()), None)
    if status_col:
        lower = combined_details[status_col].astype(str).str.lower()
        combined_details["Open"]   = lower.isin(["new", "active", "reviewed"]).astype(int)
        combined_details["Closed"] = lower.isin(["approved", "resolved"]).astype(int)

        status_summary = lower.value_counts().reset_index(name="Count").rename(columns={status_col: "Status"})

    # if "Prefix" in combined_details.columns:
    #     prefix_summary = combined_details["Prefix"].dropna().value_counts().reset_index(name="Clash Count")

    if "Prefix" in combined_details.columns:

        df_prefix = combined_details.copy()

        # Total clashes per prefix
        total_counts = df_prefix.groupby("Prefix").size().reset_index(name="Clash Count")

        # Priority counts per prefix
        priority_counts = (
            df_prefix.groupby(["Prefix", "Priority"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )

        # Rename columns for clarity
        priority_counts = priority_counts.rename(columns={
            "H": "High",
            "M": "Medium",
            "L": "Low"
        })

        # Ensure columns exist even if missing in XML
        for col in ["High", "Medium", "Low"]:
            if col not in priority_counts.columns:
                priority_counts[col] = 0

        # Merge totals + priority breakdown
        prefix_summary = pd.merge(total_counts, priority_counts, on="Prefix", how="left")

        # Final column order
        prefix_summary = prefix_summary[["Prefix", "Clash Count", "High", "Medium", "Low"]]

        # Sort by most clashes
        prefix_summary = prefix_summary.sort_values("Clash Count", ascending=False)

        # 4. Weekly_Progress – always append new row
        existing_weekly = load_weekly_progress(path)
        weekly_df = append_progress(existing_weekly, open_cnt, closed_cnt)

        # 5. Write all 4 sheets back
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            combined_details.to_excel(writer, sheet_name="Clash_Details", index=False)
            if status_summary is not None:
                status_summary.to_excel(writer, sheet_name="Status_Summary", index=False)
            if prefix_summary is not None:
                prefix_summary.to_excel(writer, sheet_name="Prefix_Summary", index=False)
            weekly_df.to_excel(writer, sheet_name="Weekly_Progress", index=False)

        bytes_data = output.getvalue()
        path.write_bytes(bytes_data)

    return bytes_data, weekly_df, str(path.absolute())


# ── UI ──────────────────────────────────────────────────────────────────────

# st.title("🏗️ Navisworks Clash → Excel Analyzer")
if page == "XML Processor":

    st.title("🏗️ Navisworks Clash → Excel Analyzer")

for k in ["df", "excel_bytes", "cached_bytes", "last_file"]:
    if k not in st.session_state:
        st.session_state[k] = None

st.subheader("1. Upload Navisworks Clash XML")
uploaded = st.file_uploader("Choose XML file", type=["xml"])

if uploaded:
    st.write(f"File size: {uploaded.size / 1024 / 1024:.1f} MB")
    project_input = st.text_input(
        "2. Project name (same name = same file)",
        value="MyProject",
        help="Use the **same name every time**"
    )

    computed_fn = get_project_filepath(project_input).name
    full_path = str(get_project_filepath(project_input).absolute())

    with st.expander("📂 File that will be updated", expanded=False):
        st.markdown(f"**Filename:** `{computed_fn}`")
        st.markdown(f"**Full path:** `{full_path}`")

    if st.session_state.last_file != uploaded.name:
        st.session_state.cached_bytes = uploaded.read()
        st.session_state.last_file = uploaded.name
        st.session_state.df = None
        st.session_state.excel_bytes = None

    st.caption(f"Loaded: **{uploaded.name}**  •  {len(st.session_state.cached_bytes)/1024/1024:.1f} MB")

    col1, col2 = st.columns([3, 2])
    with col1:
        mode = st.radio("3. Prefix detection", ["Auto (Regex)", "Manual (position after _ split)"], horizontal=True)
    use_manual = mode.startswith("Manual")
    position = None
    if use_manual:
        with col2:
            position = st.number_input("Position", min_value=1, max_value=15, value=2, step=1)

    st.divider()

    if st.button("Process & Append to Project File", type="primary", use_container_width=True):
        with st.spinner("Parsing XML + appending to project file..."):
            try:
                rows = parse_xml(st.session_state.cached_bytes, use_manual, position)
                if not rows:
                    st.error("No clashes found in this XML.")
                else:
                    df_new = pd.DataFrame(rows)
                    st.session_state.df = df_new

                    status_col = next((c for c in df_new.columns if "status" in str(c).lower()), None)
                    open_c = closed_c = 0
                    if status_col:
                        lower = df_new[status_col].astype(str).str.lower()
                        open_c  = lower.isin(["new","active","reviewed"]).sum()
                        closed_c = lower.isin(["approved","resolved"]).sum()

                    bytes_data, hist_df, path_str = update_project_file(df_new, project_input, open_c, closed_c)
                    st.session_state.excel_bytes = bytes_data

                    st.success(f"**File updated** ✓\n\n"
                               f"This XML: **{len(df_new):,}** clashes • Open: **{open_c:,}** • Closed: **{closed_c:,}**\n"
                               f"Saved to: `{path_str}`")

            except Exception as e:
                st.error(f"Error: {str(e)}")

    if st.session_state.df is not None:
        df = st.session_state.df
        st.divider()

        cols = st.columns(4)
        cols[0].metric("This XML – Clashes", f"{len(df):,}")
        cols[1].metric("This XML – Tests", df["Test Name"].nunique())
        cols[2].metric("This XML – Prefixes", df["Prefix"].nunique() if "Prefix" in df.columns else 0)
        cols[3].metric("Columns", len(df.columns))

        if "Prefix" in df.columns:
            st.subheader("Prefix Breakdown (this XML only)")
            cnt = df["Prefix"].value_counts().reset_index()
            cnt.columns = ["Prefix", "Clash Count"]
            t, c = st.columns([1, 3])
            with t:
                st.dataframe(cnt.head(15), hide_index=True, use_container_width=True)
            with c:
                st.bar_chart(cnt.head(15).set_index("Prefix")["Clash Count"], color="#4e79a7")

        # ── Open Clashes by Priority (H / M / L) ───────────────────────
        if "Priority" in df.columns:
            st.subheader("Open Clashes by Priority")
            prio_open = df.groupby("Priority")["Open"].sum().reset_index()
            st.bar_chart(prio_open.set_index("Priority")["Open"], color="#2563eb")

        with st.expander("Preview this XML (top 150 rows)"):
            st.dataframe(df.head(150))

        if st.session_state.excel_bytes:
            p = get_project_filepath(project_input)
            st.download_button(
                "⬇ Download Project Excel (all accumulated data)",
                data=st.session_state.excel_bytes,
                file_name=p.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # Show accumulated weekly progress
        try:
            hist = load_weekly_progress(p)
            if not hist.empty:
                st.subheader("Weekly Progress History (accumulated)")
                st.dataframe(hist, use_container_width=True)
                if len(hist) > 1:
                    st.line_chart(hist.set_index("Date")[["Open", "Closed"]])
        except:
            pass

else:
    st.info("Upload a clash XML to start.")



if page == "Clash Dashboard":

    st.title("📊 Clash Coordination Dashboard")

    project_input = st.text_input("Project name", value="MyProject")

    path = get_project_filepath(project_input)

    if not path.exists():
        st.warning("Project file not found. Upload XML first.")
        st.stop()

    df = pd.read_excel(path, sheet_name="Clash_Details")

    # Detect status column
    status_col = next((c for c in df.columns if "status" in str(c).lower()), None)

    if status_col:
        df["Status"] = df[status_col].astype(str)

    # ------------------------------
    # Summary Metrics
    # ------------------------------

    total_clashes = len(df)

    open_count = df["Open"].sum() if "Open" in df.columns else 0
    closed_count = df["Closed"].sum() if "Closed" in df.columns else 0

    col1, col2, col3 = st.columns(3)

    col1.metric("Total Clashes", f"{total_clashes:,}")
    col2.metric("Open Clashes", f"{open_count:,}")
    col3.metric("Closed Clashes", f"{closed_count:,}")

    st.divider()


    # ------------------------------
    # Clash Progress Over Time
    # ------------------------------

    st.subheader("Clash Resolution Progress")

    try:

        weekly_df = pd.read_excel(path, sheet_name="Weekly_Progress")

        if not weekly_df.empty:

            weekly_df["Date"] = pd.to_datetime(weekly_df["Date"])

            fig_progress = px.line(
                weekly_df,
                x="Week",
                y=["Open", "Closed"],
                markers=True
            )

            # Make dots bigger and clearer
            fig_progress.update_traces(
                marker=dict(size=10),
                line=dict(width=3)
            )

            fig_progress.update_layout(
                xaxis_title="Week",
                yaxis_title="Clash Count",
                legend_title="Status"
            )

            st.plotly_chart(fig_progress, width="stretch")

        else:
            st.info("No progress data yet.")

    except Exception as e:
        st.warning("Weekly progress data not available yet.")

    # ------------------------------
    # Priority Chart
    # ------------------------------

    if "Priority" in df.columns:

        st.subheader("Priority Distribution")

        prio = df["Priority"].value_counts().reset_index()
        prio.columns = ["Priority", "Count"]

        fig = px.bar(
            prio,
            x="Priority",
            y="Count"
        )

        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ------------------------------
    # Prefix Charts
    # ------------------------------

    prefixes = df["Prefix"].dropna().unique()

    # for prefix in prefixes:

    #     sub = df[df["Prefix"] == prefix]

    #     st.subheader(f"{prefix} CLASHES")

    #     c1, c2 = st.columns(2)

    for prefix in prefixes:

        sub = df[df["Prefix"] == prefix]

        st.subheader(f"{prefix} CLASHES")

        c1, c2, c3 = st.columns(3)

        # -------------------------
        # Status Donut
        # -------------------------

        if status_col:

            status_counts = sub["Status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]

            fig1 = px.pie(
                status_counts,
                names="Status",
                values="Count",
                hole=0.6
            )

            fig1.update_traces(textinfo="value")

            fig1.update_layout(
                annotations=[
                    dict(text=f"TOTAL<br>{len(sub)}", showarrow=False)
                ]
            )

            c1.plotly_chart(fig1, width="stretch")

        # -------------------------
        # Priority Donut
        # -------------------------

        if "Priority" in sub.columns:

            prio_counts = sub["Priority"].value_counts().reset_index()
            prio_counts.columns = ["Priority", "Count"]

            fig2 = px.pie(
                prio_counts,
                names="Priority",
                values="Count",
                hole=0.6
            )
            fig2.update_traces(textinfo="value")

            fig2.update_layout(
                annotations=[
                    dict(text=f"TOTAL<br>{len(sub)}", showarrow=False)
                ]
            )

            c2.plotly_chart(fig2, width="stretch")

        


        if "Open" in sub.columns and "Closed" in sub.columns:

            oc_df = pd.DataFrame({
                "Type": ["Open", "Closed"],
                "Count": [sub["Open"].sum(), sub["Closed"].sum()]
            })

            fig3 = px.pie(
                oc_df,
                names="Type",
                values="Count",
                hole=0.6
            )

            fig3.update_traces(textinfo="value")

            fig3.update_layout(
                annotations=[
                    dict(text=f"TOTAL<br>{len(sub)}", showarrow=False)
                ]
            )

            c3.plotly_chart(fig3, width="stretch")