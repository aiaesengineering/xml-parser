import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import re
import io
import os

st.set_page_config(page_title="Clash Exporter", page_icon="🏗️", layout="centered")

st.markdown("""
<style>
    .stApp { background: #f8f9fa; }
    .block-container { max-width: 760px; padding-top: 2rem; }
    h1 { font-size: 1.6rem !important; font-weight: 700 !important; color: #1a1a1a !important; }
    .sub { color: #666; font-size: 0.9rem; margin-top: -0.5rem; margin-bottom: 1.5rem; }
    .metric { background: #f0f4ff; border-radius: 8px; padding: 12px; text-align: center; }
    .metric-val { font-size: 1.8rem; font-weight: 700; color: #2563eb; }
    .metric-label { font-size: 0.72rem; color: #666; text-transform: uppercase; letter-spacing: 0.5px; }
    .stButton > button { background: #2563eb; color: white; border: none; border-radius: 6px; padding: 0.55rem 1.5rem; font-weight: 600; width: 100%; }
    .stButton > button:hover { background: #1d4ed8; }
    .stDownloadButton > button { background: #16a34a; color: white; border: none; border-radius: 6px; padding: 0.55rem 1.5rem; font-weight: 600; width: 100%; }
    .stDownloadButton > button:hover { background: #15803d; }
</style>
""", unsafe_allow_html=True)


# ── Prefix helpers ──────────────────────────────────────────────────────────

def extract_prefix_regex(test_name):
    """Find first discipline code like AR-, EL-, FP- after start or underscore."""
    if not test_name:
        return None
    matches = re.findall(r'(?:^|_)([A-Z]{2,4})-', test_name)
    return matches[0] if matches else None


def extract_prefix_position(test_name, position):
    if not test_name:
        return None
    parts = test_name.split("_")
    if len(parts) >= position:
        return parts[position - 1].split("-")[0].strip()
    return None


# ── Parser ──────────────────────────────────────────────────────────────────
# Navisworks XML comes in two flavours:
#   - standard:   <clashresult ...>   (singular)
#   - some exports: <clashgroup ...>  (what this user's file uses)
# We detect which tag is present per clashtest and use that.

def parse_xml(file_bytes, use_manual, position):
    root = ET.parse(io.BytesIO(file_bytes)).getroot()
    rows = []

    for test in root.iter("clashtest"):
        test_name     = test.get("name")
        test_priority = test.get("priority")

        prefix = extract_prefix_position(test_name, position) if use_manual else extract_prefix_regex(test_name)

        # Support both tag names
        clash_list = list(test.iter("clashresult"))
        if not clash_list:
            clash_list = list(test.iter("clashgroup"))

        total_clashes = len(clash_list)

        for clash in clash_list:
            row = {
                "Test Name":             test_name,
                "Prefix":                prefix,
                "Test Priority":         test_priority,
                "Total Clashes In Test": total_clashes,
            }

            for attr, val in clash.attrib.items():
                row[f"clash_{attr}"] = val

            for elem in clash.iter():
                tag = elem.tag
                if elem.text and elem.text.strip():
                    row[tag] = elem.text.strip()
                for attr, val in elem.attrib.items():
                    row[f"{tag}_{attr}"] = val

            for attr_elem in clash.findall(".//objectattribute"):
                name  = attr_elem.findtext("name")
                value = attr_elem.findtext("value")
                if name and value:
                    row[name] = value

            rows.append(row)

    return rows


# ── Excel builder ───────────────────────────────────────────────────────────

def build_excel(df):
    output     = io.BytesIO()
    status_col = next((c for c in df.columns if "status" in c.lower()), None)

    if status_col:
        lower_s      = df[status_col].astype(str).str.lower()
        df["Open"]   = lower_s.isin(["new", "active", "reviewed"]).astype(int)
        df["Closed"] = lower_s.isin(["approved", "resolved"]).astype(int)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Clash_Details", index=False)
        if status_col:
            summary = df[status_col].astype(str).str.lower().value_counts().reset_index()
            summary.columns = ["clash_status", "count"]
            summary.to_excel(writer, sheet_name="Status_Summary", index=False)
        if "Prefix" in df.columns:
            ps = df.groupby("Prefix").size().reset_index(name="Clash Count")
            ps.to_excel(writer, sheet_name="Prefix_Summary", index=False)

    return output.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────

st.markdown("# 🏗️ Navisworks Clash Exporter")
st.markdown('<p class="sub">Upload a clash XML → configure prefix → generate Excel</p>', unsafe_allow_html=True)

for k in ("df", "excel_bytes", "cached_bytes", "last_file"):
    if k not in st.session_state:
        st.session_state[k] = None

# Step 1
st.markdown("**1. Upload XML file**")
uploaded_file = st.file_uploader("xml", type=["xml"], label_visibility="collapsed")

if uploaded_file:
    if st.session_state.last_file != uploaded_file.name:
        st.session_state.cached_bytes = uploaded_file.read()
        st.session_state.last_file    = uploaded_file.name
        st.session_state.df           = None
        st.session_state.excel_bytes  = None

    file_bytes   = st.session_state.cached_bytes
    file_size_mb = len(file_bytes) / (1024 * 1024)
    st.caption(f"📄 {uploaded_file.name}  ·  {file_size_mb:.1f} MB")
    st.divider()

    # Step 2
    st.markdown("**2. Prefix detection**")
    mode = st.radio(
        "mode",
        ["Auto (Regex) — detects AR-, EL-, FP- etc.", "Manual (Position)"],
        label_visibility="collapsed",
    )

    use_manual   = "Manual" in mode
    position_val = None
    if use_manual:
        position_val = st.number_input(
            "Underscore position (e.g. 2 → 2nd segment of test name split by _)",
            min_value=1, max_value=20, value=2, step=1,
        )
    st.divider()

    # Step 3
    st.markdown("**3. Generate report**")
    if st.button("⚙️ Generate Excel Report"):
        with st.spinner("Parsing XML…"):
            try:
                pos  = int(position_val) if use_manual and position_val else None
                rows = parse_xml(file_bytes, use_manual=use_manual, position=pos)

                if not rows:
                    st.error("No clash elements found. The XML has clashtest nodes but no clashresult or clashgroup children.")
                else:
                    df = pd.DataFrame(rows)
                    st.session_state.df          = df
                    st.session_state.excel_bytes = build_excel(df.copy())
                    st.success(f"✅ {len(df):,} clashes across {df['Test Name'].nunique():,} tests.")

            except ET.ParseError as e:
                st.error(f"XML parse error: {e}")
            except Exception as e:
                st.error(f"Error: {e}")

    # Results
    if st.session_state.df is not None:
        df = st.session_state.df
        st.divider()

        cols = st.columns(4)
        for col, (label, val) in zip(cols, [
            ("Total Clashes",  f"{len(df):,}"),
            ("Clash Tests",    f"{df['Test Name'].nunique():,}"),
            ("Prefixes Found", f"{df['Prefix'].nunique():,}" if "Prefix" in df.columns else "—"),
            ("Columns",        f"{len(df.columns):,}"),
        ]):
            with col:
                st.markdown(
                    f'<div class="metric"><div class="metric-val">{val}</div>'
                    f'<div class="metric-label">{label}</div></div>',
                    unsafe_allow_html=True,
                )

        st.markdown("")

        if "Prefix" in df.columns:
            st.markdown("**Prefix breakdown**")
            prefix_counts = df["Prefix"].value_counts().reset_index()
            prefix_counts.columns = ["Prefix", "Clash Count"]
            c1, c2 = st.columns([1, 2])
            with c1:
                st.dataframe(prefix_counts, width="stretch", hide_index=True)
            with c2:
                st.bar_chart(prefix_counts.set_index("Prefix"))

        with st.expander("Preview data (first 100 rows)"):
            st.dataframe(df.head(100), width="stretch", hide_index=True)

        st.divider()
        out_name = os.path.splitext(uploaded_file.name)[0] + "_clash_report.xlsx"
        st.download_button(
            "⬇️ Download Excel Report",
            data=st.session_state.excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Upload a Navisworks clash XML file above to get started.")