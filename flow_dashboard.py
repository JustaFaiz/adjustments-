"""
flow_dashboard.py
-----------------
Streamlit dashboard for adjustment flows. Upload one or more Excel files,
pick a flow in the sidebar, and see:

  * summary cards (steps, inputs, calculations, outputs, file links, issues)
  * the interactive flow (Inputs -> Calculations -> Outputs, top to bottom),
    with click-for-details, search and the arrow checkboxes
  * a table of all steps, and the issues list
  * a button to download the standalone HTML (all flows) to send to others

Needs adjustment_flow.py in the SAME folder - it does the reading, arrows,
checks and drawing.

Requires: pip install streamlit openpyxl pandas

Run:
    streamlit run flow_dashboard.py
"""

import io
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import adjustment_flow as af

st.set_page_config(page_title="Adjustment Flow Dashboard", page_icon="🔀", layout="wide")

st.markdown("""
<style>
  .block-container { padding-top: 1.6rem; padding-bottom: 1rem; }
  div[data-testid="stMetric"] { background: rgba(128,128,128,0.08); border-radius: 8px;
                                padding: 10px 14px; border-top: 4px solid #888; }
</style>
""", unsafe_allow_html=True)

BUCKETS = [("Input", "Inputs", "#2ecc71"), ("Calculation", "Calculations", "#2a9d8f"),
           ("Output", "Outputs", "#f39c12")]


class NamedBytes(io.BytesIO):
    """In-memory Excel file that keeps its file name (the loader uses .name)."""
    def __init__(self, data, name):
        super().__init__(data)
        self.name = name


@st.cache_data(show_spinner="Reading tables...")
def parse(files, sheet):
    """files: tuple of (name, bytes). Cached, so switching flows is instant."""
    handles = [NamedBytes(data, name) for name, data in files]
    return af.build_payload(af.load_flows(handles, sheet or None))


@st.cache_data(show_spinner=False)
def page_html(flow, embedded):
    return af.build_html([flow] if embedded else flow, embedded=embedded)


# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Adjustment Flow")
    uploads = st.file_uploader("Excel file(s)", type=["xlsx", "xlsm"], accept_multiple_files=True)
    sheet = st.text_input("Only this sheet (optional)", placeholder="exact tab name")

if not uploads:
    st.title("Adjustment Flow Dashboard")
    st.info("Upload one or more Excel files in the sidebar to begin.")
    st.markdown(
        "Each table needs a header row with **Stage, Step No., Report/Schedule, "
        "System/Location, Action, I/O, Standard**. Several tables per sheet, several "
        "sheets and several files all work - each table becomes one flow.")
    st.stop()

flows = parse(tuple((u.name, u.getvalue()) for u in uploads), sheet.strip())
if not flows:
    st.error("No tables found. Each table needs a header row with at least "
             "Stage, Step No., Report/Schedule and System/Location.")
    st.stop()

with st.sidebar:
    labels = [f"{f['name']}  ({len(f['steps'])} steps" +
              (f", ⚠ {len(f['issues'])})" if f["issues"] else ")") for f in flows]
    idx = st.selectbox("Flow", range(len(flows)), format_func=lambda i: labels[i])
    st.caption(f"{len(flows)} flow(s) found")
    st.divider()
    st.download_button(
        "⬇ Download HTML (all flows)", data=page_html(flows, False),
        file_name=(Path(uploads[0].name).stem if len(uploads) == 1 else "adjustment_flows") + "_flow.html",
        mime="text/html", use_container_width=True,
        help="Standalone file you can send - works offline in Chrome or Edge.")

flow = flows[idx]
steps = flow["steps"]

# ---------------------------------------------------------------- header + cards
st.title(flow["name"])
st.caption(f"{flow['file']} › {flow['sheet']} · header row {flow['headerRow']}")

counts = {b: sum(1 for s in steps if s["bucket"] == b) for b, _, _ in BUCKETS}
cards = [("Steps", len(steps), "#FBCE07")] + \
        [(label, counts[b], col) for b, label, col in BUCKETS] + \
        [("File links", len(flow["dataEdges"]), "#FBCE07"),
         ("Issues", len(flow["issues"]), "#e74c3c" if flow["issues"] else "#2ecc71")]
for col, (label, value, _) in zip(st.columns(len(cards)), cards):
    col.metric(label, value)
# colour each card's top border
st.markdown("<style>" + "".join(
    f'div[data-testid="stHorizontalBlock"] > div:nth-child({i + 1}) div[data-testid="stMetric"]'
    f'{{border-top-color:{c};}}' for i, (_, _, c) in enumerate(cards)) + "</style>",
    unsafe_allow_html=True)

# ---------------------------------------------------------------- tabs
tab_flow, tab_steps, tab_issues = st.tabs(
    ["Flow", "Steps table", f"Issues ({len(flow['issues'])})"])

with tab_flow:
    st.caption("Click a step for its details on the right. Tick **Step-by-step only** "
               "to see just the 1 → 2 → 3 order.")
    components.html(page_html(flow, True), height=780, scrolling=False)

with tab_steps:
    rows = []
    for s in steps:
        row = {"Step No.": s["step"], "Stage": s["stage"], "Report/Schedule": s["report"],
               "System/Location": s["location"], "Action": s["action"], "I/O": s["io"],
               "Standard": s["standard"]}
        row.update(dict(s["extras"]))
        row["Excel row"] = s["row"]
        rows.append(row)
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with tab_issues:
    if not flow["issues"]:
        st.success("No issues found.")
    else:
        by_id = {s["id"]: s for s in steps}
        st.dataframe(pd.DataFrame([{
            "Step": (f"Step {by_id[i['id']]['step'] or '?'} · {by_id[i['id']]['report']}"
                     if i["id"] else "General"),
            "Excel row": by_id[i["id"]]["row"] if i["id"] else "",
            "Issue": i["msg"],
        } for i in flow["issues"]]), use_container_width=True, hide_index=True)
