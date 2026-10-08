"""
flow_dashboard.py
-----------------
Streamlit dashboard for adjustment flows. Upload one or more Excel files,
pick a flow in the sidebar, and the page shows, top to bottom:

  1. summary cards (steps, inputs, calculations, outputs, file links, issues)
  2. a simple, static flow - Inputs -> Process -> Output - Internal ->
     Output - External as light coloured bands, steps as cards, one arrow
     from each step to the next
  3. the interactive flow (click-for-details, search, arrow checkboxes)
  4. a table of all steps, and the issues list
Plus sidebar buttons to download, as standalone HTML for all adjustments,
either the Flow overview (light, with a dropdown) or the interactive view.

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

BUCKETS = [("Input", "Inputs", "#2ecc71"), ("Calculation", "Process", "#2a9d8f"),
           ("OutputInternal", "Output - Internal", "#f39c12"),
           ("OutputExternal", "Output - External", "#e76f51")]
OPTIONAL_BUCKETS = [("Output", "Outputs (not split)", "#f4a261"),
                    ("Unassigned", "Unknown stage", "#9ca3af")]
# light styling for the static flow: (band tint, band accent)
BAND_STYLE = {"Input": ("#effaf3", "#2ecc71"), "Calculation": ("#ecf7f6", "#2a9d8f"),
              "OutputInternal": ("#fef6ea", "#f39c12"), "OutputExternal": ("#fdf0ec", "#e76f51"),
              "Output": ("#fef4ec", "#f4a261"), "Unassigned": ("#f3f4f6", "#9ca3af")}
BAND_NAME = {"Input": "Inputs", "Calculation": "Process", "OutputInternal": "Output - Internal",
             "OutputExternal": "Output - External", "Output": "Outputs (not split)",
             "Unassigned": "Unknown stage"}
TYPE_BADGE = {"Temporary": ("#e8f1fb", "#1f6fb2"), "Permanent": ("#f3ebfa", "#7b3fa8")}


def fmt_amount(v):
    return "-" if v is None else f"{v:,.2f}"
ACTION_STYLE = [("download", "#e8f1fb", "#1f6fb2"), ("upload", "#f3ebfa", "#7b3fa8"),
                ("roll", "#fdf5d9", "#8a6d00")]

def esc(v):
    return (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def clip(text, n):
    text = str(text)
    return text if len(text) <= n else text[:n - 1] + "…"


def wrap2(text, n):
    """Split into at most two lines of ~n chars (second line clipped)."""
    words, lines, cur = str(text).split(), [], ""
    for w in words:
        if len(cur) + len(w) + (1 if cur else 0) <= n:
            cur = f"{cur} {w}".strip()
        else:
            lines.append(cur or w[:n])
            cur = w if cur else ""
    lines.append(cur)
    lines = [l for l in lines if l] or [""]
    return lines if len(lines) <= 2 else [lines[0], clip(" ".join(lines[1:]), n)]


# layout of the static flow (pixels)
CARD_W, CARD_H, GAP_X, ROW_GAP, BAND_GAP = 200, 104, 48, 34, 46
PER_ROW, PAD_X, BAND_TOP, BAND_BOTTOM = 4, 24, 42, 18
FLOW_W = 2 * PAD_X + PER_ROW * CARD_W + (PER_ROW - 1) * GAP_X
FONT = "Source Sans Pro, Segoe UI, Arial, sans-serif"


def static_flow_svg(flow, uid="ah"):
    """Static flow: light bands top to bottom, one arrow from each step to the next."""
    issue_count = {}
    for i in flow["issues"]:
        if i["id"]:
            issue_count[i["id"]] = issue_count.get(i["id"], 0) + 1
    always = ["Input", "Calculation", "OutputInternal", "OutputExternal"]
    order = [b for b in af.BUCKET_ORDER
             if b in always or any(s["bucket"] == b for s in flow["steps"])]

    bands, pos, y = [], {}, 0
    for b in order:
        in_band = [s for s in flow["steps"] if s["bucket"] == b]
        rows = [in_band[i:i + PER_ROW] for i in range(0, len(in_band), PER_ROW)] or [[]]
        top = y
        cy = top + BAND_TOP
        for row in rows:
            row_w = len(row) * CARD_W + max(0, len(row) - 1) * GAP_X
            x = (FLOW_W - row_w) / 2
            for s in row:
                pos[s["id"]] = (x, cy)
                x += CARD_W + GAP_X
            cy += CARD_H + ROW_GAP
        height = BAND_TOP + len(rows) * CARD_H + (len(rows) - 1) * ROW_GAP + BAND_BOTTOM
        bands.append((b, top, height, len(in_band)))
        y = top + height + BAND_GAP
    total_h = y - BAND_GAP + 4

    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {FLOW_W} {total_h}" '
           f'width="100%" style="max-width:{FLOW_W}px;display:block;font-family:{FONT}">',
           f'<defs><marker id="{uid}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" '
           'markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" '
           'fill="#4b5563"/></marker></defs>']

    for b, top, h, n in bands:
        tint, accent = BAND_STYLE[b]
        out.append(f'<rect x="1" y="{top}" width="{FLOW_W - 2}" height="{h}" rx="12" fill="{tint}"/>')
        out.append(f'<rect x="1" y="{top}" width="6" height="{h}" rx="3" fill="{accent}"/>')
        # title on the right, so arrows coming down into the first card never cross it
        out.append(f'<text x="{FLOW_W - PAD_X}" y="{top + 27}" text-anchor="end" font-size="14" '
                   f'font-weight="700" fill="#374151" letter-spacing="0.6">{BAND_NAME[b].upper()}'
                   f'<tspan font-weight="400" fill="#6b7280" letter-spacing="0">  ·  {n} '
                   f'step{"s" if n != 1 else ""}</tspan></text>')

    # arrows: every step to the next one, drawn under the cards
    steps = flow["steps"]
    for prev, nxt in zip(steps, steps[1:]):
        (x1, y1), (x2, y2) = pos[prev["id"]], pos[nxt["id"]]
        if y1 == y2:   # same row: straight across
            a, b_ = (x1 + CARD_W, x2) if x2 > x1 else (x1, x2 + CARD_W)
            d = f"M{a},{y1 + CARD_H / 2} L{b_ - 2},{y2 + CARD_H / 2}"
        else:          # next row / next band: down, across, down (rounded corners)
            sx, sy, ex, ey = x1 + CARD_W / 2, y1 + CARD_H, x2 + CARD_W / 2, y2 - 2
            my = (sy + ey) / 2
            r = min(10, abs(ex - sx) / 2)
            if r < 1:
                d = f"M{sx},{sy} L{ex},{ey}"
            else:
                dx = 1 if ex > sx else -1
                d = (f"M{sx},{sy} L{sx},{my - r} Q{sx},{my} {sx + dx * r},{my} "
                     f"L{ex - dx * r},{my} Q{ex},{my} {ex},{my + r} L{ex},{ey}")
        out.append(f'<path d="{d}" fill="none" stroke="#4b5563" stroke-width="2" '
                   f'stroke-linejoin="round" marker-end="url(#{uid})"/>')

    for s in steps:
        x, y0 = pos[s["id"]]
        a = (s["action"] or "").lower()
        bg, fg = next(((bb, f) for k, bb, f in ACTION_STYLE if k in a), ("#f3f4f6", "#4b5563"))
        act = clip(s["action"] or "-", 14)
        pw = 8 + 6.2 * len(act)
        tip = "\n".join(f"{k}: {v}" for k, v in s["fields"] if v)
        out.append(f'<g><title>{esc(tip)}</title>')
        out.append(f'<rect x="{x}" y="{y0}" width="{CARD_W}" height="{CARD_H}" rx="10" '
                   f'fill="#ffffff" stroke="#e5e7eb"/>')
        out.append(f'<text x="{x + 12}" y="{y0 + 21}" font-size="11" font-weight="700" '
                   f'fill="#6b7280" letter-spacing="0.5">STEP {esc(s["step"] or "?")}</text>')
        out.append(f'<rect x="{x + CARD_W - 10 - pw}" y="{y0 + 9}" width="{pw}" height="17" '
                   f'rx="8.5" fill="{bg}"/>')
        out.append(f'<text x="{x + CARD_W - 10 - pw / 2}" y="{y0 + 21.5}" font-size="10.5" '
                   f'font-weight="600" fill="{fg}" text-anchor="middle">{esc(act)}</text>')
        for k, line in enumerate(wrap2(s["report"] or "(no report name)", 25)):
            out.append(f'<text x="{x + 12}" y="{y0 + 44 + k * 16}" font-size="13.5" '
                       f'font-weight="600" fill="#111827">{esc(line)}</text>')
        wp = s["location"] or s["source"] or "no work paper"
        if s["location"] and s["wpSheet"]:
            wp += " › " + s["wpSheet"]
        out.append(f'<text x="{x + 12}" y="{y0 + 82}" font-size="11" fill="#6b7280">'
                   f'📄 {esc(clip(wp, 30))}</text>')
        warn = issue_count.get(s["id"])
        if warn:
            out.append(f'<text x="{x + 12}" y="{y0 + 97}" font-size="10.5" font-weight="700" '
                       f'fill="#c0392b">⚠ {warn} issue{"s" if warn > 1 else ""}</text>')
        out.append('</g>')
    out.append('</svg>')
    return "".join(out), total_h


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


OVERVIEW_PAGE = """<!doctype html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Adjustment Flow Overview</title>
<style>
  body { margin:0; background:#fff; color:#111827; font-family:Source Sans Pro, Segoe UI, Arial, sans-serif; }
  header { position:sticky; top:0; z-index:5; background:#f9fafb; border-bottom:1px solid #e5e7eb;
           padding:12px 24px; display:flex; align-items:center; gap:12px; flex-wrap:wrap; }
  header h1 { font-size:18px; margin:0 12px 0 0; }
  select { font-size:15px; padding:7px 10px; min-width:320px; border:1px solid #d1d5db; border-radius:6px;
           background:#fff; cursor:pointer; }
  .nav { cursor:pointer; font-size:22px; color:#6b7280; padding:0 6px; user-select:none; }
  .nav:hover { color:#111827; }
  #pos { color:#6b7280; font-size:13px; }
  main { max-width:1000px; margin:0 auto; padding:22px 24px 40px; }
  .flow { display:none; }
  .flow.on { display:block; }
  h2 { font-size:30px; margin:4px 0 8px; }
  .meta { display:flex; gap:18px; align-items:center; flex-wrap:wrap; margin-bottom:6px; }
  .badge { padding:3px 12px; border-radius:12px; font-weight:600; font-size:14px; }
  .amt { font-size:15px; color:#374151; }
  .amt b { font-size:17px; }
  .src { color:#6b7280; font-size:13px; margin-bottom:12px; }
  .counts { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:18px; }
  .count { background:#f3f4f6; border-radius:8px; padding:6px 12px; font-size:13px; color:#374151;
           border-top:3px solid #9ca3af; }
  .count b { font-size:16px; margin-right:4px; }
  .hint { color:#6b7280; font-size:13px; margin:0 0 10px; }
  @media print { header { position:static; } .nav, #pos { display:none; } }
</style></head><body>
<header>
  <h1>Adjustment Flow Overview</h1>
  <span class="nav" id="prev" title="Previous">&lsaquo;</span>
  <select id="pick">__OPTIONS__</select>
  <span class="nav" id="next" title="Next">&rsaquo;</span>
  <span id="pos"></span>
</header>
<main>__FLOWS__</main>
<script>
  var pick = document.getElementById("pick"), flows = document.querySelectorAll(".flow");
  function show(i) {
    i = (i + flows.length) % flows.length;
    flows.forEach(function (f, k) { f.classList.toggle("on", k === i); });
    pick.value = String(i);
    document.getElementById("pos").textContent = (i + 1) + " of " + flows.length;
    try { history.replaceState(null, "", "#adj=" + i); } catch (e) {}
  }
  pick.addEventListener("change", function () { show(+pick.value); });
  document.getElementById("prev").addEventListener("click", function () { show(+pick.value - 1); });
  document.getElementById("next").addEventListener("click", function () { show(+pick.value + 1); });
  var m = /adj=(\\d+)/.exec(location.hash);
  show(m && +m[1] < flows.length ? +m[1] : 0);
</script>
</body></html>"""


@st.cache_data(show_spinner=False)
def overview_html(flows):
    """Standalone light page: the Flow overview for every adjustment + a dropdown."""
    options, sections = [], []
    for i, f in enumerate(flows):
        options.append(f'<option value="{i}">{esc(f["name"])}'
                       + (f' ({esc(f["adjType"])})' if f["adjType"] else "") + "</option>")
        meta = []
        if f["adjType"]:
            bg, fg = TYPE_BADGE.get(f["adjType"], ("#f3f4f6", "#4b5563"))
            meta.append(f'<span class="badge" style="background:{bg};color:{fg}">'
                        f'{esc(f["adjType"])} adjustment</span>')
        if f["amount"] is not None:
            meta.append(f'<span class="amt">{esc(f["amountLabel"])}: <b>{fmt_amount(f["amount"])}</b></span>')
        counts = {b: sum(1 for s in f["steps"] if s["bucket"] == b) for b in af.BUCKET_ORDER}
        shown = BUCKETS + [bk for bk in OPTIONAL_BUCKETS if counts[bk[0]]]
        chips = [f'<span class="count" style="border-top-color:#FBCE07"><b>{len(f["steps"])}</b>Steps</span>'] + \
                [f'<span class="count" style="border-top-color:{c}"><b>{counts[b]}</b>{esc(lbl)}</span>'
                 for b, lbl, c in shown]
        svg, _ = static_flow_svg(f, uid=f"ah{i}")
        sections.append(
            f'<section class="flow"><h2>{esc(f["name"])}</h2>'
            + (f'<div class="meta">{"".join(meta)}</div>' if meta else "")
            + f'<div class="src">{esc(f["file"])} › {esc(f["sheet"])} · from row {f["headerRow"]}</div>'
            + f'<div class="counts">{"".join(chips)}</div>'
            + '<p class="hint">Each arrow goes to the next step. The pill shows the type; '
              '📄 is the work paper. Hover a card for its full details.</p>'
            + f'{svg}</section>')
    return (OVERVIEW_PAGE.replace("__OPTIONS__", "".join(options))
            .replace("__FLOWS__", "".join(sections)))


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
        "Each adjustment starts with a details row (**Temp/Perm, name, Return (CY) + amount**), "
        "followed by up to four tables - **Input, Process, Output - Internal, Output - External** - "
        "each with the columns **Sequence, Name, Type, Source, Classification Basis, "
        "Work Paper Name, Work Paper Sheet** (outputs use **Process** and **Format / Basis**).")
    st.stop()

flows = parse(tuple((u.name, u.getvalue()) for u in uploads), sheet.strip())
if not flows:
    st.error("No adjustments found. Each table needs a header row with at least "
             "Sequence and two other known columns.")
    st.stop()

with st.sidebar:
    labels = [f"{f['name']}" + (f" · {f['adjType'][:4]}" if f["adjType"] else "") +
              f"  ({len(f['steps'])} steps" +
              (f", ⚠ {len(f['issues'])})" if f["issues"] else ")") for f in flows]
    idx = st.selectbox("Flow", range(len(flows)), format_func=lambda i: labels[i])
    st.caption(f"{len(flows)} flow(s) found")
    st.divider()
    base_name = Path(uploads[0].name).stem if len(uploads) == 1 else "adjustment_flows"
    st.download_button(
        "⬇ Download Flow overview (HTML)", data=overview_html(flows),
        file_name=base_name + "_overview.html", mime="text/html", use_container_width=True,
        help="The light flow overview for every adjustment, with a dropdown. "
             "Standalone - works offline in Chrome or Edge.")
    st.download_button(
        "⬇ Download Interactive view (HTML)", data=page_html(flows, False),
        file_name=base_name + "_flow.html",
        mime="text/html", use_container_width=True,
        help="The dark interactive view for every adjustment. Standalone - works offline in Chrome or Edge.")

flow = flows[idx]
steps = flow["steps"]

# ---------------------------------------------------------------- header + cards
st.title(flow["name"])
badges = []
if flow["adjType"]:
    bg, fg = TYPE_BADGE.get(flow["adjType"], ("#f3f4f6", "#4b5563"))
    badges.append(f'<span style="background:{bg};color:{fg};padding:3px 12px;border-radius:12px;'
                  f'font-weight:600;font-size:0.9rem">{esc(flow["adjType"])} adjustment</span>')
if flow["amount"] is not None:
    badges.append(f'<span style="font-size:0.95rem;color:#374151">{esc(flow["amountLabel"])}: '
                  f'<b style="font-size:1.1rem">{fmt_amount(flow["amount"])}</b></span>')
if badges:
    st.markdown('<div style="display:flex;gap:18px;align-items:center;margin:-6px 0 6px">'
                + "".join(badges) + "</div>", unsafe_allow_html=True)
st.caption(f"{flow['file']} › {flow['sheet']} · from row {flow['headerRow']}")

counts = {b: sum(1 for s in steps if s["bucket"] == b) for b in af.BUCKET_ORDER}
shown = BUCKETS + [bk for bk in OPTIONAL_BUCKETS if counts[bk[0]]]
cards = [("Steps", len(steps), "#FBCE07")] + \
        [(label, counts[b], col) for b, label, col in shown] + \
        [("File links", len(flow["dataEdges"]), "#FBCE07"),
         ("Issues", len(flow["issues"]), "#e74c3c" if flow["issues"] else "#2ecc71")]
for col, (label, value, _) in zip(st.columns(len(cards)), cards):
    col.metric(label, value)
# colour each card's top border
st.markdown("<style>" + "".join(
    f'div[data-testid="stHorizontalBlock"] > div:nth-child({i + 1}) div[data-testid="stMetric"]'
    f'{{border-top-color:{c};}}' for i, (_, _, c) in enumerate(cards)) + "</style>",
    unsafe_allow_html=True)

# ---------------------------------------------------------------- 1. static flow
st.subheader("Flow overview")
st.caption("Steps in order, grouped by stage - each arrow goes to the next step. "
           "The pill shows the type; 📄 is the file used. Hover a card for its full details.")
svg, svg_h = static_flow_svg(flow)
components.html(f'<div style="display:flex;justify-content:center">{svg}</div>',
                height=int(svg_h) + 12, scrolling=False)

# ---------------------------------------------------------------- 2. interactive flow
st.divider()
st.subheader("Interactive view")
st.caption("Click a step for its details on the right. Tick **Step-by-step only** "
           "to see just the 1 → 2 → 3 order.")
components.html(page_html(flow, True), height=780, scrolling=False)

# ---------------------------------------------------------------- 3. tables
st.divider()
tab_steps, tab_issues = st.tabs(["Steps table", f"Issues ({len(flow['issues'])})"])

with tab_steps:
    rows = []
    for s in steps:
        # one set of columns for every stage (outputs call Name "Process"
        # and Classification Basis "Format / Basis")
        row = {"Stage": s["stage"], "Sequence": f"Step {s['step']}" if s["step"] else "",
               "Name / Process": s["report"], "Type": s["action"], "Source": s["source"],
               "Classification / Format Basis": s["standard"],
               "Work Paper Name": s["location"], "Work Paper Sheet": s["wpSheet"]}
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
