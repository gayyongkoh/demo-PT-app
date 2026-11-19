import uuid
from io import BytesIO

import openpyxl
import streamlit as st
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

from charts import result_chart, score_chart

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="HSA PT Report Generator", layout="centered")
st.title("HSA Proficiency Testing Scheme")
st.subheader("Interim Report Generator")


# ── HELPERS: reading the Excel workbook ───────────────────────────────────────
def find_value_by_label(ws, label):
    """
    Scans the sheet for a cell whose text equals 'label', then returns the
    value in the cell immediately to its right. Robust to row shifts.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(cell.value).strip() == label:
                return ws.cell(row=cell.row, column=cell.column + 1).value
    return None


def get_lab_results(ws):
    """
    Finds the lab results table by locating the header 'Data ID or Lab Code',
    then reads rows below it until an empty lab code is encountered.
    """
    header_row = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(cell.value).strip() == "Data ID or Lab Code":
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        raise ValueError("Could not find 'Data ID or Lab Code' header in PT Scores sheet.")

    labs = []
    for r in range(header_row + 1, ws.max_row + 1):
        lab_code   = ws.cell(row=r, column=2).value  # col B
        result     = ws.cell(row=r, column=3).value  # col C
        std_u      = ws.cell(row=r, column=4).value  # col D
        z_score    = ws.cell(row=r, column=6).value  # col F
        zeta_score = ws.cell(row=r, column=8).value  # col H

        # Stop at first empty lab code row
        if lab_code is None or str(lab_code).strip() == "":
            break

        labs.append({
            "lab_code":   str(int(lab_code)) if isinstance(lab_code, float) else str(lab_code),
            "result":     f"{result:.3f}"     if result     is not None else "—",
            "std_u":      f"{std_u:.3f}"      if std_u      is not None else "Not reported",
            "z_score":    f"{z_score:.1f}"    if z_score    is not None else "—",
            "zeta_score": f"{zeta_score:.1f}" if zeta_score is not None else "Not applicable",
            "z_raw":      z_score,      # raw numbers for chart plotting
            "zeta_raw":   zeta_score,
        })

    # Table in the Word report is ordered by lab code ascending
    labs.sort(key=lambda x: int(x["lab_code"]))
    return labs


def get_comparison_data(ws):
    """
    Reads the table on the 'Comparison of Results' sheet. Columns are located
    relative to the 'Data ID or Lab Code' header column:
      +0 lab code, +1 assigned value, +2 assigned u, +3 assigned U,
      +4 lab result, +5 lab reported u, +6 lab reported U
    Check these offsets against the real sheet.
    """
    header_row = header_col = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(cell.value).strip() == "Data ID or Lab Code":
                header_row, header_col = cell.row, cell.column
                break
        if header_row:
            break
    if not header_row:
        raise ValueError("Could not find 'Data ID or Lab Code' header in Comparison of Results sheet.")

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        def g(off, r=r):
            return ws.cell(row=r, column=header_col + off).value

        code = g(0)
        if code is None or str(code).strip() == "":
            break
        if not isinstance(g(4), (int, float)):
            continue
        rows.append({
            "lab_code": str(int(code)) if isinstance(code, float) else str(code),
            "av":     g(1),
            "u":      g(2),
            "U":      g(3),
            "result": g(4),
            "lab_u":  g(5),
            "lab_U":  g(6),   # "Lab reported U", used for the error bars
        })
    return rows


# ── PER-ANALYTE PROCESSING ────────────────────────────────────────────────────
def process_analyte(excel_file, doc, name, subcode):
    """Read one analyte's workbook and return everything the template needs."""
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    ws_pt = wb["PT Scores"]
    ws_rs = wb["Robust Statistics"]

    vals = {
        "Assigned Value": find_value_by_label(ws_pt, "Assigned Value"),
        "Assigned u":     find_value_by_label(ws_pt, "Assigned u"),
        "Assigned U":     find_value_by_label(ws_pt, "Assigned U"),
        "Target SD":      find_value_by_label(ws_pt, "Target SD"),
        "Mean":           find_value_by_label(ws_rs, "Mean"),
    }
    labs = get_lab_results(ws_pt)

    missing = [k for k, v in vals.items() if v is None]
    if not labs:
        missing.append("Lab results table")
    if missing:
        raise ValueError(f"could not find {', '.join(missing)}. Please check the sheet names and layout.")

    # ── Score charts: each sorted by its own score, ascending ────────────────
    z_pts = sorted(
        [(l["lab_code"], l["z_raw"]) for l in labs if isinstance(l["z_raw"], (int, float))],
        key=lambda p: p[1])
    zeta_pts = sorted(
        [(l["lab_code"], l["zeta_raw"]) for l in labs if isinstance(l["zeta_raw"], (int, float))],
        key=lambda p: p[1])

    def score_img(pts, label, fallback):
        if not pts:
            return fallback
        codes, scores = zip(*pts)
        return InlineImage(doc, score_chart(codes, scores, label), width=Mm(78))

    # ── Lab results chart (Comparison of Results sheet) ──────────────────────
    cmp_rows = get_comparison_data(wb["Comparison of Results"])
    cmp_rows.sort(key=lambda r: r["result"])   # ascending by result

    if cmp_rows and cmp_rows[0]["av"] is not None and cmp_rows[0]["U"] is not None:
        av, U = cmp_rows[0]["av"], cmp_rows[0]["U"]
        result_img = InlineImage(
            doc,
            result_chart(
                [r["lab_code"] for r in cmp_rows],
                [r["result"] for r in cmp_rows],
                [r["lab_U"] for r in cmp_rows],
                av, av + U, av - U, vals["Mean"],
                f"Mass fraction of {name} (mg/kg)",
            ),
            width=Mm(200),
        )
    else:
        result_img = "No results available"

    return {
        "test_analyte":       name,
        "project_subcode":    subcode,
        "assigned_value":     f"{vals['Assigned Value']:.5f}",
        "assigned_u":         f"{vals['Assigned u']:.5f}",
        "assigned_U":         f"{vals['Assigned U']:.5f}",
        "arithmetic_mean":    f"{vals['Mean']:.5f}",
        "standard_deviation": f"{vals['Target SD']:.6f}",
        "lab_rows":           labs,
        "z_chart":            score_img(z_pts,    "z-Scores", "No z-scores available"),
        "zeta_chart":         score_img(zeta_pts, "ζ-Scores", "No ζ-scores available"),
        "result_chart":       result_img,
    }


# ── STEP 1: REPORT DETAILS ────────────────────────────────────────────────────
st.markdown("### Step 1: Report Details")
word_template    = st.file_uploader("Upload Word template (.docx)", type=["docx"])
project_code     = st.text_input("Project Code", placeholder="e.g. STY-0200A")
address          = st.text_area("Address / Letterhead")
coordinator_name = st.text_input("Coordinator Name", placeholder="e.g. Dr Jane Tan")

# ── STEP 2: TEST ANALYTES (add / remove) ──────────────────────────────────────
st.markdown("### Step 2: Test Analytes")

if "analyte_ids" not in st.session_state:
    st.session_state.analyte_ids = [uuid.uuid4().hex]

analyte_inputs = []
for n, aid in enumerate(st.session_state.analyte_ids, start=1):
    typed_name = st.session_state.get(f"name_{aid}", "")
    title = f"Analyte {n}: {typed_name}" if typed_name else f"Analyte {n}"
    with st.expander(title, expanded=True):
        analyte_inputs.append({
            "file":    st.file_uploader("Excel data file (.xlsx)", type=["xlsx"], key=f"xlsx_{aid}"),
            "name":    st.text_input("Test Element", placeholder="e.g. Cadmium", key=f"name_{aid}"),
            "subcode": st.text_input("Project Sub-code", placeholder="e.g. STY-0200A-1", key=f"sub_{aid}"),
        })
        if len(st.session_state.analyte_ids) > 1:
            if st.button("Remove this analyte", key=f"rm_{aid}"):
                st.session_state.analyte_ids.remove(aid)
                st.rerun()

if st.button("➕ Add another analyte"):
    st.session_state.analyte_ids.append(uuid.uuid4().hex)
    st.rerun()

# ── STEP 3: GENERATE ──────────────────────────────────────────────────────────
st.markdown("### Step 3: Generate Report")
st.caption(f"{len(analyte_inputs)} analyte(s) in this report")

if st.button("Generate Report"):
    st.session_state.pop("report", None)   # clear any previous result

    if not word_template:
        st.error("Please upload the Word template.")
    elif not project_code:
        st.error("Please enter the Project Code.")
    elif any(a["file"] is None or not a["name"] for a in analyte_inputs):
        st.error("Every analyte needs an Excel file and a Test Element name.")
    else:
        doc = DocxTemplate(word_template)   # one doc object shared by all InlineImages
        analytes, errors = [], []

        for i, a in enumerate(analyte_inputs, start=1):
            try:
                analytes.append(process_analyte(a["file"], doc, a["name"], a["subcode"]))
            except KeyError as e:
                errors.append(
                    f"Analyte {i} ({a['name']}): sheet not found: {e}. Expected sheets named "
                    f"'PT Scores', 'Robust Statistics' and 'Comparison of Results'.")
            except Exception as e:
                errors.append(f"Analyte {i} ({a['name']}): {e}")

        if errors:
            for msg in errors:
                st.error(msg)
        else:
            try:
                doc.render({
                    "address":          address,
                    "project_code":     project_code,
                    "coordinator_name": coordinator_name,
                    "analytes":         analytes,
                })
                output = BytesIO()
                doc.save(output)

                st.session_state["report"] = {
                    "data":    output.getvalue(),
                    "name":    f"{project_code}_interim_report.docx",
                    "summary": [
                        f"✔ {a['test_analyte']}: {len(a['lab_rows'])} labs "
                        f"({', '.join(l['lab_code'] for l in a['lab_rows'])})"
                        for a in analytes
                    ],
                }
            except Exception as e:
                st.error(f"Something went wrong while building the report: {e}")

# ── DOWNLOAD (outside the button block so it survives reruns) ────────────────
if "report" in st.session_state:
    rep = st.session_state["report"]
    st.success("Report generated successfully!")
    for line in rep["summary"]:
        st.write(line)
    st.download_button(
        label="⬇ Download Report (.docx)",
        data=rep["data"],
        file_name=rep["name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )