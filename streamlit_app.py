import streamlit as st
import openpyxl
from docxtpl import DocxTemplate
from io import BytesIO

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="HSA PT Report Generator", layout="centered")
st.title("HSA Proficiency Testing Scheme")
st.subheader("Interim Report Generator")

# ── HELPER: find a labelled value in column B, return column C of same row ───
def find_value_by_label(ws, label):
    """
    Scans column B for a cell matching 'label' (case-insensitive),
    then returns the value in column C of the same row.
    This is robust to row shifts if the Excel layout changes slightly.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and str(cell.value).strip().lower() == label.lower():
                # Return the value one column to the right
                return ws.cell(row=cell.row, column=cell.column + 1).value
    return None


def get_lab_results(ws):
    """
    Finds the lab results table by locating the header row
    containing 'Data ID or Lab Code', then reads rows below it
    until an empty lab code is encountered.
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
        lab_code  = ws.cell(row=r, column=2).value  # col B
        result    = ws.cell(row=r, column=3).value  # col C
        std_u     = ws.cell(row=r, column=4).value  # col D
        z_score   = ws.cell(row=r, column=6).value  # col F
        zeta_score = ws.cell(row=r, column=8).value # col H

        # Stop at first empty lab code row
        if lab_code is None or str(lab_code).strip() == "":
            break

        labs.append({
            "lab_code":   str(int(lab_code)) if isinstance(lab_code, float) else str(lab_code),
            "result":     f"{result:.3f}"    if result    is not None else "—",
            "std_u":      f"{std_u:.3f}"     if std_u     is not None else "Not reported",
            "z_score":    f"{z_score:.1f}"   if z_score   is not None else "—",
            "zeta_score": f"{zeta_score:.1f}" if zeta_score is not None else "Not applicable",
        })

    # Re-sort by lab code ascending for the Word report
    labs.sort(key=lambda x: int(x["lab_code"]))
    return labs


# ── STEP 1: FILE UPLOADS ──────────────────────────────────────────────────────
st.markdown("### Step 1: Upload Files")
excel_file    = st.file_uploader("Upload Excel data file (.xlsx)", type=["xlsx"])
word_template = st.file_uploader("Upload Word template (.docx)",   type=["docx"])

# ── STEP 2: MANUAL INPUTS ─────────────────────────────────────────────────────
st.markdown("### Step 2: Enter Report Details")
project_code      = st.text_input("Project Code",     placeholder="e.g. STY-0200A")
project_subcode_1 = st.text_input("Project Sub-code", placeholder="e.g. STY-0200A-1")
test_analyte_1    = st.text_input("Test Element",     placeholder="e.g. Cadmium")
address           = st.text_area("Address / Letterhead")
coordinator_name  = st.text_input("Coordinator Name", placeholder="e.g. Dr Jane Tan")

# ── STEP 3: GENERATE ──────────────────────────────────────────────────────────
st.markdown("### Step 3: Generate Report")

if st.button("Generate Report"):

    if not excel_file:
        st.error("Please upload the Excel data file.")
    elif not word_template:
        st.error("Please upload the Word template.")
    elif not project_code:
        st.error("Please enter the Project Code.")
    else:
        try:
            # ── READ EXCEL ────────────────────────────────────────────────────
            wb = openpyxl.load_workbook(excel_file, data_only=True)

            # Use PT Scores sheet for summary stats and lab results
            ws_pt = wb["PT Scores"]

            # Use Robust Statistics sheet for arithmetic mean
            ws_rs = wb["Robust Statistics"]

            # Read summary stats by label search — robust to row shifts
            assigned_value     = find_value_by_label(ws_pt, "Assigned Value")
            assigned_u         = find_value_by_label(ws_pt, "Assigned u")
            assigned_U         = find_value_by_label(ws_pt, "Assigned U")
            standard_deviation = find_value_by_label(ws_pt, "Target SD")
            arithmetic_mean    = find_value_by_label(ws_rs, "Mean")

            # Read lab results table
            labs = get_lab_results(ws_pt)

            # ── VALIDATE ──────────────────────────────────────────────────────
            missing = []
            if assigned_value     is None: missing.append("Assigned Value")
            if assigned_u         is None: missing.append("Assigned u")
            if assigned_U         is None: missing.append("Assigned U")
            if standard_deviation is None: missing.append("Target SD")
            if arithmetic_mean    is None: missing.append("Mean (Robust Statistics sheet)")
            if len(labs) == 0:             missing.append("Lab results table")

            if missing:
                st.error(f"Could not find the following values in the Excel file: {', '.join(missing)}. "
                         f"Please check the sheet names and layout.")
            else:
                # ── BUILD CONTEXT ─────────────────────────────────────────────
                context = {
                    "address":            address,
                    "project_code":       project_code,
                    "project_subcode_1":  project_subcode_1,
                    "test_analyte_1":     test_analyte_1,
                    "coordinator_name":   coordinator_name,
                    "assigned_value":     f"{assigned_value:.5f}",
                    "assigned_u":         f"{assigned_u:.5f}",
                    "assigned_U":         f"{assigned_U:.5f}",
                    "arithmetic_mean":    f"{arithmetic_mean:.5f}",
                    "standard_deviation": f"{standard_deviation:.6f}",
                    "lab_rows":           labs,  # for table loop (future use)
                }

                # ── RENDER WORD TEMPLATE ──────────────────────────────────────
                doc = DocxTemplate(word_template)
                doc.render(context)

                output = BytesIO()
                doc.save(output)
                output.seek(0)

                st.success(f"Report generated successfully! "
                           f"Found {len(labs)} labs: "
                           f"{', '.join([l['lab_code'] for l in labs])}")

                st.download_button(
                    label="⬇ Download Report (.docx)",
                    data=output,
                    file_name=f"{project_code}_interim_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except KeyError as e:
            st.error(f"Sheet not found in Excel file: {e}. "
                     f"Expected sheets named 'PT Scores' and 'Robust Statistics'.")
        except Exception as e:
            st.error(f"Something went wrong: {e}")