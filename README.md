# HSA Proficiency Testing Scheme: Interim Report Generator

A Streamlit app that builds a Word interim report for a proficiency testing (PT) round. Staff upload one Excel results file per test analyte plus the Word template. The app extracts the values, draws the charts and fills the template. The output is one `.docx` report covering all analytes.

## What it produces

For every test analyte, the report contains:

| Page | Orientation | Content |
|---|---|---|
| 1 (once per report) | Portrait | Cover page: letterhead, reference number, evaluation notes, coordinator |
| Per analyte, page A | Portrait | Table of laboratory results, plus the z-score and ζ-score bar charts side by side |
| Per analyte, page B | Landscape | Assigned value and uncertainty table, plus the lab results chart |

The analyte pages repeat once per analyte, in the order the analytes were added in the app.

## Project files

| File | Purpose |
|---|---|
| `streamlit_app.py` | The app: UI, Excel reading, per-analyte processing, report rendering |
| `charts.py` | matplotlib charts: `score_chart` (z and ζ bar charts) and `result_chart` (lab results with assigned value band and mean) |
| `pyproject.toml` | Dependencies (see below). If a `uv.lock` exists, regenerate it after any change |
| Word template (`.docx`) | Uploaded by staff each time. Not stored in the repo, but keep the current master copy under version control |

## Dependencies

Listed in `pyproject.toml`:

```toml
dependencies = [
    "streamlit",
    "openpyxl",
    "docxtpl",
    "matplotlib>=3.7",
]
```

`matplotlib>=3.7` is needed because the chart legend uses features from recent versions. After editing dependencies, run `uv lock` (if the repo has `uv.lock`), commit both files, and reboot the app on Streamlit Cloud.

## Running locally

```
uv sync            # or: pip install streamlit openpyxl docxtpl matplotlib
streamlit run streamlit_app.py
```

## Using the app

1. **Step 1: Report details.** Upload the Word template and enter the Project Code, Address / Letterhead and Coordinator Name.
2. **Step 2: Test analytes.** Each analyte has its own block:
   - the Excel data file (`.xlsx`)
   - the Test Element name (for example Cadmium)
   - the Project Sub-code (for example STY-0200A-1)

   Use **➕ Add another analyte** for more, and **Remove this analyte** to delete a block. The block title shows the element name once typed.
3. **Step 3: Generate Report.** The app processes every analyte. If any analyte has a problem, the app lists each error by analyte number and name, and no report is built. If all succeed, a summary line per analyte shows the labs found, and a **Download Report (.docx)** button appears. The button stays available after clicking, so staff can download the file again without regenerating.

Check the summary lines before sending a report. They confirm that the right Excel file went with the right analyte name.

## Excel file requirements

Every analyte's file must use identical tab names and layout. The values differ by analyte.

| Tab | What the app reads |
|---|---|
| `PT Scores` | Cells labelled `Assigned Value`, `Assigned u`, `Assigned U` and `Target SD` (the value is read from the cell to the right of each label). The lab table starts below the header `Data ID or Lab Code`. |
| `Robust Statistics` | The cell labelled `Mean`, with the value in the cell to its right. This is the arithmetic mean of reported results. |
| `Comparison of Results` | The table below the header `Data ID or Lab Code`, used for the lab results chart. |

**Lab table in `PT Scores`.** Columns are read by position:

| Column | Content |
|---|---|
| B | Lab code |
| C | Reported result |
| D | Reported standard uncertainty |
| F | z-score |
| H | ζ-score |

Reading stops at the first empty lab code. Labels are matched by exact text (after trimming spaces), so they must not be renamed.

**Comparison of Results table.** Columns are read relative to the lab-code column:

| Offset | Content |
|---|---|
| +0 | Lab code |
| +1 | Assigned value |
| +2 | Assigned u |
| +3 | Assigned U |
| +4 | Lab result |
| +5 | Lab reported u |
| +6 | Lab reported U (used for error bars) |

These offsets were set from screenshots of the sheet. Verify them against the real file when the layout is first used or changed. If a marker or error bar looks wrong in the chart, check this mapping first.

## How values and charts are handled

- **Table order.** Lab rows in the Word table are sorted by lab code, ascending.
- **z-score chart.** Bars are sorted by z-score, ascending. Labs without a z-score are left out.
- **ζ-score chart.** Bars are sorted by ζ-score, ascending. Labs without a ζ-score are left out. Both score charts have the same size, y-axis from -4 to 4, grid lines every 1, and a dashed zero line. A score outside ±4 is drawn as a bar clipped at the axis limit, with its true value in the label.
- **Lab results chart.** Labs are sorted by result, ascending. Each lab is a square marker. Error bars use the lab's reported U, and labs with no reported U get a marker without a bar. The thick black line is the assigned value. The shaded band is the assigned value ± U. The red line is the arithmetic mean. The y-axis range and tick decimals follow the data, with no fixed limits.
- **Missing data.** If an analyte has no scores or no comparison data, the template shows a short text message in place of the chart.
- **Number formats.** Result and std. uncertainty to 3 decimals, scores to 1 decimal, assigned values and mean to 5 decimals, and standard deviation to 6 decimals.

## Word template

The template is rendered with [docxtpl](https://docxtpl.readthedocs.io) (Jinja2 syntax inside Word).

**Report-level fields (outside the loop):**

| Placeholder | Source |
|---|---|
| `{{ address }}` | Address / Letterhead |
| `{{ project_code }}` | Project Code |
| `{{ coordinator_name }}` | Coordinator Name |

**Analyte loop.** Everything for one analyte sits between a `{%p for a in analytes %}` paragraph and a `{%p endfor %}` paragraph. Inside the loop:

| Placeholder | Content |
|---|---|
| `{{ a.test_analyte }}` | Test Element name |
| `{{ a.project_subcode }}` | Project Sub-code |
| `{{ a.assigned_value }}` | Assigned value |
| `{{ a.assigned_u }}` | Standard uncertainty of assigned value |
| `{{ a.assigned_U }}` | Expanded uncertainty of assigned value |
| `{{ a.arithmetic_mean }}` | Arithmetic mean of reported results |
| `{{ a.standard_deviation }}` | Standard deviation for proficiency assessment |
| `{%tr for row in a.lab_rows %}` ... `{%tr endfor %}` | Repeats the table row. Inside it: `{{ row.lab_code }}`, `{{ row.result }}`, `{{ row.std_u }}`, `{{ row.z_score }}`, `{{ row.zeta_score }}` |
| `{{ a.z_chart }}` and `{{ a.zeta_chart }}` | Score charts, each alone in a cell of a 1 x 2 borderless table (78 mm wide images) |
| `{{ a.result_chart }}` | Lab results chart, alone in its paragraph (200 mm wide) |

**Page layout.** The template is divided into sections so each analyte gets a portrait page then a landscape page:

- The last paragraph of the cover ends a portrait section.
- The paragraph after the z/ζ chart table ends a portrait section (replaces a page break).
- Inside `{%p if not loop.last %}` ... `{%p endif %}`, a paragraph ends a landscape section. It is skipped after the last analyte.
- The document's final section settings are landscape, and govern the last analyte's landscape page.

**Rules when editing the template:**

- Put each `{%p ... %}` and `{%tr ... %}` tag alone in its own paragraph or table row, with nothing else in it.
- Type each tag in one go, without changing formatting partway through. Word can split a tag into several runs, which stops docxtpl from reading it.
- Don't delete the tag paragraphs or the tiny section-break paragraphs. Doing so breaks the loop or the page orientation.
- Keep the loop around the whole analyte section: from the results table to the result chart.
- The cover page still contains typed text that isn't filled by the app: the number of participating laboratories, the sample dispatch date and the interim report date. Update these by hand for each round, or add placeholders and inputs to the app.

## Troubleshooting

| Symptom | Likely cause and fix |
|---|---|
| `ModuleNotFoundError: matplotlib` or `ImportError` for `result_chart` | Dependencies or `charts.py` on GitHub are out of date. Push the current `charts.py` and `pyproject.toml` (and `uv.lock`), then reboot the app |
| All lab results appear in one table row | The row loop tags are wrong. Use `{%tr for row in a.lab_rows %}` and `{%tr endfor %}` in their own rows or cells, not `{% for %}` |
| A chart does not appear in the report | The placeholder is missing, misspelled, or sharing a paragraph or cell with other text. Check for the `a.` prefix |
| `Could not find ... header` or `could not find Assigned Value` | A tab or label was renamed, or the layout changed. Check the Excel file against the requirements above |
| `sheet not found` | The Excel file lacks `PT Scores`, `Robust Statistics` or `Comparison of Results` |
| Error bars or markers look wrong on the lab results chart | Check the `Comparison of Results` column offsets in `get_comparison_data` |
| Landscape page spills onto a second page | Remove an empty spacer paragraph above the chart heading, or reduce the chart width in `process_analyte` |
| Template fails with a Jinja error | A tag is split across runs or has a syntax error. Retype the tag as plain text |
| `openpyxl` warning that a `wmf` image is dropped | Harmless. The app only reads cell values and never saves the workbook |

## Design notes

- **One function per Excel file.** `process_analyte()` reads one workbook and returns a dictionary with everything the template needs. The generate step calls it once per analyte, so extraction and chart logic exist in one place.
- **Errors are collected, not fatal.** A problem in one analyte's file does not hide problems in the others.
- **Charts are drawn in memory.** matplotlib writes PNGs to `BytesIO` objects, not to disk, so simultaneous users on Streamlit Cloud can't overwrite each other's files.
- **Stable widget keys.** Each analyte block has a random ID used in its widget keys, so removing a middle block does not shift another block's uploaded file or typed values.
- **Download survives reruns.** The generated file is held in `st.session_state`, and the download button sits outside the Generate block.
