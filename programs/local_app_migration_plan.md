# Local App Migration Plan

## Goal

Replace the current notebook-first workflow with a simple local web app that allows an admin to:

1. Upload a DonSheet Excel file.
2. Choose one or more outputs to generate.
3. Preview the generated result in the browser.
4. Download the generated Excel or PDF files.
5. Run the tool locally without manually opening Python or Jupyter.

---

## Recommended Stack

Use **Streamlit** for the first version.

Why Streamlit is the best fit here:

- Very fast to build and maintain.
- Excellent support for file upload and download buttons.
- Easy dataframe/table preview in the browser.
- Runs locally at a URL like `http://localhost:8501`.
- Easier than Flask for this use case because you do not need to manually build much frontend code.
- Easier than NiceGUI if the main need is forms, buttons, previews, and downloads rather than a more custom UI.

Recommendation:

- **Phase 1:** Streamlit
- **Phase 2:** Only move to Flask or NiceGUI if later you need a more custom multi-page application

---

## Target Architecture

The new system should be separated into four layers:

### 1. Input Layer

Purpose:
- Read the uploaded DonSheet workbook.
- Load all relevant department sheets.
- Ignore helper sheets such as dropdown or mapping sheets unless needed.

Output:
- A raw combined dataframe from DonSheet.

Suggested module:
- `programs/app/io_donsheet.py`

### 2. Normalization Layer

Purpose:
- Convert DonSheet column names into one standard internal schema.
- Derive missing fields needed by the old reports and contract logic.
- Clean values, dates, blanks, and code mappings.

Output:
- A canonical dataframe used by every report and contract generator.

Suggested module:
- `programs/app/normalize.py`

Important idea:
- This is the key migration step.
- The app should not generate reports directly from raw DonSheet columns.
- Everything should first be converted into one stable internal dataframe schema.

### 3. Output Generator Layer

Purpose:
- Use the canonical dataframe to create:
  - report tables
  - Excel files
  - contract PDFs

Output:
- pandas dataframes for preview
- Excel files in memory
- PDF files in memory or temporary output folders

Suggested modules:
- `programs/app/reports/department_overview.py`
- `programs/app/reports/department_totals.py`
- `programs/app/reports/program_comparison.py`
- `programs/app/contracts/pdf_contracts.py`

Important idea:
- Each output should become a function.
- Example:
  - `generate_department_overview(df, options) -> ReportResult`
  - `generate_contract_pdfs(df, options) -> ContractResult`

### 4. UI Layer

Purpose:
- Show upload widget.
- Let the user choose output type.
- Preview tables.
- Offer download buttons.

Suggested module:
- `programs/app/streamlit_app.py`

---

## Canonical Dataframe Strategy

The old notebooks are currently doing two jobs at once:

1. Transforming old ShariSheet files into a usable dataframe.
2. Using that dataframe to generate outputs.

The new DonSheet workbook has a much cleaner structure, so the best path is:

### Do not port old notebook ingestion logic as-is

Instead:

1. Study the dataframe columns that the report logic actually needs.
2. Build a new DonSheet-to-canonical transformation layer.
3. Point all report generators to the canonical dataframe.

This will make the system:

- simpler
- easier to test
- easier to maintain
- less dependent on workbook layout changes

### Likely canonical columns

Based on the existing notebooks, the canonical dataframe will likely need fields such as:

- `Department`
- `Program`
- `Semester`
- `Fiscal Year`
- `Course`
- `Course Section`
- `Catalog Title`
- `CRN`
- `Mode`
- `Location`
- `Faculty Name`
- `ID#`
- `email`
- `telephone`
- `load/contract`
- `dept budget`
- `Cr.`
- `rate per credit`
- `total contract amount`
- `Begin Date`
- `End Date`
- `pre-work period start`
- `pre-work period end`
- `pre-work # of weeks`
- `pre-work hours/week`
- `pre-work hours/period`
- `intensive period start`
- `intensive period end`
- `intensive # of weeks`
- `intensive hours/week`
- `intensive hours/period`
- `post-work period start`
- `post-work period end`
- `post-work # of weeks`
- `post-work hours/week`
- `post-work hours/period`
- `total contract hours`
- `account to be charged`
- `international`
- `remote employee`
- `Dept contact name`
- `Dept contact ID#`

Some of these are already present in DonSheet under different names.
Some may need to be derived or merged from helper sheets.

---

## UI Design for Version 1

Keep the first version very simple.

### Page flow

1. Upload DonSheet file.
2. App validates workbook and shows summary:
   - sheets found
   - row counts
   - warnings
3. User chooses output type:
   - Department overview
   - Department totals
   - Program comparison
   - Contract PDFs
   - Combined workbook
4. User clicks `Generate`.
5. App shows:
   - preview table for table-based outputs
   - success/warning messages
   - download buttons

### For Excel-based reports

Show:
- HTML table preview in Streamlit
- Download button for `.xlsx`

### For PDF contracts

Show:
- number of contracts generated
- downloadable `.zip` file of PDFs

---

## Packaging for Local Use

Goal:
- Admin double-clicks one launcher.
- Browser opens automatically.
- No manual Python commands required.

### Best packaging approach for first working deployment

Use:

- a local Python environment
- Streamlit app
- small launcher script

Examples:
- macOS: `.command` file
- Windows: `.bat` file or packaged `.exe`

Launcher behavior:

1. Start the Streamlit app.
2. Open browser to `http://localhost:8501`.

### Best packaging approach for broader rollout

After the app is stable, package it with:

- **PyInstaller**

That gives admins a single app-like executable and avoids asking them to run Python manually.

Recommended rollout order:

1. Build the app.
2. Test with real DonSheets.
3. Package with PyInstaller.
4. Distribute to admins.

---

## Proposed Folder Structure

```text
programs/
  app/
    streamlit_app.py
    config.py
    models.py
    io_donsheet.py
    normalize.py
    excel_utils.py
    reports/
      department_overview.py
      department_totals.py
      fy_comparison.py
      combined_workbook.py
    contracts/
      pdf_contracts.py
    templates/
  local_app_migration_plan.md
```

Optional later:

```text
programs/
  tests/
    test_normalize.py
    test_department_overview.py
    test_contracts.py
```

---

## Migration Strategy

### Phase 1: Extract reusable logic from notebooks

Goal:
- move notebook code into plain Python functions

Work:
- identify output-producing cells
- extract them into modules
- remove notebook-only display code

### Phase 2: Build DonSheet normalization

Goal:
- create one canonical dataframe from DonSheet

Work:
- read all department sheets
- standardize column names
- derive semester/fiscal year
- map DonSheet fields to report fields

### Phase 3: Rebuild 2-3 priority outputs first

Start with:

1. `FY2027_OTST_DepartmentOverview_ContractsPerProgram.xlsx`
2. `FY2027_OTST_DepartmentOverview_LoadContractPerSemester.xlsx`
3. one contract PDF generation flow

Why:
- these give fast visible proof that the new architecture works
- they cover both table reports and PDF generation

### Phase 4: Build Streamlit interface

Goal:
- upload, select, preview, download

### Phase 5: Package the app

Goal:
- make it easy for admins to run locally

---

## Recommended Development Priorities

Priority order:

1. Define the canonical dataframe schema.
2. Build DonSheet reader and transformer.
3. Recreate one Excel report from DonSheet.
4. Recreate one contract PDF workflow from DonSheet.
5. Wrap both in Streamlit.
6. Package the app.

This order reduces risk because the hardest part is not the UI.
The hardest part is building a reliable canonical dataframe from DonSheet.

---

## Risks to Plan For

### 1. Hidden logic inside notebooks

Risk:
- some report rules may be embedded in ad hoc notebook cells

Mitigation:
- extract one report at a time
- compare output to existing sample workbooks

### 2. DonSheet may not contain every contract field directly

Risk:
- some PDF fields may depend on values that were previously manual or inferred

Mitigation:
- define fallback rules
- show warnings in UI when required fields are missing

### 3. Formatting parity with current Excel files

Risk:
- admins may expect current formatting, grouping, blank lines, and section labels

Mitigation:
- preserve output structure first for the highest-priority reports
- separate calculation logic from formatting logic

### 4. Packaging complexity

Risk:
- packaging too early will slow development

Mitigation:
- package only after the app works reliably in normal Python execution

---

## Recommended First Deliverable

Build a minimum viable app that can:

1. Upload one DonSheet.
2. Create one department overview Excel report.
3. Show the report as an HTML table.
4. Download the generated `.xlsx`.

After that:

5. Add the second department report.
6. Add contract PDF generation.
7. Add combined multi-report workbook generation.

---

## Recommendation Summary

The best path is:

1. Use **Streamlit**.
2. Create a **DonSheet-to-canonical-dataframe** layer.
3. Refactor notebook logic into reusable Python modules.
4. Rebuild a few priority outputs first.
5. Package later with **PyInstaller**.

The most important design decision is this:

**Treat DonSheet as the new source of truth and create one stable internal dataframe schema that every report and contract generator uses.**
