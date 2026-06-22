from __future__ import annotations

import os
from pathlib import Path
import signal
import threading

import streamlit as st
import streamlit.components.v1 as components

from app.contracts_from_donsheet import cleanup_result, generate_contracts_from_uploads
from app.contracts_from_sharisheet import generate_contracts_from_sharisheet_uploads
from app.config import (
    APP_TITLE,
    DEFAULT_DONSHEET_PATH,
    DEFAULT_SHARISHEET_DIR,
    DEPARTMENT_OPTIONS,
    ENABLE_CLOSE_BUTTON,
    ROLE_PASSWORDS,
)
from app.normalize import build_canonical_dataframe, combine_normalization_results
from app.schedule_from_donsheet import generate_schedule_outputs
from app.sharisheet import normalize_sharisheet_uploads, sharisheet_payloads_from_directory
from app.reports.demo_reports import (
    generate_budget_reports_bundle,
    generate_department_overview_bundle,
)


ROLE_KEY = "smart_seminary_role"
AUTH_KEY_PREFIX = "smart_seminary_auth_"
SHARI_CONTRACT_RESULT_KEY = "smart_seminary_shari_contract_result"
SHARI_BUDGET_RESULT_KEY = "smart_seminary_shari_budget_result"
DEPARTMENT_ADMIN_RESULT_KEY = "smart_seminary_department_admin_result"

ACTION_CONTRACT_DONSHEET = "Create contract PDFs - based on DonSheet"
ACTION_BUDGET_DONSHEET = "Create budget report workbook - based on DonSheet"
ACTION_CONTRACT_SHARISHEET = "Create contract PDFs - based on ShariSheet"
ACTION_BUDGET_SHARISHEET = "Create budget report workbook - based on ShariSheet"


def _shutdown_app() -> None:
    os.kill(os.getpid(), signal.SIGTERM)


def _schedule_shutdown(delay_seconds: float = 1.0) -> None:
    timer = threading.Timer(delay_seconds, _shutdown_app)
    timer.daemon = True
    timer.start()


def _credits_per_teacher_html_table(dataframe) -> str:
    columns = list(dataframe.columns)
    header_cells = "".join(f"<td>{column}</td>" for column in columns)
    rows_html: list[str] = []

    for _, row in dataframe.iterrows():
        values = [row[column] for column in columns]
        non_empty = [value for value in values if value not in ("", None)]
        is_header = non_empty == columns
        first_value = values[0]
        is_total = isinstance(first_value, str) and first_value.startswith("total credits:")

        cell_html: list[str] = []
        for idx, value in enumerate(values):
            display = "" if value is None else value
            styles = ["padding:6px 10px", "border:1px solid #d1d5db"]
            if is_header:
                styles.extend(["background-color:#d9d9d9", "font-weight:700", "text-align:center"])
            elif is_total:
                if idx == 0:
                    styles.extend(["color:#0000ff", "font-weight:700", "text-align:right"])
                elif display != "":
                    styles.append("color:#0000ff")
            cell_html.append(f"<td style=\"{'; '.join(styles)}\">{display}</td>")
        rows_html.append(f"<tr>{''.join(cell_html)}</tr>")

    return (
        "<table style=\"border-collapse:collapse; width:100%;\">"
        f"<tbody><tr style=\"display:none\">{header_cells}</tr>{''.join(rows_html)}</tbody></table>"
    )


def _build_result(uploaded_files, use_sample: bool):
    results = []
    source_names = []

    for uploaded_file in uploaded_files:
        results.append(build_canonical_dataframe(uploaded_file.getvalue()))
        source_names.append(uploaded_file.name)

    if use_sample and not uploaded_files:
        default_path = Path(DEFAULT_DONSHEET_PATH)
        results.append(build_canonical_dataframe(default_path.read_bytes()))
        source_names.append(default_path.name)

    combined_result = combine_normalization_results(results)
    return combined_result, source_names


def _set_role(role: str) -> None:
    st.session_state[ROLE_KEY] = role


def _back_to_home() -> None:
    st.session_state.pop(ROLE_KEY, None)


def _is_authenticated(role: str) -> bool:
    return bool(st.session_state.get(f"{AUTH_KEY_PREFIX}{role}", False))


def _render_password_gate(role: str) -> bool:
    if _is_authenticated(role):
        return True

    with st.form(f"{role.lower()}_login_form"):
        st.subheader(f"{role} access")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Unlock", type="primary")
        if submitted:
            if password == ROLE_PASSWORDS[role]:
                st.session_state[f"{AUTH_KEY_PREFIX}{role}"] = True
                st.rerun()
            st.error("Incorrect password.")
    return False


def _render_workbook_summary(result, source_names: list[str]) -> None:
    st.subheader("Workbook summary")
    st.write(f"Source files: {', '.join(f'`{name}`' for name in source_names)}")
    st.write(f"Uploaded workbooks: {len(source_names)}")
    st.write(f"Data sheets: {', '.join(result.summary.data_sheet_names)}")
    st.write(f"Normalized rows: {len(result.canonical_df)}")

    if result.warnings:
        for warning in result.warnings:
            st.warning(warning)

    st.subheader("Canonical dataframe preview")
    st.dataframe(result.canonical_df.head(50), use_container_width=True)


def _render_generator_output(result, source_names: list[str], report) -> None:
    _render_workbook_summary(result, source_names)
    st.subheader(report.report_name)
    st.dataframe(report.preview_df.head(100), use_container_width=True)
    st.download_button(
        label=f"Download {report.output_filename}",
        data=report.excel_bytes,
        file_name=report.output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _render_report_bundle(result, source_names: list[str], bundle) -> None:
    _render_workbook_summary(result, source_names)
    preview_key = bundle.preview_report_key
    selected_key = st.selectbox(
        "Preview report",
        options=list(bundle.reports.keys()),
        index=list(bundle.reports.keys()).index(preview_key) if preview_key in bundle.reports else 0,
        key="budget_bundle_preview_select",
    )
    selected_report = bundle.reports[selected_key]
    st.subheader(selected_report.report_name)
    st.dataframe(selected_report.preview_df.head(100), use_container_width=True)

    st.download_button(
        label=f"Download {bundle.combined_output_filename}",
        data=bundle.combined_workbook_bytes,
        file_name=bundle.combined_output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="budget_bundle_download_all",
    )

    st.subheader("Individual report downloads")
    for report_key, report in bundle.reports.items():
        st.download_button(
            label=f"Download {report_key}",
            data=report.excel_bytes,
            file_name=report.output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{report_key}",
        )


def _load_result_for_current_role(uploaded_files, use_sample: bool):
    if not uploaded_files and not use_sample:
        st.error("Upload a DonSheet workbook or enable the bundled DonSheet option.")
        return None, []

    with st.spinner("Normalizing DonSheet workbook..."):
        result, source_names = _build_result(uploaded_files, use_sample)
    return result, source_names


def _sharisheet_payloads(uploaded_files, use_folder: bool) -> list[tuple[str, bytes]]:
    if uploaded_files:
        return [(file.name, file.getvalue()) for file in uploaded_files]
    if use_folder:
        return sharisheet_payloads_from_directory(Path(DEFAULT_SHARISHEET_DIR))
    return []


def _load_sharisheet_result(uploaded_files, use_folder: bool):
    payloads = _sharisheet_payloads(uploaded_files, use_folder)
    if not payloads:
        st.error(
            "Upload ShariSheet workbook(s) or enable the SmartBudgeting ShariSheets folder option."
        )
        return None, []
    with st.spinner("Normalizing ShariSheet workbook(s)..."):
        result = normalize_sharisheet_uploads(payloads)
    return result, [name for name, _ in payloads]


def _render_department_admins() -> None:
    st.subheader("Department Admins")
    st.write("Choose a department, upload a DonSheet, and generate the department overview workbook.")

    selected_department = st.radio(
        "Department",
        options=DEPARTMENT_OPTIONS,
        horizontal=True,
    )
    uploaded_files = st.file_uploader(
        "Upload DonSheet workbook",
        type=["xlsx"],
        accept_multiple_files=True,
        key="department_admin_upload",
    )
    use_sample = st.checkbox(
        "Use bundled default DonSheet when no file is uploaded",
        value=True,
        key="department_admin_sample",
    )

    if st.button("Generate department overview reports", type="primary", key="department_admin_generate"):
        result, source_names = _load_result_for_current_role(uploaded_files, use_sample)
        if result is None:
            return
        report = generate_department_overview_bundle(result.canonical_df, selected_department)
        st.session_state[DEPARTMENT_ADMIN_RESULT_KEY] = {
            "result": result,
            "source_names": source_names,
            "report": report,
        }

    payload = st.session_state.get(DEPARTMENT_ADMIN_RESULT_KEY)
    if payload is not None:
        top_left, top_right = st.columns([5, 1])
        with top_right:
            if st.button("Clear admin results", use_container_width=True, key="department_admin_clear"):
                st.session_state.pop(DEPARTMENT_ADMIN_RESULT_KEY, None)
                st.rerun()
        _render_generator_output(payload["result"], payload["source_names"], payload["report"])


def _render_karen() -> None:
    st.subheader("Karen")
    st.write("Upload one DonSheet and create a printable SmartSchedule Excel file plus a PDF version.")

    if not _render_password_gate("Karen"):
        return

    uploaded_file = st.file_uploader(
        "Upload DonSheet workbook",
        type=["xlsx"],
        accept_multiple_files=False,
        key="karen_upload",
    )
    use_sample = st.checkbox(
        "Use bundled default DonSheet when no file is uploaded",
        value=False,
        key="karen_sample",
    )

    if st.button("Generate SmartSchedule file", type="primary", key="karen_generate"):
        if not uploaded_file and not use_sample:
            st.error("Upload one DonSheet workbook or enable the bundled DonSheet option.")
            return
        if uploaded_file:
            source_name = uploaded_file.name
            source_bytes = uploaded_file.getvalue()
        else:
            default_path = Path(DEFAULT_DONSHEET_PATH)
            source_name = default_path.name
            source_bytes = default_path.read_bytes()
        with st.spinner("Building SmartSchedule outputs..."):
            schedule_result = generate_schedule_outputs(source_name, source_bytes)

        st.subheader("Schedule preview")
        st.dataframe(schedule_result.preview_df.head(100), use_container_width=True)
        st.download_button(
            label=f"Download {schedule_result.excel_filename}",
            data=schedule_result.excel_bytes,
            file_name=schedule_result.excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="karen_download_excel",
        )
        st.download_button(
            label=f"Download {schedule_result.pdf_filename}",
            data=schedule_result.pdf_bytes,
            file_name=schedule_result.pdf_filename,
            mime="application/pdf",
            key="karen_download_pdf",
        )
        st.subheader("Normalized meetings")
        st.dataframe(schedule_result.normalized_meetings_df.head(100), use_container_width=True)
        if not schedule_result.excluded_df.empty:
            st.subheader("Excluded rows")
            st.dataframe(schedule_result.excluded_df.head(100), use_container_width=True)


def _render_shari() -> None:
    st.subheader("Shari")
    st.write("Upload a DonSheet or ShariSheet and create contract PDFs or budget report workbooks.")

    if not _render_password_gate("Shari"):
        return

    action = st.radio(
        "Output",
        options=[
            ACTION_CONTRACT_DONSHEET,
            ACTION_BUDGET_DONSHEET,
            ACTION_CONTRACT_SHARISHEET,
            ACTION_BUDGET_SHARISHEET,
        ],
        key="shari_action",
    )
    is_sharisheet_action = action in {ACTION_CONTRACT_SHARISHEET, ACTION_BUDGET_SHARISHEET}
    uploaded_files = st.file_uploader(
        "Upload ShariSheet workbook(s)" if is_sharisheet_action else "Upload DonSheet workbook",
        type=["xlsx"],
        accept_multiple_files=True,
        key="shari_upload",
    )
    if is_sharisheet_action:
        use_sample = st.checkbox(
            f"Use SmartBudgeting ShariSheets folder when no file is uploaded ({DEFAULT_SHARISHEET_DIR})",
            value=True,
            key="shari_sample",
        )
    else:
        use_sample = st.checkbox(
            "Use bundled default DonSheet when no file is uploaded",
            value=True,
            key="shari_sample",
        )

    if action in {ACTION_CONTRACT_DONSHEET, ACTION_CONTRACT_SHARISHEET}:
        create_clicked = st.button("Generate contract PDFs", type="primary", key="shari_generate_contracts")
        if create_clicked:
            previous = st.session_state.pop(SHARI_CONTRACT_RESULT_KEY, None)
            if previous is not None:
                cleanup_result(previous)
            if is_sharisheet_action:
                upload_payloads = _sharisheet_payloads(uploaded_files, use_sample)
                if not upload_payloads:
                    st.error(
                        "Upload ShariSheet workbook(s) or enable the SmartBudgeting ShariSheets folder option."
                    )
                    return
            else:
                if not uploaded_files and not use_sample:
                    st.error("Upload a DonSheet workbook or enable the bundled DonSheet option.")
                    return
                upload_payloads = [(file.name, file.getvalue()) for file in uploaded_files]
                if use_sample and not uploaded_files:
                    default_path = Path(DEFAULT_DONSHEET_PATH)
                    upload_payloads = [(default_path.name, default_path.read_bytes())]
            with st.spinner("Creating contract PDFs..."):
                result = (
                    generate_contracts_from_sharisheet_uploads(upload_payloads)
                    if is_sharisheet_action
                    else generate_contracts_from_uploads(upload_payloads)
                )
            st.session_state[SHARI_CONTRACT_RESULT_KEY] = result

        contract_result = st.session_state.get(SHARI_CONTRACT_RESULT_KEY)
        if contract_result is not None:
            st.success(
                f"Created {len(contract_result.contracts)} contract PDF(s) from {contract_result.records_processed} extracted contract row(s)."
            )
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Download all contracts (.zip)",
                    data=contract_result.zip_path.read_bytes(),
                    file_name="SmartSeminary_Contracts.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
            with col2:
                if st.button("Clear contract results", use_container_width=True, key="shari_clear_contracts"):
                    cleanup_result(contract_result)
                    st.session_state.pop(SHARI_CONTRACT_RESULT_KEY, None)
                    st.rerun()

            if not contract_result.extracted_rows.empty:
                st.subheader("Extracted contract rows")
                preview_cols = [
                    "Source File",
                    "Source Sheet",
                    "Source Row",
                    "Department",
                    "Semester",
                    "Year",
                    "Instructor",
                    "CourseRaw",
                    "CourseSection",
                    "CourseTitle",
                    "Credits",
                    "Rate",
                    "total contract amount",
                    "Reason",
                ]
                preview_df = contract_result.extracted_rows.loc[
                    :, [col for col in preview_cols if col in contract_result.extracted_rows.columns]
                ]
                st.dataframe(preview_df, use_container_width=True, hide_index=True)

            if contract_result.contracts:
                st.subheader("Individual downloads")
                contract_names = [contract.filename for contract in contract_result.contracts]
                selected_name = st.selectbox(
                    "Choose one generated PDF",
                    options=contract_names,
                    key="shari_contract_pdf_select",
                )
                selected_contract = next(
                    contract for contract in contract_result.contracts if contract.filename == selected_name
                )
                st.download_button(
                    label=f"Download {selected_contract.filename}",
                    data=selected_contract.path.read_bytes(),
                    file_name=selected_contract.filename,
                    mime="application/pdf",
                    key=f"download-{selected_contract.filename}",
                )
        return

    if st.button("Generate Shari output", type="primary", key="shari_generate_budget"):
        if is_sharisheet_action:
            result, source_names = _load_sharisheet_result(uploaded_files, use_sample)
            source_label = "sharisheet"
        else:
            result, source_names = _load_result_for_current_role(uploaded_files, use_sample)
            source_label = "donsheet"
        if result is None:
            return
        bundle = generate_budget_reports_bundle(result.canonical_df, source_label=source_label)
        st.session_state[SHARI_BUDGET_RESULT_KEY] = {
            "result": result,
            "source_names": source_names,
            "bundle": bundle,
        }

    budget_payload = st.session_state.get(SHARI_BUDGET_RESULT_KEY)
    if budget_payload is not None:
        left_col, right_col = st.columns([5, 1])
        with right_col:
            if st.button("Clear budget results", use_container_width=True, key="shari_clear_budget_results"):
                st.session_state.pop(SHARI_BUDGET_RESULT_KEY, None)
                st.rerun()
        _render_report_bundle(
            budget_payload["result"],
            budget_payload["source_names"],
            budget_payload["bundle"],
        )


def _render_home() -> None:
    st.write("Choose which workspace you want to open.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.button("Department Admins", use_container_width=True, on_click=_set_role, args=("Department Admins",))
    with col2:
        st.button("Karen", use_container_width=True, on_click=_set_role, args=("Karen",))
    with col3:
        st.button("Shari", use_container_width=True, on_click=_set_role, args=("Shari",))


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    title_col, action_col = st.columns([6, 1])
    with title_col:
        st.title(APP_TITLE)
    with action_col:
        close_clicked = st.button("Close App", use_container_width=True, disabled=not ENABLE_CLOSE_BUTTON)

    if close_clicked and ENABLE_CLOSE_BUTTON:
        components.html(
            """
            <script>
            const closeAttempts = () => {
              try { window.top.open('', '_self'); } catch (e) {}
              try { window.top.close(); } catch (e) {}
              try { window.open('', '_self'); } catch (e) {}
              try { window.close(); } catch (e) {}
              try { self.close(); } catch (e) {}
            };
            closeAttempts();
            const shutdownDoc = `
              <!doctype html>
              <html>
                <head>
                  <meta charset="utf-8" />
                  <title>App Closed</title>
                  <style>
                    body {
                      font-family: Arial, sans-serif;
                      display: flex;
                      align-items: center;
                      justify-content: center;
                      min-height: 100vh;
                      margin: 0;
                      background: #f8fafc;
                      color: #0f172a;
                    }
                    .message {
                      padding: 24px 32px;
                      border: 1px solid #cbd5e1;
                      background: white;
                      border-radius: 12px;
                      box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
                      font-size: 24px;
                      font-weight: 600;
                    }
                  </style>
                </head>
                <body>
                  <div class="message">You can now close this tab</div>
                </body>
              </html>`;
            setTimeout(() => {
              closeAttempts();
              window.top.location.href = "data:text/html;charset=utf-8," + encodeURIComponent(shutdownDoc);
            }, 150);
            </script>
            """,
            height=0,
        )
        _schedule_shutdown(delay_seconds=1.5)
        st.stop()

    st.caption(f"Bundled default DonSheet: `{Path(DEFAULT_DONSHEET_PATH).name}`")

    role = st.session_state.get(ROLE_KEY)
    if role:
        top_left, top_right = st.columns([6, 1])
        with top_right:
            st.button("Back", use_container_width=True, on_click=_back_to_home)
    if role == "Department Admins":
        _render_department_admins()
    elif role == "Karen":
        _render_karen()
    elif role == "Shari":
        _render_shari()
    else:
        _render_home()


if __name__ == "__main__":
    main()
