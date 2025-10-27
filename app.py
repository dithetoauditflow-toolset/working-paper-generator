# streamlit_app.py
import streamlit as st
import os
import tempfile
import time
from typing import List, Tuple, Any
import shutil
from datetime import datetime

# Import existing logic
from helper_funcs import get_company_info, create_folder_structure_for_all_working_papers
from helper_funcs import load_data_file
from tp_1 import process_files as process_tp1, process_files_for_all_processing as process_tp1_all
from tp_2 import process_files as process_tp2, process_files_for_all_processing as process_tp2_all
from tp_3 import process_files as process_tp3, process_files_for_all_processing as process_tp3_all
from tp_4 import process_files as process_tp4, process_files_for_all_processing as process_tp4_all


def get_template_paths() -> List[str]:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    templates_dir = os.path.join(script_dir, "TEMPLATES", "Working_Papers_Templates")
    if not os.path.exists(templates_dir):
        raise FileNotFoundError(
            f"Working_Papers_Templates folder not found in {os.path.join(script_dir, 'TEMPLATES')}"
        )
    templates = sorted([
        os.path.join(templates_dir, f) for f in os.listdir(templates_dir) if f.endswith(".xlsx")
    ])
    if len(templates) < 4:
        raise FileNotFoundError(
            "Not enough template files found. Ensure there are at least 4 templates in the Working_Papers_Templates folder."
        )
    return templates


def persist_uploaded_files(uploaded_files: List[Any]) -> List[str]:
    """Save uploaded files to a temp folder and return their paths."""
    if not uploaded_files:
        return []
    if "temp_dir" not in st.session_state:
        st.session_state.temp_dir = tempfile.mkdtemp(prefix="auditflow_in_")
    temp_dir = st.session_state.temp_dir
    saved_paths = []
    for uf in uploaded_files:
        dest = os.path.join(temp_dir, uf.name)
        with open(dest, "wb") as f:
            f.write(uf.getbuffer())
        saved_paths.append(dest)
    return saved_paths


REQUIRED_COLUMNS = [
    "TRADENAME",
    "UIFREFERENCENUMBER",
    "SHUTDOWN_FROM",
    "SHUTDOWN_TILL",
    "IDNUMBER",
    "PAYMENT_STATUS_ID",
    "PAYMENTMEDIUMID",
    "BANK_PAY_AMOUNT",
]


def check_required_columns(file_path: str) -> List[str]:
    """Return list of missing required columns for the given Excel file."""
    try:
        _, sheet = load_data_file(file_path)
        headers = [cell.value for cell in sheet[1]]
        missing = [c for c in REQUIRED_COLUMNS if c not in headers]
        return missing
    except Exception as e:
        # If file fails to load, treat as missing all to surface the error clearly
        return REQUIRED_COLUMNS + [f"Error: {e}"]


def show_company_overview(file_paths: List[str]) -> None:
    st.subheader("Review extracted details")
    for fp in file_paths:
        try:
            company_name, uif_ref, periods_claimed, number_of_employees, total_amount_claimed = get_company_info(fp)
            periods_list = periods_claimed.split(",") if isinstance(periods_claimed, str) and periods_claimed else []

            with st.container(border=True):
                st.markdown(f"### {company_name} · {os.path.basename(fp)}")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("UIF Ref", uif_ref or "-")
                c2.metric("Periods", str(len(periods_list)))
                c3.metric("Employees", int(number_of_employees) if number_of_employees is not None else 0)
                c4.metric("Total Claimed", f"R{total_amount_claimed}")

                with st.expander("View periods claimed", expanded=False):
                    if periods_list:
                        for p in periods_list:
                            st.write(f"- {p.strip()}")
                    else:
                        st.caption("No periods detected")
        except Exception as e:
            st.error(f"{os.path.basename(fp)}: {e}")


def validate_ready(files: List[str], consultant: str, templates: List[str]) -> Tuple[bool, List[str]]:
    issues = []
    if not files:
        issues.append("No data files uploaded")
    if not consultant:
        issues.append("Consultant name is required")
    if not templates or len(templates) < 4:
        issues.append("Templates missing or incomplete")
    return (len(issues) == 0, issues)


def process_single_wp(wp_index: int, name: str, file_path: str, template_paths: List[str], consultant: str, outdir: str):
    funcs = [process_tp1, process_tp2, process_tp3, process_tp4]
    return funcs[wp_index](file_path, template_paths[wp_index], consultant, outdir)


def process_all_for_file(file_path: str, template_paths: List[str], consultant: str, outdir: str):
    # Create structure once per file
    company_name, uif_ref, periods_claimed, number_of_employees, total_amount_claimed = get_company_info(file_path)
    audit_working_papers_folder = create_folder_structure_for_all_working_papers(
        outdir, company_name, uif_ref, file_path, template_paths
    )
    funcs_all = [process_tp1_all, process_tp2_all, process_tp3_all, process_tp4_all]
    for i in range(4):
        funcs_all[i](file_path, template_paths[i], consultant, audit_working_papers_folder)


def main():
    st.set_page_config(page_title="AuditFlow Working Paper Generator", page_icon="📄", layout="wide")
    st.title("AuditFlow Working Paper Generator")
    st.markdown(
        """
        Streamlined generation of audit working papers (TP.1–TP.4) from UIF Excel data files.

        Instructions:
        1) Upload one or more UIF Excel data files (.xlsx) in the sidebar.
        2) Click "Load Templates" to load working paper templates from `TEMPLATES/Working_Papers_Templates/`.
        3) Review the extracted company details below. If required columns are missing, fix your files and re-upload.
        4) When everything looks good, use the action buttons to generate working papers and download a ZIP.
        """
    )

    with st.sidebar:
        st.header("Setup")
        consultant = st.text_input("Consultant Name", key="consultant_name")
        uploaded = st.file_uploader(
            "Upload UIF Excel data files (.xlsx)", type=["xlsx"], accept_multiple_files=True
        )
        # Auto-load templates on first launch
        if "template_paths" not in st.session_state:
            try:
                st.session_state.template_paths = get_template_paths()
                st.success(f"Templates loaded: {len(st.session_state.template_paths)} found")
            except Exception as e:
                st.error(f"Failed to load templates: {e}")

    files = persist_uploaded_files(uploaded) if uploaded else []
    template_paths = st.session_state.get("template_paths", [])

    can_generate = False
    validation_report = []
    if files:
        # Validate templates loaded
        if not template_paths:
            st.warning("Templates not loaded. Ensure the templates directory exists and contains at least 4 .xlsx files.")
        # Validate required columns per file
        validation_report = [
            (os.path.basename(fp), check_required_columns(fp)) for fp in files
        ]
        missing_any = [(name, missing) for name, missing in validation_report if missing]
        if missing_any:
            st.error("Some files are missing required columns. Fix and re-upload:")
            with st.expander("See validation details"):
                for name, missing in missing_any:
                    st.write(f"{name}: missing -> {', '.join(map(str, missing))}")
        else:
            show_company_overview(files)
            can_generate = bool(template_paths)

    # Action selections (only after upload + validation success)
    btn_tp1 = btn_tp2 = btn_tp3 = btn_tp4 = btn_all = False
    if can_generate:
        st.markdown("---")
        st.subheader("Generate Working Papers")

        # Row 1: Four buttons spanning the width
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            btn_tp1 = st.button("Generate TP.1 for All Files", use_container_width=True)
        with c2:
            btn_tp2 = st.button("Generate TP.2 for All Files", use_container_width=True)
        with c3:
            btn_tp3 = st.button("Generate TP.3 for All Files", use_container_width=True)
        with c4:
            btn_tp4 = st.button("Generate TP.4 for All Files", use_container_width=True)

        # Row 2: Single full-width 'Generate ALL' button
        btn_all = st.button("Generate ALL (TP.1 - TP.4) for All Files", use_container_width=True)

    if any([btn_tp1, btn_tp2, btn_tp3, btn_tp4, btn_all]):
        ready, issues = validate_ready(files, consultant, template_paths)
        if not ready:
            for issue in issues:
                st.error(issue)
            st.stop()

        results = []
        overall_start = time.time()
        progress = st.progress(0)
        status_area = st.empty()
        # Prepare temporary output directory per session
        if "output_dir" not in st.session_state:
            st.session_state.output_dir = tempfile.mkdtemp(prefix="auditflow_out_")
        outdir = st.session_state.output_dir

        try:
            for idx, fp in enumerate(files):
                file_name = os.path.basename(fp)
                status_area.info(f"Processing: {file_name} ({idx+1}/{len(files)})")
                start = time.time()
                try:
                    if btn_all:
                        process_all_for_file(fp, template_paths, consultant, outdir)
                    else:
                        if btn_tp1:
                            process_single_wp(0, "TP.1", fp, template_paths, consultant, outdir)
                        if btn_tp2:
                            process_single_wp(1, "TP.2", fp, template_paths, consultant, outdir)
                        if btn_tp3:
                            process_single_wp(2, "TP.3", fp, template_paths, consultant, outdir)
                        if btn_tp4:
                            process_single_wp(3, "TP.4", fp, template_paths, consultant, outdir)
                    duration = time.time() - start
                    results.append({
                        "File": file_name,
                        "Status": "Success",
                        "Time": f"{int(duration//60)}m {int(duration%60)}s {int((duration%1)*1000)}ms",
                    })
                except Exception as e:
                    duration = time.time() - start
                    results.append({
                        "File": file_name,
                        "Status": f"Failed: {e}",
                        "Time": f"{int(duration//60)}m {int(duration%60)}s {int((duration%1)*1000)}ms",
                    })
                progress.progress(int(((idx + 1) / len(files)) * 100))
        finally:
            total = time.time() - overall_start
            progress.progress(100)
            status_area.empty()

        st.success("Processing complete")
        st.subheader("Results")
        st.dataframe(results, use_container_width=True, hide_index=True)
        st.info(f"Total time: {int(total//60)}m {int(total%60)}s {int((total%1)*1000)}ms")

        # Zip the output folder and provide download with dynamic date and company count
        date_str = datetime.now().strftime("%Y-%m-%d")
        company_count = len(files)
        zip_base_name = f"auditflow_working_papers_{date_str}_x{company_count}"
        zip_base = os.path.join(tempfile.gettempdir(), zip_base_name)
        zip_path = shutil.make_archive(zip_base, 'zip', outdir)
        with open(zip_path, 'rb') as f:
            st.download_button(
                label="Download Generated Working Papers (ZIP)",
                data=f,
                file_name=f"{zip_base_name}.zip",
                mime="application/zip",
            )


if __name__ == "__main__":
    main()
