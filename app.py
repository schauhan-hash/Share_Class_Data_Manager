from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
import json

import streamlit as st

from share_class_manager.core import (
    diff_records,
    fill_missing_isins,
    normalize_records,
    validate_dataset,
    validate_record,
)
from share_class_manager.schema import (
    DATE_FIELDS,
    DEFAULT_COLUMNS,
    FIELD_GROUPS,
    FIELD_ORDER,
    ROLE_PERMISSIONS,
    SELECT_OPTIONS,
)
from share_class_manager.xlsx_io import read_workbook, write_workbook

ROOT = Path(__file__).parent
SOURCE_FILE = ROOT / "data" / "Genericized_Sample_Data.xlsx"
TEMPLATE_FILE = ROOT / "data" / "Corrected_Excel_File_Format.xlsx"

st.set_page_config(
    page_title="HVP Share Class Data Manager",
    page_icon="H",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_style():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');
        :root {
            --ink: #173a3a;
            --teal: #147d72;
            --teal-dark: #0d5f57;
            --gold: #c48a2c;
            --paper: #f7f4ed;
            --card: #fffefa;
            --muted: #667775;
            --line: #dce3de;
        }
        html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
        .stApp { background: var(--paper); color: var(--ink); }
        h1, h2, h3 { font-family: 'Manrope', sans-serif !important; color: var(--ink); }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #263332 0%, #182524 100%);
            border-right: 5px solid var(--gold);
        }
        [data-testid="stSidebar"] * { color: #edf3ef !important; }
        [data-testid="stSidebar"] input, [data-testid="stSidebar"] select,
        [data-testid="stSidebar"] [data-baseweb="select"] * { color: var(--ink) !important; }
        [data-testid="stMetric"] {
            background: var(--card);
            border: 1px solid var(--line);
            border-top: 4px solid var(--teal);
            border-radius: 14px;
            padding: 18px;
            box-shadow: 0 5px 18px rgba(24, 56, 53, .06);
        }
        [data-testid="stMetricValue"] { color: var(--teal-dark); }
        .hero {
            background: linear-gradient(120deg, #fffefa, #e6f1ec);
            border: 1px solid var(--line);
            border-left: 7px solid var(--teal);
            border-radius: 16px;
            padding: 24px 28px;
            margin-bottom: 18px;
        }
        .eyebrow {
            color: var(--teal);
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: .12em;
            font-size: .74rem;
        }
        .badge {
            display: inline-block; padding: 5px 10px; border-radius: 999px;
            background: #dceee8; color: var(--teal-dark); font-weight: 700;
        }
        div.stButton > button, div.stDownloadButton > button {
            border-radius: 10px; border: 1px solid var(--teal);
            background: var(--teal); color: white; font-weight: 700;
        }
        div.stButton > button:hover, div.stDownloadButton > button:hover {
            background: var(--teal-dark); border-color: var(--teal-dark);
        }
        .demo-note {
            padding: 10px 14px; border-radius: 10px; background: rgba(196,138,44,.15);
            border: 1px solid rgba(196,138,44,.5); color: #6b4b16; font-size: .87rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def initialize_state():
    if "records" not in st.session_state:
        st.session_state.records = normalize_records(read_workbook(SOURCE_FILE))
        st.session_state.pending = []
        st.session_state.history = []
        st.session_state.role = "Viewer"
        st.session_state.data_version = 0


def permissions():
    return ROLE_PERMISSIONS[st.session_state.role]


def actor():
    role = st.session_state.role
    return {
        "Viewer": "demo.viewer@harbourvest.com",
        "Editor": "schauhan@harbourvest.com",
        "Reviewer": "demo.reviewer@harbourvest.com",
        "Admin": "demo.admin@harbourvest.com",
    }[role]


def now_utc():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def filtered_records():
    records = st.session_state.records
    search = st.session_state.get("search_filter", "").strip().lower()
    if search:
        records = [
            row
            for row in records
            if any(search in str(row.get(field, "")).lower() for field in FIELD_ORDER)
        ]
    for label, field in [
        ("lifecycle_filter", "Share Class Lifecycle"),
        ("currency_filter", "Share Class Currency"),
        ("type_filter", "Share Class Type"),
    ]:
        selected = st.session_state.get(label, "All")
        if selected and selected != "All":
            records = [row for row in records if row.get(field, "") == selected]
    return records


def page_title(eyebrow, title, description):
    st.markdown(
        f'<div class="hero"><div class="eyebrow">{eyebrow}</div>'
        f'<h1 style="margin:.25rem 0 .4rem">{title}</h1>'
        f'<div style="color:#607270;max-width:900px">{description}</div></div>',
        unsafe_allow_html=True,
    )


def compact_table(records, columns=DEFAULT_COLUMNS):
    selected = [column for column in columns if column in FIELD_ORDER]
    st.dataframe(
        [{column: record.get(column, "") for column in selected} for record in records],
        use_container_width=True,
        hide_index=True,
        height=430,
    )


def dashboard_page():
    page_title(
        "Control centre",
        "Share class overview",
        "A concise view of master-data coverage, identifiers, and quality signals.",
    )
    records = st.session_state.records
    issues = validate_dataset(records)
    errors = [issue for issue in issues if issue["severity"] == "Error"]
    unique_isins = len({row["ISIN"] for row in records if row["ISIN"]})
    active = sum(row.get("Share Class Lifecycle") == "Active" for row in records)
    columns = st.columns(4)
    columns[0].metric("Share classes", len(records), "demo dataset")
    columns[1].metric("Active", active)
    columns[2].metric("Unique ISINs", unique_isins, f"{len(records) - unique_isins} missing/duplicate")
    columns[3].metric("Data issues", len(issues), f"{len(errors)} blocking")

    left, right = st.columns([1.15, 1])
    with left:
        st.subheader("Classes by currency")
        currency_counts = {}
        for row in records:
            value = row.get("Share Class Currency") or "Missing"
            currency_counts[value] = currency_counts.get(value, 0) + 1
        st.bar_chart(currency_counts, color="#147d72")
        st.caption("Distribution of the current working dataset.")
    with right:
        st.subheader("Data completeness")
        field_counts = []
        for field in FIELD_ORDER:
            completed = sum(bool(row.get(field)) and str(row[field]).lower() != "n/a" for row in records)
            if completed < len(records):
                field_counts.append(
                    {
                        "Field": field,
                        "Completion": round(100 * completed / max(len(records), 1), 1),
                    }
                )
        field_counts.sort(key=lambda row: row["Completion"])
        st.dataframe(field_counts[:10], hide_index=True, use_container_width=True, height=360)
    st.info(
        "This MVP stores edits in the browser session. Download the workbook before the app restarts "
        "or use a database before moving to production."
    )


def records_page():
    page_title(
        "Master data",
        "Share class records",
        "Search a compact grid, then open a grouped editor without managing all 37 fields at once.",
    )
    controls = st.columns([2, 1, 1, 1])
    controls[0].text_input("Search any field", key="search_filter", placeholder="ISIN, share class, TPA ID...")
    records = st.session_state.records
    lifecycle_values = ["All"] + sorted({row.get("Share Class Lifecycle") for row in records} - {""})
    currency_values = ["All"] + sorted({row.get("Share Class Currency") for row in records} - {""})
    type_values = ["All"] + sorted({row.get("Share Class Type") for row in records} - {""})
    controls[1].selectbox("Lifecycle", lifecycle_values, key="lifecycle_filter")
    controls[2].selectbox("Currency", currency_values, key="currency_filter")
    controls[3].selectbox("Type", type_values, key="type_filter")
    visible = filtered_records()
    st.caption(f"Showing {len(visible)} of {len(records)} records")
    compact_table(visible)

    if "edit" not in permissions():
        st.info(f"{st.session_state.role} can view records. Switch to Editor or Admin to propose changes.")
        return

    editor_tab, delete_tab = st.tabs(["Propose add / edit", "Propose deletion"])
    with editor_tab:
        render_editor()
    with delete_tab:
        render_delete_request()


def field_widget(record, field, prefix, key_suffix):
    value = record.get(field, "")
    label = field + (" *" if field in __import__(
        "share_class_manager.schema", fromlist=["REQUIRED_FIELDS"]
    ).REQUIRED_FIELDS else "")
    if field == "ISIN":
        st.text_input(
            label,
            value=value,
            disabled=False,
            help="Leave blank only when proposing a new record. The next generated ISIN is assigned on submit.",
            key=f"{prefix}_{key_suffix}_{field}",
        )
        return st.session_state[f"{prefix}_{key_suffix}_{field}"]
    if field in SELECT_OPTIONS:
        options = SELECT_OPTIONS[field]
        if value and value not in options:
            options = [value] + options
        selected = st.selectbox(label, options, index=options.index(value) if value in options else 0, key=f"{prefix}_{key_suffix}_{field}")
        return selected
    if field == "Comments" or field == "Initial Management Fee Note":
        return st.text_area(label, value=value, height=90, key=f"{prefix}_{key_suffix}_{field}")
    return st.text_input(label, value=value, key=f"{prefix}_{key_suffix}_{field}")


def render_editor():
    record_options = {
        f"{row.get('ISIN') or 'No ISIN'} — {row.get('TPA Share Class ID') or row.get('Share Class', 'Unnamed')}": index
        for index, row in enumerate(st.session_state.records)
    }
    mode = st.radio("Action", ["Add a record", "Edit an existing record"], horizontal=True)
    if mode == "Add a record":
        before = {}
        record = {field: "" for field in FIELD_ORDER}
        prefix = "new"
    else:
        label = st.selectbox("Select record", list(record_options))
        before = st.session_state.records[record_options[label]]
        record = deepcopy(before)
        prefix = f"edit_{record_options[label]}"

    with st.form(f"{prefix}_form", clear_on_submit=prefix == "new"):
        tabs = st.tabs(list(FIELD_GROUPS))
        edited = deepcopy(record)
        for tab, (group, fields) in zip(tabs, FIELD_GROUPS.items()):
            with tab:
                widget_columns = st.columns(2)
                for position, field in enumerate(fields):
                    with widget_columns[position % 2]:
                        edited[field] = field_widget(edited, field, prefix, mode.replace(" ", "_"))
        reason = st.text_area(
            "Reason for change",
            placeholder="Describe the purpose of this proposed change.",
            height=80,
        )
        submitted = st.form_submit_button("Submit for review", use_container_width=True)

    if submitted:
        if not reason.strip():
            st.error("A reason is required for the audit history.")
            return
        if mode == "Add a record":
            edited = fill_missing_isins([edited], st.session_state.records)[0]
        elif not edited["ISIN"]:
            st.error("ISIN cannot be removed from an existing record.")
            return
        candidate = deepcopy(st.session_state.records)
        if before:
            candidate[next(i for i, row in enumerate(candidate) if row["ISIN"] == before["ISIN"])] = edited
        else:
            candidate.append(edited)
        issues = validate_record(edited)
        dataset_issues = validate_dataset(candidate)
        blockers = [issue for issue in issues if issue["severity"] == "Error"]
        duplicate_blockers = [
            issue
            for issue in dataset_issues
            if issue["severity"] == "Error"
            and (edited.get(issue["field"]) == issue["record"] or "Duplicate value" in issue["message"])
        ]
        if blockers or duplicate_blockers:
            st.error("This proposal contains blocking issues and was not submitted.")
            st.dataframe(blockers + duplicate_blockers, hide_index=True, use_container_width=True)
            return
        st.session_state.pending.append(
            {
                "id": f"CR-{len(st.session_state.pending) + len(st.session_state.history) + 1:04d}",
                "kind": "record",
                "title": "Add share class" if not before else "Edit share class",
                "before": deepcopy(before),
                "after": edited,
                "reason": reason,
                "actor": actor(),
                "created": now_utc(),
            }
        )
        st.success(f"Proposal submitted. ISIN: {edited['ISIN']}. A Reviewer can approve it from the review queue.")


def render_delete_request():
    labels = {
        f"{row.get('ISIN')} — {row.get('TPA Share Class ID') or row.get('Share Class', 'Unnamed')}": row
        for row in st.session_state.records
    }
    label = st.selectbox("Record to delete", list(labels), key="delete_record")
    reason = st.text_input("Deletion reason", key="delete_reason")
    confirm = st.checkbox("I confirm this is a proposal, not an immediate deletion.")
    if st.button("Submit deletion for review", type="secondary"):
        if not reason or not confirm:
            st.error("Provide a reason and confirm the proposal.")
            return
        record = labels[label]
        st.session_state.pending.append(
            {
                "id": f"CR-{len(st.session_state.pending) + len(st.session_state.history) + 1:04d}",
                "kind": "delete",
                "title": "Delete share class",
                "before": deepcopy(record),
                "after": {},
                "reason": reason,
                "actor": actor(),
                "created": now_utc(),
            }
        )
        st.success("Deletion proposal submitted.")


def quality_page():
    page_title(
        "Control checks",
        "Data quality",
        "Errors block proposed changes. Warnings highlight conditions that a reviewer should confirm.",
    )
    issues = validate_dataset(st.session_state.records)
    errors = [issue for issue in issues if issue["severity"] == "Error"]
    warnings = [issue for issue in issues if issue["severity"] == "Warning"]
    metric_columns = st.columns(3)
    metric_columns[0].metric("Total issues", len(issues))
    metric_columns[1].metric("Blocking errors", len(errors))
    metric_columns[2].metric("Warnings", len(warnings))
    severity = st.multiselect("Severity", ["Error", "Warning"], default=["Error", "Warning"])
    filtered = [issue for issue in issues if issue["severity"] in severity]
    if filtered:
        st.dataframe(filtered, use_container_width=True, hide_index=True)
    else:
        st.success("No data-quality issues match the selected severity.")


def import_export_page():
    page_title(
        "Workbook exchange",
        "Import and export",
        "Preview a workbook, generate missing ISINs, validate the dataset, and submit a controlled import.",
    )
    left, right = st.columns(2)
    with left:
        st.subheader("Download current data")
        st.download_button(
            "Download working dataset",
            data=write_workbook(st.session_state.records),
            file_name="Share_Class_Master_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with right:
        st.subheader("Download field template")
        st.download_button(
            "Download taxonomy template",
            data=TEMPLATE_FILE.read_bytes(),
            file_name="Share_Class_Field_Taxonomy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    if "import" not in permissions():
        st.info(f"Switch to Editor or Admin to import an Excel workbook.")
        return
    uploaded = st.file_uploader("Upload an .xlsx file with the corrected field headers", type=["xlsx"])
    if not uploaded:
        st.caption("The first sheet is read as the import source. All 37 field headers should be present.")
        return
    try:
        imported, headers = read_workbook(uploaded.getvalue(), with_headers=True)
        missing_headers = [field for field in FIELD_ORDER if field not in headers]
        unexpected_headers = [header for header in headers if header not in FIELD_ORDER]
        if missing_headers or unexpected_headers:
            st.error("The worksheet headers do not match the corrected taxonomy.")
            if missing_headers:
                st.write("Missing:", missing_headers)
            if unexpected_headers:
                st.write("Unexpected:", unexpected_headers)
            return
        if not imported:
            st.error("The worksheet contains no data rows.")
            return
        imported = fill_missing_isins(imported, st.session_state.records)
        issues = validate_dataset(imported)
        errors = [issue for issue in issues if issue["severity"] == "Error"]
        st.success(f"Loaded {len(imported)} records. Blank ISINs were allocated in the preview.")
        compact_table(imported)
        with st.expander(f"Validation results ({len(issues)} issues)"):
            if issues:
                st.dataframe(issues, use_container_width=True, hide_index=True)
            else:
                st.write("No validation issues found.")
        reason = st.text_input("Import reason", placeholder="Why is this dataset being submitted?")
        st.warning("Approval replaces the entire working dataset. This is clear and safe for a demo, but production imports should use a merge strategy.")
        if st.button("Submit validated import for review", disabled=bool(errors)):
            if not reason.strip():
                st.error("An import reason is required.")
            else:
                st.session_state.pending.append(
                    {
                        "id": f"CR-{len(st.session_state.pending) + len(st.session_state.history) + 1:04d}",
                        "kind": "import",
                        "title": "Replace workbook",
                        "before": deepcopy(st.session_state.records),
                        "after": imported,
                        "reason": reason,
                        "actor": actor(),
                        "created": now_utc(),
                    }
                )
                st.success("Import proposal submitted for review.")
    except (ValueError, KeyError, OSError) as exc:
        st.error(f"The workbook could not be read: {exc}")


def review_page():
    page_title(
        "Separation of duties",
        "Review queue",
        "Inspect proposed field-level changes, imports, and deletions before updating the working dataset.",
    )
    if "review" not in permissions():
        st.info("Switch to Reviewer or Admin to inspect and decide pending proposals.")
        st.caption(f"{len(st.session_state.pending)} pending demo proposal(s).")
        return
    if not st.session_state.pending:
        st.success("There are no pending proposals.")
        return
    for request in list(st.session_state.pending):
        with st.expander(f"{request['id']} — {request['title']} — {request['actor']}"):
            st.write(request["reason"])
            if request["kind"] == "record":
                changes = diff_records(request["before"], request["after"])
                st.dataframe(changes or [{"field": "", "before": "", "after": ""}], hide_index=True, use_container_width=True)
            elif request["kind"] == "import":
                st.write(
                    f"Replace {len(request['before'])} current records with "
                    f"{len(request['after'])} imported records."
                )
                compact_table(request["after"])
            else:
                compact_table([request["before"]])
            approve_column, reject_column = st.columns(2)
            reviewer_note = st.text_input("Decision note", key=f"note_{request['id']}")
            if approve_column.button("Approve", key=f"approve_{request['id']}", use_container_width=True):
                apply_request(request, "Approved", reviewer_note)
                st.rerun()
            if reject_column.button("Reject", key=f"reject_{request['id']}", use_container_width=True):
                if not reviewer_note.strip():
                    st.error("A decision note is required to reject a proposal.")
                else:
                    record_decision(request, "Rejected", reviewer_note)
                    st.rerun()


def apply_request(request, status, note):
    if request["kind"] == "import":
        st.session_state.records = deepcopy(request["after"])
    elif request["kind"] == "delete":
        target = request["before"]["ISIN"]
        st.session_state.records = [
            row for row in st.session_state.records if row.get("ISIN") != target
        ]
    else:
        incoming = request["after"]
        replaced = False
        for index, row in enumerate(st.session_state.records):
            if row.get("ISIN") == incoming.get("ISIN"):
                st.session_state.records[index] = deepcopy(incoming)
                replaced = True
                break
        if not replaced:
            st.session_state.records.append(deepcopy(incoming))
    st.session_state.data_version += 1
    record_decision(request, status, note)


def record_decision(request, status, note):
    st.session_state.history.append(
        {
            **deepcopy(request),
            "status": status,
            "reviewer": actor(),
            "decision": note,
            "decided": now_utc(),
        }
    )
    st.session_state.pending = [
        item for item in st.session_state.pending if item["id"] != request["id"]
    ]


def history_page():
    page_title(
        "Audit trail",
        "Change history",
        "Approved and rejected proposals, including actors, reasons, and decision notes.",
    )
    if not st.session_state.history:
        st.info("No decisions have been recorded in this session.")
        return
    summary = [
        {
            "Request": item["id"],
            "Action": item["title"],
            "Status": item["status"],
            "Submitted by": item["actor"],
            "Submitted": item["created"],
            "Reviewed by": item["reviewer"],
            "Reason": item["reason"],
            "Decision note": item["decision"],
        }
        for item in reversed(st.session_state.history)
    ]
    st.dataframe(summary, use_container_width=True, hide_index=True)
    st.download_button(
        "Download audit log",
        json.dumps(st.session_state.history, indent=2),
        file_name="share_class_audit_log.json",
        mime="application/json",
    )


def admin_page():
    page_title(
        "Configuration",
        "Demo administration",
        "Review role capabilities and reset the session to the genericized source workbook.",
    )
    if "admin" not in permissions():
        st.warning("Switch to Admin to view configuration actions.")
        return
    st.subheader("Demo role matrix")
    role_rows = [
        {
            "Role": role,
            "View": "Yes",
            "Export": "Yes",
            "Propose changes": "Yes" if "edit" in granted else "No",
            "Import": "Yes" if "import" in granted else "No",
            "Review": "Yes" if "review" in granted else "No",
            "Admin": "Yes" if "admin" in granted else "No",
        }
        for role, granted in ROLE_PERMISSIONS.items()
    ]
    st.dataframe(role_rows, hide_index=True, use_container_width=True)
    st.subheader("Production identity mapping")
    st.code(
        "Corporate identity: Microsoft Entra ID\n"
        "Allowed domain: @harbourvest.com\n"
        "Example identity: schauhan@harbourvest.com\n"
        "Recommended groups: HVP-ShareClass-Viewers, Editors, Reviewers, Admins\n"
        "No assigned group -> no access",
        language="text",
    )
    st.warning(
        "The sidebar role selector demonstrates permissions only. It is not authentication and "
        "must not be used with real or sensitive information."
    )
    confirm = st.checkbox("Reset this browser session to the original demo workbook.")
    if st.button("Reset demo data", disabled=not confirm):
        st.session_state.records = normalize_records(read_workbook(SOURCE_FILE))
        st.session_state.pending = []
        st.session_state.history = []
        st.session_state.data_version = 0
        st.success("Demo data reset.")


inject_style()
initialize_state()

with st.sidebar:
    st.markdown("## HVP  /  Master Data")
    st.caption("Share Class Data Manager")
    st.divider()
    st.markdown("### Demo identity")
    st.session_state.role = st.selectbox(
        "Role",
        list(ROLE_PERMISSIONS),
        index=list(ROLE_PERMISSIONS).index(st.session_state.role),
    )
    st.caption(actor())
    st.markdown(
        '<div class="demo-note">Demo role switching is not production security.</div>',
        unsafe_allow_html=True,
    )
    st.divider()
    page = st.radio(
        "Workspace",
        ["Overview", "Records", "Data quality", "Import / export", "Review queue", "Change history", "Admin"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption(f"Working records: {len(st.session_state.records)}")
    st.caption(f"Pending proposals: {len(st.session_state.pending)}")
    st.caption("Deployment target: Streamlit Community Cloud")

{
    "Overview": dashboard_page,
    "Records": records_page,
    "Data quality": quality_page,
    "Import / export": import_export_page,
    "Review queue": review_page,
    "Change history": history_page,
    "Admin": admin_page,
}[page]()

