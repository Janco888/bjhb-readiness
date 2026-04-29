"""
BJHB Job Readiness Board — Streamlit App

Upload COOIS + MB52 (and optionally ZMPO), click Build, download the Excel.
"""

import io
import logging
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from scripts.build_readiness import (
    aggregate_to_jobs,
    annotate_with_pos,
    build_component_detail,
    build_how_it_works,
    build_readiness_board,
    build_stock_ledger,
    load_components,
    load_pos,
    load_stock,
    simulate_picks,
)

READINESS_BG = {"READY": "#C6EFCE", "PARTIAL": "#FFEB9C", "NOT READY": "#FFC7CE"}
COMPONENT_BG = {"COVERED": "#C6EFCE", "PARTIAL": "#FFEB9C", "SHORT": "#FFC7CE"}

st.set_page_config(page_title="BJHB Job Readiness", layout="wide")

st.title("BJHB Job Readiness Board")
st.caption(
    "Upload your SAP exports and click **Build Dashboard** to generate the weekly Excel report."
)

st.divider()

col1, col2, col3 = st.columns(3)
with col1:
    coois_file = st.file_uploader("COOIS Components *", type=["xlsx"])
with col2:
    mb52_file = st.file_uploader("MB52 Stock *", type=["xlsx"])
with col3:
    po_file = st.file_uploader("ZMPO Purchase Orders (optional)", type=["xlsx"])

st.divider()

ready_to_run = bool(coois_file and mb52_file)

if st.button("Build Dashboard", type="primary", disabled=not ready_to_run):
    log_lines = []

    class ListHandler(logging.Handler):
        def emit(self, record):
            log_lines.append(self.format(record))

    log = logging.getLogger("readiness_app")
    log.handlers.clear()
    handler = ListHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
    log.addHandler(handler)
    log.setLevel(logging.INFO)
    log.propagate = False

    try:
        with st.spinner("Loading data..."):
            today = pd.Timestamp.today().normalize()
            today_str = today.strftime("%d %b %Y")
            stock = load_stock(mb52_file, log)
            df_comp_raw = load_components(coois_file, today, log)
            df_pos = load_pos(po_file, log) if po_file else None

        if df_comp_raw.empty:
            st.error(
                "**No components to process after filtering.**\n\n"
                "Likely causes:\n"
                "- All components already fully withdrawn (nothing left to pick)\n"
                "- All job start dates are more than 30 days in the past\n"
                "- COOIS export is missing material codes\n\n"
                "Check that your COOIS export covers current/upcoming jobs "
                "and includes the *Quantity withdrawn* column."
            )
            st.stop()

        with st.spinner("Running simulation..."):
            df_comp = simulate_picks(df_comp_raw, stock, log)
            df_comp = annotate_with_pos(df_comp, df_pos, log)
            df_jobs = aggregate_to_jobs(df_comp, log)

        with st.spinner("Building workbook..."):
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "READINESS_BOARD"
            build_readiness_board(ws1, df_jobs, today_str)
            build_component_detail(wb.create_sheet("COMPONENT_DETAIL"), df_comp, df_jobs, today_str)
            build_stock_ledger(wb.create_sheet("STOCK_LEDGER"), df_comp, stock, today_str)
            build_how_it_works(wb.create_sheet("HOW_IT_WORKS"))
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

        st.session_state["df_jobs"] = df_jobs
        st.session_state["df_comp"] = df_comp

        n_ready = int((df_jobs["Readiness"] == "READY").sum())
        n_partial = int((df_jobs["Readiness"] == "PARTIAL").sum())
        n_not = int((df_jobs["Readiness"] == "NOT READY").sum())

        st.success(f"Dashboard built — {today_str}")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("READY", n_ready, help="Safe to release to workshop")
        c2.metric("PARTIAL", n_partial, help="Small shortage — planner decides")
        c3.metric("NOT READY", n_not, help="Blocked — do not release")
        c4.metric("Total Jobs", len(df_jobs))

        stamp = datetime.now().strftime("%Y-%m-%d")
        st.download_button(
            label="Download Excel Dashboard",
            data=out,
            file_name=f"Job_Readiness_Board_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as e:
        st.error(f"Build failed: {e}")

    if log_lines:
        with st.expander("Build log"):
            st.text("\n".join(log_lines))

elif not ready_to_run:
    st.info("Upload COOIS Components and MB52 Stock files above to enable the build.")


# ─── MO LOOKUP ──────────────────────────────────────────────────────────────
if "df_jobs" in st.session_state:
    st.divider()
    st.subheader("MO Lookup")

    df_jobs: pd.DataFrame = st.session_state["df_jobs"]
    df_comp: pd.DataFrame = st.session_state["df_comp"]

    search = st.text_input(
        "Search by MO number or job description",
        placeholder="e.g. 1234567 or pump",
    )

    if search:
        mask = (
            df_jobs["Order"].astype(str).str.contains(search, case=False, na=False)
            | df_jobs["Job_Desc"].astype(str).str.contains(search, case=False, na=False)
        )
        hits = df_jobs[mask].copy()
    else:
        hits = df_jobs.copy()

    # Summary table
    display = hits[[
        "Order", "Job_Desc", "SD_Order", "Start_Date", "Days_to_Start",
        "Readiness", "Components_Ready", "Components_Short", "Purchased_Short",
    ]].copy()
    display["Start_Date"] = display["Start_Date"].dt.strftime("%d %b %Y")

    def _highlight_readiness(row):
        bg = READINESS_BG.get(row["Readiness"], "")
        return [
            f"background-color: {bg}; font-weight: bold" if col == "Readiness" else ""
            for col in display.columns
        ]

    st.dataframe(
        display.style.apply(_highlight_readiness, axis=1),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Order": st.column_config.TextColumn("MO Number", width="small"),
            "Job_Desc": st.column_config.TextColumn("Job Description"),
            "SD_Order": st.column_config.TextColumn("SD Order", width="small"),
            "Start_Date": st.column_config.TextColumn("Start Date", width="small"),
            "Days_to_Start": st.column_config.NumberColumn("Days to Start", width="small"),
            "Readiness": st.column_config.TextColumn("Status", width="small"),
            "Components_Ready": st.column_config.NumberColumn("Parts Ready", width="small"),
            "Components_Short": st.column_config.NumberColumn("Parts Short", width="small"),
            "Purchased_Short": st.column_config.NumberColumn("Purchased Short", width="small"),
        },
    )
    st.caption(f"{len(hits)} job(s) shown")

    # Component breakdown for a selected MO
    st.markdown("#### Component Breakdown")
    mo_options = hits["Order"].tolist()

    if mo_options:
        selected_mo = st.selectbox(
            "Select MO",
            mo_options,
            format_func=lambda o: (
                f"{o}  —  "
                f"{hits.loc[hits['Order'] == o, 'Job_Desc'].values[0]}  "
                f"({hits.loc[hits['Order'] == o, 'Readiness'].values[0]})"
            ),
            label_visibility="collapsed",
        )

        job_row = df_jobs[df_jobs["Order"] == selected_mo].iloc[0]
        readiness = job_row["Readiness"]
        bg = READINESS_BG.get(readiness, "#f0f0f0")

        st.markdown(
            f'<div style="background:{bg};padding:10px 18px;border-radius:6px;'
            f'font-weight:bold;font-size:15px;margin:8px 0 12px 0">'
            f'MO {selected_mo} &nbsp;|&nbsp; {job_row["Job_Desc"]} &nbsp;|&nbsp; {readiness}'
            f'</div>',
            unsafe_allow_html=True,
        )

        comps = df_comp[df_comp["Order"] == selected_mo].copy()

        comp_cols = [
            "Material", "Material_Desc", "Proc_Type",
            "Required", "Withdrawn", "To_Pick",
            "Stock_At_Turn", "Allocated", "Short_Qty",
            "Component_Status",
        ]
        comp_display = comps[comp_cols].copy()

        if "PO_Delivery_Date" in comps.columns:
            comp_display["PO_Doc"] = comps["PO_Doc"].where(comps["PO_Doc"].notna(), "")
            comp_display["PO_Due"] = (
                comps["PO_Delivery_Date"].dt.strftime("%d %b %Y")
                .where(comps["PO_Delivery_Date"].notna(), "")
            )

        def _highlight_comp(row):
            bg = COMPONENT_BG.get(row["Component_Status"], "")
            return [
                f"background-color: {bg}; font-weight: bold" if col == "Component_Status" else ""
                for col in comp_display.columns
            ]

        st.dataframe(
            comp_display.style.apply(_highlight_comp, axis=1),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Material": st.column_config.TextColumn("Material", width="small"),
                "Material_Desc": st.column_config.TextColumn("Description"),
                "Proc_Type": st.column_config.TextColumn("Type", width="small"),
                "Required": st.column_config.NumberColumn("Required", format="%.2f", width="small"),
                "Withdrawn": st.column_config.NumberColumn("Withdrawn", format="%.2f", width="small"),
                "To_Pick": st.column_config.NumberColumn("To Pick", format="%.2f", width="small"),
                "Stock_At_Turn": st.column_config.NumberColumn("Stock Available", format="%.2f", width="small"),
                "Allocated": st.column_config.NumberColumn("Allocated", format="%.2f", width="small"),
                "Short_Qty": st.column_config.NumberColumn("Short Qty", format="%.2f", width="small"),
                "Component_Status": st.column_config.TextColumn("Status", width="small"),
                "PO_Doc": st.column_config.TextColumn("PO Number", width="small"),
                "PO_Due": st.column_config.TextColumn("PO Due Date", width="small"),
            },
        )
    else:
        st.info("No jobs match your search.")
