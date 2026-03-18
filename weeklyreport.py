#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jul 11 18:10:42 2024

@author: khosravip2

This script is set to generate weekly reporting for Dr. Pine

REMINDERS:
IRTAs remember to enter your visits to the cumulative report 
IRTAs please share the output weekly with Dr. Pine 

"""

#!/usr/bin/env python3

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import base64
from io import BytesIO

import jinja2
import matplotlib.pyplot as plt
import pandas as pd


# Created lists for : Subject types , task names, visit types, location of visit, and consent options

SUBJECT_TYPE = [
    "Anxious Kids",
    "Depressed Kids",
    "Healthy Kids",
    "Healthy Adults",
    "ASD Kids",
]

TASK_CATEGORY_MAP = {
    "Anti-Saccade": "Eyetracking",
    "Dwell": "Eyetracking",
    "GCMRT-Music": "Eyetracking",
    "YIKES": "Eyetracking",
    "Doors": "Eyetracking",
    "Scream": "Pain Tasks",
    "EMA (CTADS)": "Other",
    "New EMA (Anx/Irritability)": "Other",
    "Flanker Behavior": "Other",
    "AXCPT": "Other",
    "STOP": "Other",
    "Memory": "Other",
    "Forms at Home": "Other",
    "Forms at Clinic": "Other",
    "CRDM": "Other",
    "4-Min Speech Sample": "Other",
    "ratings": "Other",
    "DECIBELS baseline": "Other",
}

TASK_ORDER = [
    "GCMRT-Music",
    "Dwell",
    "Anti-Saccade",
    "EMA (CTADS)",
    "New EMA (Anx/Irritability)",
    "STOP",
    "Flanker Behavior",
    "AXCPT",
    "Memory",
    "Forms at Home",
    "Forms at Clinic",
    "PCDT",
    "4-Min Speech Sample",
    "CRDM",
    "ratings",
]

TASK_CATEGORY_ORDER = ["Eyetracking", "Pain Tasks", "Other", "Scan"]


CONSENT_OPTIONS = ["Opt IN", "Opt OUT"]

CUMULATIVE_TASK_ORDER = [
    "EMA (CTADS)",
    "New EMA (Anx/Irritability)",
    "STOP",
    "Flanker Behavior",
    "AXCPT",
    "Memory",
    "Forms at Home",
    "Forms at Clinic",
    "CRDM",
    "PCDT",
    "4-Min Speech Sample",
    "Other",
    "AXCPT/Flanker4 (fMRI) Pre",
    "AXCPT/Flanker4 (fMRI) Post",
    "AXCPT/Flanker4 (MEG) Pre",
    "AXCPT/Flanker4 (MEG) Post",
    "TAU3.0 Pre",
    "TAU3.0 Post",
    "AXCPT/Flanker3 (fMRI) (T1)",
    "AXCPT/Flanker3 (fMRI) (T2)",
    "AXCPT/Flanker3 (MEG)",
    "Movies-Sintel + Francis",
    "TMS",
    "TMS-kids",
    "Carnival",
    "ratings",
    "DECIBELS baseline",
    "DECIBELS MRI+cTBS",
    "DECIBELS MEG+cTBS",
]

SCAN_TASK_ORDER = [
    "AXCPT/Flanker4 (fMRI) Pre",
    "AXCPT/Flanker4 (fMRI) Post",
    "AXCPT/Flanker4 (MEG) Pre",
    "AXCPT/Flanker4 (MEG) Post",
    "TAU3.0 Pre",
    "TAU3.0 Post",
    "Movies-Sintel + Francis",
    "TMS",
    "TMS-kids",
    "Carnival",
    "DECIBELS MRI+cTBS",
    "DECIBELS MEG+cTBS",
]

EYETRACKING_TASK_ORDER = [
    "Anti-Saccade",
    "Dwell",
    "GCMRT-Music",
    "YIKES",
    "Doors",
]

@dataclass
class ReportConfig:
    workbook_path: Path
    tracker_sheet: str = "Tracker"
    irta_sheet: str = "IRTA Subject Count"
    output_html: Path = Path("/Users/khosravip2/Documents/weekly_report.html")


def load_data(config: ReportConfig) -> tuple[pd.DataFrame, pd.DataFrame]:
    tracker = pd.read_excel(config.workbook_path, sheet_name=config.tracker_sheet)
    irta = pd.read_excel(config.workbook_path, sheet_name=config.irta_sheet)

    tracker["DATE"] = pd.to_datetime(tracker["DATE"], errors="coerce")
    tracker = tracker.sort_values("DATE").copy()

    return tracker, irta


def get_reporting_window(reference_date: pd.Timestamp | None = None) -> tuple[pd.Timestamp, pd.Timestamp]:
    today = pd.Timestamp.today().normalize() if reference_date is None else pd.Timestamp(reference_date).normalize()

    # Friday through Thursday style window
    weekday = today.weekday()
    last_friday = today - pd.Timedelta(days=(weekday - 4) % 7)
    return last_friday, today


def summarize_visits(data: pd.DataFrame, subject_type: list[str]) -> pd.DataFrame:
    VISIT_ORDER = ["V1", "V2", "Research Only", "Tx", "Tx-GCMRT"]

    summary = data.pivot_table(
        index="SUBJECTTYPE",
        columns="VISITTYPE",
        aggfunc="size",
        fill_value=0,
    )

    # Force columns to exist
    summary = summary.reindex(columns=VISIT_ORDER, fill_value=0)

    summary = summary.reindex(subject_type, fill_value=0)
    summary["Total"] = summary.sum(axis=1)
    summary.loc["Total"] = summary.sum(axis=0)

    return summary


def summarize_tasks(data: pd.DataFrame, task_order: list[str], task_map: dict[str, str]) -> pd.DataFrame:
    melted = data.melt(
        id_vars=["SUBJECTTYPE"],
        value_vars=["TASK1", "TASK2", "TASK3", "TASK4"],
        var_name="TASK_SLOT",
        value_name="TASK",
    )

    melted = melted.dropna(subset=["TASK"]).copy()
    melted["CATEGORY"] = melted["TASK"].map(task_map).fillna("Other")

    summary = melted.pivot_table(
        index="TASK",
        columns="CATEGORY",
        aggfunc="size",
        fill_value=0,
    )

    summary = summary.reindex(task_order, fill_value=0)
    summary["Total"] = summary.sum(axis=1)
    summary.loc["Total"] = summary.sum(axis=0)
    return summary


def summarize_irta(irta: pd.DataFrame) -> pd.DataFrame:
    active = irta[irta["STATUS"] == "Active"]
    active_counts = active.groupby(["IRTA", "SUBJECTTYPE"]).size().unstack(fill_value=0)

    summary = active_counts.copy()
    summary["Total"] = summary.sum(axis=1)
    return summary


def monthly_completed_counts(data: pd.DataFrame) -> pd.Series:
    completed = data[data["SESSIONSTATUS"] == "Completed"].copy()
    completed["DATE"] = pd.to_datetime(completed["DATE"], errors="coerce")

    monthly = (
        completed.assign(MONTH=completed["DATE"].dt.to_period("M").astype(str))
        .groupby("MONTH")
        .size()
        .sort_index()
    )
    return monthly


def create_monthly_plot_base64(monthly_counts: pd.Series) -> str:
    labels = monthly_counts.index.tolist()
    values = monthly_counts.values.tolist()

    fig, ax = plt.subplots(figsize=(10, 3))
    ax.plot(labels, values, marker="o")
    ax.set_title("Completed Sessions Over Time")
    ax.set_ylabel("Completed Sessions")
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=45, ha="right")
    plt.tight_layout()

    buf = BytesIO()
    plt.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)

    return base64.b64encode(buf.getvalue()).decode("utf-8")


######################################################
###### CUMULATIVE REPORT SINCE 05.31.2024 ######
######################################################

def summarize_consent(irta: pd.DataFrame, column_name: str, subject_type: list[str]) -> pd.DataFrame:
    summary = irta.pivot_table(
        index="SUBJECTTYPE",
        columns=column_name,
        aggfunc="size",
        fill_value=0,
    )
    summary = summary.reindex(subject_type, fill_value=0)
    summary.loc["Total"] = summary.sum(axis=0)
    return summary


def summarize_consenttim(irta: pd.DataFrame, subject_type: list[str], consent_order: list[str]) -> pd.DataFrame:
    summary = irta.pivot_table(
        index="SUBJECTTYPE",
        columns="CONSENTTIM",
        aggfunc="size",
        fill_value=0,
    )
    summary = summary.reindex(subject_type, fill_value=0)
    summary = summary.T.reindex(consent_order, fill_value=0).T
    summary.loc["Total"] = summary.sum(axis=0)
    return summary


def melt_tasks(data: pd.DataFrame) -> pd.DataFrame:
    melted = data.melt(
        id_vars=["SUBJECTTYPE"],
        value_vars=["TASK1", "TASK2", "TASK3", "TASK4"],
        var_name="TASK_SLOT",
        value_name="TASK",
    )
    return melted.dropna(subset=["TASK"]).copy()


def summarize_cumulative_tasks_by_subject(
    melted_tasks: pd.DataFrame,
    task_order: list[str],
    subject_type: list[str],
) -> pd.DataFrame:
    summary = melted_tasks.pivot_table(
        index="TASK",
        columns="SUBJECTTYPE",
        aggfunc="size",
        fill_value=0,
    )
    summary = summary.reindex(task_order, fill_value=0)
    summary = summary.T.reindex(subject_type, fill_value=0)
    summary["Total"] = summary.sum(axis=1)
    summary.loc["Total"] = summary.sum(axis=0)
    return summary


def summarize_selected_tasks_by_subject(
    melted_tasks: pd.DataFrame,
    selected_tasks: list[str],
    subject_type: list[str],
) -> pd.DataFrame:
    summary = melted_tasks[melted_tasks["TASK"].isin(selected_tasks)].pivot_table(
        index="TASK",
        columns="SUBJECTTYPE",
        aggfunc="size",
        fill_value=0,
    )
    summary = summary.reindex(selected_tasks, fill_value=0)
    summary = summary.T.reindex(subject_type, fill_value=0)
    summary["Total"] = summary.sum(axis=1)
    summary.loc["Total"] = summary.sum(axis=0)
    return summary

def render_report(
    output_path: Path,
    weekly_completed_visits: pd.DataFrame,
    weekly_canceled_visits: pd.DataFrame,
    weekly_completed_tasks: pd.DataFrame,
    irta_summary: pd.DataFrame,
    cumulative_completed_visits: pd.DataFrame,
    cumulative_completed_tasks: pd.DataFrame,
    cumulative_consent_movies: pd.DataFrame,
    cumulative_consent_tim: pd.DataFrame,
    cumulative_scan_tasks: pd.DataFrame,
    cumulative_eyetracking_tasks: pd.DataFrame,
    plot_base64: str,
    start_date: pd.Timestamp,
    end_date: pd.Timestamp,
) -> None:
    
    # Define the template for the HTML report
    template_text = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="utf-8">
        <title>Weekly Reporting</title>
        <style>
            body { font-family: Aptos, sans-serif;; margin: 24px; }
            h1, h2 { margin-bottom: 8px; }
            table { border-collapse: collapse; margin-bottom: 24px; width: 70%; font-size: 13px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background: #f2f2f2; }
            img { max-width: 100%; height: auto; }
            .spaced { margin-top: 20px; }
            .page-break { page-break-before: always; }
        </style>
    </head>
    <body>
        <h1>Weekly Reporting</h1>
        <p><strong>Reporting window:</strong> {{ start_date }} to {{ end_date }}</p>

        <h2>Completed Sessions Trend</h2>
        <img src="data:image/png;base64,{{ plot_base64 }}" alt="Completed sessions trend">

        <h2>Completed Visits</h2>
        {{ weekly_completed_visits_html | safe }}

        <h2>Canceled Visits</h2>
        {{ weekly_canceled_visits_html | safe }}

        <h2>Completed Tasks</h2>
        {{ weekly_completed_tasks_html | safe }}

        <h2>IRTA Active Subject Totals</h2>
        {{ irta_summary_html | safe }}

        <div class="page-break"></div>

        <h1>Cumulative Reporting</h1>

        <h2>Cumulative Completed Visits</h2>
        {{ cumulative_completed_visits_html | safe }}

        <h2>Cumulative Completed Tasks</h2>
        {{ cumulative_completed_tasks_html | safe }}

        <h2>Consented Movies</h2>
        {{ cumulative_consent_movies_html | safe }}

        <h2>Consented TIM</h2>
        {{ cumulative_consent_tim_html | safe }}

        <h2>Cumulative Scan Tasks</h2>
        {{ cumulative_scan_tasks_html | safe }}

        <h2>Cumulative Eyetracking Tasks</h2>
        {{ cumulative_eyetracking_tasks_html | safe }}
        <div class="Footer" style="font-size: 15px; padding: 8x 0;">
            <i>Report Generated By P.Khosravi</i>
    </body>
    </html>
    """

    template = jinja2.Template(template_text)
    html = template.render(
        start_date=start_date.strftime("%Y-%m-%d"),
        end_date=end_date.strftime("%Y-%m-%d"),
        plot_base64=plot_base64,
        weekly_completed_visits_html=weekly_completed_visits.to_html(),
        weekly_canceled_visits_html=weekly_canceled_visits.to_html(),
        weekly_completed_tasks_html=weekly_completed_tasks.to_html(),
        irta_summary_html=irta_summary.to_html(),
        cumulative_completed_visits_html=cumulative_completed_visits.to_html(),
        cumulative_completed_tasks_html=cumulative_completed_tasks.to_html(),
        cumulative_consent_movies_html=cumulative_consent_movies.to_html(),
        cumulative_consent_tim_html=cumulative_consent_tim.to_html(),
        cumulative_scan_tasks_html=cumulative_scan_tasks.to_html(),
        cumulative_eyetracking_tasks_html=cumulative_eyetracking_tasks.to_html(),
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")

def main() -> None:
    config = ReportConfig(
        workbook_path=Path(
            "/Users/khosravip2/Library/CloudStorage/OneDrive-SharedLibraries-NationalInstitutesofHealth/Emotion and Development Branch - Documents/CumulativeList_FridayAgenda.xlsx"
        )
    )
    tracker, irta = load_data(config)

    start_date, end_date = get_reporting_window()
    weekly = tracker[(tracker["DATE"] >= start_date) & (tracker["DATE"] <= end_date)]

    weekly_completed = weekly[weekly["SESSIONSTATUS"] == "Completed"].copy()
    weekly_canceled = weekly[weekly["SESSIONSTATUS"] == "Canceled"].copy()

    completed_visits = summarize_visits(weekly_completed, SUBJECT_TYPE)
    canceled_visits = summarize_visits(weekly_canceled, SUBJECT_TYPE)
    completed_tasks = summarize_tasks(weekly_completed, TASK_ORDER, TASK_CATEGORY_MAP)
    irta_summary = summarize_irta(irta)

    monthly_counts = monthly_completed_counts(tracker)
    plot_base64 = create_monthly_plot_base64(monthly_counts)

    # CUMULATIVE
    cumulative_completed = tracker[tracker["SESSIONSTATUS"] == "Completed"].copy()

    cumulative_completed_visits = summarize_visits(cumulative_completed, SUBJECT_TYPE)

    melted_cumulative_tasks = melt_tasks(cumulative_completed)

    cumulative_completed_tasks = summarize_cumulative_tasks_by_subject(
        melted_tasks=melted_cumulative_tasks,
        task_order=CUMULATIVE_TASK_ORDER,
        subject_type=SUBJECT_TYPE,
    )

    cumulative_consent_movies = summarize_consent(
        irta=irta,
        column_name="CONSENTMOVIES",
        subject_type=SUBJECT_TYPE,
    )

    cumulative_consent_tim = summarize_consenttim(
        irta=irta,
        subject_type=SUBJECT_TYPE,
        consent_order=CONSENT_OPTIONS,
    )

    cumulative_scan_tasks = summarize_selected_tasks_by_subject(
        melted_tasks=melted_cumulative_tasks,
        selected_tasks=SCAN_TASK_ORDER,
        subject_type=SUBJECT_TYPE,
    )

    cumulative_eyetracking_tasks = summarize_selected_tasks_by_subject(
        melted_tasks=melted_cumulative_tasks,
        selected_tasks=EYETRACKING_TASK_ORDER,
        subject_type=SUBJECT_TYPE,
    )

    render_report(
        output_path=config.output_html,
        weekly_completed_visits=completed_visits,
        weekly_canceled_visits=canceled_visits,
        weekly_completed_tasks=completed_tasks,
        irta_summary=irta_summary,
        cumulative_completed_visits=cumulative_completed_visits,
        cumulative_completed_tasks=cumulative_completed_tasks,
        cumulative_consent_movies=cumulative_consent_movies,
        cumulative_consent_tim=cumulative_consent_tim,
        cumulative_scan_tasks=cumulative_scan_tasks,
        cumulative_eyetracking_tasks=cumulative_eyetracking_tasks,
        plot_base64=plot_base64,
        start_date=start_date,
        end_date=end_date,
    )


if __name__ == "__main__":
    main()
