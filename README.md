# Weekly Report

The script generates weekly and overall reports from participant tracking data for all the projects within a research branch. It reads information from formatted Excel trackers that is on a shared drive, sums sessions by visit status and task type, as well as patient counts for each research assistances and constant numbers and generates a formatted HTML report from Jinja2 templates and matplotlib plots.

Developed to assist clinical research staff and labs efficiently summarize participation trends, task rates, and group-level data.

### Features of this script: ###

* generates a weekly summary of completed and canceled visits by participant group (e.g., HV, Anxious, MDD etc.) and visit type (e.g., V1, V2, research visit, etc). 
* Groups task by categories (e.g., Eyetracking, Scan, Behavioral, etc)
* It provides cumulative participantion reporting (i.e., data collected thus far for each task, and participant group)
* Automatically generates an interactive HTML output with pre-defined page separation for better pdf reporting if needed
* It is modular and adapatable to new task types or visit types


### Inputs: ###
* Excel workbook with cumulative participant/session information saved and editable on a shared drive

* Sheet 1: Visit/timing/task type and status data

* Sheet 2: Subject-level metadata (status, consent, assigned research assistants - IRTAs)

### Outputs: ###
* HTML report with structured tables and session plots *(It has set page breaks in case needs tobe saved as pdf)*

* PNG chart of completed sessions over time

### Needs the following packages for _python 3.8+_
```
pandas
jinja2 *#for HTML templating*
matplotlib
openpyxl *#for Excel parsing*
```
_This script can be adapated to generate reports for specific date range (calculating based on week numbers and the current day), patient type, task type/categories, visit types, and status of these visits (completed, canceled), cumulative reporting of collected date by task type and patient type. The script can also be adapted for its utlizatio of HTML reporting._
