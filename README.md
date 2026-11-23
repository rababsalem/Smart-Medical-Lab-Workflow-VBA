 # Automated Medical Lab Reporting System (Excel VBA)

An automated medical laboratory reporting workflow built in **Microsoft Excel** using **VBA (Visual Basic for Applications)**. It streamlines transferring patient records, recording timestamps, exporting PDF reports, printing documents, and triggering WhatsApp notifications.

## Overview

Preparing and distributing medical laboratory reports manually often slows down daily operations and increases the chance of human error. This project automates key steps of the reporting lifecycle inside an Excel macro-enabled workbook (`.xlsm`) so laboratory records can be processed consistently and efficiently.

## Key Features

* **Patient Record Transfer:** Moves patient details and test records between sheets/tables in a structured workflow.
* **Automatic Timestamps:** Logs date and time entries automatically for clear record tracking and auditing.
* **PDF Export:** Generates clean, ready-to-share or archive PDF report files.
* **Report Printing:** Prepares and sends reports to print using configured page/print settings.
* **WhatsApp Notifications:** Triggers WhatsApp notification alerts when reports are processed or ready.

## Repository Structure

* `data.xlsm` — Macro-enabled Excel workbook containing worksheets, print settings, visual elements, and the VBA project (`vbaProject.bin`).
* `vba-modules/` *(recommended)* — Plain-text exports of VBA modules (`.bas`, `.cls`, `.frm`) for easy code review on GitHub.
* `screenshots/` *(recommended)* — Images showing the workbook layout, exported PDF sample, or workflow output.

## Prerequisites & Setup

1. **Operating System & Software:** Microsoft Excel on Windows with macros supported.
2. **Enable Macros:** When opening `data.xlsm`, enable editing and enable macro content when prompted.
3. **Configure Paths & Settings:** Update any local folder paths, printer settings, or WhatsApp notification parameters in the VBA code to match your environment.
4. **Use Sample Data:** Ensure all patient records in the repository are dummy/sample data only.

## How to Use

1. Open `data.xlsm` in Microsoft Excel and enable macros.
2. Enter or select the test/patient record in the input sheet.
3. Run the VBA macro to transfer the record, stamp the time, export the PDF, print the report, and trigger the WhatsApp notification.

## Privacy & Data Notice

This repository is intended for demonstration and portfolio purposes only. Never upload real patient data, personal identifiable information (PII), or sensitive medical records. Always use synthetic or anonymized test data.
