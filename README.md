# Automated Hospital Survey Data Processor

This project is a Python automation built for a hospital-run Narrative Medicine program to process participant survey data and consolidate it into a single master spreadsheet.

It replaces a fully manual workflow where staff previously had to review pre-surveys, post-surveys, and attendance logs and copy results by hand, one participant at a time.

---

## Features

- **Automated processing of pre-survey, post-survey, and attendance data**
- **Consolidation of multiple Google Sheets into a single master sheet**
- **Conversion of word-based survey responses into standardized numeric values**
- **Automatic attendance counting from repeated identifier entries**
- **Resilient to missing, partial, or incomplete survey responses**
- **Handles spreadsheets with columns in different orders**
- **Fuzzy column matching to tolerate minor header wording changes**
- **Idempotent design allowing safe re-runs without overwriting existing data**

---

## How It Works

The script connects to Google Sheets using the Google Sheets API and reads data from four sources: a pre-survey export, a post-survey export, an attendance log, and a pre-existing master sheet.

Participant identifiers are used to align rows across all sources. Survey responses are mapped from human-readable text (e.g. “Strongly Agree”, “PGY2”, “Yes”) into numeric codes, attendance is calculated based on repeated entries, and all results are written into the correct rows and columns of the master sheet in a single run.

---

## Handling Real-World Data Issues

This project was designed around the realities of messy, human-generated data. It accounts for situations where:

- Participants complete only a pre-survey or post-survey
- Participants attend sessions without completing surveys
- Survey columns appear in different orders between exports
- Headers change slightly between form versions
- Responses are missing or partially filled out

Rather than assuming ideal inputs, the script uses defensive checks and adaptive column matching to ensure correct alignment whenever possible.

---

## Impact

Before this tool existed, hospital staff spent an estimated **30–40 hours** manually reviewing surveys and updating the master spreadsheet.

This automation reduced that workflow to **approximately two minutes**, while also improving consistency and reducing human error.

---

## Technologies Used

- **Python**
- **Google Sheets API**
- **gspread**
- **OAuth2 Service Accounts**

---
