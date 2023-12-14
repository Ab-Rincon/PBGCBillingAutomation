# Invoice Processing Application
### Version 1.0.4

# Invoice Processing Instructions

Follow these steps to process your invoice workbook using the provided application.

## Getting Started

1. **Prepare Invoice Workbook:**
   - Locate the `dist/input` directory on your computer.
   - Place your invoice workbook (`.xlsx` or `.xls` file) into the `dist/input` directory.
   - Please ensure that all the necessary headers are included in the file.

2. **Run the Application:**
   - Navigate to the `dist` directory.
   - Double-click on `main.exe` to start the processing of the invoice workbook.

## Post-Processing

1. **Check Output:**
   - After the application has run, open the `dist/output` directory.
   - Retrieve the processed invoice workbook from this directory.

2. **Review Log Files:**
   - Find `logfile.log` in the `dist` directory to review any operational logs for errors or confirmation of successful processing.
   - If you experience any issues, the log file may contain details that can help in troubleshooting.

## Requirements
 - It's necessary to have excel installed prior to running the application.

## Developer Section
 - Ensure the version is updated in `.github/workflows/ci.yml` and `ReadMe.md` when creating new versions

### To-Do List
 - Rename overbilled section: time mistmatch

# Notes
- Comments are removed from the Summary's sheet when flagged for error, which can only occur once per comment. If a problem sheet exists for the WBS, the summary sheet will be incomplete unless manually updated or the comments are fixed prior to the automation running once more.
- Format issues include but are not limited to:
   - Disorganized time stamps not following the format "Time in: HH:MM - Time Out: HH:MM." There is some leniency but that is the expected format.
   - Mismatched counts of time in and time out entries.
   - Time stamps not ending with 00, 15, 30, or 45 minutes.
- Military time issues:
   - "Time out" recorded before "Time in" within the same day.
   - Failures in time worked calculations due to invalid military time entries. Ex. 12:62
   - An out time of "00:00" or "24:00" though not a valid military time will be treated like "23:59" and have special handling to make it a valid time out comment
- Conflicting daily comments entail discrepancies in time entries for the same day.
- Comment mismatch concerns:
   - Situations where the total billed hours exceeds the amounts discussed in comments are noted in the mismatch worksheet.
   - Items in the Summary sheet include flagged mismatch instances.
- Worksheet Management:
   - Detail, problem, and overbilling sheets are deleted when no issues are present.
   - A clean workbook will contain only summary sheets after processing.
