# MSG to PDF Automation

This project contains a Power Automate Desktop (PAD) workflow developed to batch convert Outlook `.msg` email files into `.pdf` format securely and offline. Created during my internship with the Fair Political Practices Commission (FPPC), this solution is designed for environments where sensitive emails must be handled without cloud-based services.

## Project Summary

- **Purpose**: Automate the offline conversion of `.msg` files to `.pdf`.
- **Approach**: Built using Power Automate Desktop with minimal UI automation and reliable error handling.
- **Context**: Developed for FPPC to handle sensitive documentation securely.

## Repository Contents

- [`msg_to_pdf_flow.txt`](./msg_to_pdf_flow.txt) – Exported Power Automate Desktop flow in plain text format
- [`Batch MSG to PDF Automation – Documentation.pdf`](./Batch%20MSG%20to%20PDF%20Automation%20%E2%80%93%20Documentation.pdf) – Full PDF documentation of the project
- [`screenshots/`](./screenshots/) – Folder containing a visual overview of the PAD flow
- [`.gitignore`](./.gitignore) – Git ignore rules to keep the repo clean
- [`README.md`](./README.md) – You're reading it

## Requirements

- Windows 10 or 11
- Microsoft Outlook (signed in with a test or designated email account)
- Power Automate Desktop installed and configured

## Notes

- The folders for input, output, and failed files are configured as variables in the flow (e.g., `InputMSG`, `OutputPDF`, `FailedMSG`). These must be created locally before running the automation.
- The flow skips existing files and moves any failures to the designated folder for later review.
- Minimal UI automation is used (only to trigger the Outlook print dialog); most steps rely on keyboard input for reliability.

## Status

- Successfully converted 552 `.msg` files in a test run
- 1 failure, automatically handled by moving to the failure folder
- Total runtime: ~3 hours and 36 minutes

## Future Enhancements

- Add retry logic for failed `.msg` files
- Extract and log metadata (subject, sender, date) to a CSV
- Timestamped logging for auditing or reporting purposes

## Author

Niv Reddy  
IT Intern, Fair Political Practices Commission (FPPC)  
May 2025
