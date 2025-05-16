# MSG to PDF Automation

This project contains a Power Automate Desktop (PAD) workflow developed to batch convert Outlook `.msg` email files into `.pdf` format securely and offline. Created during my internship with the Fair Political Practices Commission (FPPC), this solution is designed for environments where sensitive emails must be handled without cloud-based services.

## Project Summary

- **Purpose**: Automate the offline conversion of `.msg` files to `.pdf`.
- **Approach**: Built using Power Automate Desktop with minimal UI automation and reliable error handling.
- **Context**: Developed for FPPC to handle sensitive documentation securely.

## Repository Contents

- **msg_to_pdf_flow.txt**: The full Power Automate Desktop flow exported in text format.
- **Batch MSG to PDF Automation – Documentation.pdf**: Complete documentation of the flow, including motivation, architecture, key features, results, and future plans.
- **screenshots/**: Contains a visual overview of the PAD flow (e.g., `pad_flow.png`).
- **.gitignore**: Configuration to exclude system and temporary files.
- **README.md**: You're reading it.

## Requirements

- Windows 10 or 11
- Microsoft Outlook (signed in with a test or designated email account)
- Power Automate Desktop installed and configured

## Notes

- The actual folders for input and output (`InputMSG`, `OutputPDF`, `FailedMSG`) are variables defined inside the flow. You’ll need to create these locally before running the automation.
- The flow is built to run unattended, skipping duplicates and logging any failed conversions by moving them to a specified folder.
- Minimal UI automation is used to maximize reliability — only the Outlook print command uses UI elements.

## Status

- Successfully converted 552 `.msg` files in a test run
- Only 1 failure, which was automatically handled
- Total runtime: ~3 hours and 36 minutes

## Future Enhancements

- Add retry logic for failed files
- Implement CSV logging of metadata (subject, sender, date)
- Improve notification and logging visibility

## Author

Niv Reddy  
IT Intern, Fair Political Practices Commission (FPPC)  
May 2025
