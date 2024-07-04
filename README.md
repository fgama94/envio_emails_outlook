# Email Sender Program for Outlook



# Overview

This Python program automates the process of sending personalized emails using data from an Excel spreadsheet. It utilizes the win32com.client library for interfacing with Outlook, pandas for data manipulation, and BeautifulSoup for HTML parsing.



# Features

- Reads email data from an Excel file with multiple sheets (Emails, PT, EN, ES).

- Validates required fields in each sheet to ensure data completeness.

- Checks for the existence of attachments specified in the spreadsheet.

- Dynamically generates email content based on recipient details and language preferences.

- Handles duplicate emails and provides logging for errors and successful operations.



# Prerequisites

Ensure the following before running the program:

- Microsoft Outlook must be running.

- Excel file (Envio_Emails.xlsx) with specific sheet names (Emails, PT, EN, ES) and required columns.

- Attachments for emails must be located in the Anexos folder.



# Usage

1. Setup: Save the Excel file (Envio_Emails.xlsx) and ensure the Anexos folder contains all required attachments.

2. Execution: Run the program and follow on-screen prompts to send emails.

3. Logs: Check generated log files (Relatório_<timestamp>.log) for any errors or notifications.



# Installation

1. Clone the repository:

        git clone https://github.com/fgama94/envio_emails_outlook.git
        cd envio_emails_outlook

2. Ensure required Python libraries are installed:

        pip install pandas beautifulsoup4



# Troubleshooting

Missing Attachments: Ensure all specified attachments exist in the Anexos folder.

Incomplete Data: Check the Excel file for missing or incomplete data in required fields.
