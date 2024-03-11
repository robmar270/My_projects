# Project Monthly Reviews

# Introduction
On my current internship I did a solution for a project that it's purpose was to create an automation ETL for our montly repports that's gonna be send out by email. So the only thing I had to consider was that they wanted the data to always be stored in the google spreadsheet, so the sales team could always have access to it.

## Project Overview

This script is designed to fetch data from Google Sheets using the Google Sheets API, perform data manipulation and merging using Pandas, and generate personalized HTML email templates using Jinja2. The emails are sent with information obtained from the merged data, and the script incorporates an image fetched from GitHub.

### Key Components:

- **Google Sheets API Integration:** Fetches data from specified Google Sheets.
- **Data Processing:** Utilizes Pandas for data manipulation and merging.
- **Dynamic HTML Email Templates:** Jinja2.
- **GitHub Image Integration:** Fetches an image from a GitHub repository for email content.
- **`MIMEMultipart`:** Represents a MIME multipart message, organizing different types of content.
- **`MIMEText`:** Represents plain text or HTML content in a MIME message. Used to attach HTML content to emails.
- **`MIMEImage`:** Represents an image attachment in a MIME message. Used to attach images with the 'cid:' scheme.




