from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd
import os
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from jinja2 import Template
from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd
from jinja2 import Template
import requests

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "keys.json"

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

# Spreadsheet ID
SPREADSHEET_ID = "1LgsSEzw0KfmUqvQNqTLRKGaEq2swM9ZjCMzGRQxxFiY"

# Build the Google Sheets API
service = build("sheets", "v4", credentials=creds)


# Function to fetch data from a sheet
def fetch_data(sheet_name):
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}")
        .execute()
    )
    return result.get("values", [])


# Fetch data from the 2024 sheet
data_2024 = fetch_data("2024")

# Fetch data from the email sheet
data_email = fetch_data("email")

# Convert data to Pandas DataFrames
df_2024 = pd.DataFrame(data_2024[1:], columns=data_2024[0])
df_email = pd.DataFrame(data_email[1:], columns=data_email[0])

# Merge DataFrames based on 'metric_id'
merged_df = pd.merge(
    df_2024, df_email[["email", "metric_id"]], on="metric_id", how="left"
)

# Print the final merged DataFrame
# print(merged_df)
# ... (previous code)

# Assuming merged_df has columns like 'email', 'metric_id', and other relevant columns

# Group the merged DataFrame by 'metric_id'
merged_data_dict = merged_df.to_dict(orient="records")
print(merged_data_dict)

### GET LOGO FROM GITHUB AND INSERT IT INTO THE HTML

# Replace 'YOUR_TOKEN' with your actual Personal Access Token
github_token = "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"

# GitHub API URL for raw content
github_raw_url = "https://raw.githubusercontent.com/robmar270/Vintly/df15b612bbdef6397318c2d867b98244f5763cd9/Vintly_logo_edit.png"

# Fetch the image from GitHub with the Personal Access Token in the headers
response = requests.get(github_raw_url, headers={"Authorization": f"token {github_token}"})

# Check if the request was successful (status code 200)
if response.status_code == 200:
    # Get the image content
    image_content = response.content


# HTML template
html_body_template = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
   body {
      font-family: 'Open Sauce Sans', sans-serif;
      background-color: #f4f5f6db;
      margin: 0;
      padding: 0;
    }

    .email-container {
      width: 100%;
      max-width: 600px; /* Adjust as needed for your email layout */
      margin: auto;
      padding: 24px;
      background: #958fb0;
      border: 1px solid #eaebed;
      border-radius: 16px;
      text-align: center;
    }

    .logo img {
      max-width: 100%;
      height: auto;
      text-align: center;
    }

    .introduction {
      text-align: center;
      margin-bottom: 40px;
    }

    .introduction h1 {
      color: #1f2a37;
      font-family: 'Open Sauce Sans', sans-serif;
      font-size: 42px;
      margin-bottom: 10px;
    }

    .introduction h2 {
      color: #1f2a37;
      font-family: 'Open Sauce Sans', sans-serif;
      font-size: 22px;
      margin-top: 10px;

    }

    .introduction h3 {
      color: #1f2a37;
      font-family: 'Open Sauce Sans', sans-serif;
      font-size: 18px;      
    }

    .first-section{
      background-color: #958fb0;
      display: flex;
      justify-content: space-between;
    }

    .box1, .box2 {
      background-color: #eaebed;
      border-radius: 5px;
      width: 280px;
      height: 125px;
      margin: 5px auto;
      box-shadow: 0 7px 1px rgba(0, 0, 0, 0.1);
      height: auto;
      padding-bottom: 10px;
    }
    .box1 h4, .box2 h4 {
      text-align: center;
      font-size: 20px;
      margin-top: 15px;
    }


    .box1 p, .box2 p {
      text-align: center;
      font-size: 18px;
    }


    .second-section{
      background-color: #958fb0;
      display: flex;
      justify-content: space-between;
      padding-top: 2%;
    }
    .box3 {
      background-color: #eaebed;
      border-radius: 5px;
      width: 590px;
      height: auto;
      padding-bottom: 10px;
      align-items: center;
      margin: auto;  
      margin: 5px auto;
      box-shadow: 0 7px 1px rgba(0, 0, 0, 0.1);
      line-height: -1; /* Adjust the value to make the text closer */

    }

    .box3 h4 {
      text-align: center;
      font-size: 20px;
      margin-bottom: 10px;
      margin-top: 15px;
    }

    .box3 p{
    text-align: center;
    font-size: 18px;
    }
    .box3 div {
    text-align: center;
    margin-top: 5px;
    font-size: 16px;
    }

    .logo-lines {
      border-top: 2px solid #1f2a37;
      border-bottom: 2px solid #1f2a37;
      margin: 20px 0;
      border-radius: 2px;
    }

</style>
</head>
<body>
<!-- ... (previous template code) -->
<div class="email-container">
    <div class="logo-section logo-lines">
        <div class="logo">
            <img src="cid:logo_image" alt="Company Logo">
        </div>
    </div>    
        {% for metric_id, metric_entries in merged_data_dict|groupby('metric_id') %}
            <div class="metric-container" id="metric_{{ metric_id }}">
                {% if metric_entries|length > 0 %}
                    <div class="introduction">
                        <h1>{{ metric_entries[0].company_name }}</h1>
                        <h2>Monthly Performance</h2>
                        <h3>Let's see what you have achieved with Vintly</h3>
                    </div>
                {% endif %}

                <div class="first-section">
                    {% for entry in metric_entries %}
                        {% if entry.metric_name == 'Total Count Of Transports (This Month)' %}
                            <div class="box1">
                                <h4>Total Count of Transports</h4>
                                <p><b>Last Month</b></p>
                                <p>{{ entry.metric_value }}</p>
                                {% if 'Total Count Of Transports (This Year)' in metric_entries|map(attribute='metric_name') %}
                                    <p><b>Year to Date</b></p>
                                    <p>{{ metric_entries|selectattr('metric_name', 'equalto', 'Total Count Of Transports (This Year)')|map(attribute='metric_value')|first }}</p>
                                {% endif %}
                            </div>
                        {% elif entry.metric_name == 'Total Deliveries Not On Time (This Month)' %}
                            <div class="box2">
                                <h4>Total Deliveries Not On Time</h4>
                                <p><b>Last Month</b></p>
                                <p>{{ entry.metric_value }}</p>
                                {% if 'Total Deliveries Not On Time (This Year)' in metric_entries|map(attribute='metric_name') %}
                                    <p><b>Year to Date</b></p>
                                    <p>{{ metric_entries|selectattr('metric_name', 'equalto', 'Total Deliveries Not On Time (This Year)')|map(attribute='metric_value')|first }}</p>
                                {% endif %}
                            </div>
                        {% endif %}
                    {% endfor %}
                </div>

                <!-- Add the second-section here -->
                <div class="second-section">
                    <div class="box3">
                        <h4>Top Destinations</h4>
                        {% for entry in metric_entries %}
                            {% if entry.metric_name == 'Top Destinations (This Month)' or entry.metric_name == 'Top Destinations (This Year)' %}
                                {% if entry.metric_name == 'Top Destinations (This Month)' %}
                                    <p><b>Last Month</b></p>
                                {% elif entry.metric_name == 'Top Destinations (This Year)' %}
                                    <p><b>Year to Date</b></p>
                                {% endif %}
                                <div>{{ entry.metric_value }}</div>
                            {% endif %}
                        {% endfor %}
                    </div>
                </div>
            </div>
        {% endfor %}
</div>
</body>
</html>










"""

# Create a Jinja2 template from the string
template = Template(html_body_template)

# Render the template with the merged data
rendered_template = template.render(merged_data_dict=merged_data_dict)
# print(rendered_template)  


# Group the merged DataFrame by 'metric_id'
grouped_df = merged_df.groupby("metric_id")

# Define email sender
email_sender = "robin.martin@vintly.com"
email_password = "jdvmktdtzuxmbzpe"
# email_password = os.environ.get("EMAIL_PASSWORD")

# ... (previous code)

# Group the merged DataFrame by 'metric_id'
grouped_df = merged_df.groupby("metric_id")

# ... (previous code)

# ... (previous code)

# Set to keep track of unique email addresses that have received emails
sent_emails = set()

# Loop through the grouped data and send personalized emails
for (metric_id, email_address), group_data in merged_df.groupby(['metric_id', 'email']):
    # Ensure that group_data is not empty
    if not group_data.empty:
        # Use the first row's company_name as the subject
        current_subject = f"Your Subject for {group_data.iloc[0]['company_name']} - Metric: {metric_id}"

        # Render the template with group-specific data
        html_content = template.render(
            merged_data_dict=group_data.to_dict(orient="records"),
        )

        # Create the email message
        em = MIMEMultipart()
        em["From"] = email_sender
        em["To"] = email_address
        em["Subject"] = current_subject

        # Attach HTML content
        html_part = MIMEText(html_content, "html")
        em.attach(html_part)

        # Attach the image with 'cid:' scheme
        image_attachment = MIMEImage(image_content)
        image_attachment.add_header("Content-ID", "<logo_image>")
        em.attach(image_attachment)

        # Add SSL (layer of security)
        context = ssl.create_default_context()

        # Log in and send the email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
            smtp.login(email_sender, email_password)
            smtp.sendmail(email_sender, email_address, em.as_string())

        # Add the email address to the set to mark it as sent
        sent_emails.add(email_address)

