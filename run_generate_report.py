import pandas as pd
import matplotlib.pyplot as plt
from jinja2 import Template
from datetime import date
import smtplib
import sys
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import datetime
import time
from credentials.config import url, sender_email, sender_password, receiver_email, cc_email, subject, body
import base64
import subprocess
import tkinter as tk
from tkinter import messagebox

def check_internet():
    print("\n (!) ...checking internet \n ")
    url='http://www.google.com'
    timeout=5
    try:
        requests.get(url, timeout=timeout)
        print("You're connected to the internet")
    except (requests.ConnectionError, requests.Timeout):
        print("No internet connection available")
        sys.exit(1)

def update_logs():
    print("\n (!) ...update logs \n ")
    # Get today's date
    today = datetime.date.today().strftime("%Y-%m-%d")

    # Append today's date to the beginning of the file
    with open("runLogs.txt", "r+") as f:
        content = f.read()
        f.seek(0, 0)
        f.write(today + "\n" + content)
        print("Today's date added to runLogs.txt")

    time.sleep(2)

def send_email(sender_email, sender_password, receiver_email, cc_email, subject, body, attachment_file, image_files = ["last_update_chart.png", "weapon_category_chart.png"]):
    print("\n (!) ...send email \n ")
    # Create a multipart message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Cc"] = cc_email
    message["Subject"] = subject

    # Read the content of the HTML file
    with open(attachment_file, "r") as html_file:
        html_content = html_file.read()

    # Create the body of the email
    email_body = f"{body}\n\n{html_content}"

    # Attach the body as HTML to the email
    message.attach(MIMEText(email_body, "html"))

    # Attach the image files
    for image_file in image_files:
        # Open the image file in binary mode
        with open(image_file, "rb") as img:
            # Create a MIME image part
            image_part = MIMEBase("image", "png")  # Modify the MIME type if necessary
            image_part.set_payload(img.read())
            # Encode the image in base64
            encoders.encode_base64(image_part)
            # Add the necessary header information
            image_part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(image_file)}")
            image_part.add_header("Content-ID", f"<{os.path.basename(image_file)}>")

        # Attach the image part to the message
        message.attach(image_part)

    # Convert message to string
    text = message.as_string()

    # Connect to the email server and send the email
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [receiver_email] + [cc_email], text)

def generate_report_html(df):
    print("\n (!) ...generate html report \n ")
    # Create the HTML template
    template = Template('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Weapon Data Report - {{ report_date }}</title>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
    </head>
    <body>
        <h1>Weapon Data Report - {{ report_date }}</h1>

        <table class="table table-bordered table-hover">
              <thead>
                  <tr>
                      <th scope="col">Title</th>
                      <th scope="col">Link URL</th>
                  </tr>
              </thead>
              <tbody>
                  <tr>
                      <td>OneDrive Word File</td>
                      <td><a href="https://www.example1.com">OneDrive</a></td>
                  </tr>
                  <tr>
                      <td>Excel File</td>
                      <td><a href="https://www.example2.com">Excel File</a></td>
                  </tr>
              </tbody>
          </table>
        
        <h2>Count of Weapons by Category:</h2>
        {{ category_count }}

        <h2>Status Summary:</h2>
        {{ status_summary }}

        <h2>Stage Summary:</h2>
        {{ stage_summary }}

        <h2>Recent Update:</h2>
        <p>{{ recent_update }}</p>

        <h2>Most Viewed Weapons:</h2>
        {{ most_viewed_weapons }}

        <h2>Blocked Weapons:</h2>
        {{ blocked_weapons }}

        <h2>Blocked By Count:</h2>
        {{ blocked_reasons }}

        <h2>Weapon URLs:</h2>
        {{ weapon_urls }}

        <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
        <script src="https://cdn.jsdelivr.net/npm/popper.js@1.12.9/dist/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
    </body>
    </html>
    ''')

    # Count of Weapons by Category
    category_count = df['WEAPON CATEGORY'].value_counts().to_frame().to_html()

    # Status Summary
    status_summary = df['STATUS'].value_counts().to_frame().to_html()

    # Stage Summary
    stage_summary = df['STAGE'].value_counts().to_frame().to_html()

    # Recent Updates
    recent_update = df['LAST UPDATE'].max()

    # Most Viewed Weapons
    most_viewed_weapons = df.sort_values('VIZIONARI', ascending=False).head(5)[['WEAPON NAME', 'VIZIONARI']].to_html(index=False)

    # Blocked Weapons
    blocked_weapons = df[df['STATUS'] == 'Weapon Locked'][['WEAPON NAME', 'BLOCKED BY']].to_html(index=False)

    # Blocked Reasons
    blocked_reasons = df[df['STATUS'] == 'Weapon Locked']['BLOCKED BY'].value_counts().to_frame().to_html()

    # Weapon URLs
    weapon_urls = df[['WEAPON NAME', 'URL']].to_html(index=False)

    # Get the current date for the report title
    report_date = date.today().strftime("%Y-%m-%d")

    # Render the HTML template with the report data
    html_content = template.render(
        report_date=report_date,
        category_count=category_count,
        status_summary=status_summary,
        stage_summary=stage_summary,
        recent_update=recent_update,
        most_viewed_weapons=most_viewed_weapons,
        blocked_weapons=blocked_weapons,
        blocked_reasons=blocked_reasons,
        weapon_urls=weapon_urls
    )

    return html_content

def is_valid_onedrive_api_url(url):
    return url.startswith("https://api.onedrive.com/v1.0/shares/u!")

def create_onedrive_directdownload(onedrive_link):
    data_bytes64 = base64.b64encode(bytes(onedrive_link, 'utf-8'))
    data_bytes64_String = data_bytes64.decode('utf-8').replace('/','_').replace('+','-').rstrip("=")
    resultUrl = f"https://api.onedrive.com/v1.0/shares/u!{data_bytes64_String}/root/content"
    return resultUrl

def create_onedrive_directdownload_no_replace(onedrive_link):
    return onedrive_link

def check_config():
    CONFIG_FILE = "credentials/config.py"
    print("\n (!) ...checking config file \n ")
    
    if not os.path.exists(CONFIG_FILE):
        print(f"{CONFIG_FILE} not found. Running create_cfg.py to create the configuration.")
        try:
            subprocess.run(["python", "create_cfg.py"])
        except FileNotFoundError:
            print("create_cfg.py script not found. Make sure it is located in the same directory.")
            exit(1)
        return
    
    # Check if the config.py file is not empty
    if os.path.getsize(CONFIG_FILE) == 0:
        print(f"{CONFIG_FILE} is empty. Running create_cfg.py to create the configuration.")
        try:
            subprocess.run(["python", "create_cfg.py"])
            exit(1)
        except FileNotFoundError:
            print("create_cfg.py script not found. Make sure it is located in the same directory.")
            exit(1)
        return
    
    # Load the config.py file as a Python module
    try:
        config_module = {}
        with open(CONFIG_FILE, "r") as f:
            exec(f.read(), config_module)
    except Exception as e:
        print(f"Error loading {CONFIG_FILE}:", str(e))
        exit(1)
    
    # Check if all the required variables are defined and not empty
    required_variables = ["url", "sender_email", "sender_password", "receiver_email", "cc_email", "subject", "body"]
    for var in required_variables:
        if var not in config_module:
            print(f"{CONFIG_FILE} is missing the variable '{var}'. Running create_cfg.py to create the configuration.")
            try:
                subprocess.run(["python", "create_cfg.py"])
            except FileNotFoundError:
                print("create_cfg.py script not found. Make sure it is located in the same directory.")
                exit(1)
            return
        if not config_module[var]:
            print(f"Value for '{var}' in {CONFIG_FILE} is empty. Running create_cfg.py to create the configuration.")
            try:
                subprocess.run(["python", "create_cfg.py"])
                exit(1)
            except FileNotFoundError:
                print("create_cfg.py script not found. Make sure it is located in the same directory.")
                exit(1)
            return
    
    print(f"{CONFIG_FILE} exists and is valid. Continuing with the main script.")

def generate_weapon_category_chart(df):
    weapon_counts = df["WEAPON CATEGORY"].value_counts()
    plt.figure(figsize=(8, 6))
    weapon_counts.plot(kind="bar")
    plt.xlabel("Weapon Category")
    plt.ylabel("Count")
    plt.title("Weapon Category Counts")
    plt.tight_layout()
    chart_file = "weapon_category_chart.png"
    plt.savefig(chart_file)
    plt.close()
    return chart_file


def generate_last_update_chart(df):
    df["LAST UPDATE"] = pd.to_datetime(df["LAST UPDATE"])
    df["DATE"] = df["LAST UPDATE"].dt.date
    update_counts = df["DATE"].value_counts().sort_index()
    plt.figure(figsize=(10, 6))
    update_counts.plot(kind="line", marker="o")
    plt.xlabel("Date")
    plt.ylabel("Count")
    plt.title("Weapon Updates Over Time")
    plt.tight_layout()
    chart_file = "last_update_chart.png"
    plt.savefig(chart_file)
    plt.close()
    return chart_file


def main():
    
    if is_valid_onedrive_api_url(url):
        url2 = create_onedrive_directdownload_no_replace(url)
    else:
        url2 = create_onedrive_directdownload(url)
        
    print(url2)

    # Read the Excel file from the URL
    df = pd.read_excel(url2)

    # Count of Weapons by Category
    category_count = df['WEAPON CATEGORY'].value_counts()

    # Status Summary
    status_summary = df['STATUS'].value_counts()

    # Stage Summary
    stage_summary = df['STAGE'].value_counts()

    # Recent Updates
    recent_update = df['LAST UPDATE'].max()

    # Most Viewed Weapons
    df['VIZIONARI'] = pd.to_numeric(df['VIZIONARI'], errors='coerce').astype('Int64')
    most_viewed_weapons = df.sort_values('VIZIONARI', ascending=False).head(5)

    # Blocked Weapons
    blocked_weapons = df[df['STATUS'] == 'Weapon Locked']
    blocked_reasons = blocked_weapons['BLOCKED BY'].value_counts()

    # Weapon URL
    weapon_urls = df[['WEAPON NAME', 'URL']]

    # Print the statistics
    print("Count of Weapons by Category:")
    print(category_count)
    print("\nStatus Summary:")
    print(status_summary)
    print("\nStage Summary:")
    print(stage_summary)
    print("\nRecent Update:", recent_update)
    print("\nMost Viewed Weapons:")
    print(most_viewed_weapons[['WEAPON NAME', 'VIZIONARI']])
    print("\nBlocked Weapons:")
    print(blocked_weapons[['WEAPON NAME', 'BLOCKED BY']])
    print("\nBlocked Reasons:")
    print(blocked_reasons)
    print("\nWeapon URLs:")
    print(weapon_urls)

    
    # Generate the HTML report
    generate_weapon_category_chart(df)
    generate_last_update_chart(df)
    html_content = generate_report_html(df)

    # Save the HTML to a file
    html_file = "weapon_data_report.html"
    with open(html_file, "w") as f:
        f.write(html_content)

    print("HTML report created successfully.")

    

    # List of image files to attach
    # image_files = ["last_update_chart.png", "weapon_category_chart.png"] 

    # Send the email with the report as an attachment
    # send_email(sender_email, sender_password, receiver_email, cc_email, subject, body, "weapon_data_report.html", image_files)
    send_email(sender_email, sender_password, receiver_email, cc_email, subject, body, "weapon_data_report.html")


    print("Email sent successfully.")


if __name__ == "__main__":
    today = datetime.date.today().strftime("%Y-%m-%d")

    with open("runLogs.txt", "r") as f:
        first_line = f.readline().strip()
        if first_line == today:
            print("Today's date already exists in runLogs.txt")
            exit()
    
    check_internet()

    check_config()
    main()

    update_logs()