import tkinter as tk
from tkinter import messagebox
import os
import smtplib
from email.message import EmailMessage

CONFIG_FILE = 'credentials/config.py'

# Function to save the config to the config.py file
def save_config():
    content = f"""url = "{url_entry.get()}"
sender_email = "{sender_email_entry.get()}"
sender_password = "{sender_password_entry.get()}"
receiver_email = "{receiver_email_entry.get()}"
cc_email = "{cc_email_entry.get()}"
subject = "{subject_entry.get()}"
body = "{body_entry.get()}"
"""

    with open(CONFIG_FILE, 'w') as configfile:
        configfile.write(content)

    messagebox.showinfo("Config Saved", "Config successfully saved.")

# Function to load the config from the config.py file
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as configfile:
            content = configfile.read()
            exec(content, globals())
            url_entry.insert(0, url)
            sender_email_entry.insert(0, sender_email)
            sender_password_entry.insert(0, sender_password)
            receiver_email_entry.insert(0, receiver_email)
            cc_email_entry.insert(0, cc_email)
            subject_entry.insert(0, subject)
            body_entry.insert(0, body)
    else:
        # Set default values here
        url_entry.insert(0, "your_onedrive_url")
        sender_email_entry.insert(0, "your_sender_email@example.com")
        sender_password_entry.insert(0, "")
        receiver_email_entry.insert(0, "your_receiver_email@example.com")
        cc_email_entry.insert(0, "")
        subject_entry.insert(0, "Weapon Data Report")
        body_entry.insert(0, "Please find attached the latest weapon data report.")

# Function to check if the sender password and confirm sender password match
def check_password():
    password = sender_password_entry.get()
    confirm_password = confirm_sender_password_entry.get()

    if password == confirm_password:
        messagebox.showinfo("Password Match", "Passwords match!")
    else:
        messagebox.showerror("Password Mismatch", "Passwords do not match. Please try again.")

# Function to send a test email
def send_test_email():
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()
    receiver_email = receiver_email_entry.get()
    cc_email = cc_email_entry.get()

    try:
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        msg = EmailMessage()
        msg.set_content("This is a test email.")

        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Cc'] = cc_email
        msg['Subject'] = "Test Email"

        server.sendmail(sender_email, [receiver_email] + [cc_email], msg.as_string())
        server.quit()

        messagebox.showinfo("Email Sent", "Test email sent successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while sending the test email: {e}")


# GUI
root = tk.Tk()
root.title("Config Editor")

# Labels
tk.Label(root, text="URL:").grid(row=0, column=0)
tk.Label(root, text="Sender Email:").grid(row=1, column=0)
tk.Label(root, text="Sender Password:").grid(row=2, column=0)
tk.Label(root, text="Confirm Sender Password:").grid(row=3, column=0)
tk.Label(root, text="Receiver Email:").grid(row=4, column=0)
tk.Label(root, text="CC Email:").grid(row=5, column=0)
tk.Label(root, text="Subject:").grid(row=6, column=0)
tk.Label(root, text="Body:").grid(row=7, column=0)

# Entries
url_entry = tk.Entry(root, width=40)
url_entry.grid(row=0, column=1)

sender_email_entry = tk.Entry(root, width=40)
sender_email_entry.grid(row=1, column=1)

sender_password_entry = tk.Entry(root, width=40, show="*")
sender_password_entry.grid(row=2, column=1)

confirm_sender_password_entry = tk.Entry(root, width=40, show="*")
confirm_sender_password_entry.grid(row=3, column=1)

receiver_email_entry = tk.Entry(root, width=40)
receiver_email_entry.grid(row=4, column=1)

cc_email_entry = tk.Entry(root, width=40)
cc_email_entry.grid(row=5, column=1)

subject_entry = tk.Entry(root, width=40)
subject_entry.grid(row=6, column=1)

body_entry = tk.Entry(root, width=40)
body_entry.grid(row=7, column=1)

# Buttons
save_button = tk.Button(root, text="Save Config", command=save_config)
save_button.grid(row=9, column=0, columnspan=2, pady=10)

load_button = tk.Button(root, text="Load Config", command=load_config)
load_button.grid(row=10, column=0, columnspan=2, pady=10)

check_password_button = tk.Button(root, text="Check Password", command=check_password)
check_password_button.grid(row=11, column=0, columnspan=2, pady=10)

# New "TEST EMAIL" button
test_email_button = tk.Button(root, text="TEST EMAIL", command=send_test_email)
test_email_button.grid(row=12, column=0, columnspan=2, pady=10)

# Load the config on startup if it exists
load_config()

root.mainloop()