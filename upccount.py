import io
import threading
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import pyodbc
from tkinter import filedialog
from configparser import ConfigParser
from tkinter import ttk
from tkcalendar import DateEntry  # Make sure to install this library
import schedule
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from io import BytesIO
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import date, datetime, timedelta
import re
from datetime import datetime, timedelta
from queue import Queue
import time
import tkinter as tk
from tkinter import ttk, filedialog
from io import StringIO
import base64
import smtplib
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt

from tkinter import filedialog
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
import tempfile
import os

server_entry = None
database_entry = None
username_entry = None
password_entry = None
smtp_server_entry = None
smtp_username_entry = None
smtp_password_entry = None
smtp_from_entry = None
to_email_entry = None
time_entry = None

# Define config globally
config = ConfigParser()
# Function to save both SQL Server and SMTP details to config.ini file
def save_config(config_window):
    config["DATABASE"] = {
        "server": base64.b64encode(server_entry.get().encode()).decode(),
        "database": base64.b64encode(database_entry.get().encode()).decode(),
        "username": base64.b64encode(username_entry.get().encode()).decode(),
        "password": base64.b64encode(password_entry.get().encode()).decode(),
    }

    config["SMTP"] = {
        "server": base64.b64encode(smtp_server_entry.get().encode()).decode(),
        "username": base64.b64encode(smtp_username_entry.get().encode()).decode(),
        "password": base64.b64encode(smtp_password_entry.get().encode()).decode(),
        "from": base64.b64encode(smtp_from_entry.get().encode()).decode(),
        "to": base64.b64encode(to_email_entry.get().encode()).decode(),
        "time": time_entry.get(),
    }

    with open("upccount_config.ini", "w") as configfile:
        config.write(configfile)

    status_label.config(text="Configuration saved successfully!", fg="green")
    config_window.destroy()


def generate_report_3(start_date_time_str, end_date_time_str):

    print(f"Start Date-Time: {start_date_time_str}")
    print(f"End Date-Time: {end_date_time_str}")
    try:
        config = ConfigParser()
        config.read("upccount_config.ini")
        server = base64.b64decode(config.get("DATABASE", "server").encode()).decode()
        database = base64.b64decode(config.get("DATABASE", "database").encode()).decode()
        username = base64.b64decode(config.get("DATABASE", "username").encode()).decode()
        password = base64.b64decode(config.get("DATABASE", "password").encode()).decode()

        if username and password:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        else:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes"

        connection = pyodbc.connect(connection_string)

        report_query = f"""
            
                SELECT 
            TH.Branch,
            TH.Station,
            TL.UPC, 
            TH.Receipt AS TransactionID, 
            TH.TheUser AS Operator, 
            TL.SubAfterTax AS Amount,
            COALESCE(TL.TL_FieldText1, 'No Line Note') AS LineNote
        FROM TransLines TL
        INNER JOIN TransHeaders TH ON TL.TransNo = TH.TransNo
        WHERE TH.Logged >= '{start_date_time_str}' 
            AND TH.Logged <= '{end_date_time_str}' 
            AND TL.UPC = '0000';
        """

        df = pd.read_sql_query(report_query, connection)

        # Format the 'Amount' column
        df['Amount'] = df['Amount'].apply(lambda x: f"${x:.2f}")

        if df.empty:
            # Generate an empty dataframe with headers
            headers = ["Branch", "Station", "UPC", "TransactionID", "Operator", "Amount", "LineNote"]
            df = pd.DataFrame(columns=headers)
        else:
            # Convert UPC to string and add single quotes
            df["UPC"] = "'" + df["UPC"].astype(str) + "'"

            # Append total count and total amount
            total_amount = df["Amount"].replace('[\$,]', '', regex=True).astype(float).sum()
            total_row = pd.DataFrame(
                {
                    "UPC": ["'Total count'"],
                    "TransactionID": [df["UPC"].count()],
                    "Operator": [""],
                    "Amount": [f"${total_amount:.2f}"],
                }
            )
            df = pd.concat([df, total_row], ignore_index=True)
            
        # Generate a unique temporary filename to save the CSV
        temp_file_name = tempfile.mktemp(suffix=".csv")

        try:
            df.to_csv(temp_file_name, index=False)
            return temp_file_name  # Return the path to the temporary file
        except Exception as e:
            print(f"An error occurred while writing to CSV: {e}")
            return None

    except Exception as e:
        print(f"An error occurred: {e}")


def daily_report_task():
    yesterday = date.today() - timedelta(days=1)
    start_date_time_str = yesterday.strftime("%Y-%m-%d 00:00:00")
    end_date_time_str = yesterday.strftime("%Y-%m-%d 23:59:59")

    report_path = generate_report_3(start_date_time_str, end_date_time_str)
    
    if report_path:  # Only send the report if the path exists (not None)
        send_report_via_email(report_path)
        os.remove(report_path)  # Delete the temporary file after sending the email
    else:
        print("Report generation failed. No report sent.")


def schedule_reports():
    config.read("upccount_config.ini")
    time_to_send = config.get("SMTP", "time")

    schedule.every().day.at(time_to_send).do(daily_report_task)

    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute


def generate_both_reports(
    start_date_str,
    start_time_str,
    end_date_str,
    end_time_str,
    current_date_time_str,
    previous_date_time_str,
):
    start_date_time_str = f"{start_date_str} {start_time_str}"
    end_date_time_str = f"{end_date_str} {end_time_str}"

    generate_report_3(start_date_time_str, end_date_time_str)

def send_report_via_email(file_path):
    config.read("upccount_config.ini")
    
    smtp_server = base64.b64decode(config.get("SMTP", "server").encode()).decode()
    smtp_username = base64.b64decode(config.get("SMTP", "username").encode()).decode()
    smtp_password = base64.b64decode(config.get("SMTP", "password").encode()).decode()
    smtp_from = base64.b64decode(config.get("SMTP", "from").encode()).decode()
    smtp_to = base64.b64decode(config.get("SMTP", "to").encode()).decode()
    
    # Current date to be appended in the email subject
    yesterday = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")


    msg = MIMEMultipart()
    msg["From"] = smtp_from
    msg["To"] = smtp_to
    msg["Subject"] = f"UPC Count Daily Report for {yesterday}"

    # Attach the file
    with open(file_path, "rb") as file:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file.read())
        encoders.encode_base64(part)
        # Set desired name for the email attachment
        part.add_header("Content-Disposition", "attachment; filename=UPC count Sales.csv")
        msg.attach(part)

    text = MIMEText("Attached is the daily report.")
    msg.attach(text)

    try:
        with smtplib.SMTP_SSL(smtp_server, 465) as server:
            server.login(smtp_username, smtp_password)
            server.sendmail(smtp_from, [smtp_to], msg.as_string())
    except Exception as e:
        print(f"An error occurred while sending the email: {e}")

if __name__ == "__main__":
    schedule_reports()

def open_config_window():
    global server_entry, database_entry, username_entry, password_entry
    global smtp_server_entry, smtp_username_entry, smtp_password_entry
    global smtp_from_entry, to_email_entry, time_entry

    config_window = tk.Toplevel(window)
    config_window.title("Update Configuration")

    # All your Labels and Entries come here

    # Server Label and Entry
    server_label = tk.Label(config_window, text="Server:")
    server_label.pack()
    server_entry = tk.Entry(config_window)
    server_entry.pack()

    # Database Label and Entry
    database_label = tk.Label(config_window, text="Database:")
    database_label.pack()
    database_entry = tk.Entry(config_window)
    database_entry.pack()

    # Username Label and Entry
    username_label = tk.Label(
        config_window, text="Username (leave blank for Windows Authentication):"
    )
    username_label.pack()
    username_entry = tk.Entry(config_window)
    username_entry.pack()

    # Password Label and Entry
    password_label = tk.Label(
        config_window, text="Password (leave blank for Windows Authentication):"
    )
    password_label.pack()
    password_entry = tk.Entry(config_window, show="*")
    password_entry.pack()

    # SMTP Server Label and Entry
    smtp_server_label = tk.Label(config_window, text="SMTP Server:")
    smtp_server_label.pack()
    smtp_server_entry = tk.Entry(config_window)
    smtp_server_entry.pack()

    # SMTP Username Label and Entry
    smtp_username_label = tk.Label(config_window, text="SMTP Username:")
    smtp_username_label.pack()
    smtp_username_entry = tk.Entry(config_window)
    smtp_username_entry.pack()

    # SMTP Password Label and Entry
    smtp_password_label = tk.Label(config_window, text="SMTP Password:")
    smtp_password_label.pack()
    smtp_password_entry = tk.Entry(config_window, show="*")
    smtp_password_entry.pack()

    # 'From' Email Address Label and Entry
    smtp_from_label = tk.Label(config_window, text="'From' Email Address:")
    smtp_from_label.pack()
    smtp_from_entry = tk.Entry(config_window)
    smtp_from_entry.pack()

    # 'To' Email Address Label and Entry
    to_email_label = tk.Label(config_window, text="'To' Email Address:")
    to_email_label.pack()
    to_email_entry = tk.Entry(config_window)
    to_email_entry.pack()

    # Time to Send Report Label and Entry
    time_entry_label = tk.Label(config_window, text="Time to Send Report (HH:MM):")
    time_entry_label.pack()
    time_entry = tk.Entry(config_window)
    time_entry.pack()

    # Save SMTP Config Button
    # Pass the config_window to the save_config function
    save_config_button = tk.Button(
        config_window, text="Save Config", command=lambda: save_config(config_window)
    )
    save_config_button.pack()


# Create the main application window
window = tk.Tk()
window.title("Moore Wilsons Debtor Report")
window.geometry("300x300")

# # Start the status label update function
# update_status_label()


start_date_label = ttk.Label(window, text="Start date:")
start_date_label.pack()
start_date_entry = DateEntry(window)
start_date_entry.pack()

start_time_label = tk.Label(window, text="Start time (HH:MM):")
start_time_label.pack()
start_time_entry = ttk.Entry(window)
start_time_entry.pack()

end_date_label = ttk.Label(window, text="End date:")
end_date_label.pack()
end_date_entry = DateEntry(window)
end_date_entry.pack()

end_time_label = tk.Label(window, text="End time (HH:MM):")
end_time_label.pack()
end_time_entry = ttk.Entry(window)
end_time_entry.pack()

generate_report_button = tk.Button(
    window,
    text="Generate Report",
    command=lambda: generate_both_reports(
        start_date_entry.get(),
        start_time_entry.get(),
        end_date_entry.get(),
        end_time_entry.get(),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        (datetime.now() - timedelta(seconds=1)).strftime("%Y-%m-%d %H:%M:%S"),
    ),
)
generate_report_button.pack()


# 'Update Configuration' Button
update_config_button = tk.Button(
    window, text="Update Configuration", command=open_config_window
)
update_config_button.pack()


# Create status label
status_label = ttk.Label(window)
status_label.pack()

# # Schedule Email Button
# schedule_email_button = tk.Button(window, text='Schedule Email', command=schedule_and_start)
# schedule_email_button.pack()

# Status Label
status_label = tk.Label(window, text="")
status_label.pack()

# Start the GUI event loop
window.mainloop()
