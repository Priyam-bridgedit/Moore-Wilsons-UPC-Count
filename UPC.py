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


# Function to schedule the report generation and email sending
time_format = re.compile("^([0-1]?[0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$")


def schedule_report():
    print("Running schedulereport...")
    global config
    config.read("upccount_config.ini")
    time_to_send = config.get("SMTP", "time")
    
    if not time_format.match(time_to_send):
        status_queue.put(
            (
                "Invalid time format! Please enter time in HH:MM or HH:MM:SS format.",
                "red",
            )
        )
        return

    def scheduled_task():
            now = datetime.now()
            end_date_time = now - timedelta(days=97)
            start_date_time = end_date_time

            start_date_time_str = start_date_time.strftime("%Y-%m-%d %H:%M:%S")
            end_date_time_str = end_date_time.strftime("%Y-%m-%d %H:%M:%S")

            

    schedule.every().day.at(time_to_send).do(scheduled_task)

    print("Report Scheduled")




# Create a queue for status updates
status_queue = Queue()

# Define window as a global variable
window = None


def generate_report_3(start_date_time_str, end_date_time_str):

    print(f"Start Date-Time: {start_date_time_str}")
    print(f"End Date-Time: {end_date_time_str}")
    try:
        config = ConfigParser()
        config.read("upccount_config.ini")
        server = base64.b64decode(config.get("DATABASE", "server").encode()).decode()
        database = base64.b64decode(
            config.get("DATABASE", "database").encode()
        ).decode()
        username = base64.b64decode(
            config.get("DATABASE", "username").encode()
        ).decode()
        password = base64.b64decode(
            config.get("DATABASE", "password").encode()
        ).decode()

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

        # If dataframe is not empty
        if not df.empty:
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
            root = tk.Tk()
            root.withdraw()

            file_path = filedialog.asksaveasfilename(defaultextension=".csv")
            root.destroy()  # Close the tkinter window

            if file_path:
                # Save the DataFrame to CSV
                try:
                    df.to_csv(file_path, index=False)
                except Exception as e:
                    print(f"An error occurred while writing to CSV: {e}")

        connection.close()
    except Exception as e:
        print(f"An error occurred: {e}")


def generate_report_3_auto(start_date_time_str, end_date_time_str):
    # # Calculate the date range for the last week
    # end_date_time = datetime.now()
    # start_date_time = end_date_time - timedelta(weeks=1)

    # # Convert datetime objects to strings
    # start_date_time_str = start_date_time.strftime("%Y-%m-%d %H:%M:%S")
    # end_date_time_str = end_date_time.strftime("%Y-%m-%d %H:%M:%S")

    # Use the rest of your generate_report_3 function but replace the input date strings with these generated ones
    # Rest of the code...
    try:
        config = ConfigParser()
        config.read("upccount_config.ini")
        server = base64.b64decode(config.get("DATABASE", "server").encode()).decode()
        database = base64.b64decode(
            config.get("DATABASE", "database").encode()
        ).decode()
        username = base64.b64decode(
            config.get("DATABASE", "username").encode()
        ).decode()
        password = base64.b64decode(
            config.get("DATABASE", "password").encode()
        ).decode()

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
        FROM AKPOS.dbo.TransLines TL
        INNER JOIN AKPOS.dbo.TransHeaders TH ON TL.TransNo = TH.TransNo
        WHERE TH.Logged >= '{start_date_time_str}' 
            AND TH.Logged <= '{end_date_time_str}' 
            AND TL.UPC = '0000';
        """

        df = pd.read_sql_query(report_query, connection)

        # Format the 'Amount' column
        df['Amount'] = df['Amount'].apply(lambda x: f"${x:.2f}")

        # If dataframe is not empty
        if not df.empty:
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

        connection.close()

        return df  # Add this line to return df
    except Exception as e:
        print(f"An error occurred: {e}")


from pandas import ExcelWriter


def send_report(df3, start_date_time_str, end_date_time_str):
    print("Running scheduled_email...")
    try:
        # Check if df1, df3, and df4 are not None
        # if df1 is None:
        #     raise ValueError("df1 is None")
        if df3 is None:
            raise ValueError("df3 is None")
        # if df4 is None:
        #     raise ValueError("df4 is None")

        # Create StringIO objects and save the DataFrames to them
        csv_buffer3 = StringIO()
        df3.to_csv(csv_buffer3, index=False)
        print(csv_buffer3.getvalue())
        #csv_buffer3.seek(0)  # Reset buffer position

        
        # Read SMTP details from config.ini file
        config = ConfigParser()
        config.read("upccount_config.ini")
        smtp_server = base64.b64decode(config.get("SMTP", "server").encode()).decode()
        smtp_username = base64.b64decode(
            config.get("SMTP", "username").encode()
        ).decode()
        smtp_password = base64.b64decode(
            config.get("SMTP", "password").encode()
        ).decode()
        smtp_from = base64.b64decode(config.get("SMTP", "from").encode()).decode()
        to_email = (
            base64.b64decode(config.get("SMTP", "to").encode()).decode().split(",")
        )

        with smtplib.SMTP_SSL(smtp_server, 465) as server:
            server.login(smtp_username, smtp_password)
            for to_address in to_email:
                to_address = to_address.strip()

                # Initialize the MIMEMultipart instance inside the loop
                msg = MIMEMultipart()
                msg["From"] = smtp_from
                msg["To"] = to_address
                msg[
                    "Subject"
                ] = f"UPC Sales Daily Reports from {start_date_time_str} and {end_date_time_str}"

                body = "Please find the daily report attached."
                msg.attach(MIMEText(body, "plain"))

                
                part3 = MIMEBase("application", "octet-stream")
                part3.set_payload(csv_buffer3.getvalue())
                encoders.encode_base64(part3)
                part3.add_header(
                    "Content-Disposition",
                    f"attachment; filename= UPC Count Report.csv",
                )

                
                #msg.attach(part1)
                msg.attach(part3)
                #msg.attach(part4)

                # if df2 is not None:  # Attach the PDF report only if df2 is not None
                #     part2 = MIMEBase("application", "octet-stream")
                #     part2.set_payload(pdf_buffer2)
                #     encoders.encode_base64(part2)
                #     part2.add_header(
                #         "Content-Disposition",
                #         f"attachment; filename= Weekly_Report2_{start_date_time_str}_to_{end_date_time_str}.pdf",
                #     )
                #     msg.attach(part2)

                server.send_message(msg)
                print(f"Reports sent to {to_address}")

    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback

        print(traceback.format_exc())


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
    


def start_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)


# Function to schedule the report and start the scheduler
def schedule_and_start():
    schedule_report()
    scheduler_thread = threading.Thread(target=start_scheduler)
    scheduler_thread.start()


if __name__ == "__main__":
    schedule_and_start()


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
