import os.path
import json
# import threading
# import logging
import schedule
import time
# import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

class GoogleSheetsApp:
    def __init__(self):
        # self.schedule_check()
        # Schedule the check and send warning function to run on Friday nights
        schedule.every().friday.at("23:55").do(self.check_and_send_warning)
         # Schedule the submit mail button function to run on Friday nights
        schedule.every().friday.at("23:55").do(self.on_submit)


    def authenticate(self):
        creds = None
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                credentials_file = self.read_settings()['credentials_file']
                flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
                creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        return creds  

    def read_settings(self):
        with open("settings.json", "r") as f:
            settings = json.load(f)
        return settings

    def send_negative_report_email(self, recipient_email, message):
        settings = self.read_settings()
        sender_email = settings['sender_email']
        sender_password = settings['sender_password']
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Items Quantity Report"
        message_with_heading = "<html><body><h2>Negative Shortage Items List </h2>"
        message_with_heading += message 
        message_with_heading += "</table></body></html>"
        msg.attach(MIMEText(message_with_heading, 'html'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()


    def send_warning_email(self, recipient_email, message):
        settings = self.read_settings()
        sender_email = settings['sender_email']
        sender_password = settings['sender_password']
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Warning"
        message_with_heading = "<html><body><h2>Minimum Stock Qty Report</h2>"
        message_with_heading += message  
        message_with_heading += "</table></body></html>"
        msg.attach(MIMEText(message_with_heading, 'html'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()     

    def check_and_send_warning(self):
        try:
            settings = self.read_settings()
            spreadsheet_id = settings['spreadsheet_id']
            item_name_column_index = settings.get('item_name_column_index', 1)

            selected_sheet_name = "Minimum Inventory"  

            creds = self.authenticate()
            if not creds:
                return

            service = build("sheets", "v4", credentials=creds)
            sheet = service.spreadsheets()

            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f"'{selected_sheet_name}'!A:ZZ").execute()
            values = result.get("values", [])

            if not values:
                print("No data found.")
                return

            header_row = values[0]
            minimum_index = None
            required_index = None

            for index, column_name in enumerate(header_row):
                if column_name == "Minimum":
                    minimum_index = index
                elif column_name == "Required":
                    required_index = index
            
            if minimum_index is None or required_index is None:
                print("Minimum or Required column not found.")
                return

            warnings = []
            for row in values[1:]:
                try:
                    minimum_value = float(row[minimum_index])
                    required_value = float(row[required_index])
                    if minimum_value >= 0.7 * required_value:
                        warnings.append(row)
                except (ValueError, IndexError):
                    pass

            if warnings:
                recipient_email = settings['recipient_email']
                sender_email = settings['sender_email']
                message = "<html><body><table border='1'><tr><th>Item Name</th><th>Minimum Stock Qty</th><th>Required Order Qty</th></tr>"
                for warning in warnings:
                    if item_name_column_index < len(warning):
                        message += f"<tr><td>{warning[item_name_column_index]}</td><td>{warning[minimum_index]}</td><td>{warning[required_index]}</td></tr>"
                    else:
                        print("Invalid item name index.")
                message += "</table></body></html>"
                self.send_warning_email(recipient_email, message)
                print("Warning email sent successfully!")
            else:
                print("No items found with Minimum Stock Qty less than 70% of Required Order Qty.")
        except Exception as ex:
            print(f"An error occurred: {ex}")
            
    # def schedule_check(self):
    #     threading.Timer(60.0, self.check_and_send_warning).start()

    def on_submit(self):
        settings = self.read_settings()
        spreadsheet_id = settings['spreadsheet_id']
        sheet_name = "Minimum Inventory"
        column_name = "Shortage"
        creds = self.authenticate()
        if not creds:
            return
        try:
            service = build("sheets", "v4", credentials=creds)
            sheet = service.spreadsheets()

            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!1:1").execute()
            header_row = result.get("values", [])[0]
            
            column_index = None
            for index, column in enumerate(header_row):
                if column.strip().lower() == column_name.lower():
                    column_index = index
                    break
            
            if column_index is None:
                print(f"Column '{column_name}' not found.")
                return

            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A:ZZ").execute()
            values = result.get("values", [])

            if not values:
                print("No data found.")
                return
            
            negative_values = []
            for row in values:
                try:
                    value = float(row[column_index])
                    if value < 0:
                        negative_values.append(row)
                except (ValueError, IndexError):
                    pass
            
            if negative_values:
                settings = self.read_settings()
                recipient_email = settings['recipient_email']
                html_table = "<table border='1'><tr><th>Item Name</th><th>Shortage Qty</th></tr>"
                for row in negative_values:
                    html_table += f"<tr><td>{row[1]}</td><td>{row[column_index]}</td></tr>"
                html_table += "</table>"
                self.send_negative_report_email(recipient_email, html_table)
                print("Email sent successfully!")
            else:
                print(f"No negative values found in column '{column_name}'.")
        except HttpError as err:
            print(f"An HTTP error occurred: {err}")

def main():
    app = GoogleSheetsApp()
    while True:
        schedule.run_pending()
        time.sleep(60)# Wait for 60 seconds

if __name__ == "__main__":
    main()
