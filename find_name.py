import os.path
import json
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox, QComboBox
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

class GoogleSheetsApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.sheet_name_label = QLabel("Sheet Name:")
        self.sheet_name_dropdown = QComboBox()
        self.column_name_label = QLabel("Column Name:")
        self.column_name_dropdown = QComboBox()
        self.submit_button = QPushButton("Submit")
        self.send_mail_button = QPushButton("Send Mail")

        layout = QVBoxLayout()
        layout.addWidget(self.sheet_name_label)
        layout.addWidget(self.sheet_name_dropdown)
        layout.addWidget(self.column_name_label)
        layout.addWidget(self.column_name_dropdown)
        layout.addWidget(self.submit_button)
        layout.addWidget(self.send_mail_button)
        self.setLayout(layout)
        self.submit_button.clicked.connect(self.on_submit)
        self.send_mail_button.clicked.connect(self.send_mail)
        settings = self.read_settings()
        spreadsheet_id = settings['spreadsheet_id']
        creds = self.authenticate()

        if creds:
            service = build("sheets", "v4", credentials=creds)
            spreadsheets = service.spreadsheets()
            result = spreadsheets.get(spreadsheetId=spreadsheet_id).execute()
            sheet_names = [sheet['properties']['title'] for sheet in result['sheets']]
            self.sheet_name_dropdown.addItems(sheet_names)
            self.sheet_name_dropdown.currentIndexChanged.connect(self.load_column_names)

    def read_settings(self):
        with open("settings.json", "r") as f:
            settings = json.load(f)
        return settings
    
    def get_column_index(self, column_name, header_row):
        try:
            return header_row.index(column_name)
        except ValueError:
            QMessageBox.warning(self, "Warning", f"Column '{column_name}' not found.")
            return None

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

    def send_email(self, receiver_email, message):
        settings = self.read_settings()
        sender_email = settings['sender_email']
        sender_password = settings['sender_password']
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Negative Values Report"
        msg.attach(MIMEText(message, 'html'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()

    def check_minimum_stock(self, values):
        settings = self.read_settings()
        minimum_stock_qty = None
        required_order_qty = None
        for row in values:
            if len(row) >= 3:
                if row[0] == "Minimum Stock Qty":
                    minimum_stock_qty = float(row[1])
                elif row[0] == "Required Order Qty":
                    required_order_qty = float(row[2])

        if minimum_stock_qty is None or required_order_qty is None:
            QMessageBox.information(self, "Information", "Minimum Stock Qty or Required Order Qty not found.")
            return

        if minimum_stock_qty < 0.7 * required_order_qty:
            recipient_email = settings['recipient_email']
            sender_email = settings['sender_email']
            sender_password = settings['sender_password']
            message = f"Warning: Minimum Stock Qty ({minimum_stock_qty}) is less than 70% of Required Order Qty ({required_order_qty})."
            self.send_email(sender_email, sender_password, recipient_email, "Warning: Minimum Stock Qty Alert", message)
            QMessageBox.information(self, "Information", "Warning email sent successfully!")
        else:
            QMessageBox.information(self, "Information", "Minimum Stock Qty is above 70% of Required Order Qty.")    

    def on_submit(self):
        settings = self.read_settings()
        spreadsheet_id = settings['spreadsheet_id']
        sheet_name = self.sheet_name_dropdown.currentText()
        column_name = self.column_name_dropdown.currentText().strip()
        creds = self.authenticate()
        if not creds:
            return
        try:
            service = build("sheets", "v4", credentials=creds)
            sheet = service.spreadsheets()

            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!1:1").execute()
            header_row = result.get("values", [])[0]
            
            column_index = self.get_column_index(column_name, header_row)
            if column_index is None:
                return

            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A:ZZ").execute()
            values = result.get("values", [])

            if not values:
                QMessageBox.information(self, "Information", "No data found.")
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
                html_table = "<table border='1'><tr><th>Item Name</th><th>Shortage Qty</th></tr>"
                item_name_index = settings.get('item_name_column_index', 1)  
                for row in negative_values:
                    html_table += f"<tr><td>{row[item_name_index]}</td><td>{row[column_index]}</td></tr>"
                html_table += "</table>"
                recipient_email = settings['recipient_email']
                self.send_email(recipient_email, html_table)

                QMessageBox.information(self, "Information", "Email sent successfully!")
            else:
                QMessageBox.information(self, "Information", f"No negative values found in column '{column_name}'.")
        except HttpError as err:
            QMessageBox.critical(self, "Error", str(err))

    def load_column_names(self):
        selected_sheet = self.sheet_name_dropdown.currentText()
        settings = self.read_settings()
        spreadsheet_id = settings['spreadsheet_id']
        creds = self.authenticate()
        if creds:
            service = build("sheets", "v4", credentials=creds)
            spreadsheets = service.spreadsheets()
            result = spreadsheets.values().get(spreadsheetId=spreadsheet_id, range=f"'{selected_sheet}'!1:1").execute()
            header_row = result.get("values", [])[0]
            self.column_name_dropdown.clear()
            self.column_name_dropdown.addItems(header_row)

    def send_mail(self):
        self.on_submit()

def main():
    app = QApplication(sys.argv)
    window = GoogleSheetsApp()
    window.setWindowTitle("Google Sheets App")
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
