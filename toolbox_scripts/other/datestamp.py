import pyperclip
import datetime


def datestamp():
    current_date = datetime.datetime.now().strftime("%d%m%Y")
    filename = f"sap_log_data_{current_date}.XLSX"
    pyperclip.copy(filename)
