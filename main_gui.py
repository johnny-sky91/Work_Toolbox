import os, pyperclip, traceback
import tkinter as tk

from tkinter import filedialog, messagebox
from dotenv import load_dotenv

from toolbox_scripts.read.read_supply_data import SupplyDataReader

from toolbox_scripts.reports.report_groups_dispoview import GroupsDispoview
from toolbox_scripts.reports.report_alloaction_data import AllocationData
from toolbox_scripts.reports.report_forecast_check import ForecastCheck

from toolbox_scripts.other.sort_my_data import sort_my_data
from toolbox_scripts.other.create_pos import create_csv_pos
from toolbox_scripts.other.datestamp import datestamp

load_dotenv()

paths, po_data, passwords = {}, {}, {}
for key, value in os.environ.items():
    if key.startswith("PATH_"):
        paths[key] = value
    elif key.startswith("PO_"):
        po_data[key] = value
    elif key.startswith("PASSWORD_"):
        passwords[value.split(", ")[0]] = value.split(", ")[1]


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Work_Toolbox")
        self.call("source", "used_theme/azure.tcl")
        self.call("set_theme", "dark")

        self.entries = {}
        self.check_vars = {}

        self.labels_widgets()
        self.select_widgets()
        self.entry_widgets()
        self.reports_widgets()
        self.others_widgets()
        self.passwords_widgets()

    def labels_widgets(self):
        labels_list = ["Select", "Files", "Reports", "Others", "Passwords"]
        for index, label in enumerate(labels_list):
            label_widget = tk.Label(self, text=label, font=("bold", 14))
            label_widget.grid(row=0, column=index, padx=5, pady=5)

    def select_widgets(self):
        select_data = {
            "Dispoview_data": lambda path_name="dispoview_data": self.select_path(
                path_name
            ),
            "Supply_data": lambda path_name="supply_data": self.select_path(path_name),
            "All_dram_data": lambda path_name="all_dram_data": self.select_path(
                path_name
            ),
            "AR_data": lambda path_name="ar_data": self.select_path(path_name),
        }
        for index, (key, value) in enumerate(select_data.items()):
            bt = tk.Button(self, text=key, command=value, width=20)
            bt.grid(row=index + 1, column=0, padx=5, pady=5)

    def entry_widgets(self):
        entries_list = ["dispoview_data", "supply_data", "all_dram_data", "ar_data"]
        for index, entry_name in enumerate(entries_list):
            entry = tk.Entry(self, name=entry_name, width=40)
            entry.grid(row=index + 1, column=1, padx=5, pady=5)
            self.entries[entry_name] = entry

    def reports_widgets(self):
        reports_data = {
            "Dispoview_groups": self.run_with_error_handling(
                lambda: GroupsDispoview(
                    dispo_file_path=self.entries["dispoview_data"].get(),
                    groups_file_path=paths["PATH_DISPO_GROUPS"],
                    supply_file_path=self.entries["supply_data"].get(),
                )()
            ),
            "Allocation_data": self.run_with_error_handling(
                lambda: AllocationData(
                    dispo_file_path=self.entries["dispoview_data"].get(),
                    groups_file_path=paths["PATH_DISPO_GROUPS"],
                )()
            ),
            "Forecast_check": self.run_with_error_handling(
                lambda: ForecastCheck(
                    forecast_file_path=self.entries["ar_data"].get()
                )(),
            ),
        }
        for index, (key, value) in enumerate(reports_data.items()):
            bt = tk.Button(self, text=key, command=value, width=20)
            bt.grid(row=index + 1, column=2, padx=5, pady=5)

    def others_widgets(self):
        others_data = {
            "Create_POs": self.run_with_error_handling(
                lambda: create_csv_pos(
                    path_excel_dat=paths["PATH_PO_TEMPLATE"],
                    ccn=po_data["PO_CCN"],
                    mas_loc=po_data["PO_MAS_LOC"],
                    request_div=po_data["PO_REQUEST_DIV"],
                    pur_loc=po_data["PO_PUR_LOC"],
                    delivery=po_data["PO_DELIVERY"],
                    inspection=po_data["PO_INSPECTION"],
                )
            ),
            "New_supply_info": self.run_with_error_handling(
                lambda: SupplyDataReader(
                    path_supply=self.entries["all_dram_data"].get(),
                    path_groups=paths["PATH_GROUPS"],
                )()
            ),
            "Sort_My_Data": self.run_with_error_handling(
                lambda: sort_my_data(paths["PATH_MY_DATA"])
            ),
            "SAP_datestamp": self.run_with_error_handling(lambda: datestamp()),
        }
        for index, (key, value) in enumerate(others_data.items()):
            bt = tk.Button(self, text=key, command=value, width=20)
            bt.grid(row=index + 1, column=3, padx=5, pady=5)

    def passwords_widgets(self):
        for index, (key, value) in enumerate(passwords.items()):
            pass_bt = tk.Button(
                self,
                text=key,
                width=20,
                command=lambda pssw=value: pyperclip.copy(pssw),
            )
            pass_bt.grid(row=index + 1, column=4, padx=5, pady=5)

    def select_path(self, path_name):
        path = filedialog.askopenfilename()
        entry = self.nametowidget(path_name)
        entry.delete(0, tk.END)
        entry.insert(0, path)
        entry.xview_moveto(1)

    def run_with_error_handling(self, func):
        def wrapper():
            try:
                func()
            except Exception as e:
                error_trace = traceback.format_exc().strip().split("\n")[-1]
                messagebox.showerror("Error", error_trace)

        return wrapper


if __name__ == "__main__":
    app = Application()
    app.mainloop()
