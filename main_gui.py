import tkinter as tk
from tkinter import filedialog
import os, pyperclip
from dotenv import load_dotenv
from toolbox_scripts.other.sort_my_data import sort_my_data
from toolbox_scripts.reports.report_groups_dispoview import GroupsDispoview
from toolbox_scripts.other.create_pos import create_csv_pos
from toolbox_scripts.read.read_supply_data import SupplyDataReader

load_dotenv()

paths = {
    "po_template": os.getenv("PATH_PO_TEMPLATE"),
    "groups": os.getenv("PATH_GROUPS"),
    "dispoview_groups": os.getenv("PATH_DISPO_GROUPS"),
    "db": os.getenv("PATH_DB"),
    "my_data": os.getenv("PATH_MY_DATA"),
    "products_files": os.getenv("PATH_PRODUCTS"),
    "components_info": os.getenv("PATH_COMPONENTS"),
}

po_data = {
    "ccn": os.getenv("PO_CCN"),
    "mas_loc": os.getenv("PO_MAS_LOC"),
    "request_div": os.getenv("PO_REQUEST_DIV"),
    "pur_loc": os.getenv("PO_PUR_LOC"),
    "delivery": os.getenv("PO_DELIVERY"),
    "inspection": os.getenv("PO_INSPECTION"),
}

supply_data = {
    "supplier_1": os.getenv("SUPPLIER_1"),
    "supplier_2": os.getenv("SUPPLIER_2"),
    "supplier_3": os.getenv("SUPPLIER_3"),
    "incoterms": os.getenv("SUPPLY_INCOTERMS"),
    "t_mode": os.getenv("SUPPLY_T_MODE"),
}

passwords = {}
for i in range(1, 6):
    passwords[f"name_{i}"] = os.getenv(f"PASSWORD_{i}").split(", ")[0]
    passwords[f"pass_{i}"] = os.getenv(f"PASSWORD_{i}").split(", ")[1]


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Work_Toolbox")
        self.configure(bg="lightgray")
        self.create_widgets()

    def create_widgets(self):
        labels = ["Select", "Files", "Reports", "Others", "Passwords"]
        for index, label in enumerate(labels):
            lbl = tk.Label(self, text=label, font=("Arial", 12, "bold"), bg="lightgray")
            lbl.grid(row=0, column=index, padx=5, pady=5)

        btn_1 = tk.Button(
            self,
            text=f"Dispoview_groups_statuses",
            command=lambda: GroupsDispoview(
                dispo_file_path=entry_dispo.get(),
                groups_file_path=paths["dispoview_groups"],
                supply_file_path=entry_supply.get(),
            )(),
        )
        btn_1.grid(row=1, column=2, padx=5, pady=5)

        btn_2 = tk.Button(
            self,
            text=f"Create_POs",
            width=20,
            command=lambda: create_csv_pos(
                path_excel_dat=paths["po_template"],
                ccn=po_data["ccn"],
                mas_loc=po_data["mas_loc"],
                request_div=po_data["request_div"],
                pur_loc=po_data["pur_loc"],
                delivery=po_data["delivery"],
                inspection=po_data["inspection"],
            ),
        )
        btn_2.grid(row=1, column=3, padx=5, pady=5)

        btn_3 = tk.Button(
            self,
            text=f"Sort_My_Data",
            width=20,
            command=lambda: sort_my_data(paths["my_data"]),
        )
        btn_3.grid(row=2, column=3, padx=5, pady=5)

        btn_4 = tk.Button(
            self,
            text=f"New_supply_info",
            width=20,
            command=lambda: SupplyDataReader(
                path_supply=entry_all_dram.get(), path_groups=paths["groups"]
            )(),
        )
        btn_4.grid(row=3, column=3, padx=5, pady=5)

        path_dispo = tk.Button(
            self,
            text=f"Dispoview_data",
            command=lambda path_name=f"dispoview_data_path": self.select_path(
                path_name
            ),
            width=20,
        )
        path_dispo.grid(row=1, column=0, padx=5, pady=5)
        entry_dispo = tk.Entry(self, name=f"dispoview_data_path", width=50)
        entry_dispo.grid(row=1, column=1, padx=5, pady=5)

        path_supply = tk.Button(
            self,
            text=f"Supply_data",
            command=lambda path_name=f"supply_path": self.select_path(path_name),
            width=20,
        )
        path_supply.grid(row=2, column=0, padx=5, pady=5)
        entry_supply = tk.Entry(self, name=f"supply_path", width=50)
        entry_supply.grid(row=2, column=1, padx=5, pady=5)

        path_all_dram = tk.Button(
            self,
            text=f"All_dram_data",
            command=lambda path_name=f"supply_all_dram": self.select_path(path_name),
            width=20,
        )
        path_all_dram.grid(row=3, column=0, padx=5, pady=5)
        entry_all_dram = tk.Entry(self, name=f"supply_all_dram", width=50)
        entry_all_dram.grid(row=3, column=1, padx=5, pady=5)
        for x in range(1, 6):
            pass_bt = tk.Button(
                self,
                text=passwords[f"name_{x}"],
                width=50,
                command=lambda pssw=passwords[f"pass_{x}"]: pyperclip.copy(pssw),
            )
            pass_bt.grid(row=x, column=4, padx=5, pady=5)

    def select_path(self, path_name):
        path = filedialog.askopenfilename()
        entry = self.nametowidget(path_name)
        entry.delete(0, tk.END)
        entry.insert(0, path)
        entry.xview_moveto(1)


if __name__ == "__main__":
    app = Application()
    app.mainloop()
