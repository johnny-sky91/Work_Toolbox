import pandas as pd
import datetime, os


class DispoviewDataReader:
    def __init__(
        self,
        dispo_file_path: str,
    ):
        self.dispo_file_path = dispo_file_path
        self.raw_dispoview = None
        self.ready_dispoview = None

    def _read_dispoview(self):
        self.raw_dispoview = pd.read_excel(self.dispo_file_path)

    def _select_dispoview(self):
        columns_to_drop = [
            "SAP - Material Description",
            "SAP - Material Number (106xxx)",
            "SAP - Purchasing Group",
            "StratBuyer",
            "MRPtype",
            "SAP - MRP Controller",
            "FrameContract",
            "SAP - Planned Deliv. Time",
            "Trading Partner Code",
            "SAP - Vendor Name",
        ]
        self.ready_dispoview = self.raw_dispoview.drop(columns=columns_to_drop)

        self.ready_dispoview.rename(
            columns={"SAP - Product Number Print": "CODENUMBER", "Figure": "DATA"},
            inplace=True,
        )
        self.ready_dispoview.dropna(how="all", inplace=True)
        self.ready_dispoview.fillna(0, inplace=True)
        self.ready_dispoview = self.ready_dispoview.replace("'", "", regex=True)
        self.ready_dispoview.iloc[:, 2:] = self.ready_dispoview.iloc[:, 2:].astype(int)

    def _save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"Data_dispoview_{now.strftime('%d%m%Y_%H%M')}.xlsx"
        directory_path = os.path.dirname(self.dispo_file_path)
        report_file_path = os.path.join(directory_path, filename)

        writer = pd.ExcelWriter(report_file_path)
        self.ready_dispoview.to_excel(writer, sheet_name=f"Dispoview_data", index=False)
        writer._save()

    def __call__(self):
        self._read_dispoview()
        self._select_dispoview()
        # self._save_to_excel()
