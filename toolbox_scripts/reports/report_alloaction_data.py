import pandas as pd
import os, datetime

from toolbox_scripts.read.read_dispoview_data import DispoviewDataReader


class AllocationData:
    def __init__(
        self,
        dispo_file_path: str,
        groups_file_path: str,
    ):
        self.dispo_file_path = dispo_file_path
        self.groups_file_path = groups_file_path
        self.raw_dispoview = None
        self.raw_groups = None
        self.components_groups = None
        self.selected_data = None
        self.gross_forecast = None
        self.stock = None

    def _read_dispoview(self):
        dispoview = DispoviewDataReader(dispo_file_path=self.dispo_file_path)
        dispoview()
        self.raw_dispoview = dispoview.ready_dispoview

    def _read_groups(self):
        self.raw_groups = pd.read_excel(self.groups_file_path, sheet_name="groups")
        self.components_groups = self.raw_groups[["COMPONENT", "CODENUMBER", "GROUP"]]
        self.components_groups = self.components_groups.drop_duplicates()

    def _merge_groups_dispoview(self):
        self.raw_dispoview = pd.merge(
            self.raw_dispoview, self.components_groups, on="CODENUMBER", how="left"
        )
        self.raw_dispoview.insert(1, "COMPONENT", self.raw_dispoview.pop("COMPONENT"))
        self.raw_dispoview.insert(2, "GROUP", self.raw_dispoview.pop("GROUP"))

    def _get_stock(self):
        self.stock = self.raw_dispoview.loc[self.raw_dispoview["DATA"].isin(["Stock"])]
        self.stock = self.stock.iloc[:, [1, 4]]
        self.stock.columns = ["COMPONENT", "STOCK"]

    def _create_forecast(self):
        self.selected_data = self.raw_dispoview.loc[
            self.raw_dispoview["DATA"].isin(["NetForecast", "CustOrders RDD"])
        ]
        self.selected_data = self.selected_data.drop(
            columns=["DATA", "CODENUMBER", "GROUP"]
        )
        self.gross_forecast = (
            self.selected_data.groupby("COMPONENT").sum().reset_index()
        )
        self.gross_forecast = pd.merge(
            self.gross_forecast, self.stock, on="COMPONENT", how="left"
        )
        self.gross_forecast.insert(1, "STOCK", self.gross_forecast.pop("STOCK"))

    def _save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"EMS_forecast_{now.strftime('%d%m%Y')}.xlsx"
        directory_path = os.path.dirname(self.dispo_file_path)
        report_file_path = os.path.join(directory_path, filename)

        writer = pd.ExcelWriter(report_file_path)
        self.gross_forecast.to_excel(writer, sheet_name=f"Sheet1", index=False)
        self.raw_dispoview.to_excel(writer, sheet_name=f"Raw_dispoview", index=False)
        self.raw_groups.to_excel(writer, sheet_name=f"Raw_groups", index=False)

        writer._save()

    def __call__(self):
        self._read_dispoview()
        self._read_groups()
        self._merge_groups_dispoview()
        self._get_stock()
        self._create_forecast()
        self._save_to_excel()
