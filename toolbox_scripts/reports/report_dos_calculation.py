import pandas as pd
import os, datetime


class DosCalculation:
    def __init__(self, forecast_file_path: str, groups_file_path: str) -> None:
        self.forecast_file_path = forecast_file_path
        self.groups_file_path = groups_file_path
        self.raw_forecast = None
        self.ready_dos_soi = None
        self.ready_dos_group = None
        self.raw_groups = None
        self.ready_groups = None

    def _read_forecast_data(self):
        self.raw_forecast = pd.read_excel(self.forecast_file_path)

    def _read_groups(self):
        self.raw_groups = pd.read_excel(self.groups_file_path)

    def _prepare_groups(self):
        self.ready_groups = self.raw_groups[["SOI", "GROUP"]].drop_duplicates()

    def _prepare_dos_soi(self):
        all_forecast_columns = self.raw_forecast.columns
        choosen_columns = [x for x in all_forecast_columns if "[" in x]
        choosen_columns = ["Product Code", "Measure"] + choosen_columns
        self.ready_dos_soi = self.raw_forecast[choosen_columns].copy()
        self.ready_dos_soi["Average_5mth"] = (
            self.ready_dos_soi.iloc[:, 3:8].mean(axis=1).round(0)
        )
        self.ready_dos_soi.insert(
            2, "Average_5mth", self.ready_dos_soi.pop("Average_5mth")
        )
        self.ready_dos_soi.rename(columns={"Product Code": "SOI"}, inplace=True)
        self.ready_dos_soi = pd.merge(
            self.ready_dos_soi,
            self.ready_groups,
            on="SOI",
            how="left",
        )
        self.ready_dos_soi.insert(0, "GROUP", self.ready_dos_soi.pop("GROUP"))

    def _prepare_dos_group(self):
        self.ready_dos_group = self.ready_dos_soi[["GROUP", "Average_5mth"]].copy()
        self.ready_dos_group = self.ready_dos_group.groupby(["GROUP"]).sum()
        self.ready_dos_group.reset_index(inplace=True)

    def _save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"Report_DoS_{now.strftime('%d%m%Y_%H%M')}.xlsx"

        directory_path = os.path.dirname(self.forecast_file_path)
        report_file_path = os.path.join(directory_path, filename)

        writer = pd.ExcelWriter(report_file_path)
        self.raw_forecast.to_excel(writer, sheet_name="Raw_forecast", index=False)
        self.ready_dos_soi.to_excel(writer, sheet_name="DoS_SOI", index=False)
        self.ready_dos_group.to_excel(writer, sheet_name="DoS_groups", index=False)
        writer._save()

    def __call__(self):
        self._read_forecast_data()
        self._read_groups()
        self._prepare_groups()
        self._prepare_dos_soi()
        self._prepare_dos_group()
        self._save_to_excel()
