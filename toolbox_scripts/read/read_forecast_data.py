import pandas as pd
import os


class ForecastDataReader:
    def __init__(
        self,
        forecast_file_path: str,
    ):
        self.forecast_file_path = forecast_file_path

    def _read_data(self):
        self.data = pd.read_excel(self.forecast_file_path)
        self.data = self.data[~self.data["System SubModel"].str.contains(" CS")]

    def _drop_colmns(self):
        column_to_drop = ["SOI SubModel"]
        self.data = self.data.drop(columns=column_to_drop)

    def _rename_columns(self):
        columns_rename = {
            "System SubModel": "System",
            "System SubModel - Sales Status Code": "System - status code",
            "SOI SubModel - Product Code": "SOI",
            "SOI SubModel - Availability list -Comment": "AV comment",
            "SOI SubModel - Sales Text": "SOI descripion",
            "System SubModel Sales Status Code": "SOI - status code",
        }
        self.data = self.data.rename(columns=columns_rename)

    def _rename_months(self):
        months_rename = {}
        for column in self.data.columns[7:]:
            new_column_name = column[28:-1]
            months_rename[column] = new_column_name
        self.data.rename(columns=months_rename, inplace=True)

    def _fill_na_values(self):
        self.data.iloc[:, 6:] = self.data.iloc[:, 6:].fillna(0)

    def _test_save_to_excel(self):
        filename = "AR_TEST.xlsx"
        directory_path = os.path.dirname(self.forecast_file_path)
        report_file_path = os.path.join(directory_path, filename)
        writer = pd.ExcelWriter(report_file_path)
        self.data.to_excel(writer, sheet_name=f"all_data", index=False)
        writer._save()

    def __call__(self):
        self._read_data()
        self._drop_colmns()
        self._rename_columns()
        self._rename_months()
        self._fill_na_values()
        # self._test_save_to_excel()
