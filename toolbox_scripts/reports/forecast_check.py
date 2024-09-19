import os, datetime
import pandas as pd
from toolbox_scripts.read.forecast_data import ForecastDataReader


class ForecastCheck:
    def __init__(
        self,
        forecast_file_path: str,
    ):
        self.forecast_file_path = forecast_file_path
        self.all_forecast = None
        self.systems_list = None
        self.index_table = None
        self.current_fy_month = None

    def _get_data(self):
        forecast = ForecastDataReader(forecast_file_path=self.forecast_file_path)
        forecast()
        self.all_forecast = forecast.data
        self.all_forecast["System"] = self.all_forecast["System"].apply(
            lambda x: x.replace(" ", "_")
        )
        self.all_forecast["System"] = self.all_forecast["System"].apply(
            lambda x: x[:31]
        )

    def _get_systems(self):
        self.systems_list = self.all_forecast["System"].drop_duplicates().tolist()

    def _index_table(self):
        self.index_table = self.all_forecast[
            ["System", "System - status code"]
        ].drop_duplicates()

        self.index_table["System"] = self.index_table["System"].apply(
            lambda x: f'=HYPERLINK("#{x}!A1","{x}")'
        )

    def _one_system_data(self, system):
        one_system = self.all_forecast[self.all_forecast["System"] == system]
        one_system = one_system.drop(columns=["System", "System - status code"])

        data = one_system[one_system["Measure"] == "Attach Rate As Is"].copy()
        actual_pos = data.columns.get_loc(self.current_fy_month)

        current_ar = one_system[["SOI", "Measure", self.current_fy_month]].copy()
        current_ar = current_ar.loc[current_ar["Measure"] == "Attach Rate Override"]
        current_ar.rename(columns={self.current_fy_month: "Actual_AR"}, inplace=True)

        selected_months_6 = data.columns[(actual_pos - 6) : actual_pos]
        data["Average_6_mth"] = data.loc[:, selected_months_6].mean(axis=1)
        data["Median_6_mth"] = data.loc[:, selected_months_6].median(axis=1)
        mode_6 = data.loc[:, selected_months_6].mode(axis=1)
        data["Mode_6_mth"] = mode_6.iloc[:, 0]

        selected_months_3 = data.columns[(actual_pos - 3) : actual_pos]
        data["Average_3_mth"] = data.loc[:, selected_months_3].mean(axis=1)
        data["Median_3_mth"] = data.loc[:, selected_months_3].median(axis=1)
        mode_3 = data.loc[:, selected_months_3].mode(axis=1)
        data["Mode_3_mth"] = mode_3.iloc[:, 0]

        data = data.merge(current_ar, on="SOI", how="inner")

        new_column_order = [
            "SOI",
            "SOI descripion",
            "SOI - status code",
            "AV comment",
            "Actual_AR",
            "Average_3_mth",
            "Average_6_mth",
            "Median_3_mth",
            "Median_6_mth",
            "Mode_3_mth",
            "Mode_6_mth",
        ]
        data = data[new_column_order]
        data.sort_values(by=["SOI"], inplace=True)
        data.rename(columns={"SOI": f'=HYPERLINK("#Systems!A1", "SOI")'}, inplace=True)
        return data

    def _current_fy_month(self):
        now = datetime.datetime.now()
        current_month = now.strftime("%b")
        fiscal_year = now.year - 1 if now.month <= 3 else now.year
        self.current_fy_month = f"{current_month}-FY{fiscal_year % 100}"

    def _save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"Report_forecast_check_{now.strftime('%d%m%Y_%H%M')}.xlsx"
        # filename = f"Report_forecast_check_TEST.xlsx"

        directory_path = os.path.dirname(self.forecast_file_path)
        report_file_path = os.path.join(directory_path, filename)

        with pd.ExcelWriter(report_file_path, engine="xlsxwriter") as writer:
            self.index_table.to_excel(writer, sheet_name="Systems", index=False)
            self.all_forecast.to_excel(writer, sheet_name="All_forecast", index=False)
            for system in self.systems_list:
                system_view = self._one_system_data(system=system)
                system_view.to_excel(writer, sheet_name=system, index=False)

    def __call__(self):
        self._get_data()
        self._current_fy_month()
        self._get_systems()
        self._index_table()
        self._save_to_excel()
