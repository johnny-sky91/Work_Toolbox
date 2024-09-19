import os, datetime

import pandas as pd

from toolbox_scripts.read.dispoview_data import DispoviewDataReader


class GroupsOverview:
    def __init__(
        self,
        groups_file_path: str,
        dispo_file_path: str,
        supply_file_path: str,
    ):
        self.groups_file_path = groups_file_path
        self.dispo_file_path = dispo_file_path
        self.supply_file_path = supply_file_path

        self.groups = None
        self.groups_list = None
        self.groups_min = None
        self.index_table = None
        self.dispo_data = None
        self.supply_all = None

    def _read_groups_data(self):
        self.groups = pd.read_excel(self.groups_file_path)
        self.groups_list = self.groups["GROUP"].drop_duplicates().tolist()

    def _read_dispo_data(self):
        dispoview = DispoviewDataReader(dispo_file_path=self.dispo_file_path)
        dispoview()
        self.dispo_data = dispoview.ready_dispoview

    def _read_supply_data(self):
        self.supply_all = pd.read_excel(
            self.supply_file_path, sheet_name="supply_confirmed"
        )

    def _merge_codenumbers_supply(self):
        codenumbers = self.groups[["CODENUMBER", "COMPONENT"]].drop_duplicates()
        self.supply_all = pd.merge(
            codenumbers,
            self.supply_all,
            on="COMPONENT",
            how="right",
        )

    def _filter_dispoview(self):
        data_needed = [
            "Stock",
            "CustOrders CDD",
            "CustOrders RDD",
            "NetForecast",
            "Supply LP",
        ]
        self.dispo_data = self.dispo_data.loc[self.dispo_data["DATA"].isin(data_needed)]

    def _mergre_groups_dispoview(self):
        groups_unique = self.groups[
            ["CODENUMBER", "GROUP", "GROUP_DESCRIPTION"]
        ].drop_duplicates()
        self.dispo_data = pd.merge(
            groups_unique,
            self.dispo_data,
            on="CODENUMBER",
            how="right",
        )

    def _index_sheet(self):
        self.index_table = self.groups[
            ["GROUP", "GROUP_DESCRIPTION", "GROUP_STATUS"]
        ].drop_duplicates()
        self.index_table["GROUP"] = self.index_table["GROUP"].apply(
            lambda x: f'=HYPERLINK("#{x}!A1","{x}")'
        )

    def _one_group_bom(self, group):
        one_bom = self.groups[self.groups["GROUP"].isin([group])]
        one_bom = one_bom[
            ["SOI", "SOI_status", "CODENUMBER", "COMPONENT_status", "USAGE"]
        ]
        one_bom = one_bom.pivot_table(
            index=["SOI_status", "SOI"],
            columns=["COMPONENT_status", "CODENUMBER"],
            values="USAGE",
            fill_value=0,
            aggfunc="sum",
        ).reset_index()
        one_bom = one_bom.set_index("SOI", drop=True)
        # one_bom.index.names = ['=HYPERLINK("#Groups!A1", "SOI")']

        return one_bom

    def _one_group_dispo(self, group):
        one_dispo = self.dispo_data[self.dispo_data["GROUP"].isin([group])]
        one_dispo = one_dispo.drop(columns=["GROUP", "GROUP_DESCRIPTION"])
        one_dispo = one_dispo.sort_values(by=["DATA", "CODENUMBER"])
        one_dispo = one_dispo.set_index("CODENUMBER", drop=True)
        return one_dispo

    def _one_group_supply(self, group):
        one_supply = self.supply_all[self.supply_all["GROUP"].isin([group])]
        one_supply = one_supply.set_index("CODENUMBER", drop=True)
        return one_supply

    def _save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"Report_groups_overview_{now.strftime('%d%m%Y_%H%M')}.xlsx"
        # filename = f"Report_groups_overview_TEST.xlsx"
        directory_path = os.path.dirname(self.dispo_file_path)
        report_file_path = os.path.join(directory_path, filename)

        navigation = pd.DataFrame(
            {"Navigate ->": [], '=HYPERLINK("#Groups!A1", "Groups")': []}
        )
        navigation = navigation.set_index("Navigate ->", drop=True)

        with pd.ExcelWriter(report_file_path, engine="xlsxwriter") as writer:
            self.index_table.to_excel(writer, sheet_name="Groups", index=False)
            self.dispo_data.to_excel(writer, sheet_name="Dispo_data", index=False)
            for group in self.groups_list:
                one_group_bom = self._one_group_bom(group)
                one_dispo = self._one_group_dispo(group)
                one_supply = self._one_group_supply(group)

                pos_navigation = 0
                pos_one_group_bom = navigation.shape[0] + 2
                pos_one_dispo = pos_one_group_bom + one_group_bom.shape[0] + 4
                pos_one_supply = pos_one_dispo + one_dispo.shape[0] + 2

                navigation.to_excel(
                    writer,
                    startrow=pos_navigation,
                    sheet_name=group,
                    index=True,
                )
                one_group_bom.to_excel(
                    writer,
                    startrow=pos_one_group_bom,
                    sheet_name=group,
                    index=True,
                )
                if one_dispo.empty:
                    pass
                else:
                    one_dispo.to_excel(
                        writer,
                        startrow=pos_one_dispo,
                        sheet_name=group,
                        index=True,
                    )
                if one_supply.empty:
                    pass
                else:
                    one_supply.to_excel(
                        writer,
                        startrow=pos_one_supply,
                        sheet_name=group,
                        index=True,
                    )

    def __call__(self):
        self._read_groups_data()
        self._read_dispo_data()
        self._read_supply_data()
        self._merge_codenumbers_supply()
        self._filter_dispoview()
        self._mergre_groups_dispoview()
        self._index_sheet()
        self._save_to_excel()
