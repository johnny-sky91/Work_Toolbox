import os, shutil


def sort_my_data(directory):
    dir_mapping = {
        "sap_log_data_": "Downloads\sap_dispoview",
        "Report_dispoview_groups_": "Results\Dispoview_groups",
        "EMS_forecast_": "Results\EMS_forecast",
        "Report_forecast_check_": "Results\Forecast_check",
        "AttachRatePlanning": "Downloads\AR_data",
    }
    files = os.listdir(directory)
    for file in files:
        for file_type in dir_mapping:
            if file_type in file:
                destination_directory = os.path.join(directory, dir_mapping[file_type])
                source_file_path = os.path.join(directory, file)
                shutil.move(source_file_path, destination_directory)
