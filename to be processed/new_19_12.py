from API.google_sheets_api import DOWNLOAD_DATA, UPDATE_VALUES

def CONVERT_TO_LIST_OF_LISTS(df):
    return df.to_numpy().tolist()


def APPEND_DF_TO_SHEET(sheet_id, upload_data, col_start, col_end, tab_name, api_info):
    sheet_data = DOWNLOAD_DATA(sheet_id, tab_name + col_start + "1:" + col_end, api_info)
    last_row = len(sheet_data)
    missing_truck_tab_new_row_start = last_row + 1

    #replace nan values
    fill_nan = upload_data.fillna("N/A")

    upload_values = CONVERT_TO_LIST_OF_LISTS(fill_nan)
    range_upload = tab_name + col_start + str(missing_truck_tab_new_row_start) + ":"  + col_end

    UPDATE_VALUES(sheet_id, range_upload, upload_values, api_info)