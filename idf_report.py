import pandas as pd
import numpy as np
from utils import load_latest_file
import gspread
from gspread_dataframe import set_with_dataframe
import os

try:
    tasks = load_latest_file('task')
    tasks['Datetime'] = pd.to_datetime(tasks['Datetime'])

    previous_month_dt = (pd.Timestamp.now() - pd.DateOffset(months=1))
    previous_month = previous_month_dt.month
    previous_month_yr = f"{previous_month_dt.year}-{previous_month}"

    # Filter data to previous month
    tasks = tasks[tasks['Year/Month'] == previous_month_yr]

    # Filter to IDF tasks
    tasks = tasks[tasks['Title'].str.contains('IDF')]
    tasks = tasks[tasks['Location 2'].str.contains('IDF')]

    # Clean Data
    tasks = tasks[tasks['Location 2'].notnull()]
    tasks = tasks[tasks['Location 2'] != ""]

    # Remove duplicates
    tasks.drop_duplicates(subset=['Location 2'],inplace=True)

    # Write to Gsheets
    script_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_path, '.config/config.json')
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    gc = gspread.service_account(filename=config_path)
    sh = gc.open_by_key("1_Xlce5mbRHktQ9BQggwX-FnbOLradpQX0AutjaaDyOk")
    # Make new worksheet if doesn't exist
    if previous_month_yr not in [ws.title for ws in sh.worksheets()]:
        worksheet = sh.add_worksheet(title=f"{previous_month_yr}", rows="100", cols="20")
    else:
        worksheet = sh.worksheet(f"{previous_month_yr}")
    worksheet.clear()
    assert tasks.shape[0] > 0
    set_with_dataframe(worksheet=worksheet, dataframe=tasks, include_index=False,include_column_header=True, resize=True)
    os.system('echo "IDF Report Update Complete" | mail -s "IDF Report Update Complete" manny@mgcapitalmain.com')
except Exception as e:
    #send email with error code
    os.system(f'echo "IDF Report Update Failed with error: {e}" | mail -s "IDF Report Update Failed" manny@mgcapitalmain.com')
