import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import pandas as pd
import os
import logging

# Set up logging
logging.basicConfig(filename='applog.txt', filemode='a', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)


def write_to_gsheets():

    # Set up path to credentials
    script_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_path, '.config/config.json')
    data_path = os.path.join(script_path, 'data/processed/')
    files = os.listdir(data_path)
    # Get the latest files based on timestamp in filename
    tasks = [file for file in files if 'task' in file]
    latest_tasks = sorted(tasks, reverse=True)[0]
    exceptions = [file for file in files if 'exception' in file]
    latest_exceptions = sorted(exceptions, reverse=True)[0]
    summary = [file for file in files if 'summary' in file]
    latest_summary = sorted(summary, reverse=True)[0]
    issues = os.listdir(os.path.join(data_path, 'issues'))
    latest_issues = sorted(issues, reverse=True)[0]


    # Set up credentials
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    gc = gspread.service_account(filename=config_path)
    sh = gc.open_by_key("1drvleU73k_jTj2Bm6x1htwMbX_2RHAtnOwIBY5gfEZY")

    data_updated = False

    # Write tasks to Google Sheets
    tasks = pd.read_csv(os.path.join(data_path,latest_tasks))
    worksheet = sh.get_worksheet(0)
    assert worksheet.title == 'Tasks'
    # If Length of the dataframe > length of the worksheet, clear the worksheet
    if len(tasks)+1 > len(worksheet.get_all_values()):
        worksheet.clear()
        set_with_dataframe(worksheet=worksheet, dataframe=tasks, include_index=False,include_column_header=True, resize=True)
        data_updated = True
        logging.info('Tasks updated in Google Sheets')
    else:
        logging.info('No update in Tasks Data')

    # Write exceptions to Google Sheets
    exceptions = pd.read_csv(os.path.join(data_path,latest_exceptions))
    worksheet = sh.get_worksheet(1)
    assert worksheet.title == 'Exceptions'
    # If Length of the dataframe > length of the worksheet, clear the worksheet
    if len(exceptions)+1 > len(worksheet.get_all_values()):
        worksheet.clear()
        set_with_dataframe(worksheet=worksheet, dataframe=exceptions, include_index=False,include_column_header=True, resize=True)
        data_updated = True
        logging.info('Exceptions updated in Google Sheets')
    else:
        logging.info('No update in Exceptions Data')

    # Write summary data to Google Sheets
    if data_updated:
        summary = pd.read_csv(os.path.join(data_path,latest_summary))
        worksheet = sh.get_worksheet(2)
        assert worksheet.title == 'Summary Data'
        worksheet.clear()
        set_with_dataframe(worksheet=worksheet, dataframe=summary, include_index=False,include_column_header=True, resize=True)
        logging.info('Summary data updated in Google Sheets')
    else:
        logging.info('No update in Summary Data')

    # Write issues data to Google Sheets
    issues = pd.read_csv(os.path.join(data_path,'issues',latest_issues))
    worksheet = sh.get_worksheet(3)
    assert worksheet.title == 'Issues'
    # If Length of the dataframe + headerRow > length of the worksheet, clear the worksheet
    if len(issues)+1 > len(worksheet.get_all_values()):
        worksheet.clear()
        set_with_dataframe(worksheet=worksheet, dataframe=issues, include_index=False,include_column_header=True, resize=True)
        logging.info('Issues updated in Google Sheets')
    else:
        logging.info('No update in Issues Data')

