import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe
import pandas as pd
import os
import logging
from datetime import datetime
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

def create_monthly_report(month_year=None,gsheet_key=None):
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

    if month_year is None:
        today = datetime.now()
        month_year = f"{today.year}-{today.month:02d}"
        month_year_title = today.strftime("%B %Y") # For use in the Google Sheet title
    else:
        month_year_title = datetime.strptime(month_year, "%Y-%m").strftime("%B %Y")

    # Set up credentials
    gc = gspread.service_account(filename=config_path)

    if gsheet_key is None:
        sh = gc.create(f'Exhibit A: IFP Monthly Report {month_year_title}')
        sh.share('manny@mgcapitalmain.com', perm_type='user', role='writer')
    else:
        sh = gc.open_by_key(gsheet_key)

    # Write tasks to Google Sheets
    tasks = pd.read_csv(os.path.join(data_path,latest_tasks))
    tasks = tasks[tasks['Year/Month'] == month_year]
    tasks.drop(columns=['Year/Month','GPS Location','LatLong'], inplace=True)
    # Create a new worksheet for tasks
    worksheet = sh.add_worksheet(title='Tasks', rows=len(tasks), cols=len(tasks.columns))
    assert worksheet.title == 'Tasks'
    set_with_dataframe(worksheet=worksheet, dataframe=tasks, include_index=False,include_column_header=True, resize=True)
    
    # Write exceptions to Google Sheets
    exceptions = pd.read_csv(os.path.join(data_path,latest_exceptions))
    # Create a new worksheet for exceptions
    worksheet = sh.add_worksheet(title='Exceptions', rows=len(exceptions), cols=len(exceptions.columns))
    assert worksheet.title == 'Exceptions'
    set_with_dataframe(worksheet=worksheet, dataframe=exceptions, include_index=False,include_column_header=True, resize=True)
    
    # Write summary data to Google Sheets
    summary = pd.read_csv(os.path.join(data_path,latest_summary))
    # Create a new worksheet for summary
    worksheet = sh.add_worksheet(title='Summary Data', rows=len(summary), cols=len(summary.columns))    
    assert worksheet.title == 'Summary Data'
    set_with_dataframe(worksheet=worksheet, dataframe=summary, include_index=False,include_column_header=True, resize=True)
    
    # Write issues data to Google Sheets
    issues = pd.read_csv(os.path.join(data_path,'issues',latest_issues))
    # Create a new worksheet for issues
    worksheet = sh.add_worksheet(title='Issues', rows=len(issues), cols=len(issues.columns))
    assert worksheet.title == 'Issues'
    set_with_dataframe(worksheet=worksheet, dataframe=issues, include_index=False,include_column_header=True, resize=True)
    
    # Delete Sheet1
    sh.del_worksheet(sh.get_worksheet(0))

    return sh.url


if __name__ == "__main__":
    print(create_monthly_report())