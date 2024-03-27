# Call api to get exception data from Lighthouse Reports

import requests
import json
import os
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import logging

# Set up logging
logging.basicConfig(filename='applog.txt', filemode='a', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)


# Get the absolute path of the script
script_path = os.path.dirname(os.path.abspath(__file__))

# Define the data folder paths
data_folder = os.path.join(script_path, 'data/raw_exceptions')
processed_folder = os.path.join(script_path, 'data/processed')

# Check if the data folders exist, if not, create them
os.makedirs(data_folder, exist_ok=True)
os.makedirs(processed_folder, exist_ok=True)

# Get the list of existing json files in the data folder
existing_files = os.listdir(data_folder)

def get_exception_data(url,auth_token):
    headers = {"Authorization":auth_token}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data

def get_monthly_data(start_date, end_date,auth_token):
    page = 1
    exceptions_url = f"https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/exceptions?from={start_date.date()}T04%3A00%3A00.000Z&to={end_date.date()}T03%3A59%3A59.999Z&page={page}&perPage=50&sort=-createdAt"
    all_data = []
    
    while True:
        data = get_exception_data(exceptions_url,auth_token)
        all_data.extend(data)
        
        if len(data) < 50:
            break
        
        page += 1
        exceptions_url = f"https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/exceptions?from={start_date.date()}T04%3A00%3A00.000Z&to={end_date.date()}T03%3A59%3A59.999Z&page={page}&perPage=50&sort=-createdAt"
    
    return all_data

# Save data as json file
def save_data(data, filename):
    with open(filename, 'w') as f:
        json.dump(data, f)

# Load data from json file
def load_data(filename):
    with open(filename, 'r') as f:
        data = json.load(f)
    return data


def extract_exceptions(start_date, end_date,auth_token):
    # Get the absolute path of the script
    script_path = os.path.dirname(os.path.abspath(__file__))

    # Define the data folder paths
    data_folder = os.path.join(script_path, 'data/raw_exceptions')
    processed_folder = os.path.join(script_path, 'data/processed')

    # Check if the data folders exist, if not, create them
    os.makedirs(data_folder, exist_ok=True)
    os.makedirs(processed_folder, exist_ok=True)

    # Get the list of existing json files in the data folder
    existing_files = os.listdir(data_folder)

    # Loop through each month and check if the data file exists
    current_date = start_date
    while current_date < end_date:
        # Format the month and year
        month_year = current_date.strftime('%Y_%m')
        
        # Check if the data file already exists
        if f'{month_year}.json' not in existing_files:
            # Collect the data for the current month
            data = get_monthly_data(current_date, current_date + relativedelta(months=1),auth_token)
            
            # Save the data as a json file
            filename = os.path.join(data_folder, f'{month_year}.json')
            save_data(data, filename)
            logging.info(f'Saved {filename}')
        else:
            logging.info(f'{month_year}.json already exists')
        
        # Move to the next month
        current_date += relativedelta(months=1)



    # Load all data and combine into a single list
    all_data = []
    existing_files = os.listdir(data_folder)
    for file in existing_files:
        if file.endswith('.json'):
            data = load_data(os.path.join(data_folder,file))
            all_data.append(data)

    df = pd.concat([pd.json_normalize(data) for data in all_data])

    

    # Filter and rename columns
    column_mapping = {
        'opened': 'Opened Datetime',
        'duration': 'Duration',
        'resolved': 'Resolved Datetime',
        'area.shortLabel':'Location 1'
    }

    df = df.rename(columns=column_mapping)
    df = df[column_mapping.values()]

    
    # Reorder the columns
    column_order = ['Location 1','Opened Datetime','Resolved Datetime','Duration']
    df = df[column_order]

    # Convert the datetime columns to string readable format
    df['Opened Datetime'] = pd.to_datetime(df['Opened Datetime']).dt.strftime('%Y-%m-%d %H:%M:%S')
    df['Resolved Datetime'] = pd.to_datetime(df['Resolved Datetime']).dt.strftime('%Y-%m-%d %H:%M:%S')

    # Convert the duration from ms to hours
    df['Duration'] = df['Duration'] / 3600000

    # Filter out bad data
    df.drop(df[(df['Opened Datetime'] < '2023-11-01') & (df['Location 1'].str.contains('IDF'))].index, inplace=True)

    df.sort_values(by='Opened Datetime', inplace=True)

    # Save the data as a csv file in processed folder with timestamp
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f'exception_data_{timestamp}.csv'
    df.to_csv(os.path.join(processed_folder,filename), index=False)
    logging.info(f'Exception data saved successfully! Filename: {filename}')