# Call api to get task data from Lighthouse Reports

import requests
import json
import os
import sys
import time
import datetime
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os
import logging

# Set up logging
logging.basicConfig(filename='applog.txt', filemode='a', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)


def get_user_data(url,auth_token,user_agent):
    headers = {"Authorization":auth_token, "User-Agent":user_agent}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data

def get_task_data(url,auth_token):
    headers = {"Authorization":auth_token}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data

def get_monthly_data(start_date, end_date,auth_token):
    page = 1
    tasks_url = f"https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/tasks/entries?from={start_date.date()}T04%3A00%3A00.000Z&to={end_date.date()}T03%3A59%3A59.999Z&page={page}&perPage=50&sort=-createdAt"
    all_data = []
    
    while True:
        data = get_task_data(tasks_url,auth_token)
        all_data.extend(data)
        
        if len(data) < 50:
            break
        
        page += 1
        tasks_url = f"https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/tasks/entries?from={start_date.date()}T04%3A00%3A00.000Z&to={end_date.date()}T03%3A59%3A59.999Z&page={page}&perPage=50&sort=-createdAt"
    
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

def extract_tasks(start_date,end_date,auth_token,user_agent):
    # Get the absolute path of the script
    script_path = os.path.dirname(os.path.abspath(__file__))

    # Define the data folder paths
    data_folder = os.path.join(script_path, 'data/raw_tasks')
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
        'sequenceId':'ID',
        'title':'Title',
        'user':'user',
        'uid':'uid',
        'createdAt':'Datetime',
        'gps.reverseGeocoded.shortLabel':'GPS Location',
        'gps.geometry.coordinates':'LatLong',
        'area.location.name':'Location 1',
        'area.point.name':'Location 2',
    }

    df = df.rename(columns=column_mapping)
    df = df[column_mapping.values()]

    # Map username to uid
    user_url = "https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/users"

    user_data = get_user_data(user_url,auth_token,user_agent)
    user_df = pd.json_normalize(user_data)

    user_df = user_df[['user._id', 'user.fullName']].rename(columns={'user._id':'user', 'user.fullName':'Name'})

    df = df.merge(user_df, on='user', how='left')

    df.drop(columns=['user','uid'], inplace=True)

    # Convert the datetime columns to string readable format
    df['Datetime'] = pd.to_datetime(df['Datetime']).dt.strftime('%Y-%m-%d %H:%M:%S')

    df.sort_values(by='Datetime', inplace=True)

    # Save the data as a csv file in processed folder with timestamp
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f'task_data_{timestamp}.csv'
    df.to_csv(os.path.join(processed_folder,filename), index=False)
    logging.info(f"Task data saved successfully! Filename: {filename}")