# Call api to get exception data from Lighthouse Reports

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
auth_token = '6482338a76be3774d91cc9b0_31de04c6-0a75-4e5f-a052-b71416c97bdd'
user_url = "https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/users"
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"

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

# Rest of the code...
def get_user_data(url):
    headers = {"Authorization":auth_token, "User-Agent":user_agent}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data

def get_exception_data(url):
    headers = {"Authorization":auth_token}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data

def get_monthly_data(start_date, end_date):
    page = 1
    exceptions_url = f"https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/exceptions?from={start_date}&to={end_date}&page={page}&perPage=50&sort=-createdAt"
    all_data = []
    
    while True:
        data = get_exception_data(exceptions_url)
        all_data.extend(data)
        
        if len(data) < 50:
            break
        
        page += 1
        exceptions_url = f"https://api.us.lighthouse.io/applications/64788f940b7634da94ffdcbd/exceptions?from={start_date}&to={end_date}&page={page}&perPage=50&sort=-createdAt"
    
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

# Define the start and end dates for data collection
start_date = datetime.strptime('2023-10-01', '%Y-%m-%d')
end_date = datetime.now().replace(day=1) - timedelta(days=1)

# Loop through each month and check if the data file exists
current_date = start_date
while current_date < end_date:
    # Format the month and year
    month_year = current_date.strftime('%Y_%m')
    
    # Check if the data file already exists
    if f'{month_year}.json' not in existing_files:
        # Collect the data for the current month
        data = get_monthly_data(current_date, current_date + relativedelta(months=1))
        
        # Save the data as a json file
        filename = os.path.join(data_folder, f'{month_year}.json')
        save_data(data, filename)
        print(f'Saved {filename}')
    else:
        print(f'{month_year}.json already exists')
    
    # Move to the next month
    current_date += relativedelta(months=1)



# Load all data and combine into a single list
all_data = []
existing_files = os.listdir(data_folder)
for file in existing_files:
    if file.endswith('.json'):
        data = load_data(os.path.join(data_folder,file))
        all_data.extend(data)

len(all_data)

# Convert json data to dataframe
df = pd.json_normalize(data)

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
user_data = get_user_data(user_url)
user_df = pd.json_normalize(user_data)

user_df = user_df[['user._id', 'user.fullName']].rename(columns={'user._id':'user', 'user.fullName':'Name'})

df = df.merge(user_df, on='user', how='left')

df.drop(columns=['user','uid'], inplace=True)

# Save the data as a csv file in processed folder with timestamp
timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
filename = f'exception_data_{timestamp}.csv'
data_folder = './data/processed'
df.to_csv(os.path.join(data_folder,filename), index=False)