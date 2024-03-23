import pandas as pd
import numpy as np
import os
from datetime import datetime

script_path = os.path.dirname(os.path.abspath(__file__))
data_path = os.path.join(script_path, 'data/processed/')
processed_folder = os.path.join(script_path, 'data/processed')

files = os.listdir(data_path)
# Get the latest files based on timestamp in filename
tasks = [file for file in files if 'task' in file]
latest_tasks = sorted(tasks, reverse=True)[0]
exceptions = [file for file in files if 'exception' in file]
latest_exceptions = sorted(exceptions, reverse=True)[0]

# Load the data
tasks = pd.read_csv(os.path.join(data_path,latest_tasks))
tasks.rename(columns={'ID': 'Task ID'}, inplace=True)
exceptions = pd.read_csv(os.path.join(data_path,latest_exceptions))
exceptions = exceptions.reset_index(names=['Exception ID'])

# Filter out where location 2 is missing
exceptions = exceptions[~exceptions['Location 1'].isnull()]


# Convert Datetime column to datetime object
tasks['Datetime'] = pd.to_datetime(tasks['Datetime'])
exceptions['Opened Datetime'] = pd.to_datetime(exceptions['Opened Datetime'])
exceptions['Resolved Datetime'] = pd.to_datetime(exceptions['Resolved Datetime'])

# Calculate duration in hours and overwrite the column since its wrong
exceptions['Duration'] = (exceptions['Resolved Datetime'] - exceptions['Opened Datetime']).dt.total_seconds() / 3600

# Add Year/Month column
tasks['Year/Month'] = tasks['Datetime'].dt.year.astype(str)+"-"+tasks['Datetime'].dt.month.astype(str)
exceptions['OpenedYear/Month'] = exceptions['Opened Datetime'].dt.year.astype(str)+"-"+exceptions['Opened Datetime'].dt.month.astype(str)

# Add Flag for exceptions duration under 12 hours
exceptions['Under 12 Hours'] = np.where(exceptions['Duration'] < 12, 1, 0)
exceptions['Under 24 Hours'] = np.where(exceptions['Duration'] < 24, 1, 0)

# Group by Year/Month
grouped_tasks = tasks.groupby('Year/Month').agg({'Task ID': 'count'}).reset_index()
grouped_exceptions = exceptions.groupby('OpenedYear/Month').agg({'Exception ID': 'count','Under 12 Hours': 'sum','Under 24 Hours': 'sum','Duration':'median'}).reset_index()

# Join the two dataframes
joined = grouped_tasks.merge(grouped_exceptions, left_on='Year/Month', right_on='OpenedYear/Month', how='left')
joined.drop(columns=['OpenedYear/Month'], inplace=True)
joined.rename(columns={'Task ID': 'Number of Tasks', 'Exception ID': 'Number of Exceptions','Duration':'Median Time to Close'}, inplace=True)

joined['Total Tasks'] = joined['Number of Tasks'] + joined['Number of Exceptions']
joined['Completion Rate'] = joined['Number of Tasks'] / joined['Total Tasks']

# Write to processed folder
timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
filename = f'summary_data_{timestamp}.csv'
joined.to_csv(os.path.join(processed_folder,filename), index=False)

filename = f'task_data_{timestamp}.csv'
tasks.to_csv(os.path.join(processed_folder,filename), index=False)

filename = f'exception_data_{timestamp}.csv'
exceptions.to_csv(os.path.join(processed_folder,filename), index=False)


