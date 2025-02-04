import logging
from extract_tasks import extract_tasks
from extract_exceptions import extract_exceptions
from extract_issues import extract_issues
from process_data import process_lighthouse_data
from write_to_google import write_to_gsheets
from datetime import datetime, timedelta
import os
import json

## Cron Job:
# 0 0 2 * * /home/mgoti/anaconda3/envs/lighthouse_api/bin/python /home/mgoti/Projects/ifp_report/main.py


# Set up logging with time and date
logging.basicConfig(filename='applog.txt', filemode='a', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Load auth token and user agent from json file
with open('./config/lighthouse.json') as f:
    data = json.load(f)
    auth_token = data['auth_token']
    user_agent = data['user_agent']

# Define the start and end dates for data collection
start_date = datetime.strptime('2023-10-01', '%Y-%m-%d')
end_date = datetime.now().replace(day=1) - timedelta(days=1)

def main():
    logging.info('**************Starting IFP Data Extraction/Update**************')
    extract_tasks(start_date, end_date,auth_token,user_agent)
    extract_exceptions(start_date, end_date,auth_token)
    extract_issues(start_date, end_date,auth_token,user_agent)
    process_lighthouse_data()
    # write_to_gsheets()
    logging.info('**************IFP Data Extraction/Update Complete**************')
    
    # Simple email without attachment
    logging.info('**************Sending Email Notification**************')
    os.system('echo "IFP Data Extraction/Update Complete" | mail -s "IFP Data Extraction/Update Complete" -A applog.txt manny@mgcapitalmain.com')
    logging.info('**************Email Notification Sent**************')


if __name__ == "__main__":
    main()