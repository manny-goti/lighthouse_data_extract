import os
import pandas as pd

def load_latest_file(file_type: str):
    """
    Load the latest file based on timestamp in filename
    :param file_type: The type of file to load (task, exception, summary)
    """
    script_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_path, '.config/config.json')
    data_path = os.path.join(script_path, 'data/processed/')
    files = os.listdir(data_path)
    # Get the latest files based on timestamp in filename
    file_subset = [file for file in files if file_type in file]
    latest_file = sorted(file_subset, reverse=True)[0]
    data = pd.read_csv(os.path.join(data_path,latest_file))

    return data
