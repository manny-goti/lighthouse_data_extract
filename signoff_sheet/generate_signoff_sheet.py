import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import calendar
import json
from typing import Dict, List
import os
import argparse
from datetime import datetime

# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def save_default_configs():
    """Save the default configurations as JSON files."""
    production_config = {
        "Production": {
            "Weekly": {
                "description": "Sweep Floor, Clean walls/windows/door frames",
                "columns": ["1.94,1.97", "1.95", "1.18,1.20,1.50", "8.17", "8.51,8.73", "8.7"]
            },
            "Quarterly": {
                "description": "Scrub Floors",
                "columns": ["1.94,1.97", "1.95", "1.18,1.20,1.50", "8.17", "8.51,8.73", "8.7"]
            }
        }
    }
    
    warehouse_config = {
        "Warehouse": {
            "Daily": {
                "description": "Dust mop floor, empty trash/replace liners, clean corners, wipe down walls/windows/door frames, change dust mop head weekly, inspect walls/windows for damage",
                "columns": ["8.72","8.71,8.46,8.47","7.08, 7.09, 7.17, 7.19, 7.20","8.77A, 8.77B","7.00-7.03", "1.58","1.02,1.08,5.29"]
            },
            "2x Weekly": {
                "description": "Scrub/Mop Floor",
                "columns": ["8.72","8.71,8.46,8.47","7.08, 7.09, 7.17, 7.19, 7.20","8.77A, 8.77B","7.00-7.03", "1.58","1.02,1.08,5.29"]
            },
            "Monthly": {
                "description": "Deep clean storage",
                "columns": ["8.72","8.71,8.46,8.47","7.08, 7.09, 7.17, 7.19, 7.20","8.77A, 8.77B","7.00-7.03", "1.58","1.02,1.08,5.29"]
            }
        }
    }

    idf_config = {
        "IDF/Server Room": {
            "Monthly": {
                "description": "Dust and sweep/mop floors",
                "columns": ["9.93","9.71","9.48","9.44","9.20","9.05","8.74","8.68","8.59","8.33","8.07","7.14","7.04","6.54","5.04","4.01","2.38","2.24,2.25","1.92","1.68","1.53","0.42,0.43","0.97"]
            }
        }
    }
    
    # Create configs directory if it doesn't exist
    configs_dir = os.path.join(SCRIPT_DIR, 'configs')
    os.makedirs(configs_dir, exist_ok=True)
    
    # Save configurations
    with open(os.path.join(configs_dir, 'production_config.json'), 'w') as f:
        json.dump(production_config, f, indent=4)
    
    with open(os.path.join(configs_dir, 'warehouse_config.json'), 'w') as f:
        json.dump(warehouse_config, f, indent=4)

    with open(os.path.join(configs_dir, 'idf_config.json'), 'w') as f:
        json.dump(idf_config, f, indent=4)

def load_config(config_file: str) -> dict:
    """
    Load configuration from a JSON file.
    
    Args:
        config_file (str): Path to the configuration JSON file
    
    Returns:
        dict: Configuration dictionary
    """
    config_path = os.path.join(SCRIPT_DIR, config_file)
    with open(config_path, 'r') as f:
        return json.load(f)

def get_month_year_input():
    """
    Get month and year input from the user if not provided via command line.
    
    Returns:
        tuple: (month, year)
    """
    while True:
        try:
            date_str = input("Enter month and year (MM/YYYY): ")
            date = datetime.strptime(date_str, "%m/%Y")
            return date.month, date.year
        except ValueError:
            print("Invalid format. Please use MM/YYYY (e.g., 01/2025)")

def create_task_report(
    month: int,
    year: int,
    output_filename: str,
    config: Dict[str, Dict[str, List[str]]]
):
    """
    Create a task report based on provided configuration.
    
    Args:
        month (int): Month number (1-12)
        year (int): Year
        output_filename (str): Name of the output Excel file
        config (dict): Configuration dictionary loaded from JSON file
    """
    # Create a new workbook
    wb = Workbook()
    ws = wb.active
    
    # Define styles
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    center_aligned = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_aligned = Alignment(horizontal='left', vertical='center')
    
    # Calculate column positions
    current_col = 2  # Start from column B (A is for dates)
    section_spans = {}
    
    # First pass: calculate column spans and positions
    for section_name, frequencies in config.items():
        section_start = current_col
        for freq_name, freq_config in frequencies.items():
            num_columns = len(freq_config['columns'])
            section_spans[(section_name, freq_name)] = {
                'start': current_col,
                'end': current_col + num_columns - 1
            }
            current_col += num_columns
        section_spans[section_name] = {
            'start': section_start,
            'end': current_col - 1
        }
    
    # Set up the headers
    for section_name, frequencies in config.items():
        section_span = section_spans[section_name]
        start_col = get_column_letter(section_span['start'])
        end_col = get_column_letter(section_span['end'])
        
        # Merge and set section header
        ws.merge_cells(f'{start_col}1:{end_col}1')
        ws[f'{start_col}1'] = section_name
        ws[f'{start_col}1'].alignment = center_aligned
        ws[f'{start_col}1'].font = Font(bold=True)
        ws[f'{start_col}1'].border = border
        
        # Add frequency headers
        for freq_name, freq_config in frequencies.items():
            freq_span = section_spans[(section_name, freq_name)]
            start_col = get_column_letter(freq_span['start'])
            end_col = get_column_letter(freq_span['end'])
            
            # Add frequency header with description
            header_text = f"{freq_name}\n({freq_config['description']})"
            ws.merge_cells(f'{start_col}2:{end_col}2')
            ws[f'{start_col}2'] = header_text
            ws[f'{start_col}2'].alignment = center_aligned
            ws[f'{start_col}2'].font = Font(bold=True)
            
            # Add thick border if this is the last frequency in a section
            freq_border = border
            if freq_span['end'] < section_spans[section_name]['end']:
                freq_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='medium'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
            ws[f'{start_col}2'].border = freq_border
            
            # Add column headers
            for i, col_header in enumerate(freq_config['columns']):
                col = get_column_letter(freq_span['start'] + i)
                ws[f'{col}3'] = col_header
                ws[f'{col}3'].alignment = center_aligned
                ws[f'{col}3'].font = Font(bold=True)
                
                # Add thick border between task sections in the header
                header_border = border
                if (freq_span['start'] + i) == freq_span['end'] and freq_span['end'] < section_spans[section_name]['end']:
                    header_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='medium'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                ws[f'{col}3'].border = header_border
    
    # Add date header
    ws['A2'] = 'Date'
    ws['A2'].alignment = center_aligned
    ws['A2'].font = Font(bold=True)
    ws['A2'].border = border
    
    # Generate dates for the month
    num_days = calendar.monthrange(year, month)[1]
    start_date = datetime(year, month, 1)
    dates = [(start_date + timedelta(days=i)).strftime('%a, %b %-d %Y') 
             for i in range(num_days)]
    
    # Add dates and empty cells
    for i, date in enumerate(dates, start=4):  # Start from row 4 (after headers)
        # Get the day of the week (0 = Monday, 6 = Sunday)
        current_date = start_date + timedelta(days=i-4)
        day_of_week = current_date.weekday()
        
        # Create thicker bottom border for end of week (Sunday)
        current_border = border
        if day_of_week == 6:  # Sunday
            current_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='medium')  # Thicker bottom border
            )
        
        # Add date
        ws[f'A{i}'] = date
        ws[f'A{i}'].alignment = left_aligned
        ws[f'A{i}'].border = current_border
        
        # Add empty cells for all columns
        for col in range(2, current_col):
            cell = ws.cell(row=i, column=col)
            
            # Check if this column is the last column before a new task section
            is_section_boundary = False
            for section_name, frequencies in config.items():
                for freq_name, freq_config in frequencies.items():
                    freq_span = section_spans[(section_name, freq_name)]
                    if col == freq_span['end'] and col < section_spans[section_name]['end']:
                        is_section_boundary = True
            
            # Create border with thick right edge if it's a section boundary
            cell_border = current_border
            if is_section_boundary:
                cell_border = Border(
                    left=current_border.left,
                    right=Side(style='medium'),
                    top=current_border.top,
                    bottom=current_border.bottom
                )
            
            cell.border = cell_border
            cell.alignment = center_aligned
    
    # Add "Reviewed by" section
    review_row = len(dates) + 6
    ws[f'A{review_row}'] = 'Reviewed by'
    ws[f'A{review_row}'].font = Font(bold=True)
    
    # Add signature lines
    fields = ['Sign:', 'Name:', 'Date:']
    for i, field in enumerate(fields):
        current_row = review_row + i + 2
        ws[f'A{current_row}'] = field
        ws.merge_cells(f'B{current_row}:D{current_row}')
        ws[f'B{current_row}'].border = Border(bottom=Side(style='thin'))
    
    # Set column widths
    ws.column_dimensions['A'].width = 20
    for col in range(2, current_col):
        ws.column_dimensions[get_column_letter(col)].width = 12
    
    # Set row heights for headers
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 45  # Increased height for description row
    
    # Save the workbook
    wb.save(output_filename)

def normalize_string(s: str) -> str:
    """Normalize string by removing extra spaces and converting to lowercase"""
    # Remove all spaces and convert to lowercase
    return ''.join(s.lower().split())

def write_unmatched_tasks(unmatched_records: list, excel_file: str, month: int, year: int):
    """Write unmatched tasks to a review file"""
    review_dir = os.path.join(SCRIPT_DIR, 'data', 'review')
    os.makedirs(review_dir, exist_ok=True)
    
    config_type = os.path.basename(excel_file).split('_')[0]  # Get Production/Warehouse/Idf part
    review_file = os.path.join(review_dir, f'unmatched_tasks_{config_type}_{year}_{month:02d}.txt')
    
    with open(review_file, 'w') as f:
        f.write(f"Unmatched Tasks Review for {config_type} - {calendar.month_name[month]} {year}\n")
        f.write("=" * 80 + "\n\n")
        
        # Group by frequency and room number
        from collections import defaultdict
        grouped_errors = defaultdict(int)
        for error in unmatched_records:
            if "No matching column for frequency" in error:
                # Extract frequency and room from error message
                parts = error.split("'")
                if len(parts) >= 4:
                    freq = parts[1]
                    room = parts[3]
                    grouped_errors[(freq, room)] += 1
        
        # Write summary of unmatched combinations
        f.write("Summary of Unmatched Combinations:\n")
        f.write("-" * 40 + "\n")
        for (freq, room), count in sorted(grouped_errors.items()):
            f.write(f"Frequency: '{freq}', Room: '{room}' - {count} occurrences\n")
        
        # Write all individual errors
        f.write("\n\nDetailed Error List:\n")
        f.write("-" * 40 + "\n")
        for error in unmatched_records:
            f.write(f"{error}\n")
    
    print(f"\nUnmatched tasks written to: {review_file}")

def populate_report_data(excel_file: str, completion_data_csv: str, month: int, year: int):
    """
    Populate an existing task report template with completion data from a CSV file.
    
    Args:
        excel_file (str): Path to the existing Excel report template
        completion_data_csv (str): Path to the CSV file containing completion data
        month (int): Month number (1-12)
        year (int): Year (e.g., 2025)
    """
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
    
    # Create target date object
    target_date = datetime(year, month, 1)
    
    # Read the completion data
    try:
        df = pd.read_csv(completion_data_csv)
    except FileNotFoundError:
        raise FileNotFoundError(f"Completion data file not found: {completion_data_csv}")
    except pd.errors.EmptyDataError:
        raise ValueError(f"Completion data file is empty: {completion_data_csv}")
    except Exception as e:
        raise ValueError(f"Error reading completion data file: {str(e)}")
    
    # Validate required columns
    required_columns = ['Datetime', 'Name', 'Room Number', 'Frequency', 'Task Type']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns in CSV: {', '.join(missing_columns)}")
    
    # Convert datetime to proper format
    try:
        df['Datetime'] = pd.to_datetime(df['Datetime'])
    except Exception as e:
        raise ValueError(f"Error converting datetime column: {str(e)}")
    
    # Get the configuration type from filename
    filename = os.path.basename(excel_file).lower()
    if 'production' in filename:
        config_type = 'production'
        config = load_config('configs/production_config.json')
    elif 'warehouse' in filename:
        config_type = 'warehouse'
        config = load_config('configs/warehouse_config.json')
    elif 'idf' in filename:
        config_type = 'idf'
        config = load_config('configs/idf_config.json')
    else:
        raise ValueError(f"Unable to determine config type from filename: {os.path.basename(excel_file)}")
    
    # Filter data for the target month/year and config type
    df['Year/Month'] = df['Datetime'].dt.strftime('%Y-%m')
    target_year_month = target_date.strftime('%Y-%m')
    df_filtered = df[
        (df['Year/Month'] == target_year_month) & 
        (df['Task Type'].str.lower() == config_type)
    ]
    
    if len(df_filtered) == 0:
        print(f"No {config_type} tasks found for {target_date.strftime('%B %Y')} in the completion data")
        return
    
    # Load the existing Excel file
    try:
        wb = load_workbook(excel_file)
        ws = wb.active
    except Exception as e:
        raise ValueError(f"Error loading Excel template: {str(e)}")
    
    # Create a mapping of room numbers to column indices
    column_mapping = {}
    current_col = 2  # Start from column B
    
    for section_name, frequencies in config.items():
        for freq_name, freq_config in frequencies.items():
            for i, room in enumerate(freq_config['columns']):
                # Store with normalized frequency and room number
                column_mapping[(normalize_string(freq_name), normalize_string(room))] = current_col + i
            current_col += len(freq_config['columns'])
    
    # Track statistics for reporting
    total_records = len(df_filtered)
    matched_records = 0
    unmatched_records = []
    
    # Process each completion record
    for _, row in df_filtered.iterrows():
        # Format the date string to match Excel format
        completion_date = row['Datetime'].strftime('%a, %b %-d %Y')
        
        # Find the row for this date
        date_cell = None
        for row_idx in range(4, ws.max_row + 1):
            if ws.cell(row=row_idx, column=1).value == completion_date:
                date_cell = row_idx
                break
        
        if date_cell is None:
            unmatched_records.append(f"Date not found: {completion_date}")
            continue
        
        # Get task frequency and room number (normalized)
        task_freq = normalize_string(row['Frequency'])
        room_number = normalize_string(str(row['Room Number']))
        
        # Find the column for this task type and room
        column = column_mapping.get((task_freq, room_number))
        
        if column is not None:
            # Add the completion data (full name)
            cell = ws.cell(row=date_cell, column=column)
            cell.value = row['Name']
            # Center align the text and adjust column width if needed
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            # Adjust column width to fit the name
            current_width = ws.column_dimensions[get_column_letter(column)].width
            name_width = len(row['Name']) + 2  # Add some padding
            if name_width > current_width:
                ws.column_dimensions[get_column_letter(column)].width = min(name_width, 20)  # Cap at 20 characters
            matched_records += 1
        else:
            unmatched_records.append(f"No matching column for frequency '{task_freq}' and room '{room_number}'")
    
    # Save the workbook back to the same file
    try:
        wb.save(excel_file)
    except Exception as e:
        raise ValueError(f"Error saving workbook: {str(e)}")
    
    # Print summary
    print(f"\nPopulation Summary for {config_type.capitalize()} tasks - {target_date.strftime('%B %Y')}:")
    print(f"Total records processed: {total_records}")
    print(f"Successfully matched: {matched_records}")
    print(f"Unmatched records: {total_records - matched_records}")
    
    if unmatched_records:
        print("\nUnmatched Record Details:")
        for error in unmatched_records[:10]:  # Show first 10 errors
            print(f"- {error}")
        if len(unmatched_records) > 10:
            print(f"... and {len(unmatched_records) - 10} more errors")
        
        # Write unmatched tasks to review file
        write_unmatched_tasks(unmatched_records, excel_file, month, year)
    
    print(f"\nUpdated report saved: {os.path.basename(excel_file)}")

def find_latest_task_data(month: int, year: int) -> str:
    """
    Find the most recent task data file for the given month and year.
    
    Args:
        month (int): Month number (1-12)
        year (int): Year (e.g., 2025)
    
    Returns:
        str: Path to the most recent task data file
    """
    # Construct the pattern for the month/year
    pattern = f"task_data_{year}_{month:02d}_*.csv"
    processed_dir = os.path.join(SCRIPT_DIR, 'data', 'processed')
    
    # Get all matching files
    matching_files = []
    for file in os.listdir(processed_dir):
        if file.startswith(f"task_data_{year}_{month:02d}_") and file.endswith(".csv"):
            matching_files.append(os.path.join(processed_dir, file))
    
    if not matching_files:
        raise FileNotFoundError(f"No task data files found for {year}-{month:02d}")
    
    # Return the most recent file (based on filename timestamp)
    return max(matching_files)

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Generate task reports')
    parser.add_argument('--month', type=int, help='Month (1-12)')
    parser.add_argument('--year', type=int, help='Year (e.g., 2025)')
    parser.add_argument('--config', type=str, choices=['production', 'warehouse', 'idf', 'all'],
                       default='all', help='Which config to use')
    
    args = parser.parse_args()
    
    try:
        # Get month and year from command line or user input
        if args.month and args.year:
            if not (1 <= args.month <= 12):
                raise ValueError(f"Invalid month: {args.month}. Month must be between 1 and 12.")
            month = args.month
            year = args.year
        else:
            month, year = get_month_year_input()
        
        # Format month name for filename
        month_name = datetime.strptime(f"{month}/1/2000", "%m/%d/%Y").strftime("%b")
        
        # Create output directory for sign-off sheets
        output_dir = os.path.join(SCRIPT_DIR, 'data', 'signoff_sheets')
        os.makedirs(output_dir, exist_ok=True)
        
        # Find the latest task data file for the given month/year
        try:
            completion_data = find_latest_task_data(month, year)
            print(f"\nFound task data file: {os.path.basename(completion_data)}")
        except FileNotFoundError as e:
            print(f"\nWarning: {str(e)}")
            print("Will generate empty sign-off sheets.")
            completion_data = None
        
        # Generate reports based on config selection
        configs_to_run = []
        if args.config == 'all':
            configs_to_run = ['production', 'warehouse', 'idf']
        else:
            configs_to_run = [args.config]
        
        for config_name in configs_to_run:
            try:
                config = load_config(f'configs/{config_name}_config.json')
                output_filename = os.path.join(output_dir, f'{config_name.capitalize()}_Tasks_{month_name}{year}.xlsx')
                
                print(f"\nGenerating {output_filename}...")
                create_task_report(
                    month=month,
                    year=year,
                    output_filename=output_filename,
                    config=config
                )
                print(f"Created {output_filename}")
                
                # If completion data was found, populate it
                if completion_data:
                    print(f"\nPopulating {output_filename} with completion data...")
                    populate_report_data(output_filename, completion_data, month, year)
            
            except Exception as e:
                print(f"\nError processing {config_name} report: {str(e)}")
                print("Continuing with next report...")
                continue
    
    except Exception as e:
        print(f"\nError: {str(e)}")
        return 1
    
    return 0

if __name__ == '__main__':
    exit(main())