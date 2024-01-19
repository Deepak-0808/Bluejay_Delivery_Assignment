import pandas as pd

def parse_timedelta_string(value):
    try:
        # Handle float values separately
        if isinstance(value, float):
            return pd.NaT

        # Assuming the format is 'hh:mm:ss'
        hours, minutes, seconds = map(int, value.split(':'))
        return pd.Timedelta(hours=hours, minutes=minutes, seconds=seconds)
    except ValueError:
        # Handle the case where the format is not as expected
        return pd.NaT

def analyze_excel_file(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Check if the required columns exist in the DataFrame
    # I assume 'Time' is 'Time In'
    required_columns = ['Employee Name', 'Position ID', 'Time', 'Time Out', 'Timecard Hours (as Time)']
    missing_columns = set(required_columns) - set(df.columns)

    if missing_columns:
        print(f"Error: Missing columns in the Excel file: {', '.join(missing_columns)}")
        return

    # Assuming the 'Time' and 'Time Out' columns are in datetime format
    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])

    # Convert 'Timecard Hours (as Time)' to Timedelta using the custom parser
    df['Timecard Hours (as Time)'] = df['Timecard Hours (as Time)'].apply(parse_timedelta_string)

    # Calculate the time between shifts
    df['Time Between Shifts'] = df['Time'].shift(-1) - df['Time Out']

    # Filter employees based on the specified criteria
    consecutive_days = df.groupby('Employee Name').filter(lambda group: group['Time'].diff().dt.days.eq(1).all())
    
    # Handle NaN values in 'Timecard Hours (as Time)' for comparison
    long_shifts = df[df['Timecard Hours (as Time)'].fillna(pd.Timedelta(seconds=0)) > pd.Timedelta(hours=14)]

    # Calculate short_time_between_shifts
    short_time_between_shifts = df[(df['Time Between Shifts'] < pd.Timedelta('10 hours')) & (df['Time Between Shifts'] > pd.Timedelta('1 hour'))]

    # Print the results
    print("Employees who have worked for 7 consecutive days:")
    print(consecutive_days[['Employee Name', 'Position ID']])
    print("\nEmployees who have less than 10 hours of time between shifts but greater than 1 hour:")
    print(short_time_between_shifts[['Employee Name', 'Position ID']])
    print("\nEmployees who have worked for more than 14 hours in a single shift:")
    print(long_shifts[['Employee Name', 'Position ID']])

# Replace 'your_file.xlsx' with the actual file path
file_path = 'Assignment_Timecard.xlsx'
analyze_excel_file(file_path)
