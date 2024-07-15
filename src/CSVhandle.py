import csv
import pandas as pd
import re
import pyodbc
import os

def clean_job_label(job_label):
    first_word = re.split(r"[ ]+", job_label)[0]
    cleaned_label = re.sub(r"[^a-zA-Z]", "", first_word)
    return cleaned_label.upper()

def empSummary_to_dataFrame(filepath):
    def parse_employee_name(name):
        suffixes = ['jr', 'sr', 'ii', 'iii', 'iv', 'v']  # Add any other suffixes as needed
        parts = name.split()
        
        # Initialize the parsed components
        first_name = parts[0]
        middle_initial = ''
        last_name = ''
        suffix = ''
        
        # Check for suffix at the end
        if parts[-1].lower in suffixes:
            suffix = parts.pop()
        
        # Determine the middle and last names based on remaining parts
        if len(parts) == 2:
            last_name = parts[1]
        elif len(parts) == 3:
            if len(parts[1]) == 1:
                middle_initial = parts[1]
                last_name = parts[2]
            else:
                last_name = parts[1] + ' ' + parts[2]
        elif len(parts) == 4:
            middle_name = parts[1]
            last_name = parts[2] + ' ' + parts[3]
        elif len(parts) == 5:
            middle_name = parts[1]
            last_name = parts[2] + ' ' + parts[3] + ' ' + parts[4]

        return first_name, middle_initial, last_name, suffix
             
    df = pd.read_csv(filepath)
    df['EmployeeFName'], df['MiddleInitial'], df['EmployeeLName'], df['Suffix'] = zip(*df['Employee Name'].apply(parse_employee_name))
    
    return df

def empList_to_dataFrame(filepath):
    def extract_name_parts(name):
        split = name.split(' ')
        first = split[0]
        last = split[1]
        if len(first) == 2:
            if first[1].lower() in ['jr', 'sr', 'ii', 'iii', 'iv', 'v']:
                return first[0], '', last, first[1]
            else:
                return first[0], first[1], last, ''
        return first[0], '', last, ''
             
    df = pd.read_csv(filepath)
    df['EmployeeFName'], df['MiddleInitial'], df['EmployeeLName'], df['Suffix'] = zip(*df['Name'].apply(extract_name_parts))
    
    return df

        
def csv_to_dataFrame(filepath):
    def extract_middle_initial_and_suffix(name):
        parts = name.split()
        if len(parts) == 2:
            if parts[1].lower() in ['jr', 'sr', 'ii', 'iii', 'iv', 'v']:
                return '', parts[1]
            else:
                return parts[1], ''
        return '', ''
    
    df = pd.read_csv(filepath)
    df['MiddleInitial'], df['Suffix'] = zip(*df['EmployeeFName'].apply(extract_middle_initial_and_suffix))
    df['EmployeeFName'] = df['EmployeeFName'].apply(lambda x: x.split()[0])
    df['FullName'] = df['EmployeeLName'] + ', ' + df['EmployeeFName']
    df['FullNameNC'] = df['EmployeeFName'] + ' ' + df['EmployeeLName']
    
    return df

def csv_to_datArr(filepath):
    df = pd.read_csv(filepath)
    df['FullName'] = df['EmployeeLName'] + ', ' + df['EmployeeFName']
    df['FullNameNC'] = df['EmployeeFName'] + ' ' + df['EmployeeLName']

    job_dataframes = {}
    for jobNum, group in df.groupby('JobNumber'):
        job_dataframes[jobNum] = group

    result = {}
    for jobNum, dataframe in job_dataframes.items():
        dataframe['Start'] = pd.to_datetime(dataframe['Start'])
        dataframe['End'] = pd.to_datetime(dataframe['End'])
        day_dataframes = {}
        for date, date_group in dataframe.groupby(dataframe['Start'].dt.date):
            day_dataframes[date] = date_group
        result[jobNum] = day_dataframes
 
    output_dir = 'output_csvs'
    os.makedirs(output_dir, exist_ok=True)

    for jobNum, days in result.items():
        for date, day_df in days.items():
            date_str = date.strftime('%Y-%m-%d')
            filename = f'{output_dir}/Job_{jobNum}_Date_{date_str}.csv'
            day_df.to_csv(filename, index=False)
            
    return result

def print_datArr(result):
    for jobNum, day_dataframes in result.items():
        print(f"Job Number: {jobNum}")
        for date, dataframe in day_dataframes.items():
            print(f"  Date: {date}")
            for _, row in dataframe.iterrows():  # Iterate through each row of the dataframe
                print(f"    Employee Name: {row['FullName']}")

def csv_to_datArr2(filepath):
    df = pd.read_csv(filepath)
    df['FullName'] = df['EmployeeLName'] + ', ' + df['EmployeeFName']
    df['FullNameNC'] = df['EmployeeFName'] + ' ' + df['EmployeeLName']
    df['Start'] = pd.to_datetime(df['Start'])
    df['End'] = pd.to_datetime(df['End'])

    grouped = df.groupby(df['Start'].dt.date)
    return {date: group for date, group in grouped}

def fetch_row_from_name(conn, Fname, Lname):
    SQL_QUERY = f"""
    SELECT *
    FROM [dbo].[Employees]
    WHERE FirstName = '{Fname}' AND LastName = '{Lname}';
    """
    cursor = conn.cursor()
    cursor.execute(SQL_QUERY)
    records = cursor.fetchall()
    cursor.close()
    if len(records) < 1 :
        print(f"Error: No employee with name {Fname} {Lname}")
    elif len(records) > 1 :
        print(f"Error: Multiple employees named {Fname} {Lname}")
        print("Employee ID#:")
        print("EmployeeID\tFirstName\tLastName")
        for r in records:
            print(f"{r.EmployeeID}\t{r.FirstName}\t{r.LastName}")
        return records[0]
    else:
        return records[0]
    
def fetch_alias(conn, name):
    SQL_QUERY = f"""
    SELECT *
    FROM [dbo].[EmployeeAliases]
    WHERE ClockSharkName = '{name}';
    """
    cursor = conn.cursor()
    cursor.execute(SQL_QUERY)
    record = cursor.fetchone()
    cursor.close()
    if record == None:
        return None
    return record

def db_update_times(conn, csname, rt, ot):
    cursor = conn.cursor()

    # SQL command to reset values to NULL
    reset_query = f"""
    UPDATE [dbo].[EmployeeAliases]
    SET RegularTimeLog = {rt}, OverTimeLog = {ot}
    WHERE ClockSharkName = '{csname}';
    """
    try:
        cursor.execute(reset_query)
        conn.commit()
    except Exception as e:
        print(f"An error occurred: {e}")
        conn.rollback()
    finally:
        cursor.close()

def db_clear_times(conn):
    cursor = conn.cursor()

    # SQL command to reset values to NULL
    reset_query = """
    UPDATE [dbo].[EmployeeAliases]
    SET RegularTimeLog = 0, OverTimeLog = 0;
    """
    try:
        cursor.execute(reset_query)
        conn.commit()
        print("Values have been reset to 0 successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")
        conn.rollback()
    finally:
        cursor.close()