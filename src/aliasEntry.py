import CSVhandle
import pyodbc
import pandas as pd
from tkinter import * 
from tkinter import filedialog

#Select Clockshark File
root = Tk()
root.title('Shit')
root.filename = filedialog.askopenfilename(title="Select Clockshark CSV", filetypes=(("CSV Files", "*.csv"),)) 

#Connect to database
SERVER = 'infinitysrv.database.windows.net'
DATABASE = 'infinitydb'
USERNAME = 'infinityadmin'
PASSWORD = 'Cranes24Canyon&'
connectionString = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
conn = pyodbc.connect(connectionString)
cursor = conn.cursor()

df = pd.read_csv(str(root.filename))
print("CSV file headers:")
print(df.columns.tolist())
#datadf = CSVhandle.empList_to_dataFrame(str(root.filename))

for index, row in df.iterrows():
    samsung_value = row["SAMSUNG"]
    clockshark_value = row["CLOCKSHARK - ADP"]
    position_value = row["Position"]
    shift_value = row["Shift"]
    labor_type = row["Skilled/Semi-Skilled"]
    craft_value = row["Craft"]

    # Check if the person already exists in the table
    query_check = f"""
        SELECT COUNT(*) FROM [EmployeeAliases]
        WHERE SECAIAlias = '{samsung_value}' AND ClockSharkName = '{clockshark_value}'
    """
    cursor_check = conn.cursor()
    cursor_check.execute(query_check)
    person_exists = cursor_check.fetchone()[0]

    if person_exists:
        print(f"Person with SECAIAlias='{samsung_value}' and ClockSharkName='{clockshark_value}' already exists. Inserting new values.")
        query = f"""
            UPDATE [EmployeeAliases]
            SET Position = '{position_value}', Shift = '{shift_value}', [Labor Category] = '{labor_type}', Craft = '{craft_value}'
            WHERE SECAIAlias = '{samsung_value}' AND [ClockSharkName] = '{clockshark_value}'
        """
        cursor = conn.cursor()
        cursor.execute(query)
        conn.commit()
    else:
        # Execute an SQL query to insert the values
        query_insert = f"""
            INSERT INTO [EmployeeAliases] (SECAIAlias, ClockSharkName, Position, Shift, [Labor Category], Craft)
            VALUES ('{samsung_value}', '{clockshark_value}', '{position_value}', '{shift_value}', '{labor_type}', '{craft_value}')
        """
        cursor_insert = conn.cursor()
        cursor_insert.execute(query_insert)
        conn.commit()
        print(f"Person with SECAIAlias='{samsung_value}' and ClockSharkName='{clockshark_value}' added to the table.")

# Close the connection
conn.close()

print("Data updated successfully in the SQL table.")