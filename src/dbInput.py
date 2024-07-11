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

#Update convert csv to dataframe and include middle initials and suffixes, see CSVhandle
datadf = CSVhandle.empSummary_to_dataFrame(str(root.filename))

unique_names_with_initials = datadf[['EmployeeFName', 'EmployeeLName', 'MiddleInitial', 'Suffix', 'Title']].drop_duplicates()
cursor = conn.cursor()
not_in_datadf = []
for index, row in unique_names_with_initials.iterrows():
    SQL_QUERY = """
    SELECT COUNT(*) FROM [dbo].[Employees] 
        WHERE [FirstName] = ? AND [MiddleInitial] = ? AND [LastName] = ? AND [Suffix] = ?
    """
    cursor.execute(SQL_QUERY, row['EmployeeFName'], row['MiddleInitial'], row['EmployeeLName'], row['Suffix'])
    exists = cursor.fetchone()[0]
    
    if exists:
        print(f"Employee: {row['EmployeeFName']} {row['MiddleInitial']} {row['EmployeeLName']} {row['Suffix']} already exists in the database.")
        
        update_query = """
        UPDATE [dbo].[Employees]
        SET [Title] = ?
        WHERE [FirstName] = ? AND [MiddleInitial] = ? AND [LastName] = ? AND [Suffix] = ?
        """
        try:
            cursor.execute(update_query, row['Title'], row['EmployeeFName'], row['MiddleInitial'], row['EmployeeLName'], row['Suffix'])
            print(f"Title for {row['EmployeeFName']} {row['MiddleInitial']} {row['EmployeeLName']} {row['Suffix']} updated to {row['Title']}")
        except pyodbc.Error as e:
            print(f"Error updating title for {row['EmployeeFName']} {row['MiddleInitial']} {row['EmployeeLName']} {row['Suffix']}: {e}")
    else:        
        SQL_QUERY = """
        INSERT INTO [dbo].[Employees]
            ([FirstName]
            ,[MiddleInitial]
            ,[LastName]
            ,[Suffix]
            ,[Status]
            ,[Title])
        VALUES (?, ?, ?, ?, ?, ?)
        """
        try:
            cursor.execute(SQL_QUERY, row['EmployeeFName'], row['MiddleInitial'], row['EmployeeLName'], row['Suffix'], 'Active', row['Title'])
            print(f"Employee: {row['EmployeeFName']} {row['MiddleInitial']} {row['EmployeeLName']} {row['Suffix']} added to Database")
        except pyodbc.Error as e:
            print(f"Error inserting {row['EmployeeFName']} {row['MiddleInitial']} {row['EmployeeLName']} {row['Suffix']}: {e}")
        
# Commit the changes to the database
conn.commit()

# Get all employees from the database
cursor.execute("SELECT [FirstName], [MiddleInitial], [LastName], [Suffix] FROM [dbo].[Employees]")
all_db_employees = cursor.fetchall()


# Ensure the data structure is correct
all_db_employees = [tuple(employee) for employee in all_db_employees]

# Convert the database employees to a DataFrame
db_df = pd.DataFrame(all_db_employees, columns=['EmployeeFName', 'MiddleInitial', 'EmployeeLName', 'Suffix'])

# Find employees in the database who are not in the datadf
merged_df = db_df.merge(unique_names_with_initials, how='left', on=['EmployeeFName', 'MiddleInitial', 'EmployeeLName', 'Suffix'], indicator=True)
not_in_datadf = merged_df[merged_df['_merge'] == 'left_only'][['EmployeeFName', 'MiddleInitial', 'EmployeeLName', 'Suffix']].to_dict('records')

# Close the cursor and connection
cursor.close()
conn.close()

# Return the list of people in the database who are not in the datadf
for person in not_in_datadf:
    print(person)