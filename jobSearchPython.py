import pandas as pd

excel_file = 'jobs.xlsx'

df = pd.read_excel(excel_file, engine='openpyxl')

# ['Company Name', 'Where', 'Applied Date', 'Deadline', 
#'Location', 'Start Date', 'End Date', 'Pay (Hourly)', 'Response']

while True:
    name = input('Please enter a Company Name or type exit: ').strip().lower()

    if name == 'exit':
        break

    if name in df['Company Name'].str.lower().values:
        row = df[df['Company Name'].str.lower() == name]
        print(row)
        changes = input('Would you like to delete, view or edit the company entry? (delete, view, edit and cancel) ').lower().strip()
        
        if changes == 'delete':
            df = df[df['Company Name'].str.lower() != name]
            df.to_excel(excel_file, index=False, engine='openpyxl')
            print(f'{name} has been removed')
        
        elif changes == 'edit':
            columns = df.columns.tolist()    
            print(columns)        
            
            while True:
                columnChoice = input(f'Enter the data you want to change or type exit: ').lower().strip()
                
                if columnChoice == 'exit':
                    break

                if columnChoice in df.columns.str.lower():
                    actual_column = df.columns[df.columns.str.lower() == columnChoice][0]
                    
                    current_value = row[actual_column].values[0]
                    print(f"Current value for {actual_column}: {current_value}")
                    
                    new_value = input(f"Enter the new value for {actual_column}: ").strip()
                    
                    df.loc[df['Company Name'].str.lower() == name, actual_column] = new_value
                    
                    df.to_excel(excel_file, index=False, engine='openpyxl')
                    print(f"The value for '{actual_column}' has been updated to '{new_value}'.")
                else:
                    print(f"'{columnChoice}' is not a valid column name.")
        elif changes == 'view':
            for column, value in row.iloc[0].items():
                print(f"{column}: {value}")
        else:
            print('No changes have been made')


    else:
        print(f"The company '{name}' does not exist.")
        addCompany = input("Do you want to add it? (yes/no): ").strip().lower()
        if addCompany == 'yes' or addCompany == 'y' or addCompany == 'add':
            addPosition = input('Enter position name: ')
            addWhere = input('Enter a where you applied: ')
            addDocument = input('Enter documents submitted: ')
            addAppDate = input('Enter the date you applied: ')
            addDeadline = input('Enter the deadline for the application: ')
            addLocation = input('Enter the location for the application: ')
            addStartDate = input('Enter the Start Date for the job: ')
            addEndDate = input('Enter the End Date for the job: ')
            addPay = input('Enter the pay (Hourly): ')
            addResponse = input('Recieved a response?: ')
            new_entry = pd.DataFrame({'Company Name': [name], 'Position Name':[addPosition],
                                    'Where': [addWhere], 'Documents': [addDocument],
                                    'Applied Date': [addAppDate], 
                                    'Deadline': [addDeadline], 'Location': [addLocation], 
                                    'Start Date': [addStartDate],'End Date': [addEndDate], 
                                    'Pay': [addPay], 'Response': [addResponse]})
            
            df = pd.concat([df, new_entry], ignore_index=True)
            df.to_excel(excel_file, index=False, engine='openpyxl')
            print(f"Company '{name}' has been added")
        else:
            print('No changes have been made')