import gooeypie as gp
import os
import pandas as pd
import shutil
from datetime import datetime
from openpyxl import load_workbook


def backup_file(file_path):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = file_path.replace('.xlsx', f'_backup_{timestamp}.xlsx')
    shutil.copyfile(file_path, backup_path)
    print(f"Backup saved as {backup_path}")


def clean_sheet(df):
    """Clean the data in the sheet."""
    df.dropna(inplace=True)  # Remove rows with any missing values
    df.drop_duplicates(inplace=True)  # Remove duplicate rows
    df.columns = [col.lower() for col in df.columns if isinstance(col, str)]  # Convert string column names to lowercase
    return df


def clean_excel(file_path):
    backup_file(file_path)  # Save a backup before making changes

    # Load the Excel file
    xls = pd.ExcelFile(file_path)
    print(xls.sheet_names)  # Debugging

    workbook = load_workbook(file_path)
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            print(df.head())  # Debugging

            # Perform cleaning operations
            df = clean_sheet(df)

            # Save the cleaned sheet back to Excel
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        except Exception as e:
            print(f"Error reading sheet {sheet_name}: {e}")

    # Save the workbook
    workbook.save(file_path)


def clean_file(event):
    file_path = file_input.text
    if file_path and os.path.isfile(file_path) and file_path.endswith('.xlsx'):
        file_name = os.path.basename(file_path)
        clean_excel(file_path)
        result_lbl.text = "File cleaned: " + file_name
    else:
        result_lbl.text = "Invalid file path. Please enter a valid Excel file path."


# Create the GUI
app = gp.GooeyPieApp('Logistikk Advanced Cleaner')
app.width = 1000

# Create the label, input field, and button
label = gp.Label(app, 'Input')
file_input = gp.Input(app)
file_btn = gp.Button(app, 'Clean File', clean_file)
result_lbl = gp.Label(app, '')
file_input.width = 100

# Add the components to the app
app.set_grid(4, 1)
app.add(label, 1, 1, align='center')
app.add(file_input, 2, 1, align='center')
app.add(file_btn, 3, 1, align='center')
app.add(result_lbl, 4, 1, align='center')

# Run the app
app.run()
