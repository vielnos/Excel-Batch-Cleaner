#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import shutil
import win32com.client as win32

def get_sheet_config():
    """
    Ask the user to input sheet names and number of top rows to delete.
    The same config will be applied to all files.
    Returns a dictionary: {sheet_name: rows_to_delete}
    """
    sheet_config = {}
    print("Enter sheet names and number of top rows to delete for ALL files in the folder.")
    print("Type 'done' when finished.\n")
    while True:
        sheet_name = input("Sheet name (or 'done' to finish): ").strip()
        if sheet_name.lower() == 'done':
            break
        if not sheet_name:
            print("Sheet name cannot be empty!")
            continue
        try:
            rows_to_delete = int(input(f"Number of top rows to delete for '{sheet_name}': ").strip())
            sheet_config[sheet_name] = rows_to_delete
        except ValueError:
            print("Please enter a valid integer for rows to delete.")
    return sheet_config

def clean_excel_files(input_folder, output_folder, sheet_config):
    """
    Clean top rows and shapes from Excel sheets in all Excel files from input_folder
    and save cleaned files in output_folder.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    files = [f for f in os.listdir(input_folder) if f.endswith((".xlsx", ".xls"))]
    if not files:
        print("No Excel files found in the input folder.")
        return

    for filename in files:
        input_path = os.path.join(input_folder, filename)
        output_path = os.path.join(output_folder, filename)
        try:
            wb = excel.Workbooks.Open(input_path)
            cleaned_any_sheet = False

            for sheet_name, rows_to_delete in sheet_config.items():
                try:
                    ws = wb.Sheets(sheet_name)

                    # Remove shapes in top rows
                    for shape in ws.Shapes:
                        try:
                            if shape.TopLeftCell.Row <= rows_to_delete:
                                shape.Delete()
                        except Exception as shape_err:
                            print(f"Could not delete shape in '{sheet_name}' of {filename}: {shape_err}")

                    # Delete top N rows
                    ws.Range(f"1:{rows_to_delete}").Delete()
                    cleaned_any_sheet = True
                    print(f"Cleaned '{sheet_name}' in {filename}")

                except Exception:
                    print(f"Sheet '{sheet_name}' not found in {filename}")

            if cleaned_any_sheet:
                wb.SaveAs(output_path)
                print(f"Saved cleaned file to: {output_path}")
            wb.Close(SaveChanges=False)

        except Exception as file_err:
            print(f"Failed to open {filename}: {file_err}")

    excel.Quit()
    print("\nAll sheets processed successfully!")

if __name__ == "__main__":
    print("=== Excel Batch Sheet Cleaner ===\n")
    input_folder = input("Enter the input folder path: ").strip()
    output_folder = input("Enter the output folder path: ").strip()
    sheet_config = get_sheet_config()

    if not sheet_config:
        print("No sheets configured. Exiting.")
    else:
        clean_excel_files(input_folder, output_folder, sheet_config)

