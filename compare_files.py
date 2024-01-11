from tkinter import filedialog, messagebox
import numpy as np
import pandas as pd


def select_file(title="Select a file"):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx;*.xls")])


def compare_files():
    # Define the columns you want to keep for missing data
    columns_to_keep = [
        'ORG_CODE', 'CUSTOMER_NAME', 'COMMODITY_CODE', 'IPN', 'PRIME_MPN_MFG', 'PRIME_MPN',
        'DESCRIPTION', 'ITEM_STATUS', 'BUYER', 'ITEM_TYPE', 'ABC_CLASS', 'PURCHASING_LT',
        'FACTORY_LT', 'MOQ', 'MPQ', 'NCNR', 'ONHAND_QTY', 'CONSIGNED_QTY', 'ONORDER_QTY',
        'PFEP_PROGRAM', 'BOND_QTY', 'BOND_OWNER', 'ATS_QTY', 'ATS_OWNER', 'SOURCED',
        'AWARD_DATE', 'SOURCING_VENDOR', 'NET_FCST_DMD_QTY', 'THIRTY_DAY_DMD_QTY',
        'SIXTY_DAY_DMD_QTY', 'NINETY_DAY_DMD_QTY', 'ANNUAL_DMD_QTY', 'LAST_CONSUMPTION',
        'BOND_NEED', 'DAY30', 'PASS30', 'PASS60', 'PASS90', 'VISIBILITY'
    ]

    # Prompt user to select last week's file
    last_week_file = select_file("Select last week's file")
    if not last_week_file:  # Check if a file was selected
        return

    # Prompt user to select current week's file
    current_week_file = select_file("Select the new file")
    if not current_week_file:  # Check if a file was selected
        return

    # Load the selected files
    last_week_xls = pd.ExcelFile(last_week_file)
    current_week_xls = pd.ExcelFile(current_week_file)

    # Function to find sheet name (case-insensitive)
    def find_sheet_name(xls, sheet_name_to_find):
        for sheet in xls.sheet_names:
            if sheet.lower() == sheet_name_to_find.lower():
                return sheet
        return None  # Return None if not found

    # Find 'detail' sheet in both files (case-insensitive)
    last_week_detail_sheet = find_sheet_name(last_week_xls, 'detail')
    current_week_detail_sheet = find_sheet_name(current_week_xls, 'detail')

    # Check if the 'detail' sheet was found in both files
    if not last_week_detail_sheet or not current_week_detail_sheet:
        messagebox.showerror("Error", "The 'detail' sheet was not found in one or both files.")
        return

    # Read the 'detail' sheets
    last_week_data = pd.read_excel(last_week_file, sheet_name=last_week_detail_sheet)
    current_week_data = pd.read_excel(current_week_file, sheet_name=current_week_detail_sheet)

    # Convert column names to uppercase
    last_week_data.columns = last_week_data.columns.str.upper()
    current_week_data.columns = current_week_data.columns.str.upper()

    # Debug prints
    print("Unique IPNs in last week's data:", last_week_data['IPN'].nunique())
    print("Unique IPNs in current week's data:", current_week_data['IPN'].nunique())

    # Convert IPNs in both DataFrames to strings and remove any leading or trailing white spaces
    last_week_data['IPN'] = last_week_data['IPN'].astype(str).str.strip()
    current_week_data['IPN'] = current_week_data['IPN'].astype(str).str.strip()

    # Filter last week's data, and this is where the Removed from last week data is compared
    last_week_data_filtered = last_week_data[~last_week_data['IPN'].isin(current_week_data['IPN'])]

    # Debug printsinstead
    print("Rows in last_week_data:", len(last_week_data))
    print("Rows in last_week_data_filtered:", len(last_week_data_filtered))

    # Merge the necessary columns only for comparison
    merged_data = pd.merge(last_week_data[['IPN', 'ITEM_TYPE', 'SOURCED']],
                           current_week_data[['IPN', 'ITEM_TYPE', 'SOURCED']],
                           on='IPN', suffixes=('', '_This_Week'), how='inner', indicator=True)

    # Debug prints
    print(merged_data['_merge'].value_counts())

    # Detect changes in ITEM_TYPE
    condition_item = (merged_data['ITEM_TYPE'] != merged_data['ITEM_TYPE_This_Week'])
    merged_data['ITEM_TYPE_CHANGED_FROM'] = np.where(condition_item, merged_data['ITEM_TYPE'], "")

    # Detect changes in SOURCED
    condition_sourced = (merged_data['SOURCED'] != merged_data['SOURCED_This_Week'])
    merged_data['SOURCE_TYPE_CHANGED_FROM'] = np.where(condition_sourced, merged_data['SOURCED'], "")

    # Debug prints
    print("Detected ITEM_TYPE changes:", merged_data['ITEM_TYPE_CHANGED_FROM'].nunique())
    print("Detected SOURCE_TYPE changes:", merged_data['SOURCE_TYPE_CHANGED_FROM'].nunique())

    # Update current_week_data with the new columns for changes
    current_week_data = pd.merge(current_week_data,
                                 merged_data[['IPN', 'ITEM_TYPE_CHANGED_FROM', 'SOURCE_TYPE_CHANGED_FROM']],
                                 on='IPN', how='left')

    # Debug prints
    print(current_week_data['ITEM_TYPE_CHANGED_FROM'].value_counts())
    print(current_week_data['SOURCE_TYPE_CHANGED_FROM'].value_counts())

    # Identify rows that were in last week's file but not in the current week's file
    missing_data = last_week_data[~last_week_data['IPN'].isin(current_week_data['IPN'])][columns_to_keep]

    # Update Excel with the new data
    with pd.ExcelWriter(current_week_file, engine='openpyxl', mode='a') as writer:
        book = writer.book

        # # Debug prints
        # print("ITEM_TYPE_CHANGED_TO non-empty rows:",
        #       current_week_data[current_week_data['ITEM_TYPE_CHANGED_FROM'] != ""].shape[0]
        #
        # print("SOURCED_TYPE_CHANGED_TO non-empty rows:",
        #       current_week_data[current_week_data['SOURCED_TYPE_CHANGED_TO'] != ""].shape[0])

        # Update the 'detail' sheet
        if 'detail' in book.sheetnames:
            book.remove(book['detail'])
        current_week_data.to_excel(writer, index=False, sheet_name='detail')

        # Write the 'Removed From Prev File' sheet
        if 'Removed From Prev File' in book.sheetnames:
            book.remove(book['Removed From Prev File'])
        missing_data.to_excel(writer, sheet_name='Removed From Prev File', index=False)

    # Show success message
    messagebox.showinfo("Success",
                        "Last Week's File and This Week's File have been analyzed and compared. "
                        "The 'detail' sheet has been updated with the 'ITEM_TYPE_CHANGED_TO' column. "
                        "You can also see your removed data in the 'Removed From Prev File' sheet.")
