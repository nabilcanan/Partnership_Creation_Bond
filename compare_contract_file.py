import pandas as pd
from compare_files import select_file
from tkinter import messagebox


def compare_contract_file():
    # Prompt user to select file with 'Removed From Prev File' sheet
    missing_ipn_file = select_file("Choose your new file with the Removed From Prev File Sheet")

    # Check sheet names in the file
    xls = pd.ExcelFile(missing_ipn_file)
    print(xls.sheet_names)

    # Prompt user to select active supplier contracts file
    contract_file = select_file("Choose the Active Supplier Contracts file that came in that week")

    # Load 'Missing From Last Week' data into a pandas DataFrame
    missing_data = pd.read_excel(missing_ipn_file, sheet_name='Removed From Prev File')

    # Load details data
    detail_data = pd.read_excel(missing_ipn_file, sheet_name='detail')
    detail_data['IPN'] = detail_data['IPN'].astype(str).str.strip()

    # Load active supplier contracts data into a pandas DataFrame
    contract_data = pd.read_excel(contract_file, skiprows=1)
    contract_data['IPN'] = contract_data['IPN'].astype(str).str.strip()

    # Normalize IPN values by removing leading zeros in both datasets
    detail_data['Normalized_IPN'] = detail_data['IPN'].apply(lambda x: x.lstrip('0'))
    contract_data['Normalized_IPN'] = contract_data['IPN'].apply(lambda x: x.lstrip('0'))

    # Merge detail_data with contract_data on Normalized_IPN
    merged_data = pd.merge(detail_data, contract_data[['Normalized_IPN', 'Price']],
                           left_on='Normalized_IPN', right_on='Normalized_IPN', how='left')

    # Format the Corporate Contract Price column
    merged_data['Corporate Contract Price'] = merged_data['Price'].apply(
        lambda x: "{:.4f}".format(x) if pd.notnull(x) else "")

    # Drop the columns not needed including the temporary Normalized_IPN and Price columns
    merged_data.drop(columns=['Normalized_IPN', 'Price'], inplace=True)

    # Check if IPNs in 'Removed From Prev File' are in contract data
    missing_data['On Corporate Contract'] = missing_data['IPN'].apply(
        lambda x: 'Yes' if any(x.lstrip('0') == ipn.lstrip('0') for ipn in contract_data['IPN']) else 'No')

    with pd.ExcelWriter(missing_ipn_file, engine='openpyxl', mode='a') as writer:
        # Get all sheet names in the workbook
        sheet_names = writer.book.sheetnames

        # Find 'detail' or 'DETAIL' (case-insensitive) and update
        detail_sheet_name = next((sheet for sheet in sheet_names if sheet.lower() == 'detail'), None)
        if detail_sheet_name:
            writer.book.remove(writer.book[detail_sheet_name])
        merged_data.to_excel(writer, sheet_name='detail', index=False)

        # Find 'Removed From Prev File' sheet and update
        removed_sheet_name = next((sheet for sheet in sheet_names if sheet.lower() == 'removed from prev file'), None)
        if removed_sheet_name:
            writer.book.remove(writer.book[removed_sheet_name])
        missing_data.to_excel(writer, sheet_name='Removed From Prev File', index=False)

        # Rearrange the sheets
        writer.book._sheets = [writer.book['detail'], writer.book['Removed From Prev File']] + [sheet for sheet in
                                                                                                writer.book if
                                                                                                sheet.title not in [
                                                                                                    'detail',
                                                                                                    'Removed From Prev File']]

    # Show success message
    messagebox.showinfo("Success",
                        "Comparison with contract file complete! The 'Corporate Contract Price' column has been added "
                        "to the 'detail' sheet and the 'On Corporate Contract' column to the 'Removed From Prev File' "
                        "sheet. You can now see if these IPNs are in our Active Contract File and their corresponding "
                        "prices.")
