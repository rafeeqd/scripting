import pandas as pd
import os
import glob
import xlsxwriter

def convert_json_to_excel(path_to_json, output_file):
    """
    Converts JSON files in a directory to an Excel file.

    Args:
        path_to_json (str): The path to the directory containing JSON files.
        output_file (str): The name of the output Excel file.

    Returns:
        None
    """
    # Find all JSON files in the directory
    json_pattern = os.path.join(path_to_json, '*.json')
    file_list = glob.glob(json_pattern)

    # Create an Excel workbook and worksheet
    xbook = xlsxwriter.Workbook(output_file)
    xsheet = xbook.add_worksheet('Test')

    count = 0  # Counter variable

    # Process each JSON file
    for files in file_list:
        # Read the JSON file
        json_read = pd.read_json(files)
        get_attributes = json_read.attributes.values.tolist()

        # Write attribute names and values to the worksheet
        for index, row in enumerate(get_attributes):
            xsheet.write(count, index, row['trait_type'])
            if row['value'] == 'Blank':
                xsheet.write(count + 1, index, '')
            else:
                xsheet.write(count + 1, index, row['value'])

        count += 2  # Increment the counter

    xbook.close()  # Close the workbook

# Usage example
path_to_json = r'Yourpath'
output_file = 'Test1.xlsx'
convert_json_to_excel(path_to_json, output_file)
