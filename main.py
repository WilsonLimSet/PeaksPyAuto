import pandas as pd
import random
import os

def select_people(file_list, output_file):
    """
    Select 10 random people from each input Excel file based on a unique email address,
    and create a new Excel file with two sheets: one with all the selected people and one
    with the rest of the data. Set the column widths of the output Excel file.

    :param file_list: List of input file names
    :param output_file: Name of output file
    """
    with pd.ExcelWriter(output_file) as writer:
        # Keep track of selected people across all sheets
        selected_people_across_sheets = set()

        # Create a DataFrame to store all the selected people
        df_all_selected = pd.DataFrame()

        # Add a sheet for all selected people
        df_all_selected.to_excel(writer, index=False, header=True, startrow=0, sheet_name='All Selected')

        # Set column widths of output file for the 'All Selected' sheet
        worksheet = writer.sheets['All Selected']
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 20
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 15
        worksheet.column_dimensions['E'].width = 25
        worksheet.column_dimensions['F'].width = 15
        worksheet.column_dimensions['G'].width = 20
        worksheet.column_dimensions['H'].width = 20
        worksheet.column_dimensions['I'].width = 20
        worksheet.column_dimensions['J'].width = 20
        worksheet.column_dimensions['K'].width = 25

        for file in file_list:
            # Read input file
            try:
                df = pd.read_excel(os.path.join(os.getcwd(), file))
            except FileNotFoundError:
                # Skip over file if it does not exist
                print(f"File not found: {file}")
                continue
            except:
                # Skip over file if there is an error reading it
                print(f"Error reading file: {file}")
                continue

            # Remove already selected people from consideration
            remaining_people = df[~df['USC Email'].isin(selected_people_across_sheets)]

            # Select 10 random people based on unique email address
            if len(remaining_people) >= 10:
                selected_people_sheet = remaining_people.sample(n=10)
            else:
                selected_people_sheet = remaining_people

            # Add selected people from current sheet to overall list
            selected_people_across_sheets.update(selected_people_sheet['USC Email'])

            # Append selected people to the DataFrame of all selected people
            df_all_selected = df_all_selected.append(selected_people_sheet)

            # Split data into selected and rest
            df_selected = selected_people_sheet
            df_rest = df[~df['USC Email'].isin(selected_people_across_sheets)]

            # Write data to output file
            sheet_name = os.path.splitext(file)[0][:31] # Truncate sheet name to 31 characters or less
            df_selected.to_excel(writer, index=False, header=True, startrow=0, sheet_name=sheet_name)
            df_rest.to_excel(writer, index=False, header=True, startrow=0, sheet_name= "waitlist_"+ sheet_name)

            # Set column widths of output file
            worksheet = writer.sheets[sheet_name]
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 15
            worksheet.column_dimensions['E'].width = 25
            worksheet.column_dimensions['F'].width = 15
            worksheet.column_dimensions['G'].width = 20
            worksheet.column_dimensions['H'].width = 20
            worksheet.column_dimensions['I'].width = 20
            worksheet.column_dimensions['J'].width = 20
            worksheet.column_dimensions['K'].width = 25

            # Set column widths of output file for the 'Waitlist' sheets
        for sheet_name in writer.sheets:
            if sheet_name.startswith("waitlist_"):
                worksheet = writer.sheets[sheet_name]
                worksheet.column_dimensions['A'].width = 20
                worksheet.column_dimensions['B'].width = 20
                worksheet.column_dimensions['C'].width = 15
                worksheet.column_dimensions['D'].width = 15
                worksheet.column_dimensions['E'].width = 25
                worksheet.column_dimensions['F'].width = 15
                worksheet.column_dimensions['G'].width = 20
                worksheet.column_dimensions['H'].width = 20
                worksheet.column_dimensions['I'].width = 20
                worksheet.column_dimensions['J'].width = 20
                worksheet.column_dimensions['K'].width = 25

            # Save the Excel file
        writer.save()

    # Get list of input files
    path = os.getcwd()
    file_list = [f for f in os.listdir(path) if f.endswith('.xlsx')]

    # Set name of output file
    output_file = 'selected_people.xlsx'

    # Run script
    select_people(file_list, output_file)