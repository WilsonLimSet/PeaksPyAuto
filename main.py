import pandas as pd
import random
import os
import openpyxl

def select_people(file_list, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for file in file_list:
            selected_people = []
            data = []
            df = pd.read_excel(file, engine='openpyxl')
            people = df[df['USC Email'].duplicated(keep=False) == False]['USC Email'].tolist()
            if len(people) >= 10:
                selected_people = random.sample(people, 10)
            else:
                selected_people = people
            data = df[df['USC Email'].isin(selected_people)]
            df_selected = pd.DataFrame(data)
            df_rest = df[~df['USC Email'].isin(selected_people)]
            sheet_name = file[:31] # Truncate sheet name to 31 characters or less
            df_selected.to_excel(writer, index=False, header=True, startrow=0, sheet_name=sheet_name)
            df_rest.to_excel(writer, index=False, header=True, startrow=0, sheet_name=sheet_name + "_rest")

            # Get the active worksheet
            worksheet = writer.sheets[sheet_name]
            # Set column widths
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
            # Add more columns as needed

        writer.save()

path = os.getcwd()
files = os.listdir(path)
file_list = [f for f in files if f[-3:] == 'lsx'] # List of file names
output_file = 'selected_people.xlsx'
select_people(file_list, output_file)