import openpyxl
import data_manager as dm

file_path = dm._get_apontamento_file()
wb = openpyxl.load_workbook(file_path, read_only=True)
print("Sheets:", wb.sheetnames)
