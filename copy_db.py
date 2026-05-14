import shutil
import os

source_file = r"z:\PCP\PROJETOS MARLON\CalculadorWIP\COSTA GRAN. - PROGRAMAÇÕES - WIP NOVO.xlsm"
dest_file = r"z:\PCP\PROJETOS MARLON\ProgramarProd\Base_de_Dados_Copia.xlsm"

try:
    shutil.copy2(source_file, dest_file)
    print(f"File successfully copied to {dest_file}")
except Exception as e:
    print(f"Error copying file: {e}")
