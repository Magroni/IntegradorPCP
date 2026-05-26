import sys
sys.path.append("z:/PCP/PROJETOS MARLON/ProgramarProd")
import pandas as pd
import data_manager as dm

def check():
    df_base = dm.get_base_dados()
    print("Mapeamentos cadastrados na Base de Dados (Processo -> Setor):")
    for idx, row in df_base.iterrows():
        print(f"  {row.get('PROCESSO')} -> {row.get('SETOR')}")

if __name__ == "__main__":
    check()
