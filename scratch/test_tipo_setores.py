import sys
sys.path.append("z:/PCP/PROJETOS MARLON/ProgramarProd")
import data_manager as dm

def run():
    print("Testando carregamento/inicializacao da aba TIPO_SETORES...")
    df = dm.get_tipo_setores()
    print("DataFrame retornado:")
    print(df)
    
    print("\nTestando atualizacao da aba...")
    # Vamos adicionar uma nova linha de teste
    df_novo = df.copy()
    # Adiciona linha se nao tiver "Teste"
    if "Teste" not in df_novo["TIPO_PROCESSO"].values:
        df_novo = df_novo._append({"TIPO_PROCESSO": "Teste", "SETORES": "TESTE1, TESTE2"}, ignore_index=True)
        
    sucesso = dm.update_tipo_setores(df_novo)
    print("Atualizacao bem sucedida?", sucesso)
    
    if sucesso:
        df_verificado = dm.get_tipo_setores()
        print("\nVerificado apos salvar:")
        print(df_verificado)
        
        # Limpa o teste para deixar a planilha original limpa
        df_clean = df_verificado[df_verificado["TIPO_PROCESSO"] != "Teste"]
        dm.update_tipo_setores(df_clean)
        print("\nLimpeza de teste concluida!")

if __name__ == "__main__":
    run()
