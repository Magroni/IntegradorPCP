# Gerenciador de Programação WIP — Costa Granitos

Sistema web de **PCP (Planejamento e Controle da Produção)** para gestão de blocos de pedra natural. Desenvolvido em Python + Streamlit.

## Funcionalidades

- **📊 Cadastro de Blocos** — Adicionar e editar blocos com roteiro completo de processos
- **👁️ Base de Dados** — Visualização e filtro da programação completa
- **🗓️ Janela de Programações** — Fila de trabalho com alocação de máquinas em lote
- **✅ Apontamento** — Cruzamento automático com planilha de apontamento + indicador de aderência
- **🖨️ Exportação** — Relatório HTML pronto para impressão (Ctrl+P)
- **⚙️ Opções Gerais** — Configuração dinâmica de arquivos e abas do Excel

## Instalação

```powershell
# Instalar dependências
pip install -r requirements.txt

# Iniciar aplicação
run_app.bat
# ou
streamlit run app.py
```

## Configuração dos Arquivos

O sistema lê os caminhos dos arquivos Excel a partir do `config.json` (não versionado).  
Copie o arquivo de exemplo e ajuste os caminhos para a sua máquina:

```powershell
copy config.json.example config.json
```

Depois ajuste os caminhos em **⚙️ Opções Gerais** dentro da própria aplicação.

## Estrutura de Arquivos

| Arquivo | Descrição |
|---|---|
| `app.py` | Aplicação principal Streamlit |
| `data_manager.py` | Leitura e escrita nas planilhas Excel |
| `export_manager.py` | Gerador de relatórios (legado) |
| `config.json` | Caminhos e nomes de abas (**não versionado**) |
| `AGENT_CONTEXT.md` | Contexto detalhado para agentes de IA |

## Dependências

Ver `requirements.txt` — principais: `streamlit`, `pandas`, `openpyxl`.
