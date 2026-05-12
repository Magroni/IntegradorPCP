# AGENT_CONTEXT.md — Contexto do Projeto PCP Costa Granitos
> **Atualizado em:** 2026-05-12  
> **Propósito:** Arquivo de contexto para agentes de IA. Leia este arquivo antes de qualquer alteração no projeto.

---

## 1. Visão Geral do Sistema

Sistema web de **PCP (Planejamento e Controle da Produção)** para a **Costa Granitos**, uma empresa de beneficiamento de pedras naturais (granito, mármore, quartzito).

- **Stack:** Python + Streamlit (UI web local)
- **Servidor:** Roda localmente via `run_app.bat`
- **Persistência:** Arquivos Excel (`.xlsm` e `.xlsx`) — sem banco de dados relacional
- **Caminhos configuráveis** via `config.json` (não hardcoded)

---

## 2. Estrutura de Arquivos

```
ProgramarProd/
├── app.py                  # Aplicação principal Streamlit (6 abas)
├── data_manager.py         # Toda lógica de leitura/escrita nas planilhas Excel
├── export_manager.py       # (LEGADO) Gerador de Excel para exportação — não mais usado pela app
├── config.json             # Caminhos das bases de dados (editável pelo usuário na UI)
├── requirements.txt        # Dependências Python
├── run_app.bat             # Script para iniciar o app
├── AGENT_CONTEXT.md        # Este arquivo
└── analise_apontamento.py  # Script temporário de análise (pode ser ignorado/deletado)
```

---

## 3. Bases de Dados (Planilhas Excel)

### 3.1 DB.xlsm (Programação)
Caminho configurável via `config.json` → chave `"DB_FILE"`.
*(Anteriormente chamado: `COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm`)*

| Aba | Conteúdo |
|-----|----------|
| `PROGRAMAÇÃO` | Tabela principal: todos os blocos e etapas de produção |
| `ENTREGUES` | Histórico de blocos finalizados e entregues (usada para calcular médias de capacidade) |
| `BASE DE DADOS` | Mapeamento Processo → Setor (Máquina) padrão |

#### Colunas da aba PROGRAMAÇÃO (principais):
| Coluna | Descrição |
|--------|-----------|
| `BLOCO` | Número identificador do bloco de pedra |
| `MATERIAL` | Nome do material (ex: ALPINUS, BLACK VULKON) |
| `DEMANDA` | Tipo de demanda/cliente |
| `PROCESSO` | Nome do processo de produção (ex: SERRAR, RESINAR, LEVIGAR) |
| `SETOR` | Máquina/setor responsável pelo processo |
| `DATA` | Data **programada** para o processo |
| `DATA REALIZADA` | Data em que o processo foi **efetivamente realizado** |
| `STATUS PROCESSO` | `NÃO REALIZADO` / `EM PROCESSO` / `REALIZADO` |
| `QTD. CHAPAS` | Quantidade de chapas do bloco |
| `VOLUME M²` | Volume em metros quadrados |
| `OBSERVAÇÃO DE PRODUÇÃO` | Campo livre para anotações |

### 3.2 DB.xlsx (Apontamento de Produção)
Caminho configurável via `config.json` → chave `"APONTAMENTO_FILE"`.
*(Anteriormente chamado: `Apontamento Produção (REV 1).xlsx`)*

Planilha mantida pelo operador. O sistema a lê em **modo somente leitura**.

| Aba | Uso |
|-----|-----|
| `BD` | Registros de produção diária (cabeçalho na linha 7) |
| `BASE DADOS` | Mapeamento PROCESSO_COMPLETO → RESUMIDO (ex: "19-POLIMENTO (S)" → "POLIMENTO") |

#### Colunas relevantes da aba BD:
- `DATA REG` — data do apontamento
- `MATERIAL+BLOCO` — texto concatenado (ex: "ALPINUS-1234")
- `Nº BLOCO` — número do bloco (às vezes vazio — extrair de MATERIAL+BLOCO)
- `PROCESSO` — nome completo do processo conforme o operador registrou
- `SETOR` — máquina usada
- `QTD CH (SEM RET & REPASSE)` — quantidade de chapas produzidas

---

## 4. Arquitetura da Aplicação (app.py)

### Abas (tabs):
| # | Tab | Propósito |
|---|-----|-----------|
| 1 | 🛠️ Adicionar / Editar Bloco | CRUD de blocos. Busca por número, edita todas as etapas do roteiro. |
| 2 | 👁️ Base de Dados | Visualização completa da tabela de programação. Filtros e busca. |
| 3 | 🗓️ Janela de Programações | Painel duplo: Fila de Trabalho (esquerda) + Painel de Alocação (direita). Agendamento em lote. |
| 4 | ✅ Apontamento | Cruzamento Programação × Apontamento. Indicador de aderência. Confirmação de REALIZADOS. |
| 5 | 🖨️ Exportação | Relatório HTML gerado em tela, pronto para Ctrl+P. |
| 6 | ⚙️ Opções Gerais | Configuração de caminhos dos arquivos + mapeamento Processo×Máquina. |

---

## 5. Módulo data_manager.py

### Configuração
```python
get_config()       # Lê config.json → dicionário com DB_FILE e APONTAMENTO_FILE
save_config(cfg)   # Salva config.json
_get_db_file()     # Retorna o caminho atual do DB de programação (leitura dinâmica)
_get_apontamento_file()  # Retorna o caminho atual do apontamento
```

### Leitura de Dados
```python
get_data()                         # Lê aba PROGRAMAÇÃO → DataFrame principal
get_base_dados()                   # Lê aba BASE DE DADOS → mapeamento Processo×Setor
get_historico_medias_entregues()   # Lê aba ENTREGUES → médias históricas de chapas/dia por máquina
get_apontamentos_do_dia(date)      # Lê aba BD do Apontamento → DataFrame do dia filtrado
get_mapa_resumido_processos()      # Lê BASE DADOS do Apontamento → dict {processo_completo: resumido}
```

### Escrita de Dados
```python
add_record(record_dict)            # Adiciona nova linha na aba PROGRAMAÇÃO
update_cell_by_row(idx, updates)   # Atualiza colunas específicas de uma linha pelo índice pandas
save_bloco_edits(bloco_id, rows_data) # Salva edições completas de um bloco
update_base_dados(df)              # Regrava a aba BASE DE DADOS com o df editado
```

### Validações de Negócio
```python
validar_sequencia_bloco(df, bloco, idx, nova_data)
# Garante que o processo anterior do bloco já está agendado ou realizado
# antes de permitir o agendamento do processo atual.
# Retorna (True, "") ou (False, "mensagem de erro")
```

---

## 6. Regras de Negócio Importantes

1. **Sequência obrigatória:** Não se pode agendar o processo N sem que o processo N-1 tenha pelo menos uma data programada. (`validar_sequencia_bloco`)

2. **Proteção de histórico:** Processos com `STATUS PROCESSO == "REALIZADO"` não podem ter suas datas editadas pela interface normal.

3. **Cálculo de Aderência:** `aderencia = (já_confirmados + encontrados_no_apontamento) / total_programados * 100`

4. **Bloqueio Visual na Fila:** A coluna "Liberado?" na Aba 3 é **temporal** — ela compara a data do processo anterior com a **Data Alvo** selecionada no topo da tela. Um processo bloqueado pode aparecer como 🟢 Sim se a Data Alvo for futura o suficiente.

5. **Extração de BLOCO do Apontamento:** O campo `Nº BLOCO` frequentemente está vazio. Nesses casos, o sistema extrai o número do campo `MATERIAL+BLOCO` (ex: `"ALPINUS-1234"` → `"1234"`).

6. **Normalização de nomes de processo:** Para cruzar programação × apontamento, o sistema usa a tabela `BASE DADOS` do arquivo de Apontamento como dicionário de equivalência (processo_completo → resumido).

---

## 7. Configuração de Caminhos

O arquivo `config.json` na raiz do projeto controla os caminhos:

```json
{
    "DB_FILE": "z:\\PCP\\PROJETOS MARLON\\ProgramarProd\\DB.xlsm",
    "APONTAMENTO_FILE": "z:\\PCP\\PROJETOS MARLON\\ProgramarProd\\DB.xlsx"
}
```

- **Editável pela UI** na Aba ⚙️ Opções Gerais → Seção "Caminhos das Bases de Dados"
- O sistema **valida** se o arquivo existe antes de salvar
- Alterações entram em vigor após **F5** na página

---

## 8. Histórico de Mudanças

| Data | Mudança |
|------|--------|
| 2026-05-12 | Arquivos renomeados para `DB.xlsm` e `DB.xlsx`; config.json e data_manager.py atualizados |
| 2026-05-12 | data_manager.py reescrito do zero para corrigir encoding corrompido pelo PowerShell |
| 2026-05-12 | Criação do arquivo AGENT_CONTEXT.md |
| 2026-05-12 | Aba 6 (Opções Gerais): seção de configuração de caminhos adicionada |
| 2026-05-12 | config.json criado; data_manager.py refatorado para leitura dinâmica dos caminhos |
| 2026-05-12 | Aba 4 (Apontamento): criada como aba independente; cruzamento automático com planilha de Apontamento |
| 2026-05-12 | Aba 4: Indicador de aderência à programação (5 métricas) |
| 2026-05-12 | Aba 4: Coluna "Máquina" preenchida via SETOR_AP do Apontamento quando SETOR da prog. está vazio |
| 2026-05-12 | Aba 4: Filtro de máquinas expandido com setores lidos do Apontamento |
| 2026-05-12 | Aba 4: Processos já REALIZADOS aparecem na tabela com status 🟢 Já Confirmado |
| 2026-05-12 | Aba 3 (Programações): Fila mostra 🟢/🔴 baseado na Data Alvo (temporal) |
| 2026-05-12 | Aba 5 (Exportação): Substituído download Excel por relatório HTML com @media print |
| 2026-05-12 | data_manager.py: Adicionado get_apontamentos_do_dia() e get_mapa_resumido_processos() |

---

## 9. Como Iniciar o Sistema

```bat
run_app.bat
```
Ou manualmente:
```powershell
streamlit run app.py
```

---

## 10. Pendências / Backlog

- [ ] Integração automática bidirecional com o Apontamento (ler e confirmar sem intervenção manual)
- [ ] Tabela de equivalência de nomes de processo editável pela UI (para melhorar o cruzamento)
- [ ] Migração para servidor de rede / SharePoint (alterar `DB_FILE` no config.json)
- [ ] Log de auditoria: registrar quem alterou qual data e quando
- [ ] Melhorar performance do `data_editor` na fila quando houver muitos blocos pendentes
