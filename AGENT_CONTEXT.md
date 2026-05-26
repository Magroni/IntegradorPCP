# AGENT_CONTEXT.md — Contexto do Projeto PCP Costa Granitos
> **Atualizado em:** 2026-05-26 (v19)  
> **Propósito:** Arquivo de contexto para agentes de IA. Leia este arquivo antes de qualquer alteração no projeto.

---

## 1. Visão Geral do Sistema

Sistema web de **PCP (Planejamento e Controle da Produção)** para a **Costa Granitos**, uma empresa de beneficiamento de pedras naturais (granito, mármore, quartzito).

- **Stack:** Python + Streamlit (UI web local)
- **Servidor:** Roda localmente via `run_app.bat`
- **Persistência:** Arquivos Excel (`.xlsm`, `.xlsx`, `.xlsb`) — sem banco de dados relacional
- **Caminhos configuráveis** via `config.json` (não hardcoded)
- **Versão:** Controle de versão via Git
- **Repositório Remoto:** [https://github.com/Magroni/IntegradorPCP.git](https://github.com/Magroni/IntegradorPCP.git)

---

## 2. Estrutura de Arquivos

```
ProgramarProd/
├── app.py                  # Aplicação principal Streamlit (Arquitetura Híbrida: Form + Live)
├── data_manager.py         # Lógica de dados, mapeamentos e suporte a Insumos/Paradas
├── config.json             # Configurações de caminhos e abas (Novo: SHEET_AP_INSUMOS)
├── run_app.bat             # Atalho para execução
├── AGENT_CONTEXT.md        # Guia de contexto do projeto (Manter atualizado!)
└── README.md               # Instruções gerais
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
| `BASE DADOS` | Mapeamento PROCESSO_COMPLETO → RESUMIDO |
| `PARADAS` | Ocorrências de máquinas (ID, MOTIVO, DIA_INICIO, HORA_INICIO, DIA_FIM, HORA_FIM, TEMPO) |
| `INSUMOS` | Detalhamento (ID, TIPO, DESCRICAO, QTD, UNID, TEMPO_SECAGEM, CABECAS, INSUMO_DETALHE) |

#### Colunas reais da aba DB (cabeçalho na linha 0, sem sujeira):

| Coluna | Descrição |
|---|---|
| `ID` | Identificador auto-incremento |
| `DATA_REG` | Data em que o operador digitou o registro (NÃO usar para confronto) |
| `MATERIAL` | Nome do material (ex: MAGMA, BLANC DU BLANC) |
| `NUMERO_BLOCO` | Número do bloco (ex: 3706, 288, 7179) |
| `PROCESSO` | Nome do processo (ex: RESINAR BRUTO, POLIR REFEITO) |
| `SETOR` | Máquina/setor (ex: RESINA, CIMEF) |
| `ESP` | Espessura |
| `QTD_CHAPAS` | Quantidade de chapas produzidas |
| `COMP`, `ALT`, `QTDM2` | Comprimento, Altura, Metros quadrados |
| `OPERADOR` | Nome do operador |
| **`DATA_INICIO`** | **Data real do trabalho na fábrica (USAR ESTA para confronto!)** |
| `DATA_FIM` | Data de término do trabalho |
| `HORA_INICIO`, `HORA_FIM` | Horários |
| `TEMPO_PROCESSO` | Duração |
| `TURNO` | D (Diurno) ou N (Noturno) |
| `TIPO_RES`, `QTDKG_RES` | Tipo e Quantidade de resina |
| `TIPO_ENDUR`, `QDKG_ENDUR` | Tipo e Quantidade de endurecedor |
| `TEMPO_SECAGEM` | Tempo de secagem |
| `SAT1..SAT20` | Sequência de abrasivos |

> **IMPORTANTE**: Para o Confronto (Aba 4), a coluna correta é `DATA_INICIO`, NÃO `DATA_REG`.

### 3.3 PLANILHA BLOCOS.xlsb (Cadastro de Blocos)
Caminho configurável via `config.json` → chave `"BLOCKS_FILE"`.

- **Aba:** `PLAN. BLOCOS`
- **Header:** Linha 9 (skiprows=8 no pandas)
- **Colunas relevantes** (nomes reais com encoding variável):
  - `Nº BLOCO` → Identificador do bloco (⚠️ encoding pode variar: `N° BLOCO`, `Nº BLOCO`)
  - `MATERIAL` → Nome do material
  - `COMP. (LIQUIDO)` → Comprimento líquido (m)
  - `ALT. (LIQUIDO)` → Altura líquida (m)
  - `LARG. (LIQUIDO)` → Largura líquida (m)

> **CRÍTICO:** Os nomes de colunas neste arquivo .xlsb contêm caracteres especiais e encoding instável. A busca de colunas DEVE ser **dinâmica por keywords** (ex: buscar "BLOCO", "COMP"+"LIQUIDO") e **NUNCA hardcoded** como `N_BLOCO` ou `COMP_LIQUIDO`.

### 3.4 Estoque Chapas.xlsx (Cadastro Secundário)
Caminho configurável via `config.json` → chave `"CHAPAS_FILE"`.

- **Aba:** `ENTRADAS`
- Usado como **fallback** quando o bloco não é encontrado na PLANILHA BLOCOS
- Busca flexível por colunas: `BLOCO`, `Nº BLOCO`, `N_BLOCO`, etc.

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
| 6 | 📈 Análises e Indicadores | Dashboard de performance (M², Chapas, Turnos, Máquinas e Refeito). |
| 7 | ⚙️ Opções Gerais | Configuração de caminhos dos arquivos + mapeamento Processo×Máquina. |

### Padrão de gráficos empilhados (Aba 6)
Os gráficos de barras empilhadas (Normal/Refeito) usam **cálculo explícito de ponto médio** para centralizar labels:
1. Calcula `_y0` (início) e `_y1` (fim) de cada segmento via acumulação no pandas
2. `_mid = (_y0 + _y1) / 2` → posição Y do texto
3. `_label` → string formatada (vazio para valor=0, evita texto em barras vazias)
4. **Todas as camadas** (bars, text_seg, text_totals) usam o **mesmo DataFrame** — obrigatório para `facet`
5. Ordenação cronológica via **lista explícita** `sort=ordem_dias` (não `EncodingSortField`)

---

## 5. Módulo data_manager.py

### Configuração
```python
get_config()       # Lê config.json → dicionário com DB_FILE, APONTAMENTO_FILE, BLOCKS_FILE, CHAPAS_FILE
save_config(cfg)   # Salva config.json
_get_db_file()     # Retorna o caminho atual do DB de programação (leitura dinâmica)
_get_apontamento_file()  # Retorna o caminho atual do apontamento
```

### Leitura de Dados
```python
get_data()                         # Lê aba PROGRAMAÇÃO → DataFrame principal (normaliza BLOCO)
get_base_dados()                   # Lê aba BASE DE DADOS → mapeamento Processo×Setor
get_historico_medias_entregues()   # Lê aba ENTREGUES → médias históricas de chapas/dia por máquina
get_apontamentos_do_dia(date)      # Lê aba BD do Apontamento → DataFrame do dia filtrado
get_all_apontamentos()             # Lê todos os apontamentos (para Aba 6 - Indicadores)
get_mapa_resumido_processos()      # Lê BASE DADOS do Apontamento → dict {processo_completo: resumido}
get_bloco_info(bloco_id)           # Busca material e medidas na PLANILHA BLOCOS (1ª) ou ESTOQUE CHAPAS (fallback)
get_apontamentos_por_bloco(bloco_id)  # Histórico de apontamentos de um bloco específico
```

> **CRÍTICO (get_bloco_info):** Usa busca **dinâmica de colunas** (`find_col` com keywords) para lidar com encoding variável do `.xlsb`. Nunca referenciar colunas como `N_BLOCO` diretamente.

### Escrita de Dados
```python
add_record(record_dict)            # Adiciona nova linha na aba PROGRAMAÇÃO
update_cell_by_row(idx, updates)   # Atualiza colunas específicas de uma linha pelo índice pandas
salvar_edicao_bloco_excel(...)     # Salva edições completas de um bloco (insere/remove linhas)
update_base_dados(df)              # Regrava a aba BASE DE DADOS com o df editado
add_apontamento_batch(batch)       # Grava múltiplos apontamentos de uma vez (carrinho)
add_paradas(paradas_list)          # Grava paradas na aba PARADAS
add_insumos(insumos_list)          # Grava insumos na aba INSUMOS
```

### Validações de Negócio
```python
validar_sequencia_bloco(df, bloco, idx, nova_data)
# Garante que o processo anterior do bloco já está agendado ou realizado
# antes de permitir o agendamento do processo atual.
# Retorna (True, "") ou (False, "mensagem de erro")
```

### Normalização
```python
normalize_bloco(bloco)  # Padroniza ID: remove .0, espaços, força upper. Usado em get_data() e confronto.
```

---

## 6. Regras de Negócio Importantes

1. **Sequência obrigatória:** Não se pode agendar o processo N sem que o processo N-1 tenha pelo menos uma data programada. (`validar_sequencia_bloco`)

2. **Proteção de histórico:** Processos com `STATUS PROCESSO == "REALIZADO"` não podem ter suas datas editadas pela interface normal.

3. **Cálculo de Aderência:** `aderencia = (já_confirmados + encontrados_no_apontamento) / total_programados * 100`

4. **Bloqueio Visual na Fila:** A coluna "Liberado?" na Aba 3 é **temporal** — ela compara a data do processo anterior com a **Data Alvo** selecionada no topo da tela. Um processo bloqueado pode aparecer como 🟢 Sim se a Data Alvo for futura o suficiente.

5. **Self-Healing de Máquinas (SETOR)**: Se a coluna `SETOR` estiver vazia no Excel (ex: fórmula não calculada), o `data_manager.py` autocompleta o valor cruzando o `PROCESSO` com a aba `BASE DE DADOS`. Isso garante que os filtros de máquina sempre funcionem.

6. **Gestão de Programação (Remoção)**: É possível remover blocos de um dia específico apenas desmarcando-os na Janela de Programações e clicando em "Confirmar Alterações". O sistema limpa a data no Excel e volta o status para "NÃO REALIZADO".

7. **% Refeito dinâmico (Aba 6)**: O cálculo de `% Refeito` usa a **métrica selecionada** (Chapas ou M²). Se "Chapas" está selecionado → refeito_ch / total_ch; se "M²" → refeito_m2 / total_m2. O label também muda: `% Refeito (Chapas)` ou `% Refeito (M²)`.

8. **Classificação Refeito/Normal**: Processos cujo nome contém `REPASSE`, `REFEITO` ou `REPROCESSO` são classificados como "Refeito". O setor produtivo (máquina/setor) **RETOQUE** e o seu processo associado **RETOCAR** são **incluídos** normalmente nas análises por solicitação do usuário, permitindo o correto apontamento de materiais de processo novo.

9. **Data Fim e Dia de Produção (Aba 6)**: Para as análises e indicadores, a data e hora de finalização (`DIA_FIM` / `HORA_FIM`) são usadas como parâmetro de data. Se vazias, realiza fallback para `DIA_INICIO` / `HORA_INICIO`. Como o dia de produção começa às 07:00 e termina às 06:59 do dia subsequente (atendendo aos turnos de produção), qualquer processo finalizado antes de 07:00 AM é contabilizado na data de produção do dia anterior.

10. **Blocos com Duplo Código (Equivalência)**: Suporte para blocos identificados por códigos compostos com barra `/` (ex: `4244/771418`). A comparação/busca flexível (`blocos_match`) é baseada na interseção das partes dos códigos. Assim, buscar por `4244` ou `771418` casará automaticamente com o bloco `4244/771418` em todo o sistema (Planilha de Blocos, Estoque de Chapas, Formulários de Edição, Fila de Cruzamento PCP/Apontamento e Consultas de Histórico).

11. **Filtro de Setores e Processos por Tipo de Produção (Aba 4)**: No formulário de apontamento, as listas de seleção dos campos "Máquina/Setor" e "2. Processo/Etapa" são totalmente dinâmicas. Ao selecionar o "Tipo de Processo", o campo "Máquina/Setor" filtra os setores permitidos e o campo "Processo/Etapa" filtra dinamicamente as etapas cujas máquinas pertencem àquele tipo de produção (resolvido por cruzamento de chaves entre a aba `TIPO_SETORES` e a aba `BASE DE DADOS`). Essas vinculações são governadas por tabelas editáveis na interface (Aba 7).

12. **Padronização de Paradas e Farol Lean (Aba 6 & Aba 4)**: A lista oficial de motivos de parada e suas classificações (Farol Lean: `Operacional`, `Intervenção`, `Crítica`) é cadastrada na aba `TIPO_PARADAS` de `DB.xlsm` via Aba 7 (⚙️ Opções Gerais). Essa lista gera dinamicamente uma caixa de seleção suspensa (dropdown `SelectboxColumn`) no formulário de apontamento de paradas (Aba 4) para garantir a padronização na entrada dos dados. No painel de análises, o Pareto e os cartões Lean utilizam esse cadastro oficial, mantendo fallback de mapeamento (`BASE PARADAS` em `DB.xlsx`) e por palavra-chave para dados históricos.

---

## 7. Configuração de Caminhos

O arquivo `config.json` na raiz do projeto controla os caminhos:

```json
{
    "DB_FILE": "z:\\PCP\\PROJETOS MARLON\\ProgramarProd\\COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm",
    "APONTAMENTO_FILE": "z:\\PCP\\PROJETOS MARLON\\ProgramarProd\\Apontamento Produção (REV 2).xlsx",
    "BLOCKS_FILE": "z:\\PCP\\PROJETOS MARLON\\ProgramarProd\\PLANILHA BLOCOS.xlsb",
    "CHAPAS_FILE": "z:\\PCP\\PROJETOS MARLON\\ProgramarProd\\Estoque Chapas 2026.xlsx",
    "SHEET_PROGRAMACAO": "DB",
    "SHEET_ENTREGUES": "ENTREGUES",
    "SHEET_BASE_DADOS": "BASE DE DADOS",
    "SHEET_AP_BD": "DB",
    "SHEET_AP_BASE": "BASE DADOS",
    "SHEET_AP_PARADAS": "PARADAS",
    "SHEET_AP_INSUMOS": "INSUMOS"
}
```

- **Editável pela UI** na Aba ⚙️ Opções Gerais → Seção "Caminhos das Bases de Dados"
- O sistema **valida** se o arquivo existe antes de salvar
- Alterações entram em vigor após **F5** na página

---

## 8. Histórico de Mudanças

| Data | Alteração |
|------|-----------| 
| 2026-05-26 | **Modo Alto Contraste Unificado e Eliminação de Textos Cinza (v19)**: Implementado e unificado o Modo Alto Contraste nas duas principais interfaces de impressão do sistema: **Aba 5 (Exportação para o Chão de Fábrica)** e **Aba 6 (Relatório Formal A3/A4)**. O novo layout B&W de alto contraste elimina absolutamente todas as cores e tons de cinza do texto, forçando-os a preto puro (#000000) e os fundos a branco puro (#ffffff), prevenindo textos desbotados em impressoras monocromáticas. A tabela de totais de exportação foi redefinida para usar a elegante linha dupla de contabilidade clássica, garantindo ótima legibilidade e economia de toner no processo de impressão ou exportação para PDF. |
| 2026-05-26 | **Ultra-compactação e Resiliência a Cabeçalhos/Rodapés (v18)**: Identificado que o overflow na exportação para PDF era causado pela opção **"Cabeçalhos e rodapés"** (Headers and Footers) ativada nas configurações da caixa de diálogo do Chrome (o que reduz a altura útil da folha em ~3cm). Para tornar o Relatório A3/A4 100% resiliente a essa configuração, compactei ainda mais a planilha: reduzi o padding das seções para `6px`, o gap entre colunas para `8px`, o tamanho das fontes das tabelas para `7.2px` e os paddings das células para `1.5px 3px` (cabeçalhos) e `1px 2px` (conteúdo), garantindo o enquadramento de **exatamente 1 página** mesmo com cabeçalhos/rodapés e margens padrões ativados no PDF. |
| 2026-05-26 | **Ajuste de Proporção para Target A4 e Escala do Browser (v17)**: Configurada a resolução e layout do Relatório A3 para focar nas dimensões nativas de **A4 Landscape** (`size: A4 landscape;` no CSS, e limites máximos de largura compactados de `1500px`/`1540px` para `1100px`/`1120px`), mantendo a mesma proporção ideal (1:1.4142). Dessa forma, o relatório se ajusta perfeitamente em A4, e ao ser impresso ou gerado em folhas A3, a funcionalidade nativa de escala automática do navegador ("Ajustar à página" ou "Fit to Page") se encarrega de preencher e expandir o relatório pela folha inteira sem deixar espaços em branco ou sobras na margem inferior. |
| 2026-05-26 | **Relatório A3 Compactado para Impressão de 1 Página (v16)**: Otimizados e compactados todos os espaçamentos, margens e tamanhos do Relatório A3 para garantir que ele caiba perfeitamente em **exatamente 1 página** no formato A3 Landscape (Horizontal) ao salvar como PDF ou imprimir. Reduzidas as margens de página na regra `@page` para `0.4cm 0.6cm`, reduzido o padding geral e das colunas, ajustada a altura global de progress bars para `3px` (e `2px` nos severities), diminuídos os tamanhos de fonte de tabelas para `8px` e de KPIs, e reduzidas as lacunas e paddings de células para valores super compactos. Corrigidos também duplicados residuais nos componentes do HTML final. |
| 2026-05-26 | **Cadastro Editável de Paradas & Dropdown de Apontamentos (v15)**: Criada uma nova tabela interativa na Aba 7 (⚙️ Opções Gerais) para permitir o cadastro e edição direta de motivos oficiais de parada e seus tipos (Farol Lean: Operacional, Intervenção, Crítica). Essa lista é salva na aba **`TIPO_PARADAS`** do arquivo `DB.xlsm` do PCP (auto-inicializada com padrões enxutos). Integrado esse cadastro na tabela de paradas do Apontamento de Produção (Aba 4), transformando o campo texto livre em uma **lista suspensa (dropdown)** dinâmica baseada em `st.column_config.SelectboxColumn`, garantindo que os novos registros apontados pelos operadores sigam a padronização exata. O cálculo de indicadores e o relatório A3 agora cruzam a prioridade desse cadastro oficial de paradas com o mapeamento histórico do Excel. |
| 2026-05-26 | **Relatório A3 — Padronização de Paradas & Farol Lean (v14)**: Implementada a padronização e saneamento de ocorrências de inatividade no A3 baseando-se no mesmo padrão do mapeamento de processos. O sistema agora lê a aba **`BASE PARADAS`** (configurável) do arquivo de apontamentos, associando cada motivo bruto digitado a um nome padronizado (RESUMIDO) e a uma classificação do **Farol Lean** (TIPO_PARADA: Operacional, Intervenção, Crítica). Implementado fallback robusto por palavras-chave e tempo. O Pareto do A3 agora agrupa as ocorrências automaticamente pelo motivo padronizado, e o painel de severidade reflete as categorias reais mapeadas, enquanto a tabela detalhada exibe o "Motivo Real" bruto para auditoria. |
| 2026-05-26 | **Relatório A3 — Análise Lean de Severidade e Simplificação (v13)**: Otimizado o relatório A3 removendo a antiga seção 10 (Plano de Ação Lean) e assinaturas para que a gestão de contramedidas seja trabalhada à parte. Para ocupar de forma inteligente o espaço vertical sobressalente da coluna da direita, foi adicionada a nova seção **9. Análise Lean: Severidade e Frequência das Paradas**, com um painel visual de três cartões de desempenho categorizando as perdas por severidade: *Operacionais/Ajustes Rápidos* (<15 min), *Intervenções/Setups* (15 a 60 min) e *Críticas/Manutenções* (>60 min), cada uma com contagem de ocorrências, duration total, barra de progresso e % do impacto. A tabela de Ocorrências Críticas foi renomeada para Seção 10 e estendida para ocupar toda a largura inferior de forma simétrica. |
| 2026-05-26 | **Definição de Setor Retoque e Processo Retocar (v12)**: Alinhado e documentado no contexto que **RETOQUE** é o setor produtivo (máquina/setor) e **RETOCAR** é o nome real do processo (processo/etapa). Ambos permanecem 100% incluídos nas análises operacionais (qualidade e produtividade) para permitir a divisão futura dos materiais de processo novo no apontamento. |
| 2026-05-26 | **Inclusão Completa de Retoque / Retocar (v11)**: Revertida a regra de descarte de processos de retoque por solicitação do usuário. Os processos de `"RETOQUE"` e `"RETOCAR"` foram totalmente reintroduzidos nos dataframes de análise (`df_an`), permitindo que apareçam na qualidade e produtividade. O usuário fará uma divisão/saneamento dos apontamentos no Excel de modo que apenas novos materiais de processo novo sejam apontados como retoque normal. |
| 2026-05-26 | **Exclusão de Processos de Retoque / Retocar (v10)**: Corrigido o vazamento de dados do processo "RETOCAR" nas análises de indicadores. A regra de negócio original determinava que processos de retoque deveriam ser excluídos, mas a checagem no Pandas usava apenas `.str.contains("RETOQUE")`, não filtrando o infinitivo "RETOCAR" usado nos apontamentos. Atualizado o filtro para `"RETOQUE|RETOCAR"`, removendo o setor/máquina RETOQUE dos relatórios de qualidade e produtividade e sanando a distorção (onde ele aparecia indevidamente com 18% de volume e 0% de refugo). |
| 2026-05-26 | **Relatório A3 — Distribuição por Processo Totalizando 100% (v9)**: Corrigido o bug onde a tabela de Distribuição por Processo Produtivo no relatório A3 não somava 100% devido à exibição apenas dos top 6 processos. Implementado o agrupamento Pareto-style sob a linha "OUTROS PROCESSOS (X proc.)" para agrupar dinamicamente todos os demais processos fora do top 6, garantindo que o somatório da participação sempre inteire 100%. |
| 2026-05-26 | **Relatório A3 Lean Operacional & Produtividade Integrada (Aba 6)**: Expandido o Relatório A3 para ser um painel de desempenho integrado de toda a operação da Costa Granitos. O novo layout A3 Landscape (420mm x 297mm) consolida de forma integrada a produtividade física total de Chapas e M² processados (com taxas de refugo/reprocesso normal vs refeito), uma matriz detalhada de produção por máquina e turno (Diurno D vs Noturno N), além dos KPIs de inatividade/paradas, Pareto de motivos de perdas e tabela de distribuição de impactos cruzando a taxa horária de custos ociosos editada. Otimizado com tipografia Inter e espaçamentos densos ideais para salvamento em PDF ou impressão em papel A3 real. |
| 2026-05-26 | **Custos de Ociosidade Customizáveis por Setor (Aba 6)**: Substituído o custo hora global único por uma tabela interativa editável via `st.data_editor`. O usuário agora pode definir taxas horárias individualizadas por setor produtivo (máquina). O sistema salva as taxas automaticamente no `config.json` (chave `CUSTOS_SETORES`), recalculando em tempo real o prejuízo financeiro acumulado, com tooltips dinâmicos nos gráficos e detalhamento na tabela de ocorrências. |
| 2026-05-26 | **Filtro de Máquinas e Análise de Paradas (Aba 6)**: Implementado filtro de multiselect para marcar/desmarcar de forma interativa quais máquinas visualizar nos indicadores. Adicionada a nova seção **⏹️ Análise de Paradas e Ociosidade** no final dos dashboards, cruzando a aba PARADAS com os apontamentos para revelar as principais causas e tempos de inatividade por motivo e por máquina de forma altamente visual (Altair). |
| 2026-05-26 | **Relatório A3 — Produtividade Detalhada (v8)**: A coluna esquerda do A3 foi expandida de 3 para 5 seções. Adicionadas: (4) **Tabela de Qualidade por Máquina** com chapas e m² Normal vs. Refeito por máquina e barra de % índice de rejeito com cor semafórica (verde <3%, amarelo 3–10%, vermelho >10%); (5) **Distribuição por Processo Produtivo** com top 6 processos por m² e barra de participação percentual. Os KPIs globais agora exibem desdobramento Normal/Refeito inline. O contexto introdutório passou a incluir número de dias ativos, média de chapas/dia e média de m²/dia calculados dinamicamente. Seções da coluna direita renumeradas de 6 a 10. |
| 2026-05-26 | **Relatório A3 — Performance Integrada (v7)**: Reformulação completa do relatório A3 na Aba 6. Criado relatório boardroom-ready A3 Landscape consolidando TODOS os indicadores: produtividade (chapas + m² por máquina e turno), qualidade (refeito), paradas (pareto de motivos, impacto por máquina com taxa horária), ocorrências mais críticas (Top 5) e plano de ação Lean com 4 contramedidas e espaço para assinaturas. |
| 2026-05-26 | **Configurações Gerais Resilientes (Aba 7)**: Resolvido o travamento no apontamento de novos arquivos. Substituída a validação rígida de existência física (`os.path.exists`) por avisos não impeditivos (`st.warning`), permitindo salvar caminhos de rede temporariamente inacessíveis. Implementado reset inteligente do `st.session_state` ao Salvar/Restaurar para sincronização visual imediata. Adicionado painel de emergência auto-suficiente caso a planilha principal falhe no carregamento, evitando o travamento do app (`st.stop()`). |
| 2026-05-26 | **Filtros por Tipo de Produção**: Implementado isolamento de setores permitidos por tipo de processo. Adicionada a aba persistente `TIPO_SETORES` no Excel do PCP com auto-inicialização e tabela totalmente editável na interface (Aba 7). O formulário de apontamentos (Aba 4) agora restringe dinamicamente tanto a seleção do **Setor** quanto a do **Processo/Etapa** com base no tipo de processo escolhido para mitigar erros operacionais (ex: impede registrar processo de levigamento na máquina de resina ou selecionar processos incompatíveis). |
| 2026-05-25 | **Blocos com Dupla Identificação**: Adicionado suporte robusto para blocos com códigos múltiplos separados por barra `/` (ex: `4244/771418`). A lógica `dm.blocos_match` analisa a intersecção de códigos e foi integrada no formulário de busca de blocos (Aba 1), no preenchimento de máquinas por apontamento e filtros de correspondência PCP-fábrica (Aba 4), bem como na busca dinâmica no Estoque de Chapas e na validação WIP de blocos. |
| 2026-05-21 | **Uso de Data Fim para Indicadores**: Alterado os indicadores de produção (Aba 6) para usar a data e hora de término (`DIA_FIM` / `HORA_FIM`) como base para contabilização de chapas e M², garantindo que processos longos (como serrada) sejam contabilizados no dia de término da produção. Mantido fallback para data inicial se a finalização estiver vazia. |
| 2026-05-21 | **Fix Busca de Blocos (.xlsb)**: `get_bloco_info` reescrita com busca dinâmica de colunas via keywords (`find_col`). Antes falhava silenciosamente por hardcodar `N_BLOCO` que não existia (real: `Nº BLOCO` com encoding instável). Mesmo padrão aplicado para `COMP. (LIQUIDO)`, `ALT. (LIQUIDO)`, `LARG. (LIQUIDO)` |
| 2026-05-21 | **Fix % Refeito**: Percentual agora respeita a métrica selecionada (Chapas ou M²). Antes sempre calculava com M² independente da seleção. Label também é dinâmico |
| 2026-05-21 | **Fix Gráficos Empilhados (Aba 6)**: Labels centralizados dentro de cada segmento usando cálculo explícito de `_mid`. Corrigido erro "Facet charts require same data" unificando todas as camadas no mesmo DataFrame. Ordenação cronológica via lista explícita |
| 2026-05-15 | **Fix Crítico DATA_INICIO**: Corrigido o campo de data usado no Confronto. Antes usava `DATA_REG` (data de digitação), agora usa `DATA_INICIO` (data real do trabalho na fábrica) |
| 2026-05-15 | **Mapeamento Direto de Colunas**: Eliminada lógica genérica de "adivinhação" de colunas. Agora usa os nomes reais: `NUMERO_BLOCO`, `PROCESSO`, `SETOR`, `QTD_CHAPAS`, `MATERIAL` |
| 2026-05-15 | **Confronto Lado a Lado**: Nova interface na Aba 4 que divide a tela entre Plano (PCP) e Realizado (Fábrica) para auditoria instantânea |
| 2026-05-15 | **Sincronização Automática**: Carregamento inteligente de apontamentos ao mudar a data, com botão de refresh |
| 2026-05-14 | **Self-Healing de Setores**: Autopreenchimento de máquinas baseado na Base de Dados (corrige erros de fórmulas do Excel) |
| 2026-05-14 | **Clusterização UI**: Janela de Programação organizada em blocos lógicos (resiliência a erros de escopo e NameError) |
| 2026-05-14 | **Remoção de Agendamentos**: Interface permite desmarcar blocos para limpar datas e resetar status no banco de dados |
| 2026-05-14 | **Normalização de Blocos**: IDs forçados para String para evitar bugs de tipo `int` vs `str` |
| 2026-05-14 | **Indicadores (Aba 6)**: Dashboard de produção diária, por turno, máquina e análise de Refeito |
| 2026-05-14 | **Resinagem Manual**: Substituída lista fixa por entrada manual de nome, catalisador e proporção |
| 2026-05-14 | **Validação Estrita**: Paradas agora devem estar contidas no intervalo do processo |
| 2026-05-14 | **Base Granular**: Novas colunas `CABECAS`, `INSUMO_DETALHE`, `DIA_INICIO` e `DIA_FIM` no banco |
| 2026-05-14 | **Carrinho Auto-Reset**: Limpeza automática e recarregamento após salvar no Excel |
| 2026-05-13 | **Formulário Híbrido**: Separação em Dados Básicos (Form) + Insumos/Paradas (Live) |
| 2026-05-13 | **Upgrade REV 2**: Migração para nova estrutura de planilha de apontamento |
| 2026-05-13 | **Mapeamento Centralizado**: Refatoração do `data_manager` para suportar aliases de colunas |
| 2026-05-13 | **Validação de Tempos**: Bloqueio de horas inválidas e tempos negativos |
| 2026-05-13 | **Turno Automático**: Cálculo de turno D/N baseado na hora de início |
| 2026-05-13 | **Endurentes (LISTAS)**: Integração com tabela de proporções e tempo de secagem |
| 2026-05-13 | **Excel Pro**: Expansão automática de Tabelas (ListObjects) e cópia de estilos/bordas |
| 2026-05-13 | **Histórico Dinâmico**: Detecção de cabeçalhos variável para busca de histórico |
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
| 2026-05-12 | Git configurado no projeto; README.md e .gitignore criados |
| 2026-05-12 | Repositório conectado ao GitHub: `Magroni/IntegradorPCP` |

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

---

## 11. Armadilhas Conhecidas (Gotchas)

1. **Encoding `.xlsb`**: O engine `pyxlsb` lê caracteres especiais (º, ², ç, ã) com encoding instável. **NUNCA** hardcodar nomes de colunas de arquivos `.xlsb`. Use busca por keywords.

2. **Facet + Layer no Altair**: Gráficos facetados com múltiplas camadas (bars + text) exigem que **todas as camadas usem o mesmo DataFrame**. Usar DataFrames filtrados separados causa erro.

3. **Ordenação de eixo nominal no Altair**: `sort=alt.EncodingSortField(field=...)` nem sempre funciona com camadas ou facets. Prefira **lista explícita** de valores ordenados: `sort=['14/05', '15/05', ...]`.

4. **Stack order em barras empilhadas**: Para controlar qual segmento fica na base, use `order=alt.Order('_sort:Q')` com coluna numérica de ordenação (ex: Refeito=0, Normal=1).

5. **Validação Rígida no Streamlit**: Evite o uso de `st.stop()` global logo após a carga inicial de dados sem fornecer uma rota de emergência. Sempre que houver falha crítica ao carregar arquivos Excel, exiba a interface de configurações para que o usuário possa corrigir os caminhos de forma interativa na própria tela sem precisar editar manualmente o arquivo `config.json`.

