# -*- coding: utf-8 -*-
"""
Script de Transposição Avançada com Atrelamento Cronológico Preciso de Paradas (Corrigido)
REV 1 (Tabela Única) -> REV 2 (Estrutura Normalizada: DB, INSUMOS, PARADAS)
Costa Granitos - PCP
"""

import os
import pandas as pd
import openpyxl
import datetime as dt

# Caminhos dos arquivos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_REV1 = os.path.join(BASE_DIR, "Apontamento Produção (REV 1).xlsx")
FILE_REV2 = os.path.join(BASE_DIR, "Apontamento Produção (REV 2).xlsx")
FILE_OUTPUT = os.path.join(BASE_DIR, "Apontamento_Maio_Transposto_Auditoria.xlsx")
FILE_BLOCKS = os.path.join(BASE_DIR, "PLANILHA BLOCOS.xlsb")

print("======================================================================")
print("INICIANDO TRANSPOSIÇÃO DE DADOS: MAIO/2026 (ATRELAMENTO CHRONOLÓGICO DE PARADAS)")
print("======================================================================")

# Auxiliar de normalização de bloco
def normalize_bloco(b):
    if pd.isna(b) or str(b).strip() == "" or str(b).lower() == "nan":
        return ""
    b_str = str(b).strip().upper()
    if b_str.endswith(".0"):
        b_str = b_str[:-2]
    return b_str

# Auxiliar de normalização de setor
def normalize_sector(s):
    s = str(s).strip().upper()
    if "5-POLITRIZ" in s:
        return "5-POLITRIZ"
    if "3-POLITRIZ" in s:
        return "3-POLITRIZ"
    if "2-RESINA" in s:
        return "2-RESINA"
    if "RETOQUE" in s:
        return "RETOQUE"
    return s

# Auxiliar de criação de datetime robusto
def make_datetime(date_val, time_val):
    if pd.isna(date_val) or pd.isna(time_val):
        return None
    if hasattr(date_val, "to_pydatetime"):
        d = date_val.to_pydatetime().date()
    elif isinstance(date_val, dt.datetime):
        d = date_val.date()
    else:
        d = date_val
    
    if not isinstance(time_val, dt.time):
        try:
            t_str = str(time_val).strip()
            parts = t_str.split(":")
            h = int(parts[0])
            m = int(parts[1])
            s = int(parts[2]) if len(parts) > 2 else 0
            time_val = dt.time(h, m, s)
        except:
            return None
            
    return dt.datetime.combine(d, time_val)

# Auxiliar para formatar hora/tempo como string HH:MM
def format_time_hhmm(t):
    if pd.isna(t) or t == "" or str(t).strip() == "" or str(t).lower() == "nan":
        return ""
    if isinstance(t, dt.time):
        return f"{t.hour:02d}:{t.minute:02d}"
    try:
        t_str = str(t).strip()
        parts = t_str.split(":")
        h = int(parts[0])
        m = int(parts[1])
        return f"{h:02d}:{m:02d}"
    except:
        return str(t)

# Auxiliar para calcular duração entre dois tempos (cobrindo viradas de meia-noite)
def calculate_duration_hhmm(ini, fim):
    if pd.isna(ini) or pd.isna(fim):
        return ""
    
    def parse_to_time(val):
        if isinstance(val, dt.time):
            return val
        try:
            t_str = str(val).strip()
            parts = t_str.split(":")
            h = int(parts[0])
            m = int(parts[1])
            s = int(parts[2]) if len(parts) > 2 else 0
            return dt.time(h, m, s)
        except:
            return None
            
    t_ini = parse_to_time(ini)
    t_fim = parse_to_time(fim)
    
    if t_ini is None or t_fim is None:
        return ""
        
    d = dt.date(2026, 5, 1)
    dt_ini = dt.datetime.combine(d, t_ini)
    dt_fim = dt.datetime.combine(d, t_fim)
    
    if dt_fim < dt_ini:
        dt_fim += dt.timedelta(days=1)
        
    diff = dt_fim - dt_ini
    total_seconds = int(diff.total_seconds())
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    return f"{h:02d}:{m:02d}"

# --- PREPARAÇÃO DOS LOOKUPS ---

# 1. Carregar colunas e obter o último ID do novo template (REV 2)
print("\n[Passo 1] Analisando o template novo (REV 2)...")
try:
    df_temp_db = pd.read_excel(FILE_REV2, sheet_name="DB", engine="openpyxl")
    cols_rev2_db = df_temp_db.columns.tolist()
    if "ID" in df_temp_db.columns:
        ids = pd.to_numeric(df_temp_db["ID"], errors="coerce").dropna()
        start_id = int(ids.max()) + 1 if not ids.empty else 1
    else:
        start_id = 1
    print(f"-> Estrutura 'DB' carregada. Próximo ID inicial: {start_id}")
except Exception as e:
    print(f"Aviso ao ler REV 2: {e}. Usando ID inicial = 1.")
    start_id = 1
    cols_rev2_db = [
        "ID", "DATA_REG", "MATERIAL", "NUMERO_BLOCO", "PROCESSO", "SETOR", "ESP", "QTD_CHAPAS",
        "COMP", "ALT", "QTDM2", "OPERADOR", "DATA_INICIO", "DATA_FIM", "HORA_INICIO", "HORA_FIM",
        "TEMPO_PROCESSO", "TURNO"
    ]

# Mapeamento de dimensões oficiais da Planilha de Blocos
block_dims_map = {}
if os.path.exists(FILE_BLOCKS):
    try:
        print("-> Mapeando dimensões da Planilha de Blocos...")
        df_blocks = pd.read_excel(FILE_BLOCKS, sheet_name="PLAN. BLOCOS", skiprows=8, engine="pyxlsb")
        df_blocks.columns = [str(c).strip().upper() for c in df_blocks.columns]
        col_b = next((c for c in df_blocks.columns if "BLOCO" in c), None)
        col_comp = next((c for c in df_blocks.columns if "COMP" in c and "LIQ" in c), None)
        col_alt = next((c for c in df_blocks.columns if "ALT" in c and "LIQ" in c), None)
        if col_b:
            for _, row in df_blocks.iterrows():
                b_norm = normalize_bloco(row[col_b])
                if b_norm:
                    comp_val = float(pd.to_numeric(row[col_comp], errors="coerce") or 0) if col_comp else 0.0
                    alt_val = float(pd.to_numeric(row[col_alt], errors="coerce") or 0) if col_alt else 0.0
                    block_dims_map[b_norm] = {"COMP": comp_val, "ALT": alt_val}
    except Exception as e:
        print(f"Aviso ao mapear Planilha de Blocos: {e}")

# 2. Carregar arquivo antigo (REV 1)
print(f"\n[Passo 2] Lendo o arquivo antigo '{os.path.basename(FILE_REV1)}'...")
try:
    df_rev1 = pd.read_excel(FILE_REV1, sheet_name="BD", skiprows=6, engine="openpyxl")
    print(f"-> Planilha antiga carregada! Total de linhas brutas: {len(df_rev1)}")
except Exception as e:
    print(f"Erro ao carregar REV 1: {e}")
    exit(1)

# Normaliza nomes de colunas do REV 1
df_rev1.columns = [str(c).strip().upper() for c in df_rev1.columns]

# Criar coluna com bloco normalizado
def extract_bloco_rev1(row):
    mat_blo = str(row.get("MATERIAL+BLOCO", ""))
    num_bloco = str(row.get("Nº BLOCO", ""))
    if pd.notna(row.get("Nº BLOCO")) and num_bloco.strip() != "" and num_bloco.lower() != "nan":
        return normalize_bloco(num_bloco)
    parts = []
    if " - " in mat_blo:
        parts = mat_blo.split(" - ")
    elif "-" in mat_blo:
        parts = mat_blo.split("-")
    if len(parts) > 1:
        return normalize_bloco(parts[1])
    return normalize_bloco(mat_blo)

df_rev1["BLOCO_NORM"] = df_rev1.apply(extract_bloco_rev1, axis=1)

# 3. Filtrar registros do mês de Maio de 2026
print("\n[Passo 3] Filtrando registros de Maio/2026...")
col_data_filt = "DATA REG" if "DATA REG" in df_rev1.columns else df_rev1.columns[0]
df_rev1["DATETIME_CONVERTED"] = pd.to_datetime(df_rev1[col_data_filt], errors="coerce")

df_may = df_rev1[
    (df_rev1["DATETIME_CONVERTED"].dt.month == 5) & 
    (df_rev1["DATETIME_CONVERTED"].dt.year == 2026)
].copy()

print(f"-> Encontrados {len(df_may)} registros no total em Maio/2026.")

# 4. Separar registros entre PRODUÇÃO e PARADAS (Classificação Avançada Refinada)
print("\n[Passo 4] Classificando e separando apontamentos de PRODUÇÃO das PARADAS de máquina...")

# Função de classificação precisa de Parada (inclui Checklist e Lixamento/Levantamento)
def is_parada(row):
    proc = str(row.get("PROCESSO", "")).strip().upper()
    mat_blo = str(row.get("MATERIAL+BLOCO", "")).strip().upper()
    
    if proc == "PARADA":
        return True
        
    parada_keywords = [
        'ALMOÇO', 'ALMOCO', 'ABRASIVO', 'PARADA', 'TROCA', 'AJUSTE', 'SETUP', 
        'ABASTECER', 'AGUARDANDO', 'INTERVALO', 'MANUTENÇÃO', 'MANUTENCAO', 
        'FALTA', 'JANTA', 'ELÉTRICA', 'ELETRICA',
        'CHECK', 'CHECKLIST', 'LIXAMENTO', 'LEVANTAMENTO', 'LEVANTANDO'
    ]
    
    for kw in parada_keywords:
        if kw in mat_blo or kw in proc:
            return True
            
    return False

df_may["IS_PARADA"] = df_may.apply(is_parada, axis=1)

df_prod_raw = df_may[~df_may["IS_PARADA"]].copy()
df_paradas_raw = df_may[df_may["IS_PARADA"]].copy()

print(f"-> Apontamentos de Produção Efetiva: {len(df_prod_raw)} linhas.")
print(f"-> Ocorrências de Paradas de Máquina: {len(df_paradas_raw)} linhas.")

# 5. Criar e preencher a nova aba "DB" (Produção)
print("\n[Passo 5] Transpondo dados de produção (Aba DB)...")
df_final_db = pd.DataFrame(columns=cols_rev2_db)

if not df_prod_raw.empty:
    # Gerar ID sequencial dinâmico
    df_prod_raw["NEW_ID"] = range(start_id, start_id + len(df_prod_raw))

    # Mapear campos básicos
    df_final_db["ID"] = df_prod_raw["NEW_ID"]
    df_final_db["DATA_REG"] = df_prod_raw["DATETIME_CONVERTED"]
    df_final_db["DATA_INICIO"] = df_prod_raw["DATETIME_CONVERTED"]
    df_final_db["DATA_FIM"] = df_prod_raw["DATETIME_CONVERTED"]

    # Processo, Setor, Operador e Turno
    df_final_db["PROCESSO"] = df_prod_raw["PROCESSO"].fillna("").astype(str).str.strip().str.upper()
    df_final_db["SETOR"] = df_prod_raw["SETOR"].fillna("").astype(str).str.strip().str.upper()
    df_final_db["OPERADOR"] = df_prod_raw["OPERADOR"].fillna("").astype(str).str.strip().str.upper()
    df_final_db["TURNO"] = df_prod_raw["TURNO"].fillna("").astype(str).str.strip().str.upper()

    # Horários e tempos formatados de forma robusta e amigável para o Excel
    df_final_db["HORA_INICIO"] = df_prod_raw["HORIM. INI"].apply(format_time_hhmm)
    df_final_db["HORA_FIM"] = df_prod_raw["HORIM. FIM"].apply(format_time_hhmm)
    df_final_db["TEMPO_PROCESSO"] = df_prod_raw.apply(
        lambda r: calculate_duration_hhmm(r["HORIM. INI"], r["HORIM. FIM"]), axis=1
    )

    # Material e Bloco
    def split_material(row):
        mat_blo = str(row.get("MATERIAL+BLOCO", ""))
        nome_mat = str(row.get("NOME MATERIAL", ""))
        if pd.notna(row.get("NOME MATERIAL")) and nome_mat.strip() != "" and nome_mat.lower() != "nan":
            return nome_mat.strip().upper()
        if " - " in mat_blo:
            return mat_blo.split(" - ")[0].strip().upper()
        if "-" in mat_blo:
            return mat_blo.split("-")[0].strip().upper()
        return mat_blo.strip().upper()

    df_final_db["MATERIAL"] = df_prod_raw.apply(split_material, axis=1)
    df_final_db["NUMERO_BLOCO"] = df_prod_raw["BLOCO_NORM"]

    # Mapear quantidade de chapas a partir da coluna '3CM'
    df_final_db["QTD_CHAPAS"] = pd.to_numeric(df_prod_raw["3CM"], errors="coerce").fillna(0).astype(int)

    # Definir espessura baseado em 3CM ou 2CM
    def extrair_espessura(row):
        val_3cm = pd.to_numeric(row.get("3CM", 0), errors="coerce")
        val_2cm = pd.to_numeric(row.get("2CM", 0), errors="coerce")
        if pd.notna(val_3cm) and val_3cm > 0: return 3
        if pd.notna(val_2cm) and val_2cm > 0: return 2
        return 3 # default

    df_final_db["ESP"] = df_prod_raw.apply(extrair_espessura, axis=1)

    # Dimensões (Comprimento e Altura) com preenchimento da Planilha de Blocos se estiver vazio
    df_final_db["COMP"] = pd.to_numeric(df_prod_raw["COMP."], errors="coerce")
    df_final_db["ALT"] = pd.to_numeric(df_prod_raw["ALT."], errors="coerce")

    for idx, row in df_prod_raw.iterrows():
        db_idx = df_final_db[df_final_db["ID"] == row["NEW_ID"]].index[0]
        b_norm = row["BLOCO_NORM"]
        
        comp_val = df_final_db.at[db_idx, "COMP"]
        alt_val = df_final_db.at[db_idx, "ALT"]
        
        if (pd.isna(comp_val) or comp_val <= 0 or pd.isna(alt_val) or alt_val <= 0) and b_norm:
            if b_norm in block_dims_map:
                df_final_db.at[db_idx, "COMP"] = block_dims_map[b_norm]["COMP"]
                df_final_db.at[db_idx, "ALT"] = block_dims_map[b_norm]["ALT"]

    # Cálculo do M² baseado na coluna '3CM'
    m2_calculado_cnt = 0
    for idx, row in df_prod_raw.iterrows():
        db_idx = df_final_db[df_final_db["ID"] == row["NEW_ID"]].index[0]
        
        chapas = df_final_db.at[db_idx, "QTD_CHAPAS"]
        comp = df_final_db.at[db_idx, "COMP"]
        alt = df_final_db.at[db_idx, "ALT"]
        
        m2_orig = pd.to_numeric(row.get("QTD M² (SEM RET & REPASSE)"), errors="coerce")
        m2_com_rep = pd.to_numeric(row.get("QTD M² (COM RET & REPASSE)"), errors="coerce")
        
        m2_final = 0.0
        if pd.notna(m2_orig) and m2_orig > 0:
            m2_final = float(m2_orig)
        elif pd.notna(m2_com_rep) and m2_com_rep > 0:
            m2_final = float(m2_com_rep)
            
        if (m2_final == 0.0 or pd.isna(m2_final)) and chapas > 0:
            if pd.notna(comp) and comp > 0 and pd.notna(alt) and alt > 0:
                m2_final = float(round(comp * alt * chapas, 3))
                m2_calculado_cnt += 1
                
        df_final_db.at[db_idx, "QTDM2"] = m2_final

    print(f"-> Cálculo de M² concluído: {m2_calculado_cnt} metragens calculadas.")

# 6. Criar e preencher a nova aba "INSUMOS" (Resinas, Endurecedores e Abrasivos)
print("\n[Passo 6] Extraindo e estruturando Insumos de Produção (Aba INSUMOS)...")
insumos_list = []

if not df_prod_raw.empty:
    for _, row in df_prod_raw.iterrows():
        rec_id = row["NEW_ID"]
        
        # A. RESINA
        resina_tipo = str(row.get("TPO RESINA", "")).strip()
        resina_qtd = pd.to_numeric(row.get("QTD.KG", 0), errors="coerce")
        if pd.notna(resina_tipo) and resina_tipo != "" and resina_tipo.lower() != "nan" and resina_tipo != "--":
            insumos_list.append({
                "ID_APONTAMENTO": rec_id,
                "TIPO_INSUMO": "RESINA",
                "DESCRICAO": resina_tipo.upper(),
                "QUANTIDADE": resina_qtd if pd.notna(resina_qtd) else 0.0,
                "UNIDADE": "KG",
                "TEMPO_SECAGEM": pd.NA
            })
            
        # B. ENDURECEDOR
        endur_tipo = str(row.get("TIPO.ENDUR", "")).strip()
        endur_qtd = pd.to_numeric(row.get("QTD.KG.1", 0), errors="coerce")
        tempo_secagem = pd.to_numeric(row.get("24H", 24.0), errors="coerce")
        if pd.notna(endur_tipo) and endur_tipo != "" and endur_tipo.lower() != "nan" and endur_tipo != "--":
            insumos_list.append({
                "ID_APONTAMENTO": rec_id,
                "TIPO_INSUMO": "ENDURECEDOR",
                "DESCRICAO": endur_tipo.upper(),
                "QUANTIDADE": endur_qtd if pd.notna(endur_qtd) else 0.0,
                "UNIDADE": "KG",
                "TEMPO_SECAGEM": tempo_secagem if pd.notna(tempo_secagem) else 24.0
            })

        # C. ABRASIVOS (SAT1 a SAT20)
        for i in range(1, 21):
            col_name = f"SEQ. ABR. {i}"
            if col_name in row:
                abr_val = str(row[col_name]).strip()
                if pd.notna(row[col_name]) and abr_val != "" and abr_val.lower() != "nan" and abr_val != "-----":
                    insumos_list.append({
                        "ID_APONTAMENTO": rec_id,
                        "TIPO_INSUMO": "ABRASIVO",
                        "DESCRICAO": f"CAB {i:02d}: {abr_val.upper()}",
                        "QUANTIDADE": 1.0,
                        "UNIDADE": "UNID",
                        "TEMPO_SECAGEM": pd.NA
                    })

df_final_insumos = pd.DataFrame(insumos_list, columns=["ID_APONTAMENTO", "TIPO_INSUMO", "DESCRICAO", "QUANTIDADE", "UNIDADE", "TEMPO_SECAGEM"])

# 7. Criar e preencher a nova aba "PARADAS" (Com atrelamento cronológico inteligente corrigido de forma definitiva)
print("\n[Passo 7] Transpondo ocorrências de inatividade (Aba PARADAS) com Atrelamento Inteligente...")

# Pré-construir datetimes para busca cronológica precisa e rápida
if not df_prod_raw.empty:
    df_prod_raw["PROD_DATETIME"] = df_prod_raw.apply(
        lambda r: make_datetime(r["DATETIME_CONVERTED"], r["HORIM. INI"]), axis=1
    )

paradas_list = []

for idx, row in df_paradas_raw.iterrows():
    # O motivo da parada na REV 1 fica no campo MATERIAL+BLOCO
    motivo = str(row.get("MATERIAL+BLOCO", "PARADA")).strip()
    data_dia = row["DATETIME_CONVERTED"]
    hora_ini = row.get("HORIM. INI")
    hora_fim = row.get("HORIM. FIM")
    duracao = row.get("TEMPO EFETIVO PRODUÇÃO/PARADA")
    setor_parada = str(row.get("SETOR", "")).strip().upper()
    
    stop_dt = make_datetime(data_dia, hora_ini)
    norm_setor = normalize_sector(setor_parada)
    
    id_apontamento = pd.NA
    if not df_prod_raw.empty and stop_dt is not None:
        # A. Primeiro tenta atrelar ao processo do mesmo setor/máquina (usando a normalização de setor)
        df_same_sector = df_prod_raw[
            df_prod_raw["SETOR"].apply(normalize_sector) == norm_setor
        ].copy()
        
        if not df_same_sector.empty:
            # Calcula distâncias cronológicas absolutas em segundos
            df_same_sector["TIME_DIFF"] = (df_same_sector["PROD_DATETIME"] - stop_dt).apply(lambda x: abs(x.total_seconds()))
            idx_closest = df_same_sector["TIME_DIFF"].idxmin()
            id_apontamento = int(df_same_sector.loc[idx_closest, "NEW_ID"])
            
        # B. Fallback caso não ache no mesmo setor: pega o processo mais próximo cronologicamente de forma geral
        if pd.isna(id_apontamento):
            df_temp = df_prod_raw.copy()
            df_temp["TIME_DIFF"] = (df_temp["PROD_DATETIME"] - stop_dt).apply(lambda x: abs(x.total_seconds()))
            idx_closest = df_temp["TIME_DIFF"].idxmin()
            id_apontamento = int(df_temp.loc[idx_closest, "NEW_ID"])
            
    paradas_list.append({
        "ID_APONTAMENTO": id_apontamento,
        "MOTIVO": motivo.upper(),
        "DATA_INICIO": data_dia,
        "DATA_FIM": data_dia,
        "HORA_INICIO": format_time_hhmm(hora_ini),
        "HORA_FIM": format_time_hhmm(hora_fim),
        "TEMPO": calculate_duration_hhmm(hora_ini, hora_fim)
    })

df_final_paradas = pd.DataFrame(paradas_list, columns=["ID_APONTAMENTO", "MOTIVO", "DATA_INICIO", "DATA_FIM", "HORA_INICIO", "HORA_FIM", "TEMPO"])

# Formatação de datas para string DD/MM/AAAA
for date_col in ["DATA_REG", "DATA_INICIO", "DATA_FIM"]:
    if date_col in df_final_db.columns:
        df_final_db[date_col] = pd.to_datetime(df_final_db[date_col], errors="coerce").dt.strftime("%d/%m/%Y")
        
for date_col in ["DATA_INICIO", "DATA_FIM"]:
    if date_col in df_final_paradas.columns:
        df_final_paradas[date_col] = pd.to_datetime(df_final_paradas[date_col], errors="coerce").dt.strftime("%d/%m/%Y")

# 8. Gravar o novo arquivo de Auditoria com as 3 abas separadas
print(f"\n[Passo 8] Gravando a planilha final de Auditoria '{os.path.basename(FILE_OUTPUT)}'...")
try:
    with pd.ExcelWriter(FILE_OUTPUT, engine="openpyxl") as writer:
        df_final_db.to_excel(writer, sheet_name="DB", index=False)
        df_final_insumos.to_excel(writer, sheet_name="INSUMOS", index=False)
        df_final_paradas.to_excel(writer, sheet_name="PARADAS", index=False)
        
    print("\n======================================================================")
    print("SUCESSO! Transposição estrutural concluída com Atrelamento de Paradas!")
    print(f"-> Planilha gerada: {FILE_OUTPUT}")
    print(f"-> Registros de Produção (Aba 'DB'): {len(df_final_db)} linhas")
    print(f"-> Abrasivos e Insumos Químicos (Aba 'INSUMOS'): {len(df_final_insumos)} linhas")
    print(f"-> Paradas e Inatividades (Aba 'PARADAS'): {len(df_final_paradas)} linhas")
    print("   * Todas as paradas foram atreladas aos IDs de processos correspondentes de forma definitiva!")
    print("======================================================================")
except Exception as e:
    print(f"Erro ao salvar arquivo Excel: {e}")

input("\nPressione ENTER para fechar...")
