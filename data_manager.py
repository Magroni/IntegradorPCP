# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
import os
import json

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DE CAMINHOS (lidos do config.json em runtime)
# ---------------------------------------------------------------------------
_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

_DEFAULT_CONFIG = {
    "DB_FILE": r"z:\PCP\PROJETOS MARLON\ProgramarProd\COSTA GRAN. - PROGRAMAÇÕES - BASE DE DADOS.xlsm",
    "APONTAMENTO_FILE": r"z:\PCP\PROJETOS MARLON\ProgramarProd\Apontamento Produção (REV 1).xlsx",
    # Nomes das abas (configuráveis pelo usuário)
    "SHEET_PROGRAMACAO": "DB",
    "SHEET_ENTREGUES": "ENTREGUES",
    "SHEET_BASE_DADOS": "BASE DE DADOS",
    "SHEET_AP_BD": "DB",
    "SHEET_AP_BASE": "BASE DADOS"
}


def get_config():
    """Lê o config.json e retorna o dicionário de configurações."""
    try:
        if os.path.exists(_CONFIG_PATH):
            with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in _DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
    except Exception as e:
        print(f"Erro ao ler config.json: {e}")
    return dict(_DEFAULT_CONFIG)


def save_config(cfg: dict):
    """Salva o dicionário de configurações no config.json."""
    try:
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"Erro ao salvar config.json: {e}")
        return False


def _get_db_file():
    return get_config()["DB_FILE"]


def _get_apontamento_file():
    return get_config()["APONTAMENTO_FILE"]


def _get_sheet(key: str) -> str:
    """Retorna o nome da aba configurada, com fallback para o default."""
    return get_config().get(key, _DEFAULT_CONFIG[key])


def get_sheet_names(filepath: str):
    """Retorna a lista de nomes de abas de um arquivo Excel."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception as e:
        print(f"Erro ao ler abas de {filepath}: {e}")
        return []


# Mantém compatibilidade com código que usa as constantes diretamente
DB_FILE = _get_db_file()
APONTAMENTO_FILE = _get_apontamento_file()

# ---------------------------------------------------------------------------
# LEITURA DE DADOS
# ---------------------------------------------------------------------------

def get_data():
    """Lê a aba de Programação do arquivo de Programação."""
    try:
        df = pd.read_excel(_get_db_file(), sheet_name=_get_sheet("SHEET_PROGRAMACAO"), engine="openpyxl")
        return df
    except Exception as e:
        print(f"Erro ao carregar os dados: {e}")
        return pd.DataFrame()


def get_headers():
    """Lê os cabeçalhos diretamente com openpyxl para ter um mapa de colunas exato."""
    try:
        wb = openpyxl.load_workbook(_get_db_file(), read_only=True, data_only=True)
        ws = wb[_get_sheet("SHEET_PROGRAMACAO")]
        headers = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col).value
            if val:
                headers[str(val).strip()] = col
        wb.close()
        return headers
    except Exception as e:
        print(f"Erro ao ler cabeçalhos: {e}")
        return {}


def get_base_dados():
    """Lê a aba de Base de Dados e retorna o DataFrame."""
    try:
        df = pd.read_excel(_get_db_file(), sheet_name=_get_sheet("SHEET_BASE_DADOS"), engine="openpyxl")
        return df
    except Exception as e:
        print(f"Erro ao carregar base de dados: {e}")
        return pd.DataFrame()


def get_data_entregues():
    """Lê a aba de Entregues para cálculo de capacidade."""
    try:
        df = pd.read_excel(_get_db_file(), sheet_name=_get_sheet("SHEET_ENTREGUES"), engine="openpyxl")
        return df
    except Exception as e:
        print(f"Erro ao carregar entregues: {e}")
        return pd.DataFrame()


# ---------------------------------------------------------------------------
# CÁLCULOS
# ---------------------------------------------------------------------------

def get_historico_medias(df):
    """
    Calcula a média histórica de QTD. CHAPAS por dia para cada PROCESSO.
    Considera apenas registros com STATUS PROCESSO == 'REALIZADO'.
    """
    try:
        realizados = df[df["STATUS PROCESSO"] == "REALIZADO"].copy()
        if realizados.empty:
            return {}

        def parse_data(val):
            if pd.isna(val) or val == "": return pd.NaT
            if isinstance(val, pd.Timestamp): return val
            try: return pd.to_datetime(str(val), format="%d/%m/%Y")
            except:
                try: return pd.to_datetime(str(val))
                except: return pd.NaT

        realizados["DATA_REALIZADA_DT"] = realizados["DATA REALIZADA"].apply(parse_data)
        realizados = realizados.dropna(subset=["DATA_REALIZADA_DT"])
        realizados["QTD. CHAPAS"] = pd.to_numeric(realizados["QTD. CHAPAS"], errors="coerce").fillna(0)

        producao_diaria = realizados.groupby(["PROCESSO", "DATA_REALIZADA_DT"])["QTD. CHAPAS"].sum().reset_index()
        medias = producao_diaria.groupby("PROCESSO")["QTD. CHAPAS"].mean().to_dict()
        return medias
    except Exception as e:
        print(f"Erro ao calcular médias: {e}")
        return {}


def get_historico_medias_entregues():
    """Calcula a média de capacidade usando a base histórica ENTREGUES."""
    df = get_data_entregues()
    if df.empty:
        return {}
    try:
        def parse_data(val):
            if pd.isna(val) or val == "": return pd.NaT
            if isinstance(val, pd.Timestamp): return val
            try: return pd.to_datetime(str(val), format="%d/%m/%Y")
            except:
                try: return pd.to_datetime(str(val))
                except: return pd.NaT

        col_data = "DATA REALIZADA" if "DATA REALIZADA" in df.columns else "DATA"
        df["DATA_REALIZADA_DT"] = df[col_data].apply(parse_data)
        df = df.dropna(subset=["DATA_REALIZADA_DT"])
        df["QTD. CHAPAS"] = pd.to_numeric(df.get("QTD. CHAPAS", 0), errors="coerce").fillna(0)

        producao_diaria = df.groupby(["SETOR", "DATA_REALIZADA_DT"])["QTD. CHAPAS"].sum().reset_index()
        medias = producao_diaria.groupby("SETOR")["QTD. CHAPAS"].mean().to_dict()
        return medias
    except Exception as e:
        print(f"Erro ao calcular médias de entregues: {e}")
        return {}


# ---------------------------------------------------------------------------
# VALIDAÇÕES DE NEGÓCIO
# ---------------------------------------------------------------------------

def validar_sequencia_bloco(df, bloco_id, index_atual, nova_data_str):
    """
    Valida se a etapa anterior do bloco foi finalizada ou tem data agendada
    anterior ou igual à nova_data.
    """
    try:
        if pd.isna(bloco_id) or str(bloco_id).strip() == "":
            return True, ""

        df_bloco = df[df["BLOCO"] == bloco_id].copy()
        df_bloco = df_bloco.sort_index()
        indices_bloco = df_bloco.index.tolist()

        if index_atual not in indices_bloco:
            return True, ""

        pos = indices_bloco.index(index_atual)

        if pos == 0:
            return True, ""

        index_anterior = indices_bloco[pos - 1]
        linha_anterior = df_bloco.loc[index_anterior]

        status_anterior = str(linha_anterior.get("STATUS PROCESSO", ""))
        if status_anterior == "REALIZADO":
            return True, ""

        data_anterior_str = linha_anterior.get("DATA", "")
        if pd.isna(data_anterior_str) or str(data_anterior_str).strip() == "":
            proc = linha_anterior.get("PROCESSO", "Desconhecido")
            return False, f"O processo anterior '{proc}' não foi agendado nem concluído."

        nova_data_dt = pd.to_datetime(nova_data_str, format="%d/%m/%Y")

        try:
            if isinstance(data_anterior_str, pd.Timestamp):
                data_anterior_dt = data_anterior_str
            else:
                data_anterior_str_clean = str(data_anterior_str).strip()
                if "/" in data_anterior_str_clean:
                    data_anterior_dt = pd.to_datetime(data_anterior_str_clean, format="%d/%m/%Y")
                else:
                    data_anterior_dt = pd.to_datetime(data_anterior_str_clean)

            if data_anterior_dt.date() > nova_data_dt.date():
                proc = linha_anterior.get("PROCESSO", "Desconhecido")
                return False, f"O processo anterior '{proc}' está agendado para {data_anterior_dt.strftime('%d/%m/%Y')}, que é depois da data escolhida."
        except Exception:
            pass  # Se não conseguir parsear, deixa passar

        return True, ""
    except Exception as e:
        print(f"Erro ao validar bloco: {e}")
        return False, "Erro interno na validação."


# ---------------------------------------------------------------------------
# ESCRITA DE DADOS
# ---------------------------------------------------------------------------

def add_record(record_dict):
    """Adiciona um novo registro preservando o formato e macros."""
    try:
        wb = openpyxl.load_workbook(_get_db_file(), keep_vba=True)
        ws = wb[_get_sheet("SHEET_PROGRAMACAO")]

        next_row = ws.max_row + 1
        for row in range(2, ws.max_row + 2):
            if ws.cell(row=row, column=1).value is None:
                next_row = row
                break

        headers = get_headers()

        for key, value in record_dict.items():
            if key in headers:
                col_idx = headers[key]
                ws.cell(row=next_row, column=col_idx, value=value)

        wb.save(_get_db_file())
        return True
    except Exception as e:
        print(f"Erro ao adicionar registro: {e}")
        return False


def add_records(records_list):
    """Adiciona múltiplos registros de uma vez (útil para roteiros de blocos)."""
    try:
        wb = openpyxl.load_workbook(_get_db_file(), keep_vba=True)
        ws = wb[_get_sheet("SHEET_PROGRAMACAO")]

        next_row = ws.max_row + 1
        for row in range(2, ws.max_row + 2):
            if ws.cell(row=row, column=1).value is None:
                next_row = row
                break

        headers = get_headers()

        for record_dict in records_list:
            for key, value in record_dict.items():
                if key in headers:
                    col_idx = headers[key]
                    ws.cell(row=next_row, column=col_idx, value=value)
            next_row += 1

        wb.save(_get_db_file())
        return True
    except Exception as e:
        print(f"Erro ao adicionar múltiplos registros: {e}")
        return False


def update_cell_by_row(df_index, updates_dict):
    """
    Atualiza células específicas baseando-se no índice do DataFrame.
    O índice do DataFrame 0 corresponde à linha 2 do Excel (linha 1 = cabeçalho).
    """
    try:
        wb = openpyxl.load_workbook(_get_db_file(), keep_vba=True)
        ws = wb[_get_sheet("SHEET_PROGRAMACAO")]

        excel_row = df_index + 2
        headers = get_headers()

        for col_name, new_value in updates_dict.items():
            if col_name in headers:
                col_idx = headers[col_name]
                ws.cell(row=excel_row, column=col_idx, value=new_value)

        wb.save(_get_db_file())
        return True
    except Exception as e:
        print(f"Erro ao atualizar célula: {e}")
        return False


def update_base_dados(df_novo):
    """Sobrescreve as colunas PROCESSO e SETOR na aba de Base de Dados."""
    try:
        wb = openpyxl.load_workbook(_get_db_file(), keep_vba=True)
        ws = wb[_get_sheet("SHEET_BASE_DADOS")]

        col_proc = None
        col_setor = None

        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col).value
            if val == "PROCESSO": col_proc = col
            elif val == "SETOR": col_setor = col

        if not col_proc or not col_setor:
            return False

        max_row = ws.max_row
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_proc).value = None
            ws.cell(row=r, column=col_setor).value = None

        for idx, row in df_novo.iterrows():
            proc_val = str(row.get("PROCESSO", "")).strip()
            setor_val = str(row.get("SETOR", "")).strip()
            if proc_val and proc_val != "nan":
                ws.cell(row=idx + 2, column=col_proc, value=proc_val)
                ws.cell(row=idx + 2, column=col_setor, value=setor_val if setor_val != "nan" else "")

        wb.save(_get_db_file())
        return True
    except Exception as e:
        print(f"Erro ao atualizar base de dados: {e}")
        return False


def salvar_edicao_bloco_excel(bloco_id, material, demanda, qtd_chapas, vol_m2, roteiro_atual):
    """
    Sincroniza as edições de um bloco existente diretamente no Excel.
    Reescreve as linhas nas posições originais, adiciona novas linhas logo abaixo
    se o roteiro cresceu, ou deleta linhas finais se o roteiro diminuiu.
    """
    try:
        wb = openpyxl.load_workbook(_get_db_file(), keep_vba=True)
        ws = wb[_get_sheet("SHEET_PROGRAMACAO")]
        headers = get_headers()

        col_bloco = headers.get("BLOCO")
        old_rows = []
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=col_bloco).value).strip() == str(bloco_id).strip():
                old_rows.append(r)

        num_new = len(roteiro_atual)
        num_old = len(old_rows)
        last_written_row = max(old_rows) if old_rows else ws.max_row

        for i, step in enumerate(roteiro_atual):
            if i < num_old:
                row_to_write = old_rows[i]
            else:
                insert_idx = last_written_row + 1
                ws.insert_rows(insert_idx, 1)
                row_to_write = insert_idx

            last_written_row = row_to_write

            if headers.get("BLOCO"): ws.cell(row=row_to_write, column=headers["BLOCO"], value=bloco_id)
            if headers.get("MATERIAL"): ws.cell(row=row_to_write, column=headers["MATERIAL"], value=material)
            if headers.get("DEMANDA"): ws.cell(row=row_to_write, column=headers["DEMANDA"], value=demanda)
            if headers.get("QTD. CHAPAS"): ws.cell(row=row_to_write, column=headers["QTD. CHAPAS"], value=qtd_chapas)
            if headers.get("VOLUME M²"): ws.cell(row=row_to_write, column=headers["VOLUME M²"], value=vol_m2)

            if headers.get("PROCESSO"): ws.cell(row=row_to_write, column=headers["PROCESSO"], value=step.get("PROCESSO", ""))
            if headers.get("SETOR"): ws.cell(row=row_to_write, column=headers["SETOR"], value=step.get("SETOR", ""))
            if headers.get("OBSERVAÇÃO DE PRODUÇÃO"): ws.cell(row=row_to_write, column=headers["OBSERVAÇÃO DE PRODUÇÃO"], value=step.get("OBSERVACAO", ""))

            status = step.get("STATUS PROCESSO", "NÃO REALIZADO")
            if headers.get("STATUS PROCESSO"): ws.cell(row=row_to_write, column=headers["STATUS PROCESSO"], value=status)

            data_str = step.get("DATA", "")
            if data_str and headers.get("DATA"): ws.cell(row=row_to_write, column=headers["DATA"], value=data_str)

            data_realizada = step.get("DATA REALIZADA", "")
            if data_realizada and headers.get("DATA REALIZADA"): ws.cell(row=row_to_write, column=headers["DATA REALIZADA"], value=data_realizada)

        # Deletar as sobras físicas, de baixo para cima
        if num_old > num_new:
            sobras = old_rows[num_new:]
            for r in reversed(sobras):
                ws.delete_rows(r, 1)

        wb.save(_get_db_file())
        return True
    except Exception as e:
        print(f"Erro ao salvar edição do bloco no excel: {e}")
        return False


# ---------------------------------------------------------------------------
# APONTAMENTO DE PRODUÇÃO
# ---------------------------------------------------------------------------

def get_mapa_resumido_processos():
    """
    Lê a aba de mapeamento do arquivo de Apontamento e retorna um dicionário
    mapeando o nome completo do processo para o nome resumido.
    Ex: '19-POLIMENTO (S)' -> 'POLIMENTO'
    """
    try:
        df_bd = pd.read_excel(
            _get_apontamento_file(),
            sheet_name=_get_sheet("SHEET_AP_BASE"),
            header=0,
            engine="openpyxl",
            usecols=[0, 1]
        )
        df_bd.columns = ["PROCESSO_COMPLETO", "RESUMIDO"]
        df_bd = df_bd.dropna(subset=["PROCESSO_COMPLETO"])
        mapa = {}
        for _, row in df_bd.iterrows():
            proc = str(row["PROCESSO_COMPLETO"]).strip().upper()
            res = str(row["RESUMIDO"]).strip().upper() if pd.notna(row["RESUMIDO"]) else proc
            mapa[proc] = res
        return mapa
    except Exception as e:
        print(f"Erro ao ler BASE DADOS do Apontamento: {e}")
        return {}


def get_apontamentos_do_dia(data_alvo_date):
    """
    Lê a aba de apontamentos diários do arquivo de Apontamento (cabeçalho na linha 7)
    e retorna um DataFrame filtrado pela data_alvo_date.
    Colunas retornadas: BLOCO, NOME_MATERIAL, PROCESSO_APONTADO, RESUMIDO, SETOR_AP, QTD_CH, DATA_REG
    """
    try:
        # Busca dinâmica do cabeçalho (procura pela linha que contém 'PROCESSO' ou 'DATA REG')
        df_full = pd.read_excel(_get_apontamento_file(), sheet_name=_get_sheet("SHEET_AP_BD"), header=None, engine="openpyxl", nrows=20)
        header_row = 6 # fallback padrão
        for i, row in df_full.iterrows():
            row_vals = [str(v).strip().upper() for v in row.values if pd.notna(v)]
            if "PROCESSO" in row_vals or "DATA REG" in row_vals or "MATERIAL+BLOCO" in row_vals:
                header_row = i
                break
        
        df = pd.read_excel(
            _get_apontamento_file(),
            sheet_name=_get_sheet("SHEET_AP_BD"),
            skiprows=header_row,
            engine="openpyxl"
        )
        # Limpar nomes de colunas (há espaços e normalizar para maiúsculo)
        orig_cols = [str(c).strip().upper() for c in df.columns]
        df.columns = orig_cols

        # Mapeamento robusto (procura por nomes exatos ou parciais)
        mapping = {
            "DATA_REG": ["DATA REG", "DATA", "DATA_REG"],
            "MAT_BLOCO": ["MATERIAL+BLOCO", "MAT+BLO", "MATERIAL BLOCO"],
            "NOME_MATERIAL": ["NOME MATERIAL", "MATERIAL", "NOME_MATERIAL"],
            "BLOCO_RAW": ["Nº BLOCO", "NUM BLOCO", "BLOCO", "N BLOCO"],
            "PROCESSO_APONTADO": ["PROCESSO", "PROC", "PROCESSO_APONTADO"],
            "SETOR_AP": ["SETOR", "MAQUINA", "MÁQUINA", "SETOR_AP"],
            "QTD_CH": ["QTD CH (SEM RET & REPASSE)", "QTD CH", "CHAPAS", "QTD_CH"],
            "QTD_M2": ["QTD M² (SEM RET & REPASSE)", "QTD M2", "METRAGEM", "QTD_M2"]
        }

        rename_dict = {}
        for target, aliases in mapping.items():
            for alias in aliases:
                if alias.upper() in orig_cols:
                    rename_dict[alias.upper()] = target
                    break
        
        df = df.rename(columns=rename_dict)

        if "DATA_REG" not in df.columns:
            # Tenta encontrar qualquer coluna que tenha 'DATA' no nome se não achou a exata
            for c in df.columns:
                if "DATA" in c:
                    df = df.rename(columns={c: "DATA_REG"})
                    break

        df = df.dropna(subset=["DATA_REG"])
        df["DATA_REG"] = pd.to_datetime(df["DATA_REG"], errors="coerce")
        df = df.dropna(subset=["DATA_REG"])

        df_dia = df[df["DATA_REG"].dt.date == data_alvo_date].copy()

        if df_dia.empty:
            return pd.DataFrame()

        def extrair_bloco(row):
            b = row.get("BLOCO_RAW")
            if pd.notna(b) and str(b).strip() not in ["", "nan", "None"]:
                return str(b).strip().split(".")[0].upper()
            
            mat_bloco = str(row.get("MAT_BLOCO", "")).strip()
            if "-" in mat_bloco:
                return mat_bloco.rsplit("-", 1)[-1].strip().split(".")[0].upper()
            return ""

        df_dia["BLOCO"] = df_dia.apply(extrair_bloco, axis=1)

        mapa = get_mapa_resumido_processos()
        def resumir(proc):
            proc_up = str(proc).strip().upper()
            return mapa.get(proc_up, proc_up)

        df_dia["RESUMIDO"] = df_dia["PROCESSO_APONTADO"].apply(resumir)
        df_dia["QTD_CH"] = pd.to_numeric(df_dia["QTD_CH"], errors="coerce").fillna(0)
        df_dia["QTD_M2"] = pd.to_numeric(df_dia.get("QTD_M2", 0), errors="coerce").fillna(0)

        colunas = ["BLOCO", "NOME_MATERIAL", "MAT_BLOCO", "PROCESSO_APONTADO", "RESUMIDO", "SETOR_AP", "QTD_CH", "QTD_M2", "DATA_REG"]
        return df_dia[[c for c in colunas if c in df_dia.columns]].reset_index(drop=True)
    except Exception as e:
        print(f"Erro ao ler apontamentos do dia: {e}")
        return pd.DataFrame()
