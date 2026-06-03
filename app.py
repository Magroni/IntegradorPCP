import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime, timedelta
import importlib
import altair as alt
import data_manager as dm

importlib.reload(dm)

# --- FUNÇÕES DE SUPORTE GLOBAIS (DENTRO DO APP) ---
def parse_dt(d):
    if pd.isna(d) or str(d).strip().upper() in ["", "NAN", "NAT", "NONE", "NULL", "-", "0"]: return None
    try:
        if isinstance(d, pd.Timestamp): return d.date()
        d_str = str(d).strip()
        # Tenta limpar nomes de dias da semana (comum em formatos longos do Excel)
        if "," in d_str: d_str = d_str.split(",")[-1].strip()
        
        # Tenta formatos comuns
        for fmt in ["%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%y"]:
            try: return pd.to_datetime(d_str, format=fmt).date()
            except: continue
        return pd.to_datetime(d_str).date()
    except: return None

def is_finished(row):
    st_proc = str(row.get("STATUS PROCESSO", "")).upper()
    dt_real = parse_dt(row.get("DATA REALIZADA"))
    return st_proc == "REALIZADO" or dt_real is not None

def is_planned(row):
    return parse_dt(row.get("DATA")) is not None

def format_data_view(d_val):
    if pd.isna(d_val) or str(d_val).strip() in ["", "nan", "NaT", "None"]: return "-"
    dt = parse_dt(d_val)
    if dt: return dt.strftime("%d/%m/%Y")
    return str(d_val)

def get_edited_df(df_base, state_key):
    """
    Reconstrói de forma robusta o DataFrame de um st.data_editor a partir dos estados
    brutos de alteração da Session State (edited_rows, added_rows, deleted_rows).
    Isso é necessário para contornar o bug crônico do Streamlit onde st.data_editor
    dentro de st.form retorna o DataFrame antigo sem as edições na submissão do formulário.
    """
    if state_key not in st.session_state:
        return df_base
        
    editor_state = st.session_state[state_key]
    df_edited = df_base.copy()
    
    # 1. Aplica edições em linhas existentes
    edited_rows = editor_state.get("edited_rows", {})
    for idx_str, row_changes in edited_rows.items():
        try:
            idx = int(idx_str)
            if idx in df_edited.index:
                for col, val in row_changes.items():
                    df_edited.at[idx, col] = val
        except Exception as e:
            print(f"Erro ao aplicar edição no índice {idx_str}: {e}")
            
    # 2. Adiciona novas linhas
    added_rows = editor_state.get("added_rows", [])
    if added_rows:
        new_rows = []
        for row in added_rows:
            new_rows.append(row)
        if new_rows:
            new_df = pd.DataFrame(new_rows)
            df_edited = pd.concat([df_edited, new_df], ignore_index=True)
            
    # 3. Deleta linhas removidas
    deleted_rows = editor_state.get("deleted_rows", [])
    if deleted_rows:
        try:
            indices_to_drop = [int(idx) for idx in deleted_rows]
            df_edited = df_edited.drop(index=[idx for idx in indices_to_drop if idx in df_edited.index])
        except Exception as e:
            print(f"Erro ao deletar linhas no editor: {e}")
            
    return df_edited.reset_index(drop=True)

st.set_page_config(page_title="Apontamento & Indicadores — Costa Granitos", layout="wide", page_icon="📊")

st.title("📊 Gestor de Apontamentos & Indicadores de Produção")

# Custom CSS to improve navigation
st.markdown("""
    <style>
    /* Esconde ABSOLUTAMENTE todos os botões internos de campos numéricos (+, -, x) */
    [data-testid="stNumberInputContainer"] button {
        display: none !important;
    }
    /* Garante que o input ocupe todo o espaço e não pare o TAB nos botões escondidos */
    [data-testid="stNumberInputContainer"] input {
        padding-right: 1rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE SELEÇÃO E CONFIGURAÇÃO ---
def gui_select_file(current_path, is_excel=True):
    import tkinter as tk
    from tkinter import filedialog
    import os
    try:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1)
        init_dir = os.path.dirname(current_path) if current_path and os.path.exists(current_path) else "."
        types = [("Excel files", "*.xlsx *.xlsm *.xls *.xlsb"), ("All files", "*.*")] if is_excel else [("All files", "*.*")]
        
        file_path = filedialog.askopenfilename(
            initialdir=init_dir,
            title="Selecione o Arquivo",
            filetypes=types
        )
        root.destroy()
        if file_path:
            return file_path.replace("/", "\\")
        return None
    except Exception as e:
        print(f"Erro no seletor: {e}")
        return None

def render_opcoes_gerais(cfg_atual, df_base):
    import os as _osc
    st.header("⚙️ Opções Gerais")
    st.subheader("📁 Caminhos e Abas das Bases de Dados")
    st.markdown("Configure os arquivos e **qual aba** de cada planilha o sistema deve usar. Após salvar, atualize a página.")

    # --- Arquivo de Programação ---
    st.markdown("#### 📊 Arquivo de Programação (`.xlsm`)")
    
    # Inicializa session_state para manter valor digitado ou selecionado
    if "tmp_db_path" not in st.session_state:
        st.session_state["tmp_db_path"] = cfg_atual.get("DB_FILE", "")

    col_p1, col_p2, col_p3 = st.columns([3, 1, 1])
    with col_p1:
        novo_db = st.text_input(
            "Caminho completo do Banco de Dados",
            value=st.session_state["tmp_db_path"],
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\DB.xlsm"
        )
        if novo_db != st.session_state["tmp_db_path"]:
            st.session_state["tmp_db_path"] = novo_db
            st.rerun()
    with col_p2:
        if st.button("🔍 Buscar...", key="btn_busc_db"):
            selecionado = gui_select_file(st.session_state["tmp_db_path"])
            if selecionado:
                st.session_state["tmp_db_path"] = selecionado
                st.rerun()
    with col_p3:
        existe_db_live = _osc.path.exists(st.session_state["tmp_db_path"])
        st.markdown(f"<div style='margin-top: 5px;'>{'🟢 OK' if existe_db_live else '🔴 Não encontrado'}</div>", unsafe_allow_html=True)

    if existe_db_live:
        abas_db = dm.get_sheet_names(st.session_state["tmp_db_path"])
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            idx_prog = abas_db.index(cfg_atual.get("SHEET_PROGRAMACAO", "")) if cfg_atual.get("SHEET_PROGRAMACAO","") in abas_db else 0
            aba_prog = st.selectbox("Aba: Programação", abas_db, index=idx_prog, key="cfg_aba_prog")
        with col_s2:
            idx_entregues = abas_db.index(cfg_atual.get("SHEET_ENTREGUES", "")) if cfg_atual.get("SHEET_ENTREGUES","") in abas_db else 0
            aba_entregues = st.selectbox("Aba: Entregues", abas_db, index=idx_entregues, key="cfg_aba_entregues")
        with col_s3:
            idx_base = abas_db.index(cfg_atual.get("SHEET_BASE_DADOS", "")) if cfg_atual.get("SHEET_BASE_DADOS","") in abas_db else 0
            aba_base = st.selectbox("Aba: Base de Dados", abas_db, index=idx_base, key="cfg_aba_base")
    else:
        aba_prog = cfg_atual.get("SHEET_PROGRAMACAO", "DB")
        aba_entregues = cfg_atual.get("SHEET_ENTREGUES", "ENTREGUES")
        aba_base = cfg_atual.get("SHEET_BASE_DADOS", "BASE DE DADOS")
        if novo_db:
            st.caption("Informe um caminho válido para ver as abas disponíveis.")

    st.markdown("---")

    # --- Arquivo de Apontamento ---
    st.markdown("#### ✅ Arquivo de Apontamento (`.xlsx`)")
    
    if "tmp_ap_path" not in st.session_state:
        st.session_state["tmp_ap_path"] = cfg_atual.get("APONTAMENTO_FILE", "")

    col_a1, col_a2, col_a3 = st.columns([3, 1, 1])
    with col_a1:
        novo_ap = st.text_input(
            "Caminho completo do Apontamento",
            value=st.session_state["tmp_ap_path"],
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\DB.xlsx"
        )
        if novo_ap != st.session_state["tmp_ap_path"]:
            st.session_state["tmp_ap_path"] = novo_ap
            st.rerun()
    with col_a2:
        if st.button("🔍 Buscar...", key="btn_busc_ap"):
            selecionado = gui_select_file(st.session_state["tmp_ap_path"])
            if selecionado:
                st.session_state["tmp_ap_path"] = selecionado
                st.rerun()
    with col_a3:
        existe_ap_live = _osc.path.exists(st.session_state["tmp_ap_path"])
        st.markdown(f"<div style='margin-top: 5px;'>{'🟢 OK' if existe_ap_live else '🔴 Não encontrado'}</div>", unsafe_allow_html=True)

    if existe_ap_live:
        abas_ap = dm.get_sheet_names(st.session_state["tmp_ap_path"])
        col_sa1, col_sa2 = st.columns(2)
        with col_sa1:
            idx_ap_bd = abas_ap.index(cfg_atual.get("SHEET_AP_BD", "")) if cfg_atual.get("SHEET_AP_BD","") in abas_ap else 0
            aba_ap_bd = st.selectbox("Aba: Apontamentos Diários", abas_ap, index=idx_ap_bd, key="cfg_aba_ap_bd")
        with col_sa2:
            idx_ap_base = abas_ap.index(cfg_atual.get("SHEET_AP_BASE", "")) if cfg_atual.get("SHEET_AP_BASE","") in abas_ap else 0
            aba_ap_base = st.selectbox("Aba: Mapeamento de Processos", abas_ap, index=idx_ap_base, key="cfg_aba_ap_base")
    else:
        aba_ap_bd = cfg_atual.get("SHEET_AP_BD", "DB")
        aba_ap_base = cfg_atual.get("SHEET_AP_BASE", "BASE DADOS")
        if novo_ap:
            st.caption("Informe um caminho válido para ver as abas disponíveis.")

    st.markdown("---")

    # --- Planilha de Blocos ---
    st.markdown("#### 🧱 Planilha de Blocos (`.xlsb`)")
    
    if "tmp_bl_path" not in st.session_state:
        st.session_state["tmp_bl_path"] = cfg_atual.get("BLOCKS_FILE", "")

    col_bl1, col_bl2, col_bl3 = st.columns([3, 1, 1])
    with col_bl1:
        novo_bl = st.text_input(
            "Caminho completo de Blocos",
            value=st.session_state["tmp_bl_path"],
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\PLANILHA BLOCOS.xlsb"
        )
        if novo_bl != st.session_state["tmp_bl_path"]:
            st.session_state["tmp_bl_path"] = novo_bl
            st.rerun()
    with col_bl2:
        if st.button("🔍 Buscar...", key="btn_busc_bl"):
            selecionado = gui_select_file(st.session_state["tmp_bl_path"])
            if selecionado:
                st.session_state["tmp_bl_path"] = selecionado
                st.rerun()
    with col_bl3:
        existe_bl_live = _osc.path.exists(st.session_state["tmp_bl_path"])
        st.markdown(f"<div style='margin-top: 5px;'>{'🟢 OK' if existe_bl_live else '🔴 Não encontrado'}</div>", unsafe_allow_html=True)

    st.markdown("---")

    # --- Planilha de Chapas ---
    st.markdown("#### 📋 Planilha de Estoque de Chapas (`.xlsx`)")
    
    if "tmp_ch_path" not in st.session_state:
        st.session_state["tmp_ch_path"] = cfg_atual.get("CHAPAS_FILE", "")

    col_ch1, col_ch2, col_ch3 = st.columns([3, 1, 1])
    with col_ch1:
        novo_ch = st.text_input(
            "Caminho completo de Chapas",
            value=st.session_state["tmp_ch_path"],
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\Estoque Chapas.xlsx"
        )
        if novo_ch != st.session_state["tmp_ch_path"]:
            st.session_state["tmp_ch_path"] = novo_ch
            st.rerun()
    with col_ch2:
        if st.button("🔍 Buscar...", key="btn_busc_ch"):
            selecionado = gui_select_file(st.session_state["tmp_ch_path"])
            if selecionado:
                st.session_state["tmp_ch_path"] = selecionado
                st.rerun()
    with col_ch3:
        existe_ch_live = _osc.path.exists(st.session_state["tmp_ch_path"])
        st.markdown(f"<div style='margin-top: 5px;'>{'🟢 OK' if existe_ch_live else '🔴 Não encontrado'}</div>", unsafe_allow_html=True)

    st.markdown("---")

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 Salvar Configurações", type="primary", width="stretch", key="btn_salvar_cfg"):
            alertas_path = []
            if not _osc.path.exists(novo_db):
                alertas_path.append(f"⚠️ Arquivo de Programação não encontrado localmente: `{novo_db}`")
            if not _osc.path.exists(novo_ap):
                alertas_path.append(f"⚠️ Arquivo de Apontamento não encontrado localmente: `{novo_ap}`")
            if not _osc.path.exists(novo_bl):
                alertas_path.append(f"⚠️ Planilha de Blocos não encontrada localmente: `{novo_bl}`")
            if not _osc.path.exists(novo_ch):
                alertas_path.append(f"⚠️ Planilha de Chapas não encontrada localmente: `{novo_ch}`")

            # Sempre salvamos, mas emitimos avisos caso algum arquivo não tenha sido encontrado
            novo_cfg = dict(cfg_atual)
            novo_cfg["DB_FILE"] = novo_db
            novo_cfg["APONTAMENTO_FILE"] = novo_ap
            novo_cfg["BLOCKS_FILE"] = novo_bl
            novo_cfg["CHAPAS_FILE"] = novo_ch
            novo_cfg["SHEET_PROGRAMACAO"] = aba_prog
            novo_cfg["SHEET_ENTREGUES"] = aba_entregues
            novo_cfg["SHEET_BASE_DADOS"] = aba_base
            novo_cfg["SHEET_AP_BD"] = aba_ap_bd
            novo_cfg["SHEET_AP_BASE"] = aba_ap_base
            
            if dm.save_config(novo_cfg):
                # Limpa as variáveis temporárias de sessão para que o st.rerun() recarregue os novos caminhos salvos
                for key in ["tmp_db_path", "tmp_ap_path", "tmp_bl_path", "tmp_ch_path",
                            "cfg_db_path_input", "cfg_ap_path_input", "cfg_bl_path_input", "cfg_ch_path_input"]:
                    if key in st.session_state:
                        del st.session_state[key]
                
                if alertas_path:
                    st.success("✅ Configurações salvas!")
                    for ap in alertas_path:
                        st.warning(ap)
                    st.info("💡 As configurações foram salvas com sucesso. No entanto, verifique se os caminhos acima estão corretos ou se a máquina atual possui permissão de acesso a eles.")
                else:
                    st.success("✅ Configurações salvas com sucesso! Aplicando...")
                
                st.rerun()
            else:
                st.error("Erro ao salvar o arquivo de configuração.")

    with col_btn2:
        if st.button("↩️ Restaurar Padrões", width="stretch", key="btn_restaurar_cfg"):
            if dm.save_config(dm._DEFAULT_CONFIG):
                # Limpa as variáveis temporárias de sessão para que o st.rerun() recarregue os padrões salvos
                for key in ["tmp_db_path", "tmp_ap_path", "tmp_bl_path", "tmp_ch_path",
                            "cfg_db_path_input", "cfg_ap_path_input", "cfg_bl_path_input", "cfg_ch_path_input"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.success("↩️ Configurações restauradas para os padrões com sucesso!")
                st.rerun()
            else:
                st.error("Erro ao restaurar configuração.")

    st.divider()

    # ---- SEÇÃO 2: PROCESSOS E SETORES ----
    st.subheader("🔧 Roteiro Padrão: Processos × Máquinas")
    st.markdown("Gerencie a correspondência padrão entre Processos e Setores Produtivos (Máquinas). Estes dados alimentam a inteligência do sistema.")
    
    if df_base.empty:
        st.warning("Não há dados na aba BASE DE DADOS.")
        df_edit = pd.DataFrame(columns=["PROCESSO", "SETOR"])
    else:
        df_edit = df_base[["PROCESSO", "SETOR"]].copy()
        
    st.markdown("Edite os nomes abaixo, adicione novas linhas clicando na última célula ou exclua linhas selecionando a borda esquerda e apertando *Delete*.")
    editado = st.data_editor(
        df_edit,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "PROCESSO": st.column_config.TextColumn("Nome do Processo", required=True),
            "SETOR": st.column_config.TextColumn("Setor Produtivo (Máquina) Padrão", required=True)
        }
    )
    
    if st.button("💾 Salvar Configurações no Banco de Dados", type="primary"):
        with st.spinner("Salvando na aba BASE DE DADOS..."):
            editado = editado.dropna(subset=["PROCESSO", "SETOR"], how="all")
            sucesso = dm.update_base_dados(editado)
            if sucesso:
                st.success("Configurações salvas com sucesso! As novas sugestões já estão valendo para a aba de blocos.")
            else:
                st.error("Ocorreu um erro ao salvar as configurações.")

    st.divider()
    
    # ---- SEÇÃO 3: TIPO DE PROCESSO × SETORES PERMITIDOS ----
    st.subheader("📋 Filtro por Tipo de Produção: Tipos × Setores Permitidos")
    st.markdown("Configure quais Máquinas/Setores devem ser exibidos no apontamento de produção para cada Tipo de Processo.")
    
    df_ts_base = dm.get_tipo_setores()
    if df_ts_base.empty:
        df_ts_edit = pd.DataFrame(columns=["TIPO_PROCESSO", "SETORES"])
    else:
        df_ts_edit = df_ts_base[["TIPO_PROCESSO", "SETORES"]].copy()
        
    st.markdown("Edite abaixo. No campo **Setores Permitidos**, insira os nomes dos setores separados por vírgula (ex: `CIMEF, BARSANTI`).")
    editado_ts = st.data_editor(
        df_ts_edit,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "TIPO_PROCESSO": st.column_config.TextColumn("Tipo de Processo", required=True),
            "SETORES": st.column_config.TextColumn("Setores Permitidos (separados por vírgula)", required=True)
        },
        key="editor_tipo_setores_cfg"
    )
    
    if st.button("💾 Salvar Filtros de Produção no Banco de Dados", type="primary", key="btn_salvar_ts"):
        with st.spinner("Salvando na aba TIPO_SETORES..."):
            editado_ts = editado_ts.dropna(subset=["TIPO_PROCESSO", "SETORES"], how="all")
            sucesso = dm.update_tipo_setores(editado_ts)
            if sucesso:
                st.success("Filtros de produção salvos com sucesso! As novas opções de filtros já estão ativas no apontamento.")
                st.rerun()
            else:
                st.error("Ocorreu um erro ao salvar as configurações dos filtros.")

    st.divider()
    
    # ---- SEÇÃO 4: CADASTRO E FAROL DE MOTIVOS DE PARADA ----
    st.subheader("🛑 Cadastro e Farol Lean de Motivos de Parada")
    st.markdown("Crie e edite a lista oficial de motivos de paradas que os operadores poderão selecionar no apontamento. Defina o **Farol Lean** correspondente para cada parada (`Operacional`, `Intervenção`, ou `Crítica`).")
    
    df_tp_base = dm.get_tipo_paradas()
    if df_tp_base.empty:
        df_tp_edit = pd.DataFrame(columns=["MOTIVO", "TIPO_PARADA"])
    else:
        df_tp_edit = df_tp_base[["MOTIVO", "TIPO_PARADA"]].copy()
        
    st.markdown("Edite a lista abaixo. O campo **Farol Lean (Tipo)** aceita as opções: `Operacional`, `Intervenção` ou `Crítica`.")
    editado_tp = st.data_editor(
        df_tp_edit,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "MOTIVO": st.column_config.TextColumn("Motivo da Parada Padronizado", required=True),
            "TIPO_PARADA": st.column_config.SelectboxColumn("Farol Lean (Tipo)", options=["Operacional", "Intervenção", "Crítica"], required=True)
        },
        key="editor_tipo_paradas_cfg"
    )
    
    if st.button("💾 Salvar Cadastro de Paradas no Banco de Dados", type="primary", key="btn_salvar_tp"):
        with st.spinner("Salvando na aba TIPO_PARADAS..."):
            editado_tp = editado_tp.dropna(subset=["MOTIVO", "TIPO_PARADA"], how="all")
            sucesso = dm.update_tipo_paradas(editado_tp)
            if sucesso:
                st.success("Cadastro e Farol Lean de Paradas salvos com sucesso! A lista suspensa já está disponível no apontamento.")
                st.rerun()
            else:
                st.error("Ocorreu um erro ao salvar as configurações das paradas.")


# Load base data and pointing data
df_base = dm.get_base_dados()
df_all_ap = dm.get_all_apontamentos()

if df_all_ap.empty:
    aba_ap = dm._get_sheet("SHEET_AP_BD")
    st.error(f"❌ Não foi possível carregar os dados de apontamentos. O arquivo de apontamento pode estar temporariamente inacessível ou o nome da aba '{aba_ap}' está incorreto.")
    st.info("💡 Ajuste os caminhos e as configurações das bases de dados abaixo para apontar para arquivos corretos:")
    
    # Renderiza a interface de configurações gerais de emergência para que o usuário possa consertar
    render_opcoes_gerais(dm.get_config(), df_base)
    st.stop()

if not df_base.empty:
    df_base = df_base.fillna("")

mapa_processos = {}
lista_processos = []
if not df_base.empty and "PROCESSO" in df_base.columns and "SETOR" in df_base.columns:
    for idx, row in df_base.iterrows():
        p = str(row.get("PROCESSO", "")).strip()
        s = str(row.get("SETOR", "")).strip()
        if p and p != "nan":
            mapa_processos[p] = s
            lista_processos.append(p)

setores_list = sorted([str(x) for x in df_base["SETOR"].unique() if str(x) != "" and str(x) != "nan"])

tab_apontamento, tab_consulta, tab_analises, tab_config = st.tabs([
    "✅ Lançar Apontamento",
    "🔍 Consulta de Apontamentos",
    "📈 Análises e Indicadores",
    "⚙️ Opções Gerais"
])


# ----------------- ABA 1: ADICIONAR / EDITAR BLOCO -----------------
# ----------------- ABA 4: APONTAMENTO DE PRODUÇÃO -----------------
with tab_apontamento:
    st.header("✅ Apontamento de Produção")
    st.markdown("Registre os apontamentos de produção da fábrica diretamente no Excel, incluindo tempos, paradas de máquina e insumos.")

    # ---- NOVO FORMULÁRIO DE APONTAMENTO MANUAL DINÂMICO ----
    with st.expander("➕ Lançar Novo Apontamento Manual", expanded=True):
        st.markdown("### Registro de Produção")
        
        # 1. BUSCA DE BLOCO (Sempre disponível no topo)
        c_bl1, c_bl2 = st.columns([1, 3])
        with c_bl1:
            f_bloco = st.text_input("Nº Bloco*", value=st.session_state.get("ap_bloco_val", ""), placeholder="Ex: 1234", key="ap_bloco_trigger_v2")
        
        if f_bloco != st.session_state.get("ap_last_bloco"):
            if f_bloco:
                with st.spinner("Buscando informações do bloco..."):
                    info = dm.get_bloco_info(f_bloco)
                    if info:
                        st.session_state["ap_mat_val"] = info["MATERIAL"]
                        st.session_state["ap_comp_val"] = info["COMP"]
                        st.session_state["ap_alt_val"] = info["ALT"]
                        st.session_state["ap_source"] = info["SOURCE"]
                        st.session_state["ap_found"] = True
                    else:
                        st.session_state["ap_mat_val"] = ""
                        st.session_state["ap_comp_val"] = 0.0
                        st.session_state["ap_alt_val"] = 0.0
                        st.session_state["ap_source"] = ""
                        st.session_state["ap_found"] = False
                st.session_state["ap_last_bloco"] = f_bloco
                st.rerun()
            else:
                st.session_state["ap_mat_val"] = ""
                st.session_state["ap_comp_val"] = 0.0
                st.session_state["ap_alt_val"] = 0.0
                st.session_state["ap_last_bloco"] = ""
                st.session_state["ap_source"] = ""
                st.session_state["ap_found"] = None
                st.rerun()

        # Exibição do Status da Busca e Histórico
        if f_bloco:
            if st.session_state.get("ap_found"):
                st.success(f"✅ Bloco encontrado na base: **{st.session_state.get('ap_source')}**")
            elif st.session_state.get("ap_found") is False:
                st.warning("⚠️ Bloco não encontrado nas bases de dados (Blocos/Chapas). Preencha manualmente.")
            
            with st.expander(f"📜 Ver histórico de apontamentos do bloco {f_bloco}", expanded=False):
                with st.spinner("Carregando histórico..."):
                    historico = dm.get_apontamentos_por_bloco(f_bloco)
                    if not historico.empty:
                        st.dataframe(historico, width="stretch", hide_index=True)
                    else:
                        st.info("Nenhum apontamento anterior encontrado para este bloco.")

        # 2. Escolha do TIPO e ETAPA (Aparecem após o bloco ser digitado)
        if f_bloco:
            # Carrega dinamicamente os tipos de processos da aba de configuração
            df_ts = dm.get_tipo_setores()
            tipos_processo = [""] + sorted(list(df_ts["TIPO_PROCESSO"].dropna().unique()))
            f_tipo = st.selectbox("1. Tipo de Processo*", tipos_processo, key="ap_tipo_proc_v3")
            
            if f_tipo:
                # Determina os processos permitidos com base nos setores permitidos para este tipo
                processos_permitidos = []
                match_row = df_ts[df_ts["TIPO_PROCESSO"] == f_tipo]
                if not match_row.empty:
                    setores_str = str(match_row.iloc[0].get("SETORES", ""))
                    setores_permitidos = [s.strip().upper() for s in setores_str.split(",") if s.strip()]
                    
                    # Filtra do mapa_processos apenas os processos cujo setor associado está na lista permitida
                    for proc, setor in mapa_processos.items():
                        if str(setor).strip().upper() in setores_permitidos:
                            processos_permitidos.append(proc)
                            
                # Fallback se não houver nenhum mapeamento resolvido
                if not processos_permitidos:
                    opcoes_proc_master = sorted(list(mapa_processos.keys()))
                else:
                    opcoes_proc_master = sorted(list(set(processos_permitidos)))
                    
                f_processo = st.selectbox("2. Processo/Etapa*", [""] + opcoes_proc_master, key="f_proc_master_v5")
                
                if f_processo:
                    # 1. SEÇÃO LIVE: RESINAGEM (Manual e Dinâmica via Form para evitar lag)
                    # Fora do form principal pois Streamlit não aceita forms aninhados
                    f_qtd_resina, f_tipo_resina, f_tipo_endur, f_qtd_endur_calc, f_sec, f_prop = 0.0, "", "", 0.0, 24, 0.0
                    if f_tipo == "Resinagem / Tela / Manta / Estuque":
                        with st.form("form_resinagem_live"):
                            st.markdown("#### 🧪 Insumos Principais (Cálculo ao pressionar ENTER)")
                            c_m1, c_m2, c_m3, c_m4 = st.columns(4)
                            with c_m1:
                                f_tipo_resina = st.text_input("Nome da Resina", value="", key="live_res_nome")
                                f_qtd_resina = st.number_input("Qtd KG (Resina)", min_value=0.0, step=0.1, value=None, key="live_res_kg")
                            with c_m2:
                                f_tipo_endur = st.text_input("Nome do Catalisador", value="", key="live_cat_nome")
                                f_prop = st.number_input("Proporção (%)", min_value=0.0, max_value=100.0, step=0.1, value=None, key="live_cat_prop")
                            
                            if f_prop and f_prop > 0:
                                f_qtd_endur_calc = round((f_qtd_resina if f_qtd_resina else 0.0) * (f_prop / 100), 3)
                                with c_m3: st.number_input("Qtd KG (Catalisador)", value=f_qtd_endur_calc, disabled=True)
                                with c_m4: f_sec = st.number_input("Tempo Secagem (h)", min_value=0, value=24)
                            
                            st.form_submit_button("🔄 RECALCULAR / CONFIRMAR PROPORÇÃO", width="stretch")
                            if not f_prop or f_prop <= 0:
                                st.info("ℹ️ Digite a proporção e pressione ENTER ou o botão acima para calcular.")
                        st.markdown("---")

                    # 2. O RESTANTE NO FORM (Para evitar atualizações constantes ao usar TAB)
                    with st.form("form_apontamento_completo"):
                        st.markdown("#### 📄 Dados de Produção e Insumos")
                        
                        c_p1, c_p2, c_p3, c_p4 = st.columns(4)
                        with c_p1:
                            f_data = st.date_input("Data Reg.*", value=datetime.now(), format="DD/MM/YYYY")
                            f_material = st.text_input("Material*", value=st.session_state.get("ap_mat_val", ""), placeholder="Ex: ALPINUS")
                        with c_p2:
                            # Filtra as máquinas permitidas para o tipo de processo selecionado
                            setores_disponiveis = []
                            if f_tipo:
                                match_row = df_ts[df_ts["TIPO_PROCESSO"] == f_tipo]
                                if not match_row.empty:
                                    setores_str = str(match_row.iloc[0].get("SETORES", ""))
                                    setores_disponiveis = sorted([s.strip().upper() for s in setores_str.split(",") if s.strip()])
                            
                            # Fallback caso não esteja cadastrado ou esteja em branco no mapeamento
                            if not setores_disponiveis:
                                setores_ativos = set(str(x) for x in df["SETOR"].unique() if str(x) not in ["", "nan"])
                                setores_mapa = set(v for v in mapa_processos.values() if v)
                                setores_disponiveis = sorted(list(setores_ativos | setores_mapa))
                                
                            f_setor = st.selectbox("Máquina/Setor*", [""] + setores_disponiveis)
                            f_operador = st.text_input("Operador")
                        with c_p3:
                            f_qtd_ch = st.number_input("Qtd Chapas*", min_value=0, step=1, value=None)
                            f_esp = st.number_input("Espessura", min_value=0.0, step=0.1, value=None)
                        with c_p4:
                            f_comp = st.number_input("Comprimento (m)", min_value=0.0, step=0.01, value=float(st.session_state.get("ap_comp_val", 0.0)))
                            f_alt = st.number_input("Altura (m)", min_value=0.0, step=0.01, value=float(st.session_state.get("ap_alt_val", 0.0)))

                        c_t1, c_t2, c_t3, c_t4 = st.columns(4)
                        with c_t1: f_dia_ini = st.date_input("Dia Início", value=datetime.now() - timedelta(days=1), format="DD/MM/YYYY")
                        with c_t2: f_dia_fim = st.date_input("Dia Fim", value=datetime.now() - timedelta(days=1), format="DD/MM/YYYY")
                        with c_t3: f_hora_ini_str = st.text_input("Hora Início", value="", placeholder="Ex: 0800")
                        with c_t4: f_hora_fim_str = st.text_input("Hora Fim", value="", placeholder="Ex: 1530")

                        if f_tipo == "Levigamento / Polimento":
                            st.markdown("---")
                            st.caption("⚙️ Configuração da Máquina (C01 a C20)")
                            if "polishing_heads_state" not in st.session_state: st.session_state["polishing_heads_state"] = [""] * 20
                            for r in range(2):
                                cols_c = st.columns(10)
                                for c in range(10):
                                    idx = r * 10 + c
                                    st.session_state["polishing_heads_state"][idx] = cols_c[c].text_input(f"C{idx+1:02d}", value=st.session_state["polishing_heads_state"][idx], key=f"f_cab_form_{idx+1}")

                        st.markdown("---")
                        c_i1, c_i2 = st.columns([3, 2])
                        with c_i1:
                            st.caption("➕ Insumos Adicionais")
                            if "df_ins_add" not in st.session_state or st.session_state.get("last_proc_v3") != f_processo:
                                st.session_state["last_proc_v3"] = f_processo
                                if f_tipo == "Resinagem / Tela / Manta / Estuque": rows_add = [{"TIPO": "MANTA", "DESCRICAO": "", "QTD": 0.0, "UNID": "M²"}]
                                elif f_tipo == "Levigamento / Polimento": rows_add = [{"TIPO": "ABRASIVO", "DESCRICAO": "", "QTD": 0.0, "UNID": "UNID"}]
                                else: rows_add = [{"TIPO": "OUTROS", "DESCRICAO": "", "QTD": 0.0, "UNID": "UNID"}]
                                st.session_state["df_ins_add"] = pd.DataFrame(rows_add)
                            editado_ins_add = st.data_editor(st.session_state["df_ins_add"], num_rows="dynamic", width="stretch", key="editor_ins_final")
                        
                        with c_i2:
                            st.caption("⏹️ Paradas de Máquina")
                            
                            # Carrega os motivos oficiais cadastrados
                            df_tp_op = dm.get_tipo_paradas()
                            lista_motivos_ap = sorted(list(df_tp_op["MOTIVO"].dropna().unique())) if not df_tp_op.empty else []
                            
                            # Sincroniza a data das paradas com a data do processo
                            if "df_paradas_state" not in st.session_state or st.session_state.get("last_f_dia_ini") != f_dia_ini or st.session_state.get("last_f_dia_fim") != f_dia_fim:
                                st.session_state["last_f_dia_ini"] = f_dia_ini
                                st.session_state["last_f_dia_fim"] = f_dia_fim
                                is_empty_or_default = True
                                if "df_paradas_state" in st.session_state:
                                    df_p = st.session_state["df_paradas_state"]
                                    if not df_p.empty and (len(df_p) > 1 or df_p.iloc[0]["MOTIVO"] != ""):
                                        is_empty_or_default = False
                                
                                if is_empty_or_default:
                                    st.session_state["df_paradas_state"] = pd.DataFrame([{"MOTIVO": "", "DIA_INI": f_dia_ini, "HORA_INI": "", "DIA_FIM": f_dia_fim, "HORA_FIM": ""}])
                                else:
                                    df_p = st.session_state["df_paradas_state"]
                                    if df_p.iloc[0]["MOTIVO"] == "":
                                        df_p.at[df_p.index[0], "DIA_INI"] = f_dia_ini
                                        df_p.at[df_p.index[0], "DIA_FIM"] = f_dia_fim
                                    st.session_state["df_paradas_state"] = df_p
                            
                            col_conf_paradas = {
                                "DIA_INI": st.column_config.DateColumn("D.Início", format="DD/MM/YYYY", default=f_dia_ini),
                                "DIA_FIM": st.column_config.DateColumn("D.Fim", format="DD/MM/YYYY", default=f_dia_fim),
                                "HORA_INI": st.column_config.TextColumn("H.Início"),
                                "HORA_FIM": st.column_config.TextColumn("H.Fim")
                            }
                            if lista_motivos_ap:
                                col_conf_paradas["MOTIVO"] = st.column_config.SelectboxColumn("Motivo da Parada*", options=lista_motivos_ap, required=True)
                            else:
                                col_conf_paradas["MOTIVO"] = st.column_config.TextColumn("Motivo da Parada*", required=True)
                                
                            editado_paradas = st.data_editor(
                                st.session_state["df_paradas_state"], 
                                num_rows="dynamic", 
                                width="stretch", 
                                column_config=col_conf_paradas,
                                key="editor_paradas_final"
                            )

                        btn_carrinho = st.form_submit_button("🛒 ADICIONAR AO CARRINHO", width="stretch", type="primary")


                    if btn_carrinho:
                        # Reconstroi os dataframes dos editores a partir do session state para evitar o bug do st.form
                        editado_ins_add = get_edited_df(st.session_state["df_ins_add"], "editor_ins_final")
                        editado_paradas = get_edited_df(st.session_state["df_paradas_state"], "editor_paradas_final")
                        
                        # Validação e lógica de salvamento (mesma lógica do callback anterior)
                        # ... processamento dos dados lidos dos widgets acima ...
                        def parse_time_local(t_str):
                            if not t_str: return None
                            import re
                            from datetime import time
                            t = re.sub(r'\D', '', str(t_str))
                            if len(t) == 4:
                                try: return time(int(t[:2]), int(t[2:]))
                                except: return None
                            elif len(t) == 3:
                                try: return time(int(t[:1]), int(t[1:]))
                                except: return None
                            elif ":" in str(t_str):
                                parts = str(t_str).split(":")
                                if len(parts) == 2:
                                    try: return time(int(parts[0]), int(parts[1]))
                                    except: return None
                            return None

                        h_ini = parse_time_local(f_hora_ini_str)
                        h_fim = parse_time_local(f_hora_fim_str)
                        
                        if not f_material or not f_setor or f_qtd_ch is None or not h_ini or not h_fim:
                            st.error("❌ Preencha todos os campos obrigatórios (*) e as horas corretamente.")
                        else:
                            # Calcula tempo processo
                            dt1 = datetime.combine(f_dia_ini, h_ini)
                            dt2 = datetime.combine(f_dia_fim, h_fim)
                            m_tot = (dt2 - dt1).total_seconds() / 60
                            
                            if m_tot <= 0:
                                st.error("❌ Tempo total deve ser positivo.")
                            else:
                                # Coleta Insumos das Cabeças
                                ins_finais = []
                                if f_tipo == "Levigamento / Polimento":
                                    for i in range(1, 21):
                                        val_c = st.session_state.get(f"f_cab_form_{i}")
                                        if val_c:
                                            ins_finais.append({
                                                "TIPO_INSUMO": "ABRASIVO", 
                                                "DESCRICAO": f"CAB {i:02d}: {val_c}", 
                                                "QUANTIDADE": 1.0, 
                                                "UNIDADE": "UNID", 
                                                "TEMPO_SECAGEM": "",
                                                "CABECAS": f"{i:02d}",
                                                "INSUMO_DETALHE": val_c
                                            })
                                
                                # Coleta Resinagem (Vem do scope 'Live' acima)
                                if f_qtd_resina and f_qtd_resina > 0:
                                    ins_finais.append({
                                        "TIPO_INSUMO": "RESINA", 
                                        "DESCRICAO": f_tipo_resina.upper(), 
                                        "QUANTIDADE": f_qtd_resina, 
                                        "UNIDADE": "KG", 
                                        "TEMPO_SECAGEM": "",
                                        "CABECAS": "",
                                        "INSUMO_DETALHE": f_tipo_resina.upper()
                                    })
                                if f_qtd_endur_calc and f_qtd_endur_calc > 0:
                                    ins_finais.append({
                                        "TIPO_INSUMO": "ENDURECEDOR", 
                                        "DESCRICAO": f_tipo_endur.upper(), 
                                        "QUANTIDADE": f_qtd_endur_calc, 
                                        "UNIDADE": "KG", 
                                        "TEMPO_SECAGEM": f_sec,
                                        "CABECAS": "",
                                        "INSUMO_DETALHE": f_tipo_endur.upper()
                                    })
                                
                                # Insumos da tabela e Paradas
                                for _, row in editado_ins_add.iterrows():
                                    if row.get("TIPO") and row.get("QTD", 0) > 0:
                                        ins_finais.append({
                                            "TIPO_INSUMO": row["TIPO"], 
                                            "DESCRICAO": row["DESCRICAO"], 
                                            "QUANTIDADE": row["QTD"], 
                                            "UNIDADE": row["UNID"], 
                                            "TEMPO_SECAGEM": "",
                                            "CABECAS": "",
                                            "INSUMO_DETALHE": row["DESCRICAO"]
                                        })
                                
                                # 4. Coleta Paradas (Downtimes) e Validação Estrita
                                par_finais = []
                                total_mp = 0
                                erro_parada = None
                                
                                for _, p_row in editado_paradas.iterrows():
                                    if p_row.get("MOTIVO"):
                                        h_ini_p = parse_time_local(p_row.get("HORA_INI"))
                                        h_fim_p = parse_time_local(p_row.get("HORA_FIM"))
                                        
                                        if not h_ini_p or not h_fim_p:
                                            erro_parada = f"❌ Hora inválida ou incompleta na parada '{p_row['MOTIVO']}'."
                                            break
                                            
                                        # Conversão extremamente segura usando parse_dt
                                        d_ini_p_parsed = parse_dt(p_row["DIA_INI"])
                                        d_fim_p_parsed = parse_dt(p_row["DIA_FIM"])
                                        
                                        if not d_ini_p_parsed or not d_fim_p_parsed:
                                            erro_parada = f"❌ Data de início ou fim inválida na parada '{p_row['MOTIVO']}'."
                                            break
                                            
                                        dt_ini_p = datetime.combine(d_ini_p_parsed, h_ini_p)
                                        dt_fim_p = datetime.combine(d_fim_p_parsed, h_fim_p)
                                        mp = (dt_fim_p - dt_ini_p).total_seconds() / 60
                                        
                                        # Validação: Dentro do intervalo do processo?
                                        if dt_ini_p < dt1 or dt_fim_p > dt2:
                                            erro_parada = f"❌ Parada '{p_row['MOTIVO']}' ({h_ini_p.strftime('%H:%M')} às {h_fim_p.strftime('%H:%M')}) está fora do intervalo do processo ({h_ini.strftime('%H:%M')} às {h_fim.strftime('%H:%M')})."
                                            break
                                        
                                        if mp <= 0:
                                            erro_parada = f"❌ Tempo da parada '{p_row['MOTIVO']}' deve ser positivo."
                                            break

                                        total_mp += mp
                                        par_finais.append({
                                            "MOTIVO": p_row["MOTIVO"], 
                                            "DIA_INICIO": d_ini_p_parsed.strftime("%d/%m/%Y"), 
                                            "HORA_INICIO": h_ini_p.strftime("%H:%M"), 
                                            "DIA_FIM": d_fim_p_parsed.strftime("%d/%m/%Y"), 
                                            "HORA_FIM": h_fim_p.strftime("%H:%M"), 
                                            "TEMPO": f"{int(mp // 60):02d}:{int(mp % 60):02d}"
                                        })

                                if erro_parada:
                                    st.error(erro_parada)
                                elif total_mp > m_tot:
                                    st.error(f"❌ Soma das paradas ({int(total_mp)} min) é maior que o tempo total de processo ({int(m_tot)} min).")
                                else:
                                    # ... resto da lógica de sucesso ...
                                    # Calcula Turno (N = 22:00 as 06:59, D = resto)
                                    f_turno = "D"
                                    if h_ini.hour >= 22 or h_ini.hour < 7:
                                        f_turno = "N"

                                    novo_rec = {
                                        "DATA_REG": f_data.strftime("%d/%m/%Y"), "BLOCO_RAW": f_bloco, "NOME_MATERIAL": f_material.upper(),
                                        "PROCESSO_APONTADO": f_processo, "SETOR_AP": f_setor, "QTD_CH": f_qtd_ch, 
                                        "QTD_M2": round(f_comp * f_alt * f_qtd_ch, 3) if f_comp and f_alt else 0.0,
                                        "ESP": f_esp if f_esp else "", "COMP": f_comp if f_comp else "", "ALT": f_alt if f_alt else "",
                                        "OPERADOR": f_operador.upper() if f_operador else "", "DIA_INICIO": f_dia_ini.strftime("%d/%m/%Y"),
                                        "DIA_FIM": f_dia_fim.strftime("%d/%m/%Y"), "HORA_INICIO": h_ini.strftime("%H:%M"),
                                        "HORA_FIM": h_fim.strftime("%H:%M"), "TEMPO_PROCESSO": f"{int(m_tot // 60):02d}:{int(m_tot % 60):02d}", "TURNO": f_turno
                                    }
                                    if "carrinho_ap" not in st.session_state: st.session_state["carrinho_ap"] = []
                                    st.session_state["carrinho_ap"].append((novo_rec, par_finais, ins_finais))
                                    st.toast(f"📍 Bloco {f_bloco} no carrinho!", icon="🛒")
                                    st.rerun()

                # --- EXIBIÇÃO DO CARRINHO (FORA DO FORM) ---
                if "carrinho_ap" in st.session_state and st.session_state["carrinho_ap"]:
                    st.markdown("### 🛒 Carrinho de Apontamentos")
                    st.info(f"Você tem {len(st.session_state['carrinho_ap'])} apontamento(s) no carrinho.")
                    
                    dados_view = []
                    for idx, (rec, par, ins) in enumerate(st.session_state["carrinho_ap"]):
                        dados_view.append({
                            "Nº": idx + 1,
                            "Bloco": rec["BLOCO_RAW"], 
                            "Material": rec["NOME_MATERIAL"], 
                            "Processo": rec["PROCESSO_APONTADO"], 
                            "Chapas": rec["QTD_CH"]
                        })
                    st.table(pd.DataFrame(dados_view))
                    
                    with st.expander("❌ Remover Item Específico do Carrinho"):
                        item_to_remove = st.selectbox(
                            "Selecione o apontamento do carrinho para remover:",
                            range(len(st.session_state["carrinho_ap"])),
                            format_func=lambda i: f"{i+1}. Bloco {st.session_state['carrinho_ap'][i][0]['BLOCO_RAW']} - {st.session_state['carrinho_ap'][i][0]['PROCESSO_APONTADO']}",
                            key="select_item_remove_cart"
                        )
                        if st.button("Confirmar Remoção do Item Selecionado", width="stretch", key="btn_remove_item_cart"):
                            st.session_state["carrinho_ap"].pop(item_to_remove)
                            st.success("Item removido do carrinho com sucesso!")
                            st.rerun()
                    
                    c_b1, c_b2 = st.columns(2)
                    with c_b1:
                        if st.button("🗑️ Limpar Carrinho", width="stretch"):
                            st.session_state["carrinho_ap"] = []
                            st.rerun()
                    with c_b2:
                        if st.button("🚀 FINALIZAR E SALVAR TUDO", type="primary", width="stretch"):
                            with st.spinner("Gravando no Excel..."):
                                next_id = dm.get_next_apontamento_id()
                                batch = []
                                for rec, par, ins in st.session_state["carrinho_ap"]:
                                    rec["ID"] = next_id
                                    batch.append((rec, par, ins))
                                    next_id += 1
                                
                                ok, err = dm.add_apontamento_batch(batch)
                                if ok:
                                    st.success(f"✅ {len(batch)} itens salvos com sucesso!")
                                    st.session_state["carrinho_ap"] = []
                                    st.balloons()
                                    st.rerun()
                                else:
                                    st.error(f"❌ Erro ao salvar: {err}")
                else: st.info("Selecione um processo acima para abrir o formulário.")
            else: st.info("Selecione um Tipo de Processo acima.")

    st.divider()


# ----------------- ABA 5: CONSULTA DE APONTAMENTOS -----------------
with tab_consulta:
    st.header("🔍 Consulta de Apontamentos Realizados")
    st.markdown("Busque e visualize a ficha consolidada detalhada de qualquer apontamento de produção, incluindo processos, paradas de máquina e insumos utilizados.")

    # 1. Carrega todos os apontamentos
    with st.spinner("Carregando apontamentos do banco de dados..."):
        df_all_ap = dm.get_all_apontamentos()

    if df_all_ap.empty:
        st.warning("Nenhum apontamento encontrado na base de dados.")
    else:
        # Prepara mapeamento e colunas de forma robusta
        mapping = dm._get_system_mapping()
        
        # Identifica a coluna do Bloco
        col_bloco = None
        for alias in mapping.get("BLOCO_RAW", []):
            if alias.upper() in df_all_ap.columns:
                col_bloco = alias.upper()
                break
        if not col_bloco:
            col_bloco = "NUMERO_BLOCO" if "NUMERO_BLOCO" in df_all_ap.columns else df_all_ap.columns[3]

        # Identifica colunas
        col_data = "DATA_REG" if "DATA_REG" in df_all_ap.columns else "DATA"
        col_setor = "SETOR" if "SETOR" in df_all_ap.columns else "SETOR_AP"
        col_proc = "PROCESSO" if "PROCESSO" in df_all_ap.columns else "PROCESSO_APONTADO"
        col_operador = "OPERADOR" if "OPERADOR" in df_all_ap.columns else "OPERADOR"
        col_chapas = "QTD_CHAPAS" if "QTD_CHAPAS" in df_all_ap.columns else "QTD_CH"
        col_m2 = "QTDM2"
        for alias in mapping.get("QTD_M2", []):
            if alias.upper() in df_all_ap.columns:
                col_m2 = alias.upper()
                break

        # Converte a coluna de data de forma robusta
        df_all_ap["DATA_FILTRO"] = df_all_ap[col_data].apply(dm.parse_excel_date)
        
        # Inicializa o contador de resets de datas de forma segura se não existir
        if "dates_reset_counter" not in st.session_state:
            st.session_state["dates_reset_counter"] = 0
            
        # --- CAMPOS DE PESQUISA ---
        with st.container(border=True):
            st.markdown("#### 🔍 Filtros de Pesquisa")
            c_f1, c_f2, c_f3, c_f4 = st.columns(4)
            
            with c_f1:
                filtro_bloco = st.text_input("Nº Bloco", placeholder="Ex: 3706", key="c_filtro_bloco_val")
                
            with c_f2:
                unique_setores = sorted(list(str(s).strip() for s in df_all_ap[col_setor].dropna().unique() if str(s).strip()))
                filtro_setor = st.multiselect("Máquina/Setor", unique_setores, placeholder="Todos os setores", key="c_filtro_setor_val")
                
            with c_f3:
                unique_processos = sorted(list(str(p).strip() for p in df_all_ap[col_proc].dropna().unique() if str(p).strip()))
                filtro_processo = st.multiselect("Processo/Etapa", unique_processos, placeholder="Todos os processos", key="c_filtro_proc_val")
                
            with c_f4:
                filtro_operador = st.text_input("Operador (Busca parcial)", placeholder="Ex: MARLON", key="c_filtro_op_val")
            
            # Seletor de intervalo de datas de registro
            st.markdown("---")
            c_date, c_clear = st.columns([3, 1])
            with c_date:
                min_date = df_all_ap["DATA_FILTRO"].dropna().min()
                max_date = df_all_ap["DATA_FILTRO"].dropna().max()
                if pd.isna(min_date): min_date = datetime.now() - timedelta(days=30)
                if pd.isna(max_date): max_date = datetime.now()
                
                min_date_val = min_date.date() if hasattr(min_date, "date") else min_date
                max_date_val = max_date.date() if hasattr(max_date, "date") else max_date
                
                # Gera uma chave única dinâmica vinculada ao dates_reset_counter para forçar remota do componente
                date_widget_key = f"c_filtro_datas_reg_{st.session_state['dates_reset_counter']}"
                
                filtro_datas = st.date_input(
                    "Intervalo de Data de Registro",
                    value=(min_date_val, max_date_val),
                    format="DD/MM/YYYY",
                    key=date_widget_key
                )
            with c_clear:
                st.write(""); st.write("")
                # Botão para resetar as datas de pesquisa
                if st.button("Resetar Datas", width="stretch", key="c_btn_reset_dates"):
                    # Remove o cache de chave antigo se houver
                    st.session_state.pop(date_widget_key, None)
                    # Incrementa o contador para alterar a chave do date_input e forçar remontagem limpa no React
                    st.session_state["dates_reset_counter"] += 1
                    st.rerun()

        # Aplica os filtros ao DataFrame
        df_filtrado = df_all_ap.copy()
        
        if filtro_bloco:
            df_filtrado = df_filtrado[df_filtrado[col_bloco].apply(lambda x: dm.blocos_match(x, filtro_bloco))]
            
        if filtro_setor:
            df_filtrado = df_filtrado[df_filtrado[col_setor].astype(str).str.strip().isin(filtro_setor)]
            
        if filtro_processo:
            df_filtrado = df_filtrado[df_filtrado[col_proc].astype(str).str.strip().isin(filtro_processo)]
            
        if filtro_operador:
            df_filtrado = df_filtrado[df_filtrado[col_operador].astype(str).str.contains(filtro_operador, case=False, na=False)]

        # Aplica o filtro do intervalo de datas de registro
        if isinstance(filtro_datas, tuple) and len(filtro_datas) == 2:
            start_date, end_date = filtro_datas
            df_filtrado = df_filtrado[
                (df_filtrado["DATA_FILTRO"].dt.date >= start_date) & 
                (df_filtrado["DATA_FILTRO"].dt.date <= end_date)
            ]
        elif isinstance(filtro_datas, tuple) and len(filtro_datas) == 1:
            start_date = filtro_datas[0]
            df_filtrado = df_filtrado[df_filtrado["DATA_FILTRO"].dt.date >= start_date]

        # Ordena por data decrescente e ID decrescente
        df_filtrado = df_filtrado.sort_values(by=["DATA_FILTRO", "ID"], ascending=[False, False])

        st.subheader(f"📋 Lista de Apontamentos ({len(df_filtrado)} encontrados)")
        
        if df_filtrado.empty:
            st.info("Nenhum apontamento corresponde aos filtros selecionados.")
        else:
            # --- CÁLCULO DOS SUBTOTAIS (ATUALIZADOS CONFORME O FILTRO) ---
            total_ap = len(df_filtrado)
            total_chapas = int(pd.to_numeric(df_filtrado[col_chapas], errors="coerce").fillna(0).sum())
            total_m2 = pd.to_numeric(df_filtrado[col_m2], errors="coerce").fillna(0.0).sum()
            
            c_sub1, c_sub2, c_sub3 = st.columns(3)
            with c_sub1:
                st.metric("Total Apontamentos", f"{total_ap}")
            with c_sub2:
                st.metric("Total Chapas", f"{total_chapas:,}".replace(",", "."))
            with c_sub3:
                st.metric("Total Metragem (M²)", f"{total_m2:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
            # Prepara dados para exibição amigável
            df_display = df_filtrado.copy()
            df_display["DATA_FORMATTED"] = df_display["DATA_FILTRO"].apply(lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "-")
            
            # Converte e formata a coluna de data de início de produção (Data Produção)
            df_display["DATA_PROD_FILTRO"] = df_display["DATA_INICIO"].apply(dm.parse_excel_date)
            df_display["DATA_PROD_FORMATTED"] = df_display["DATA_PROD_FILTRO"].apply(lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "-")
            
            df_display["ID_SHOW"] = df_display["ID"].astype(str)
            df_display["BLOCO_SHOW"] = df_display[col_bloco].astype(str)
            df_display["CHAPAS_SHOW"] = df_display[col_chapas].apply(lambda x: int(float(x)) if pd.notna(x) and str(x).strip() != "" else 0)
            df_display["M2_SHOW"] = pd.to_numeric(df_display[col_m2], errors="coerce").fillna(0.0).round(3)
            
            # Colunas limpas para a tabela (adicionadas as colunas de data de produção e metragem M²)
            cols_show = ["ID_SHOW", "DATA_FORMATTED", "DATA_PROD_FORMATTED", "BLOCO_SHOW", "MATERIAL", col_proc, col_setor, "CHAPAS_SHOW", "M2_SHOW", col_operador]
            df_display_clean = df_display[[c for c in cols_show if c in df_display.columns]].copy()
            df_display_clean.columns = [
                "ID", "Data Registro", "Data Produção", "Bloco", "Material", "Processo", "Máquina/Setor", "Qtd. Chapas", "Qtd. M²", "Operador"
            ][:len(df_display_clean.columns)]
            
            # Adiciona nota amigável instruindo o usuário a clicar na linha
            st.caption("💡 *Dica: Clique em qualquer linha da tabela abaixo para carregar a Ficha Consolidada completa daquele apontamento!*")
            
            # Instancia o dataframe com seleção habilitada
            selection_event = st.dataframe(
                df_display_clean,
                width="stretch",
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row"
            )
            
            # Extrai a seleção de forma extremamente robusta
            selected_rows = []
            if selection_event:
                if hasattr(selection_event, "selection"):
                    sel = selection_event.selection
                    if isinstance(sel, dict):
                        selected_rows = sel.get("rows", [])
                    elif hasattr(sel, "rows"):
                        selected_rows = sel.rows
                elif isinstance(selection_event, dict):
                    sel = selection_event.get("selection", {})
                    if isinstance(sel, dict):
                        selected_rows = sel.get("rows", [])
                    elif hasattr(sel, "rows"):
                        selected_rows = sel.rows
            
            # Ficha Consolidada do apontamento selecionado
            st.write("---")
            st.subheader("📑 Visualizador de Ficha Consolidada")
            
            selected_id = None
            if selected_rows:
                selected_row_idx = selected_rows[0]
                # Mapeia de volta para o ID correto no df_filtrado
                selected_id = df_filtrado.iloc[selected_row_idx]["ID"]
            else:
                st.info("💡 Selecione uma linha na lista acima clicando nela para carregar a Ficha Consolidada detalhada.")
            
            if selected_id:
                ap_row = df_filtrado[df_filtrado["ID"] == selected_id].iloc[0]
                
                # Busca as paradas e insumos no data_manager
                with st.spinner("Buscando detalhes do apontamento..."):
                    df_paradas = dm.get_apontamento_paradas(selected_id)
                    df_insumos = dm.get_apontamento_insumos(selected_id)
                
                # Renderiza a ficha consolidada estilizada
                st.markdown(
                    f"""
                    <div style="background-color: #0E1117; padding: 25px; border-radius: 12px; border: 1px solid #30363D; margin-top: 15px;">
                        <h2 style="text-align: center; color: #58A6FF; margin-bottom: 5px;">📄 FICHA DE APONTAMENTO CONSOLIDADA</h2>
                        <h4 style="text-align: center; color: #8B949E; margin-bottom: 25px;">Apontamento ID #{selected_id}</h4>
                        <div style="display: flex; justify-content: space-around; font-size: 15px; margin-bottom: 25px; color: #C9D1D9; border-bottom: 1px solid #30363D; padding-bottom: 15px;">
                            <span><b>📅 Data:</b> {ap_row['DATA_FILTRO'].strftime('%d/%m/%Y') if pd.notna(ap_row['DATA_FILTRO']) else '-'}</span>
                            <span><b>⚙️ Máquina/Setor:</b> {ap_row.get(col_setor, '-')}</span>
                            <span><b>👤 Operador:</b> {ap_row.get(col_operador, '-')}</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                
                import datetime as dt_module

                def format_time_display(t_val):
                    if pd.isna(t_val) or t_val is None:
                        return "-"
                    if isinstance(t_val, dt_module.time):
                        return t_val.strftime("%H:%M")
                    if isinstance(t_val, dt_module.datetime):
                        return t_val.strftime("%H:%M")
                    s = str(t_val).strip()
                    if not s or s.lower() == "nan":
                        return "-"
                    if len(s) >= 5 and ":" in s:
                        parts = s.split(":")
                        try:
                            return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
                        except:
                            return s
                    return s

                raw_d_ini = ap_row.get("DIA_INICIO") or ap_row.get("DATA_INICIO") or ap_row.get("DATA_REG")
                p_d_ini = dm.parse_excel_date(raw_d_ini)
                d_ini_str = p_d_ini.strftime("%d/%m/%Y") if pd.notna(p_d_ini) else "-"
                
                raw_d_fim = ap_row.get("DIA_FIM") or ap_row.get("DATA_FIM") or raw_d_ini
                p_d_fim = dm.parse_excel_date(raw_d_fim)
                d_fim_str = p_d_fim.strftime("%d/%m/%Y") if pd.notna(p_d_fim) else "-"
                
                h_ini_str = format_time_display(ap_row.get("HORA_INICIO"))
                h_fim_str = format_time_display(ap_row.get("HORA_FIM"))

                # Duas colunas para dados principais
                col_res1, col_res2 = st.columns(2)
                
                with col_res1:
                    with st.container(border=True):
                        st.markdown("#### 🧱 Identificação do Bloco & Material")
                        st.write(f"**Número do Bloco:** {ap_row.get(col_bloco, '-')}")
                        st.write(f"**Material:** {ap_row.get('MATERIAL', '-')}")
                        st.write(f"**Espessura:** {ap_row.get('ESP', '-')}")
                        st.write(f"**Dimensões (Comp x Alt):** {ap_row.get('COMP', '-')}m × {ap_row.get('ALT', '-')}m")
                        
                        m2_val = ap_row.get("QTDM2") or ap_row.get("QTD_M2") or 0.0
                        if isinstance(m2_val, (int, float)):
                            st.write(f"**Volume Total Produzido:** {m2_val:.3f} M²")
                        else:
                            st.write(f"**Volume Total Produzido:** {m2_val} M²")
                            
                with col_res2:
                    with st.container(border=True):
                        st.markdown("#### ⚙️ Detalhes do Processo & Tempos")
                        st.write(f"**Processo / Etapa:** {ap_row.get(col_proc, '-')}")
                        st.write(f"**Turno:** {ap_row.get('TURNO', '-')}")
                        st.write(f"**Quantidade de Chapas:** {int(float(ap_row.get(col_chapas, 0))) if pd.notna(ap_row.get(col_chapas)) else 0} chapas")
                        st.write(f"**Data / Hora Início:** {d_ini_str} às {h_ini_str}")
                        st.write(f"**Data / Hora Fim:** {d_fim_str} às {h_fim_str}")
                        st.write(f"**Tempo de Processo:** {ap_row.get('TEMPO_PROCESSO', '-')} h/min")
                
                st.write("")
                
                # Paradas e Insumos lado a lado
                col_det1, col_det2 = st.columns(2)
                
                with col_det1:
                    st.markdown("#### 🛑 Paradas de Máquina (Downtimes)")
                    if df_paradas.empty:
                        st.info("Nenhuma parada de máquina registrada para este apontamento.")
                    else:
                        df_p_clean = df_paradas.copy()
                        cols_p = ["MOTIVO", "DATA_INICIO", "HORA_INICIO", "DATA_FIM", "HORA_FIM", "TEMPO"]
                        df_p_clean = df_p_clean[[c for c in cols_p if c in df_p_clean.columns]].copy()
                        df_p_clean.columns = ["Motivo", "Data Ini", "Hora Ini", "Data Fim", "Hora Fim", "Tempo"]
                        st.dataframe(df_p_clean, width="stretch", hide_index=True)
                        
                        # Totalização
                        st.success(f"**Total de paradas:** {len(df_p_clean)}")
                        
                with col_det2:
                    st.markdown("#### 🧪 Insumos Utilizados")
                    if df_insumos.empty:
                        st.info("Nenhum insumo ou material adicional registrado para este apontamento.")
                    else:
                        df_i_clean = df_insumos.copy()
                        cols_i = ["TIPO_INSUMO", "DESCRICAO", "QUANTIDADE", "UNIDADE", "TEMPO_SECAGEM"]
                        df_i_clean = df_i_clean[[c for c in cols_i if c in df_i_clean.columns]].copy()
                        df_i_clean.columns = ["Tipo Insumo", "Descrição", "Qtd", "Unidade", "Secagem (h)"]
                        st.dataframe(df_i_clean, width="stretch", hide_index=True)
                        
                        st.success(f"**Total de insumos:** {len(df_i_clean)}")

                # --- GERENCIAMENTO DE APONTAMENTOS (EXCLUSÃO E EDIÇÃO) ---
                st.write("")
                st.markdown("---")
                st.markdown("### ⚠️ Gerenciamento do Apontamento")
                
                c_del1, c_del2, c_del3 = st.columns([1.5, 1.5, 2])
                with c_del1:
                    btn_excluir = st.button("🗑️ EXCLUIR ESTE APONTAMENTO", type="primary", width="stretch", key=f"c_btn_excluir_{selected_id}")
                with c_del2:
                    btn_editar = st.button("📝 EDITAR APONTAMENTO", width="stretch", key=f"c_btn_editar_{selected_id}")
                
                if btn_excluir:
                    st.session_state[f"conf_del_{selected_id}"] = True
                    st.rerun()
                    
                if btn_editar:
                    st.session_state[f"edit_ap_{selected_id}"] = True
                    st.rerun()
                    
                if st.session_state.get(f"edit_ap_{selected_id}"):
                    # Carrega as paradas e insumos existentes para a edição
                    df_p_exist = dm.get_apontamento_paradas(selected_id)
                    df_i_exist = dm.get_apontamento_insumos(selected_id)
                    
                    # Prepara df_p_edit
                    df_p_edit = pd.DataFrame(columns=["MOTIVO", "DIA_INICIO", "HORA_INICIO", "DIA_FIM", "HORA_FIM"])
                    if not df_p_exist.empty:
                        # Mapeia DATA_INICIO / DATA_FIM para DIA_INICIO / DIA_FIM se vier do Excel antigo
                        col_map = {}
                        for c in df_p_exist.columns:
                            if c in ["DATA_INICIO", "DIA_INICIO"]:
                                col_map[c] = "DIA_INICIO"
                            elif c in ["DATA_FIM", "DIA_FIM"]:
                                col_map[c] = "DIA_FIM"
                            else:
                                col_map[c] = c
                        df_p_norm = df_p_exist.rename(columns=col_map)
                        df_p_edit = df_p_norm.reindex(columns=["MOTIVO", "DIA_INICIO", "HORA_INICIO", "DIA_FIM", "HORA_FIM"]).copy()
                        df_p_edit["DIA_INICIO"] = df_p_edit["DIA_INICIO"].apply(lambda x: dm.parse_excel_date(x).date() if pd.notna(dm.parse_excel_date(x)) else datetime.now().date())
                        df_p_edit["DIA_FIM"] = df_p_edit["DIA_FIM"].apply(lambda x: dm.parse_excel_date(x).date() if pd.notna(dm.parse_excel_date(x)) else datetime.now().date())
                    
                    df_p_edit["MOTIVO"] = df_p_edit["MOTIVO"].fillna("").astype(str).str.strip()
                    df_p_edit["HORA_INICIO"] = df_p_edit["HORA_INICIO"].fillna("").astype(str).str.replace("nan", "", regex=False).str.strip()
                    df_p_edit["HORA_FIM"] = df_p_edit["HORA_FIM"].fillna("").astype(str).str.replace("nan", "", regex=False).str.strip()
                    
                    # Prepara df_i_edit safely
                    df_i_edit = pd.DataFrame(columns=["TIPO_INSUMO", "DESCRICAO", "QUANTIDADE", "UNIDADE", "TEMPO_SECAGEM", "CABECAS", "INSUMO_DETALHE"])
                    if not df_i_exist.empty:
                        df_i_edit = df_i_exist.reindex(columns=["TIPO_INSUMO", "DESCRICAO", "QUANTIDADE", "UNIDADE", "TEMPO_SECAGEM", "CABECAS", "INSUMO_DETALHE"]).copy()
                    
                    df_i_edit["QUANTIDADE"] = pd.to_numeric(df_i_edit["QUANTIDADE"], errors="coerce").fillna(0.0)
                    df_i_edit["TIPO_INSUMO"] = df_i_edit["TIPO_INSUMO"].fillna("").astype(str).str.strip()
                    df_i_edit["DESCRICAO"] = df_i_edit["DESCRICAO"].fillna("").astype(str).str.strip()
                    df_i_edit["UNIDADE"] = df_i_edit["UNIDADE"].fillna("").astype(str).str.strip()
                    df_i_edit["TEMPO_SECAGEM"] = df_i_edit["TEMPO_SECAGEM"].fillna("").astype(str).str.replace(".0", "", regex=False).str.replace("nan", "", regex=False).str.strip()
                    df_i_edit["CABECAS"] = df_i_edit["CABECAS"].fillna("").astype(str).str.replace("nan", "", regex=False).str.strip()
                    df_i_edit["INSUMO_DETALHE"] = df_i_edit["INSUMO_DETALHE"].fillna("").astype(str).str.replace("nan", "", regex=False).str.strip()

                    st.write("")
                    with st.form(f"form_editar_ap_{selected_id}"):
                        st.markdown("### 📝 Editar Dados do Apontamento")
                        
                        c_ed1, c_ed2, c_ed3 = st.columns(3)
                        with c_ed1:
                            edit_proc = st.text_input("Processo*", value=str(ap_row.get(col_proc, "")))
                            edit_material = st.text_input("Material*", value=str(ap_row.get("MATERIAL", "")))
                            raw_ch = ap_row.get(col_chapas, 0)
                            try: edit_qtd_val = int(float(raw_ch)) if pd.notna(raw_ch) and str(raw_ch).strip() != "" else 0
                            except: edit_qtd_val = 0
                            edit_qtd = st.number_input("Qtd Chapas*", min_value=0, step=1, value=edit_qtd_val)
                            raw_esp = ap_row.get("ESP", 0)
                            try: edit_esp_val = float(raw_esp) if pd.notna(raw_esp) and str(raw_esp).strip() != "" else 0.0
                            except: edit_esp_val = 0.0
                            edit_esp = st.number_input("Espessura", min_value=0.0, step=0.1, value=edit_esp_val)
                            
                        with c_ed2:
                            unique_setores_clean = [s for s in unique_setores if s != "Todos"]
                            current_setor = str(ap_row.get(col_setor, "")).strip()
                            idx_setor = 0
                            if current_setor in unique_setores_clean:
                                idx_setor = unique_setores_clean.index(current_setor)
                            edit_setor = st.selectbox("Máquina/Setor*", unique_setores_clean, index=idx_setor)
                            raw_comp = ap_row.get("COMP", 0)
                            try: edit_comp_val = float(raw_comp) if pd.notna(raw_comp) and str(raw_comp).strip() != "" else 0.0
                            except: edit_comp_val = 0.0
                            edit_comp = st.number_input("Comprimento (m)", min_value=0.0, step=0.01, value=edit_comp_val)
                            raw_alt = ap_row.get("ALT", 0)
                            try: edit_alt_val = float(raw_alt) if pd.notna(raw_alt) and str(raw_alt).strip() != "" else 0.0
                            except: edit_alt_val = 0.0
                            edit_alt = st.number_input("Altura (m)", min_value=0.0, step=0.01, value=edit_alt_val)
                            edit_op = st.text_input("Operador", value=str(ap_row.get(col_operador, "")))
                            
                        with c_ed3:
                            raw_dia_ini = ap_row.get("DIA_INICIO") or ap_row.get("DATA_INICIO") or ap_row.get("DATA_REG")
                            parsed_dia_ini = dm.parse_excel_date(raw_dia_ini)
                            if pd.isna(parsed_dia_ini): parsed_dia_ini = datetime.now()
                            edit_dia_ini = st.date_input("Dia Início Produção*", value=parsed_dia_ini.date() if hasattr(parsed_dia_ini, "date") else parsed_dia_ini, format="DD/MM/YYYY")
                            raw_dia_fim = ap_row.get("DIA_FIM") or ap_row.get("DATA_FIM") or raw_dia_ini
                            parsed_dia_fim = dm.parse_excel_date(raw_dia_fim)
                            if pd.isna(parsed_dia_fim): parsed_dia_fim = datetime.now()
                            edit_dia_fim = st.date_input("Dia Fim Produção*", value=parsed_dia_fim.date() if hasattr(parsed_dia_fim, "date") else parsed_dia_fim, format="DD/MM/YYYY")
                            c_t1, c_t2 = st.columns(2)
                            with c_t1: edit_hora_ini = st.text_input("Hora Início", value=str(ap_row.get("HORA_INICIO", "")))
                            with c_t2: edit_hora_fim = st.text_input("Hora Fim", value=str(ap_row.get("HORA_FIM", "")))
                            raw_turno = str(ap_row.get("TURNO", "D")).strip().upper()
                            edit_turno = st.selectbox("Turno", ["D", "N"], index=0 if raw_turno == "D" else 1)
                            st.text_input("Tempo Processo (Calculado automaticamente)", value=str(ap_row.get("TEMPO_PROCESSO", "")), disabled=True)

                        st.markdown("---")
                        st.markdown("### 🛑 Paradas & 🧪 Insumos do Apontamento")
                        c_edi, c_edp = st.columns(2)
                        
                        with c_edi:
                            st.markdown("#### 🧪 Insumos Utilizados")
                            edit_df_insumos = st.data_editor(
                                df_i_edit,
                                num_rows="dynamic",
                                width="stretch",
                                column_config={
                                    "TIPO_INSUMO": st.column_config.SelectboxColumn("Tipo Insumo*", options=["RESINA", "ENDURECEDOR", "MANTA", "ABRASIVO", "OUTROS"], required=True),
                                    "DESCRICAO": st.column_config.TextColumn("Descrição*", required=True),
                                    "QUANTIDADE": st.column_config.NumberColumn("Qtd*", min_value=0.0, format="%.3f"),
                                    "UNIDADE": st.column_config.SelectboxColumn("Unidade", options=["KG", "M²", "UNID", "OUTROS"]),
                                    "TEMPO_SECAGEM": st.column_config.TextColumn("Tempo Secagem (h)"),
                                    "CABECAS": st.column_config.TextColumn("Cab."),
                                    "INSUMO_DETALHE": st.column_config.TextColumn("Detalhe")
                                },
                                key=f"editor_edit_insumos_{selected_id}"
                            )
                            
                        with c_edp:
                            st.markdown("#### 🛑 Paradas de Máquina (Downtimes)")
                            df_tp_op = dm.get_tipo_paradas()
                            lista_motivos_ap = sorted(list(df_tp_op["MOTIVO"].dropna().unique())) if not df_tp_op.empty else []
                            col_conf_paradas_edit = {
                                "DIA_INICIO": st.column_config.DateColumn("D.Início", format="DD/MM/YYYY"),
                                "DIA_FIM": st.column_config.DateColumn("D.Fim", format="DD/MM/YYYY"),
                                "HORA_INICIO": st.column_config.TextColumn("H.Início"),
                                "HORA_FIM": st.column_config.TextColumn("H.Fim")
                            }
                            if lista_motivos_ap:
                                col_conf_paradas_edit["MOTIVO"] = st.column_config.SelectboxColumn("Motivo da Parada*", options=lista_motivos_ap, required=True)
                            else:
                                col_conf_paradas_edit["MOTIVO"] = st.column_config.TextColumn("Motivo da Parada*", required=True)
                            edit_df_paradas = st.data_editor(
                                df_p_edit,
                                num_rows="dynamic",
                                width="stretch",
                                column_config=col_conf_paradas_edit,
                                key=f"editor_edit_paradas_{selected_id}"
                            )

                        c_ebtn1, c_ebtn2 = st.columns(2)
                        with c_ebtn1: btn_salvar_edit = st.form_submit_button("💾 Salvar Alterações", type="primary", width="stretch")
                        with c_ebtn2: btn_cancelar_edit = st.form_submit_button("↩️ Cancelar Edição", width="stretch")
                            
                        if btn_salvar_edit:
                            # Reconstroi os dataframes dos editores a partir do session state para evitar o bug do st.form
                            edit_df_insumos = get_edited_df(df_i_edit, f"editor_edit_insumos_{selected_id}")
                            edit_df_paradas = get_edited_df(df_p_edit, f"editor_edit_paradas_{selected_id}")
                            
                            if not edit_proc or not edit_setor or not edit_material or edit_qtd is None or not edit_dia_ini or not edit_dia_fim:
                                st.error("❌ Preencha todos os campos obrigatórios (*) com valores válidos.")
                            else:
                                def parse_time_local(t_str):
                                    if not t_str: return None
                                    import re
                                    from datetime import time
                                    t = re.sub(r'\D', '', str(t_str))
                                    if len(t) == 4:
                                        try: return time(int(t[:2]), int(t[2:]))
                                        except: return None
                                    elif len(t) == 3:
                                        try: return time(int(t[:1]), int(t[1:]))
                                        except: return None
                                    elif ":" in str(t_str):
                                        parts = str(t_str).split(":")
                                        if len(parts) == 2:
                                            try: return time(int(parts[0]), int(parts[1]))
                                            except: return None
                                    return None
                                    
                                h_ini = parse_time_local(edit_hora_ini)
                                h_fim = parse_time_local(edit_hora_fim)
                                if not h_ini or not h_fim:
                                    st.error("❌ Formato de Hora de Início ou Fim inválido (use HH:MM ou HHMM).")
                                else:
                                    from datetime import datetime as dt_class
                                    dt_ini = dt_class.combine(edit_dia_ini, h_ini)
                                    dt_fim = dt_class.combine(edit_dia_fim, h_fim)
                                    diff_minutes = (dt_fim - dt_ini).total_seconds() / 60
                                    
                                    if diff_minutes <= 0:
                                        st.error("❌ A data/hora de fim deve ser posterior à data/hora de início.")
                                    else:
                                        # Processamento e Validação Estrita de Paradas
                                        par_finais = []
                                        total_mp = 0
                                        erro_parada = None
                                        
                                        for _, p_row in edit_df_paradas.iterrows():
                                            if p_row.get("MOTIVO"):
                                                h_ini_p = parse_time_local(p_row.get("HORA_INICIO"))
                                                h_fim_p = parse_time_local(p_row.get("HORA_FIM"))
                                                
                                                if not h_ini_p or not h_fim_p:
                                                    erro_parada = f"❌ Hora inválida ou incompleta na parada '{p_row['MOTIVO']}'."
                                                    break
                                                    
                                                d_ini_p_parsed = p_row.get("DIA_INICIO")
                                                d_fim_p_parsed = p_row.get("DIA_FIM")
                                                
                                                if isinstance(d_ini_p_parsed, str):
                                                    d_ini_p_parsed = dm.parse_excel_date(d_ini_p_parsed)
                                                if isinstance(d_fim_p_parsed, str):
                                                    d_fim_p_parsed = dm.parse_excel_date(d_fim_p_parsed)
                                                    
                                                if hasattr(d_ini_p_parsed, "date"): d_ini_p_parsed = d_ini_p_parsed.date()
                                                if hasattr(d_fim_p_parsed, "date"): d_fim_p_parsed = d_fim_p_parsed.date()
                                                
                                                if not d_ini_p_parsed or not d_fim_p_parsed:
                                                    erro_parada = f"❌ Data de início ou fim inválida na parada '{p_row['MOTIVO']}'."
                                                    break
                                                    
                                                dt_ini_p = datetime.combine(d_ini_p_parsed, h_ini_p)
                                                dt_fim_p = datetime.combine(d_fim_p_parsed, h_fim_p)
                                                mp = (dt_fim_p - dt_ini_p).total_seconds() / 60
                                                
                                                # Validação: Dentro do intervalo do processo?
                                                if dt_ini_p < dt_ini or dt_fim_p > dt_fim:
                                                    erro_parada = f"❌ Parada '{p_row['MOTIVO']}' ({h_ini_p.strftime('%H:%M')} às {h_fim_p.strftime('%H:%M')}) está fora do intervalo do processo."
                                                    break
                                                
                                                if mp <= 0:
                                                    erro_parada = f"❌ Tempo da parada '{p_row['MOTIVO']}' deve ser positivo."
                                                    break

                                                total_mp += mp
                                                par_finais.append({
                                                    "MOTIVO": p_row["MOTIVO"], 
                                                    "DIA_INICIO": d_ini_p_parsed.strftime("%d/%m/%Y"), 
                                                    "HORA_INICIO": h_ini_p.strftime("%H:%M"), 
                                                    "DIA_FIM": d_fim_p_parsed.strftime("%d/%m/%Y"), 
                                                    "HORA_FIM": h_fim_p.strftime("%H:%M"), 
                                                    "TEMPO": f"{int(mp // 60):02d}:{int(mp % 60):02d}"
                                                })

                                        if erro_parada:
                                            st.error(erro_parada)
                                        elif total_mp > diff_minutes:
                                            st.error(f"❌ Soma das paradas ({int(total_mp)} min) é maior que o tempo total de processo ({int(diff_minutes)} min).")
                                        else:
                                            # Processamento e Formatação Estrita de Insumos
                                            ins_finais = []
                                            for _, i_row in edit_df_insumos.iterrows():
                                                if i_row.get("TIPO_INSUMO") and i_row.get("DESCRICAO"):
                                                    try: q_val = float(i_row.get("QUANTIDADE", 0.0))
                                                    except: q_val = 0.0
                                                    
                                                    ins_finais.append({
                                                        "TIPO_INSUMO": str(i_row["TIPO_INSUMO"]).upper(),
                                                        "DESCRICAO": str(i_row["DESCRICAO"]).upper(),
                                                        "QUANTIDADE": q_val,
                                                        "UNIDADE": str(i_row.get("UNIDADE", "UNID")).upper(),
                                                        "TEMPO_SECAGEM": i_row.get("TEMPO_SECAGEM", "") if pd.notna(i_row.get("TEMPO_SECAGEM")) else "",
                                                        "CABECAS": i_row.get("CABECAS", "") if pd.notna(i_row.get("CABECAS")) else "",
                                                        "INSUMO_DETALHE": i_row.get("INSUMO_DETALHE", "") if pd.notna(i_row.get("INSUMO_DETALHE")) else ""
                                                    })
                                                    
                                            calculated_tempo = f"{int(diff_minutes // 60):02d}:{int(diff_minutes % 60):02d}"
                                            
                                            updates = {
                                                "PROCESSO_APONTADO": edit_proc,
                                                "SETOR_AP": edit_setor,
                                                "NOME_MATERIAL": edit_material.upper(),
                                                "QTD_CH": edit_qtd,
                                                "ESP": edit_esp if edit_esp else "",
                                                "COMP": edit_comp if edit_comp else "",
                                                "ALT": edit_alt if edit_alt else "",
                                                "OPERADOR": edit_op.upper() if edit_op else "",
                                                "TURNO": edit_turno,
                                                "TEMPO_PROCESSO": calculated_tempo,
                                                "HORA_INICIO": h_ini.strftime("%H:%M"),
                                                "HORA_FIM": h_fim.strftime("%H:%M"),
                                                "DIA_INICIO": edit_dia_ini.strftime("%d/%m/%Y"),
                                                "DIA_FIM": edit_dia_fim.strftime("%d/%m/%Y"),
                                                "QTD_M2": round(edit_comp * edit_alt * edit_qtd, 3) if edit_comp and edit_alt else 0.0
                                            }
                                            with st.spinner("Salvando alterações no Excel..."):
                                                ok_main = dm.update_apontamento(selected_id, updates)
                                                ok_rel = dm.update_apontamento_relations(selected_id, par_finais, ins_finais)
                                                
                                                if ok_main and ok_rel:
                                                    st.success("✅ Apontamento, Paradas e Insumos atualizados com sucesso!")
                                                    st.session_state.pop(f"edit_ap_{selected_id}", None)
                                                    st.session_state.pop("df_ap_cache", None)
                                                    st.rerun()
                                                else:
                                                    st.error("❌ Ocorreu um erro ao salvar no Excel. Verifique se o arquivo está aberto em outro programa.")
                                        
                        if btn_cancelar_edit:
                            st.session_state.pop(f"edit_ap_{selected_id}", None)
                            st.rerun()
                
                if st.session_state.get(f"conf_del_{selected_id}"):
                    with st.container(border=True):
                        st.warning(f"⚠️ **Atenção:** Você tem certeza de que deseja excluir permanentemente o apontamento ID #{selected_id} (Bloco {ap_row.get(col_bloco)})? Esta ação limpará o processo principal, bem como todas as paradas e insumos vinculados no Excel e não poderá ser desfeita.")
                        c_conf1, c_conf2 = st.columns(2)
                        with c_conf1:
                            if st.button("🔴 Sim, confirmar exclusão", type="primary", width="stretch", key=f"c_btn_yes_del_{selected_id}"):
                                with st.spinner("Excluindo registro do Excel..."):
                                    if dm.delete_apontamento(selected_id):
                                        st.success("✅ Apontamento excluído com sucesso!")
                                        # Reseta caches de apontamento diário e da busca
                                        st.session_state.pop("df_ap_cache", None)
                                        st.session_state.pop(f"conf_del_{selected_id}", None)
                                        # Recarrega a página
                                        st.rerun()
                                    else:
                                        st.error("❌ Ocorreu um erro ao excluir o apontamento no Excel. Verifique se o arquivo está aberto em outro programa.")
                        with c_conf2:
                            if st.button("↩️ Cancelar", width="stretch", key=f"c_btn_no_del_{selected_id}"):
                                st.session_state.pop(f"conf_del_{selected_id}", None)
                                st.rerun()


with tab_analises:
    st.header("📈 Análises e Indicadores de Produção")
    
    df_raw_an = dm.get_all_apontamentos()
    mapping_an = dm._get_system_mapping()
    
    def find_col_an(key):
        for alias in mapping_an.get(key, []):
            if alias.upper() in df_raw_an.columns: return alias.upper()
        return None

    c_dt = find_col_an("DIA_FIM") or find_col_an("DIA_INICIO")
    c_hr = find_col_an("HORA_FIM") or find_col_an("HORA_INICIO")
    c_dt_ini = find_col_an("DIA_INICIO")
    c_hr_ini = find_col_an("HORA_INICIO")
    c_m2 = find_col_an("QTD_M2")
    c_ch = find_col_an("QTD_CH")
    c_st = find_col_an("SETOR_AP")
    c_tr = find_col_an("TURNO")
    c_pr = find_col_an("PROCESSO_APONTADO")

    if df_raw_an.empty:
        st.warning("⚠️ Base de dados vazia ou não carregada.")
    elif not c_dt or not c_m2 or not c_ch:
        st.warning("⚠️ Colunas essenciais não encontradas no DB para gerar análises.")
    else:
        hoje = datetime.now().date()
        
        # Filtro de Período
        col_f1, col_f2 = st.columns([1, 3])
        with col_f1:
            periodo = st.selectbox("Período", ["Últimos 7 dias", "Últimos 30 dias", "Mês Atual", "Personalizado", "Todo o Histórico"], index=1)
            
            hoje = datetime.now().date()
            data_ini_custom = hoje - timedelta(days=7)
            data_fim_custom = hoje
            if periodo == "Personalizado":
                res_datas = st.date_input("Intervalo", [hoje - timedelta(days=7), hoje], key="analise_date_range")
                if isinstance(res_datas, (list, tuple)) and len(res_datas) == 2:
                    data_ini_custom, data_fim_custom = res_datas
                elif isinstance(res_datas, (list, tuple)) and len(res_datas) == 1:
                    data_ini_custom = data_fim_custom = res_datas[0]
                else:
                    data_ini_custom = data_fim_custom = res_datas

        # Inicializa a lista de máquinas disponíveis e adiciona o multiselect
        maquinas_lista = sorted([str(x) for x in df_raw_an[c_st].unique() if str(x) != "" and str(x) != "nan"])
        with col_f2:
            maquinas_sel = st.multiselect(
                "Marcar/Desmarcar Máquinas (Setores)",
                options=maquinas_lista,
                default=maquinas_lista,
                placeholder="Selecione as máquinas para filtrar os indicadores..."
            )

        # Processamento
        df_an = df_raw_an.copy()
        for col_num in [c_m2, c_ch]:
            df_an[col_num] = pd.to_numeric(df_an[col_num], errors='coerce').fillna(0)

        if c_dt:
            df_an[c_dt] = pd.to_datetime(df_an[c_dt], errors='coerce', dayfirst=True)
        if c_dt_ini and c_dt_ini != c_dt:
            df_an[c_dt_ini] = pd.to_datetime(df_an[c_dt_ini], errors='coerce', dayfirst=True)
        
        def get_dia_producao(row):
            dt = row[c_dt]
            val_hr = row.get(c_hr)
            
            # Fallback a nível de linha se a data fim estiver vazia
            if pd.isna(dt) and c_dt_ini:
                dt = row[c_dt_ini]
                val_hr = row.get(c_hr_ini)
                
            if pd.isna(dt): return None
            try:
                hour = val_hr.hour if hasattr(val_hr, 'hour') else int(str(val_hr).split(":")[0])
                if hour < 7: return (dt - timedelta(days=1)).date()
                return dt.date()
            except: return dt.date()
        
        df_an["DIA_PROD"] = df_an.apply(get_dia_producao, axis=1)
        df_an = df_an.dropna(subset=["DIA_PROD"])
        dias_unicos = []
        if periodo == "Últimos 7 dias":
            dias_unicos = sorted(df_an["DIA_PROD"].unique(), reverse=True)[:7]
            if dias_unicos:
                df_an = df_an[df_an["DIA_PROD"] >= min(dias_unicos)]
        elif periodo == "Últimos 30 dias":
            dias_unicos = sorted(df_an["DIA_PROD"].unique(), reverse=True)[:30]
            if dias_unicos:
                df_an = df_an[df_an["DIA_PROD"] >= min(dias_unicos)]
        elif periodo == "Mês Atual":
            df_an = df_an[pd.to_datetime(df_an["DIA_PROD"]).dt.month == hoje.month]
        elif periodo == "Personalizado":
            df_an = df_an[(df_an["DIA_PROD"] >= data_ini_custom) & (df_an["DIA_PROD"] <= data_fim_custom)]
        
        # Filtro de Máquinas selecionadas
        if maquinas_sel:
            df_an = df_an[df_an[c_st].isin(maquinas_sel)]
        
        df_an["REFEITO"] = df_an[c_pr].astype(str).str.upper().str.contains("REPASSE|REFEITO|REPROCESSO")
        df_an["TIPO_PROD"] = df_an["REFEITO"].map({True: "Refeito", False: "Normal"})

        st.write("---")
        c_opt1, c_opt2 = st.columns([2, 2])
        with c_opt1:
            metrica_board = st.radio("Métrica para o Board:", ["Chapas", "M²"], horizontal=True)
            col_valor = c_ch if metrica_board == "Chapas" else c_m2

        m_tot_m2 = df_an[c_m2].sum(); m_tot_ch = df_an[c_ch].sum()
        m_refeito_m2 = df_an[df_an["REFEITO"]][c_m2].sum(); m_refeito_ch = df_an[df_an["REFEITO"]][c_ch].sum()

        c_met1, c_met2, c_met3, c_met4 = st.columns(4)
        if metrica_board == "Chapas":
            p_refeito = (m_refeito_ch / m_tot_ch * 100) if m_tot_ch > 0 else 0
            c_met1.metric("Total Chapas", f"{int(m_tot_ch)}")
            c_met2.metric("Produção Normal", f"{int(df_an[df_an['REFEITO']==False][c_ch].sum())}")
            c_met3.metric("Produção Refeito", f"{int(m_refeito_ch)}")
        else:
            p_refeito = (m_refeito_m2 / m_tot_m2 * 100) if m_tot_m2 > 0 else 0
            c_met1.metric("Total Produzido (M²)", f"{m_tot_m2:,.2f}")
            c_met2.metric("Produção Normal", f"{df_an[df_an['REFEITO']==False][c_m2].sum():,.2f}")
            c_met3.metric("Produção Refeito", f"{m_refeito_m2:,.2f}")
        c_met4.metric(f"% Refeito ({metrica_board})", f"{p_refeito:.1f}%")

        st.divider()
        st.subheader(f"📊 Board de Produtividade Diária ({metrica_board})")
        
        try:
            df_base_pivot = df_an.pivot_table(index="DIA_PROD", columns=[c_st, c_tr, "TIPO_PROD"], values=col_valor, aggfunc='sum').fillna(0)
            combinacoes = sorted(list(set([(c[0], c[1]) for c in df_base_pivot.columns])))
            df_final = pd.DataFrame(index=df_base_pivot.index)
            for st_val, tr_val in combinacoes:
                vn = df_base_pivot[(st_val, tr_val, "Normal")] if (st_val, tr_val, "Normal") in df_base_pivot.columns else pd.Series(0, index=df_base_pivot.index)
                vr = df_base_pivot[(st_val, tr_val, "Refeito")] if (st_val, tr_val, "Refeito") in df_base_pivot.columns else pd.Series(0, index=df_base_pivot.index)
                df_final[f"{st_val} / {tr_val}"] = [f"{int(n)} ({int(r)})" if metrica_board == "Chapas" else f"{n:,.1f} ({r:,.1f})" if (n > 0 or r > 0) else "-" for n, r in zip(vn, vr)]
            
            df_final["TOTAL GERAL"] = [f"{int(t)}" if metrica_board == "Chapas" else f"{t:,.1f}" for t in df_base_pivot.sum(axis=1)]
            df_final = df_final.sort_index(ascending=False)
            df_final.index = [d.strftime('%d/%m/%Y') for d in df_final.index]
            
            # Tabela em cima ocupando a largura total para visão rápida
            st.dataframe(df_final, width="stretch")
            st.caption("📌 Legenda: **Normal (Refeita)**.")
            
            # Gráfico embaixo ocupando a largura total para detalhamento visual
            df_board_gr = df_an.groupby(["DIA_PROD", c_st, c_tr, "TIPO_PROD"])[col_valor].sum().reset_index()
            df_board_gr["MAQ_TURNO"] = df_board_gr[c_st] + " (" + df_board_gr[c_tr] + ")"
            df_board_gr["DIA"] = df_board_gr["DIA_PROD"].apply(lambda d: d.strftime('%d/%m'))
            
            # Calcula ponto médio de cada segmento para centralizar texto
            tipo_order = {"Refeito": 0, "Normal": 1}  # Refeito na base, Normal em cima
            df_board_gr["_sort"] = df_board_gr["TIPO_PROD"].map(tipo_order)
            df_board_gr = df_board_gr.sort_values(["DIA_PROD", "MAQ_TURNO", "_sort"]).reset_index(drop=True)
            
            df_board_gr["_y0"] = 0.0
            df_board_gr["_y1"] = 0.0
            for (dia, maq), grp in df_board_gr.groupby(["DIA_PROD", "MAQ_TURNO"]):
                cum = 0.0
                for idx in grp.index:
                    val = df_board_gr.loc[idx, col_valor]
                    df_board_gr.loc[idx, "_y0"] = cum
                    df_board_gr.loc[idx, "_y1"] = cum + val
                    cum += val
            df_board_gr["_mid"] = (df_board_gr["_y0"] + df_board_gr["_y1"]) / 2
            
            # Label vazio para segmentos sem valor (evita texto em barras vazias)
            fmt = '.0f' if metrica_board == 'Chapas' else '.1f'
            df_board_gr["_label"] = df_board_gr[col_valor].apply(lambda v: f"{v:{fmt}}" if v > 0 else "")
            
            bars = alt.Chart(df_board_gr).mark_bar().encode(
                x=alt.X('MAQ_TURNO:N', title=None, sort='ascending'),
                y=alt.Y(f'{col_valor}:Q', title=metrica_board, stack='zero'),
                color=alt.Color('TIPO_PROD:N', title='Tipo', scale=alt.Scale(domain=['Normal', 'Refeito'], range=['#00CC96', '#EF553B'])),
                order=alt.Order('_sort:Q')
            )
            
            # Valores centralizados dentro de cada segmento
            text_seg = alt.Chart(df_board_gr).mark_text(
                align='center', baseline='middle', color='white', fontSize=15, fontWeight='bold', stroke='black', strokeWidth=0.5
            ).encode(
                x=alt.X('MAQ_TURNO:N', sort='ascending'),
                y=alt.Y('_mid:Q'),
                text=alt.Text('_label:N')
            )
            
            # Totais no topo
            text_totals = alt.Chart(df_board_gr).mark_text(
                align='center', baseline='bottom', dy=-5, fontSize=18, fontWeight='bold', color='white', stroke='black', strokeWidth=0.6
            ).encode(x=alt.X('MAQ_TURNO:N', sort='ascending'), y=alt.Y(f'sum({col_valor}):Q'), text=alt.Text(f'sum({col_valor}):Q', format=fmt))
            
            chart_layered = (bars + text_seg + text_totals).properties(width=alt.Step(100), height=450)
            
            faceted_chart = chart_layered.facet(
                column=alt.Column('DIA:N', title='Dia de Produção', sort=alt.SortOrder('ascending'))
            ).configure_view(stroke=None).configure_axis(labelFontSize=13, titleFontSize=15)
            
            st.altair_chart(faceted_chart, width="stretch")
            
        except Exception as e:
            st.info(f"Aguardando dados suficientes. (Erro: {e})")

        st.divider()
        c_gr1, c_gr2 = st.columns(2)
        with c_gr1:
            st.subheader(f"📅 Evolução Diária ({metrica_board})")
            df_evol = df_an.groupby(["DIA_PROD", "TIPO_PROD"])[col_valor].sum().reset_index()
            df_evol["DIA"] = df_evol["DIA_PROD"].apply(lambda d: d.strftime('%d/%m'))
            
            # Calcula ponto médio de cada segmento para centralizar texto
            tipo_order_e = {"Refeito": 0, "Normal": 1}
            df_evol["_sort"] = df_evol["TIPO_PROD"].map(tipo_order_e)
            df_evol = df_evol.sort_values(["DIA_PROD", "_sort"]).reset_index(drop=True)
            df_evol["_y0"] = 0.0; df_evol["_y1"] = 0.0
            for dia, grp in df_evol.groupby("DIA_PROD"):
                cum = 0.0
                for idx in grp.index:
                    val = df_evol.loc[idx, col_valor]
                    df_evol.loc[idx, "_y0"] = cum
                    df_evol.loc[idx, "_y1"] = cum + val
                    cum += val
            df_evol["_mid"] = (df_evol["_y0"] + df_evol["_y1"]) / 2
            
            # Label e ordenação cronológica
            fmt_e = '.0f' if metrica_board == 'Chapas' else '.1f'
            df_evol["_label"] = df_evol[col_valor].apply(lambda v: f"{v:{fmt_e}}" if v > 0 else "")
            dias_ordenados = sorted(df_evol["DIA_PROD"].unique())
            mapa_dias = {d: d.strftime('%d/%m') for d in dias_ordenados}
            ordem_dias = [mapa_dias[d] for d in dias_ordenados]
            df_evol["DIA"] = df_evol["DIA_PROD"].map(mapa_dias)
            
            bars_evol = alt.Chart(df_evol).mark_bar().encode(
                x=alt.X('DIA:N', title='Dia', sort=ordem_dias),
                y=alt.Y(f'{col_valor}:Q', title=metrica_board, stack='zero'),
                color=alt.Color('TIPO_PROD:N', title='Tipo', scale=alt.Scale(domain=['Normal', 'Refeito'], range=['#00CC96', '#EF553B'])),
                order=alt.Order('_sort:Q')
            )
            text_evol = alt.Chart(df_evol).mark_text(
                align='center', baseline='middle', color='white', fontSize=16, fontWeight='bold', stroke='black', strokeWidth=0.5
            ).encode(
                x=alt.X('DIA:N', sort=ordem_dias),
                y=alt.Y('_mid:Q'),
                text=alt.Text('_label:N')
            )
            totals_evol = alt.Chart(df_evol).mark_text(
                align='center', baseline='bottom', dy=-5, fontSize=18, fontWeight='bold', color='white', stroke='black', strokeWidth=0.5
            ).encode(
                x=alt.X('DIA:N', sort=ordem_dias),
                y=alt.Y(f'sum({col_valor}):Q'),
                text=alt.Text(f'sum({col_valor}):Q', format=fmt_e)
            )
            st.altair_chart((bars_evol + text_evol + totals_evol).properties(height=450), width="stretch")
            
        with c_gr2:
            st.subheader(f"🌙 Produtividade por Turno ({metrica_board})")
            df_turno = df_an.groupby([c_tr, "TIPO_PROD"])[col_valor].sum().reset_index()
            
            # Calcula ponto médio de cada segmento para centralizar texto
            tipo_order_t = {"Refeito": 0, "Normal": 1}
            df_turno["_sort"] = df_turno["TIPO_PROD"].map(tipo_order_t)
            df_turno = df_turno.sort_values([c_tr, "_sort"]).reset_index(drop=True)
            df_turno["_y0"] = 0.0; df_turno["_y1"] = 0.0
            for turno, grp in df_turno.groupby(c_tr):
                cum = 0.0
                for idx in grp.index:
                    val = df_turno.loc[idx, col_valor]
                    df_turno.loc[idx, "_y0"] = cum
                    df_turno.loc[idx, "_y1"] = cum + val
                    cum += val
            df_turno["_mid"] = (df_turno["_y0"] + df_turno["_y1"]) / 2
            df_turno_text = df_turno[df_turno[col_valor] > 0].copy()
            
            bars_turno = alt.Chart(df_turno).mark_bar().encode(
                x=alt.X(f'{c_tr}:N', title='Turno'),
                y=alt.Y(f'{col_valor}:Q', title=metrica_board, stack='zero'),
                color=alt.Color('TIPO_PROD:N', title='Tipo', scale=alt.Scale(domain=['Normal', 'Refeito'], range=['#00CC96', '#EF553B'])),
                order=alt.Order('_sort:Q')
            )
            text_turno = alt.Chart(df_turno_text).mark_text(
                align='center', baseline='middle', color='white', fontSize=16, fontWeight='bold', stroke='black', strokeWidth=0.5
            ).encode(
                x=alt.X(f'{c_tr}:N'),
                y=alt.Y('_mid:Q'),
                text=alt.Text(f'{col_valor}:Q', format='.0f' if metrica_board == 'Chapas' else '.1f')
            )
            totals_turno = alt.Chart(df_turno).mark_text(
                align='center', baseline='bottom', dy=-5, fontSize=18, fontWeight='bold', color='white', stroke='black', strokeWidth=0.5
            ).encode(
                x=alt.X(f'{c_tr}:N'),
                y=alt.Y(f'sum({col_valor}):Q'),
                text=alt.Text(f'sum({col_valor}):Q', format='.0f' if metrica_board == 'Chapas' else '.1f')
            )
            st.altair_chart((bars_turno + text_turno + totals_turno).properties(height=450), width="stretch")

        # ----------------- SEÇÃO: ANÁLISE DE PARADAS E OCIOSIDADE -----------------
        st.write("---")
        st.subheader("⏹️ Análise de Paradas e Ociosidade de Máquinas")
        st.markdown("Veja as principais causas de inatividade e ociosidade das máquinas registradas nos apontamentos.")
        
        try:
            ap_file = dm._get_apontamento_file()
            sheet_p = dm._get_sheet("SHEET_AP_PARADAS")
            df_paradas = pd.read_excel(ap_file, sheet_name=sheet_p)
            df_paradas.columns = [str(c).strip().upper() for c in df_paradas.columns]
            
            if df_paradas.empty:
                st.info("ℹ️ Não foram encontrados registros de paradas.")
            else:
                # 1. Definir função de parsing de tempo em minutos
                def parse_downtime_minutes(row):
                    t_val = row.get("TEMPO")
                    if pd.notna(t_val) and str(t_val).strip() != "":
                        t_str = str(t_val).strip()
                        if ":" in t_str:
                            try:
                                parts = t_str.split(":")
                                return int(parts[0]) * 60 + int(parts[1])
                            except:
                                pass
                        try:
                            num = float(t_str.replace(",", "."))
                            if num > 0:
                                if num < 1.0:
                                    return int(round(num * 24 * 60))
                                return int(round(num))
                        except:
                            pass
                    
                    # Fallback com HORA_INICIO e HORA_FIM
                    def parse_time_helper(t_str):
                        if not t_str or pd.isna(t_str): return None
                        import re
                        from datetime import time
                        t = re.sub(r'\D', '', str(t_str))
                        if len(t) == 4:
                            try: return time(int(t[:2]), int(t[2:]))
                            except: return None
                        elif len(t) == 3:
                            try: return time(int(t[:1]), int(t[1:]))
                            except: return None
                        return None

                    h_ini = parse_time_helper(row.get("HORA_INICIO"))
                    h_fim = parse_time_helper(row.get("HORA_FIM"))
                    
                    if h_ini and h_fim:
                        # Tenta usar datas se disponíveis
                        d_ini = pd.to_datetime(row.get("DATA_INICIO"), errors='coerce', dayfirst=True)
                        d_fim = pd.to_datetime(row.get("DATA_FIM"), errors='coerce', dayfirst=True)
                        
                        if pd.notna(d_ini) and pd.notna(d_fim):
                            try:
                                dt_ini = datetime.combine(d_ini.date(), h_ini)
                                dt_fim = datetime.combine(d_fim.date(), h_fim)
                                diff = (dt_fim - dt_ini).total_seconds() / 60
                                if diff > 0: return int(round(diff))
                            except:
                                pass
                        # Fallback mesmo dia
                        try:
                            dt_ini = datetime.combine(datetime.today().date(), h_ini)
                            dt_fim = datetime.combine(datetime.today().date(), h_fim)
                            diff = (dt_fim - dt_ini).total_seconds() / 60
                            if diff < 0:
                                diff += 24 * 60
                            if diff > 0: return int(round(diff))
                        except:
                            pass
                    return 0

                # Aplicar parsing de tempo
                df_paradas["TEMPO"] = df_paradas.apply(parse_downtime_minutes, axis=1)

                # 2. Fazer o merge com a base bruta (df_raw_an) primeiro, para obter os setores de todas as paradas
                # independentemente do filtro de data/máquinas geral do topo!
                col_id_ap = find_col_an("ID") or "ID"
                c_st_effective = c_st or "SETOR"
                
                # Criar um subconjunto de df_raw_an contendo as colunas necessárias para calcular DIA_PROD e fazer merge
                cols_to_use = [col_id_ap]
                if c_st_effective and c_st_effective in df_raw_an.columns:
                    cols_to_use.append(c_st_effective)
                
                # Adicionar colunas de data/hora para o cálculo do dia de produção
                for col_name in [c_dt, c_dt_ini, c_hr, c_hr_ini]:
                    if col_name and col_name in df_raw_an.columns and col_name not in cols_to_use:
                        cols_to_use.append(col_name)
                        
                df_raw_subset = df_raw_an[cols_to_use].copy()
                
                # Converter as colunas de data para datetime na cópia
                if c_dt and c_dt in df_raw_subset.columns:
                    df_raw_subset[c_dt] = pd.to_datetime(df_raw_subset[c_dt], errors='coerce', dayfirst=True)
                if c_dt_ini and c_dt_ini in df_raw_subset.columns and c_dt_ini != c_dt:
                    df_raw_subset[c_dt_ini] = pd.to_datetime(df_raw_subset[c_dt_ini], errors='coerce', dayfirst=True)
                    
                # Aplicar get_dia_producao com segurança!
                df_raw_subset["DIA_PROD"] = df_raw_subset.apply(get_dia_producao, axis=1)
                
                # Manter apenas as colunas necessárias para o merge final
                df_raw_an_subset = df_raw_subset[[col_id_ap, c_st_effective, "DIA_PROD"]].copy()
                df_raw_an_subset[col_id_ap] = pd.to_numeric(df_raw_an_subset[col_id_ap], errors='coerce')
                df_paradas["ID_APONTAMENTO"] = pd.to_numeric(df_paradas["ID_APONTAMENTO"], errors='coerce')
                
                df_paradas_m = df_paradas.merge(df_raw_an_subset, left_on="ID_APONTAMENTO", right_on=col_id_ap, how="inner")
                
                if df_paradas_m.empty:
                    st.info("ℹ️ Não foram encontradas correspondências de paradas com os apontamentos.")
                else:
                    # 3. Filtrar pelo mesmo período de data selecionado
                    if periodo == "Últimos 7 dias" and dias_unicos:
                        df_paradas_m = df_paradas_m[df_paradas_m["DIA_PROD"] >= min(dias_unicos)]
                    elif periodo == "Últimos 30 dias" and dias_unicos:
                        df_paradas_m = df_paradas_m[df_paradas_m["DIA_PROD"] >= min(dias_unicos)]
                    elif periodo == "Mês Atual":
                        df_paradas_m = df_paradas_m[pd.to_datetime(df_paradas_m["DIA_PROD"]).dt.month == hoje.month]
                    elif periodo == "Personalizado":
                        df_paradas_m = df_paradas_m[(df_paradas_m["DIA_PROD"] >= data_ini_custom) & (df_paradas_m["DIA_PROD"] <= data_fim_custom)]
                    
                    if df_paradas_m.empty:
                        st.info("ℹ️ Não há registros de paradas no período selecionado.")
                    else:
                        # --- CONTROLES DE MANIPULAÇÃO LOCAL ---
                        st.markdown("#### 🛠️ Controles de Filtros Locais para a Análise de Paradas")
                        
                        col_local1, col_local2 = st.columns([1, 1])
                        
                        # 1. Filtro local de Máquinas
                        maquinas_parada_disponiveis = sorted([str(x) for x in df_paradas_m[c_st].unique() if str(x) != "" and str(x) != "nan"])
                        with col_local1:
                            maquinas_parada_sel = st.multiselect(
                                "Filtrar Máquinas nas Paradas",
                                options=maquinas_parada_disponiveis,
                                default=maquinas_parada_disponiveis,
                                key="multiselect_maquinas_paradas"
                            )
                        
                        # 2. Filtro local de Motivos de Parada
                        motivos_disponiveis = sorted([str(x) for x in df_paradas_m["MOTIVO"].unique() if str(x) != "" and str(x) != "nan"])
                        default_motivos = [m for m in motivos_disponiveis if "ALMOÇO" not in m.upper()]
                        if not default_motivos: default_motivos = motivos_disponiveis # Fallback se só tiver almoço
                        
                        with col_local2:
                            motivos_sel = st.multiselect(
                                "Filtrar Motivos de Parada (Almoço removido por padrão)",
                                options=motivos_disponiveis,
                                default=default_motivos,
                                key="multiselect_motivos_paradas"
                            )
                            
                        # Segunda linha de controles: Unidade de tempo + Custos de Setores
                        col_c1, col_c2 = st.columns([1, 2])
                        with col_c1:
                            unidade_tempo = st.radio("Unidade do Gráfico de Motivos", ["Minutos", "Horas"], horizontal=True, key="radio_unidade_tempo")
                        
                        with col_c2:
                            # 3. Custo por Hora Ociosa por Setor
                            # Ler o dicionário atual do config.json
                            cfg_atual = dm.get_config()
                            custos_cfg = cfg_atual.get("CUSTOS_SETORES", {})
                            
                            # Obter todos os setores presentes na base de paradas para garantir que aparecem
                            setores_parada = sorted([str(x) for x in df_paradas_m[c_st].unique() if str(x) != "" and str(x) != "nan"])
                            
                            # Construir dataframe para edição
                            custos_lista = []
                            for s in setores_parada:
                                custos_lista.append({
                                    "Setor": s,
                                    "Custo Hora (R$/h)": float(custos_cfg.get(s, 150.0))
                                })
                            df_custos = pd.DataFrame(custos_lista)
                            
                            st.markdown("💰 **Custos Ociosos por Setor (R$/h)**")
                            df_custos_edit = st.data_editor(
                                df_custos,
                                column_config={
                                    "Setor": st.column_config.TextColumn("Setor / Máquina", disabled=True),
                                    "Custo Hora (R$/h)": st.column_config.NumberColumn("Custo Hora (R$)", min_value=0.0, step=10.0, format="R$ %.2f")
                                },
                                hide_index=True,
                                width="stretch",
                                key="data_editor_custos_setores"
                            )
                            
                            # Salvar alterações de custos no config.json se mudou
                            novos_custos = dict(zip(df_custos_edit["Setor"], df_custos_edit["Custo Hora (R$/h)"]))
                            if novos_custos != custos_cfg:
                                cfg_atual["CUSTOS_SETORES"] = novos_custos
                                dm.save_config(cfg_atual)
                                st.rerun()
                        
                        # Aplicar os filtros locais
                        df_paradas_filtrado = df_paradas_m.copy()
                        if maquinas_parada_sel:
                            df_paradas_filtrado = df_paradas_filtrado[df_paradas_filtrado[c_st].isin(maquinas_parada_sel)]
                        if motivos_sel:
                            df_paradas_filtrado = df_paradas_filtrado[df_paradas_filtrado["MOTIVO"].isin(motivos_sel)]
                            
                        if df_paradas_filtrado.empty:
                            st.warning("⚠️ Nenhum registro de parada atende aos filtros locais selecionados.")
                        else:
                            # 1. Função de formatação para HH:MM
                            def format_to_hhmm(minutes):
                                if pd.isna(minutes) or minutes <= 0:
                                    return "00:00"
                                total_mins = int(round(minutes))
                                h = total_mins // 60
                                m = total_mins % 60
                                return f"{h:02d}:{m:02d}"

                            # 2. Calcular prejuízo financeiro individualizado por linha PRIMEIRO de tudo!
                            custos_map = novos_custos
                            def calc_prejuizo_linha(row):
                                setor_row = str(row.get(c_st))
                                custo_s = custos_map.get(setor_row, 150.0)
                                tempo_min = float(row.get("TEMPO", 0))
                                return (tempo_min / 60.0) * custo_s
                                
                            df_paradas_filtrado["PREJUIZO"] = df_paradas_filtrado.apply(calc_prejuizo_linha, axis=1)
                            prejuizo_estimado = df_paradas_filtrado["PREJUIZO"].sum()

                            # --- MAPEAMENTO PADRONIZADO E CLASSIFICAÇÃO DE PARADAS (FAROL LEAN) ---
                            mapa_paradas = dm.get_mapeamento_paradas()
                            
                            # Carrega também o cadastro de paradas oficial editável (TIPO_PARADAS do DB.xlsm)
                            df_tp_db = dm.get_tipo_paradas()
                            tipo_paradas_db_map = {}
                            if not df_tp_db.empty and "MOTIVO" in df_tp_db.columns and "TIPO_PARADA" in df_tp_db.columns:
                                tipo_paradas_db_map = dict(zip(df_tp_db["MOTIVO"].astype(str).str.strip().str.upper(), df_tp_db["TIPO_PARADA"].astype(str).str.strip()))
                            
                            def map_motivo_padrao(row):
                                raw_m = str(row.get("MOTIVO", "")).strip().upper()
                                if raw_m in mapa_paradas:
                                    return mapa_paradas[raw_m]["RESUMIDO"]
                                return raw_m
                                
                            def map_tipo_parada(row):
                                raw_m = str(row.get("MOTIVO", "")).strip().upper()
                                # 1. Prioridade: cadastro oficial da Seção 4 (DB.xlsm)
                                if raw_m in tipo_paradas_db_map:
                                    return tipo_paradas_db_map[raw_m]
                                    
                                # 2. Mapeamento histórico da aba 'BASE PARADAS' (DB.xlsx)
                                if raw_m in mapa_paradas:
                                    return mapa_paradas[raw_m]["TIPO"]
                                    
                                # Fallback inteligente
                                raw_m_upper = raw_m.upper()
                                if "ALMOÇO" in raw_m_upper or "REFEIÇÃO" in raw_m_upper:
                                    return "Intervenção"
                                elif "MANUTENÇÃO" in raw_m_upper or "QUEBRA" in raw_m_upper or "MECÂNICA" in raw_m_upper or "MECANICA" in raw_m_upper or "ELÉTRICA" in raw_m_upper or "ELETRICA" in raw_m_upper:
                                    return "Crítica"
                                    
                                t_min = float(row.get("TEMPO", 0))
                                if t_min < 15:
                                    return "Operacional"
                                elif t_min <= 60:
                                    return "Intervenção"
                                else:
                                    return "Crítica"
                                    
                            df_paradas_filtrado["MOTIVO_ORIGINAL"] = df_paradas_filtrado["MOTIVO"]
                            df_paradas_filtrado["MOTIVO"] = df_paradas_filtrado.apply(map_motivo_padrao, axis=1)
                            df_paradas_filtrado["TIPO_PARADA"] = df_paradas_filtrado.apply(map_tipo_parada, axis=1)

                            # 3. Métricas Rápidas
                            tempo_tot_min = df_paradas_filtrado["TEMPO"].sum()
                            
                            # 4. Agrupamentos seguros (agora com PREJUIZO já calculado!)
                            c_motivos_totais = df_paradas_filtrado.groupby("MOTIVO")[["TEMPO", "PREJUIZO"]].sum().reset_index()
                            c_motivos_totais["TEMPO_HHMM"] = c_motivos_totais["TEMPO"].apply(format_to_hhmm)
                            c_motivos_totais = c_motivos_totais.sort_values("TEMPO", ascending=False)
                            motivo_top = c_motivos_totais.iloc[0]["MOTIVO"] if not c_motivos_totais.empty else "-"
                            motivos_ordenados = c_motivos_totais["MOTIVO"].tolist()
                             
                            c_motivos = df_paradas_filtrado.groupby(["MOTIVO", c_st])[["TEMPO", "PREJUIZO"]].sum().reset_index()
                            c_motivos["TEMPO_HHMM"] = c_motivos["TEMPO"].apply(format_to_hhmm)
                             
                            c_setores = df_paradas_filtrado.groupby(c_st)[["TEMPO", "PREJUIZO"]].sum().reset_index()
                            c_setores["TEMPO_HHMM"] = c_setores["TEMPO"].apply(format_to_hhmm)
                            setor_top = c_setores.sort_values("TEMPO", ascending=False).iloc[0][c_st] if not c_setores.empty else "-"
                            
                            st.divider()
                            
                            # Exibição de Indicadores Premium
                            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                            col_m1.metric("Tempo Total Parado (HH:MM)", format_to_hhmm(tempo_tot_min))
                            col_m2.metric("Principal Motivo", str(motivo_top).upper())
                            col_m3.metric("Máquina Mais Ociosa", str(setor_top).upper())
                            col_m4.metric("Prejuízo Estimado (R$)", f"R$ {prejuizo_estimado:,.2f}", delta=f"- {format_to_hhmm(tempo_tot_min)} de atividade", delta_color="inverse")
                            
                            st.divider()
                            col_g1, col_g2 = st.columns(2)
                            
                            with col_g1:
                                if unidade_tempo == "Horas":
                                    c_motivos["TEMPO_GRAFICO"] = c_motivos["TEMPO"] / 60
                                    c_motivos_totais["TEMPO_GRAFICO"] = c_motivos_totais["TEMPO"] / 60
                                    titulo_eixo_x = "Tempo Parado (Horas)"
                                    fmt_tooltip = ".1f"
                                else:
                                    c_motivos["TEMPO_GRAFICO"] = c_motivos["TEMPO"]
                                    c_motivos_totais["TEMPO_GRAFICO"] = c_motivos_totais["TEMPO"]
                                    titulo_eixo_x = "Tempo Parado (Minutos)"
                                    fmt_tooltip = ".0f"
                                    
                                # Ordenar totais no Pandas para passar ordem explícita ao Altair
                                c_motivos_totais_sorted = c_motivos_totais.sort_values("TEMPO_GRAFICO", ascending=False)
                                motivos_ordenados = c_motivos_totais_sorted["MOTIVO"].tolist()
                                    
                                st.markdown(f"**⏱️ Tempo de Parada por Motivo ({unidade_tempo})**")
                                # Camada 1: Barras empilhadas e coloridas por Máquina
                                bars_motivo = alt.Chart(c_motivos).mark_bar().encode(
                                    y=alt.Y("MOTIVO:N", title="Motivo", sort=motivos_ordenados, axis=alt.Axis(labelLimit=350, labelOverlap=False)),
                                    x=alt.X("TEMPO_GRAFICO:Q", title=titulo_eixo_x),
                                    color=alt.Color(f"{c_st}:N", title="Máquina", scale=alt.Scale(scheme="tableau10")),
                                    tooltip=[
                                        "MOTIVO", 
                                        alt.Tooltip(f"{c_st}:N", title="Máquina"),
                                        alt.Tooltip("TEMPO_GRAFICO:Q", format=fmt_tooltip, title=f"Tempo ({unidade_tempo})"),
                                        alt.Tooltip("TEMPO_HHMM:N", title="Tempo (HH:MM)"),
                                        alt.Tooltip("PREJUIZO:Q", format=",.2f", title="Prejuízo Est. (R$)")
                                    ]
                                )
                                
                                # Camada 2: Rótulos de texto com a duração TOTAL no fim da barra empilhada
                                text_motivo = alt.Chart(c_motivos_totais_sorted).mark_text(
                                    align="left",
                                    baseline="middle",
                                    dx=5,
                                    color="white"
                                ).encode(
                                    y=alt.Y("MOTIVO:N", sort=motivos_ordenados),
                                    x=alt.X("TEMPO_GRAFICO:Q"),
                                    text=alt.Text("TEMPO_HHMM:N")
                                )
                                
                                chart_p_motivo = (bars_motivo + text_motivo).properties(height=320)
                                st.altair_chart(chart_p_motivo, width="stretch")
                                
                            with col_g2:
                                st.markdown("**🏭 Tempo de Parada por Máquina/Setor (Horas)**")
                                c_setores["TEMPO_HORAS"] = c_setores["TEMPO"] / 60
                                
                                # Ordenar no Pandas para passar ordem explícita ao Altair (maiores tempos à esquerda)
                                c_setores_sorted = c_setores.sort_values("TEMPO_HORAS", ascending=False)
                                setores_ordenados = c_setores_sorted[c_st].tolist()
                                
                                # Camada 1: Barras
                                bars_setor = alt.Chart(c_setores_sorted).mark_bar().encode(
                                    x=alt.X(f"{c_st}:N", title="Máquina", sort=setores_ordenados),
                                    y=alt.Y("TEMPO_HORAS:Q", title="Tempo Parado (Horas)"),
                                    color=alt.Color(f"{c_st}:N", title="Máquina", scale=alt.Scale(scheme="tableau10"), legend=None),
                                    tooltip=[
                                        c_st, 
                                        alt.Tooltip("TEMPO_HORAS:Q", format=".1f", title="Horas"),
                                        alt.Tooltip("TEMPO_HHMM:N", title="Tempo (HH:MM)"),
                                        alt.Tooltip("PREJUIZO:Q", format=",.2f", title="Prejuízo Est. (R$)")
                                    ]
                                )
                                
                                # Camada 2: Texto em cima das barras (mostrando HH:MM)
                                text_setor = alt.Chart(c_setores_sorted).mark_text(
                                    align="center",
                                    baseline="bottom",
                                    dy=-5,
                                    color="white"
                                ).encode(
                                    x=alt.X(f"{c_st}:N", sort=setores_ordenados),
                                    y=alt.Y("TEMPO_HORAS:Q"),
                                    text=alt.Text("TEMPO_HHMM:N")
                                )
                                
                                chart_p_setor = (bars_setor + text_setor).properties(height=320)
                                st.altair_chart(chart_p_setor, width="stretch")
                                
                                # Tabela Detalhada Expansível
                                with st.expander("📋 Visualizar Ocorrências de Paradas Detalhadas"):
                                    st.markdown("Lista completa de todas as ocorrências de paradas com base nos filtros locais ativos (ordenadas da maior duração para a menor):")
                                    df_view_paradas = df_paradas_filtrado.sort_values("TEMPO", ascending=False).copy()
                                    df_view_paradas["TEMPO_HHMM"] = df_view_paradas["TEMPO"].apply(format_to_hhmm)
                                    
                                    # Conversão robusta de tipo para evitar ArrowTypeError no st.dataframe do Streamlit
                                    if "DATA_INICIO" in df_view_paradas.columns:
                                        df_view_paradas["DATA_INICIO"] = pd.to_datetime(df_view_paradas["DATA_INICIO"], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y").fillna("")
                                    if "HORA_INICIO" in df_view_paradas.columns:
                                        df_view_paradas["HORA_INICIO"] = df_view_paradas["HORA_INICIO"].apply(lambda t: t.strftime("%H:%M") if hasattr(t, "strftime") else str(t).replace("nan", "").strip() if pd.notna(t) else "")
                                    if "HORA_FIM" in df_view_paradas.columns:
                                        df_view_paradas["HORA_FIM"] = df_view_paradas["HORA_FIM"].apply(lambda t: t.strftime("%H:%M") if hasattr(t, "strftime") else str(t).replace("nan", "").strip() if pd.notna(t) else "")
                                        
                                    df_view_paradas = df_view_paradas[["ID_APONTAMENTO", c_st, "MOTIVO", "DATA_INICIO", "HORA_INICIO", "HORA_FIM", "TEMPO_HHMM", "PREJUIZO"]]
                                    df_view_paradas.columns = ["ID Apontamento", "Máquina", "Motivo", "Data", "Hora Início", "Hora Fim", "Duração (HH:MM)", "Prejuízo Estimado (R$)"]
                                    df_view_paradas["Prejuízo Estimado (R$)"] = df_view_paradas["Prejuízo Estimado (R$)"].apply(lambda v: f"R$ {v:,.2f}")
                                    st.dataframe(df_view_paradas, width="stretch", hide_index=True)
                                
                            # --- SEÇÃO: RELATÓRIO A3 LEAN ---
                            st.write("---")
                            st.subheader("🖨️ Relatório Formal A3 de Desempenho (Lean Manufacturing)")
                            st.markdown("Gere um relatório operacional de inatividade e impactos financeiros estruturado no tradicional **Layout A3 Landscape (420mm x 297mm)**, ideal para apresentações formais de diretoria, impressão física ou salvamento em PDF.")
                            
                            show_a3 = st.checkbox("📂 Visualizar Relatório A3 de Paradas e Ociosidade", value=False, key="chk_relatorio_a3")
                            show_bw = st.checkbox("🕶️ Modo Alto Contraste (Preto e Branco para Impressão)", value=False, key="chk_bw_print")
                            if show_a3:
                                body_class = "bw-contrast" if show_bw else ""
                                # Calcular faixa exata de datas cobertas no período filtrado
                                faixa_datas = ""
                                if not df_paradas_filtrado.empty and "DIA_PROD" in df_paradas_filtrado.columns:
                                    try:
                                        data_min_fmt = pd.to_datetime(df_paradas_filtrado["DIA_PROD"].min()).strftime("%d/%m/%Y")
                                        data_max_fmt = pd.to_datetime(df_paradas_filtrado["DIA_PROD"].max()).strftime("%d/%m/%Y")
                                        faixa_datas = f"({data_min_fmt} a {data_max_fmt})"
                                    except:
                                        pass
                                        
                                periodo_exibicao = f"{periodo} {faixa_datas}".strip()

                                # A. INDICADORES DE PRODUTIVIDADE OPERACIONAL (Aba 6)
                                prod_tot_m2 = df_an[c_m2].sum() if c_m2 in df_an.columns else 0.0
                                prod_tot_ch = df_an[c_ch].sum() if c_ch in df_an.columns else 0.0
                                
                                prod_refeito_m2 = df_an[df_an["REFEITO"]][c_m2].sum() if c_m2 in df_an.columns else 0.0
                                prod_refeito_ch = df_an[df_an["REFEITO"]][c_ch].sum() if c_ch in df_an.columns else 0.0
                                
                                prod_normal_m2 = prod_tot_m2 - prod_refeito_m2
                                prod_normal_ch = prod_tot_ch - prod_refeito_ch
                                
                                idx_refeito_m2 = (prod_refeito_m2 / prod_tot_m2 * 100) if prod_tot_m2 > 0 else 0.0
                                idx_refeito_ch = (prod_refeito_ch / prod_tot_ch * 100) if prod_tot_ch > 0 else 0.0

                                # Estilos saudáveis originais para o cartão de Chapas
                                ch_border = "#dcfce7"
                                ch_bg = "#f0fdf4"
                                ch_color = "#16a34a"
                                ch_alert_lbl = ""

                                # Estilos saudáveis originais para o cartão de M²
                                m2_border = "#dcfce7"
                                m2_bg = "#f0fdf4"
                                m2_color = "#16a34a"
                                m2_alert_lbl = ""

                                # B. MATRIZ DE PRODUTIVIDADE POR MÁQUINA & TURNO
                                prod_rows_html = ""
                                if not df_an.empty:
                                    df_prod_g = df_an.groupby([c_st, c_tr])[[c_ch, c_m2]].sum().reset_index()
                                    maquinas_prod = sorted(df_prod_g[c_st].unique())
                                    
                                    for maq in maquinas_prod:
                                        df_maq = df_prod_g[df_prod_g[c_st] == maq]
                                        
                                        # Diurno: contém 'D' ou '1' (ex: 'D', 'DIURNO', 'TURNO D')
                                        is_d = df_maq[c_tr].astype(str).str.upper().str.contains("D|DIUR|1", regex=True)
                                        ch_d = df_maq[is_d][c_ch].sum()
                                        m2_d = df_maq[is_d][c_m2].sum()
                                        
                                        # Noturno: contém 'N' ou '2' (ex: 'N', 'NOTU', 'TURNO N')
                                        is_n = df_maq[c_tr].astype(str).str.upper().str.contains("N|NOT|2", regex=True)
                                        ch_n = df_maq[is_n][c_ch].sum()
                                        m2_n = df_maq[is_n][c_m2].sum()
                                        
                                        # Outros turnos
                                        is_other = ~(is_d | is_n)
                                        ch_other = df_maq[is_other][c_ch].sum()
                                        m2_other = df_maq[is_other][c_m2].sum()
                                        
                                        ch_n_exib = ch_n + ch_other
                                        m2_n_exib = m2_n + m2_other
                                        
                                        ch_tot = ch_d + ch_n_exib
                                        m2_tot = m2_d + m2_n_exib
                                        
                                        prod_rows_html += f"""
                                        <tr>
                                            <td style='font-weight:600; color:#334155;'>{maq}</td>
                                            <td style='text-align:center; font-weight:500;'>{int(ch_d)} Ch</td>
                                            <td style='text-align:center; color:#475569;'>{m2_d:,.1f} m²</td>
                                            <td style='text-align:center; font-weight:500;'>{int(ch_n_exib)} Ch</td>
                                            <td style='text-align:center; color:#475569;'>{m2_n_exib:,.1f} m²</td>
                                            <td style='text-align:right; font-weight:700; color:#1e3a8a;'>{int(ch_tot)} Ch / {m2_tot:,.1f} m²</td>
                                        </tr>"""
                                        
                                    # Linha de Total Geral da Tabela de Produtividade
                                    tot_is_d = df_prod_g[c_tr].astype(str).str.upper().str.contains("D|DIUR|1", regex=True)
                                    tot_is_n = df_prod_g[c_tr].astype(str).str.upper().str.contains("N|NOT|2", regex=True)
                                    tot_is_other = ~(tot_is_d | tot_is_n)
                                    
                                    tot_ch_d = df_prod_g[tot_is_d][c_ch].sum()
                                    tot_m2_d = df_prod_g[tot_is_d][c_m2].sum()
                                    
                                    tot_ch_n = df_prod_g[tot_is_n][c_ch].sum() + df_prod_g[tot_is_other][c_ch].sum()
                                    tot_m2_n = df_prod_g[tot_is_n][c_m2].sum() + df_prod_g[tot_is_other][c_m2].sum()
                                    
                                    prod_rows_html += f"""
                                    <tr style='background:#f8fafc; font-weight:700; border-top:2px solid #cbd5e1;'>
                                         <td style='color:#1e3a8a;'>TOTAL GERAL</td>
                                         <td style='text-align:center;'>{int(tot_ch_d)} Ch</td>
                                         <td style='text-align:center;'>{tot_m2_d:,.1f} m²</td>
                                         <td style='text-align:center;'>{int(tot_ch_n)} Ch</td>
                                         <td style='text-align:center;'>{tot_m2_n:,.1f} m²</td>
                                         <td style='text-align:right; font-weight:700; color:#1e3a8a;'>{int(prod_tot_ch)} Ch / {prod_tot_m2:,.1f} m²</td>
                                    </tr>"""
                                else:
                                    prod_rows_html = "<tr><td colspan='6' style='text-align:center; color:#64748b;'>Nenhum registro de produção no período</td></tr>"

                                # F1. QUALIDADE POR MÁQUINA: Normal vs. Refeito com % índice de refeito
                                qualidade_rows_html = ""
                                if not df_an.empty:
                                    df_qual = df_an.groupby(c_st).apply(lambda g: pd.Series({
                                        "ch_normal": g[~g["REFEITO"]][c_ch].sum(),
                                        "ch_refeito": g[g["REFEITO"]][c_ch].sum(),
                                        "m2_normal": g[~g["REFEITO"]][c_m2].sum(),
                                        "m2_refeito": g[g["REFEITO"]][c_m2].sum(),
                                        "m2_total": g[c_m2].sum(),
                                        "ch_total": g[c_ch].sum(),
                                    })).reset_index()
                                    df_qual = df_qual.sort_values("m2_total", ascending=False)
                                    for _, rq in df_qual.iterrows():
                                        idx_ref = (rq["m2_refeito"] / rq["m2_total"] * 100) if rq["m2_total"] > 0 else 0
                                        bar_color = "#ef4444" if idx_ref > 10 else ("#f59e0b" if idx_ref > 3 else "#22c55e")
                                        pct_maq = (rq["m2_total"] / prod_tot_m2 * 100) if prod_tot_m2 > 0 else 0
                                        
                                        # Identificar se a produção refeita é maior que a produção normal
                                        refeito_maior = rq["m2_refeito"] > rq["m2_normal"] or rq["ch_refeito"] > rq["ch_normal"]
                                        
                                        # Background diferenciado e ícone de alerta se refeito > normal
                                        row_style = "background-color: #fef2f2 !important;" if refeito_maior else ""
                                        
                                        if refeito_maior:
                                            maq_label = f"🚨 <strong>{rq[c_st]}</strong> <span style='font-size:7.5px; color:#ef4444; font-weight:700; display:block; margin-top:1px;'>(Refeito > Normal)</span>"
                                        else:
                                            maq_label = rq[c_st]
                                        
                                        # Formatar cores dos valores de reprocesso (só vermelho se > 0)
                                        color_ch_ref = "color:#ef4444; font-weight:700;" if rq['ch_refeito'] > 0 else "color:#64748b;"
                                        color_m2_ref = "color:#ef4444; font-weight:700;" if rq['m2_refeito'] > 0 else "color:#64748b;"
                                        
                                        qualidade_rows_html += f"""
                                        <tr style='{row_style}'>
                                            <td style='font-weight:600; color:#334155;'>{maq_label}</td>
                                            <td style='text-align:center;'>{int(rq['ch_normal'])} Ch</td>
                                            <td style='text-align:center; {color_ch_ref}'>{int(rq['ch_refeito'])} Ch</td>
                                            <td style='text-align:center;'>{rq['m2_normal']:,.1f} m²</td>
                                            <td style='text-align:center; {color_m2_ref}'>{rq['m2_refeito']:,.1f} m²</td>
                                            <td style='text-align:right;'>
                                                <div style='display:flex; justify-content:space-between; font-size:8px; color:#64748b; margin-bottom:1px;'>
                                                    <span style='font-weight:700; color:{bar_color};'>{idx_ref:.1f}%</span>
                                                    <span>{pct_maq:.0f}% vol.</span>
                                                </div>
                                                <div class="progress-container">
                                                    <div class="progress-bar" style="width: {idx_ref:.1f}%; background: {bar_color};"></div>
                                                </div>
                                            </td>
                                        </tr>"""
                                else:
                                    qualidade_rows_html = "<tr><td colspan='6' style='text-align:center; color:#64748b;'>Sem dados</td></tr>"

                                # F2. ESTATÍSTICAS DIÁRIAS
                                num_dias_ativos = int(df_an["DIA_PROD"].nunique()) if not df_an.empty else 0
                                media_diaria_ch = (prod_tot_ch / num_dias_ativos) if num_dias_ativos > 0 else 0
                                media_diaria_m2 = (prod_tot_m2 / num_dias_ativos) if num_dias_ativos > 0 else 0

                                # F3. TOP PROCESSOS POR M² (de df_an usando c_pr)
                                processos_rows_html = ""
                                if not df_an.empty and c_pr:
                                    df_proc = df_an.groupby(c_pr)[[c_m2, c_ch]].sum().reset_index()
                                    df_proc = df_proc[df_proc[c_m2] > 0].sort_values(c_m2, ascending=False)
                                    total_m2_proc = df_proc[c_m2].sum()
                                    
                                    top_n_proc = 6
                                    df_proc_top = df_proc.head(top_n_proc)
                                    df_proc_others = df_proc.iloc[top_n_proc:]
                                    
                                    for _, rp in df_proc_top.iterrows():
                                        pct_p = (rp[c_m2] / total_m2_proc * 100) if total_m2_proc > 0 else 0
                                        processos_rows_html += f"""
                                        <tr>
                                            <td style='font-weight:600; color:#334155;'>{rp[c_pr]}</td>
                                            <td style='text-align:center;'>{int(rp[c_ch])} Ch</td>
                                            <td style='text-align:right;'>{rp[c_m2]:,.1f} m²</td>
                                            <td style='width:90px;'>
                                                <div style='display:flex; justify-content:space-between; font-size:8px; color:#64748b; margin-bottom:1px;'>
                                                    <span>{pct_p:.1f}%</span>
                                                </div>
                                                <div class="progress-container">
                                                    <div class="progress-bar" style="width: {pct_p}%; background: #3b82f6;"></div>
                                                </div>
                                            </td>
                                        </tr>"""
                                        
                                    if not df_proc_others.empty:
                                        others_m2 = df_proc_others[c_m2].sum()
                                        others_ch = df_proc_others[c_ch].sum()
                                        pct_others = (others_m2 / total_m2_proc * 100) if total_m2_proc > 0 else 0
                                        processos_rows_html += f"""
                                        <tr style='background:#f8fafc; font-style:italic;'>
                                            <td style='font-weight:600; color:#64748b;'>OUTROS PROCESSOS ({len(df_proc_others)} proc.)</td>
                                            <td style='text-align:center;'>{int(others_ch)} Ch</td>
                                            <td style='text-align:right;'>{others_m2:,.1f} m²</td>
                                            <td style='width:90px;'>
                                                <div style='display:flex; justify-content:space-between; font-size:8px; color:#64748b; margin-bottom:1px;'>
                                                    <span>{pct_others:.1f}%</span>
                                                </div>
                                                <div class="progress-container">
                                                    <div class="progress-bar" style="width: {pct_others}%; background: #94a3b8;"></div>
                                                </div>
                                            </td>
                                        </tr>"""
                                else:
                                    processos_rows_html = "<tr><td colspan='4' style='text-align:center; color:#64748b;'>Sem dados de processo</td></tr>"

                                # C. INDICADORES DE PARADAS & MOTIVOS - PARETO COMPACTO (Top 5 + Outros)
                                motivos_rows_html = ""
                                if not c_motivos_totais.empty:
                                    df_motivos_sorted = c_motivos_totais.sort_values("TEMPO", ascending=False)
                                    total_tempo_m = df_motivos_sorted["TEMPO"].sum()
                                    total_prej_m = df_motivos_sorted["PREJUIZO"].sum()
                                    
                                    top_n = 5
                                    df_top = df_motivos_sorted.head(top_n)
                                    df_others = df_motivos_sorted.tail(len(df_motivos_sorted) - top_n)
                                    
                                    for idx, r in df_top.iterrows():
                                        pct = (r["TEMPO"] / total_tempo_m * 100) if total_tempo_m > 0 else 0
                                        prej_fmt = f"R$ {r['PREJUIZO']:,.2f}"
                                        motivos_rows_html += f"""
                                        <tr>
                                            <td style='font-weight:600; color:#334155;'>{r['MOTIVO']}</td>
                                            <td style='text-align:center;'>{r['TEMPO_HHMM']}</td>
                                            <td style='text-align:right; font-weight:600; color:#1e3a8a;'>{prej_fmt}</td>
                                            <td style='width:100px;'>
                                                <div style='display:flex; justify-content:space-between; font-size:8.5px; color:#64748b; margin-bottom:1px;'>
                                                    <span>{pct:.1f}%</span>
                                                </div>
                                                <div class="progress-container">
                                                    <div class="progress-bar" style="width: {pct}%; background: #EF553B;"></div>
                                                </div>
                                            </td>
                                        </tr>"""
                                        
                                    if not df_others.empty:
                                        others_tempo = df_others["TEMPO"].sum()
                                        others_prej = df_others["PREJUIZO"].sum()
                                        others_tempo_hhmm = format_to_hhmm(others_tempo)
                                        pct_others = (others_tempo / total_tempo_m * 100) if total_tempo_m > 0 else 0
                                        
                                        motivos_rows_html += f"""
                                        <tr style='background:#f8fafc; font-style:italic;'>
                                            <td style='font-weight:600; color:#64748b;'>OUTRAS CAUSAS ({len(df_others)} motivos)</td>
                                            <td style='text-align:center;'>{others_tempo_hhmm}</td>
                                            <td style='text-align:right; font-weight:600; color:#64748b;'>R$ {others_prej:,.2f}</td>
                                            <td style='width:100px;'>
                                                <div style='display:flex; justify-content:space-between; font-size:8.5px; color:#64748b; margin-bottom:1px;'>
                                                    <span>{pct_others:.1f}%</span>
                                                </div>
                                                <div class="progress-container">
                                                    <div class="progress-bar" style="width: {pct_others}%; background: #94a3b8;"></div>
                                                </div>
                                            </td>
                                        </tr>"""
                                else:
                                    motivos_rows_html = "<tr><td colspan='4' style='text-align:center; color:#64748b;'>Nenhuma parada registrada</td></tr>"
                                    
                                # D. INDICADORES DE IMPACTO POR MÁQUINA & TAXA HORÁRIA
                                setores_rows_html = ""
                                total_tempo_s = c_setores["TEMPO"].sum()
                                for idx, r in c_setores.sort_values("TEMPO", ascending=False).iterrows():
                                    pct = (r["TEMPO"] / total_tempo_s * 100) if total_tempo_s > 0 else 0
                                    prej_fmt = f"R$ {r['PREJUIZO']:,.2f}"
                                    maq_name = r[c_st]
                                    c_val = float(custos_cfg.get(maq_name, 150.0))
                                    
                                    # Mapeamento dinâmico de cores Tableau10 para consistência
                                    cores_tableau = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                                    try:
                                        color_idx = setores_ordenados.index(maq_name) % len(cores_tableau)
                                        cores_s = cores_tableau[color_idx]
                                    except:
                                        cores_s = "#AB63FA"
                                        
                                    setores_rows_html += f"""
                                    <tr>
                                        <td style='font-weight:600; color:#334155;'>{maq_name}</td>
                                        <td style='text-align:center;'>{r['TEMPO_HHMM']}</td>
                                        <td style='text-align:center; font-weight:600; color:#1e3a8a;'>R$ {c_val:,.2f}/h</td>
                                        <td style='text-align:right; font-weight:600; color:#b91c1c;'>{prej_fmt}</td>
                                        <td style='width:100px;'>
                                            <div style='display:flex; justify-content:space-between; font-size:8.5px; color:#64748b; margin-bottom:1px;'>
                                                <span>{pct:.1f}%</span>
                                            </div>
                                            <div class="progress-container">
                                                <div class="progress-bar" style="width: {pct}%; background: {cores_s};"></div>
                                            </div>
                                        </td>
                                    </tr>"""
                                    
                                # F4. LEAN SEVERIDADE BREAKDOWN (Farol Lean: Operacional, Intervenção, Crítica)
                                lean_severity_cards_html = ""
                                if not df_paradas_filtrado.empty:
                                    total_tempo = df_paradas_filtrado["TEMPO"].sum()
                                    
                                    # Operacionais (Farol Verde)
                                    df_micro = df_paradas_filtrado[df_paradas_filtrado["TIPO_PARADA"] == "Operacional"]
                                    micro_cnt = len(df_micro)
                                    micro_sum = df_micro["TEMPO"].sum()
                                    micro_pct = (micro_sum / total_tempo * 100) if total_tempo > 0 else 0
                                    
                                    # Intervenções (Farol Amarelo)
                                    df_media = df_paradas_filtrado[df_paradas_filtrado["TIPO_PARADA"] == "Intervenção"]
                                    media_cnt = len(df_media)
                                    media_sum = df_media["TEMPO"].sum()
                                    media_pct = (media_sum / total_tempo * 100) if total_tempo > 0 else 0
                                    
                                    # Críticas (Farol Vermelho)
                                    df_longa = df_paradas_filtrado[df_paradas_filtrado["TIPO_PARADA"] == "Crítica"]
                                    longa_cnt = len(df_longa)
                                    longa_sum = df_longa["TEMPO"].sum()
                                    longa_pct = (longa_sum / total_tempo * 100) if total_tempo > 0 else 0
                                    
                                    severity_data = [
                                        {"title": "Operacionais", "cnt": micro_cnt, "time": micro_sum, "pct": micro_pct, "color": "#10b981", "bg": "#f0fdf4", "border": "#dcfce7", "tag": "Farol Verde"},
                                        {"title": "Intervenções", "cnt": media_cnt, "time": media_sum, "pct": media_pct, "color": "#f59e0b", "bg": "#fffbeb", "border": "#fef3c7", "tag": "Farol Amarelo"},
                                        {"title": "Críticas", "cnt": longa_cnt, "time": longa_sum, "pct": longa_pct, "color": "#ef4444", "bg": "#fef2f2", "border": "#fee2e2", "tag": "Farol Vermelho"}
                                    ]
                                    
                                    for card in severity_data:
                                        dur_str = format_to_hhmm(card["time"])
                                        lean_severity_cards_html += f"""
                                        <div style="border: 1px solid {card['border']}; background: {card['bg']}; padding: 2px 3px; border-radius: 6px; text-align: center;">
                                            <div style="font-size: 7px; font-weight: 700; color: #475569; text-transform: uppercase;">{card['title']}</div>
                                            <div style="font-size: 9.5px; font-weight: 700; color: {card['color']}; margin: 1px 0;">{card['cnt']} ocor.</div>
                                            <div style="font-size: 7px; font-weight: 600; color: #334155;">{dur_str} total</div>
                                            <div style="width: 100%; margin-top: 1px;">
                                                <div style="display:flex; justify-content:space-between; font-size:6px; color:#64748b; margin-bottom:1px;">
                                                    <span style="font-weight:700; color:{card['color']};">{card['pct']:.1f}%</span>
                                                    <span style="font-style:italic;">{card['tag']}</span>
                                                </div>
                                                <div class="progress-container" style="height: 1.5px; background: #e2e8f0; border-radius: 1px; overflow: hidden; margin-top: 1px;">
                                                    <div class="progress-bar" style="width: {card['pct']:.1f}%; height: 100%; background: {card['color']};"></div>
                                                </div>
                                            </div>
                                        </div>"""
                                else:
                                    lean_severity_cards_html = "<div style='grid-column: span 3; text-align:center; color:#64748b; font-size:9px;'>Sem dados de paradas</div>"

                                # E. OCORRÊNCIAS MAIS CRÍTICAS (Top 5 para layout compacto)
                                ocorrencias_rows_html = ""
                                df_top_5 = df_paradas_filtrado.sort_values("TEMPO", ascending=False).head(5)
                                for idx, r in df_top_5.iterrows():
                                    duration_str = format_to_hhmm(r["TEMPO"])
                                    prej_fmt = f"R$ {r['PREJUIZO']:,.2f}"
                                    dt_str = "-"
                                    if pd.notna(r.get("DATA_INICIO")):
                                        try:
                                            dt_str = pd.to_datetime(r["DATA_INICIO"], dayfirst=True).strftime("%d/%m/%Y")
                                        except:
                                            dt_str = str(r["DATA_INICIO"])
                                            
                                    ocorrencias_rows_html += f"""
                                    <tr>
                                        <td style='text-align:center;'>{int(r['ID_APONTAMENTO']) if pd.notna(r['ID_APONTAMENTO']) else '-'}</td>
                                        <td style="font-weight:600; color:#334155;">{r[c_st]}</td>
                                        <td>{r['MOTIVO_ORIGINAL']}</td>
                                        <td style='text-align:center;'>{dt_str}</td>
                                        <td style='text-align:center;'>{r['HORA_INICIO']} a {r['HORA_FIM']}</td>
                                        <td style="font-weight:700; color:#EF553B; text-align:center;">{duration_str}</td>
                                         <td style="font-weight:700; color:#1e3a8a; text-align:right;">{prej_fmt}</td>
                                     </tr>"""
                                    
                                html_a3 = f"""
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="utf-8">
                                    <title>Relatório A3 Costa Granitos - Performance Integrada</title>
                                    <style>
                                        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
                                        
                                        body {{
                                            font-family: 'Inter', sans-serif;
                                            background: #f8fafc;
                                            color: #1e293b;
                                            margin: 0;
                                            padding: 6px;
                                        }}
                                        
                                        .no-print {{
                                            text-align: right;
                                            margin-bottom: 8px;
                                            max-width: 1200px;
                                            margin-left: auto;
                                            margin-right: auto;
                                        }}
                                        
                                        .btn-print-a3 {{
                                            background: #1e3a8a;
                                            color: white;
                                            padding: 8px 18px;
                                            border: none;
                                            border-radius: 6px;
                                            cursor: pointer;
                                            font-weight: 600;
                                            font-size: 12px;
                                            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
                                            transition: background 0.2s;
                                        }}
                                        
                                        .btn-print-a3:hover {{
                                            background: #15295f;
                                        }}
                                        
                                        .a3-card {{
                                            background: #ffffff;
                                            width: 100%;
                                            max-width: 1200px;
                                            margin: 0 auto;
                                            border: 1px solid #cbd5e1;
                                            border-radius: 8px;
                                            padding: 10px 14px;
                                            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
                                            box-sizing: border-box;
                                        }}
                                        
                                        .header-a3 {{
                                            border-bottom: 2px solid #1e3a8a;
                                            padding-bottom: 3px;
                                            margin-bottom: 6px;
                                            display: flex;
                                            justify-content: space-between;
                                            align-items: flex-end;
                                        }}
                                        
                                        .header-a3 h1 {{
                                            margin: 0;
                                            font-size: 18px;
                                            color: #1e3a8a;
                                            font-weight: 700;
                                            letter-spacing: -0.5px;
                                        }}
                                        
                                        .header-a3 .subtitle {{
                                            margin: 2px 0 0 0;
                                            font-size: 10px;
                                            color: #64748b;
                                            text-transform: uppercase;
                                            font-weight: 600;
                                            letter-spacing: 1px;
                                        }}
                                        
                                        .meta-grid {{
                                            display: grid;
                                            grid-template-columns: repeat(4, 1fr);
                                            gap: 6px;
                                            background: #f8fafc;
                                            border: 1px solid #e2e8f0;
                                            padding: 4px 8px;
                                            border-radius: 6px;
                                            margin-top: 2px;
                                            font-size: 9.5px;
                                        }}
                                        
                                        .meta-item b {{
                                            color: #334155;
                                        }}
                                        
                                        .a3-columns {{
                                            display: grid;
                                            grid-template-columns: 1.15fr 1fr;
                                            gap: 10px;
                                            margin-bottom: 6px;
                                        }}
                                        
                                        .column-section {{
                                            border: 1px solid #cbd5e1;
                                            border-radius: 6px;
                                            padding: 8px;
                                            background: #ffffff;
                                        }}
                                        
                                        .section-title {{
                                            font-size: 11px;
                                            font-weight: 700;
                                            color: #1e3a8a;
                                            border-bottom: 1px solid #cbd5e1;
                                            padding-bottom: 2px;
                                            margin-top: 0;
                                            margin-bottom: 4px;
                                            text-transform: uppercase;
                                            letter-spacing: 0.5px;
                                        }}
                                        
                                        .kpi-grid {{
                                            display: grid;
                                            grid-template-columns: repeat(2, 1fr);
                                            gap: 6px;
                                            margin-bottom: 6px;
                                        }}
                                        
                                        .kpi-card {{
                                            border: 1px solid #e2e8f0;
                                            padding: 5px 8px;
                                            border-radius: 6px;
                                            background: #f8fafc;
                                            text-align: center;
                                        }}
                                        
                                        .kpi-val {{
                                            font-size: 15px;
                                            font-weight: 700;
                                            color: #1e3a8a;
                                            margin-top: 2px;
                                        }}
                                        
                                        .kpi-lbl {{
                                            font-size: 9px;
                                            color: #64748b;
                                            text-transform: uppercase;
                                            font-weight: 600;
                                        }}
                                        
                                        table.a3-table {{
                                            width: 100%;
                                            border-collapse: collapse;
                                            font-size: 9.5px;
                                            margin-bottom: 4px;
                                        }}
                                        
                                        table.a3-table th {{
                                            background: #f8fafc;
                                            color: #475569;
                                            font-weight: 700;
                                            padding: 3px 5px;
                                            border: 1px solid #cbd5e1;
                                            text-align: left;
                                        }}
                                        
                                        table.a3-table td {{
                                            padding: 2.5px 4px;
                                            border: 1px solid #cbd5e1;
                                            color: #334155;
                                        }}
                                        
                                        table.a3-table tr:nth-child(even) {{
                                            background: #f8fafc;
                                        }}
                                        
                                        .progress-container {{
                                            width: 100%;
                                            background: #f1f5f9;
                                            border: 1px solid #e2e8f0;
                                            border-radius: 3px;
                                            height: 4px;
                                            overflow: hidden;
                                        }}
                                        
                                        .progress-bar {{
                                            height: 100%;
                                            border-radius: 3px;
                                        }}
                                        
                                        .bottom-section {{
                                            border: 1px solid #cbd5e1;
                                            border-radius: 6px;
                                            padding: 8px 10px;
                                            background: #ffffff;
                                            margin-bottom: 0;
                                        }}
                                        
                                        .action-cell {{
                                            border: 1px dashed #94a3b8;
                                            padding: 6px 10px;
                                            border-radius: 4px;
                                            background: #f8fafc;
                                            font-size: 10.5px;
                                            color: #475569;
                                            line-height: 1.35;
                                        }}
                                        
                                        .signatures {{
                                            display: flex;
                                            justify-content: space-around;
                                            margin-top: 16px;
                                            padding-top: 8px;
                                            border-top: 1px solid #e2e8f0;
                                            font-size: 11px;
                                        }}
                                        
                                        .sig-line {{
                                            text-align: center;
                                            width: 260px;
                                        }}
                                        
                                        .sig-line .line {{
                                            border-top: 1px solid #475569;
                                            margin-top: 30px;
                                            margin-bottom: 4px;
                                        }}
                                        
                                        @media print {{
                                            body {{
                                                background: #ffffff;
                                                padding: 0 !important;
                                                margin: 0 !important;
                                            }}
                                            .no-print {{
                                                display: none !important;
                                            }}
                                            .a3-card {{
                                                border: none !important;
                                                box-shadow: none !important;
                                                padding: 0 !important;
                                                margin: 0 !important;
                                                max-width: 100% !important;
                                                height: 100% !important;
                                                page-break-inside: avoid !important;
                                                zoom: 100% !important;
                                            }}
                                            tr {{
                                                page-break-inside: avoid !important;
                                            }}
                                        }}
                                        
                                        @page {{
                                            
                                            margin: 0.8cm 1cm;
                                        }}
                                        
                                        /* Modo Alto Contraste (Preto e Branco) */
                                        .bw-contrast {{
                                            color: #000000 !important;
                                            background: #ffffff !important;
                                        }}
                                        .bw-contrast .a3-card {{
                                            border: 2.5px solid #000000 !important;
                                            background: #ffffff !important;
                                            box-shadow: none !important;
                                        }}
                                        .bw-contrast .header-a3 {{
                                            border-bottom: 3px solid #000000 !important;
                                        }}
                                        .bw-contrast .header-a3 h1,
                                        .bw-contrast .header-a3 span,
                                        .bw-contrast .header-a3 .subtitle {{
                                            color: #000000 !important;
                                        }}
                                        .bw-contrast .section-title {{
                                            color: #000000 !important;
                                            border-bottom: 2px solid #000000 !important;
                                            font-weight: 700 !important;
                                        }}
                                        .bw-contrast .kpi-card {{
                                            background: #ffffff !important;
                                            border: 2px solid #000000 !important;
                                        }}
                                        .bw-contrast .kpi-val, 
                                        .bw-contrast .kpi-lbl,
                                        .bw-contrast .kpi-card div {{
                                            color: #000000 !important;
                                        }}
                                        .bw-contrast table.a3-table th {{
                                            background: #cbd5e1 !important;
                                            color: #000000 !important;
                                            border: 1.5px solid #000000 !important;
                                            font-weight: 700 !important;
                                        }}
                                        .bw-contrast table.a3-table td {{
                                            color: #000000 !important;
                                            border: 1px solid #000000 !important;
                                            background: #ffffff !important;
                                        }}
                                        .bw-contrast table.a3-table tr {{
                                            background: #ffffff !important;
                                        }}
                                        .bw-contrast .progress-container {{
                                            background: #ffffff !important;
                                            border: 1px solid #000000 !important;
                                        }}
                                        .bw-contrast .progress-bar {{
                                            background: #000000 !important;
                                        }}
                                        .bw-contrast .meta-grid {{
                                            background: #ffffff !important;
                                            border: 1.5px solid #000000 !important;
                                        }}
                                        .bw-contrast .meta-item,
                                        .bw-contrast .meta-item b {{
                                            color: #000000 !important;
                                        }}
                                        .bw-contrast .bottom-section {{
                                            border: 2.5px solid #000000 !important;
                                            background: #ffffff !important;
                                        }}
                                        .bw-contrast .column-section {{
                                            border: 2px solid #000000 !important;
                                        }}
                                        .bw-contrast .action-cell {{
                                            background: #ffffff !important;
                                            border: 1.5px dashed #000000 !important;
                                            color: #000000 !important;
                                        }}
                                        .bw-contrast .sig-line .line {{
                                            border-top: 1.5px solid #000000 !important;
                                        }}
                                        .bw-contrast div[style*="border"] {{
                                            border: 1.5px solid #000000 !important;
                                        }}
                                        .bw-contrast div[style*="background"]:not(.progress-bar) {{
                                            background: #ffffff !important;
                                        }}
                                        .bw-contrast,
                                        .bw-contrast * {{
                                            color: #000000 !important;
                                        }}
                                    </style>
                                </head>
                                <body class="{body_class}">
                                    <div class="no-print">
                                        <button class="btn-print-a3" onclick="window.print()">🖨️ Imprimir / Salvar PDF (A3 / A4)</button>
                                    </div>
                                    
                                    <div class="a3-card">
                                        <div class="header-a3">
                                            <div>
                                                <h1>COSTA GRANITOS &mdash; PERFORMANCE OPERACIONAL E PRODUTIVIDADE</h1>
                                                <div class="subtitle">RELATÓRIO A3 LEAN: ANÁLISE INTEGRADA DE PRODUTIVIDADE, OCUPAÇÃO E PARADAS DE MÁQUINAS</div>
                                            </div>
                                            <div style="text-align: right; font-size: 11px; color:#475569; font-weight: 500;">
                                                Status: <span style="color:#16a34a; font-weight:700;">🟢 AUDITADO</span>
                                            </div>
                                        </div>
                                        
                                        <div class="meta-grid">
                                            <div class="meta-item"><b>Período Analisado:</b> {periodo_exibicao}</div>
                                            <div class="meta-item"><b>Data de Emissão:</b> {datetime.now().strftime("%d/%m/%Y %H:%M")}</div>
                                            <div class="meta-item"><b>Responsável:</b> PCP Costa Granitos</div>
                                            <div class="meta-item"><b>Filtros Ativos:</b> {len(maquinas_parada_sel)} Máqs. / {len(motivos_sel)} Motivos</div>
                                        </div>
                                        
                                        <div style="height: 6px;"></div>
                                        
                                        <div class="a3-columns">
                                            <!-- Coluna Esquerda: Produtividade -->
                                            <div class="column-section" style="display:flex; flex-direction:column; gap:6px;">
 
                                                <div class="section-title">1. Contexto &amp; Escopo Operacional</div>
                                                <p style="font-size: 9.5px; margin: 0 0 2px 0; line-height: 1.35; color: #475569; text-align:justify;">
                                                    Este A3 consolida a performance integrada da fábrica. O período cobriu <strong>{num_dias_ativos} dia(s) ativo(s)</strong> de produção, com produtividade média de <strong>{media_diaria_ch:,.0f} chapas/dia</strong> e <strong>{media_diaria_m2:,.1f} m²/dia</strong>. Os dados discriminam produção normal e refeita por máquina e turno, correlacionando com o índice de qualidade (% refeito).
                                                </p>
 
                                                <div class="section-title">2. KPIs Globais de Produtividade</div>
                                                <div class="kpi-grid">
                                                    <div class="kpi-card" style="border-color:{ch_border}; background:{ch_bg};">
                                                        <div class="kpi-lbl" style="color:{ch_color}; font-weight: 700;">Total Chapas{ch_alert_lbl}</div>
                                                        <div class="kpi-val" style="color:{ch_color}; font-size: 16px;">{int(prod_tot_ch)}</div>
                                                        <div style="font-size:9px; color:#64748b; margin-top:2px;">Normal: {int(prod_normal_ch)} | Refeito: {int(prod_refeito_ch)}</div>
                                                    </div>
                                                    <div class="kpi-card" style="border-color:{m2_border}; background:{m2_bg};">
                                                        <div class="kpi-lbl" style="color:{m2_color}; font-weight: 700;">Total M²{m2_alert_lbl}</div>
                                                        <div class="kpi-val" style="color:{m2_color}; font-size: 16px;">{prod_tot_m2:,.1f}</div>
                                                        <div style="font-size:9px; color:#64748b; margin-top:2px;">Normal: {prod_normal_m2:,.1f} | Ref: {prod_refeito_m2:,.1f}</div>
                                                    </div>
                                                    <div class="kpi-card">
                                                        <div class="kpi-lbl">Média Diária</div>
                                                        <div class="kpi-val" style="font-size:14px; padding-top:3px;">{media_diaria_ch:,.0f} Ch</div>
                                                        <div style="font-size:9px; color:#64748b; margin-top:2px;">{media_diaria_m2:,.1f} m²/dia — {num_dias_ativos} dias</div>
                                                    </div>
                                                    <div class="kpi-card">
                                                        <div class="kpi-lbl">Índice Global Refeito</div>
                                                        <div class="kpi-val" style="font-size:14px; padding-top:2px; color:{'#ef4444' if idx_refeito_m2 > 5 else '#f59e0b' if idx_refeito_m2 > 2 else '#16a34a'};">{idx_refeito_m2:.1f}%</div>
                                                        <div style="font-size:9px; color:#64748b; margin-top:2px;">em M² processados</div>
                                                    </div>
                                                </div>
 
                                                <div class="section-title">3. Matriz de Produtividade por Máquina e Turno</div>
                                                <table class="a3-table">
                                                    <thead>
                                                        <tr>
                                                            <th>Máquina / Setor</th>
                                                            <th style='text-align:center;'>Diurno (Ch)</th>
                                                            <th style='text-align:center;'>Diurno (M²)</th>
                                                            <th style='text-align:center;'>Noturno (Ch)</th>
                                                            <th style='text-align:center;'>Noturno (M²)</th>
                                                            <th style='text-align:right;'>Total (Ch / M²)</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>{prod_rows_html}</tbody>
                                                </table>
 
                                                <div class="section-title">4. Qualidade por Máquina — Normal vs. Refeito</div>
                                                <table class="a3-table">
                                                    <thead>
                                                        <tr>
                                                            <th>Máquina</th>
                                                            <th style='text-align:center;'>Normal (Ch)</th>
                                                            <th style='text-align:center;'>Refeito (Ch)</th>
                                                            <th style='text-align:center;'>Normal (M²)</th>
                                                            <th style='text-align:center;'>Refeito (M²)</th>
                                                            <th style='text-align:right;'>% Refeito / Vol.</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>{qualidade_rows_html}</tbody>
                                                </table>
 
                                                <div class="section-title">5. Distribuição por Processo Produtivo</div>
                                                <table class="a3-table">
                                                    <thead>
                                                        <tr>
                                                            <th>Processo</th>
                                                            <th style='text-align:center;'>Chapas</th>
                                                            <th style='text-align:right;'>M²</th>
                                                            <th style='text-align:right;'>Participação</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>{processos_rows_html}</tbody>
                                                </table>
 
                                            </div>
                                            
                                            <!-- Coluna Direita: Paradas & Inatividade -->
                                            <div class="column-section" style="display:flex; flex-direction:column; justify-content:space-between; gap:6px;">
                                                <div>
                                                    <div class="section-title">6. KPIs de Paradas &amp; Ociosidade</div>
                                                    <div class="kpi-grid">
                                                        <div class="kpi-card" style="border-color:#fee2e2; background:#fef2f2;">
                                                            <div class="kpi-lbl" style="color:#ef4444;">Tempo Total Ocioso</div>
                                                            <div class="kpi-val" style="color:#ef4444; font-size:15px;">{format_to_hhmm(tempo_tot_min)}</div>
                                                        </div>
                                                        <div class="kpi-card" style="border-color:#fee2e2; background:#fef2f2;">
                                                            <div class="kpi-lbl" style="color:#ef4444;">Prejuízo Estimado</div>
                                                            <div class="kpi-val" style="color:#ef4444; font-size:15px;">R$ {prejuizo_estimado:,.2f}</div>
                                                        </div>
                                                        <div class="kpi-card">
                                                            <div class="kpi-lbl">Principal Causa</div>
                                                            <div class="kpi-val" style="font-size:11px; padding-top:4px;">{str(motivo_top).upper()}</div>
                                                        </div>
                                                        <div class="kpi-card">
                                                            <div class="kpi-lbl">Máquina Mais Ociosa</div>
                                                            <div class="kpi-val" style="font-size:11px; padding-top:4px;">{str(setor_top).upper()}</div>
                                                        </div>
                                                    </div>
                                                </div>
                                                
                                                <div>
                                                    <div class="section-title">7. Pareto de Paradas por Motivo</div>
                                                    <table class="a3-table">
                                                        <thead>
                                                            <tr>
                                                                <th>Motivo da Parada</th>
                                                                <th style='text-align:center; width:60px;'>Duração</th>
                                                                <th style='text-align:right; width:100px;'>Custo (R$)</th>
                                                                <th style='text-align:left; width:110px;'>Representação (%)</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {motivos_rows_html}
                                                        </tbody>
                                                    </table>
                                                </div>
                                                
                                                <div>
                                                    <div class="section-title">8. Ocupação de Máquina &amp; Impacto Financeiro Mapeado</div>
                                                    <table class="a3-table">
                                                        <thead>
                                                            <tr>
                                                                <th>Máquina / Setor</th>
                                                                <th style='text-align:center;'>Duração</th>
                                                                <th style='text-align:center;'>Taxa (R$/h)</th>
                                                                <th style='text-align:right;'>Prejuízo</th>
                                                                <th style='text-align:left; width:110px;'>Representação (%)</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {setores_rows_html}
                                                        </tbody>
                                                    </table>
                                                </div>
                                                
                                                <div>
                                                    <div class="section-title">9. Análise Lean: Severidade e Frequência das Paradas</div>
                                                    <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 6px; margin-top: 2px;">
                                                        {lean_severity_cards_html}
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                        
                                        <!-- Seção Inferior: Ocorrências Críticas (Top 5) -->
                                        <div class="bottom-section" style="margin-bottom:0; padding: 6px 10px;">
                                            <div class="section-title" style="margin-bottom:4px;">10. Ocorrências Mais Críticas de Paradas (Top 5 por Duração)</div>
                                            <table class="a3-table" style="font-size: 9px;">
                                                <thead>
                                                    <tr>
                                                        <th style='text-align:center; width:60px;'>ID Apont.</th>
                                                        <th>Máquina</th>
                                                        <th>Motivo Real da Parada</th>
                                                        <th style='text-align:center; width:80px;'>Data</th>
                                                        <th style='text-align:center; width:150px;'>Intervalo de Tempo</th>
                                                        <th style='text-align:center; width:80px;'>Duração (HH:MM)</th>
                                                        <th style='text-align:right; width:120px;'>Prejuízo Individual</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {ocorrencias_rows_html}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>

                                </body>
                                </html>
                                """
                                st.components.v1.html(html_a3, height=980, scrolling=True)
                                st.info("💡 **Dica de Impressão:** Ao clicar no botão acima, a tela de impressão do navegador será aberta diretamente. Selecione a impressora desejada (como **Microsoft Print to PDF** ou **Salvar como PDF**), mude a orientação (Layout) para **Paisagem (Landscape)**, defina o tamanho do papel como **A3** (ou **A4** com a opção **Ajustar à página** ativa) e ajuste a escala para preencher a folha perfeitamente!")

                                
        except Exception as e:
            st.warning(f"Não foi possível carregar a análise detalhada de paradas. (Motivo: {e})")


# ----------------- ABA 7: OPÇÕES GERAIS -----------------
with tab_config:
    render_opcoes_gerais(dm.get_config(), df_base)

