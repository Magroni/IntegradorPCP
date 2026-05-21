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

st.set_page_config(page_title="PCP Costa Granitos", layout="wide", page_icon="📊")

st.title("📊 Gerenciador de Programação WIP")

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

# Load data
df = dm.get_data()
df_base = dm.get_base_dados()

if df.empty:
    aba_p = dm._get_sheet("SHEET_PROGRAMACAO")
    st.error(f"Não foi possível carregar os dados. Verifique se o arquivo está acessível e possui a aba '{aba_p}'.")
    st.info("💡 Você pode ajustar o nome do arquivo e da aba na guia 'Opções Gerais' (se conseguir acessar) ou no arquivo config.json.")
    st.stop()

df = df.fillna("")
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
# --- PREPARAÇÃO DE DADOS GLOBAL ---
df_aberto = df[~df.apply(is_finished, axis=1)].copy()
setores_list = sorted([str(x) for x in df["SETOR"].unique() if str(x) != "" and str(x) != "nan"])

tab_block, tab_view, tab_machine, tab_apontamento, tab_export, tab_analises, tab_config = st.tabs([
    "🛠️ Adicionar / Editar Bloco", 
    "👁️ Base de Dados", 
    "🗓️ Janela de Programações",
    "✅ Apontamento",
    "🖨️ Exportação",
    "📈 Análises e Indicadores",
    "⚙️ Opções Gerais"
])

# ----------------- ABA 1: ADICIONAR / EDITAR BLOCO -----------------
with tab_block:
    st.header("Adicionar Novo / Editar Bloco Existente")
    
    modo = st.radio("Selecione a ação:", ["Criar Novo Bloco", "Editar Bloco Existente"], horizontal=True)
    st.divider()
    
    if "modo_atual" not in st.session_state or st.session_state["modo_atual"] != modo:
        st.session_state["roteiro_atual"] = []
        st.session_state["modo_atual"] = modo
        st.session_state["bloco_carregado"] = False
        st.session_state["bloco_edit_id"] = ""
        st.session_state["mat_edit"] = ""
        st.session_state["dem_edit"] = ""
        st.session_state["qtd_edit"] = 0
        st.session_state["vol_edit"] = 0.0
        
    if modo == "Criar Novo Bloco":
        # --- BUSCA AUTOMÁTICA DE INFO PARA NOVO BLOCO ---
        bloco = st.text_input("Número do Bloco*", key="c_blo")
        
        if bloco and bloco != st.session_state.get("c_last_bloco"):
            info = dm.get_bloco_info(bloco)
            if info:
                st.session_state["c_mat_val"] = info["MATERIAL"]
                st.session_state["c_source"] = info["SOURCE"]
                st.session_state["c_found"] = True
                st.session_state["c_last_bloco"] = bloco
                st.rerun()
            else:
                st.session_state["c_mat_val"] = ""
                st.session_state["c_source"] = ""
                st.session_state["c_found"] = False
                st.session_state["c_last_bloco"] = bloco
                st.rerun()
        elif not bloco:
            st.session_state["c_mat_val"] = ""
            st.session_state["c_source"] = ""
            st.session_state["c_found"] = None
            st.session_state["c_last_bloco"] = ""

        if bloco:
            if st.session_state.get("c_found"):
                st.success(f"✅ Bloco já cadastrado na base: **{st.session_state.get('c_source')}**")
            elif st.session_state.get("c_found") is False:
                st.info("ℹ️ Bloco novo (não encontrado nas bases de Blocos/Chapas).")

        c1, c2, c3 = st.columns(3)
        with c1:
            material = st.text_input("Material*", value=st.session_state.get("c_mat_val", ""), key="c_mat")
        with c2:
            demanda = st.text_input("Demanda", key="c_dem")
            qtd_chapas = st.number_input("Qtd. Chapas", min_value=0, step=1, key="c_qtd")
        with c3:
            volume_m2 = st.number_input("Volume M²", min_value=0.0, format="%.2f", key="c_vol")
            
        st.divider()
        st.subheader("Construtor de Roteiro")
        
        st.markdown("💡 **Roteiros Rápidos (Presets)**")
        col_p1, col_p2, col_p3, col_p4 = st.columns(4)
        def apply_preset(lista_proc):
            st.session_state["roteiro_atual"] = [{"PROCESSO": p, "SETOR": mapa_processos.get(p, ""), "OBSERVACAO": ""} for p in lista_proc]
            
        if col_p1.button("Simples (Polir Simples)", use_container_width=True): apply_preset(["SERRAR", "LEVIGAR 1", "POLIR SIMPLES", "RETOCAR"])
        if col_p2.button("Resinado 1x", use_container_width=True): apply_preset(["SERRAR", "LEVIGAR 1", "RESINAR 1", "POLIR RESINADO", "RETOCAR"])
        if col_p3.button("Resinado 2x", use_container_width=True): apply_preset(["SERRAR", "LEVIGAR 1", "RESINAR 1", "LEVIGAR 2", "RESINAR 2", "POLIR RESINADO", "RETOCAR"])
        if col_p4.button("Telado + Resinado 1x", use_container_width=True): apply_preset(["SERRAR", "TELAR/MANTAR", "LEVIGAR 1", "RESINAR 1", "POLIR RESINADO", "RETOCAR"])
            
        st.markdown("---")
        st.markdown("**Adicionar Etapas Manualmente:**")
        
        col_proc, col_setor, col_btn = st.columns([2, 2, 1])
        with col_proc:
            processo_sel = st.selectbox("Escolha um Processo", [""] + lista_processos + ["Outro (Digitar Novo)..."], key="c_psel")
            processo_final = st.text_input("Digite o Novo Processo", key="c_ptxt") if processo_sel == "Outro (Digitar Novo)..." else processo_sel
        with col_setor:
            setor_sugerido = mapa_processos.get(processo_final, "") if processo_final else ""
            setor_final = st.text_input("Setor Produtivo", value=setor_sugerido, key=f"c_setor_{processo_final}")
        with col_btn:
            st.write(""); st.write("")
            if st.button("➕ Adicionar Etapa", use_container_width=True, key="c_btn_add"):
                if processo_final and setor_final:
                    st.session_state["roteiro_atual"].append({"PROCESSO": processo_final, "SETOR": setor_final, "OBSERVACAO": ""})
                    st.rerun()
                else: st.error("Preencha Processo e Setor")
                    
        if st.session_state["roteiro_atual"]:
            st.markdown("### 📋 Roteiro Construído")
            for i, etapa in enumerate(st.session_state["roteiro_atual"]):
                st.info(f"**{i+1}. {etapa['PROCESSO']}** ➡️ Setor: {etapa['SETOR']}")
                
            if st.button("🗑️ Limpar Roteiro", key="c_limpar"):
                st.session_state["roteiro_atual"] = []
                st.rerun()
                
            st.divider()
            if st.button("✅ Salvar Bloco e Roteiro Completo", type="primary", key="c_salvar"):
                if not bloco or not material:
                    st.error("Preencha Material e Bloco antes de salvar.")
                else:
                    records = []
                    for p in st.session_state["roteiro_atual"]:
                        records.append({
                            "MATERIAL": material, "BLOCO": bloco, "QTD. CHAPAS": qtd_chapas, "VOLUME M²": volume_m2,
                            "DEMANDA": demanda, "PROCESSO": p["PROCESSO"], "SETOR": p["SETOR"], "STATUS PROCESSO": "NÃO REALIZADO",
                            "DATA": "", "DATA REALIZADA": "", "STATUS ATUAL": "Aguardando Produção",
                            "OBSERVAÇÃO": "", "OBSERVAÇÃO DE PRODUÇÃO": "", "ENTREGA?": ""
                        })
                    with st.spinner("Salvando no Excel..."):
                        if dm.add_records(records):
                            st.success(f"Bloco {bloco} cadastrado com {len(records)} etapas!")
                            st.session_state["roteiro_atual"] = []
                        else: st.error("Erro ao criar roteiro.")

    else: # MODO EDITAR
        with st.form("form_busca_bloco"):
            col_b, col_b_btn = st.columns([3, 1])
            with col_b:
                bloco_busca = st.text_input("Digite o número do Bloco para buscar:")
            with col_b_btn:
                st.write(""); st.write("")
                submit_busca = st.form_submit_button("🔍 Buscar Bloco", use_container_width=True)
                
            if submit_busca:
                if bloco_busca:
                    # Garantir que a busca seja flexível (removendo .0 de floats, espaços, etc)
                    bloco_limpo = str(bloco_busca).strip().upper()
                    df["BLOCO_STR"] = df["BLOCO"].astype(str).str.replace(".0", "", regex=False).str.strip().str.upper()
                    
                    df_b = df[df["BLOCO_STR"] == bloco_limpo].copy()
                    if df_b.empty:
                        st.warning("Bloco não encontrado.")
                        st.session_state["bloco_carregado"] = False
                    else:
                        st.session_state["bloco_carregado"] = True
                        st.session_state["bloco_edit_id"] = str(bloco_busca)
                        primeira = df_b.iloc[0]
                        
                        # Valores para o Session State (base)
                        mat_val = str(primeira.get("MATERIAL", ""))
                        dem_val = str(primeira.get("DEMANDA", ""))
                        try: qtd_val = int(pd.to_numeric(primeira.get("QTD. CHAPAS", 0)))
                        except: qtd_val = 0
                        try: vol_val = float(pd.to_numeric(primeira.get("VOLUME M²", 0.0)))
                        except: vol_val = 0.0
                        
                        st.session_state["mat_edit"] = mat_val
                        st.session_state["dem_edit"] = dem_val
                        st.session_state["qtd_edit"] = qtd_val
                        st.session_state["vol_edit"] = vol_val
                        
                        # FORÇA ATUALIZAÇÃO DOS WIDGETS (Keys)
                        st.session_state["e_mat"] = mat_val
                        st.session_state["e_dem"] = dem_val
                        st.session_state["e_qtd"] = qtd_val
                        st.session_state["e_vol"] = vol_val
                        
                        rot = []
                        for _, r in df_b.iterrows():
                            rot.append({
                                "PROCESSO": str(r.get("PROCESSO", "")),
                                "SETOR": str(r.get("SETOR", "")),
                                "OBSERVACAO": str(r.get("OBSERVACAO DE PRODUÇÃO", "")),
                                "STATUS PROCESSO": str(r.get("STATUS PROCESSO", "")),
                                "DATA": str(r.get("DATA", "")),
                                "DATA REALIZADA": str(r.get("DATA REALIZADA", ""))
                            })
                        st.session_state["roteiro_atual"] = rot
        
        if st.session_state.get("bloco_carregado"):
            st.success(f"Bloco {st.session_state['bloco_edit_id']} carregado!")
            c1, c2, c3 = st.columns(3)
            with c1:
                mat_edit = st.text_input("Material*", value=st.session_state["mat_edit"], key="e_mat")
            with c2:
                dem_edit = st.text_input("Demanda", value=st.session_state["dem_edit"], key="e_dem")
                qtd_edit = st.number_input("Qtd. Chapas", value=st.session_state["qtd_edit"], min_value=0, step=1, key="e_qtd")
            with c3:
                vol_edit = st.number_input("Volume M²", value=st.session_state["vol_edit"], min_value=0.0, format="%.2f", key="e_vol")
                
            st.divider()
            st.subheader("Roteiro do Bloco (Ficha de Produção)")
            
            novas_etapas = []
            
            def format_data_view(d_val):
                if pd.isna(d_val) or str(d_val).strip() in ["", "nan", "NaT"]: return "-"
                try:
                    if isinstance(d_val, pd.Timestamp): return d_val.strftime("%d/%m/%Y")
                    if "/" in str(d_val): return pd.to_datetime(str(d_val), format="%d/%m/%Y").strftime("%d/%m/%Y")
                    return pd.to_datetime(str(d_val)).strftime("%d/%m/%Y")
                except: return str(d_val)
            
            # --- VISUALIZAÇÃO EM ESTILO TABELA (DATA EDITOR) ---
            df_edit_roteiro = pd.DataFrame(st.session_state["roteiro_atual"])
            
            # Formatação para exibição
            df_view_rot = pd.DataFrame()
            df_view_rot["Ordem"] = range(1, len(df_edit_roteiro) + 1)
            df_view_rot["Processo"] = df_edit_roteiro["PROCESSO"]
            df_view_rot["Máquina"] = df_edit_roteiro["SETOR"]
            
            # Lógica de Fidelidade para o Status Visual
            def get_visual_status(row):
                if is_finished(row): return "✅ REALIZADO"
                st_txt = str(row.get("STATUS PROCESSO", "")).upper()
                if st_txt == "EM PROCESSO": return "⏳ EM PROCESSO"
                return "⚪ NÃO REALIZADO"
                
            df_view_rot["Status"] = df_edit_roteiro.apply(get_visual_status, axis=1)
            df_view_rot["Data Prog."] = df_edit_roteiro["DATA"].apply(format_data_view)
            df_view_rot["Data Realiz."] = df_edit_roteiro["DATA REALIZADA"].apply(format_data_view)
            df_view_rot["Observação"] = df_edit_roteiro["OBSERVACAO"]
            
            edited_df_rot = st.data_editor(
                df_view_rot,
                column_config={
                    "Ordem": st.column_config.NumberColumn("Nº", width="small", disabled=True),
                    "Processo": st.column_config.TextColumn("Processo", disabled=True),
                    "Máquina": st.column_config.SelectboxColumn("Máquina", options=sorted(setores_list), required=True),
                    "Status": st.column_config.TextColumn("Status (Fórmula)", disabled=True),
                    "Data Prog.": st.column_config.TextColumn("Data Prog.", disabled=True),
                    "Data Realiz.": st.column_config.TextColumn("Data Realiz.", disabled=True),
                    "Observação": st.column_config.TextColumn("Observação", width="large")
                },
                hide_index=True,
                use_container_width=True,
                key="editor_roteiro_bloco"
            )
            
            # Sincroniza as alterações de volta para o session_state
            novas_etapas = []
            for i, row in edited_df_rot.iterrows():
                orig = st.session_state["roteiro_atual"][i].copy()
                # Só permite alterar se não estiver realizado (ou permite tudo e o data_manager filtra)
                orig["SETOR"] = row["Máquina"]
                orig["OBSERVACAO"] = row["Observação"]
                novas_etapas.append(orig)
            
            st.session_state["roteiro_atual"] = novas_etapas
            
            # Botão de remover (ainda necessário para excluir linhas)
            if len(st.session_state["roteiro_atual"]) > 0:
                with st.expander("❌ Remover Etapas Pendentes"):
                    idx_remover = st.multiselect("Selecione as etapas para remover (apenas não realizadas):", 
                                                 range(1, len(st.session_state["roteiro_atual"]) + 1),
                                                 format_func=lambda x: f"{x}. {st.session_state['roteiro_atual'][x-1]['PROCESSO']}")
                    if st.button("Confirmar Remoção das Etapas Selecionadas"):
                        # Remove de trás para frente para não bagunçar os índices
                        for idx in sorted(idx_remover, reverse=True):
                            if not is_finished(st.session_state["roteiro_atual"][idx-1]):
                                st.session_state["roteiro_atual"].pop(idx-1)
                            else:
                                st.error(f"Não é possível remover a etapa {idx} porque ela já foi realizada.")
                        st.rerun()

            st.markdown("---")
            st.markdown("**➕ Adicionar Nova Etapa ao Final:**")
            col_eproc, col_eset, col_ebtn = st.columns([2, 2, 1])
            with col_eproc:
                eproc_sel = st.selectbox("Processo", [""] + lista_processos + ["Outro (Digitar Novo)..."], key="e_psel")
                eproc_final = st.text_input("Digite o Processo", key="e_ptxt") if eproc_sel == "Outro (Digitar Novo)..." else eproc_sel
            with col_eset:
                eset_sug = mapa_processos.get(eproc_final, "") if eproc_final else ""
                eset_final = st.text_input("Setor Produtivo", value=eset_sug, key=f"e_setor_{eproc_final}")
            with col_ebtn:
                st.write(""); st.write("")
                if st.button("Adicionar", use_container_width=True, key="e_btn_add"):
                    if eproc_final and eset_final:
                        st.session_state["roteiro_atual"].append({"PROCESSO": eproc_final, "SETOR": eset_final, "OBSERVACAO": ""})
                        st.rerun()
                        
            st.divider()
            if st.button("💾 Salvar Edição Completa no Excel", type="primary", use_container_width=True):
                with st.spinner("Sincronizando as alterações fisicamente no Excel (pode demorar alguns segundos)..."):
                    sucesso = dm.salvar_edicao_bloco_excel(
                        st.session_state["bloco_edit_id"], mat_edit, dem_edit, qtd_edit, vol_edit, st.session_state["roteiro_atual"]
                    )
                    if sucesso:
                        st.success("Edição salva com sucesso!")
                        st.session_state["bloco_carregado"] = False # Reseta a tela
                        st.rerun()
                    else:
                        st.error("Erro ao salvar edição.")

# ----------------- ABA 2: BASE DE DADOS -----------------
with tab_view:
    st.header("Base de Dados Completa")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        setores = ["Todos"] + sorted(list(set([str(x) for x in df["SETOR"] if str(x) != ""])))
        setor_filtro = st.selectbox("Filtrar por Setor", setores)
    with col2:
        status_list = ["Todos"] + sorted(list(set([str(x) for x in df["STATUS PROCESSO"] if str(x) != ""])))
        status_filtro = st.selectbox("Filtrar por Status", status_list)
    with col3:
        processos = ["Todos"] + sorted(list(set([str(x) for x in df["PROCESSO"] if str(x) != ""])))
        processo_filtro = st.selectbox("Filtrar por Processo", processos)
        
    filtered_df = df.copy()
    if setor_filtro != "Todos": filtered_df = filtered_df[filtered_df["SETOR"].astype(str) == setor_filtro]
    if status_filtro != "Todos": filtered_df = filtered_df[filtered_df["STATUS PROCESSO"].astype(str) == status_filtro]
    if processo_filtro != "Todos": filtered_df = filtered_df[filtered_df["PROCESSO"].astype(str) == processo_filtro]
        
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    st.metric("Total de Registros Filtrados", len(filtered_df))

# ----------------- ABA 3: JANELA DE PROGRAMAÇÕES -----------------
with tab_machine:
    st.header("Janela de Programações")
    st.markdown("Agende os processos em lote. Selecione os blocos na esquerda e verifique o impacto na capacidade na direita.")
    
    # 1. FUNÇÕES DE UTILIDADE (Clusterizadas para evitar erros de escopo)
    def match_date(d_val, target_date):
        """Confere se uma data do Excel bate com a data alvo selecionada."""
        if pd.isna(d_val) or str(d_val).strip() in ["", "nan", "NaT", "None"]: return False
        try:
            if isinstance(d_val, pd.Timestamp): return d_val.date() == target_date
            d_str = str(d_val).strip()
            if "/" in d_str: return pd.to_datetime(d_str, format="%d/%m/%Y").date() == target_date
            return pd.to_datetime(d_str).date() == target_date
        except: return False

    def safe_date_str(d):
        """Converte data do Excel em string amigável DD/MM, tratando nulos com '-'."""
        d_str = str(d).strip().upper()
        if pd.isna(d) or d_str in ["", "NAN", "NAT", "NONE", "NULL", "-"]: return "-"
        try:
            if isinstance(d, pd.Timestamp): return d.strftime("%d/%m")
            # Tenta converter string de data
            if "/" in str(d): return pd.to_datetime(str(d), format="%d/%m/%Y").strftime("%d/%m")
            return pd.to_datetime(str(d)).strftime("%d/%m")
        except: return "-"

    # 2. FILTROS E DADOS INICIAIS
    medias_hist = dm.get_historico_medias_entregues()
    
    st.write("---")
    c_top1, c_top2 = st.columns(2)
    with c_top1:
        data_alvo = st.date_input("Data Alvo (Para onde agendar?)", value=datetime.now(), format="DD/MM/YYYY")
        data_alvo_date = pd.to_datetime(data_alvo).date()
    with c_top2:
        maquina_foco = st.selectbox("Máquina/Setor Específico (Filtro)", ["Todos"] + sorted(setores_list), key="maq_foco_top")
    st.write("---")

    status_liberado = {}
    hoje = datetime.now().date()
    
    for bloco_id, group_bloco in df.groupby("BLOCO"):
        group_bloco = group_bloco.sort_index()
        for idx, row in group_bloco.iterrows():
            if is_finished(row): continue
            
            # Localiza a posição deste processo no roteiro total do bloco
            idx_pos = group_bloco.index.get_loc(idx)
            dt_prog = parse_dt(row.get("DATA"))
            
            # PRIORIDADE 1: Atrasados (Programado para trás e não feito)
            if dt_prog and dt_prog < hoje:
                status_liberado[idx] = "🟡 Atrasado"
                continue
            
            # PRIORIDADE 2: Já agendado em OUTRA data futura
            if dt_prog and dt_prog != data_alvo_date:
                status_liberado[idx] = "🔴 Agendado"
                continue

            # PRIORIDADE 3: Lógica de Sequência (Sim/Não)
            if idx_pos == 0:
                status_liberado[idx] = "🟢 Sim"
            else:
                idx_ant = group_bloco.index[idx_pos - 1]
                row_ant = group_bloco.loc[idx_ant]
                if is_finished(row_ant) or is_planned(row_ant):
                    status_liberado[idx] = "🟢 Sim"
                else:
                    status_liberado[idx] = "🔴 Não"
    
    # 4. COLUNA ESQUERDA: FILA DE TRABALHO
    col_esquerda, col_direita = st.columns([1.5, 1])
    with col_esquerda:
        st.subheader("📋 Fila de Trabalho")
        df_fila = df_aberto.copy()
        if maquina_foco != "Todos":
            df_fila = df_fila[df_fila["SETOR"] == maquina_foco]
            
        # Cálculo de Dias Parado
        hoje = datetime.now().date()
        dict_ult_proc = {}; dict_dias_parado = {}
        
        for b_id, group_b in df.groupby("BLOCO"):
            group_b = group_b.sort_index()
            # Encontra o último processo que está REALMENTE finalizado
            finalizados = [idx for idx, r in group_b.iterrows() if is_finished(r)]
            if finalizados:
                last_idx = finalizados[-1]
                row_f = group_b.loc[last_idx]
                dict_ult_proc[b_id] = row_f["PROCESSO"]
                dt_f = parse_dt(row_f.get("DATA REALIZADA")) or parse_dt(row_f.get("DATA"))
                if dt_f:
                    dict_dias_parado[b_id] = max((hoje - dt_f).days, 0)
            else:
                dict_ult_proc[b_id] = "Novo / Pátio"
                dict_dias_parado[b_id] = 0

        df_view = pd.DataFrame()
        df_view["Selecionar"] = df_fila["DATA"].apply(lambda x: match_date(x, data_alvo_date)).tolist()
        df_view["Liberado"] = df_fila.index.map(lambda x: status_liberado.get(x, "🔴 Não")).tolist()
        df_view["ID"] = df_fila.index.tolist()
        df_view["Bloco"] = df_fila["BLOCO"].tolist()
        df_view["Material"] = df_fila.get("MATERIAL", "").tolist()
        df_view["Máquina"] = df_fila["SETOR"].tolist()
        df_view["Processo"] = df_fila["PROCESSO"].tolist()
        df_view["Chapas"] = pd.to_numeric(df_fila["QTD. CHAPAS"], errors="coerce").fillna(0).astype(int).tolist()
        df_view["Últ. Processo"] = df_fila["BLOCO"].map(lambda x: dict_ult_proc.get(x, "Nenhum")).tolist()
        df_view["Parado"] = df_fila["BLOCO"].map(lambda x: dict_dias_parado.get(x, 0)).tolist()
        df_view["Data Prog."] = df_fila["DATA"].apply(safe_date_str)
        col_realizada = df_fila["DATA REALIZADA"] if "DATA REALIZADA" in df_fila.columns else pd.Series([None]*len(df_fila), index=df_fila.index)
        df_view["Data Realiz."] = col_realizada.apply(safe_date_str)
        
        editado_fila = st.data_editor(
            df_view,
            column_config={
                "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False),
                "Liberado": st.column_config.TextColumn("Liberado?"),
                "Máquina": st.column_config.SelectboxColumn("Máquina", options=sorted(setores_list), required=True),
                "Parado": st.column_config.NumberColumn("Parado", format="%d d"),
                "ID": st.column_config.NumberColumn("ID", width="small")
            },
            hide_index=True,
            use_container_width=True,
            disabled=["ID", "Liberado", "Processo", "Bloco", "Material", "Chapas", "Últ. Processo", "Parado", "Data Prog.", "Data Realiz."]
        )
        
        # Filtros baseados no ID preservado
        selecionados = editado_fila[editado_fila["Selecionar"] == True] if "Selecionar" in editado_fila.columns else pd.DataFrame()
        
        # Blocos que o usuário DESMARCOU (estavam agendados e agora não estão)
        desmarcados = pd.DataFrame()
        if "Selecionar" in editado_fila.columns and "ID" in editado_fila.columns:
            desmarcados = editado_fila[(editado_fila["Selecionar"] == False) & (df_fila["DATA"].apply(lambda x: match_date(x, data_alvo_date)).values)]
        
        # Selecionados Novos (para adicionar)
        selecionados_novos = pd.DataFrame()
        if not selecionados.empty and "ID" in selecionados.columns:
            selecionados_novos = selecionados[~selecionados["ID"].map(lambda idx: match_date(df.loc[idx, "DATA"], data_alvo_date))]
        
        # Cálculo Seguro de Chapas
        chapas_adicionadas = 0
        if not selecionados_novos.empty and "Chapas" in selecionados_novos.columns:
            chapas_adicionadas = selecionados_novos["Chapas"].sum()
            
        chapas_removidas = 0
        if not desmarcados.empty and "Chapas" in desmarcados.columns:
            chapas_removidas = desmarcados["Chapas"].sum()
            
        saldo_chapas = chapas_adicionadas - chapas_removidas

    # 5. COLUNA DIREITA: PAINEL DE ALOCAÇÃO
    with col_direita:
        st.subheader("📅 Painel de Alocação")
        with st.container(border=True):
            is_global = (maquina_foco == "Todos")
            maq_alvo_label = maquina_foco if not is_global else "TODAS AS MÁQUINAS"
            st.markdown(f"**Máquina Alvo:** {maq_alvo_label} | **Data:** {data_alvo.strftime('%d/%m/%Y')}")
            
            df_dia = df_aberto[df_aberto["DATA"].apply(lambda x: match_date(x, data_alvo_date))].copy()
            df_dia["QTD. CHAPAS"] = pd.to_numeric(df_dia["QTD. CHAPAS"], errors="coerce").fillna(0)
            
            if is_global:
                carga_existente = df_dia["QTD. CHAPAS"].sum()
                media_alvo = sum(medias_hist.values())
            else:
                carga_existente = df_dia[df_dia["SETOR"] == maquina_foco]["QTD. CHAPAS"].sum()
                media_alvo = medias_hist.get(maquina_foco, 0)
                
            carga_projetada = carga_existente + saldo_chapas
            
            st.write("---")
            col_k1, col_k2, col_k3 = st.columns(3)
            col_k1.metric("Já Agendada", f"{int(carga_existente)} ch")
            
            label_add = "Adicionando" if saldo_chapas >= 0 else "Removendo"
            col_k2.metric(label_add, f"{'+' if saldo_chapas >= 0 else ''}{int(saldo_chapas)} ch")
            
            delta_val = carga_projetada - media_alvo
            col_k3.metric("Projeção Total", f"{int(carga_projetada)} ch", delta=f"{delta_val:.1f} vs Média" if media_alvo > 0 else None, delta_color="inverse")
            
            # Detalhamento de blocos (Visual Profissional)
            df_detalhe = df_dia if is_global else df_dia[df_dia["SETOR"] == maquina_foco]
            if not df_detalhe.empty:
                with st.expander(f"📦 Detalhes do Agendamento ({len(df_detalhe)} blocos)", expanded=False):
                    df_exibir = df_detalhe.copy()
                    # Prepara colunas para exibição amigável
                    cols_show = ["SETOR", "BLOCO", "MATERIAL", "PROCESSO", "QTD. CHAPAS"]
                    df_exibir = df_exibir[cols_show].copy()
                    df_exibir.columns = ["Máquina", "Bloco", "Material", "Processo", "Chapas"]
                    df_exibir["Chapas"] = df_exibir["Chapas"].astype(int)
                    
                    st.dataframe(
                        df_exibir,
                        column_config={
                            "Chapas": st.column_config.NumberColumn("Chapas", format="%d ch"),
                            "Máquina": st.column_config.TextColumn("Máquina", width="small"),
                            "Bloco": st.column_config.TextColumn("Bloco", width="small")
                        },
                        hide_index=True,
                        use_container_width=True
                    )
            
            if media_alvo > 0:
                st.progress(min(max(carga_projetada / media_alvo, 0.0), 1.0))
                if carga_projetada > media_alvo:
                    st.error(f"⚠️ Atenção: Carga Global ({int(carga_projetada)}) ultrapassa a média total ({int(media_alvo)})!")
                    
            st.write("---")
            # Permite agendar se for global ou uma máquina específica
            if False: # Placeholder para manter indentação
                pass
            else:
                btn_label = "🚀 Confirmar Alterações" if saldo_chapas != 0 else "🚀 Agendar Selecionados"
                if st.button(btn_label, type="primary", use_container_width=True):
                    erros = []
                    sucessos = 0
                    remocoes = 0
                    nova_data_str = data_alvo.strftime("%d/%m/%Y")
                    
                    with st.spinner("Atualizando Programação..."):
                        # 1. AGENDAR NOVOS
                        if not selecionados_novos.empty and "ID" in selecionados_novos.columns:
                            for idx_sel in selecionados_novos["ID"]:
                                row_edited = selecionados_novos[selecionados_novos["ID"] == idx_sel].iloc[0]
                                nova_maquina = row_edited["Máquina"]
                                
                                valido, msg = dm.validar_sequencia_bloco(df, row_edited["Bloco"], idx_sel, nova_data_str)
                                if not valido:
                                    erros.append(f"Bloco {row_edited['Bloco']}: {msg}")
                                else:
                                    updates = {
                                        "DATA": nova_data_str,
                                        "SETOR": nova_maquina
                                    }
                                    if dm.update_cell_by_row(idx_sel, updates): sucessos += 1
                                    else: erros.append(f"Bloco {row_edited['Bloco']}: Erro ao salvar.")
                        
                        # 2. REMOVER DESMARCADOS
                        if not desmarcados.empty and "ID" in desmarcados.columns:
                            for idx_rem in desmarcados["ID"]:
                                bloco_rem = df.loc[idx_rem, "BLOCO"]
                                if dm.update_cell_by_row(idx_rem, {"DATA": None}):
                                    remocoes += 1
                                else:
                                    erros.append(f"Bloco {bloco_rem}: Erro ao remover.")
                                    
                    if sucessos > 0 or remocoes > 0:
                        msg_sucesso = []
                        if sucessos > 0: msg_sucesso.append(f"{sucessos} agendados")
                        if remocoes > 0: msg_sucesso.append(f"{remocoes} removidos")
                        st.success(f"Sucesso: {', '.join(msg_sucesso)}!")
                        st.rerun()
                    if erros:
                        for e in erros: 
                            st.error(f"🛑 {e}")
                            st.info("💡 Dica: Verifique se o arquivo Excel está fechado em todos os computadores.")

# ----------------- ABA 4: APONTAMENTO DE PRODUÇÃO -----------------
with tab_apontamento:
    st.header("✅ Apontamento de Produção")
    st.markdown("Cruza o que foi **apontado na planilha de produção** com o que estava **programado**, para você confirmar o que foi REALIZADO.")

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
                        st.dataframe(historico, use_container_width=True, hide_index=True)
                    else:
                        st.info("Nenhum apontamento anterior encontrado para este bloco.")

        # 2. Escolha do TIPO e ETAPA (Aparecem após o bloco ser digitado)
        if f_bloco:
            tipos_processo = ["", "Serrada / Corte", "Levigamento / Polimento", "Resinagem / Tela / Manta / Estuque", "Retoque", "Outros"]
            f_tipo = st.selectbox("1. Tipo de Processo*", tipos_processo, key="ap_tipo_proc_v3")
            
            if f_tipo:
                opcoes_proc_master = sorted(list(mapa_processos.keys()))
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
                            
                            st.form_submit_button("🔄 RECALCULAR / CONFIRMAR PROPORÇÃO", use_container_width=True)
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
                            editado_ins_add = st.data_editor(st.session_state["df_ins_add"], num_rows="dynamic", use_container_width=True, key="editor_ins_final")
                        
                        with c_i2:
                            st.caption("⏹️ Paradas de Máquina")
                            if "df_paradas_state" not in st.session_state:
                                d_ontem = datetime.now() - timedelta(days=1)
                                st.session_state["df_paradas_state"] = pd.DataFrame([{"MOTIVO": "", "DIA_INI": d_ontem, "HORA_INI": "", "DIA_FIM": d_ontem, "HORA_FIM": ""}])
                            
                            editado_paradas = st.data_editor(
                                st.session_state["df_paradas_state"], 
                                num_rows="dynamic", 
                                use_container_width=True, 
                                column_config={
                                    "DIA_INI": st.column_config.DateColumn("D.Início", format="DD/MM/YYYY"),
                                    "DIA_FIM": st.column_config.DateColumn("D.Fim", format="DD/MM/YYYY"),
                                    "HORA_INI": st.column_config.TextColumn("H.Início"),
                                    "HORA_FIM": st.column_config.TextColumn("H.Fim")
                                },
                                key="editor_paradas_final"
                            )

                        btn_carrinho = st.form_submit_button("🛒 ADICIONAR AO CARRINHO", use_container_width=True, type="primary")


                    if btn_carrinho:
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
                                            
                                        dt_ini_p = datetime.combine(p_row["DIA_INI"], h_ini_p)
                                        dt_fim_p = datetime.combine(p_row["DIA_FIM"], h_fim_p)
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
                                            "DIA_INICIO": p_row["DIA_INI"].strftime("%d/%m/%Y"), 
                                            "HORA_INICIO": h_ini_p.strftime("%H:%M"), 
                                            "DIA_FIM": p_row["DIA_FIM"].strftime("%d/%m/%Y"), 
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
                    for rec, par, ins in st.session_state["carrinho_ap"]:
                        dados_view.append({"Bloco": rec["BLOCO_RAW"], "Material": rec["NOME_MATERIAL"], "Processo": rec["PROCESSO_APONTADO"], "Chapas": rec["QTD_CH"]})
                    st.table(pd.DataFrame(dados_view))
                    
                    c_b1, c_b2 = st.columns(2)
                    with c_b1:
                        if st.button("🗑️ Limpar Carrinho", use_container_width=True):
                            st.session_state["carrinho_ap"] = []
                            st.rerun()
                    with c_b2:
                        if st.button("🚀 FINALIZAR E SALVAR TUDO", type="primary", use_container_width=True):
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

    c_ap1, c_ap2 = st.columns(2)
    with c_ap1:
        data_ap_sel = st.date_input("Data de Produção", value=datetime.now(), format="DD/MM/YYYY", key="data_ap_v2")
    with c_ap2:
        setores_prog = set(str(x) for x in df["SETOR"].unique() if str(x) not in ["", "nan"])
        setores_ap_pre = st.session_state.get("ap_setores_extras", set())
        setores_ap_todos = sorted(setores_prog | setores_ap_pre)
        maq_ap_sel = st.selectbox("Filtrar por Máquina", ["Todos"] + setores_ap_todos, key="maq_ap_v2")

    # Função para carregar dados
    def force_load_ap():
        with st.spinner(f"Lendo apontamentos de {data_ap_sel.strftime('%d/%m/%Y')}..."):
            st.session_state["df_ap_cache"] = dm.get_apontamentos_do_dia(data_ap_sel)
            st.session_state["last_ap_date"] = data_ap_sel
            
            if not st.session_state["df_ap_cache"].empty:
                st.success(f"✅ {len(st.session_state['df_ap_cache'])} registros de produção carregados com sucesso!")
            if not st.session_state["df_ap_cache"].empty and "SETOR_AP" in st.session_state["df_ap_cache"].columns:
                extras = set(str(x).strip() for x in st.session_state["df_ap_cache"]["SETOR_AP"].unique() if str(x).strip() not in ["", "nan"])
                st.session_state["ap_setores_extras"] = extras

    # Gatilho 1: Mudança de Data (Automático)
    if "last_ap_date" not in st.session_state or st.session_state["last_ap_date"] != data_ap_sel:
        force_load_ap()

    # Gatilho 2: Botão de Refresh (Manual)
    if st.button("🔄 Atualizar Base de Apontamentos", use_container_width=True):
        force_load_ap()
        st.toast("Dados atualizados!", icon="✅")

    df_ap = st.session_state.get("df_ap_cache", pd.DataFrame())
    data_ap_str = data_ap_sel.strftime("%d/%m/%Y")

    # --- FILTRAGEM DA PROGRAMAÇÃO (USANDO PARSE_DT ROBUSTO) ---
    def is_same_day(d_val, target_date):
        dt = parse_dt(d_val)
        if dt is None: return False
        # Se for datetime, extrai a data. Se já for date, usa diretamente.
        d_only = dt.date() if hasattr(dt, "date") and callable(getattr(dt, "date")) else dt
        return d_only == target_date

    # Filtra programação e realizados do dia
    df_prog_dia = df[df["DATA"].apply(lambda x: is_same_day(x, data_ap_sel))].copy()
    df_real_dia = df[(df["STATUS PROCESSO"] == "REALIZADO") & (df["DATA REALIZADA"].apply(lambda x: is_same_day(x, data_ap_sel)))].copy()
    
    # Une os dois grupos (evitando duplicatas se o agendado for igual ao realizado)
    df_prog_dia = pd.concat([df_prog_dia, df_real_dia], ignore_index=False).drop_duplicates()
    
    if maq_ap_sel != "Todos":
        df_prog_dia = df_prog_dia[df_prog_dia["SETOR"] == maq_ap_sel]

    st.write("---")

    # Mapa de bloco -> SETOR_AP do apontamento (para preencher máquinas sem SETOR)
    bloco_para_setor_ap = {}
    if not df_ap.empty and "SETOR_AP" in df_ap.columns:
        for _, r in df_ap.iterrows():
            b_norm = dm.normalize_bloco(r["BLOCO"])
            s_ap = str(r["SETOR_AP"]).strip()
            if b_norm and s_ap and s_ap != "nan":
                bloco_para_setor_ap[b_norm] = s_ap

    # Lógica de Cruzamento Melhorada
    match_flags = []
    match_info  = []
    maquinas_resolvidas = []  # setor efetivo (nossa base ou apontamento)
    
    for _, row_prog in df_prog_dia.iterrows():
        b = dm.normalize_bloco(row_prog["BLOCO"])
        proc_prog = str(row_prog["PROCESSO"]).strip().upper()
        # Quebra o processo programado em palavras-chave (ex: TELAR/MANTAR -> [TELAR, MANTAR])
        import re
        keywords_prog = [k for k in re.split(r'[^A-Z0-9]', proc_prog) if len(k) > 2]
        
        ja_realizado = str(row_prog.get("STATUS PROCESSO", "")).strip().upper() == "REALIZADO"
        setor_prog = str(row_prog.get("SETOR", "")).strip()
        
        # Completa máquina a partir do apontamento se estiver vazio
        setor_efetivo = setor_prog if setor_prog and setor_prog != "nan" else bloco_para_setor_ap.get(b, "N/I")
        maquinas_resolvidas.append(setor_efetivo)
        
        if ja_realizado:
            data_real = str(row_prog.get("DATA REALIZADA", ""))
            match_flags.append("🟢 Já Confirmado")
            match_info.append(f"Realizado em {data_real}")
            continue
            
        if not df_ap.empty:
            # Busca flexível por bloco
            linhas_bloco = df_ap[df_ap["BLOCO"].apply(dm.normalize_bloco) == b]
            if not linhas_bloco.empty:
                # Busca inteligente por processo (qualquer keyword batendo)
                match_proc = False
                ap_row_encontrada = None
                
                for _, r_ap in linhas_bloco.iterrows():
                    proc_ap_res = str(r_ap.get("RESUMIDO", "")).upper()
                    proc_ap_full = str(r_ap.get("PROCESSO_APONTADO", "")).upper()
                    
                    # Se as palavras-chave baterem com o resumido ou o nome completo
                    if any(kw in proc_ap_res or kw in proc_ap_full for kw in keywords_prog):
                        match_proc = True
                        ap_row_encontrada = r_ap
                        break
                
                if match_proc:
                    match_flags.append("✅ Encontrado")
                    match_info.append(f"{ap_row_encontrada['PROCESSO_APONTADO']} ({int(ap_row_encontrada['QTD_CH'])} ch)")
                else:
                    procs_ap = ", ".join(linhas_bloco["PROCESSO_APONTADO"].tolist())
                    match_flags.append("⚠️ Bloco sim, processo diferente")
                    match_info.append(f"Apontado: {procs_ap}")
            else:
                match_flags.append("❌ Não encontrado")
                match_info.append("-")
        else:
            match_flags.append("❌ Sem apontamento")
            match_info.append("-")

    # --- SEÇÃO 1: RESUMO (MÉTRICAS) ---
    total_prog = len(df_prog_dia)
    total_ja_confirmados = sum(1 for f in match_flags if "🟢" in f)
    total_encontrados = sum(1 for f in match_flags if "✅" in f)
    total_confirmados_efetivo = total_ja_confirmados + total_encontrados
    aderencia = (total_confirmados_efetivo / total_prog * 100) if total_prog > 0 else 0

    st.subheader(f"📊 Resumo Geral - {data_ap_str}")
    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
    col_m1.metric("Programados", total_prog)
    col_m2.metric("✅ No Apontamento", total_encontrados)
    col_m3.metric("❌ Pendentes", total_prog - total_confirmados_efetivo)
    col_m4.metric("🟢 Já Confirmados", total_ja_confirmados)
    col_m5.metric("📈 Aderência", f"{aderencia:.1f}%")
    st.progress(min(aderencia / 100, 1.0))

    st.write("---")

    # --- SEÇÃO 2: LADO A LADO (O CONFRONTO REAL) ---
    st.subheader(f"⚔️ Confronto Direto - {data_ap_str}")
    
    col_pcp, col_fab = st.columns(2)
    
    with col_pcp:
        st.markdown("### 📋 Plano (PCP)")
        if not df_prog_dia.empty:
            # Seleciona colunas essenciais para o confronto
            df_view_pcp = df_prog_dia[["SETOR", "BLOCO", "PROCESSO", "QTD. CHAPAS"]].copy()
            st.dataframe(df_view_pcp, use_container_width=True, hide_index=True, height=400)
        else:
            st.info("Nenhuma programação para este dia.")

    with col_fab:
        st.markdown("### 🏭 Realizado (Fábrica)")
        if not df_ap.empty:
            # Seleciona colunas essenciais do apontamento
            df_view_fab = df_ap[["SETOR_AP", "BLOCO", "PROCESSO_APONTADO", "QTD_CH"]].copy()
            st.dataframe(df_view_fab, use_container_width=True, hide_index=True, height=400)
        else:
            st.warning("Nenhum apontamento carregado para este dia.")
            if st.button("🔄 Carregar Apontamentos do Excel", use_container_width=True):
                force_load_ap()
                st.rerun()
    st.write("---")
    # --- SEÇÃO 3: FERRAMENTA DE CONFIRMAÇÃO ---
    if total_prog > 0:
        with st.expander("✅ Confirmar Realização e Sincronizar com Banco de Dados", expanded=True):
            st.markdown("O sistema já marcou abaixo os itens que ele encontrou correspondência automática:")
            df_confirm = pd.DataFrame()
            df_confirm["Confirmar?"]    = [True if "✅" in f else False for f in match_flags]
            df_confirm["Index"]         = df_prog_dia.index.tolist()
            df_confirm["Bloco"]         = df_prog_dia["BLOCO"].tolist()
            df_confirm["Processo Prog."]= df_prog_dia["PROCESSO"].tolist()
            df_confirm["Status Cruzamento"] = match_flags
            df_confirm["Encontrado na Fábrica Como"] = match_info

            editado = st.data_editor(
                df_confirm,
                column_config={
                    "Confirmar?": st.column_config.CheckboxColumn("Salvar?", default=False),
                    "Index": None,
                    "Encontrado na Fábrica Como": st.column_config.TextColumn("Relatório Fábrica", width="medium")
                },
                hide_index=True, use_container_width=True,
                disabled=["Bloco", "Processo Prog.", "Status Cruzamento", "Encontrado na Fábrica Como"],
                key="editor_sinc_final_v2"
            )

            confirmados = editado[editado["Confirmar?"] == True]
            if st.button(f"💾 Salvar {len(confirmados)} Confirmações no Banco de Dados", type="primary", use_container_width=True):
                sucessos = 0
                with st.spinner("Sincronizando..."):
                    for idx in confirmados["Index"]:
                        if dm.update_cell_by_row(idx, {"STATUS PROCESSO": "REALIZADO", "DATA REALIZADA": data_ap_str}):
                            sucessos += 1
                if sucessos > 0:
                    st.success(f"✅ {sucessos} processos atualizados!")
                    st.rerun()
with tab_export:
    st.header("🖨️ Exportação para o Chão de Fábrica")
    st.markdown("Selecione as datas, filtre a máquina se quiser e clique em Gerar. O relatório aparecerá na tela pronto para você pressionar **Ctrl + P** e imprimir!")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        data_ini = st.date_input("Data Inicial", value=datetime.now(), format="DD/MM/YYYY")
    with c2:
        data_fim = st.date_input("Data Final", value=datetime.now(), format="DD/MM/YYYY")
    with c3:
        setores_exp = [str(x) for x in df["SETOR"].unique() if str(x) not in ["", "nan"]]
        maq_exp = st.selectbox("Máquina (Opcional)", ["Todas"] + sorted(setores_exp), key="maq_exp")
        
    if st.button("📄 Gerar Relatório", type="primary"):
        df_aberto_exp = df[df["STATUS PROCESSO"] != "REALIZADO"].copy()
        if maq_exp != "Todas":
            df_aberto_exp = df_aberto_exp[df_aberto_exp["SETOR"] == maq_exp]

        def filter_dates_exp(d_val):
            if pd.isna(d_val) or str(d_val).strip() in ["", "nan"]: return False
            try:
                if isinstance(d_val, pd.Timestamp): d_date = d_val.date()
                elif "/" in str(d_val): d_date = pd.to_datetime(str(d_val), format="%d/%m/%Y").date()
                else: d_date = pd.to_datetime(str(d_val)).date()
                return data_ini <= d_date <= data_fim
            except: return False

        df_export = df_aberto_exp[df_aberto_exp["DATA"].apply(filter_dates_exp)].copy()
        df_export["QTD. CHAPAS"] = pd.to_numeric(df_export["QTD. CHAPAS"], errors="coerce").fillna(0)
        df_export["VOLUME M²"] = pd.to_numeric(df_export.get("VOLUME M²", 0), errors="coerce").fillna(0)

        def fmt_d(d):
            if pd.isna(d) or str(d).strip() in ["", "nan"]: return "-"
            try:
                if isinstance(d, pd.Timestamp): return d.strftime("%d/%m/%Y")
                if "/" in str(d): return pd.to_datetime(str(d), format="%d/%m/%Y").strftime("%d/%m/%Y")
                return pd.to_datetime(str(d)).strftime("%d/%m/%Y")
            except: return str(d)

        df_export["DATA_FMT"] = df_export["DATA"].apply(fmt_d)
        df_export = df_export.sort_values(by=["SETOR", "DATA_FMT"])

        if df_export.empty:
            st.warning("Nenhum bloco agendado para o período selecionado.")
        else:
            periodo_str = f"{data_ini.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
            titulo = f"PROGRAMAÇÃO: {maq_exp.upper()} &nbsp;|&nbsp; {periodo_str}"

            rows_html = ""
            total_chapas = 0
            total_vol = 0

            for maq, df_maq in df_export.groupby("SETOR", sort=True):
                first_maq = True
                for data_val, df_dia in df_maq.groupby("DATA_FMT", sort=True):
                    sub_ch = 0
                    sub_vol = 0
                    first_data = True
                    for _, row in df_dia.iterrows():
                        ch = row["QTD. CHAPAS"]
                        vol = row["VOLUME M²"]
                        obs = str(row.get("OBSERVAÇÃO DE PRODÃO", "") or row.get("OBSERVAÇÃO DE PRODUÇÃO", "") or "").strip()
                        sub_ch += ch; sub_vol += vol
                        maq_cell = f"<b>{maq}</b>" if first_maq and first_data else ""
                        data_cell = f"<b>{data_val}</b>" if first_data else ""
                        rows_html += f"""
                        <tr>
                            <td class='td-maq'>{maq_cell}</td>
                            <td class='td-data'>{data_cell}</td>
                            <td>{row['MATERIAL']}</td>
                            <td>{row['BLOCO']}</td>
                            <td>{row['DEMANDA']}</td>
                            <td>{row['PROCESSO']}</td>
                            <td>{obs}</td>
                            <td class='td-num'>{int(ch)}</td>
                            <td class='td-num'>{vol:.2f}</td>
                            <td class='td-check'></td>
                        </tr>"""
                        first_data = False; first_maq = False

                    rows_html += f"""
                    <tr class='subtotal'>
                        <td colspan='7' class='sub-label'>Subtotal &mdash; {data_val}</td>
                        <td class='td-num'><b>{int(sub_ch)}</b></td>
                        <td class='td-num'><b>{sub_vol:.2f}</b></td>
                        <td></td>
                    </tr>"""
                    total_chapas += sub_ch; total_vol += sub_vol

            rows_html += f"""
            <tr class='total-geral'>
                <td colspan='7' class='sub-label'>TOTAL GERAL</td>
                <td class='td-num'><b>{int(total_chapas)}</b></td>
                <td class='td-num'><b>{total_vol:.2f}</b></td>
                <td></td>
            </tr>"""

            html = f"""
            <style>
              @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
              
              #relatorio-print {{ font-family: 'Inter', Arial, sans-serif; font-size: 11px; color: #111; }}
              #relatorio-print h2 {{ font-size: 15px; text-align: center; margin-bottom: 8px; letter-spacing: 1px; text-transform: uppercase; }}
              #relatorio-print .periodo {{ text-align: center; font-size: 10px; color: #555; margin-bottom: 12px; }}
              
              #relatorio-print table {{ width: 100%; border-collapse: collapse; }}
              #relatorio-print th {{
                  background: #1a56a0; color: #fff; font-weight: 700;
                  padding: 6px 4px; text-align: center; border: 1px solid #aaa; font-size: 10px;
              }}
              #relatorio-print td {{ padding: 5px 4px; border: 1px solid #ccc; vertical-align: middle; }}
              #relatorio-print tr:nth-child(even):not(.subtotal):not(.total-geral) {{ background: #f5f8ff; }}
              #relatorio-print .td-maq {{ font-weight: 700; color: #1a56a0; white-space: nowrap; }}
              #relatorio-print .td-data {{ white-space: nowrap; font-weight: 600; }}
              #relatorio-print .td-num {{ text-align: center; white-space: nowrap; }}
              #relatorio-print .td-check {{ width: 50px; min-width: 50px; }}
              #relatorio-print .subtotal td {{ background: #dce8f7; font-weight: 600; }}
              #relatorio-print .subtotal .sub-label {{ text-align: right; padding-right: 10px; }}
              #relatorio-print .total-geral td {{ background: #1a56a0; color: #fff; font-weight: 700; }}
              #relatorio-print .total-geral .sub-label {{ text-align: right; padding-right: 10px; }}
              
              .print-btn {{
                  display: inline-block; margin-bottom: 16px; padding: 8px 22px;
                  background: #1a56a0; color: #fff; border: none; border-radius: 6px;
                  cursor: pointer; font-size: 13px; font-weight: 600;
              }}
              .print-btn:hover {{ background: #154080; }}
              
              @media print {{
                  body * {{ visibility: hidden !important; }}
                  #relatorio-print, #relatorio-print * {{ visibility: visible !important; }}
                  #relatorio-print {{ position: fixed; top: 0; left: 0; width: 100%; }}
                  .print-btn {{ display: none !important; }}
                  #relatorio-print table {{ font-size: 10px; }}
              }}
            </style>
            <div id="relatorio-print">
              <h2>🏷️ {titulo}</h2>
              <p class="periodo">Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
              <button class="print-btn" onclick="window.print()">&#128424; Imprimir</button>
              <table>
                <thead>
                  <tr>
                    <th>MÁQUINA</th><th>DATA</th><th>MATERIAL</th><th>BLOCO</th>
                    <th>DEMANDA</th><th>PROCESSO</th><th>OBSERVAÇÃO</th>
                    <th>QTD.<br>CHAPAS</th><th>VOL.<br>M²</th><th>FEITO?</th>
                  </tr>
                </thead>
                <tbody>{rows_html}</tbody>
              </table>
            </div>
            """
            st.components.v1.html(html, height=max(600, 55 * len(df_export) + 200), scrolling=True)

# Função auxiliar para abrir dialog do Windows (definida fora para evitar redefinição)
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

# ----------------- ABA 6: ANÁLISES E INDICADORES -----------------
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

        if periodo == "Últimos 7 dias": df_an = df_an[df_an["DIA_PROD"] >= hoje - timedelta(days=7)]
        elif periodo == "Últimos 30 dias": df_an = df_an[df_an["DIA_PROD"] >= hoje - timedelta(days=30)]
        elif periodo == "Mês Atual": df_an = df_an[pd.to_datetime(df_an["DIA_PROD"]).dt.month == hoje.month]
        elif periodo == "Personalizado": df_an = df_an[(df_an["DIA_PROD"] >= data_ini_custom) & (df_an["DIA_PROD"] <= data_fim_custom)]
        
        df_an = df_an[~df_an[c_pr].astype(str).str.upper().str.contains("RETOQUE")]
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
            st.dataframe(df_final, use_container_width=True)
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
                x=alt.X('MAQ_TURNO:N', title=None, sort=None),
                y=alt.Y(f'{col_valor}:Q', title=metrica_board, stack='zero'),
                color=alt.Color('TIPO_PROD:N', title='Tipo', scale=alt.Scale(domain=['Normal', 'Refeito'], range=['#00CC96', '#EF553B'])),
                order=alt.Order('_sort:Q')
            )
            
            # Valores centralizados dentro de cada segmento
            text_seg = alt.Chart(df_board_gr).mark_text(
                align='center', baseline='middle', color='white', fontSize=15, fontWeight='bold', stroke='black', strokeWidth=0.5
            ).encode(
                x=alt.X('MAQ_TURNO:N', sort=None),
                y=alt.Y('_mid:Q'),
                text=alt.Text('_label:N')
            )
            
            # Totais no topo
            text_totals = alt.Chart(df_board_gr).mark_text(
                align='center', baseline='bottom', dy=-5, fontSize=18, fontWeight='bold', color='white', stroke='black', strokeWidth=0.6
            ).encode(x=alt.X('MAQ_TURNO:N', sort=None), y=alt.Y(f'sum({col_valor}):Q'), text=alt.Text(f'sum({col_valor}):Q', format=fmt))
            
            chart_layered = (bars + text_seg + text_totals).properties(width=alt.Step(100), height=450)
            
            faceted_chart = chart_layered.facet(
                column=alt.Column('DIA:N', title='Dia de Produção', sort=alt.SortOrder('ascending'))
            ).configure_view(stroke=None).configure_axis(labelFontSize=13, titleFontSize=15)
            
            st.altair_chart(faceted_chart, use_container_width=True)
            
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
            st.altair_chart((bars_evol + text_evol + totals_evol).properties(height=450), use_container_width=True)
            
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
            st.altair_chart((bars_turno + text_turno + totals_turno).properties(height=450), use_container_width=True)


# ----------------- ABA 7: OPÇÕES GERAIS -----------------
with tab_config:
    st.header("⚙️ Opções Gerais")
    st.subheader("📁 Caminhos e Abas das Bases de Dados")
    st.markdown("Configure os arquivos e **qual aba** de cada planilha o sistema deve usar. Após salvar, atualize a página.")

    cfg_atual = dm.get_config()



    # --- Arquivo de Programação ---
    st.markdown("#### 📊 Arquivo de Programação (`.xlsm`)")
    
    # Inicializa session_state para manter valor digitado ou selecionado
    if "tmp_db_path" not in st.session_state:
        st.session_state["tmp_db_path"] = cfg_atual.get("DB_FILE", "")

    col_p1, col_p2, col_p3 = st.columns([3, 1, 1])
    with col_p1:
        novo_db = st.text_input(
            "Caminho completo",
            value=st.session_state["tmp_db_path"],
            key="cfg_db_path_input",
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\DB.xlsm"
        )
        if novo_db != st.session_state["tmp_db_path"]:
            st.session_state["tmp_db_path"] = novo_db
    with col_p2:
        if st.button("🔍 Buscar...", key="btn_busc_db"):
            selecionado = gui_select_file(st.session_state["tmp_db_path"])
            if selecionado:
                st.session_state["tmp_db_path"] = selecionado
                st.rerun()
    with col_p3:
        import os as _osc
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
            "Caminho completo",
            value=st.session_state["tmp_ap_path"],
            key="cfg_ap_path_input",
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\DB.xlsx"
        )
        if novo_ap != st.session_state["tmp_ap_path"]:
            st.session_state["tmp_ap_path"] = novo_ap
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
            "Caminho completo",
            value=st.session_state["tmp_bl_path"],
            key="cfg_bl_path_input",
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\PLANILHA BLOCOS.xlsb"
        )
        if novo_bl != st.session_state["tmp_bl_path"]:
            st.session_state["tmp_bl_path"] = novo_bl
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
            "Caminho completo",
            value=st.session_state["tmp_ch_path"],
            key="cfg_ch_path_input",
            label_visibility="collapsed",
            placeholder="Ex: z:\\PCP\\Estoque Chapas.xlsx"
        )
        if novo_ch != st.session_state["tmp_ch_path"]:
            st.session_state["tmp_ch_path"] = novo_ch
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
        if st.button("💾 Salvar Configurações", type="primary", use_container_width=True, key="btn_salvar_cfg"):
            erros_path = []
            if not _osc.path.exists(novo_db):
                erros_path.append(f"❌ Arquivo de Programação não encontrado: `{novo_db}`")
            if not _osc.path.exists(novo_ap):
                erros_path.append(f"❌ Arquivo de Apontamento não encontrado: `{novo_ap}`")
            if not _osc.path.exists(novo_bl):
                erros_path.append(f"❌ Planilha de Blocos não encontrada: `{novo_bl}`")
            if not _osc.path.exists(novo_ch):
                erros_path.append(f"❌ Planilha de Chapas não encontrada: `{novo_ch}`")

            if erros_path:
                for ep in erros_path:
                    st.error(ep)
            else:
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
                    st.success("✅ Configurações salvas! Aplicando...")
                    st.rerun()
                else:
                    st.error("Erro ao salvar o arquivo de configuração.")

    with col_btn2:
        if st.button("↩️ Restaurar Padrões", use_container_width=True, key="btn_restaurar_cfg"):
            if dm.save_config(dm._DEFAULT_CONFIG):
                st.success("↩️ Configurações restauradas para os padrões. Atualize a página.")
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
        use_container_width=True,
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

