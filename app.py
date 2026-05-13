import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime
import importlib
import data_manager as dm

importlib.reload(dm)

st.set_page_config(page_title="Gerenciador de Programação", layout="wide", page_icon="📊")

st.title("📊 Gerenciador de Programação WIP")

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
lista_processos = sorted(list(set(lista_processos)))

tab_block, tab_view, tab_machine, tab_apontamento, tab_export, tab_config = st.tabs([
    "🛠️ Adicionar / Editar Bloco", 
    "👁️ Base de Dados", 
    "🗓️ Janela de Programações",
    "✅ Apontamento",
    "🖨️ Exportação",
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
        c1, c2, c3 = st.columns(3)
        with c1:
            material = st.text_input("Material*", key="c_mat")
            bloco = st.text_input("Número do Bloco*", key="c_blo")
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
                        st.session_state["mat_edit"] = str(primeira.get("MATERIAL", ""))
                        st.session_state["dem_edit"] = str(primeira.get("DEMANDA", ""))
                        try: st.session_state["qtd_edit"] = int(pd.to_numeric(primeira.get("QTD. CHAPAS", 0)))
                        except: st.session_state["qtd_edit"] = 0
                        try: st.session_state["vol_edit"] = float(pd.to_numeric(primeira.get("VOLUME M²", 0.0)))
                        except: st.session_state["vol_edit"] = 0.0
                        
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
            
            for i, etapa in enumerate(st.session_state["roteiro_atual"]):
                is_realizado = (etapa.get("STATUS PROCESSO") == "REALIZADO")
                
                d_prog = format_data_view(etapa.get('DATA'))
                d_real = format_data_view(etapa.get('DATA REALIZADA'))
                datas_html = f"<div style='font-size:0.85em; color:gray;'>Prog: <b>{d_prog}</b> <br> Real: <b>{d_real}</b></div>"
                
                with st.container(border=True):
                    col1, col2, col3, col4 = st.columns([1, 2, 4, 1])
                    if is_realizado:
                        col1.markdown(f"**{i+1}. 🔒 {etapa['PROCESSO']}**<br><small>{etapa['SETOR']}</small>", unsafe_allow_html=True)
                        col2.info("Status: REALIZADO")
                        col2.markdown(datas_html, unsafe_allow_html=True)
                        col3.write(f"Observação: {etapa['OBSERVACAO']}")
                        col4.markdown("<div style='text-align:center; padding-top:10px;'>Não Editável</div>", unsafe_allow_html=True)
                        novas_etapas.append(etapa)
                    else:
                        # Permite editar o setor ou observar
                        col1.markdown(f"**{i+1}. {etapa['PROCESSO']}**")
                        novo_setor = col1.text_input("Setor", value=etapa["SETOR"], key=f"eset_{i}", label_visibility="collapsed")
                        col2.warning(f"Status: {etapa.get('STATUS PROCESSO')}")
                        col2.markdown(datas_html, unsafe_allow_html=True)
                        nova_obs = col3.text_input("Observação (Chão de Fábrica)", value=etapa["OBSERVACAO"], key=f"eobs_{i}", label_visibility="collapsed")
                        
                        if col4.button("🗑️ Remover", key=f"edelt_{i}"):
                            st.warning("Clique em 'Salvar Edição' no final para efetivar a remoção.")
                        else:
                            etapa_atualizada = etapa.copy()
                            etapa_atualizada["SETOR"] = novo_setor
                            etapa_atualizada["OBSERVACAO"] = nova_obs
                            novas_etapas.append(etapa_atualizada)
                            
            st.session_state["roteiro_atual"] = novas_etapas
            
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
    
    medias_hist = dm.get_historico_medias_entregues()
    df_aberto = df[df["STATUS PROCESSO"] != "REALIZADO"].copy()
    
    # Mover os seletores principais para o topo para influenciar a tabela
    st.write("---")
    c_top1, c_top2 = st.columns(2)
    with c_top1:
        data_alvo = st.date_input("Data Alvo (Para onde agendar?)", value=datetime.now(), format="DD/MM/YYYY")
    with c_top2:
        setores_list = [str(x) for x in df_aberto["SETOR"].unique() if str(x) != "" and str(x) != "nan"]
        maquina_foco = st.selectbox("Máquina/Setor Específico (Filtro)", ["Todos"] + sorted(setores_list), key="maq_foco_top")
    st.write("---")
    
    data_alvo_date = pd.to_datetime(data_alvo).date()
    
    # --- AVALIAÇÃO DE BLOQUEIO DINÂMICO ---
    status_liberado = {}
    for bloco_id, group in df_aberto.groupby("BLOCO"):
        group = group.sort_index()
        for pos, idx in enumerate(group.index):
            if pos == 0:
                status_liberado[idx] = "🟢 Sim"
            else:
                pode_programar = True
                for i in range(pos):
                    idx_ant = group.index[i]
                    d_ant = group.loc[idx_ant, "DATA"]
                    if pd.isna(d_ant) or str(d_ant).strip() in ["", "nan", "NaT"]:
                        pode_programar = False
                        break
                    else:
                        # Tem data. Verifica se a data_ant é <= data_alvo
                        try:
                            if isinstance(d_ant, pd.Timestamp): d_ant_dt = d_ant.date()
                            elif "/" in str(d_ant): d_ant_dt = pd.to_datetime(str(d_ant), format="%d/%m/%Y").date()
                            else: d_ant_dt = pd.to_datetime(str(d_ant)).date()
                            
                            if d_ant_dt > data_alvo_date:
                                pode_programar = False
                                break
                        except: pass
                        
                status_liberado[idx] = "🟢 Sim" if pode_programar else "🔴 Não"
    # -----------------------------------------------------------------
    
    col_esquerda, col_direita = st.columns([1.5, 1])
    
    with col_esquerda:
        st.subheader("📋 Fila de Trabalho")
        st.info("💡 Marque os blocos na coluna 'Selecionar' para agendá-los juntos.")
            
        df_fila = df_aberto.copy()
        if maquina_foco != "Todos":
            df_fila = df_fila[df_fila["SETOR"] == maquina_foco]
            
        df_view = pd.DataFrame()
        df_view["Selecionar"] = [False] * len(df_fila)
        df_view["Liberado"] = df_fila.index.map(lambda x: status_liberado.get(x, "🔴 Não")).tolist()
        df_view["Índice"] = df_fila.index.tolist()
        df_view["Máquina"] = df_fila["SETOR"].tolist()
        df_view["Processo"] = df_fila["PROCESSO"].tolist()
        df_view["Bloco"] = df_fila["BLOCO"].tolist()
        df_view["Chapas"] = pd.to_numeric(df_fila["QTD. CHAPAS"], errors="coerce").fillna(0).astype(int).tolist()
        
        # --- Cálculo do Tempo Parado (WIP) ---
        df_realizados = df[df["STATUS PROCESSO"] == "REALIZADO"].copy()
        last_realizados = df_realizados.groupby("BLOCO").tail(1).set_index("BLOCO")
        hoje = datetime.now().date()
        
        dict_ult_proc = last_realizados["PROCESSO"].to_dict()
        dict_ult_data = {}
        dict_dias_parado = {}
        
        for b_id, row_b in last_realizados.iterrows():
            d_raw = row_b.get("DATA REALIZADA", "")
            if pd.isna(d_raw) or str(d_raw).strip() in ["", "nan", "NaT"]:
                d_raw = row_b.get("DATA", "") # Fallback
                
            d = None
            try:
                if isinstance(d_raw, pd.Timestamp): d = d_raw.date()
                elif "/" in str(d_raw): d = pd.to_datetime(str(d_raw), format="%d/%m/%Y").date()
                elif str(d_raw).strip() not in ["", "nan", "NaT"]: d = pd.to_datetime(str(d_raw)).date()
            except: pass
            
            if d:
                dict_ult_data[b_id] = d.strftime("%d/%m/%Y")
                dias = (hoje - d).days
                dict_dias_parado[b_id] = dias if dias >= 0 else 0
            else:
                dict_ult_data[b_id] = "-"
                dict_dias_parado[b_id] = 0
                
        df_view["Últ. Processo"] = df_fila["BLOCO"].map(lambda x: dict_ult_proc.get(x, "Nenhum")).tolist()
        df_view["Data ÚIt."] = df_fila["BLOCO"].map(lambda x: dict_ult_data.get(x, "-")).tolist()
        df_view["Dias Parado"] = df_fila["BLOCO"].map(lambda x: dict_dias_parado.get(x, 0)).tolist()
        # -------------------------------------
        
        def safe_date_str(d):
            if pd.isna(d) or str(d).strip() in ["", "nan", "NaT"]: return "-"
            try:
                if isinstance(d, pd.Timestamp): return d.strftime("%d/%m")
                if "/" in str(d): return pd.to_datetime(str(d), format="%d/%m/%Y").strftime("%d/%m")
                return pd.to_datetime(str(d)).strftime("%d/%m")
            except: return str(d)
            
        df_view["Data Atual"] = df_fila["DATA"].apply(safe_date_str)
        
        editado_fila = st.data_editor(
            df_view,
            column_config={
                "Selecionar": st.column_config.CheckboxColumn("Selecionar", default=False),
                "Liberado": st.column_config.TextColumn("Liberado?"),
                "Dias Parado": st.column_config.NumberColumn("Dias Parado", format="%d d"),
                "Índice": None # Escondido
            },
            hide_index=True,
            use_container_width=True,
            disabled=["Liberado", "Máquina", "Processo", "Bloco", "Chapas", "Data Atual", "Últ. Processo", "Data ÚIt.", "Dias Parado"]
        )
        
        selecionados = editado_fila[editado_fila["Selecionar"] == True]
        chapas_selecionadas = selecionados["Chapas"].sum()
        
    with col_direita:
        st.subheader("📅 Painel de Alocação")
        
        with st.container(border=True):
            maq_alvo = maquina_foco if maquina_foco != "Todos" else "Nenhuma Selecionada"
            st.markdown(f"**Máquina Alvo:** {maq_alvo} | **Data:** {data_alvo.strftime('%d/%m/%Y')}")
            
            def match_date(d_val, target_date):
                if pd.isna(d_val) or str(d_val).strip() in ["", "nan", "NaT"]: return False
                try:
                    if isinstance(d_val, pd.Timestamp): return d_val.date() == target_date
                    if "/" in str(d_val): return pd.to_datetime(str(d_val), format="%d/%m/%Y").date() == target_date
                    return pd.to_datetime(str(d_val)).date() == target_date
                except: return False
                
            df_dia = df_aberto[df_aberto["DATA"].apply(lambda x: match_date(x, data_alvo))].copy()
            df_dia["QTD. CHAPAS"] = pd.to_numeric(df_dia["QTD. CHAPAS"], errors="coerce").fillna(0)
            
            carga_existente = 0
            df_dia_maq = pd.DataFrame()
            if maq_alvo != "Nenhuma Selecionada":
                df_dia_maq = df_dia[df_dia["SETOR"] == maq_alvo]
                carga_existente = df_dia_maq["QTD. CHAPAS"].sum()
                
            media_alvo = medias_hist.get(maq_alvo, 0)
            carga_projetada = carga_existente + chapas_selecionadas
            
            st.write("---")
            col_k1, col_k2, col_k3 = st.columns(3)
            col_k1.metric("Carga Já Agendada", f"{int(carga_existente)} ch")
            col_k2.metric("Sendo Adicionado", f"+ {int(chapas_selecionadas)} ch")
            
            delta = carga_projetada - media_alvo
            delta_str = f"{delta:.1f} vs Média"
            col_k3.metric("Projeção Total", f"{int(carga_projetada)} ch", delta=delta_str, delta_color="inverse" if media_alvo > 0 else "off")
            
            if not df_dia_maq.empty:
                with st.expander("👀 Ver blocos já agendados para este dia", expanded=False):
                    for _, row_agendada in df_dia_maq.iterrows():
                        chapas_row = int(row_agendada["QTD. CHAPAS"])
                        material_row = str(row_agendada.get("MATERIAL", ""))
                        st.markdown(f"📦 **Bloco {row_agendada['BLOCO']}** ({material_row}) - {row_agendada['PROCESSO']} <span style='color:gray'>({chapas_row} chapas)</span>", unsafe_allow_html=True)
            
            if media_alvo > 0:
                st.progress(min(carga_projetada / media_alvo, 1.0))
                if carga_projetada > media_alvo:
                    st.error(f"⚠️ Atenção: A carga projetada ({int(carga_projetada)}) ultrapassa a média histórica diária ({int(media_alvo)}) desta máquina!")
            else:
                if maq_alvo != "Nenhuma Selecionada":
                    st.info("Máquina sem média histórica na aba ENTREGUES.")
                    
            st.write("---")
            if maq_alvo == "Nenhuma Selecionada":
                st.warning("Selecione uma Máquina Específica no filtro da esquerda para agendar em lote.")
            else:
                if st.button("🚀 Agendar Selecionados", type="primary", use_container_width=True, disabled=len(selecionados)==0):
                    erros = []
                    sucessos = 0
                    nova_data_str = data_alvo.strftime("%d/%m/%Y")
                    
                    with st.spinner("Validando e Agendando..."):
                        for idx_selecionado in selecionados["Índice"]:
                            row_data_prog = df.loc[idx_selecionado]
                            bloco = row_data_prog["BLOCO"]
                            valido, msg_erro = dm.validar_sequencia_bloco(df, bloco, idx_selecionado, nova_data_str)
                            
                            if not valido:
                                erros.append(f"Bloco {bloco}: {msg_erro}")
                            else:
                                updates = {
                                    "DATA": nova_data_str,
                                    "STATUS PROCESSO": "EM PROCESSO" if str(row_data_prog["STATUS PROCESSO"]).upper() == "NÃO REALIZADO" else row_data_prog["STATUS PROCESSO"]
                                }
                                if dm.update_cell_by_row(idx_selecionado, updates):
                                    sucessos += 1
                                else:
                                    erros.append(f"Bloco {bloco}: Erro ao salvar no banco.")
                                    
                    if sucessos > 0:
                        st.success(f"{sucessos} etapas agendadas com sucesso para {nova_data_str}!")
                    if erros:
                        for e in erros:
                            st.error(f"🛑 {e}")
                            
                    if sucessos > 0:
                        st.rerun()

# ----------------- ABA 4: APONTAMENTO DE PRODUÇÃO -----------------
with tab_apontamento:
    st.header("✅ Apontamento de Produção")
    st.markdown("Cruza o que foi **apontado na planilha de produção** com o que estava **programado**, para você confirmar o que foi REALIZADO.")

    # ---- NOVO FORMULÁRIO DE APONTAMENTO MANUAL DINÂMICO ----
    with st.expander("➕ Lançar Novo Apontamento Manual", expanded=False):
        st.markdown("### Registro de Produção")
        
        # 1. Pergunta o TIPO genérico de processo
        tipos_processo = ["", "Serrada / Corte", "Levigamento / Polimento", "Resinagem / Tela / Manta / Estuque", "Retoque", "Outros"]
        f_tipo = st.selectbox("1. Tipo de Processo*", tipos_processo, key="ap_tipo_proc")
        
        if f_tipo:
            # 2. Pergunta o PROCESSO específico da base
            opcoes_proc_padrao = sorted(lista_processos)
            f_processo = st.selectbox("2. Processo Específico*", [""] + opcoes_proc_padrao, key="ap_f_proc")
            
            if f_processo:
                with st.form("form_novo_apontamento_dyn", clear_on_submit=True, enter_to_submit=False):
                    st.markdown("#### Dados Principais")
                    c_p1, c_p2, c_p3, c_p4 = st.columns(4)
                    with c_p1:
                        f_data = st.date_input("Data Reg.*", value=datetime.now(), format="DD/MM/YYYY")
                        f_bloco = st.text_input("Nº Bloco*", value=None, placeholder="Ex: 1234")
                        f_material = st.text_input("Material*", value=None, placeholder="Ex: ALPINUS")
                    with c_p2:
                        setores_ativos = set(str(x) for x in df["SETOR"].unique() if str(x) not in ["", "nan"])
                        setores_mapa = set(v for v in mapa_processos.values() if v)
                        setores_disponiveis = sorted(list(setores_ativos | setores_mapa))
                        f_setor = st.selectbox("Máquina/Setor*", [""] + setores_disponiveis)
                        f_operador = st.text_input("Operador", value=None)
                        f_turno = st.selectbox("Turno", ["", "D", "N"])
                    with c_p3:
                        f_qtd_ch = st.number_input("Qtd Chapas*", min_value=0, step=1, value=None)
                        f_dureza = st.text_input("Dureza", value=None)
                    with c_p4:
                        f_esp = st.number_input("Espessura", min_value=0.0, step=0.1, value=None)
                        f_comp = st.number_input("Comprimento (m)", min_value=0.0, step=0.01, value=None)
                        f_alt = st.number_input("Altura (m)", min_value=0.0, step=0.01, value=None)

                    st.markdown("#### Tempos")
                    c_t1, c_t2, c_t3, c_t4 = st.columns(4)
                    with c_t1: f_dia_ini = st.date_input("Dia Início", value=None, format="DD/MM/YYYY")
                    with c_t2: f_dia_fim = st.date_input("Dia Fim", value=None, format="DD/MM/YYYY")
                    with c_t3: f_hora_ini_str = st.text_input("Hora Início", value="", placeholder="Ex: 1040")
                    with c_t4: f_hora_fim_str = st.text_input("Hora Fim", value="", placeholder="Ex: 1530")

                    st.markdown("#### Insumos Específicos")
                    extra_data = {}
                    
                    # Lógica agora baseada no TIPO, independente do nome do processo
                    if f_tipo == "Levigamento / Polimento":
                        c_i1, c_i2 = st.columns(2)
                        with c_i1: extra_data["VEL_ESTEIRA"] = st.text_input("Vel. Esteira", value=None)
                        with c_i2: extra_data["VEL_TRAVE"] = st.text_input("Vel. Trave", value=None)
                        
                        with st.expander("🛠️ Sequência de Abrasivos (1 a 20)", expanded=False):
                            for r in range(5):
                                cols_abr = st.columns(4)
                                for c in range(4):
                                    i = r * 4 + c + 1
                                    val_abr = cols_abr[c].text_input(f"Seq. Abr. {i}", value=None, key=f"abr_{i}")
                                    if val_abr:
                                        extra_data[f"Seq. Abr. {i}"] = val_abr
                    
                    elif f_tipo == "Resinagem / Tela / Manta / Estuque":
                        c_i1, c_i2, c_i3 = st.columns(3)
                        with c_i1:
                            extra_data["TIPO_ACIDO"] = st.text_input("Tipo Ácido", value=None)
                            extra_data["TIPO_RESINA"] = st.text_input("Tipo Resina / Manta", value=None)
                        with c_i2:
                            extra_data["QTD_KG"] = st.number_input("Qtd KG (Resina)", min_value=0.0, step=0.1, value=None)
                            extra_data["TIPO_ENDUR"] = st.text_input("Tipo Endurecedor", value=None)
                        with c_i3:
                            extra_data["QTD_KG3"] = st.number_input("Qtd KG (Endurecedor)", min_value=0.0, step=0.1, value=None)
                            extra_data["V_24H"] = st.text_input("Tempo de Secagem", value=None)
                    
                    elif f_tipo == "Retoque":
                        c_i1, c_i2 = st.columns(2)
                        with c_i1: extra_data["IMPERMEABILIZANTE"] = st.text_input("Impermeabilizante", value=None)
                        with c_i2: extra_data["LIXA"] = st.text_input("Lixa Utilizada", value=None)
                    
                    else:
                        st.info("Nenhum insumo específico necessário para este tipo de processo.")
                    
                    if st.form_submit_button("🚀 Gravar Apontamento", type="primary", use_container_width=True):
                        if not f_data or not f_bloco or not f_material or not f_setor or f_qtd_ch is None:
                            st.error("Por favor, preencha todos os campos obrigatórios (*).")
                        else:
                            # Parse custom time format (e.g. 1040 -> 10:40)
                            def parse_time(t_str):
                                if not t_str: return None
                                import re
                                from datetime import time
                                t = re.sub(r'\D', '', t_str)
                                if len(t) == 1: t = "0" + t + "00"
                                elif len(t) == 2: t = t + "00"
                                elif len(t) == 3: t = "0" + t
                                elif len(t) > 4: t = t[:4]
                                
                                if len(t) == 4:
                                    h, m = int(t[:2]), int(t[2:])
                                    if 0 <= h <= 23 and 0 <= m <= 59:
                                        return time(h, m)
                                return None
                            
                            f_hora_ini = parse_time(f_hora_ini_str)
                            f_hora_fim = parse_time(f_hora_fim_str)

                            # Cálculo do tempo de processo
                            tempo_proc = ""
                            if f_hora_ini and f_hora_fim and f_dia_ini and f_dia_fim:
                                dt_ini = datetime.combine(f_dia_ini, f_hora_ini)
                                dt_fim = datetime.combine(f_dia_fim, f_hora_fim)
                                diff = dt_fim - dt_ini
                                if diff.total_seconds() > 0:
                                    hours, remainder = divmod(diff.total_seconds(), 3600)
                                    minutes, _ = divmod(remainder, 60)
                                    tempo_proc = f"{int(hours):02d}:{int(minutes):02d}"

                            # Cálculo de M2
                            comp_val = f_comp if f_comp else 0.0
                            alt_val = f_alt if f_alt else 0.0
                            ch_val = f_qtd_ch if f_qtd_ch else 0
                            calc_m2 = comp_val * alt_val * ch_val

                            novo_ap_dict = {
                                "DATA_REG": f_data.strftime("%d/%m/%Y"),
                                "BLOCO_RAW": f_bloco,
                                "NOME_MATERIAL": f_material.upper(),
                                "PROCESSO_APONTADO": f_processo,
                                "SETOR_AP": f_setor,
                                "QTD_CH": ch_val,
                                "QTD_M2": round(calc_m2, 3) if calc_m2 > 0 else 0.0,
                                "ESP": f_esp if f_esp else "",
                                "COMP": comp_val if comp_val > 0 else "",
                                "ALT": alt_val if alt_val > 0 else "",
                                "OPERADOR": f_operador.upper() if f_operador else "",
                                "DUREZA": f_dureza.upper() if f_dureza else "",
                                "DIA_INICIO": f_dia_ini.strftime("%d/%m/%Y") if f_dia_ini else "",
                                "DIA_FIM": f_dia_fim.strftime("%d/%m/%Y") if f_dia_fim else "",
                                "HORA_INICIO": f_hora_ini.strftime("%H:%M") if f_hora_ini else "",
                                "HORA_FIM": f_hora_fim.strftime("%H:%M") if f_hora_fim else "",
                                "TEMPO_PROCESSO": tempo_proc,
                                "TURNO": f_turno
                            }
                            novo_ap_dict.update(extra_data)
                            
                            if dm.add_apontamento(novo_ap_dict):
                                st.success(f"✅ Apontamento do bloco {f_bloco} gravado com sucesso! (Tempo Calculado: {tempo_proc if tempo_proc else 'Não aplicável'})")
                                st.balloons()
                            else:
                                st.error("Erro ao gravar apontamento no arquivo Excel.")
            else:
                st.info("Selecione um processo acima para abrir o formulário correspondente.")
        else:
            st.info("Selecione um Tipo de Processo acima para começar.")


    st.divider()

    c_ap1, c_ap2 = st.columns(2)
    with c_ap1:
        data_apontamento = st.date_input("Data de Produção", value=datetime.now(), format="DD/MM/YYYY", key="data_ap")
    with c_ap2:
        # Filtro de máquinas unificado: nossa base + apontamento
        setores_prog = set(str(x) for x in df["SETOR"].unique() if str(x) not in ["", "nan"])
        setores_ap_pre = set()
        if "ap_setores_extras" in st.session_state:
            setores_ap_pre = st.session_state["ap_setores_extras"]
        setores_ap_todos = sorted(setores_prog | setores_ap_pre)
        maq_ap = st.selectbox("Filtrar por Máquina", ["Todos"] + setores_ap_todos, key="maq_ap")

    if st.button("🔄 Carregar Apontamentos do Dia", key="btn_carregar_ap", type="primary"):
        st.session_state["ap_data"] = data_apontamento
        st.session_state["ap_maq"] = maq_ap

    if "ap_data" in st.session_state:
        data_ap_sel = st.session_state["ap_data"]
        maq_ap_sel  = st.session_state["ap_maq"]
        data_ap_str = data_ap_sel.strftime("%d/%m/%Y")

        with st.spinner(f"Lendo apontamentos de {data_ap_str}..."):
            df_ap = dm.get_apontamentos_do_dia(data_ap_sel)
        
        # Guardar máquinas extras do apontamento no session_state para o filtro
        if not df_ap.empty and "SETOR_AP" in df_ap.columns:
            setores_extras = set(str(x).strip() for x in df_ap["SETOR_AP"].unique() if str(x).strip() not in ["", "nan"])
            st.session_state["ap_setores_extras"] = setores_extras

        # Programação do dia (inclui já realizados para visão completa)
        def ap_date_ok(d_val):
            if pd.isna(d_val) or str(d_val).strip() in ["", "nan", "NaT"]: return False
            try:
                if isinstance(d_val, pd.Timestamp): return d_val.date() == data_ap_sel
                if "/" in str(d_val): return pd.to_datetime(str(d_val), format="%d/%m/%Y").date() == data_ap_sel
                return pd.to_datetime(str(d_val)).date() == data_ap_sel
            except: return False
        
        # Pendentes agendados para o dia
        df_pendentes = df[df["STATUS PROCESSO"] != "REALIZADO"].copy()
        df_pendentes = df_pendentes[df_pendentes["DATA"].apply(ap_date_ok)]
        
        # Já realizados cuja DATA REALIZADA é o dia selecionado
        df_realizados_dia = df[df["STATUS PROCESSO"] == "REALIZADO"].copy()
        def ap_data_real_ok(d_val):
            if pd.isna(d_val) or str(d_val).strip() in ["", "nan", "NaT"]: return False
            try:
                if isinstance(d_val, pd.Timestamp): return d_val.date() == data_ap_sel
                if "/" in str(d_val): return pd.to_datetime(str(d_val), format="%d/%m/%Y").date() == data_ap_sel
                return pd.to_datetime(str(d_val)).date() == data_ap_sel
            except: return False
        df_realizados_dia = df_realizados_dia[df_realizados_dia["DATA REALIZADA"].apply(ap_data_real_ok)]
        
        # Une os dois grupos
        df_prog_dia = pd.concat([df_pendentes, df_realizados_dia], ignore_index=False)
        
        if maq_ap_sel != "Todos":
            df_prog_dia = df_prog_dia[df_prog_dia["SETOR"] == maq_ap_sel]

        def norm_bloco(b):
            if pd.isna(b): return ""
            # Remove .0 de floats, remove espaços e coloca em maiúsculo
            return str(b).strip().split(".")[0].upper()

        # Mapa de bloco -> SETOR_AP do apontamento (para preencher máquinas sem SETOR)
        bloco_para_setor_ap = {}
        if not df_ap.empty and "SETOR_AP" in df_ap.columns:
            for _, r in df_ap.iterrows():
                b = norm_bloco(r["BLOCO"])
                s = str(r["SETOR_AP"]).strip()
                if b and s and s != "nan":
                    bloco_para_setor_ap[b] = s
        
        match_flags = []
        match_info  = []
        maquinas_resolvidas = []  # setor efetivo (nossa base ou apontamento)
        for _, row_prog in df_prog_dia.iterrows():
            b = norm_bloco(row_prog["BLOCO"])
            proc_prog = str(row_prog["PROCESSO"]).strip().upper()
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
                linhas_bloco = df_ap[df_ap["BLOCO"].apply(norm_bloco) == b]
                if not linhas_bloco.empty:
                    match_proc = any(
                        any(kw in str(r["RESUMIDO"]).upper() for kw in proc_prog.split())
                        for _, r in linhas_bloco.iterrows()
                    )
                    if match_proc:
                        ap_row = linhas_bloco.iloc[0]
                        match_flags.append("✅ Encontrado")
                        match_info.append(f"{ap_row['PROCESSO_APONTADO']} ({int(ap_row['QTD_CH'])} ch)")
                    else:
                        procs_ap = ", ".join(linhas_bloco["PROCESSO_APONTADO"].tolist())
                        match_flags.append("⚠️ Bloco sim, processo diferente")
                        match_info.append(procs_ap)
                else:
                    match_flags.append("❌ Não encontrado")
                    match_info.append("-")
            else:
                match_flags.append("❌ Sem apontamento")
                match_info.append("-")

        total_prog = len(df_prog_dia)
        total_ja_confirmados = sum(1 for f in match_flags if "🟢" in f)
        total_encontrados = sum(1 for f in match_flags if "✅" in f)
        total_confirmados_efetivo = total_ja_confirmados + total_encontrados
        aderencia = (total_confirmados_efetivo / total_prog * 100) if total_prog > 0 else 0

        st.write("---")
        col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
        col_m1.metric("Programados no Dia", total_prog)
        col_m2.metric("🟢 Já Confirmados", total_ja_confirmados)
        col_m3.metric("✅ No Apontamento", total_encontrados)
        col_m4.metric("❌ Não Apontados", total_prog - total_confirmados_efetivo)
        cor_delta = "normal" if aderencia >= 80 else "inverse"
        col_m5.metric("📊 Aderência", f"{aderencia:.1f}%",
                      delta=f"{"✅ Boa" if aderencia>=80 else "⚠️ Baixa"}", delta_color=cor_delta)
        st.progress(min(aderencia / 100, 1.0))

        st.write("---")
        if total_prog == 0:
            st.info(f"Nenhuma produção estava programada para {data_ap_str} na máquina selecionada.")
        else:
            st.markdown(f"**Confirme abaixo o que foi realizado em {data_ap_str}:**")
            df_confirm = pd.DataFrame()
            df_confirm["Confirmar?"]    = [True if "✅" in f else False for f in match_flags]
            df_confirm["Index"]         = df_prog_dia.index.tolist()
            df_confirm["Status"]        = match_flags
            df_confirm["Máquina"]       = maquinas_resolvidas  # Usa setor efetivo (nossa base ou apontamento)
            df_confirm["Bloco"]         = df_prog_dia["BLOCO"].tolist()
            df_confirm["Material"]      = df_prog_dia["MATERIAL"].tolist()
            df_confirm["Processo Prog."]= df_prog_dia["PROCESSO"].tolist()
            df_confirm["Chapas Prog."]  = pd.to_numeric(df_prog_dia["QTD. CHAPAS"], errors="coerce").fillna(0).astype(int).tolist()
            df_confirm["Apontado Como"] = match_info

            editado_confirm = st.data_editor(
                df_confirm,
                column_config={
                    "Confirmar?": st.column_config.CheckboxColumn("✅ Confirmar?", default=False),
                    "Index": None,
                    "Chapas Prog.": st.column_config.NumberColumn("Chapas Prog."),
                },
                hide_index=True, use_container_width=True,
                disabled=["Status", "Máquina", "Bloco", "Material", "Processo Prog.", "Chapas Prog.", "Apontado Como"],
                key="editor_confirm_ap"
            )

            confirmados_ap = editado_confirm[editado_confirm["Confirmar?"] == True]
            if st.button(
                f"✅ Marcar {len(confirmados_ap)} processo(s) como REALIZADO em {data_ap_str}",
                type="primary", disabled=len(confirmados_ap) == 0, key="btn_salvar_ap"
            ):
                erros_ap = []; sucessos_ap = 0
                with st.spinner("Salvando..."):
                    for idx_ap in confirmados_ap["Index"]:
                        if dm.update_cell_by_row(idx_ap, {"STATUS PROCESSO": "REALIZADO", "DATA REALIZADA": data_ap_str}):
                            sucessos_ap += 1
                        else:
                            erros_ap.append(str(df.loc[idx_ap, "BLOCO"]))
                if sucessos_ap > 0:
                    st.success(f"✅ {sucessos_ap} processo(s) marcado(s) como REALIZADO!")
                    del st.session_state["ap_data"]
                for e in erros_ap:
                    st.error(f"🛑 Erro ao salvar bloco {e}")
                if sucessos_ap > 0:
                    st.rerun()

        if not df_ap.empty:
            blocos_prog_norm = set(norm_bloco(b) for b in df_prog_dia["BLOCO"].tolist()) if total_prog > 0 else set()
            df_ap_extras = df_ap[~df_ap["BLOCO"].apply(norm_bloco).isin(blocos_prog_norm)]
            if not df_ap_extras.empty:
                with st.expander(f"📋 {len(df_ap_extras)} apontamento(s) sem programação correspondente", expanded=False):
                    st.dataframe(df_ap_extras[["BLOCO", "PROCESSO_APONTADO", "SETOR_AP", "QTD_CH"]],
                                 hide_index=True, use_container_width=True)


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

# ----------------- ABA 6: OPÇÕES GERAIS -----------------
with tab_config:
    st.header("⚙️ Opções Gerais")

    # ---- SEÇÃO 1: CAMINHOS DAS BASES DE DADOS ----
    st.subheader("📁 Caminhos e Abas das Bases de Dados")
    st.markdown("Configure os arquivos e **qual aba** de cada planilha o sistema deve usar. Após salvar, atualize a página.")

    cfg_atual = dm.get_config()

    # Função auxiliar para abrir dialog do Windows
    def gui_select_file(current_path, is_excel=True):
        import tkinter as tk
        from tkinter import filedialog
        import os
        try:
            root = tk.Tk()
            root.withdraw()
            root.wm_attributes('-topmost', 1)
            init_dir = os.path.dirname(current_path) if current_path and os.path.exists(current_path) else "."
            types = [("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")] if is_excel else [("All files", "*.*")]
            
            file_path = filedialog.askopenfilename(
                initialdir=init_dir,
                title="Selecione o Arquivo",
                filetypes=types
            )
            root.destroy()
            return file_path.replace("/", "\\") if file_path else ""
        except Exception as e:
            return ""

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

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 Salvar Configurações", type="primary", use_container_width=True, key="btn_salvar_cfg"):
            erros_path = []
            if not _osc.path.exists(novo_db):
                erros_path.append(f"❌ Arquivo de Programação não encontrado: `{novo_db}`")
            if not _osc.path.exists(novo_ap):
                erros_path.append(f"❌ Arquivo de Apontamento não encontrado: `{novo_ap}`")

            if erros_path:
                for ep in erros_path:
                    st.error(ep)
            else:
                novo_cfg = dict(cfg_atual)
                novo_cfg["DB_FILE"] = novo_db
                novo_cfg["APONTAMENTO_FILE"] = novo_ap
                novo_cfg["SHEET_PROGRAMACAO"] = aba_prog
                novo_cfg["SHEET_ENTREGUES"] = aba_entregues
                novo_cfg["SHEET_BASE_DADOS"] = aba_base
                novo_cfg["SHEET_AP_BD"] = aba_ap_bd
                novo_cfg["SHEET_AP_BASE"] = aba_ap_base
                if dm.save_config(novo_cfg):
                    st.success("✅ Configurações salvas! Atualize a página (F5) para aplicar.")
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

