import os

with open('app.py', 'r', encoding='utf-8') as f:
    c = f.read()

# 1. Update the A3 layout controls to have columns layout and the slider
old_controls = """                            # --- SEÇÃO: RELATÓRIO A3 LEAN ---
                            st.write("---")
                            st.subheader("🖨️ Relatório Formal A3 de Desempenho (Lean Manufacturing)")
                            st.markdown("Gere um relatório operacional de inatividade e impactos financeiros estruturado no tradicional **Layout A3 Landscape (420mm x 297mm)**, ideal para apresentações formais de diretoria, impressão física ou salvamento em PDF.")
                            
                            show_a3 = st.checkbox("📂 Visualizar Relatório A3 de Paradas e Ociosidade", value=False, key="chk_relatorio_a3")
                            show_bw = st.checkbox("🕶️ Modo Alto Contraste (Preto e Branco para Impressão)", value=False, key="chk_bw_print")
                            analisar_custos = st.checkbox("💸 Analisar Custos Financeiros", value=False, key="chk_analisar_custos")
                            
                            if show_a3:"""

new_controls = """                            # --- SEÇÃO: RELATÓRIO A3 LEAN ---
                            st.write("---")
                            st.subheader("🖨️ Relatório Formal A3 de Desempenho (Lean Manufacturing)")
                            st.markdown("Gere um relatório operacional de inatividade e impactos financeiros estruturado no tradicional **Layout A3 Landscape (420mm x 297mm)**, ideal para apresentações formais de diretoria, impressão física ou salvamento em PDF.")
                            
                            col_ctrl1, col_ctrl2 = st.columns(2)
                            with col_ctrl1:
                                show_a3 = st.checkbox("📂 Visualizar Relatório A3 de Paradas e Ociosidade", value=False, key="chk_relatorio_a3")
                                show_bw = st.checkbox("🕶️ Modo Alto Contraste (Preto e Branco para Impressão)", value=False, key="chk_bw_print")
                                analisar_custos = st.checkbox("💸 Analisar Custos Financeiros", value=False, key="chk_analisar_custos")
                            with col_ctrl2:
                                if show_a3:
                                    top_n_maq_cfg = st.slider("Motivos por Máquina (Detalhamento)", min_value=1, max_value=5, value=2, step=1, key="slider_top_n_maq")
                                else:
                                    top_n_maq_cfg = 3
                            
                            if show_a3:"""

c = c.replace(old_controls, new_controls)

# 2. Update maquinas_paradas_rows_html to use top_n_maq_cfg
old_top_n = "                                        top_n_maq = 3"
new_top_n = "                                        top_n_maq = top_n_maq_cfg"
c = c.replace(old_top_n, new_top_n)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(c)

print("A3 slider configuration for Detailed stops complete!")
