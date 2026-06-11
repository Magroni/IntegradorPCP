import os

with open('app.py', 'r', encoding='utf-8') as f:
    c = f.read()

old_block = """                            # --- SEÇÃO: RELATÓRIO A3 LEAN ---
                            # Dynamic headers and columns based on analisar_custos
                            custo_header_html = "<th style='text-align:right; width:100px;'>Custo (R$)</th>" if analisar_custos else ""
                            setores_custo_headers = "<th style='text-align:center;'>Taxa (R$/h)</th><th style='text-align:right;'>Prejuízo</th>" if analisar_custos else ""
                            ocorrencias_custo_header = "<th style='text-align:right; width:120px;'>Prejuízo Individual</th>" if analisar_custos else ""
                            
                            if analisar_custos:
                                meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} (Prejuízo Est.: R$ {prejuizo_estimado:,.2f})"
                            else:
                                meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} ({len(df_paradas_filtrado)} ocor.)"
                                
                            st.write("---")
                            st.subheader("🖨️ Relatório Formal A3 de Desempenho (Lean Manufacturing)")
                            st.markdown("Gere um relatório operacional de inatividade e impactos financeiros estruturado no tradicional **Layout A3 Landscape (420mm x 297mm)**, ideal para apresentações formais de diretoria, impressão física ou salvamento em PDF.")
                            
                            show_a3 = st.checkbox("📂 Visualizar Relatório A3 de Paradas e Ociosidade", value=False, key="chk_relatorio_a3")
                            show_bw = st.checkbox("🕶️ Modo Alto Contraste (Preto e Branco para Impressão)", value=False, key="chk_bw_print")
                            if show_a3:"""

new_block = """                            # --- SEÇÃO: RELATÓRIO A3 LEAN ---
                            st.write("---")
                            st.subheader("🖨️ Relatório Formal A3 de Desempenho (Lean Manufacturing)")
                            st.markdown("Gere um relatório operacional de inatividade e impactos financeiros estruturado no tradicional **Layout A3 Landscape (420mm x 297mm)**, ideal para apresentações formais de diretoria, impressão física ou salvamento em PDF.")
                            
                            show_a3 = st.checkbox("📂 Visualizar Relatório A3 de Paradas e Ociosidade", value=False, key="chk_relatorio_a3")
                            show_bw = st.checkbox("🕶️ Modo Alto Contraste (Preto e Branco para Impressão)", value=False, key="chk_bw_print")
                            analisar_custos = st.checkbox("💸 Analisar Custos Financeiros", value=False, key="chk_analisar_custos")
                            
                            if show_a3:
                                # Dynamic headers and columns based on analisar_custos
                                custo_header_html = "<th style='text-align:right; width:100px;'>Custo (R$)</th>" if analisar_custos else ""
                                setores_custo_headers = "<th style='text-align:center;'>Taxa (R$/h)</th><th style='text-align:right;'>Prejuízo</th>" if analisar_custos else ""
                                ocorrencias_custo_header = "<th style='text-align:right; width:120px;'>Prejuízo Individual</th>" if analisar_custos else ""
                                
                                if analisar_custos:
                                    meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} (Prejuízo Est.: R$ {prejuizo_estimado:,.2f})"
                                else:
                                    meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} ({len(df_paradas_filtrado)} ocor.)\""""

c = c.replace(old_block, new_block)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(c)

print("analisar_custos toggle switch added and moved inside A3 visibility block!")
