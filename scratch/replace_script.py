with open('app.py', 'r', encoding='utf-8') as f:
    c = f.read()

# 1. Define meta_ociosidade and remove kpi2_html definition
old_kpi2_block = """                            if analisar_custos:
                                kpi2_html = f'''
                                <div class="kpi-card" style="border-color:#fee2e2; background:#fef2f2;">
                                    <div class="kpi-lbl" style="color:#ef4444;">Prejuízo Estimado</div>
                                    <div class="kpi-val" style="color:#ef4444; font-size:15px;">R$ {prejuizo_estimado:,.2f}</div>
                                </div>'''
                            else:
                                total_ocorrencias = len(df_paradas_filtrado)
                                kpi2_html = f'''
                                <div class="kpi-card">
                                    <div class="kpi-lbl">Total Ocorrências</div>
                                    <div class="kpi-val" style="font-size:15px;">{total_ocorrencias}</div>
                                </div>'''"""

new_kpi2_block = """                            if analisar_custos:
                                meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} (Prejuízo Est.: R$ {prejuizo_estimado:,.2f})"
                            else:
                                meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} ({len(df_paradas_filtrado)} ocor.)\""""

# Let's do exact text replaces on app.py:

# Clean up variables block
c = c.replace(
    '                            if analisar_custos:\n                                kpi2_html = f\'\'\'\n                                <div class="kpi-card" style="border-color:#fee2e2; background:#fef2f2;">\n                                    <div class="kpi-lbl" style="color:#ef4444;">Prejuízo Estimado</div>\n                                    <div class="kpi-val" style="color:#ef4444; font-size:15px;">R$ {prejuizo_estimado:,.2f}</div>\n                                </div>\'\'\'\n                            else:\n                                total_ocorrencias = len(df_paradas_filtrado)\n                                kpi2_html = f\'\'\'\n                                <div class="kpi-card">\n                                    <div class="kpi-lbl">Total Ocorrências</div>\n                                    <div class="kpi-val" style="font-size:15px;">{total_ocorrencias}</div>\n                                </div>\'\'\'',
    '                            if analisar_custos:\n                                meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} (Prejuízo Est.: R$ {prejuizo_estimado:,.2f})"\n                            else:\n                                meta_ociosidade = f"<b>Inatividade:</b> {format_to_hhmm(tempo_tot_min)} ({len(df_paradas_filtrado)} ocor.)"'
)

# Header metadata replacement
c = c.replace(
    '                                            <div class="meta-item"><b>Filtros Ativos:</b> {len(maquinas_parada_sel)} Máqs. / {len(motivos_sel)} Motivos</div>',
    '                                            <div class="meta-item">{meta_ociosidade}</div>'
)

# KPIs section removal
old_kpis_sec = """                                            <!-- Coluna Direita: Paradas & Inatividade -->
                                            <div class="column-section" style="display:flex; flex-direction:column; justify-content:space-between; gap:6px;">
                                                <div>
                                                    <div class="section-title">6. KPIs de Paradas &amp; Ociosidade</div>
                                                    <div class="kpi-grid">
                                                        <div class="kpi-card" style="border-color:#fee2e2; background:#fef2f2;">
                                                            <div class="kpi-lbl" style="color:#ef4444;">Tempo Total Ocioso</div>
                                                            <div class="kpi-val" style="color:#ef4444; font-size:15px;">{format_to_hhmm(tempo_tot_min)}</div>
                                                        </div>
                                                        {kpi2_html}
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
                                                    <div class="section-title">7. Pareto de Paradas por Motivo</div>"""

new_kpis_sec = """                                            <!-- Coluna Direita: Paradas & Inatividade -->
                                            <div class="column-section" style="display:flex; flex-direction:column; justify-content:space-between; gap:6px;">
                                                <div>
                                                    <div class="section-title">6. Pareto de Paradas por Motivo</div>"""

c = c.replace(old_kpis_sec, new_kpis_sec)

# Renumber other titles
c = c.replace('<div class="section-title">8. Pareto de Paradas por Máquina</div>', '<div class="section-title">7. Pareto de Paradas por Máquina</div>')
c = c.replace('<div class="section-title">9. Detalhamento de Paradas por Máquina e Motivo</div>', '<div class="section-title">8. Detalhamento de Paradas por Máquina e Motivo</div>')
c = c.replace('<div class="section-title">10. Análise Lean: Severidade e Frequência das Paradas</div>', '<div class="section-title">9. Análise Lean: Severidade e Frequência das Paradas</div>')
c = c.replace('<div class="section-title" style="margin-bottom:4px;">11. Ocorrências Mais Críticas de Paradas (Top 5 por Duração)</div>', '<div class="section-title" style="margin-bottom:4px;">10. Ocorrências Mais Críticas de Paradas (Top 5 por Duração)</div>')

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(c)

print("A3 layout KPI removal and renumbering complete!")
