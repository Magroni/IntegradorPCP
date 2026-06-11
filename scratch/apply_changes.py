import os

with open('app.py', 'r', encoding='utf-8') as f:
    c = f.read()

# 1. Update c_motivos groupby to include TIPO_PARADA
old_groupby = 'c_motivos = df_paradas_filtrado.groupby(["MOTIVO", c_st])[["TEMPO", "PREJUIZO"]].sum().reset_index()'
new_groupby = 'c_motivos = df_paradas_filtrado.groupby(["MOTIVO", c_st, "TIPO_PARADA"])[["TEMPO", "PREJUIZO"]].sum().reset_index()'
c = c.replace(old_groupby, new_groupby)

# 2. Update meta-item Filtros Ativos with meta_ociosidade
old_meta = '                                            <div class="meta-item"><b>Filtros Ativos:</b> {len(maquinas_parada_sel)} Máqs. / {len(motivos_sel)} Motivos</div>'
new_meta = '                                            <div class="meta-item">{meta_ociosidade}</div>'
c = c.replace(old_meta, new_meta)

# 3. Dynamic row columns for setores_rows_html based on analisar_custos
old_setores_row = """                                    setores_rows_html += f\"\"\"
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
                                    </tr>\"\"\""""

new_setores_row = """                                    custo_cells_html = f"<td style='text-align:center; font-weight:600; color:#1e3a8a;'>R$ {c_val:,.2f}/h</td><td style='text-align:right; font-weight:600; color:#b91c1c;'>{prej_fmt}</td>" if analisar_custos else ""
                                    setores_rows_html += f\"\"\"
                                    <tr>
                                        <td style='font-weight:600; color:#334155;'>{maq_name}</td>
                                        <td style='text-align:center;'>{r['TEMPO_HHMM']}</td>
                                        {custo_cells_html}
                                        <td style='width:100px;'>
                                            <div style='display:flex; justify-content:space-between; font-size:8.5px; color:#64748b; margin-bottom:1px;'>
                                                <span>{pct:.1f}%</span>
                                            </div>
                                            <div class="progress-container">
                                                <div class="progress-bar" style="width: {pct}%; background: {cores_s};"></div>
                                            </div>
                                        </td>
                                    </tr>\"\"\""""

c = c.replace(old_setores_row, new_setores_row)

# 4. Remove lean_severity_cards_html calculations completely (or keep it if needed, but we don't need it since Section 9 is removed)
# We can just leave lean_severity_cards_html calculations as is, but we will not render the section.

# 5. Insert maquinas_paradas_rows_html logic right after setores_rows_html loop ends
old_severity_start = '                                # F4. LEAN SEVERIDADE BREAKDOWN (Farol Lean: Operacional, Intervenção, Crítica)'

new_detalhamento_logic = """                                # D2. DETALHAMENTO DE PARADAS POR MÁQUINA E MOTIVO
                                maquinas_paradas_rows_html = ""
                                if not c_setores.empty and not c_motivos.empty:
                                    setores_tempo_total = c_setores.sort_values("TEMPO", ascending=False)
                                    cores_tableau = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                                    
                                    for idx_s, row_s in setores_tempo_total.iterrows():
                                        maq_name = row_s[c_st]
                                        maq_tempo_total = row_s["TEMPO"]
                                        maq_tempo_hhmm = row_s["TEMPO_HHMM"]
                                        
                                        try:
                                            color_idx = setores_ordenados.index(maq_name) % len(cores_tableau)
                                            cores_s = cores_tableau[color_idx]
                                        except:
                                            cores_s = "#3b82f6"
                                            
                                        colspan_val = 4 if analisar_custos else 3
                                        maquinas_paradas_rows_html += f\"\"\"
                                        <tr style="background-color: #f8fafc; font-weight: 700; border-top: 1.5px solid #cbd5e1;">
                                            <td colspan="{colspan_val}" style="color: #1e3a8a; font-size: 9.5px; padding: 4px 6px;">
                                                🛠️ {maq_name} <span style="font-weight: normal; color: #64748b; font-size: 8.5px;">(Tempo Total Parado: {maq_tempo_hhmm})</span>
                                            </td>
                                        </tr>\"\"\"
                                        
                                        df_motivos_maq = c_motivos[c_motivos[c_st] == maq_name].copy()
                                        df_motivos_maq = df_motivos_maq.sort_values("TEMPO", ascending=False)
                                        
                                        top_n_maq = 3
                                        df_top_maq = df_motivos_maq.head(top_n_maq)
                                        df_others_maq = df_motivos_maq.iloc[top_n_maq:]
                                        
                                        for idx_m, row_m in df_top_maq.iterrows():
                                            pct_maq = (row_m["TEMPO"] / maq_tempo_total * 100) if maq_tempo_total > 0 else 0
                                            prej_fmt = f"R$ {row_m['PREJUIZO']:,.2f}"
                                            tipo_p = row_m.get("TIPO_PARADA", "")
                                            
                                            badge_html = ""
                                            if tipo_p == "Operacional":
                                                badge_html = ' <span style="background-color: #f0fdf4; color: #16a34a; border: 1px solid #dcfce7; padding: 1px 4px; border-radius: 3px; font-size: 7.5px; margin-left: 6px; font-weight: bold; text-transform: uppercase;">Operacional</span>'
                                            elif tipo_p == "Intervenção":
                                                badge_html = ' <span style="background-color: #fffbeb; color: #d97706; border: 1px solid #fef3c7; padding: 1px 4px; border-radius: 3px; font-size: 7.5px; margin-left: 6px; font-weight: bold; text-transform: uppercase;">Intervenção</span>'
                                            elif tipo_p == "Crítica":
                                                badge_html = ' <span style="background-color: #fef2f2; color: #ef4444; border: 1px solid #fee2e2; padding: 1px 4px; border-radius: 3px; font-size: 7.5px; margin-left: 6px; font-weight: bold; text-transform: uppercase;">Crítica</span>'
                                            
                                            cost_cell_html = f"<td style='text-align:right; font-weight:600; color:#1e3a8a;'>{prej_fmt}</td>" if analisar_custos else ""
                                            maquinas_paradas_rows_html += f\"\"\"
                                            <tr>
                                                <td style="padding-left: 15px; color: #475569; font-weight: 500;">{row_m['MOTIVO']}{badge_html}</td>
                                                <td style="text-align:center;">{row_m['TEMPO_HHMM']}</td>
                                                {cost_cell_html}
                                                <td style="width:100px;">
                                                    <div style="display:flex; justify-content:space-between; font-size:8px; color:#64748b; margin-bottom:1px;">
                                                        <span>{pct_maq:.1f}%</span>
                                                    </div>
                                                    <div class="progress-container">
                                                        <div class="progress-bar" style="width: {pct_maq}%; background: {cores_s};"></div>
                                                    </div>
                                                </td>
                                            </tr>\"\"\"
                                            
                                        if not df_others_maq.empty:
                                            others_tempo = df_others_maq["TEMPO"].sum()
                                            others_prej = df_others_maq["PREJUIZO"].sum()
                                            others_tempo_hhmm = format_to_hhmm(others_tempo)
                                            pct_others = (others_tempo / maq_tempo_total * 100) if maq_tempo_total > 0 else 0
                                            
                                            cost_cell_html = f"<td style='text-align:right; color: #94a3b8; font-weight:600;'>R$ {others_prej:,.2f}</td>" if analisar_custos else ""
                                            maquinas_paradas_rows_html += f\"\"\"
                                            <tr style="font-style: italic; background-color: #fafafa;">
                                                <td style="padding-left: 15px; color: #94a3b8; font-weight: 500;">OUTROS ({len(df_others_maq)} motivos)</td>
                                                <td style="text-align:center; color: #94a3b8;">{others_tempo_hhmm}</td>
                                                {cost_cell_html}
                                                <td style="width:100px;">
                                                    <div style="display:flex; justify-content:space-between; font-size:8px; color:#94a3b8; margin-bottom:1px;">
                                                        <span>{pct_others:.1f}%</span>
                                                    </div>
                                                    <div class="progress-container">
                                                        <div class="progress-bar" style="width: {pct_others}%; background: #94a3b8;"></div>
                                                    </div>
                                                </td>
                                            </tr>\"\"\"
                                else:
                                    colspan_val = 4 if analisar_custos else 3
                                    maquinas_paradas_rows_html = f"<tr><td colspan='{colspan_val}' style='text-align:center; color:#64748b;'>Nenhuma parada registrada</td></tr>"

                                # F4. LEAN SEVERIDADE BREAKDOWN (Farol Lean: Operacional, Intervenção, Crítica)"""

c = c.replace(old_severity_start, new_detalhamento_logic)

# 6. Remove KPI cards block from Right Column in the HTML layout
old_kpis_sec = """                                            <!-- Coluna Direita: Paradas & Inatividade -->
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
                                                    <div class="section-title">7. Pareto de Paradas por Motivo</div>"""

new_kpis_sec = """                                            <!-- Coluna Direita: Paradas & Inatividade -->
                                            <div class="column-section" style="display:flex; flex-direction:column; justify-content:space-between; gap:6px;">
                                                <div>
                                                    <div class="section-title">6. Pareto de Paradas por Motivo</div>"""

c = c.replace(old_kpis_sec, new_kpis_sec)

# 7. Dynamic columns headers in Pareto table (Motivos)
old_motivo_header = """                                                    <table class="a3-table">
                                                        <thead>
                                                            <tr>
                                                                <th>Motivo da Parada</th>
                                                                <th style='text-align:center; width:60px;'>Duração</th>
                                                                <th style='text-align:right; width:100px;'>Custo (R$)</th>
                                                                <th style='text-align:left; width:110px;'>Representação (%)</th>
                                                            </tr>
                                                        </thead>"""

new_motivo_header = """                                                    <table class="a3-table">
                                                        <thead>
                                                            <tr>
                                                                <th>Motivo da Parada</th>
                                                                <th style='text-align:center; width:60px;'>Duração</th>
                                                                {custo_header_html}
                                                                <th style='text-align:left; width:110px;'>Representação (%)</th>
                                                            </tr>
                                                        </thead>"""

c = c.replace(old_motivo_header, new_motivo_header)

# 8. Dynamic columns and header renaming in Pareto por Máquina table
old_setores_header = """                                                    <div class="section-title">8. Ocupação de Máquina &amp; Impacto Financeiro Mapeado</div>
                                                    <table class="a3-table">
                                                        <thead>
                                                            <tr>
                                                                <th>Máquina / Setor</th>
                                                                <th style='text-align:center;'>Duração</th>
                                                                <th style='text-align:center;'>Taxa (R$/h)</th>
                                                                <th style='text-align:right;'>Prejuízo</th>
                                                                <th style='text-align:left; width:110px;'>Representação (%)</th>
                                                            </tr>
                                                        </thead>"""

new_setores_header = """                                                    <div class="section-title">7. Pareto de Paradas por Máquina</div>
                                                    <table class="a3-table">
                                                        <thead>
                                                            <tr>
                                                                <th>Máquina / Setor</th>
                                                                <th style='text-align:center;'>Duração</th>
                                                                {setores_custo_headers}
                                                                <th style='text-align:left; width:110px;'>Representação (%)</th>
                                                            </tr>
                                                        </thead>"""

c = c.replace(old_setores_header, new_setores_header)

# 9. Insert Section 8 HTML and Remove Section 9 HTML
old_lean_sec_and_bottom = """                                                <div>
                                                    <div class="section-title">9. Análise Lean: Severidade e Frequência das Paradas</div>
                                                    <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 6px; margin-top: 2px;">
                                                        {lean_severity_cards_html}
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                        
                                        <!-- Seção Inferior: Ocorrências Críticas (Top 5) -->
                                        <div class="bottom-section" style="margin-bottom:0; padding: 6px 10px;">
                                            <div class="section-title" style="margin-bottom:4px;">10. Ocorrências Mais Críticas de Paradas (Top 5 por Duração)</div>"""

new_lean_sec_and_bottom = """                                                <div>
                                                    <div class="section-title">8. Detalhamento de Paradas por Máquina e Motivo</div>
                                                    <table class="a3-table">
                                                        <thead>
                                                            <tr>
                                                                <th>Máquina / Motivo da Parada</th>
                                                                <th style='text-align:center; width:60px;'>Duração</th>
                                                                {custo_header_html}
                                                                <th style='text-align:left; width:110px;'>Representação (%)</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {maquinas_paradas_rows_html}
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </div>
                                        </div>
                                        
                                        <!-- Seção Inferior: Ocorrências Críticas (Top 5) -->
                                        <div class="bottom-section" style="margin-bottom:0; padding: 6px 10px;">
                                            <div class="section-title" style="margin-bottom:4px;">9. Ocorrências Mais Críticas de Paradas (Top 5 por Duração)</div>"""

c = c.replace(old_lean_sec_and_bottom, new_lean_sec_and_bottom)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(c)

print("A3 layout KPI removal, merge 8 and 9, and renumbering complete!")
