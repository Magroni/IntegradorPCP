import os

filepath = r"z:\PCP\PROJETOS MARLON\ProgramarProd\app.py"

with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Restore button
old_button = '<button id="btnPrintA3" class="btn-print-a3">🖨️ Imprimir / Salvar PDF (A3 / A4)</button>'
new_button = '<button class="btn-print-a3" onclick="window.print()">🖨️ Imprimir / Salvar PDF (A3 / A4)</button>'

if old_button in content:
    content = content.replace(old_button, new_button)
    print("Restored the button to standard window.print()!")

# 2. Remove the JavaScript script block at the end of body
# Let's locate the script block. We'll split the content and remove lines from <script> to </script> inside html_a3's body end
# A safer way is to find the exact target string and replace it.
# Let's define the exact target script text we wrote:
target_script_block = """                                    
                                    <script>
                                        function printInNewTab() {
                                            var printWindow = window.open('', '_blank');
                                            if (!printWindow) {
                                                alert('Por favor, permita pop-ups para este site para visualizar e imprimir o relatório.');
                                                return;
                                            }
                                            var htmlContent = document.documentElement.outerHTML;
                                            var scriptToAdd = '<script>window.onload = function() { setTimeout(function() { window.print(); }, 500); };<' + '/script>';
                                            htmlContent = htmlContent.replace('<' + '/body>', scriptToAdd + '<' + '/body>');
                                            
                                            printWindow.document.write(htmlContent);
                                            printWindow.document.close();
                                        }

                                        function initPrintButton() {
                                            var btn = document.getElementById('btnPrintA3');
                                            if (btn) {
                                                if (window.self !== window.top) {
                                                    btn.innerHTML = "🖨️ Abrir Relatório em Nova Aba (Evita Corte e Força A3 Paisagem)";
                                                    btn.onclick = printInNewTab;
                                                } else {
                                                    btn.innerHTML = "🖨️ Imprimir / Salvar PDF (A3 Paisagem)";
                                                    btn.onclick = function() { window.print(); };
                                                }
                                            }
                                        }
                                        if (document.readyState === 'loading') {
                                            window.addEventListener('DOMContentLoaded', initPrintButton);
                                        } else {
                                            initPrintButton();
                                        }
                                    </script>"""

# Since the content has double curly braces in python file, let's search with double curly braces!
target_script_block_double = """                                    
                                    <script>
                                        function printInNewTab() {{
                                            var printWindow = window.open('', '_blank');
                                            if (!printWindow) {{
                                                alert('Por favor, permita pop-ups para este site para visualizar e imprimir o relatório.');
                                                return;
                                            }}
                                            var htmlContent = document.documentElement.outerHTML;
                                            var scriptToAdd = '<script>window.onload = function() {{ setTimeout(function() {{ window.print(); }}, 500); }};<' + '/script>';
                                            htmlContent = htmlContent.replace('<' + '/body>', scriptToAdd + '<' + '/body>');
                                            
                                            printWindow.document.write(htmlContent);
                                            printWindow.document.close();
                                        }}

                                        function initPrintButton() {{
                                            var btn = document.getElementById('btnPrintA3');
                                            if (btn) {{
                                                if (window.self !== window.top) {{
                                                    btn.innerHTML = "🖨️ Abrir Relatório em Nova Aba (Evita Corte e Força A3 Paisagem)";
                                                    btn.onclick = printInNewTab;
                                                }} else {{
                                                    btn.innerHTML = "🖨️ Imprimir / Salvar PDF (A3 Paisagem)";
                                                    btn.onclick = function() {{ window.print(); }};
                                                }}
                                            }}
                                        }}
                                        if (document.readyState === 'loading') {{
                                            window.addEventListener('DOMContentLoaded', initPrintButton);
                                        }} else {{
                                            initPrintButton();
                                        }}
                                    </script>"""

if target_script_block_double in content:
    content = content.replace(target_script_block_double, "")
    print("Successfully removed the custom JavaScript printing code!")
else:
    # If indentation is slightly different, let's split by lines and rebuild
    lines = content.split('\n')
    start_idx = -1
    end_idx = -1
    for idx, line in enumerate(lines):
        if "function printInNewTab()" in line:
            start_idx = idx - 1 # Include <script>
        if "initPrintButton();" in line and start_idx != -1:
            end_idx = idx + 2 # Include </script>
            break
            
    if start_idx != -1 and end_idx != -1:
        print(f"Found script block from line {start_idx+1} to {end_idx+1}. Removing it.")
        del lines[start_idx:end_idx]
        content = '\n'.join(lines)

# 3. Update st.info message
old_info = 'st.info("💡 **Dica de Impressão:** Ao clicar no botão acima, o relatório será aberto em uma **nova aba** (o que resolve o problema de corte do Streamlit e força o layout A3 Paisagem/Landscape no Microsoft Print to PDF). Na tela de impressão que se abrirá automaticamente, selecione **Salvar como PDF** (ou **Microsoft Print to PDF**) e confirme se o tamanho está como **A3** (ou **A4** com a opção **Ajustar à página** ativada) e o layout como **Paisagem (Landscape)**.")'
new_info = 'st.info("💡 **Dica de Impressão:** Ao clicar no botão acima, a tela de impressão do navegador será aberta diretamente. Selecione a impressora desejada (como **Microsoft Print to PDF** ou **Salvar como PDF**), mude a orientação (Layout) para **Paisagem (Landscape)**, defina o tamanho do papel como **A3** (ou **A4** com a opção **Ajustar à página** ativa) e ajuste a escala para preencher a folha perfeitamente!")'

if old_info in content:
    content = content.replace(old_info, new_info)
    print("Updated st.info text successfully!")
else:
    # Fallback search
    for idx, line in enumerate(lines):
        if 'st.info("💡 **Dica de Impressão:**' in line:
            lines[idx] = '                                st.info("💡 **Dica de Impressão:** Ao clicar no botão acima, a tela de impressão do navegador será aberta diretamente. Selecione a impressora desejada (como **Microsoft Print to PDF** ou **Salvar como PDF**), mude a orientação (Layout) para **Paisagem (Landscape)**, defina o tamanho do papel como **A3** (ou **A4** com a opção **Ajustar à página** ativa) e ajuste a escala para preencher a folha perfeitamente!")'
            content = '\n'.join(lines)
            print("Fallback updated st.info text successfully!")
            break

with open(filepath, "w", encoding="utf-8") as f:
    f.write(content)
print("Done!")
