import re

file_path = "z:/PCP/PROJETOS MARLON/ProgramarProd/app.py"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Procura o ponto onde as colunas foram cortadas
target = r'(\s*<td style="font-weight:700; color:#EF553B; text-align:center;">\{duration_str\}</td>)\r?\n\s*html_a3 = f"""'

replacement = r'\1\n                                         <td style="font-weight:700; color:#1e3a8a; text-align:right;">{prej_fmt}</td>\n                                     </tr>"""\n                                    \n                                html_a3 = f"""'

new_content, count = re.subn(target, replacement, content)
print(f"Substituições realizadas: {count}")

if count > 0:
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    print("Arquivo app.py atualizado com sucesso!")
else:
    print("Falha ao localizar o trecho para substituição.")
