import os

filepath = r"z:\PCP\PROJETOS MARLON\ProgramarProd\app.py"

with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

# Let's search for scriptToAdd
found = False
lines = content.split('\n')
for idx, line in enumerate(lines):
    if "var scriptToAdd =" in line:
        print(f"Found line {idx+1}: {repr(line)}")
        # Replace the script tag
        lines[idx] = line.replace(r'<\/script>', r"<' + '/script>")
        found = True
    if "htmlContent.replace('</body>'" in line:
        print(f"Found body line {idx+1}: {repr(line)}")
        lines[idx] = line.replace("'</body>'", "'<' + '/body>'")
        found = True

if found:
    new_content = '\n'.join(lines)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(new_content)
    print("Successfully updated app.py!")
else:
    print("Could not find the target lines.")
