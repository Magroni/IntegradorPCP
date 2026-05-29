import os

filepath = r"z:\PCP\PROJETOS MARLON\ProgramarProd\app.py"

with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

# Let's search for "size: A3 landscape;"
target = "size: A3 landscape;"
if target in content:
    content = content.replace(target, "")
    print("Found and removed 'size: A3 landscape;'!")
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully updated app.py!")
else:
    # Try with double curly braces representation if it was parsed as string
    target_double = "size: A3 landscape;"
    print("Could not find literal target. Printing occurrences of '@page':")
    lines = content.split('\n')
    for idx, line in enumerate(lines):
        if "@page" in line or "size: A3" in line:
            print(f"Line {idx+1}: {repr(line)}")
