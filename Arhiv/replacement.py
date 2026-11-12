""" заменять только целые отдельные числа (а не подстроки внутри других)"""
import re

input_path = "repl_1.txt"
output_path = "repl_upd.txt"

replacements = {
    r"\b234567\b": "723456",
    r"\b345678\b": "823456",
    r"\b456789\b": "923456",
}


with open(input_path, "r", encoding="utf-8") as f:
    content=f.read()

updated_content = content
for pattern, new in replacements.items():
    updated_content = re.sub(pattern, new, updated_content)

with open(output_path, "w", encoding="utf-8") as f:
    f.write(updated_content)
