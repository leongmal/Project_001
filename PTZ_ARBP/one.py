
unique_lines = set()
duble = []

with open('number_pribor.txt', 'r', encoding = 'utf-8') as f:
    for line in f:
        cleaned_line = line.strip()
        if cleaned_line in unique_lines:
            duble.append(cleaned_line.rstrip('\n'))
            
        else:
            unique_lines.add(cleaned_line)
print(line.strip())
print(len(unique_lines))
print(duble)

