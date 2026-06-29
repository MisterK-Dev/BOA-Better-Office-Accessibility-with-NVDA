import os
import re

for root, dirs, files in os.walk('addon/'):
    for file in files:
        if file.endswith('.py'):
            filepath = os.path.join(root, file)
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            changed = False
            new_lines = []
            for line in lines:
                if 'except Exception: pass' in line:
                    indent = line[:len(line) - len(line.lstrip('\t '))]
                    new_lines.append(indent + 'except Exception:\n')
                    new_lines.append(indent + '\tpass\n')
                    changed = True
                elif 'if ' in line and ':' in line and not line.rstrip().endswith(':'):
                    # Be careful not to break dictionary comprehensions or lambda or inline if
                    if re.match(r'^[\t ]*(if|elif|else).*:\s*[a-zA-Z_]', line):
                        parts = line.split(':', 1)
                        indent = line[:len(line) - len(line.lstrip('\t '))]
                        new_lines.append(parts[0] + ':\n')
                        new_lines.append(indent + '\t' + parts[1].lstrip())
                        changed = True
                    else:
                        new_lines.append(line)
                else:
                    new_lines.append(line)
                    
            if changed:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.writelines(new_lines)

print("Fixed E701 issues.")
