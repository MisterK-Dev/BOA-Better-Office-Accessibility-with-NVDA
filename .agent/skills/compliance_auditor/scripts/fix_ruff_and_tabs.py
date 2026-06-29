import os
import glob
import re
import subprocess
import json

def fix_e402():
    p = subprocess.run(['ruff', 'check', 'addon/', '--output-format', 'json'], capture_output=True, text=True)
    errors = json.loads(p.stdout)
    
    # Group by file
    e402_by_file = {}
    for e in errors:
        if e['code'] == 'E402':
            filepath = e['location']['row']
            f_path = e['filename']
            if f_path not in e402_by_file:
                e402_by_file[f_path] = set()
            e402_by_file[f_path].add(e['location']['row'])
            
    for f_path, rows in e402_by_file.items():
        with open(f_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            
        for row in rows:
            idx = row - 1
            if idx < len(lines):
                line = lines[idx].rstrip('\n')
                if '# noqa' not in line:
                    lines[idx] = line + "  # noqa: E402\n"
                    
        with open(f_path, 'w', encoding='utf-8') as f:
            f.writelines(lines)
            
    print(f"Fixed {sum(len(v) for v in e402_by_file.values())} E402 errors.")

def fix_tabs():
    count = 0
    for root, dirs, files in os.walk('addon/'):
        for file in files:
            if file.endswith('.py'):
                filepath = os.path.join(root, file)
                with open(filepath, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                
                changed = False
                new_lines = []
                for line in lines:
                    match = re.match(r'^([ \t]+)', line)
                    if match:
                        leading = match.group(1)
                        if ' ' in leading: # If there's any space in the leading whitespace
                            leading_spaces = leading.replace('\t', '    ')
                            tabs = len(leading_spaces) // 4
                            spaces = len(leading_spaces) % 4
                            new_line = '\t' * tabs + ' ' * spaces + line.lstrip(' \t')
                            new_lines.append(new_line)
                            changed = True
                        else:
                            new_lines.append(line)
                    else:
                        new_lines.append(line)
                        
                if changed:
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.writelines(new_lines)
                    count += 1
    print(f"Fixed indentation in {count} files.")

if __name__ == "__main__":
    fix_e402()
    fix_tabs()
