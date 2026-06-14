import glob
import os

def fix_po_files():
    # Go up 3 levels from .agent/skills/build_fixer to the repository root
    # or just assume the script is executed with root as CWD.
    # To be safe, we calculate root dynamically just in case.
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.abspath(os.path.join(script_dir, '..', '..', '..'))
    
    search_pattern = os.path.join(root_dir, 'addon', 'locale', '*', 'LC_MESSAGES', 'nvda.po')
    
    for f in glob.glob(search_pattern):
        try:
            with open(f, 'r', encoding='utf-8') as file:
                lines = file.read().split('\n')
            
            changed = False
            for i in range(len(lines)):
                if lines[i].startswith('msgstr'):
                    j = i - 1
                    msgid_has_newline = False
                    while j >= 0:
                        if lines[j].startswith('msgid'):
                            break
                        if lines[j].endswith('\\n"'):
                            msgid_has_newline = True
                            break
                        elif lines[j].endswith('"') and not lines[j].endswith('\\n"'):
                            msgid_has_newline = False
                            break
                        j -= 1
                    
                    k = i
                    while k < len(lines) and (lines[k].startswith('msgstr') or lines[k].startswith('"')):
                        k += 1
                    
                    if k - 1 >= i:
                        last_line = lines[k-1]
                        msgstr_has_newline = last_line.endswith('\\n"')
                        
                        if not msgid_has_newline and msgstr_has_newline:
                            lines[k-1] = last_line[:-3] + '"'
                            changed = True
                        elif msgid_has_newline and not msgstr_has_newline and last_line.endswith('"'):
                            lines[k-1] = last_line[:-1] + '\\n"'
                            changed = True

            if changed:
                with open(f, 'w', encoding='utf-8') as file:
                    file.write('\n'.join(lines))
                print(f"Fixed newlines in {f}")
        except Exception as e:
            print(f"Error processing {f}: {e}")

if __name__ == '__main__':
    fix_po_files()
