import sys
import os
import re

def check_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    violations = []
    
    for i, line in enumerate(lines):
        # We are looking for _(" or _(' 
        # Note: sometimes they do _( " or other spacings, we can use regex
        if re.search(r'\b_\(\s*[\'"]', line):
            # Backtrack to find the previous non-empty line
            j = i - 1
            has_comment = False
            while j >= 0:
                prev_line = lines[j].strip()
                if prev_line == '':
                    j -= 1
                    continue
                if prev_line.startswith('# Translators:'):
                    has_comment = True
                    break
                else:
                    break
                
            if not has_comment:
                violations.append((i + 1, line.strip()))
                
    return violations

def main():
    target_dir = sys.argv[1] if len(sys.argv) > 1 else 'addon'
    all_violations = {}
    
    for root, dirs, files in os.walk(target_dir):
        for file in files:
            if file.endswith('.py'):
                filepath = os.path.join(root, file)
                violations = check_file(filepath)
                if violations:
                    all_violations[filepath] = violations

    if all_violations:
        print("ERROR: Found missing '# Translators:' comments above localized strings.\n")
        for filepath, violations in all_violations.items():
            print(f"File: {filepath}")
            for line_no, content in violations:
                print(f"  Line {line_no}: {content}")
        sys.exit(1)
    else:
        print("SUCCESS: All translatable strings have translator comments.")
        sys.exit(0)

if __name__ == '__main__':
    main()
