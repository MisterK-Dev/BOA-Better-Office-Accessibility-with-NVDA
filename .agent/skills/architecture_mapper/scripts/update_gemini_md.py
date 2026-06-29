import os
import re

def generate_tree(dir_path, prefix="", depth=0, max_depth=3):
    if depth > max_depth:
        return ""
    
    ignore_dirs = {'.git', '__pycache__', 'site_scons'}
    ignore_exts = {'.pyc', '.mo', '.po', '.pot'}
    
    try:
        items = os.listdir(dir_path)
    except PermissionError:
        return ""
    
    items.sort()
    dirs = []
    files = []
    for item in items:
        path = os.path.join(dir_path, item)
        if os.path.isdir(path):
            if item not in ignore_dirs:
                dirs.append(item)
        else:
            ext = os.path.splitext(item)[1]
            if ext not in ignore_exts and not item.endswith('.nvda-addon'):
                files.append(item)
                
    result = ""
    total = len(dirs) + len(files)
    count = 0
    
    for item in dirs:
        count += 1
        connector = "└── " if count == total else "├── "
        result += f"{prefix}{connector}{item}/\n"
        extension = "    " if count == total else "│   "
        result += generate_tree(os.path.join(dir_path, item), prefix + extension, depth + 1, max_depth)
        
    for item in files:
        count += 1
        connector = "└── " if count == total else "├── "
        result += f"{prefix}{connector}{item}\n"
        
    return result

def main():
    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..', '..'))
    gemini_file = os.path.join(repo_root, 'GEMINI.md')
    
    if not os.path.exists(gemini_file):
        print("Error: GEMINI.md not found in the root directory.")
        return

    print("Scanning directory structure...")
    tree_text = "```\n" + generate_tree(repo_root, max_depth=3) + "```\n"
    
    with open(gemini_file, 'r', encoding='utf-8') as f:
        content = f.read()
        
    pattern = re.compile(r'(<!-- MAP_START -->\n).*?(\n<!-- MAP_END -->)', re.DOTALL)
    if not pattern.search(content):
        print("Error: Could not find <!-- MAP_START --> and <!-- MAP_END --> markers in GEMINI.md")
        return
        
    new_content = pattern.sub(rf'\g<1>{tree_text}\g<2>', content)
    
    with open(gemini_file, 'w', encoding='utf-8') as f:
        f.write(new_content)
        
    print("Successfully updated GEMINI.md architecture map!")

if __name__ == '__main__':
    main()
