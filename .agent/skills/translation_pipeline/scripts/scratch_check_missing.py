import os
import re

def parse_po(filepath):
    if not os.path.exists(filepath):
        return []
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    msgids = []
    lines = content.split('\n')
    current_msgid = ""
    in_msgid = False
    
    for line in lines:
        if line.startswith('msgid '):
            in_msgid = True
            current_msgid = line[6:].strip().strip('"')
        elif line.startswith('msgstr '):
            if current_msgid and current_msgid != "":
                msgids.append(current_msgid)
            in_msgid = False
            current_msgid = ""
        elif in_msgid and line.startswith('"'):
            current_msgid += line.strip().strip('"')
            
    return msgids

def main():
    pot_file = "BOA.pot"
    pot_ids = parse_po(pot_file)
    print(f"Total msgids in master POT: {len(pot_ids)}")
    
    locales_dir = os.path.join("addon", "locale")
    if not os.path.exists(locales_dir):
        print("Locale dir not found.")
        return
        
    for lang in os.listdir(locales_dir):
        lang_dir = os.path.join(locales_dir, lang, "LC_MESSAGES")
        po_file = os.path.join(lang_dir, "nvda.po")
        if os.path.exists(po_file):
            po_ids = parse_po(po_file)
            missing = [m for m in pot_ids if m not in po_ids]
            if missing:
                print(f"[{lang}] Missing {len(missing)} strings.")
                for m in missing:
                    print(f"   - {m}")
        else:
            print(f"[{lang}] nvda.po does not exist!")

if __name__ == '__main__':
    main()
