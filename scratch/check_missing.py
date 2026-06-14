import re

def get_msgids(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    # Find all msgid blocks. A block is msgid followed by msgstr
    matches = re.findall(r'msgid "(.*?)"\nmsgstr "(.*?)"', content, re.DOTALL)
    # Also handle multi-line if any, but in this case they are simple strings or have explicit newlines.
    # Actually gettext uses msgid "" \n "string" for multiline, let's just use a simpler parser or polib if available.
    
    # A simple parser
    blocks = []
    current_msgid = None
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    msgids = set()
    for i, line in enumerate(lines):
        if line.startswith('msgid '):
            # Extract the full msgid
            msgid = line[6:].strip('\n"')
            j = i + 1
            while j < len(lines) and lines[j].startswith('"'):
                msgid += lines[j].strip('\n"')
                j += 1
            if msgid:
                msgids.add(msgid)
    return msgids

pot_msgids = get_msgids('BOA.pot')
po_msgids = get_msgids('addon/locale/ta/LC_MESSAGES/nvda.po')

missing = pot_msgids - po_msgids
print("Missing in PO:")
for m in missing:
    print(m)

