import re

with open('BOA.pot', encoding='utf-8') as f:
    pot_content = f.read()

with open('addon/locale/zh_CN/LC_MESSAGES/nvda.po', encoding='utf-8') as f:
    po_content = f.read()

# Improved regex to handle multiline msgid/msgstr but we mostly have single lines for these ones.
# Actually let's just parse it simply:
def parse_po(text):
    entries = {}
    blocks = text.split('\n\n')
    for block in blocks:
        lines = block.split('\n')
        msgid = ""
        msgstr = ""
        in_msgid = False
        in_msgstr = False
        for line in lines:
            if line.startswith('msgid "'):
                msgid = line[7:-1]
                in_msgid = True
                in_msgstr = False
            elif line.startswith('msgstr "'):
                msgstr = line[8:-1]
                in_msgstr = True
                in_msgid = False
            elif line.startswith('"') and line.endswith('"'):
                if in_msgid:
                    msgid += line[1:-1]
                elif in_msgstr:
                    msgstr += line[1:-1]
        if msgid:
            entries[msgid] = msgstr
    return entries

pot_entries = parse_po(pot_content)
po_entries = parse_po(po_content)

for msgid in pot_entries:
    if msgid == "": continue
    if msgid not in po_entries or po_entries[msgid] == "":
        print(f"Missing: {msgid}")
