import polib

pot = polib.pofile('c:\\Users\\kiran\\.gemini\\antigravity\\scratch\\BOA\\BOA.pot')
po = polib.pofile('c:\\Users\\kiran\\.gemini\\antigravity\\scratch\\BOA\\addon\\locale\\it\\LC_MESSAGES\\nvda.po')

po_dict = {entry.msgid: entry for entry in po}

missing_or_empty = []
for entry in pot:
    if entry.msgid not in po_dict or not po_dict[entry.msgid].msgstr.strip():
        missing_or_empty.append(entry.msgid)

print("Missing or empty msgids:")
for m in missing_or_empty:
    print(repr(m))
