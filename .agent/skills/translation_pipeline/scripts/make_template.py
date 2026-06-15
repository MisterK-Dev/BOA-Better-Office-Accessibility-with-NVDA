import json
from scratch_check_missing import parse_po

pot_ids = parse_po("BOA.pot")
es_ids = parse_po("addon/locale/es/LC_MESSAGES/nvda.po")
missing = [m for m in pot_ids if m not in es_ids]

template = {
    "es": {m: "" for m in missing},
    "fr": {m: "" for m in missing},
    "de": {m: "" for m in missing},
    "pt": {m: "" for m in missing},
    "zh_CN": {m: "" for m in missing}
}

with open("scratch/batch1_template.json", "w", encoding="utf-8") as f:
    json.dump(template, f, indent=2, ensure_ascii=False)
    
print("Generated scratch/batch1_template.json")
