import json

with open('parsed_components.json', 'r', encoding='utf-8') as f:
    items = json.load(f)

print(f'=== DSP ({len([it for it in items if it["cat"] == "DSP"])} items) ===')
for it in items:
    if it['cat'] == 'DSP':
        print(f"ERP: {it['erp']:25} | {it['rest']}")
