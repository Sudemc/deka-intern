import json

with open('parsed_components.json', 'r', encoding='utf-8') as f:
    items = json.load(f)

for cat in ['DSP', 'CON', 'CRY']:
    print(f'=== {cat} ({len([it for it in items if it["cat"] == cat])} items) ===')
    for it in items:
        if it['cat'] == cat:
            print(f"ERP: {it['erp']:25} | {it['rest']}")
