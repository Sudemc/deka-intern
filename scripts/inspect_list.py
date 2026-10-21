import json

with open('parsed_components.json', 'r', encoding='utf-8') as f:
    items = json.load(f)

for cat in ['ADP', 'ANT', 'BAH', 'CAP', 'DIO', 'BUZ']:
    print(f'=== {cat} ===')
    for it in items:
        if it['cat'] == cat:
            print(f"ERP: {it['erp']:25} | {it['rest']}")
