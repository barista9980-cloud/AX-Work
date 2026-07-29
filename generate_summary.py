import json
from collections import Counter

data = json.load(open('real_estate_parsed_catalog.json', encoding='utf-8'))

regions = Counter(d.get('region', '기타') for d in data)
types = Counter(d.get('contract_type', '기타') for d in data)
lessors = Counter(d.get('party_a_lessor', '기타') for d in data)
lessees = Counter(d.get('party_b_lessee', '기타') for d in data)

summary = {
    "total_files": len(data),
    "by_region": dict(regions),
    "by_contract_type": dict(types),
    "top_lessors": dict(lessors.most_common(5)),
    "top_lessees": dict(lessees.most_common(5))
}

with open('summary_report.json', 'w', encoding='utf-8') as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print("Summary report generated successfully.")
