import requests
import json

response = requests.get('https://st88-8ew0.onrender.com/api/planning-data/2026/01')
data = response.json()

print(f"Agents retournés: {len(data.get('agents', []))}")
agents = data.get('agents', [])

print(f"\nPremiers 5 agents:")
for a in agents[:5]:
    print(f"  - {a['nom']} {a['prenom']}")

print(f"\nDerniers 5 agents:")
for a in agents[-5:]:
    print(f"  - {a['nom']} {a['prenom']}")

print(f"\nTous les noms:")
for i, a in enumerate(agents, 1):
    print(f"{i}. {a['nom']}")
