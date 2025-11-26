from app import app
import json

with app.test_client() as client:
    response = client.get('/api/planning-data/2026/01')
    data = json.loads(response.data)
    
    agents = data.get('agents', [])
    print(f"Agents retournés par l'API locale: {len(agents)}\n")
    
    for i, a in enumerate(agents[:10], 1):
        print(f"{i}. {a['nom']} {a['prenom']}")
    
    print("\n...")
    
    for i, a in enumerate(agents[-10:], len(agents)-9):
        print(f"{i}. {a['nom']} {a['prenom']}")
