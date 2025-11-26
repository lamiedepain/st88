import openpyxl
from datetime import datetime

EXCEL_FILE = '2026 - PRESENCES_CONGES VOIRIE ESPACES VERTS ST8 (1).xlsx'

wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
sheet = wb['config']

# headers row 2
headers = [cell.value for cell in sheet[2]]

agents = []
for row_idx in range(3, 76):
    row = [cell.value for cell in sheet[row_idx]]
    if row[5]:
        nom = row[5].strip() if isinstance(row[5], str) else row[5]
        prenom = row[6].strip() if isinstance(row[6], str) else row[6]
        agents.append({'row': row_idx, 'nom': nom, 'prenom': prenom, 'data': row})


def build_aptitudes(agent):
    aptitudes = []
    for idx, val in enumerate(agent['data']):
        if idx >= 7 and val is not None and str(val).strip() != '':
            display = ''
            dateObj = None
            try:
                parsed = None
                if isinstance(val, (datetime,)):
                    parsed = val
                else:
                    parsed = datetime.fromisoformat(str(val)) if 'T' in str(val) else None
                if parsed:
                    dateObj = parsed
                    display = parsed.strftime('%d/%m/%Y')
                else:
                    # try parsing generic
                    try:
                        parsed2 = datetime.strptime(str(val), '%Y-%m-%d')
                        dateObj = parsed2
                        display = parsed2.strftime('%d/%m/%Y')
                    except Exception:
                        if isinstance(val, str) and val.strip().lower() == 'x':
                            display = '✓'
                        else:
                            display = str(val).strip()
            except Exception:
                display = str(val).strip()
            aptitudes.append({'nom': headers[idx] if idx < len(headers) else f'col{idx}', 'valeur': display, 'dateObj': dateObj})
    return aptitudes


def map_badges(aptitudes):
    badges = []
    for apt in aptitudes:
        nomLower = (apt['nom'] or '').lower()
        val = apt['valeur'] if apt['valeur'] else ''
        badge = None
        if 'r.482' in nomLower or 'r482' in nomLower or 'caces' in nomLower:
            badge = ('C', 'badge-caces')
        elif 'grue' in nomLower or 'nacelle' in nomLower or 'r.490' in nomLower:
            badge = ('C', 'badge-caces')
        elif 'chariot' in nomLower or 'r.489' in nomLower:
            badge = ('C', 'badge-caces')
        elif 'tondeuse' in nomLower:
            badge = ('Tn', 'badge-tondeuse')
        elif 'tronço' in nomLower or 'tronco' in nomLower:
            badge = ('T', 'badge-tronco')
        elif 'permis' in nomLower or 'remorque' in nomLower or 'perm' in nomLower:
            badge = ('P', 'badge-permis')
        elif 'aipr' in nomLower:
            badge = ('A', 'badge-aipr')
        elif 'fimo' in nomLower:
            badge = ('F', 'badge-fimo')
        elif 'secours' in nomLower or 'premiers secours' in nomLower:
            badge = ('S', 'badge-secours')
        else:
            if val and val != '✓':
                badge = ( (str(nomLower).split()[0][:3] or 'GEN').upper(), 'badge-generic')
        if badge:
            title = f"{apt['nom']}: {val}" if val and val != '✓' else apt['nom']
            badges.append({'text': badge[0], 'class': badge[1], 'title': title})
    return badges


for agent in agents:
    if 'MALLET' in str(agent['nom']).upper() or 'TADJROUNA' in str(agent['nom']).upper():
        print('---', agent['nom'], agent['prenom'], f'(row {agent["row"]})')
        apts = build_aptitudes(agent)
        if not apts:
            print('  (aucune aptitude non vide)')
        else:
            for apt in apts:
                print(f"  - {apt['nom']} -> {apt['valeur']}")
            badges = map_badges(apts)
            print('  Badges générés:')
            for b in badges:
                print(f"    [{b['text']}] class={b['class']} title='{b['title']}'")
        print()
