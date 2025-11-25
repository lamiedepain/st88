from flask import Flask, render_template, jsonify, request
import openpyxl
from datetime import datetime
import os
import shutil

app = Flask(__name__)

EXCEL_FILE = '2026 - PRESENCES_CONGES VOIRIE ESPACES VERTS ST8 (1).xlsx'
BACKUP_DIR = 'backups'

# Créer une sauvegarde au démarrage
if not os.path.exists(BACKUP_DIR):
    os.makedirs(BACKUP_DIR)

backup_file = os.path.join(BACKUP_DIR, f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
if os.path.exists(EXCEL_FILE):
    shutil.copy2(EXCEL_FILE, backup_file)
    print(f"✅ Sauvegarde créée: {backup_file}")

@app.route('/')
def index():
    return render_template('agents.html')

@app.route('/planning')
def planning():
    return render_template('planning.html')

@app.route('/generator')
def generator():
    return render_template('generator.html')

@app.route('/api/agents', methods=['GET'])
def get_agents():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        sheet = wb['config']
        
        # Lire les en-têtes (ligne 2)
        headers = [cell.value for cell in sheet[2]]
        
        # Lire les agents (lignes 3 à 75)
        agents = []
        for row_idx in range(3, 76):  # Lignes 3 à 75
            row = [cell.value for cell in sheet[row_idx]]
            if row[5]:  # Si le nom existe (colonne 6 = index 5)
                agent = {
                    'index': row_idx,  # Numéro de ligne réel
                    'matricule': row[4] or '',
                    'nom': row[5] or '',
                    'prenom': row[6] or '',
                    'data': row
                }
                agents.append(agent)
        
        return jsonify({
            'success': True,
            'headers': headers,
            'agents': agents
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/agents/<int:index>', methods=['PUT'])
def update_agent(index):
    try:
        data = request.json
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400
            
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb['config']
        
        # index est déjà le numéro de ligne réel
        row_num = index
        
        # Mettre à jour les cellules - s'assurer que toutes les valeurs sont sérialisables
        for col_idx, value in enumerate(data.get('data', []), start=1):
            # Convertir les valeurs None en chaîne vide
            if value is None:
                value = ''
            sheet.cell(row=row_num, column=col_idx, value=value)
        
        wb.save(EXCEL_FILE)
        return jsonify({'success': True})
    except Exception as e:
        print(f"Erreur lors de la mise à jour de l'agent {index}: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/agents/<int:index>', methods=['DELETE'])
def delete_agent(index):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb['config']
        
        # index est déjà le numéro de ligne réel
        row_num = index
        sheet.delete_rows(row_num, 1)
        
        wb.save(EXCEL_FILE)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/planning/<month>', methods=['GET'])
def get_planning(month):
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        
        # Le mois est au format "Janvier 2026"
        if month not in wb.sheetnames:
            return jsonify({'success': False, 'error': 'Mois introuvable'}), 404
        
        sheet = wb[month]
        
        # Lire toutes les données du planning
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(list(row))
        
        return jsonify({
            'success': True,
            'month': month,
            'data': data
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/planning/<month>', methods=['PUT'])
def update_planning(month):
    try:
        data = request.json
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400
            
        wb = openpyxl.load_workbook(EXCEL_FILE)
        
        if month not in wb.sheetnames:
            return jsonify({'success': False, 'error': 'Mois introuvable'}), 404
        
        sheet = wb[month]
        
        # Mettre à jour les cellules modifiées
        for update in data.get('updates', []):
            row = update['row']
            col = update['col']
            value = update['value']
            sheet.cell(row=row, column=col, value=value)
        
        wb.save(EXCEL_FILE)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/months', methods=['GET'])
def get_months():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        months = [name for name in wb.sheetnames if name != 'config']
        return jsonify({
            'success': True,
            'months': months
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
