# ST88 Planning Management

Application web Flask pour gérer les agents et plannings.

## Déploiement sur Render

1. **Préparer le repository GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/VOTRE_USERNAME/st88-planning.git
   git push -u origin main
   ```

2. **Ajouter le fichier Excel au repository**
   - Copier `2026 - PRESENCES_CONGES VOIRIE ESPACES VERTS ST8 (1).xlsx` à la racine
   - `git add "2026 - PRESENCES_CONGES VOIRIE ESPACES VERTS ST8 (1).xlsx"`
   - `git commit -m "Add Excel file"`
   - `git push`

3. **Déployer sur Render**
   - Aller sur https://render.com
   - Connecter votre compte GitHub
   - Cliquer "New +" → "Web Service"
   - Sélectionner votre repository `st88-planning`
   - Render détectera automatiquement le `render.yaml`
   - Cliquer "Create Web Service"

4. **Accéder à l'application**
   - URL: https://st88-planning.onrender.com (ou l'URL donnée par Render)
   - Les modifications seront enregistrées dans le fichier Excel sur le serveur
   - Une sauvegarde est créée à chaque démarrage dans le dossier `backups/`

## Structure

```
st88/
├── app.py                 # Application Flask
├── requirements.txt       # Dépendances Python
├── render.yaml           # Configuration Render
├── templates/
│   ├── agents.html       # Gestion des agents
│   ├── planning.html     # Vue planning
│   └── generator.html    # Générateur de planning
└── backups/              # Sauvegardes automatiques
```

## Fonctionnalités

- ✅ Gestion des agents (ajouter, modifier, supprimer)
- ✅ Affichage par groupes avec codes couleurs
- ✅ Vue planning par mois
- ✅ Sauvegarde automatique au démarrage
- 🚧 Générateur automatique de planning (en développement)

## Notes

- Le fichier Excel original est modifié directement
- Les sauvegardes sont créées dans `backups/` avec horodatage
- L'application est accessible depuis n'importe quel navigateur
- Pas besoin de Python en local, tout tourne sur Render
