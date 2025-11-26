from app import app

if __name__ == '__main__':
    # Démarrage local sur le port 5001 pour éviter conflits
    app.run(host='0.0.0.0', port=5001, debug=True)
