# Controle qualite PCB

Application web locale et deployable pour analyser des certificats qualite PCB au format PDF.

## Utilisation locale

Double-cliquer sur `Controle_PCB.bat`, puis deposer un ou plusieurs PDF dans la page web ouverte automatiquement.

## Deploiement Render

Le projet est pret pour Render avec:

- `requirements.txt`
- `Procfile`
- `render.yaml`

Commande de demarrage:

```bash
gunicorn web_app:app
```
