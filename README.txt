# Prototype unifié — Q1 / Q2 + galerie des types

## Local
python -m pip install -r requirements.txt
python app.py

## URLs
- /          Accueil + choix Q1/Q2 + paramètres
- /q1        Lance directement Q1
- /q2        Lance directement Q2
- /types     Galerie: 1 exemple de chaque type (pour vérifier ce qui est trop facile / trop dur)

## Render
Build: pip install -r requirements.txt
Start: gunicorn app:app

## Retirer un type (ex: les nombres)
Sur l'accueil, section Paramètres > “Types inclus”, décoche le type.

V3: Export Excel plus lisible (couleurs, feedbackPhase, sheet blocks + dictionary).

V4: corrections (courbe monte/descend), options modifiables après clic, ligne centrale retirée pour couleur majoritaire.
