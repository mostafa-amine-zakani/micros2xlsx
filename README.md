# micros2xlsx

Lit des **images de factures MICROS** (Oracle MICROS) par OCR local et remplit un fichier **Excel** avec :

| Date | Point de vente | Client | Montant | Fichier | À vérifier |
|------|----------------|--------|---------|---------|------------|

- 100 % local et gratuit (Tesseract), aucune API
- Ajoute à un Excel existant, ignore les images déjà traitées
- Lignes incomplètes surlignées en jaune + colonne « À vérifier »

## Installation

1. Installer Tesseract + le français
   - Ubuntu/Debian : `sudo apt install tesseract-ocr tesseract-ocr-fra`
   - macOS : `brew install tesseract tesseract-lang`
   - Windows : https://github.com/UB-Mannheim/tesseract/wiki (cocher « French »)
2. `pip install -r requirements.txt`

## Utilisation

```bash
python micros2xlsx.py factures/                 # tout un dossier
python micros2xlsx.py ticket.jpg -o mois.xlsx   # une image
python micros2xlsx.py factures/ --debug         # affiche le texte OCR brut
```

## Adapter à vos tickets

Le format MICROS varie d'un établissement à l'autre. Lancez `--debug` sur un ticket, puis ajustez les
expressions régulières en haut de `micros2xlsx.py` (`DATE_RE`, `POS_RE`, `CLIENT_RE`, `TOTAL_RE`).

- Montant : le plus grand des montants trouvés sur les lignes « Total / Montant » (hors sous-total)
- Point de vente : ligne « Point de vente / PDV / RVC », sinon 1re ligne du ticket

## Conseils OCR

Photo bien à plat, éclairée, ticket net (idéalement ≥ 1000 px de large).

## Tests

```bash
pytest
```

## Licence

MIT
