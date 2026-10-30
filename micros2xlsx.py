"""Lit des images de factures MICROS (OCR) -> Excel : Date, Point de vente, Client, Montant."""
import argparse
import re
from datetime import date
from pathlib import Path

import pytesseract
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageOps

EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
HEADERS = ["Date", "Point de vente", "Client", "Montant", "Fichier", "À vérifier"]

# --- Patterns à adapter selon le format de VOS tickets ---
DATE_RE = re.compile(r"\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\b")
POS_RE = re.compile(r"\b(?:point de vente|pdv|rvc|revenue center|outlet)\s*[:\-]?\s*(.+)", re.I)
CLIENT_RE = re.compile(r"\b(?:client|customer|guest)\s*[:\-]?\s*(.+)", re.I)
TOTAL_RE = re.compile(
    r"(?<!sous-)(?<!sous )(?<!sub-)(?<!sub )\b(?:total|montant|amount due)"
    r"(?:\s*(?:ttc|à payer|a payer|due))?\s*[:\-]?\s*(?:MAD|DH|€|\$)?\s*([\d\s.,]*\d)",
    re.I,
)


def to_float(s):
    """'1 234,50' / '1,234.50' / '220,00' -> float (None si illisible)."""
    s = re.sub(r"\s", "", s)
    try:
        if "," in s and "." in s:
            dec = "," if s.rfind(",") > s.rfind(".") else "."
            s = s.replace("." if dec == "," else ",", "").replace(dec, ".")
        elif "," in s:
            s = s.replace(",", "" if len(s.split(",")[-1]) == 3 else ".")
        return float(s)
    except ValueError:
        return None


def make_date(d, m, y):
    y = int(y)
    try:
        return date(y + 2000 if y < 100 else y, int(m), int(d))
    except ValueError:
        return None


def ocr(path):
    img = ImageOps.autocontrast(ImageOps.grayscale(ImageOps.exif_transpose(Image.open(path))))
    if img.width < 1000:  # agrandir les petits tickets améliore l'OCR
        img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    return pytesseract.image_to_string(img, lang="fra+eng", config="--psm 6")


def parse(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    r = {"date": None, "pos": None, "client": None, "total": None}
    totals = []
    for l in lines:
        if r["date"] is None and (m := DATE_RE.search(l)):
            r["date"] = make_date(*m.groups())
        if r["pos"] is None and (m := POS_RE.search(l)):
            r["pos"] = m.group(1).strip()
        if r["client"] is None and (m := CLIENT_RE.search(l)):
            r["client"] = m.group(1).strip()
        if (m := TOTAL_RE.search(l)) and (v := to_float(m.group(1))) is not None:
            totals.append(v)
    r["total"] = max(totals) if totals else None  # le total général est le plus grand montant
    if r["pos"] is None and lines:
        r["pos"] = lines[0]  # repli : en-tête du ticket (nom du restaurant)
    return r


def open_book(path):
    if Path(path).exists():
        wb = load_workbook(path)
        return wb, wb.active
    wb = Workbook()
    ws = wb.active
    ws.title = "Factures"
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, w in zip("ABCDEF", (12, 24, 28, 12, 30, 14)):
        ws.column_dimensions[col].width = w
    return wb, ws


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source", help="image ou dossier d'images")
    ap.add_argument("-o", "--output", default="factures.xlsx")
    ap.add_argument("--debug", action="store_true", help="affiche le texte OCR brut")
    a = ap.parse_args()

    src = Path(a.source)
    files = sorted(f for f in src.iterdir() if f.suffix.lower() in EXTS) if src.is_dir() else [src]
    wb, ws = open_book(a.output)
    done = {row[4] for row in ws.iter_rows(min_row=2, values_only=True)}
    yellow = PatternFill("solid", fgColor="FFF2CC")

    for f in files:
        if f.name in done:
            print(f"= {f.name} déjà traité")
            continue
        text = ocr(f)
        if a.debug:
            print(f"--- {f.name} ---\n{text}")
        r = parse(text)
        missing = [k for k in ("date", "pos", "client", "total") if not r[k]]
        ws.append([r["date"], r["pos"], r["client"], r["total"], f.name,
                   ", ".join(missing) if missing else ""])
        row = ws[ws.max_row]
        row[0].number_format = "DD/MM/YYYY"
        row[3].number_format = "#,##0.00"
        if missing:
            for c in row:
                c.fill = yellow
        print(f"{'!' if missing else '+'} {f.name} -> {r['date']} | {r['pos']} | {r['client']} | {r['total']}")

    wb.save(a.output)
    print(f"Enregistré : {a.output}")


if __name__ == "__main__":
    main()
