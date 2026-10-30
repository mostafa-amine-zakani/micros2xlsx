from micros2xlsx import parse, to_float

SAMPLE = """LE JARDIN - Casablanca
Point de vente: Terrasse
Table 12 Chk 4587
Date: 21/09/2026 20:45
Client: M. Alaoui
2 x Tajine 180,00
Sous-total 200,00
TVA 10% 20,00
TOTAL 220,00
"""


def test_parse():
    r = parse(SAMPLE)
    assert r["date"].isoformat() == "2026-09-21"
    assert r["pos"] == "Terrasse"
    assert r["client"] == "M. Alaoui"
    assert r["total"] == 220.0


def test_to_float():
    assert to_float("1 234,50") == 1234.5
    assert to_float("1,234.50") == 1234.5
    assert to_float("220,00") == 220.0
