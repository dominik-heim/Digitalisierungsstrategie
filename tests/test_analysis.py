from digistrat import analysis as A


def test_stakeholder_quadrants():
    assert A.stakeholder_quadrant(5, 5) == "Eng einbinden"
    assert A.stakeholder_quadrant(5, 1) == "Zufriedenstellen"
    assert A.stakeholder_quadrant(1, 5) == "Informieren"
    assert A.stakeholder_quadrant(2, 2) == "Beobachten"


def test_impact_effort_quadrant():
    assert A.quadrant(8, 3) == "Quick Win"
    assert A.quadrant(8, 8) == "Großes Vorhaben"
    assert A.quadrant(3, 3) == "Lückenfüller"
    assert A.quadrant(3, 8) == "Nicht machen"


def test_finance():
    flows = [-100.0, 60.0, 60.0]
    assert abs(A.npv(0.0, flows) - 20) < 1e-9
    assert abs(A.payback(flows) - (1 + 40 / 60)) < 1e-9
    r = A.irr(flows)
    assert r is not None and abs(A.npv(r, flows)) < 1e-6


def test_invest_check():
    good = {"rolle": "Werker", "ziel": "Arbeitsgänge am Terminal abmelden", "nutzen": "der Status in Echtzeit stimmt",
            "akzeptanz": "Given ein Auftrag When ich abmelde Then ist er im ERP", "story_points": 3}
    assert A.invest_check(good)["score"] == 1
    bad = dict(good, ziel="Daten erfassen und auswerten", akzeptanz="geht", story_points=13)
    res = A.invest_check(bad)
    assert not res["checks"]["I"] and not res["checks"]["S"] and not res["checks"]["T"]


def test_quarters():
    assert A.q_add("2027-Q4", 1) == "2028-Q1"
    assert A.q_diff("2027-Q1", "2028-Q1") == 4


def test_vsm_kpis():
    rows = [{"prozess": "P", "nr": 1, "schritt": "A", "bz": 10, "wz": 0, "rz": 0, "qr": 90, "wertschoepfend": "Ja",
             "medium": "Papier", "system": None},
            {"prozess": "P", "nr": 2, "schritt": "B", "bz": 10, "wz": 80, "rz": 0, "qr": 100, "wertschoepfend": "Nein",
             "medium": "Digital integriert", "system": None}]
    v = A.vsm(rows)["prozesse"]["P"]
    assert v["durchlaufzeit"] == 100 and abs(v["flussgrad"] - 0.1) < 1e-9 and abs(v["rty"] - 0.9) < 1e-9
    assert len(v["medienbrueche"]) == 1 and v["engpass"] == "B"


def test_raci_validation():
    r = A.raci(["GF", "PL"], [{"_id": "X", "GF": "A", "PL": "R"}, {"_id": "Y", "GF": "A", "PL": "A"}])
    assert [i["aufgabe"] for i in r["fehler"]] == ["Y"]
