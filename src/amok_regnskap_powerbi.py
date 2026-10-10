"""
Henter alle tilgjengelige strukturerte årsregnskap for AMOK EQUIPMENT AS
fra Brønnøysundregistrenes Regnskapsregister-API.

Lager to filer:
1. amok_equipment_alle_aar_wide.csv
2. amok_equipment_alle_aar_long.csv

Installer avhengigheter:
    pip install -r requirements.txt

Kjør:
    python amok_regnskap_powerbi.py
"""

from pathlib import Path
import time

import pandas as pd
import requests


ORGNR = "911823705"
SELSKAP = "AMOK EQUIPMENT AS"
REGNSKAPSTYPE = "SELSKAP"

BASE_DIR = Path(__file__).resolve().parent.parent
PROCESSED_DIR = BASE_DIR / "data" / "processed"
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

BASE_URL = "https://data.brreg.no/regnskapsregisteret/regnskap"
WIDE_OUTPUT_FILE = PROCESSED_DIR / "amok_equipment_alle_aar_wide.csv"
LONG_OUTPUT_FILE = PROCESSED_DIR / "amok_equipment_alle_aar_long.csv"


REQUEST_TIMEOUT = 30
PAUSE_BETWEEN_REQUESTS = 0.25


def nested_get(data, *keys, default=None):
    """Henter en verdi trygt fra en nestet dictionary."""
    current = data
    for key in keys:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current is None:
            return default
    return current


def safe_divide(numerator, denominator):
    """Utfører divisjon uten å feile ved null eller manglende verdier."""
    if numerator is None or denominator is None or denominator == 0:
        return None
    return numerator / denominator


def hent_tilgjengelige_aar(session, orgnr):
    """Henter år hvor Brønnøysund har en kopi av årsregnskapet."""
    url = f"{BASE_URL}/aarsregnskap/kopi/{orgnr}/aar"
    print(f"Henter tilgjengelige regnskapsår for {orgnr} ...")

    response = session.get(url, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    data = response.json()

    if not isinstance(data, list):
        raise ValueError("Forventet en liste med regnskapsår fra API-et.")

    aar = []
    for value in data:
        try:
            aar.append(int(value))
        except (TypeError, ValueError):
            print(f"Ignorerer ugyldig år: {value!r}")

    aar = sorted(set(aar))
    print(f"Tilgjengelige år: {aar}")
    return aar


def velg_regnskap_for_aar(data, aar):
    """Velger riktig regnskapsobjekt dersom API-et returnerer en liste."""
    if isinstance(data, dict):
        return data

    if not isinstance(data, list) or not data:
        return None

    for item in data:
        periode_til = nested_get(item, "regnskapsperiode", "tilDato")
        if periode_til and str(periode_til).startswith(str(aar)):
            return item

    return data[0]


def hent_regnskap_for_aar(session, orgnr, aar):
    """Henter strukturert selskapsregnskap for ett bestemt år."""
    url = f"{BASE_URL}/{orgnr}"
    params = {
        "år": aar,
        "regnskapstype": REGNSKAPSTYPE,
    }
    headers = {"Accept": "application/json"}

    print(f"Henter strukturert regnskap for {aar} ...")
    response = session.get(
        url,
        params=params,
        headers=headers,
        timeout=REQUEST_TIMEOUT,
    )

    if response.status_code == 404:
        print(f"  Ingen strukturerte data funnet for {aar}.")
        return None

    response.raise_for_status()
    return velg_regnskap_for_aar(response.json(), aar)


def flatten_regnskap(regnskap, aar):
    """Flater ut ett regnskap og beregner finansielle KPI-er."""
    periode_fra = nested_get(regnskap, "regnskapsperiode", "fraDato")
    periode_til = nested_get(regnskap, "regnskapsperiode", "tilDato")

    salgsinntekter = nested_get(
        regnskap, "resultatregnskapResultat", "driftsresultat",
        "driftsinntekter", "salgsinntekter"
    )
    driftsinntekter = nested_get(
        regnskap, "resultatregnskapResultat", "driftsresultat",
        "driftsinntekter", "sumDriftsinntekter"
    )
    lonnskostnad = nested_get(
        regnskap, "resultatregnskapResultat", "driftsresultat",
        "driftskostnad", "loennskostnad"
    )
    driftskostnader = nested_get(
        regnskap, "resultatregnskapResultat", "driftsresultat",
        "driftskostnad", "sumDriftskostnad"
    )
    driftsresultat = nested_get(
        regnskap, "resultatregnskapResultat", "driftsresultat", "driftsresultat"
    )
    netto_finans = nested_get(
        regnskap, "resultatregnskapResultat", "finansresultat", "nettoFinans"
    )
    resultat_for_skatt = nested_get(
        regnskap, "resultatregnskapResultat", "ordinaertResultatFoerSkattekostnad"
    )
    skattekostnad = nested_get(
        regnskap, "resultatregnskapResultat", "ordinaertResultatSkattekostnad"
    )
    arsresultat = nested_get(
        regnskap, "resultatregnskapResultat", "aarsresultat"
    )

    anleggsmidler = nested_get(
        regnskap, "eiendeler", "anleggsmidler", "sumAnleggsmidler"
    )
    omlopsmidler = nested_get(
        regnskap, "eiendeler", "omloepsmidler", "sumOmloepsmidler"
    )
    varer = nested_get(regnskap, "eiendeler", "sumVarer")
    fordringer = nested_get(regnskap, "eiendeler", "sumFordringer")
    bank = nested_get(regnskap, "eiendeler", "sumBankinnskuddOgKontanter")
    eiendeler = nested_get(regnskap, "eiendeler", "sumEiendeler")

    egenkapital = nested_get(
        regnskap, "egenkapitalGjeld", "egenkapital", "sumEgenkapital"
    )
    kortsiktig_gjeld = nested_get(
        regnskap, "egenkapitalGjeld", "gjeldOversikt",
        "kortsiktigGjeld", "sumKortsiktigGjeld"
    )
    langsiktig_gjeld = nested_get(
        regnskap, "egenkapitalGjeld", "gjeldOversikt",
        "langsiktigGjeld", "sumLangsiktigGjeld"
    )
    gjeld = nested_get(
        regnskap, "egenkapitalGjeld", "gjeldOversikt", "sumGjeld"
    )

    driftsmargin = safe_divide(driftsresultat, driftsinntekter)
    resultatmargin = safe_divide(arsresultat, driftsinntekter)
    egenkapitalandel = safe_divide(egenkapital, eiendeler)
    gjeldsgrad = safe_divide(gjeld, egenkapital)
    likviditetsgrad_1 = safe_divide(omlopsmidler, kortsiktig_gjeld)

    likvide_omlopsmidler = None
    if omlopsmidler is not None and varer is not None:
        likvide_omlopsmidler = omlopsmidler - varer

    likviditetsgrad_2 = safe_divide(likvide_omlopsmidler, kortsiktig_gjeld)
    roa = safe_divide(arsresultat, eiendeler)
    roe = safe_divide(arsresultat, egenkapital)

    return {
        "selskap": SELSKAP,
        "organisasjonsnummer": ORGNR,
        "regnskapsaar": int(aar),
        "periode_fra": periode_fra,
        "periode_til": periode_til,
        "valuta": nested_get(regnskap, "valuta"),
        "regnskapstype": nested_get(regnskap, "regnskapstype"),
        "salgsinntekter": salgsinntekter,
        "driftsinntekter": driftsinntekter,
        "lonnskostnad": lonnskostnad,
        "driftskostnader": driftskostnader,
        "driftsresultat": driftsresultat,
        "netto_finans": netto_finans,
        "resultat_for_skatt": resultat_for_skatt,
        "skattekostnad": skattekostnad,
        "arsresultat": arsresultat,
        "anleggsmidler": anleggsmidler,
        "omlopsmidler": omlopsmidler,
        "varer": varer,
        "fordringer": fordringer,
        "bank": bank,
        "sum_eiendeler": eiendeler,
        "egenkapital": egenkapital,
        "kortsiktig_gjeld": kortsiktig_gjeld,
        "langsiktig_gjeld": langsiktig_gjeld,
        "sum_gjeld": gjeld,
        "driftsmargin_pct": driftsmargin * 100 if driftsmargin is not None else None,
        "resultatmargin_pct": resultatmargin * 100 if resultatmargin is not None else None,
        "egenkapitalandel_pct": egenkapitalandel * 100 if egenkapitalandel is not None else None,
        "gjeldsgrad": gjeldsgrad,
        "likviditetsgrad_1": likviditetsgrad_1,
        "likviditetsgrad_2": likviditetsgrad_2,
        "ROA_pct": roa * 100 if roa is not None else None,
        "ROE_pct": roe * 100 if roe is not None else None,
    }


def lag_long_format(df_wide):
    """Konverterer wide-format til analysevennlig long-format for Power BI."""
    metadata = {
        "salgsinntekter": ("Salgsinntekter", "Resultatregnskap", "Driftsinntekter", "NOK", 10),
        "driftsinntekter": ("Driftsinntekter", "Resultatregnskap", "Driftsinntekter", "NOK", 20),
        "lonnskostnad": ("Lønnskostnad", "Resultatregnskap", "Driftskostnader", "NOK", 30),
        "driftskostnader": ("Driftskostnader", "Resultatregnskap", "Driftskostnader", "NOK", 40),
        "driftsresultat": ("Driftsresultat", "Resultatregnskap", "Resultat", "NOK", 50),
        "netto_finans": ("Netto finans", "Resultatregnskap", "Finansposter", "NOK", 60),
        "resultat_for_skatt": ("Resultat før skatt", "Resultatregnskap", "Resultat", "NOK", 70),
        "skattekostnad": ("Skattekostnad", "Resultatregnskap", "Skatt", "NOK", 80),
        "arsresultat": ("Årsresultat", "Resultatregnskap", "Resultat", "NOK", 90),
        "anleggsmidler": ("Anleggsmidler", "Balanse", "Eiendeler", "NOK", 110),
        "omlopsmidler": ("Omløpsmidler", "Balanse", "Eiendeler", "NOK", 120),
        "varer": ("Varer", "Balanse", "Omløpsmidler", "NOK", 130),
        "fordringer": ("Fordringer", "Balanse", "Omløpsmidler", "NOK", 140),
        "bank": ("Bankinnskudd og kontanter", "Balanse", "Omløpsmidler", "NOK", 150),
        "sum_eiendeler": ("Sum eiendeler", "Balanse", "Eiendeler", "NOK", 160),
        "egenkapital": ("Egenkapital", "Balanse", "Egenkapital", "NOK", 170),
        "kortsiktig_gjeld": ("Kortsiktig gjeld", "Balanse", "Gjeld", "NOK", 180),
        "langsiktig_gjeld": ("Langsiktig gjeld", "Balanse", "Gjeld", "NOK", 190),
        "sum_gjeld": ("Sum gjeld", "Balanse", "Gjeld", "NOK", 200),
        "driftsmargin_pct": ("Driftsmargin", "KPI", "Lønnsomhet", "Prosent", 310),
        "resultatmargin_pct": ("Resultatmargin", "KPI", "Lønnsomhet", "Prosent", 320),
        "egenkapitalandel_pct": ("Egenkapitalandel", "KPI", "Soliditet", "Prosent", 330),
        "gjeldsgrad": ("Gjeldsgrad", "KPI", "Soliditet", "Forholdstall", 340),
        "likviditetsgrad_1": ("Likviditetsgrad 1", "KPI", "Likviditet", "Forholdstall", 350),
        "likviditetsgrad_2": ("Likviditetsgrad 2", "KPI", "Likviditet", "Forholdstall", 360),
        "ROA_pct": ("Avkastning på eiendeler", "KPI", "Lønnsomhet", "Prosent", 370),
        "ROE_pct": ("Avkastning på egenkapital", "KPI", "Lønnsomhet", "Prosent", 380),
        "omsetningsvekst_pct": ("Omsetningsvekst", "KPI", "Vekst", "Prosent", 410),
        "resultatvekst_pct": ("Resultatvekst", "KPI", "Vekst", "Prosent", 420),
        "egenkapitalvekst_pct": ("Egenkapitalvekst", "KPI", "Vekst", "Prosent", 430),
    }

    id_columns = [
        "selskap", "organisasjonsnummer", "regnskapsaar",
        "periode_fra", "periode_til", "valuta", "regnskapstype"
    ]
    value_columns = [col for col in metadata if col in df_wide.columns]

    df_long = df_wide.melt(
        id_vars=id_columns,
        value_vars=value_columns,
        var_name="nokkel_kode",
        value_name="verdi",
    )

    df_long["nokkel_navn"] = df_long["nokkel_kode"].map(lambda x: metadata[x][0])
    df_long["kategori"] = df_long["nokkel_kode"].map(lambda x: metadata[x][1])
    df_long["underkategori"] = df_long["nokkel_kode"].map(lambda x: metadata[x][2])
    df_long["enhet"] = df_long["nokkel_kode"].map(lambda x: metadata[x][3])
    df_long["sortering"] = df_long["nokkel_kode"].map(lambda x: metadata[x][4])

    df_long["verdi"] = pd.to_numeric(df_long["verdi"], errors="coerce")
    df_long = df_long.dropna(subset=["verdi"])
    df_long = df_long.sort_values(
        ["organisasjonsnummer", "regnskapsaar", "sortering"]
    ).reset_index(drop=True)

    return df_long[
        [
            "selskap", "organisasjonsnummer", "regnskapsaar",
            "periode_fra", "periode_til", "valuta", "regnskapstype",
            "kategori", "underkategori", "nokkel_kode", "nokkel_navn",
            "enhet", "sortering", "verdi"
        ]
    ]


def main():
    print("=" * 72)
    print(SELSKAP)
    print(f"Organisasjonsnummer: {ORGNR}")
    print("=" * 72)

    session = requests.Session()
    session.headers.update({"User-Agent": "AMOK-regnskapsanalyse/1.0"})

    aar_liste = hent_tilgjengelige_aar(session, ORGNR)
    if not aar_liste:
        print("Ingen tilgjengelige regnskapsår ble funnet.")
        return

    rows = []
    manglende_aar = []

    for aar in aar_liste:
        try:
            regnskap = hent_regnskap_for_aar(session, ORGNR, aar)
            if regnskap is None:
                manglende_aar.append(aar)
            else:
                rows.append(flatten_regnskap(regnskap, aar))
        except requests.RequestException as exc:
            print(f"  Feil ved henting av {aar}: {exc}")
            manglende_aar.append(aar)

        time.sleep(PAUSE_BETWEEN_REQUESTS)

    if not rows:
        print("Ingen strukturerte regnskapsdata ble hentet.")
        return

    df_wide = pd.DataFrame(rows)
    df_wide = df_wide.drop_duplicates(
        subset=["organisasjonsnummer", "regnskapsaar"], keep="first"
    )
    df_wide = df_wide.sort_values("regnskapsaar").reset_index(drop=True)

    df_wide["omsetningsvekst_pct"] = (
        pd.to_numeric(df_wide["driftsinntekter"], errors="coerce")
        .pct_change(fill_method=None)
        .mul(100)
    )
    df_wide["resultatvekst_pct"] = (
        pd.to_numeric(df_wide["arsresultat"], errors="coerce")
        .pct_change(fill_method=None)
        .mul(100)
    )
    df_wide["egenkapitalvekst_pct"] = (
        pd.to_numeric(df_wide["egenkapital"], errors="coerce")
        .pct_change(fill_method=None)
        .mul(100)
    )

    df_long = lag_long_format(df_wide)

    df_wide.to_csv(
        WIDE_OUTPUT_FILE,
        index=False,
        sep=";",
        encoding="utf-8-sig",
        decimal=",",
    )
    df_long.to_csv(
        LONG_OUTPUT_FILE,
        index=False,
        sep=";",
        encoding="utf-8-sig",
        decimal=",",
    )

    print()
    print("Ferdig.")
    print(f"Wide-format: {Path(WIDE_OUTPUT_FILE).resolve()}")
    print(f"Long-format: {Path(LONG_OUTPUT_FILE).resolve()}")
    print(f"Antall strukturerte regnskapsår: {len(df_wide)}")
    print(f"Antall rader i long-format: {len(df_long)}")

    if manglende_aar:
        print(f"År uten strukturerte data: {sorted(set(manglende_aar))}")


if __name__ == "__main__":
    main()
