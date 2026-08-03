"""
NETTHANDELSPRISEN – POSTEN BRING
Agent-logikk v5 med komplett og endelig scoringsmodell
"""

import json
import re
import time
import requests
import pandas as pd
import anthropic

# ─────────────────────────────────────────────
# KONFIGURASJON
# ─────────────────────────────────────────────
LOGISTIKK_AKTORER = [
    "Posten/Bring", "PostNord", "Helthjem", "Instabox",
    "Porterbuddy", "DHL", "Budbee", "UPS/FedEx", "Egne biler"
]

BRANSJER = [
    "Klær og mote", "Sko og vesker", "Sport og friluftsliv",
    "Elektronikk og teknologi", "Helse og apotek", "Hudpleie og skjønnhet",
    "Hjem og interiør", "Mat og dagligvarer", "Bøker og media",
    "Barn og leker", "Musikk og instrument", "Friluftsliv og klatring",
    "Landbruk og maskiner", "Bil og motor", "Blomster og hage",
    "Treningsklær", "Vintage og brukt", "Annet"
]

# ─────────────────────────────────────────────
# SCORINGSMODELL – KOMPLETT VEKTING
# ─────────────────────────────────────────────
#
# KATEGORI 1: Første inntrykk (25% av total)
#   - Startside og navigasjon: 25%
#   - Bilder og film: 25%
#   - Produktinformasjon: 25%
#   - Søkefunksjon: 25%
#
# KATEGORI 2: Info, kundeservice og bærekraft (25% av total)
#   - Kjøpsvilkår, levering og retur: 35%
#     - Kjøpsvilkår: 20%
#     - Levering: 40%
#     - Retur: 40%
#   - Kundeservice: 35%
#     - Selvbetjent: 50%
#     - Betjent: 50%
#   - Bærekraft: 30%
#     - Strategi: 50%
#     - Kundeverktøy: 50%
#
# KATEGORI 3: Kassen, mersalg og inspirasjon (25% av total)
#   - Kassen: 50%
#     - Innlogging/identifisering: 20%
#     - Leveringsvalg og fleksibilitet: 30%
#     - Leveringspris: 20%
#     - Leveringstid og presisjon: 15%
#     - Betalingsalternativer: 15%
#   - Mersalg: 25%
#   - Inspirasjon: 25%
#
# KATEGORI 4: Markedsføring og kundedialog (25% av total)
#   - Sosiale medier: 40%
#   - Kundeklubb: 30% (dynamisk – se logikk)
#   - Nyhetsbrev: 30%
#
# NB: Stor-klassen bedømmes strengere enn Medium og Liten.

IKS_VEKTER = {
    "kjopsvilkar_levering_retur": 0.35,
    "kundeservice": 0.35,
    "baerekraft": 0.30,
}

KLR_VEKTER = {
    "kjopsvilkar": 0.20,
    "levering": 0.40,
    "retur": 0.40,
}

KASSEN_VEKTER = {
    "innlogging": 0.20,
    "leveringsvalg": 0.30,
    "leveringspris": 0.20,
    "leveringstid": 0.15,
    "betaling": 0.15,
}


# ─────────────────────────────────────────────
# JSON-HJELPERE
# ─────────────────────────────────────────────
def rens_json(tekst: str) -> str:
    if not tekst:
        return ""
    tekst = re.sub(r"```json\s*", "", tekst)
    tekst = re.sub(r"```\s*", "", tekst)
    tekst = tekst.strip()
    start = tekst.find("{")
    if start == -1:
        return tekst
    nivaa = 0
    slutt = -1
    inni = False
    forrige = ""
    for i in range(start, len(tekst)):
        c = tekst[i]
        if c == '"' and forrige != "\\":
            inni = not inni
        if not inni:
            if c == "{":
                nivaa += 1
            elif c == "}":
                nivaa -= 1
                if nivaa == 0:
                    slutt = i + 1
                    break
        forrige = c
    if slutt > start:
        tekst = tekst[start:slutt]
    tekst = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', tekst)
    return tekst


def fiks_linjeskift(tekst: str) -> str:
    res = []
    inni = False
    forrige = ""
    for c in tekst:
        if c == '"' and forrige != "\\":
            inni = not inni
            res.append(c)
        elif inni and c == "\n":
            res.append("\\n")
        elif inni and c == "\t":
            res.append("\\t")
        else:
            res.append(c)
        forrige = c
    return "".join(res)


def parse_json(tekst: str):
    if not tekst:
        return None
    renset = rens_json(tekst)
    for forsok in [
        lambda t: json.loads(t),
        lambda t: json.loads(fiks_linjeskift(t)),
        lambda t: json.loads(re.sub(r',\s*([}\]])', r'\1', fiks_linjeskift(t))),
    ]:
        try:
            return forsok(renset)
        except Exception:
            pass
    return None


# ─────────────────────────────────────────────
# EXCEL OG DUPLIKATER
# ─────────────────────────────────────────────
def les_excel(filbane) -> list:
    try:
        df = pd.read_excel(filbane, header=None)
        navn = []
        for col in df.columns:
            for v in df[col].dropna():
                n = str(v).strip()
                if n and n.lower() not in ("nan", "nettbutikk", "navn"):
                    navn.append(n)
        return navn
    except Exception:
        return []


def fjern_duplikater(navn: list) -> list:
    sett = {}
    for n in navn:
        key = re.sub(r"[^a-zæøå0-9]", "", n.lower())
        if key not in sett:
            sett[key] = n
    return list(sett.values())


# ─────────────────────────────────────────────
# BRØNNØYSUND OG REGNSKAP
# ─────────────────────────────────────────────
def sok_brreg(navn: str, selskapsnavn: str = "") -> dict:
    sokeliste = []
    if selskapsnavn and selskapsnavn != navn:
        sokeliste.append(selskapsnavn)
    sokeliste.append(navn)
    for sok in sokeliste:
        try:
            url = f"https://data.brreg.no/enhetsregisteret/api/enheter?navn={requests.utils.quote(sok)}&size=10"
            r = requests.get(url, timeout=8)
            data = r.json()
            enheter = data.get("_embedded", {}).get("enheter", [])
            if not enheter:
                continue
            for e in enheter:
                kode = e.get("organisasjonsform", {}).get("kode", "")
                if kode in ("AS", "ASA", "NUF", "AB"):
                    return {
                        "orgform": kode,
                        "orgnr": str(e.get("organisasjonsnummer", "")),
                        "registrert_navn": e.get("navn", navn),
                        "usikker": False,
                    }
            e = enheter[0]
            return {
                "orgform": e.get("organisasjonsform", {}).get("kode", "Ukjent"),
                "orgnr": str(e.get("organisasjonsnummer", "")),
                "registrert_navn": e.get("navn", navn),
                "usikker": True,
            }
        except Exception:
            continue
    return {"orgform": "Ukjent", "orgnr": "", "registrert_navn": navn, "usikker": True}


def hent_regnskap(orgnr: str) -> dict:
    if not orgnr:
        return {}
    try:
        url = f"https://data.brreg.no/regnskapsregisteret/regnskap/{orgnr}"
        r = requests.get(url, timeout=8)
        if r.status_code == 200:
            data = r.json()
            if isinstance(data, list) and data:
                siste = data[0]
                res = siste.get("resultatregnskapResultat", {})
                omsetning = res.get("driftsinntekter", {}).get("sumDriftsinntekter")
                aar = siste.get("regnskapsperiode", {}).get("fraDato", "")[:4]
                if omsetning:
                    return {"omsetning_kr": omsetning, "aar": aar}
    except Exception:
        pass
    return {}


def er_enk(orgform: str) -> bool:
    return orgform.upper() in ("ENK", "ENKELTPERSONFORETAK")


# ─────────────────────────────────────────────
# URL OG INFO
# ─────────────────────────────────────────────
def finn_url_og_info(navn: str, client) -> dict:
    bransje_liste = ", ".join(BRANSJER)
    prompt = f"""Søk på nettet og finn informasjon om denne norske nettbutikken: "{navn}"

Finn URL-en til nettbutikken, offisielt selskapsnavn og bransje.

Svar KUN med JSON uten markdown:
{{"url": "https://...", "selskapsnavn": "Offisielt AS-navn fra Brreg", "bransje": "velg fra: {bransje_liste}", "land": "NO", "omsetning_est": "Under 50 mill", "selger_b2c_fysisk": true, "skandinavisk_tilstedevarelse": true, "kommentar": "kort begrunnelse"}}

For omsetning_est velg: Under 50 mill, 50-250 mill, Over 250 mill, Ukjent
For selskapsnavn: skriv det juridiske selskapsnavnet (f.eks Linda JC AS for Linda Johansen)
Hvis ikke funnet: url = "" """

    for forsok in range(2):
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=800,
                tools=[{"type": "web_search_20250305", "name": "web_search"}],
                messages=[{"role": "user", "content": prompt}]
            )

            # Hent ALL tekst fra alle blokker inkludert tool_result
            tekst = ""
            for blokk in resp.content:
                if hasattr(blokk, "type"):
                    if blokk.type == "text":
                        tekst += blokk.text + "\n"
                    elif blokk.type == "tool_result":
                        # Tool result kan inneholde søkeresultater
                        if hasattr(blokk, "content"):
                            for c in blokk.content:
                                if hasattr(c, "text"):
                                    tekst += c.text + "\n"

            # Prøv å parse JSON fra teksten
            data = parse_json(tekst)
            if data and data.get("url", "").startswith("http"):
                return data

            # Hvis første forsøk ikke ga URL, prøv uten websøk
            if forsok == 0:
                resp2 = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=600,
                    messages=[{"role": "user", "content": f"""Hva er URL-en til den norske nettbutikken "{navn}"?
Svar KUN med JSON:
{{"url": "https://...", "selskapsnavn": "navn AS", "bransje": "velg fra: {bransje_liste}", "land": "NO", "omsetning_est": "Under 50 mill", "selger_b2c_fysisk": true, "skandinavisk_tilstedevarelse": true, "kommentar": "begrunnelse"}}"""}]
                )
                tekst2 = resp2.content[0].text if resp2.content else ""
                data2 = parse_json(tekst2)
                if data2 and data2.get("url", "").startswith("http"):
                    return data2

        except Exception as e:
            if forsok == 0:
                time.sleep(2)

    return {
        "url": "", "selskapsnavn": navn, "bransje": "Annet", "land": "Ukjent",
        "omsetning_est": "Ukjent", "selger_b2c_fysisk": True,
        "skandinavisk_tilstedevarelse": True, "kommentar": "URL ikke funnet"
    }


# ─────────────────────────────────────────────
# KLASSIFISERING OG KO-SJEKK
# ─────────────────────────────────────────────
def klassifiser(regnskap: dict, info: dict) -> tuple:
    omsetning_kr = regnskap.get("omsetning_kr")
    aar = regnskap.get("aar", "")
    if omsetning_kr:
        mill = omsetning_kr / 1_000_000
        tekst = f"{mill:.1f} mill ({aar})"
        if mill >= 250:
            return "Stor", tekst
        elif mill >= 50:
            return "Medium", tekst
        else:
            return "Liten", tekst
    omsetning_est = info.get("omsetning_est", "Ukjent")
    if "Over 250" in omsetning_est:
        return "Stor", "Est. over 250 mill"
    elif "50-250" in omsetning_est:
        return "Medium", "Est. 50-250 mill"
    elif "Under 50" in omsetning_est:
        return "Liten", "Est. under 50 mill"
    return "Ukjent", "Ukjent"


def ko_sjekk(brreg: dict, info: dict) -> tuple:
    if er_enk(brreg.get("orgform", "")) and not brreg.get("usikker", False):
        return True, "ENK bekreftet – enkeltmannsforetak filtreres ut"
    return False, ""


def bygg_advarsler(brreg: dict, info: dict) -> list:
    advarsler = []
    if brreg.get("usikker"):
        advarsler.append("Brreg-treff usikkert – org.form bør verifiseres manuelt")
    if not info.get("url", ""):
        advarsler.append("Ingen URL funnet – agenten klarte ikke å identifisere nettbutikken. Sjekk manuelt.")
    if not info.get("selger_b2c_fysisk", True):
        advarsler.append("Agenten er usikker på om dette er B2C fysiske varer – sjekk manuelt")
    if not info.get("skandinavisk_tilstedevarelse", True):
        advarsler.append("Agenten er usikker på skandinavisk tilstedeværelse – sjekk manuelt")
    return advarsler


# ─────────────────────────────────────────────
# SCORING – KOMPLETT OPPDATERT MODELL
# ─────────────────────────────────────────────
def score_nettbutikk(navn: str, url: str, klasse: str, client) -> dict:
    """
    Scorer i to separate kall:
    Kall 1: Første inntrykk + IKS
    Kall 2: Kassen/mersalg/inspirasjon + Markedsføring + Logistikk
    """
    strengere = "NB: Klasse Stor bedømmes strengere og med høyere krav enn Medium og Liten. " if klasse == "Stor" else ""

    prompt1 = f"""Vurder nettbutikken "{navn}" på URL: {url}
Klasse: {klasse}. {strengere}
Score fra 1-5. KORTE begrunnelser UTEN linjeskift.

KATEGORI 1 – FØRSTE INNTRYKK (25% av total, 4 kriterier à 25%):

startside: Visuelt ryddig? Lett å forstå hva de selger? Identitet og inspirasjon (ikke bare bestselgere)?
Logisk menystruktur? Filtrering på farge, størrelse, pris, bruksområde (ikke bare nyeste og pris)?

bilder_film: Produkter i miljø (klær på modell med høyde/størrelse oppgitt, sofa i møblert stue)?
Bilder fra flere vinkler? Film av produkter der relevant?

produktinfo: Materiale og mål oppgitt der relevant? Godt strukturert og lett å forstå?
Produktanmeldelser fra kjøpere med info om størrelse og bruk?

sokefunksjon: Gir korrekt resultat ved søk med to variabler?
Test: rød kjole, stekepanne 24 cm induksjon, blå genser str medium, kaffemaskin med melkeskummer.

KATEGORI 2 – INFO, KUNDESERVICE OG BÆREKRAFT (25% av total):

Kjøpsvilkår, levering og retur (35% av kat2):
kjopsvilkar: Lovpålagt info om kjøpsvilkår tilgjengelig? (ja=nøytral, nei=MEGET negativt)
levering: Leveringspris/fri frakt synlig FØR checkout? Klikkbar for mer info? Alle detaljer:
  alternativer, priser, transportører, leveringstid (inkl plukk+pakk). Brukes frakt som USP?
retur: Returinfo på startside og produktside? Kostnaden, fremgangsmåte, tilbakebetalingstid, returfrist?
  Tilbyr de returløsning med etikett? (Nei = dårlig score)

Kundeservice (35% av kat2):
selvbetjent_ks: Oppdatert FAQ? Chatbot uten krav om ordrenummer? (Chatbot med ordrenr-krav gir lav verdi)
betjent_ks: Tydelig info om kanaler, åpningstider, responstid? Lett å finne?
  (Skjult kontaktinfo/tvinger via FAQ = trekk. God AI chatbot 24/7 = pluss)

Bærekraft (30% av kat2):
baerekraft_strategi: Bærekraftsrapport eller sertifiseringer? Åpenhetsloven-rapport?
  (Pålagt over 70 mill omsetning, 35 mill balansesum eller 50 årsverk)
baerekraft_kunder: Filtrering på bærekraftige produkter? Resalg, panteordninger?
  GOTS-sertifisering eller andre produktsertifiseringer?

Svar KUN med JSON:
{{"inntrykk": {{"startside": {{"score": 3, "begrunnelse": "setning"}}, "bilder_film": {{"score": 3, "begrunnelse": "setning"}}, "produktinfo": {{"score": 3, "begrunnelse": "setning"}}, "sokefunksjon": {{"score": 3, "begrunnelse": "setning"}}}}, "iks": {{"kjopsvilkar": {{"score": 3, "begrunnelse": "setning"}}, "levering": {{"score": 3, "begrunnelse": "setning"}}, "retur": {{"score": 3, "begrunnelse": "setning"}}, "selvbetjent_ks": {{"score": 3, "begrunnelse": "setning"}}, "betjent_ks": {{"score": 3, "begrunnelse": "setning"}}, "baerekraft_strategi": {{"score": 3, "begrunnelse": "setning"}}, "baerekraft_kunder": {{"score": 3, "begrunnelse": "setning"}}}}}}"""

    prompt2 = f"""Vurder nettbutikken "{navn}" på URL: {url}
Klasse: {klasse}. {strengere}
Score fra 1-5. KORTE begrunnelser UTEN linjeskift.

KATEGORI 3 – KASSEN, MERSALG OG INSPIRASJON (25% av total):
Kassen teller 50%, Mersalg 25%, Inspirasjon 25%.

Kassen (50% av kat3) – 5 underkriterier:
innlogging: Kan kunden identifisere seg med Vipps/Klarna/tilsvarende for enklere utfylling?
  (Ja = høyere score, må fylle manuelt = lavere)

leveringsvalg: Har de minst 2 leveringsalternativer? Nedtrekksmeny for valg av hentested/pakkeboks?
  (Nedtrekksmeny = høyere score enn forhåndsvalgt/mange separate valg som krever scrolling)
  Nærhet til kunden viktig – mange utleveringssteder/pakkebokser gir høyere score.
  Relevante valg: tyngre varer bør ha hjemlevering, lette varer bør ha postkasse/Helthjem.

leveringspris: Hva er billigste leveringspris? Under 99 kr = god score, over 129 kr = trekk
  (spesielt hvis varene ikke er store/tunge/eksklusivt sortiment)

leveringstid: Oppgir de estimert leveringsdato = best. Tidsintervall (1-3 dager) = middels.
  Ingen info eller kun transporttid = dårlig. Leveringstid = plukk+pakk + transport.

betaling: Full score hvis alle vanlige: kredit/debetkort, Vipps, Klarna/etterbetaling/delbetaling.

Mersalg (25% av kat3):
mersalg: Relevante produktanbefalinger på produktside (komplementerende ELLER alternativer)?
  Tilbud på vei til kassen og i kassen? Er anbefalingene relevante for det kunden ser på?

Inspirasjon (25% av kat3):
inspirasjon: Artikler, oppskrifter med produktlenker, filmer, guider?
  Gir faglig tillit OG selger mer? (Eks: oppskrifter med kjøkkenutstyr, friluftsliv-filmer med utstyr)

KATEGORI 4 – MARKEDSFØRING OG KUNDEDIALOG (25% av total):
SoMe 40%, Kundeklubb 30% (dynamisk), Nyhetsbrev 30%.

some: Hvilke plattformer? Engasjement og dialog med kunder?
  Autentisk innhold (ikke katalogbilder)? Troverdige influensere?
  Kjøpslenker? SoMe som kundeservice?
  IKKE antall følgere – engasjement og autentisitet teller!

kundeklubb: Score slik:
  2 = Ingen kundeklubb
  3 = Kundeklubb men KUN inngangsrabatt ved første kjøp
  4 = Kundeklubb med poeng, rabatter og lojalitetsfordeler
  5 = Full lojalitetspakke med poeng, rabatter, fortrinn til salg og egne tilbud

nyhetsbrev: Har de nyhetsbrev? Rekrutterer aktivt via kjøpsreisen (popup, påmelding i kasse)?
  SMS + e-post? Aktiv rekruttering = pluss.

Kartlegg også (ikke score):
logistikk: Hvilke brukes? {", ".join(LOGISTIKK_AKTORER)}
tech: Mobiloptimalisert? SSL? Lastetid (ok/warn/bad)? Trustpilot-score?
trust: Liste med tillitselementer (garantier, sertifiseringer, kundeomtaler, fysisk adresse synlig osv)
apenhetsloven: Er rapport funnet for 2024/2025? Er de rapporteringspliktige?

Svar KUN med JSON:
{{"kommentar": "2-3 setninger om butikken totalt", "kassen": {{"innlogging": {{"score": 3, "begrunnelse": "setning"}}, "leveringsvalg": {{"score": 3, "begrunnelse": "setning"}}, "leveringspris": {{"score": 3, "begrunnelse": "setning"}}, "leveringstid": {{"score": 3, "begrunnelse": "setning"}}, "betaling": {{"score": 3, "begrunnelse": "setning"}}}}, "mersalg": {{"score": 3, "begrunnelse": "setning"}}, "inspirasjon": {{"score": 3, "begrunnelse": "setning"}}, "markedsforing": {{"some": {{"score": 3, "begrunnelse": "setning"}}, "kundeklubb": {{"score": 2, "begrunnelse": "setning", "har_kundeklubb": false}}, "nyhetsbrev": {{"score": 3, "begrunnelse": "setning"}}}}, "logistikk": {{"Posten/Bring": false, "PostNord": false, "Helthjem": false, "Instabox": false, "Porterbuddy": false, "DHL": false, "Budbee": false, "UPS/FedEx": false, "Egne biler": false}}, "tech": {{"mobil": "ok", "ssl": "ok", "lastetid": "ok", "trustpilot": "Ikke funnet"}}, "trust": ["element1", "element2"], "apenhetsloven": {{"palagt": false, "rapport_funnet": false, "kommentar": "setning"}}}}"""

    def kall(prompt, nr):
        for forsok in range(2):
            try:
                resp = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1200,
                    tools=[{"type": "web_search_20250305", "name": "web_search"}],
                    messages=[{"role": "user", "content": prompt}]
                )
                # Hent tekst fra alle blokker inkludert tool_result
                tekst = ""
                for blokk in resp.content:
                    if hasattr(blokk, "type"):
                        if blokk.type == "text":
                            tekst += blokk.text + "\n"
                        elif blokk.type == "tool_result":
                            if hasattr(blokk, "content"):
                                for c in blokk.content:
                                    if hasattr(c, "text"):
                                        tekst += c.text + "\n"
                data = parse_json(tekst)
                if data:
                    return data
                resp2 = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1200,
                    messages=[{"role": "user", "content": f"Fiks til gyldig JSON, svar KUN med JSON:\n{tekst[:2000]}"}]
                )
                data2 = parse_json(resp2.content[0].text if resp2.content else "")
                if data2:
                    return data2
            except Exception:
                if forsok == 0:
                    time.sleep(2)
        return None

    del1 = kall(prompt1, 1)
    time.sleep(3)
    del2 = kall(prompt2, 2)
    time.sleep(3)

    if not del1 and not del2:
        return None

    resultat = {}
    if del1:
        resultat["inntrykk"] = del1.get("inntrykk", {})
        resultat["iks"] = del1.get("iks", {})
    else:
        resultat["inntrykk"] = {}
        resultat["iks"] = {}

    if del2:
        resultat["kommentar"] = del2.get("kommentar", "")
        resultat["kassen"] = del2.get("kassen", {})
        resultat["mersalg"] = del2.get("mersalg", {})
        resultat["inspirasjon"] = del2.get("inspirasjon", {})
        resultat["markedsforing"] = del2.get("markedsforing", {})
        resultat["logistikk"] = del2.get("logistikk", {})
        resultat["tech"] = del2.get("tech", {})
        resultat["trust"] = del2.get("trust", [])
        resultat["apenhetsloven"] = del2.get("apenhetsloven", {})
    else:
        resultat.update({
            "kommentar": "", "kassen": {}, "mersalg": {}, "inspirasjon": {},
            "markedsforing": {}, "logistikk": {}, "tech": {}, "trust": [], "apenhetsloven": {}
        })

    return resultat


# ─────────────────────────────────────────────
# BEREGN TOTALSCORE
# ─────────────────────────────────────────────
def beregn_totalscore(scoring: dict) -> dict:
    # 1. Første inntrykk – 4 kriterier, 25% vekt hver
    inntrykk = scoring.get("inntrykk", {})
    inn_scores = [v["score"] for v in inntrykk.values() if isinstance(v, dict) and "score" in v]
    inn_snitt = round(sum(inn_scores) / len(inn_scores), 2) if inn_scores else 0

    # 2. IKS
    iks = scoring.get("iks", {})
    klr_snitt = (
        iks.get("kjopsvilkar", {}).get("score", 0) * KLR_VEKTER["kjopsvilkar"] +
        iks.get("levering", {}).get("score", 0) * KLR_VEKTER["levering"] +
        iks.get("retur", {}).get("score", 0) * KLR_VEKTER["retur"]
    )
    ks_snitt = (
        iks.get("selvbetjent_ks", {}).get("score", 0) * 0.5 +
        iks.get("betjent_ks", {}).get("score", 0) * 0.5
    )
    bk_snitt = (
        iks.get("baerekraft_strategi", {}).get("score", 0) * 0.5 +
        iks.get("baerekraft_kunder", {}).get("score", 0) * 0.5
    )
    iks_snitt = round(
        klr_snitt * IKS_VEKTER["kjopsvilkar_levering_retur"] +
        ks_snitt * IKS_VEKTER["kundeservice"] +
        bk_snitt * IKS_VEKTER["baerekraft"], 2
    )

    # 3. Kassen (50%) + Mersalg (25%) + Inspirasjon (25%)
    kassen = scoring.get("kassen", {})
    kassescore = (
        kassen.get("innlogging", {}).get("score", 0) * KASSEN_VEKTER["innlogging"] +
        kassen.get("leveringsvalg", {}).get("score", 0) * KASSEN_VEKTER["leveringsvalg"] +
        kassen.get("leveringspris", {}).get("score", 0) * KASSEN_VEKTER["leveringspris"] +
        kassen.get("leveringstid", {}).get("score", 0) * KASSEN_VEKTER["leveringstid"] +
        kassen.get("betaling", {}).get("score", 0) * KASSEN_VEKTER["betaling"]
    )
    mersalg_score = scoring.get("mersalg", {}).get("score", 0)
    insp_score = scoring.get("inspirasjon", {}).get("score", 0)
    kat3_snitt = round(kassescore * 0.50 + mersalg_score * 0.25 + insp_score * 0.25, 2)

    # 4. Markedsføring – dynamisk kundeklubb
    mf = scoring.get("markedsforing", {})
    some_score = mf.get("some", {}).get("score", 0)
    kk_score = mf.get("kundeklubb", {}).get("score", 0)
    nb_score = mf.get("nyhetsbrev", {}).get("score", 0)
    har_kk = mf.get("kundeklubb", {}).get("har_kundeklubb", kk_score > 2)

    if har_kk:
        mf_snitt = round(some_score * 0.40 + kk_score * 0.30 + nb_score * 0.30, 2)
    else:
        mf_snitt = round(some_score * (40/70) + nb_score * (30/70), 2)

    total = round(inn_snitt * 0.25 + iks_snitt * 0.25 + kat3_snitt * 0.25 + mf_snitt * 0.25, 1)

    return {
        "total": total,
        "kategorier": {
            "inntrykk": round(inn_snitt, 1),
            "iks": round(iks_snitt, 1),
            "iks_detalj": {
                "klr": round(klr_snitt, 1),
                "kundeservice": round(ks_snitt, 1),
                "baerekraft": round(bk_snitt, 1),
            },
            "kat3": round(kat3_snitt, 1),
            "kat3_detalj": {
                "kassen": round(kassescore, 1),
                "mersalg": mersalg_score,
                "inspirasjon": insp_score,
            },
            "markedsforing": round(mf_snitt, 1),
            "mf_detalj": {
                "some": some_score,
                "kundeklubb": kk_score,
                "nyhetsbrev": nb_score,
                "har_kundeklubb": har_kk,
            }
        }
    }


def bygg_scoring_for_visning(scoring_raw: dict) -> dict:
    if not scoring_raw:
        return None
    return {
        "inntrykk": [
            {"navn": "Startside og navigasjon", "score": scoring_raw.get("inntrykk", {}).get("startside", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("startside", {}).get("begrunnelse", ""), "vekt": "25%"},
            {"navn": "Bilder og film", "score": scoring_raw.get("inntrykk", {}).get("bilder_film", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("bilder_film", {}).get("begrunnelse", ""), "vekt": "25%"},
            {"navn": "Produktinformasjon", "score": scoring_raw.get("inntrykk", {}).get("produktinfo", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("produktinfo", {}).get("begrunnelse", ""), "vekt": "25%"},
            {"navn": "Søkefunksjon", "score": scoring_raw.get("inntrykk", {}).get("sokefunksjon", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("sokefunksjon", {}).get("begrunnelse", ""), "vekt": "25%"},
        ],
        "iks": [
            {"navn": "Kjøpsvilkår (lovkrav)", "score": scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("begrunnelse", ""), "vekt": "KLR 20%"},
            {"navn": "Leveringsinformasjon", "score": scoring_raw.get("iks", {}).get("levering", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("levering", {}).get("begrunnelse", ""), "vekt": "KLR 40%"},
            {"navn": "Returløsning", "score": scoring_raw.get("iks", {}).get("retur", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("retur", {}).get("begrunnelse", ""), "vekt": "KLR 40%"},
            {"navn": "Selvbetjent kundeservice", "score": scoring_raw.get("iks", {}).get("selvbetjent_ks", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("selvbetjent_ks", {}).get("begrunnelse", ""), "vekt": "KS 50%"},
            {"navn": "Betjent kundeservice", "score": scoring_raw.get("iks", {}).get("betjent_ks", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("betjent_ks", {}).get("begrunnelse", ""), "vekt": "KS 50%"},
            {"navn": "Bærekraft – strategi og rapportering", "score": scoring_raw.get("iks", {}).get("baerekraft_strategi", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("baerekraft_strategi", {}).get("begrunnelse", ""), "vekt": "BK 50%"},
            {"navn": "Bærekraft – kundeverktøy", "score": scoring_raw.get("iks", {}).get("baerekraft_kunder", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("baerekraft_kunder", {}).get("begrunnelse", ""), "vekt": "BK 50%"},
        ],
        "kassen": [
            {"navn": "Innlogging og identifisering", "score": scoring_raw.get("kassen", {}).get("innlogging", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("innlogging", {}).get("begrunnelse", ""), "vekt": "Kassen 20%"},
            {"navn": "Leveringsvalg og fleksibilitet", "score": scoring_raw.get("kassen", {}).get("leveringsvalg", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("leveringsvalg", {}).get("begrunnelse", ""), "vekt": "Kassen 30%"},
            {"navn": "Leveringspris", "score": scoring_raw.get("kassen", {}).get("leveringspris", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("leveringspris", {}).get("begrunnelse", ""), "vekt": "Kassen 20%"},
            {"navn": "Leveringstid og presisjon", "score": scoring_raw.get("kassen", {}).get("leveringstid", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("leveringstid", {}).get("begrunnelse", ""), "vekt": "Kassen 15%"},
            {"navn": "Betalingsalternativer", "score": scoring_raw.get("kassen", {}).get("betaling", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("betaling", {}).get("begrunnelse", ""), "vekt": "Kassen 15%"},
            {"navn": "Mersalg og produktanbefalinger", "score": scoring_raw.get("mersalg", {}).get("score", 0), "begrunnelse": scoring_raw.get("mersalg", {}).get("begrunnelse", ""), "vekt": "Mersalg 25%"},
            {"navn": "Inspirasjon og innhold", "score": scoring_raw.get("inspirasjon", {}).get("score", 0), "begrunnelse": scoring_raw.get("inspirasjon", {}).get("begrunnelse", ""), "vekt": "Inspirasjon 25%"},
        ],
        "markedsforing": [
            {"navn": "Sosiale medier", "score": scoring_raw.get("markedsforing", {}).get("some", {}).get("score", 0), "begrunnelse": scoring_raw.get("markedsforing", {}).get("some", {}).get("begrunnelse", ""), "vekt": "SoMe 40%"},
            {"navn": "Kundeklubb", "score": scoring_raw.get("markedsforing", {}).get("kundeklubb", {}).get("score", 0), "begrunnelse": scoring_raw.get("markedsforing", {}).get("kundeklubb", {}).get("begrunnelse", ""), "vekt": "Kundeklubb 30% (dynamisk)"},
            {"navn": "Nyhetsbrev og SMS", "score": scoring_raw.get("markedsforing", {}).get("nyhetsbrev", {}).get("score", 0), "begrunnelse": scoring_raw.get("markedsforing", {}).get("nyhetsbrev", {}).get("begrunnelse", ""), "vekt": "Nyhetsbrev 30%"},
        ],
    }


# ─────────────────────────────────────────────
# HOVED-AGENT
# ─────────────────────────────────────────────
def kjor_agent(filbane, api_key: str, maks: int = None, fremgang_callback=None) -> dict:
    client = anthropic.Anthropic(api_key=api_key)

    raa_navn = les_excel(filbane)
    if not raa_navn:
        return {}

    navn_liste = fjern_duplikater(raa_navn)
    if maks:
        navn_liste = navn_liste[:maks]

    resultater = {}
    totalt = len(navn_liste)

    for i, navn in enumerate(navn_liste):
        prosent = int((i / totalt) * 100)
        if fremgang_callback:
            fremgang_callback(f"[{i+1}/{totalt}] Behandler: {navn}", prosent)

        # A – URL og selskapsnavn
        info = finn_url_og_info(navn, client)
        selskapsnavn = info.get("selskapsnavn", navn)
        url = info.get("url", "")
        time.sleep(3)

        # B – Brønnøysund
        brreg = sok_brreg(navn, selskapsnavn)

        # C – Regnskapstall
        regnskap = {}
        if brreg.get("orgnr"):
            regnskap = hent_regnskap(brreg["orgnr"])

        klasse, omsetning_tekst = klassifiser(regnskap, info)
        skal_ut, aarsak = ko_sjekk(brreg, info)
        advarsler = bygg_advarsler(brreg, info)

        butikk = {
            "name": navn,
            "url": url,
            "orgform": brreg.get("orgform", "Ukjent"),
            "orgnr": brreg.get("orgnr", ""),
            "enk": er_enk(brreg.get("orgform", "")) and not brreg.get("usikker", False),
            "bransje": info.get("bransje", "Annet"),
            "omsetning": omsetning_tekst,
            "klasse": klasse,
            "advarsler": advarsler,
            "juryStatus": "Ikke vurdert",
            "juryScore": 0,
            "juryNote": "",
            "trust": [],
            "logistikk": None,
            "tech": None,
            "apenhetsloven": None,
            "scoring": None,
            "total": None,
            "inntrykk": None,
            "iks": None,
            "iksDetalj": None,
            "kat3": None,
            "kat3Detalj": None,
            "markedsforing": None,
            "mfDetalj": None,
        }

        if skal_ut:
            butikk["status"] = "ut"
            butikk["screeningBegrunnelse"] = f"Filtrert ut: {aarsak}"
            butikk["kommentar"] = "Filtrert ut i screening."
        elif not url:
            butikk["status"] = "usikker"
            butikk["screeningBegrunnelse"] = f"{brreg.get('orgform','?')}. Ingen URL funnet – kan ikke score."
            butikk["kommentar"] = "Ingen URL funnet – sjekk manuelt."
            if fremgang_callback:
                fremgang_callback(f"[{i+1}/{totalt}] ⚠ Ingen URL for {navn}", prosent)
        else:
            brreg_note = " | Brreg-treff usikkert" if brreg.get("usikker") else ""
            butikk["screeningBegrunnelse"] = (
                f"{brreg.get('orgform','?')}. Klasse {klasse}. {omsetning_tekst}.{brreg_note}"
            )

            if fremgang_callback:
                fremgang_callback(f"[{i+1}/{totalt}] Scorer: {navn} ({url})", prosent)

            scoring_raw = score_nettbutikk(navn, url, klasse, client)

            if scoring_raw:
                totaler = beregn_totalscore(scoring_raw)
                butikk["status"] = "inn"
                butikk["total"] = totaler["total"]
                butikk["inntrykk"] = totaler["kategorier"]["inntrykk"]
                butikk["iks"] = totaler["kategorier"]["iks"]
                butikk["iksDetalj"] = totaler["kategorier"].get("iks_detalj")
                butikk["kat3"] = totaler["kategorier"]["kat3"]
                butikk["kat3Detalj"] = totaler["kategorier"].get("kat3_detalj")
                butikk["markedsforing"] = totaler["kategorier"]["markedsforing"]
                butikk["mfDetalj"] = totaler["kategorier"].get("mf_detalj")
                butikk["kommentar"] = scoring_raw.get("kommentar", "")
                butikk["logistikk"] = scoring_raw.get("logistikk")
                butikk["tech"] = scoring_raw.get("tech")
                butikk["apenhetsloven"] = scoring_raw.get("apenhetsloven")
                butikk["trust"] = scoring_raw.get("trust", [])
                butikk["scoring"] = bygg_scoring_for_visning(scoring_raw)
            else:
                butikk["status"] = "usikker"
                butikk["kommentar"] = "Scoring feilet – krever manuell sjekk."

        resultater[navn] = butikk

        if fremgang_callback:
            fremgang_callback(f"[{i+1}/{totalt}] ✓ {navn} ferdig", prosent)

    if fremgang_callback:
        fremgang_callback("✅ Alle butikker behandlet!", 100)

    return resultater
