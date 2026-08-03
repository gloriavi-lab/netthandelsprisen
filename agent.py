"""
NETTHANDELSPRISEN – POSTEN BRING
Agent-logikk med oppdatert scoringsmodell
"""

import json
import re
import time
import requests
import pandas as pd
from pathlib import Path
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

# Scoringsmodell – vekting
# Kategori 1: Første inntrykk (25% av total) – 4 kriterier, 25% vekt hver
# Kategori 2: Info, kundeservice og bærekraft (25% av total)
#   KLR 35%: kjøpsvilkår 20%, levering 40%, retur 40%
#   Kundeservice 35%: selvbetjent 50%, betjent 50%
#   Bærekraft 30%: strategi 50%, kundeverktøy 50%
# Kategori 3: Kassen / mersalg (25% av total) – 3 kriterier, lik vekt
# Kategori 4: Markedsføring (25% av total)
#   SoMe 40%, Kundeklubb 30% (dynamisk), Nyhetsbrev 30%

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
    except Exception as e:
        return []


def fjern_duplikater(navn: list) -> list:
    sett = {}
    for n in navn:
        key = re.sub(r"[^a-zæøå0-9]", "", n.lower())
        if key not in sett:
            sett[key] = n
    return list(sett.values())


# ─────────────────────────────────────────────
# BRØNNØYSUND
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
    prompt = f"""Finn informasjon om denne norske nettbutikken: "{navn}"

Svar KUN med JSON uten markdown:
{{"url": "https://...", "selskapsnavn": "Offisielt AS-navn fra Brreg", "bransje": "velg fra: {bransje_liste}", "land": "NO", "omsetning_est": "Under 50 mill", "selger_b2c_fysisk": true, "skandinavisk_tilstedevarelse": true, "kommentar": "kort begrunnelse"}}

For omsetning_est velg: Under 50 mill, 50-250 mill, Over 250 mill, Ukjent"""

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=600,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": prompt}]
        )
        tekst = ""
        for blokk in resp.content:
            if blokk.type == "text":
                tekst += blokk.text + "\n"
        data = parse_json(tekst)
        if data:
            return data
    except Exception:
        pass
    return {
        "url": "", "selskapsnavn": navn, "bransje": "Annet", "land": "Ukjent",
        "omsetning_est": "Ukjent", "selger_b2c_fysisk": True,
        "skandinavisk_tilstedevarelse": True, "kommentar": "Ikke funnet"
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
    if not info.get("url", ""):
        return True, "Kunne ikke identifisere nettbutikk – ingen URL funnet"
    return False, ""


def b2c_og_skandinav_advarsel(info: dict) -> str:
    advarsler = []
    if not info.get("selger_b2c_fysisk", True):
        advarsler.append("Agenten er usikker på om dette er B2C fysiske varer.")
    if not info.get("skandinavisk_tilstedevarelse", True):
        advarsler.append("Agenten er usikker på skandinavisk tilstedeværelse.")
    return " | ".join(advarsler)


# ─────────────────────────────────────────────
# SCORING – OPPDATERT MODELL
# ─────────────────────────────────────────────
def score_nettbutikk(navn: str, url: str, klasse: str, client) -> dict:
    """
    Scorer i to separate kall for å unngå for lang JSON.
    Kall 1: Første inntrykk + IKS
    Kall 2: Kassen + Markedsføring (SoMe, Kundeklubb, Nyhetsbrev) + Logistikk
    """

    prompt1 = f"""Vurder nettbutikken "{navn}" på {url} (klasse: {klasse}).

KATEGORI 1 – FØRSTE INNTRYKK (4 kriterier, 25% vekt hver):
- startside: Visuelt ryddig? Identitet og inspirasjon (ikke bare bestselgere)? Logisk meny? Filtrering på farge/størrelse/pris/bruksområde?
- bilder_film: Miljøbilder? Klær på modell med høyde? Flere vinkler? Film der relevant?
- produktinfo: Materiale, mål, kjøperanmeldelser med størrelsesinfo? Godt strukturert?
- sokefunksjon: Korrekt svar ved to+ parametere? (rød kaffekopp, blå genser str M, stekepanne 24cm induksjon)

KATEGORI 2 – INFO, KUNDESERVICE OG BÆREKRAFT:
Kjøpsvilkår/levering/retur (35% av kat2):
- kjopsvilkar: Lovpålagt info om kjøpsvilkår tilgjengelig? (ja/nei – nei er svært negativt)
- levering: Leveringspris/fri frakt synlig FØR checkout? Alle detaljer: alternativer, priser, transportører, leveringstid inkl plukk og pakk?
- retur: Returinfo på startside og produktside? Kostnad, fremgangsmåte, tilbakebetalingstid? Returløsning med etikett?

Kundeservice (35% av kat2):
- selvbetjent_ks: Oppdatert FAQ? Chatbot som fungerer uten ordrenummer?
- betjent_ks: Kontaktkanaler og åpningstider lett å finne? Gode AI-chatbot scorer høyere.

Bærekraft (30% av kat2):
- baerekraft_strategi: Bærekraftsrapport? Åpenhetsloven-rapport (pålagt over 70 mill/35 mill balansesum/50 årsverk)?
- baerekraft_kunder: Filter på bærekraftige produkter? Resalg? Panteordninger? GOTS-sertifisering?

Score fra 1-5. KORTE begrunnelser UTEN linjeskift.
Svar KUN med JSON:
{{"inntrykk": {{"startside": {{"score": 3, "begrunnelse": "setning"}}, "bilder_film": {{"score": 3, "begrunnelse": "setning"}}, "produktinfo": {{"score": 3, "begrunnelse": "setning"}}, "sokefunksjon": {{"score": 3, "begrunnelse": "setning"}}}}, "iks": {{"kjopsvilkar": {{"score": 3, "begrunnelse": "setning"}}, "levering": {{"score": 3, "begrunnelse": "setning"}}, "retur": {{"score": 3, "begrunnelse": "setning"}}, "selvbetjent_ks": {{"score": 3, "begrunnelse": "setning"}}, "betjent_ks": {{"score": 3, "begrunnelse": "setning"}}, "baerekraft_strategi": {{"score": 3, "begrunnelse": "setning"}}, "baerekraft_kunder": {{"score": 3, "begrunnelse": "setning"}}}}}}"""

    prompt2 = f"""Vurder nettbutikken "{navn}" på {url}.

KATEGORI 3 – KASSEN / MERSALG:
- leveringsvalg_kassen: Hentested, pakkeboks, hjemlevering, tidsvindu, dato?
- betaling: Kort, Vipps, Klarna, faktura, delbetaling?
- mersalg: Relevante anbefalinger, artikler, oppskrifter, videoer?

KATEGORI 4 – MARKEDSFØRING / KUNDEDIALOG (SoMe 40%, Kundeklubb 30%, Nyhetsbrev 30%):

Sosiale medier (40%):
- Engasjement og dialog med kundene? Autentisk innhold (ikke katalogbilder)?
- Troverdige influensere? Kjøpslenker? Kundeservice via SoMe?
- IKKE antall følgere – det er engasjement, autentisitet og kjøpsmulighet som teller.

Kundeklubb (30%) – VIKTIG:
- Har de kundeklubb? Score slik:
  2 = Ingen kundeklubb
  3 = Har kundeklubb men kun inngangsrabatt ved første kjøp
  4 = Har kundeklubb med poeng/rabatter/lojalitetsfordeler
  5 = Har kundeklubb med full lojalitetspakke (poeng, rabatter, fortrinn til salg, egne tilbud)

Nyhetsbrev (30%):
- Har de nyhetsbrev? Rekrutterer aktivt via kjøpsreisen?
- SMS + e-post? Pop-up, påmelding i kassen?

Kartlegg også:
- Logistikkpartnere: {", ".join(LOGISTIKK_AKTORER)}
- Tech: Mobiloptimalisert? SSL? Lastetid? Trustpilot?
- Åpenhetsloven: Rapport funnet for 2024/2025?

Score fra 1-5. KORTE begrunnelser UTEN linjeskift.
Svar KUN med JSON:
{{"kommentar": "kort vurdering", "kassen": {{"leveringsvalg_kassen": {{"score": 3, "begrunnelse": "setning"}}, "betaling": {{"score": 3, "begrunnelse": "setning"}}, "mersalg": {{"score": 3, "begrunnelse": "setning"}}}}, "markedsforing": {{"some": {{"score": 3, "begrunnelse": "setning", "har_kundeklubb": false}}, "kundeklubb": {{"score": 2, "begrunnelse": "setning", "har_kundeklubb": false}}, "nyhetsbrev": {{"score": 3, "begrunnelse": "setning"}}}}, "logistikk": {{"Posten/Bring": false, "PostNord": false, "Helthjem": false, "Instabox": false, "Porterbuddy": false, "DHL": false, "Budbee": false, "UPS/FedEx": false, "Egne biler": false}}, "tech": {{"mobil": "ok", "ssl": "ok", "lastetid": "ok", "trustpilot": "Ikke funnet"}}, "apenhetsloven": {{"palagt": false, "rapport_funnet": false, "kommentar": "setning"}}, "trust": ["element1"]}}"""

    def kall(prompt, nr):
        for forsok in range(2):
            try:
                resp = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1200,
                    tools=[{"type": "web_search_20250305", "name": "web_search"}],
                    messages=[{"role": "user", "content": prompt}]
                )
                tekst = ""
                for blokk in resp.content:
                    if blokk.type == "text":
                        tekst += blokk.text + "\n"
                data = parse_json(tekst)
                if data:
                    return data
                # Reparasjon
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
        resultat["markedsforing"] = del2.get("markedsforing", {})
        resultat["logistikk"] = del2.get("logistikk", {})
        resultat["tech"] = del2.get("tech", {})
        resultat["apenhetsloven"] = del2.get("apenhetsloven", {})
        resultat["trust"] = del2.get("trust", [])
    else:
        resultat.update({"kommentar": "", "kassen": {}, "markedsforing": {},
                         "logistikk": {}, "tech": {}, "apenhetsloven": {}, "trust": []})

    return resultat


# ─────────────────────────────────────────────
# BEREGN TOTALSCORE MED NY VEKTING
# ─────────────────────────────────────────────
def beregn_totalscore(scoring: dict) -> dict:
    # 1. Første inntrykk – 4 kriterier, 25% vekt hver
    inntrykk_data = scoring.get("inntrykk", {})
    inntrykk_scores = [v["score"] for v in inntrykk_data.values() if isinstance(v, dict) and "score" in v]
    inntrykk_snitt = round(sum(inntrykk_scores) / len(inntrykk_scores), 2) if inntrykk_scores else 0

    # 2. IKS med intern vekting
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

    # 3. Kassen – lik vekt
    kassen_data = scoring.get("kassen", {})
    kassen_scores = [v["score"] for v in kassen_data.values() if isinstance(v, dict) and "score" in v]
    kassen_snitt = round(sum(kassen_scores) / len(kassen_scores), 2) if kassen_scores else 0

    # 4. Markedsføring – SoMe 40%, Kundeklubb 30% (dynamisk), Nyhetsbrev 30%
    mf = scoring.get("markedsforing", {})
    some_score = mf.get("some", {}).get("score", 0)
    kundeklubb_score = mf.get("kundeklubb", {}).get("score", 0)
    nyhetsbrev_score = mf.get("nyhetsbrev", {}).get("score", 0)
    har_kundeklubb = mf.get("kundeklubb", {}).get("har_kundeklubb", kundeklubb_score > 2)

    if har_kundeklubb:
        # Normal vekting
        mf_snitt = round(some_score * 0.40 + kundeklubb_score * 0.30 + nyhetsbrev_score * 0.30, 2)
    else:
        # Dynamisk – kundeklubb hoppes over, vekt fordeles på SoMe og nyhetsbrev
        mf_snitt = round(some_score * (40/70) + nyhetsbrev_score * (30/70), 2)

    # Totalpoeng
    total = round(
        inntrykk_snitt * 0.25 +
        iks_snitt * 0.25 +
        kassen_snitt * 0.25 +
        mf_snitt * 0.25, 1
    )

    return {
        "total": total,
        "kategorier": {
            "inntrykk": round(inntrykk_snitt, 1),
            "iks": round(iks_snitt, 1),
            "iks_detalj": {
                "klr": round(klr_snitt, 1),
                "kundeservice": round(ks_snitt, 1),
                "baerekraft": round(bk_snitt, 1),
            },
            "kassen": round(kassen_snitt, 1),
            "markedsforing": round(mf_snitt, 1),
            "mf_detalj": {
                "some": some_score,
                "kundeklubb": kundeklubb_score,
                "nyhetsbrev": nyhetsbrev_score,
                "har_kundeklubb": har_kundeklubb,
            }
        }
    }


def bygg_scoring_for_visning(scoring_raw: dict) -> dict:
    if not scoring_raw:
        return None
    return {
        "inntrykk": [
            {"navn": "Startside og navigasjon", "score": scoring_raw.get("inntrykk", {}).get("startside", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("startside", {}).get("begrunnelse", ""), "vekt": "25% av inntrykk"},
            {"navn": "Bilder og film", "score": scoring_raw.get("inntrykk", {}).get("bilder_film", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("bilder_film", {}).get("begrunnelse", ""), "vekt": "25% av inntrykk"},
            {"navn": "Produktinformasjon", "score": scoring_raw.get("inntrykk", {}).get("produktinfo", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("produktinfo", {}).get("begrunnelse", ""), "vekt": "25% av inntrykk"},
            {"navn": "Søkefunksjon", "score": scoring_raw.get("inntrykk", {}).get("sokefunksjon", {}).get("score", 0), "begrunnelse": scoring_raw.get("inntrykk", {}).get("sokefunksjon", {}).get("begrunnelse", ""), "vekt": "25% av inntrykk"},
        ],
        "iks": [
            {"navn": "Kjøpsvilkår (lovkrav)", "score": scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("begrunnelse", ""), "vekt": "KLR 20%"},
            {"navn": "Leveringsinformasjon", "score": scoring_raw.get("iks", {}).get("levering", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("levering", {}).get("begrunnelse", ""), "vekt": "KLR 40%"},
            {"navn": "Returløsning", "score": scoring_raw.get("iks", {}).get("retur", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("retur", {}).get("begrunnelse", ""), "vekt": "KLR 40%"},
            {"navn": "Selvbetjent kundeservice", "score": scoring_raw.get("iks", {}).get("selvbetjent_ks", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("selvbetjent_ks", {}).get("begrunnelse", ""), "vekt": "KS 50%"},
            {"navn": "Betjent kundeservice", "score": scoring_raw.get("iks", {}).get("betjent_ks", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("betjent_ks", {}).get("begrunnelse", ""), "vekt": "KS 50%"},
            {"navn": "Bærekraft – strategi", "score": scoring_raw.get("iks", {}).get("baerekraft_strategi", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("baerekraft_strategi", {}).get("begrunnelse", ""), "vekt": "BK 50%"},
            {"navn": "Bærekraft – kundeverktøy", "score": scoring_raw.get("iks", {}).get("baerekraft_kunder", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("baerekraft_kunder", {}).get("begrunnelse", ""), "vekt": "BK 50%"},
        ],
        "kassen": [
            {"navn": "Leveringsvalg i kassen", "score": scoring_raw.get("kassen", {}).get("leveringsvalg_kassen", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("leveringsvalg_kassen", {}).get("begrunnelse", "")},
            {"navn": "Betalingsalternativer", "score": scoring_raw.get("kassen", {}).get("betaling", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("betaling", {}).get("begrunnelse", "")},
            {"navn": "Mersalg og inspirasjon", "score": scoring_raw.get("kassen", {}).get("mersalg", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("mersalg", {}).get("begrunnelse", "")},
        ],
        "markedsforing": [
            {"navn": "Sosiale medier", "score": scoring_raw.get("markedsforing", {}).get("some", {}).get("score", 0), "begrunnelse": scoring_raw.get("markedsforing", {}).get("some", {}).get("begrunnelse", ""), "vekt": "SoMe 40%"},
            {"navn": "Kundeklubb", "score": scoring_raw.get("markedsforing", {}).get("kundeklubb", {}).get("score", 0), "begrunnelse": scoring_raw.get("markedsforing", {}).get("kundeklubb", {}).get("begrunnelse", ""), "vekt": "Kundeklubb 30% (dynamisk)"},
            {"navn": "Nyhetsbrev", "score": scoring_raw.get("markedsforing", {}).get("nyhetsbrev", {}).get("score", 0), "begrunnelse": scoring_raw.get("markedsforing", {}).get("nyhetsbrev", {}).get("begrunnelse", ""), "vekt": "Nyhetsbrev 30%"},
        ],
    }


# ─────────────────────────────────────────────
# HOVED-AGENT FUNKSJON
# ─────────────────────────────────────────────
def kjor_agent(filbane, api_key: str, maks: int = None, fremgang_callback=None) -> dict:
    """
    Kjører screeningen og returnerer resultater som dict.
    fremgang_callback(melding, prosent) kalles løpende for fremdriftsvisning.
    """
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
        time.sleep(3)

        # B – Brønnøysund
        brreg = sok_brreg(navn, selskapsnavn)

        # C – Regnskapstall
        regnskap = {}
        if brreg.get("orgnr"):
            regnskap = hent_regnskap(brreg["orgnr"])

        # Klassifisering og KO
        klasse, omsetning_tekst = klassifiser(regnskap, info)
        skal_ut, aarsak = ko_sjekk(brreg, info)
        advarsel = b2c_og_skandinav_advarsel(info)

        butikk = {
            "name": navn,
            "url": info.get("url", ""),
            "orgform": brreg.get("orgform", "Ukjent"),
            "orgnr": brreg.get("orgnr", ""),
            "enk": er_enk(brreg.get("orgform", "")) and not brreg.get("usikker", False),
            "bransje": info.get("bransje", "Annet"),
            "omsetning": omsetning_tekst,
            "klasse": klasse,
            "b2cAdvarsel": advarsel,
            "brregUsikker": brreg.get("usikker", False),
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
            "kassen": None,
            "markedsforing": None,
            "mfDetalj": None,
        }

        if skal_ut:
            butikk["status"] = "ut"
            butikk["screeningBegrunnelse"] = f"Filtrert ut: {aarsak}"
            butikk["kommentar"] = "Filtrert ut i screening."
        else:
            brreg_note = " | Brreg-treff usikkert" if brreg.get("usikker") else ""
            butikk["screeningBegrunnelse"] = (
                f"{brreg.get('orgform','?')}, {info.get('land','?')}. "
                f"Klasse {klasse}. {omsetning_tekst}.{brreg_note}"
            )
            if advarsel:
                butikk["screeningBegrunnelse"] += f" | {advarsel}"

            if fremgang_callback:
                fremgang_callback(f"[{i+1}/{totalt}] Scorer: {navn}", prosent)

            scoring_raw = score_nettbutikk(navn, info.get("url", ""), klasse, client)

            if scoring_raw:
                totaler = beregn_totalscore(scoring_raw)
                butikk["status"] = "inn"
                butikk["total"] = totaler["total"]
                butikk["inntrykk"] = totaler["kategorier"]["inntrykk"]
                butikk["iks"] = totaler["kategorier"]["iks"]
                butikk["iksDetalj"] = totaler["kategorier"].get("iks_detalj")
                butikk["kassen"] = totaler["kategorier"]["kassen"]
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
        fremgang_callback("Ferdig!", 100)

    return resultater
