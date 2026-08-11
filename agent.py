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
# SCORINGSMODELL – OPPDATERT VEKTING (v6)
# ─────────────────────────────────────────────
#
# KATEGORI 1: Første inntrykk (25% av total)
#   - Startside (kun meny/overordnet, IKKE søk): 25%
#   - Bilder og film (video er bonus, ikke krav): 25%
#   - Produktinformasjon (IKKE anmeldelser): 25%
#   - Søkefunksjon: 25%
#
# KATEGORI 2: Info, kundeservice og bærekraft (25% av total)
#   - Kjøpsvilkår, levering og retur: 40%
#     - Kjøpsvilkår: 20% (JA/NEI-portvokter – nei = filtreres helt ut av videre vurdering)
#     - Levering: 40% (ingen antakelser om kronegrenser)
#     - Retur: 40%
#   - Kundeservice: 40% (samlet kategori, ikke lenger delt selvbetjent/betjent)
#   - Bærekraft: 20% (samlet kategori)
#
# KATEGORI 3: Kassen, mersalg og inspirasjon (25% av total)
#   - Kassen: 50% (innlogging fjernet – 4 gjenværende kriterier)
#     - Leveringsvalg og fleksibilitet: 35%
#     - Leveringspris: 25%
#     - Leveringstid og presisjon: 20%
#     - Betalingsalternativer: 20%
#   - Mersalg: 25%
#   - Inspirasjon: 25%
#
# KATEGORI 4: Markedsføring og kundedialog (25% av total)
#   - Sosiale medier: 100% av kat4-poengsummen
#   - Kundeklubb: IKKE scoret – kun ja/nei + beskrivelse
#   - Nyhetsbrev/SMS: IKKE scoret – kun ja/nei + beskrivelse
#
# NB: Stor-klassen bedømmes strengere enn Medium og Liten.
# NB: Mobilvennlighet/mobiloptimalisering er fjernet som eget vurderingspunkt.

IKS_VEKTER = {
    "kjopsvilkar_levering_retur": 0.40,
    "kundeservice": 0.40,
    "baerekraft": 0.20,
}

KLR_VEKTER = {
    "kjopsvilkar": 0.20,
    "levering": 0.40,
    "retur": 0.40,
}

# innlogging fjernet – vektet på nytt blant de 4 gjenværende Kassen-kriteriene
KASSEN_VEKTER = {
    "leveringsvalg": 0.35,
    "leveringspris": 0.25,
    "leveringstid": 0.20,
    "betaling": 0.20,
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
                tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 4}],
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
def score_nettbutikk(navn: str, url: str, klasse: str, client, sidedata: dict = None) -> dict:
    """
    Scorer i to separate kall:
    Kall 1: Første inntrykk + IKS
    Kall 2: Kassen/mersalg/inspirasjon + Markedsføring + Logistikk

    sidedata (valgfritt): faktisk innhentet sidetekst fra en sitekontroll-modul
    (f.eks. søkeresultat, kjøpsvilkår-tekst, fraktinfo-tekst) – brukes som fasit
    i prompten i stedet for at Claude må gjette ut fra websøk alene.
    """
    strengere = "NB: Klasse Stor bedømmes strengere og med høyere krav enn Medium og Liten. " if klasse == "Stor" else ""

    sidedata_tekst = ""
    if sidedata:
        sidedata_tekst = f"""

FAKTISK INNHENTET SIDEINNHOLD (bruk dette som fasit – ikke gjett hvis det står her):
{json.dumps(sidedata, ensure_ascii=False, indent=None)[:4000]}
"""

    prompt1 = f"""Vurder nettbutikken "{navn}" på URL: {url}
Klasse: {klasse}. {strengere}
Score fra 1-5. KORTE begrunnelser UTEN linjeskift.
VIKTIG: Ikke gjett eller anta noe som ikke faktisk er observert eller oppgitt. Hvis informasjon mangler,
si det tydelig i begrunnelsen ("ikke funnet") i stedet for å anta et sannsynlig svar.
{sidedata_tekst}

KATEGORI 1 – FØRSTE INNTRYKK (25% av total, 4 kriterier à 25%):

startside: Vurder KUN menylinjen og det overordnede visuelle inntrykket på startsiden – IKKE søkefunksjonen
(den vurderes separat). Er menystrukturen logisk? Vises kategorier også som bilder/visuelle snarveier
(ikke bare tekst i menyen)? Visuelt ryddig? Lett å forstå hva de selger? Identitet og inspirasjon
(ikke bare bestselgere)?

bilder_film: Vurder KUN selve bildebruken på hjemmesiden, og hva som er riktig for akkurat DENNE bransjen
– IKKE innhold hentet fra sosiale medier. Er bildene profesjonelle og relevante for bransjen (produkter i
miljø der det passer, flere vinkler der det passer)? Film/video er en BONUS, ikke et krav, og relevansen
varierer med bransje (f.eks. sminkebruk-video hos en hudpleiebutikk, brette-video ved servise/bestikk) –
fravær av video skal IKKE trekke ned, kun tilstedeværelse gir pluss.

produktinfo: Vurder i hvor stor grad selve produktet er beskrevet – materiale, størrelse, mål, vekt,
bruksområde og lignende relevante detaljer der det passer for produkttypen. IKKE vurder om det finnes
kjøper-/produktanmeldelser – det er ikke en del av dette kriteriet.

sokefunksjon: Basert på faktisk gjennomført søk (se sideinnhold over hvis tilgjengelig) – gir søket riktig
resultat både med ett søkeord (f.eks. "genser") og med to parametre samtidig (f.eks. "rød genser")?
Hvis faktisk søkeresultat IKKE er tilgjengelig i sideinnholdet under, si eksplisitt i begrunnelsen at dette
er en antatt vurdering basert på generell informasjon om nettbutikken, og ikke et testet søkeresultat.

KATEGORI 2 – INFO, KUNDESERVICE OG BÆREKRAFT (25% av total):

Kjøpsvilkår, levering og retur (40% av kat2):
kjopsvilkar: JA/NEI – har nettbutikken kjøpsvilkår tilgjengelig? Dette er lovpålagt. Score 5 hvis ja
(nøytralt, forventet standard), score 1 hvis nei (MEGET negativt – dette er en portvokter-sjekk, se eget felt
"har_kjopsvilkar" i JSON-svaret).

levering: IKKE gjett. Vurder KUN basert på informasjon nettbutikken faktisk oppgir UTENFOR selve
checkout-løsningen – dvs. på produktside, i FAQ, i footer, eller på egen fraktinfo-side. Er leveringspris
oppgitt et sted (ikke i selve kassen)? Brukes "FRI FRAKT" som tydelig USP allerede før kunden legger noe i
handlekurven? Er det oppgitt cut-off-tid (f.eks. "bestilt før kl. 12 sendes samme dag") – dette er en bonus
som viser god logistikk-kontroll, IKKE et krav. IKKE ta med "pakketid" som eget punkt – det er ikke relevant.
IKKE anta noen konkret kronegrense for hva som er "for høyt" eller "for lavt" beløp for fri frakt – vurder
kun ut fra hva som faktisk står, uten egne antakelser om hva som er rimelig.

retur: Undersøk om nettbutikken tilbyr returløsning med FERDIG ADRESSELAPP i pakken, eller om kunden må
skrive ut adresselappen selv ved retur. God score krever at nettbutikken sørger for adresseetikett til en
fri eller akseptabel pris (under 100 kr). Hvis kunden selv må ordne alt, eller returkostnaden er høy/uklar,
skal dette trekke ned.

Kundeservice (40% av kat2) – ÉN samlet kategori (IKKE del i selvbetjent/betjent):
kundeservice: List opp hvilke kontaktkanaler nettbutikken faktisk tilbyr (telefon, e-post, chat, åpningstider
osv.) og vurder hvor lett tilgjengelig kundeservice er. Vurder om en eventuell chatbot er AI-drevet eller
bemannet/vanlig chat, HVIS dette faktisk kan fastslås ut fra tilgjengelig informasjon – hvis det ikke lar
seg avgjøre med sikkerhet, list heller opp kanalene uten å gjette på AI vs. bemannet.

Bærekraft (20% av kat2) – ÉN samlet kategori:
baerekraft: Har nettbutikken informasjon om bærekraftsrapporter, klimamål, Svanemerket, Åpenhetsloven,
sertifiseringer (f.eks. GOTS) eller lignende? Nevner de noe om at enkelte produkter er bærekraftige, og
i så fall hva sier de konkret? Små nettbutikker forventes IKKE nødvendigvis å ha dette – fravær skal ikke
gi like hardt trekk som for store aktører. Oppsummer i begrunnelsen konkret HVA de eventuelt nevner.

Svar KUN med JSON:
{{"inntrykk": {{"startside": {{"score": 3, "begrunnelse": "setning"}}, "bilder_film": {{"score": 3, "begrunnelse": "setning"}}, "produktinfo": {{"score": 3, "begrunnelse": "setning"}}, "sokefunksjon": {{"score": 3, "begrunnelse": "setning"}}}}, "iks": {{"kjopsvilkar": {{"score": 3, "begrunnelse": "setning", "har_kjopsvilkar": true}}, "levering": {{"score": 3, "begrunnelse": "setning"}}, "retur": {{"score": 3, "begrunnelse": "setning"}}, "kundeservice": {{"score": 3, "begrunnelse": "setning", "kanaler": ["telefon", "e-post", "chat"]}}, "baerekraft": {{"score": 3, "begrunnelse": "setning"}}}}}}"""

    prompt2 = f"""Vurder nettbutikken "{navn}" på URL: {url}
Klasse: {klasse}. {strengere}
Score fra 1-5. KORTE begrunnelser UTEN linjeskift.
VIKTIG: Ikke gjett eller anta noe som ikke faktisk er observert eller oppgitt. Hvis informasjon mangler,
si det tydelig i begrunnelsen i stedet for å anta et sannsynlig svar.
{sidedata_tekst}

KATEGORI 3 – KASSEN, MERSALG OG INSPIRASJON (25% av total):
Kassen teller 50%, Mersalg 25%, Inspirasjon 25%.

Kassen (50% av kat3) – 4 underkriterier. Disse skal IKKE gjettes ut fra generell tekst på nettbutikken –
vurder KUN basert på faktisk observert checkout-flyt hvis dette er tilgjengelig i sideinnholdet over.
Hvis ikke, si eksplisitt i begrunnelsen at scoren er en antatt vurdering og IKKE en testet observasjon:

leveringsvalg: Har de minst 2 leveringsalternativer i kassen? Nedtrekksmeny for valg av hentested/pakkeboks?
  (Nedtrekksmeny = høyere score enn forhåndsvalgt/mange separate valg som krever scrolling)
  Nærhet til kunden viktig – mange utleveringssteder/pakkebokser gir høyere score.

leveringspris: Hva er billigste leveringspris faktisk vist i kassen? Under 99 kr = god score, over 129 kr
  = trekk (spesielt hvis varene ikke er store/tunge/eksklusivt sortiment).

leveringstid: Oppgir de estimert leveringsdato i kassen = best. Tidsintervall (1-3 dager) = middels.
  Ingen info eller kun transporttid = dårlig.

betaling: Hvilke betalingsmetoder ble faktisk tilbudt i kassen? Full score hvis alle vanlige:
  kredit/debetkort, Vipps, Klarna/etterbetaling/delbetaling.

Mersalg (25% av kat3):
mersalg: Sjekk om nettbutikken viser relevante produktanbefalinger BÅDE på produktsiden OG i selve
  checkout-løsningen (hvis observert). Er anbefalingene relevante for det kunden faktisk ser på/kjøper?

Inspirasjon (25% av kat3):
inspirasjon: Finnes det guider knyttet til produktet – artikler, filmer, oppskrifter eller
  bruksanvisninger som viser hvordan produktet kan brukes? Gir faglig tillit OG selger mer?
  (Eks: oppskrifter med kjøkkenutstyr, friluftsliv-filmer med utstyr, bruksguider for klær/utstyr)

KATEGORI 4 – MARKEDSFØRING OG KUNDEDIALOG (25% av total):
Kun "some" gir tallscore i denne kategorien (100% av kat4-scoren). Kundeklubb og nyhetsbrev vurderes
som ja/nei + beskrivelse, IKKE med poengsum (se JSON-format).

some: Hvilke kanaler/plattformer er nettbutikken faktisk til stede i? Vurder innholdet de faktisk
  publiserer der (engasjement, autentisitet – IKKE antall følgere). Sjekk om det er mulig å handle
  direkte via SoMe (f.eks. "Handle her"-knapp på Instagram/TikTok-innlegg/annonser) – oppgi ja/nei/ukjent
  hvis dette faktisk kan fastslås. IKKE trekk ned for manglende influencer-samarbeid – tilstedeværelse av
  troverdige influencere er KUN en bonus, aldri et minus ved fravær.

kundeklubb: Har nettbutikken en kundeklubb? Svar KUN ja eller nei (se "har_kundeklubb" i JSON).
  Hvis ja: beskriv KORT hva kundeklubben faktisk inneholder – poeng, rabatter, fortrinn til salg osv.
  IKKE reduser dette til kun "inngangsrabatt ved første kjøp" – beskriv det reelle innholdet.
  Dette gis IKKE en poengsum, kun beskrivelse.

nyhetsbrev: Er det mulig å melde seg på nyhetsbrev og/eller SMS (f.eks. pop-up-vindu, felt i footer)?
  Svar KUN ja eller nei (se "har_nyhetsbrev" i JSON). Beskriv KORT hva nyhetsbrevet/SMS-varslene dreier
  seg om hvis det fremgår. Dette gis IKKE en poengsum, kun beskrivelse.

Kartlegg også (ikke score):
logistikk: Hvilke brukes? {", ".join(LOGISTIKK_AKTORER)}
tech: SSL? Lastetid (ok/warn/bad)? Trustpilot-score? (IKKE vurder mobilvennlighet/mobiloptimalisering –
  dette skal ikke lenger testes eller inngå i vurderingen)
trust: Liste med tillitselementer (garantier, sertifiseringer, kundeomtaler, fysisk adresse synlig osv)
apenhetsloven: Er rapport funnet for 2024/2025? Er de rapporteringspliktige?

Svar KUN med JSON:
{{"kommentar": "2-3 setninger om butikken totalt", "kassen": {{"leveringsvalg": {{"score": 3, "begrunnelse": "setning"}}, "leveringspris": {{"score": 3, "begrunnelse": "setning"}}, "leveringstid": {{"score": 3, "begrunnelse": "setning"}}, "betaling": {{"score": 3, "begrunnelse": "setning"}}}}, "mersalg": {{"score": 3, "begrunnelse": "setning"}}, "inspirasjon": {{"score": 3, "begrunnelse": "setning"}}, "markedsforing": {{"some": {{"score": 3, "begrunnelse": "setning", "handle_direkte_i_some": "ja/nei/ukjent"}}, "kundeklubb": {{"har_kundeklubb": false, "beskrivelse": "setning"}}, "nyhetsbrev": {{"har_nyhetsbrev": false, "beskrivelse": "setning"}}}}, "logistikk": {{"Posten/Bring": false, "PostNord": false, "Helthjem": false, "Instabox": false, "Porterbuddy": false, "DHL": false, "Budbee": false, "UPS/FedEx": false, "Egne biler": false}}, "tech": {{"ssl": "ok", "lastetid": "ok", "trustpilot": "Ikke funnet"}}, "trust": ["element1", "element2"], "apenhetsloven": {{"palagt": false, "rapport_funnet": false, "kommentar": "setning"}}}}"""

    def kall(prompt, nr):
        for forsok in range(2):
            try:
                resp = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1200,
                    tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 4}],
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

    # 2. IKS – kjøpsvilkår/levering/retur 40%, kundeservice 40%, bærekraft 20%
    iks = scoring.get("iks", {})
    klr_snitt = (
        iks.get("kjopsvilkar", {}).get("score", 0) * KLR_VEKTER["kjopsvilkar"] +
        iks.get("levering", {}).get("score", 0) * KLR_VEKTER["levering"] +
        iks.get("retur", {}).get("score", 0) * KLR_VEKTER["retur"]
    )
    ks_snitt = iks.get("kundeservice", {}).get("score", 0)
    bk_snitt = iks.get("baerekraft", {}).get("score", 0)
    iks_snitt = round(
        klr_snitt * IKS_VEKTER["kjopsvilkar_levering_retur"] +
        ks_snitt * IKS_VEKTER["kundeservice"] +
        bk_snitt * IKS_VEKTER["baerekraft"], 2
    )

    # 3. Kassen (50%, 4 underkriterier – innlogging fjernet) + Mersalg (25%) + Inspirasjon (25%)
    kassen = scoring.get("kassen", {})
    kassescore = (
        kassen.get("leveringsvalg", {}).get("score", 0) * KASSEN_VEKTER["leveringsvalg"] +
        kassen.get("leveringspris", {}).get("score", 0) * KASSEN_VEKTER["leveringspris"] +
        kassen.get("leveringstid", {}).get("score", 0) * KASSEN_VEKTER["leveringstid"] +
        kassen.get("betaling", {}).get("score", 0) * KASSEN_VEKTER["betaling"]
    )
    mersalg_score = scoring.get("mersalg", {}).get("score", 0)
    insp_score = scoring.get("inspirasjon", {}).get("score", 0)
    kat3_snitt = round(kassescore * 0.50 + mersalg_score * 0.25 + insp_score * 0.25, 2)

    # 4. Markedsføring – KUN "some" gir poengsum. Kundeklubb/nyhetsbrev er ja/nei + beskrivelse (informativt).
    mf = scoring.get("markedsforing", {})
    some_score = mf.get("some", {}).get("score", 0)
    har_kk = mf.get("kundeklubb", {}).get("har_kundeklubb", False)
    har_nb = mf.get("nyhetsbrev", {}).get("har_nyhetsbrev", False)
    mf_snitt = round(some_score, 2)

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
                "har_kundeklubb": har_kk,
                "har_nyhetsbrev": har_nb,
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
            {"navn": "Kjøpsvilkår (lovkrav)", "score": scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("begrunnelse", ""), "vekt": "KLR 20%", "har_kjopsvilkar": scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("har_kjopsvilkar")},
            {"navn": "Leveringsinformasjon", "score": scoring_raw.get("iks", {}).get("levering", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("levering", {}).get("begrunnelse", ""), "vekt": "KLR 40%"},
            {"navn": "Returløsning", "score": scoring_raw.get("iks", {}).get("retur", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("retur", {}).get("begrunnelse", ""), "vekt": "KLR 40%"},
            {"navn": "Kundeservice", "score": scoring_raw.get("iks", {}).get("kundeservice", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("kundeservice", {}).get("begrunnelse", ""), "vekt": "KS 40%", "kanaler": scoring_raw.get("iks", {}).get("kundeservice", {}).get("kanaler", [])},
            {"navn": "Bærekraft", "score": scoring_raw.get("iks", {}).get("baerekraft", {}).get("score", 0), "begrunnelse": scoring_raw.get("iks", {}).get("baerekraft", {}).get("begrunnelse", ""), "vekt": "BK 20%"},
        ],
        "kassen": [
            {"navn": "Leveringsvalg og fleksibilitet", "score": scoring_raw.get("kassen", {}).get("leveringsvalg", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("leveringsvalg", {}).get("begrunnelse", ""), "vekt": "Kassen 35%"},
            {"navn": "Leveringspris", "score": scoring_raw.get("kassen", {}).get("leveringspris", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("leveringspris", {}).get("begrunnelse", ""), "vekt": "Kassen 25%"},
            {"navn": "Leveringstid og presisjon", "score": scoring_raw.get("kassen", {}).get("leveringstid", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("leveringstid", {}).get("begrunnelse", ""), "vekt": "Kassen 20%"},
            {"navn": "Betalingsalternativer", "score": scoring_raw.get("kassen", {}).get("betaling", {}).get("score", 0), "begrunnelse": scoring_raw.get("kassen", {}).get("betaling", {}).get("begrunnelse", ""), "vekt": "Kassen 20%"},
            {"navn": "Mersalg og produktanbefalinger", "score": scoring_raw.get("mersalg", {}).get("score", 0), "begrunnelse": scoring_raw.get("mersalg", {}).get("begrunnelse", ""), "vekt": "Mersalg 25%"},
            {"navn": "Inspirasjon og innhold", "score": scoring_raw.get("inspirasjon", {}).get("score", 0), "begrunnelse": scoring_raw.get("inspirasjon", {}).get("begrunnelse", ""), "vekt": "Inspirasjon 25%"},
        ],
        "markedsforing": [
            {"navn": "Sosiale medier", "score": scoring_raw.get("markedsforing", {}).get("some", {}).get("score", 0), "begrunnelse": scoring_raw.get("markedsforing", {}).get("some", {}).get("begrunnelse", ""), "vekt": "100%", "handle_direkte_i_some": scoring_raw.get("markedsforing", {}).get("some", {}).get("handle_direkte_i_some")},
            {"navn": "Kundeklubb", "har_kundeklubb": scoring_raw.get("markedsforing", {}).get("kundeklubb", {}).get("har_kundeklubb", False), "begrunnelse": scoring_raw.get("markedsforing", {}).get("kundeklubb", {}).get("beskrivelse", ""), "vekt": "Ikke scoret – informativt"},
            {"navn": "Nyhetsbrev og SMS", "har_nyhetsbrev": scoring_raw.get("markedsforing", {}).get("nyhetsbrev", {}).get("har_nyhetsbrev", False), "begrunnelse": scoring_raw.get("markedsforing", {}).get("nyhetsbrev", {}).get("beskrivelse", ""), "vekt": "Ikke scoret – informativt"},
        ],
    }


# ─────────────────────────────────────────────
# HOVED-AGENT
# ─────────────────────────────────────────────
def kjor_agent(filbane, api_key: str, maks: int = None, fremgang_callback=None, eksisterende: dict = None) -> dict:
    """
    eksisterende (valgfritt): resultater-dict fra en tidligere kjøring (f.eks. lastet
    fra en gammel resultater.json). Butikker som allerede finnes der med status
    inn/ut/usikker blir IKKE scoret på nytt – sparer API-kostnad ved gjenopptak.
    """
    client = anthropic.Anthropic(api_key=api_key)

    raa_navn = les_excel(filbane)
    if not raa_navn:
        return {}

    navn_liste = fjern_duplikater(raa_navn)
    if maks:
        navn_liste = navn_liste[:maks]

    resultater = dict(eksisterende) if eksisterende else {}
    totalt = len(navn_liste)

    for i, navn in enumerate(navn_liste):
        prosent = int((i / totalt) * 100)

        if navn in resultater and resultater[navn].get("status") in ("inn", "ut", "usikker"):
            if fremgang_callback:
                fremgang_callback(f"[{i+1}/{totalt}] ⏭ Hopper over (allerede vurdert): {navn}", prosent)
            continue

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
                har_kjopsvilkar = scoring_raw.get("iks", {}).get("kjopsvilkar", {}).get("har_kjopsvilkar", True)
                if har_kjopsvilkar is False:
                    butikk["status"] = "ut"
                    butikk["screeningBegrunnelse"] += " | Filtrert ut: fant ingen kjøpsvilkår på nettbutikken."
                    butikk["kommentar"] = "Filtrert ut – ingen kjøpsvilkår funnet (lovpålagt, portvokterkriterium)."
                    butikk["scoring"] = bygg_scoring_for_visning(scoring_raw)
                    resultater[navn] = butikk
                    if fremgang_callback:
                        fremgang_callback(f"[{i+1}/{totalt}] ✓ {navn} ferdig", prosent)
                    continue

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
