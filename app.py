"""
NETTHANDELSPRISEN – POSTEN BRING
Streamlit-app for screening og juryverktøy
"""

import streamlit as st
import json
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from agent import kjor_agent, LOGISTIKK_AKTORER

# ─────────────────────────────────────────────
# SIDEKONFIGURASJON
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Netthandelsprisen – Posten Bring",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Posten Bring farger
st.markdown("""
<style>
:root { --red: #C8102E; --coal: #212121; }
.stApp { background-color: #F5F4F2; }
.main-header {
    background: #212121;
    color: white;
    padding: 16px 24px;
    border-radius: 8px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 16px;
}
.logo-badge {
    background: #C8102E;
    color: white;
    padding: 4px 10px;
    border-radius: 4px;
    font-weight: 700;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.score-card {
    background: white;
    border-radius: 8px;
    padding: 16px;
    border-top: 3px solid #C8102E;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}
.metric-label {
    font-size: 11px;
    color: #999;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    font-weight: 600;
}
.metric-value {
    font-size: 28px;
    font-weight: 700;
    color: #1A1A1A;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <span class="logo-badge">Posten Bring</span>
    <div>
        <div style="font-size: 18px; font-weight: 700;">Netthandelsprisen – Jurysystem</div>
        <div style="font-size: 13px; color: #aaa;">Automatisk screening og scoringsagent</div>
    </div>
    <div style="margin-left: auto; background: #2a2a2a; color: #888; padding: 3px 10px; border-radius: 4px; font-size: 12px;">2026</div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "resultater" not in st.session_state:
    st.session_state.resultater = {}
if "agent_kjoert" not in st.session_state:
    st.session_state.agent_kjoert = False


# ─────────────────────────────────────────────
# GOOGLE SHEETS KOBLING
# ─────────────────────────────────────────────
def koble_sheets():
    """Kobler til Google Sheets for lagring av juryvurderinger."""
    try:
        creds_info = st.secrets["gcp_service_account"]
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        gc = gspread.authorize(creds)
        sheet_id = st.secrets["google_sheets"]["sheet_id"]
        sh = gc.open_by_key(sheet_id)
        return sh.sheet1
    except Exception as e:
        return None


def lagre_juryvurdering(butikk_navn: str, score: int, status: str, notat: str):
    """Lagrer juryens vurdering til Google Sheets."""
    ws = koble_sheets()
    if ws is None:
        return False
    try:
        data = ws.get_all_records()
        # Sjekk om butikken allerede finnes
        for i, rad in enumerate(data, 2):
            if rad.get("Navn") == butikk_navn:
                ws.update(f"D{i}", [[score]])
                ws.update(f"E{i}", [[status]])
                ws.update(f"F{i}", [[notat]])
                return True
        # Legg til ny rad
        ws.append_row([butikk_navn, "", "", score, status, notat])
        return True
    except Exception:
        return False


def hent_juryvurderinger() -> dict:
    """Henter alle juryvurderinger fra Google Sheets."""
    ws = koble_sheets()
    if ws is None:
        return {}
    try:
        data = ws.get_all_records()
        return {rad["Navn"]: rad for rad in data if rad.get("Navn")}
    except Exception:
        return {}


# ─────────────────────────────────────────────
# HJELPEFUNKSJONER
# ─────────────────────────────────────────────
def score_badge(score):
    if score is None:
        return "–"
    n = float(score)
    if n >= 4.5:
        color = "#1B6B3A"
        bg = "#E6F4EC"
    elif n >= 3.5:
        color = "#0D4A8A"
        bg = "#E6F0FA"
    elif n >= 2.5:
        color = "#7A4800"
        bg = "#FEF3E2"
    else:
        color = "#C8102E"
        bg = "#FDECEA"
    return f'<span style="background:{bg};color:{color};padding:2px 8px;border-radius:4px;font-weight:700;font-size:13px">{n:.1f}</span>'


def klasse_badge(klasse):
    farger = {
        "Liten": ("#0D4A8A", "#E6F0FA"),
        "Medium": ("#7A4800", "#FEF3E2"),
        "Stor": ("#5B2D8E", "#F0E8FA"),
    }
    c, bg = farger.get(klasse, ("#666", "#eee"))
    return f'<span style="background:{bg};color:{c};padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">{klasse}</span>'


def status_badge(status, enk=False):
    if enk:
        return '<span style="background:#FDECEA;color:#C8102E;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">⛔ ENK</span>'
    if status == "inn":
        return '<span style="background:#E6F4EC;color:#1B6B3A;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">✓ Videre</span>'
    if status == "ut":
        return '<span style="background:#FDECEA;color:#C8102E;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">✕ Ut</span>'
    return '<span style="background:#FEF3E2;color:#7A4800;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">? Sjekk</span>'


# ─────────────────────────────────────────────
# SIDEBAR – NAVIGASJON
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏆 Netthandelsprisen")
    st.markdown("---")

    side = st.radio(
        "Naviger til",
        ["📂 Last opp og kjør", "📋 Screening", "⭐ Topp 100", "🏆 Finale", "📦 Logistikk", "ℹ️ Om verktøyet"],
        label_visibility="collapsed"
    )

    st.markdown("---")

    if st.session_state.resultater:
        r = st.session_state.resultater
        inn = sum(1 for v in r.values() if v.get("status") == "inn" and not v.get("enk"))
        ut = sum(1 for v in r.values() if v.get("status") == "ut" and not v.get("enk"))
        enk = sum(1 for v in r.values() if v.get("enk"))

        st.markdown("**Statistikk**")
        st.metric("Totalt", len(r))
        st.metric("Går videre", inn)
        st.metric("Filtrert ut", ut)
        st.metric("ENK", enk)

    st.markdown("---")
    api_key = st.text_input("API-nøkkel", type="password", help="Din Anthropic API-nøkkel")


# ─────────────────────────────────────────────
# SIDE 1 – LAST OPP OG KJØR
# ─────────────────────────────────────────────
if side == "📂 Last opp og kjør":
    st.header("Last opp Excel-fil og start screening")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("""
        **Slik fungerer det:**
        1. Last opp Excel-filen med nettbutikknavn
        2. Velg antall butikker (10 for test, alle for produksjon)
        3. Trykk "Start screening"
        4. Agenten screener og scorer alle butikkene automatisk
        5. Resultatene vises i Screening- og Topp 100-fanene
        """)

        fil = st.file_uploader(
            "Last opp Excel-fil (.xlsx)",
            type=["xlsx"],
            help="Filen skal ha butikknavn i én kolonne"
        )

        maks = st.selectbox(
            "Antall butikker å behandle",
            [10, 25, 50, 100, 250, 500, "Alle"],
            index=0,
            help="Start med 10 for testing"
        )

        if not api_key:
            st.warning("⚠️ Lim inn API-nøkkelen din i sidepanelet til venstre")

        start_knapp = st.button(
            "🚀 Start screening",
            type="primary",
            disabled=not fil or not api_key
        )

    with col2:
        st.markdown("**Estimert kostnad**")
        kostnader = {10: "$0.50", 25: "$1.25", 50: "$2.50", 100: "$5.00", 250: "$12.50", 500: "$25.00", "Alle": "$40-50"}
        tider = {10: "3-5 min", 25: "8-12 min", 50: "15-20 min", 100: "30-40 min", 250: "1.5-2 t", 500: "3-4 t", "Alle": "6-8 t"}
        st.info(f"""
        **Butikker:** {maks}
        **Estimert kostnad:** {kostnader.get(maks, "?")}
        **Estimert tid:** {tider.get(maks, "?")}
        """)

        if st.session_state.agent_kjoert:
            st.success("✅ Siste kjøring fullført!")

    if start_knapp and fil and api_key:
        maks_int = None if maks == "Alle" else int(maks)

        st.markdown("---")
        st.markdown("### Agenten kjører...")

        fremgang_tekst = st.empty()
        fremgang_bar = st.progress(0)
        logg = st.empty()

        logg_linjer = []

        def oppdater_fremgang(melding, prosent):
            fremgang_tekst.markdown(f"**{melding}**")
            fremgang_bar.progress(prosent / 100)
            logg_linjer.append(melding)
            if len(logg_linjer) > 10:
                logg_linjer.pop(0)
            logg.code("\n".join(logg_linjer))

        try:
            resultater = kjor_agent(
                filbane=fil,
                api_key=api_key,
                maks=maks_int,
                fremgang_callback=oppdater_fremgang
            )
            st.session_state.resultater = resultater
            st.session_state.agent_kjoert = True
            fremgang_bar.progress(1.0)
            st.success(f"✅ Ferdig! {len(resultater)} butikker behandlet.")
            st.balloons()
        except Exception as e:
            st.error(f"❌ Feil: {e}")


# ─────────────────────────────────────────────
# SIDE 2 – SCREENING
# ─────────────────────────────────────────────
elif side == "📋 Screening":
    st.header("Screening – alle butikker")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå. Gå til 'Last opp og kjør' for å starte.")
        st.stop()

    r = st.session_state.resultater
    alle = list(r.values())

    # Filtre
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        status_filter = st.selectbox("Status", ["Alle", "Går videre", "Filtrert ut", "Krever sjekk", "ENK"])
    with col2:
        klasse_filter = st.selectbox("Klasse", ["Alle", "Liten", "Medium", "Stor"])
    with col3:
        bransjer = ["Alle"] + sorted(set(b.get("bransje", "Annet") for b in alle if b.get("bransje")))
        bransje_filter = st.selectbox("Bransje", bransjer)
    with col4:
        sok = st.text_input("Søk nettbutikk", placeholder="Skriv navn...")

    # Filtrer
    vis = alle
    if status_filter == "Går videre":
        vis = [s for s in vis if s.get("status") == "inn" and not s.get("enk")]
    elif status_filter == "Filtrert ut":
        vis = [s for s in vis if s.get("status") == "ut" and not s.get("enk")]
    elif status_filter == "Krever sjekk":
        vis = [s for s in vis if s.get("status") == "usikker"]
    elif status_filter == "ENK":
        vis = [s for s in vis if s.get("enk")]
    if klasse_filter != "Alle":
        vis = [s for s in vis if s.get("klasse") == klasse_filter]
    if bransje_filter != "Alle":
        vis = [s for s in vis if s.get("bransje") == bransje_filter]
    if sok:
        vis = [s for s in vis if sok.lower() in s.get("name", "").lower()]

    vis = sorted(vis, key=lambda x: x.get("total") or 0, reverse=True)

    st.markdown(f"**Viser {len(vis)} butikker**")

    # Tabell
    if vis:
        df_data = []
        for s in vis:
            df_data.append({
                "Nettbutikk": s.get("name", ""),
                "Bransje": s.get("bransje", "–"),
                "Org.form": s.get("orgform", "–"),
                "Status": "✓ Videre" if s.get("status") == "inn" and not s.get("enk") else "⛔ ENK" if s.get("enk") else "✕ Ut" if s.get("status") == "ut" else "? Sjekk",
                "Klasse": s.get("klasse", "–"),
                "Omsetning": s.get("omsetning", "–"),
                "Score": s.get("total", "–"),
                "Inntrykk": s.get("inntrykk", "–"),
                "IKS": s.get("iks", "–"),
                "Kassen": s.get("kassen", "–"),
                "Markedsf.": s.get("markedsforing", "–"),
            })

        df = pd.DataFrame(df_data)
        st.dataframe(df, use_container_width=True, hide_index=True)

        # Eksporter
        csv = df.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            "↓ Eksporter CSV",
            data=csv.encode("utf-8-sig"),
            file_name="netthandelsprisen_screening.csv",
            mime="text/csv"
        )

    # Detaljvisning
    st.markdown("---")
    st.markdown("### Detaljvisning")
    valgt = st.selectbox("Velg butikk for detaljer", ["–"] + [s.get("name") for s in vis])

    if valgt and valgt != "–":
        butikk = r.get(valgt, {})

        col1, col2 = st.columns([2, 1])

        with col1:
            st.markdown(f"## {butikk.get('name')}")
            if butikk.get("url"):
                st.markdown(f"🌐 [{butikk.get('url')}]({butikk.get('url')})")
            st.markdown(butikk.get("kommentar", ""))

            # Advarsler
            if butikk.get("b2cAdvarsel"):
                st.warning(f"🔍 {butikk.get('b2cAdvarsel')}")
            if butikk.get("brregUsikker"):
                st.warning("🔍 Brreg-treff usikkert – org.form bør verifiseres manuelt")
            ap = butikk.get("apenhetsloven", {})
            if ap and ap.get("palagt") and not ap.get("rapport_funnet"):
                st.warning(f"⚠️ Åpenhetsloven: Ser ut til å være rapporteringspliktig men ingen rapport funnet. {ap.get('kommentar', '')}")

        with col2:
            if butikk.get("total") is not None:
                st.metric("Totalpoeng", butikk.get("total"))
                st.metric("Første inntrykk", butikk.get("inntrykk"))
                st.metric("Info/KS/BK", butikk.get("iks"))
                st.metric("Kassen", butikk.get("kassen"))
                st.metric("Markedsføring", butikk.get("markedsforing"))

        # Kriteriegjennomgang
        if butikk.get("scoring"):
            scoring = butikk["scoring"]
            tab1, tab2, tab3, tab4 = st.tabs([
                "Første inntrykk", "Info/KS/Bærekraft", "Kassen", "Markedsføring"
            ])

            def vis_kriterier(tab, kriterier):
                with tab:
                    for k in kriterier:
                        col_a, col_b = st.columns([1, 4])
                        with col_a:
                            score = k.get("score", 0)
                            farge = "🟢" if score >= 4 else "🟡" if score >= 3 else "🔴"
                            st.markdown(f"**{farge} {score}/5**")
                            if k.get("vekt"):
                                st.caption(k["vekt"])
                        with col_b:
                            st.markdown(f"**{k.get('navn')}**")
                            st.caption(k.get("begrunnelse", ""))
                        st.divider()

            vis_kriterier(tab1, scoring.get("inntrykk", []))
            vis_kriterier(tab2, scoring.get("iks", []))
            vis_kriterier(tab3, scoring.get("kassen", []))
            vis_kriterier(tab4, scoring.get("markedsforing", []))


# ─────────────────────────────────────────────
# SIDE 3 – TOPP 100
# ─────────────────────────────────────────────
elif side == "⭐ Topp 100":
    st.header("⭐ Topp 100 – Juryens arbeidsflate")
    st.markdown("Topp 33-34 per klasse basert på agentens score. Juryen setter sine vurderinger her.")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå.")
        st.stop()

    r = st.session_state.resultater

    # Hent topp per klasse
    def hent_topp(klasse, antall=34):
        return sorted(
            [v for v in r.values() if v.get("status") == "inn" and not v.get("enk") and v.get("klasse") == klasse and v.get("total") is not None],
            key=lambda x: x.get("total", 0),
            reverse=True
        )[:antall]

    liten = hent_topp("Liten")
    medium = hent_topp("Medium")
    stor = hent_topp("Stor")
    alle_topp = liten + medium + stor

    # Statistikk
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Totalt i Topp 100", len(alle_topp))
    col2.metric("Liten klasse", len(liten))
    col3.metric("Medium klasse", len(medium))
    col4.metric("Stor klasse", len(stor))

    # Klasse-filter
    klasse_tab = st.radio("Vis klasse", ["Alle", "Liten", "Medium", "Stor"], horizontal=True)
    if klasse_tab == "Liten":
        vis_liste = liten
    elif klasse_tab == "Medium":
        vis_liste = medium
    elif klasse_tab == "Stor":
        vis_liste = stor
    else:
        vis_liste = alle_topp

    st.markdown("---")

    # Last inn lagrede juryvurderinger
    juryvurderinger = hent_juryvurderinger()

    for i, butikk in enumerate(vis_liste):
        navn = butikk.get("name", "")
        lagret = juryvurderinger.get(navn, {})

        with st.expander(f"**{i+1}. {navn}** – Score: {butikk.get('total', '–')} | {butikk.get('klasse')} | {butikk.get('bransje', '–')}"):
            col1, col2, col3 = st.columns([2, 1, 1])

            with col1:
                if butikk.get("url"):
                    st.markdown(f"🌐 [{butikk.get('url')}]({butikk.get('url')})")
                st.caption(butikk.get("kommentar", ""))

            with col2:
                st.metric("Agentscore", butikk.get("total"))
                st.metric("Inntrykk", butikk.get("inntrykk"))
                st.metric("IKS", butikk.get("iks"))

            with col3:
                st.metric("Kassen", butikk.get("kassen"))
                st.metric("Markedsf.", butikk.get("markedsforing"))

            st.markdown("**Juryens vurdering**")

            jury_score = st.select_slider(
                f"Juryscore for {navn}",
                options=[1, 2, 3, 4, 5],
                value=int(lagret.get("Juryscore", 3)),
                format_func=lambda x: "⭐" * x,
                key=f"score_{navn}"
            )

            jury_status = st.selectbox(
                f"Status for {navn}",
                ["Ikke vurdert", "Kandidat", "Semifinalist", "Finalist", "Vinner"],
                index=["Ikke vurdert", "Kandidat", "Semifinalist", "Finalist", "Vinner"].index(
                    lagret.get("Status", "Ikke vurdert")
                ),
                key=f"status_{navn}"
            )

            jury_notat = st.text_area(
                f"Notat for {navn}",
                value=lagret.get("Notat", ""),
                placeholder="Skriv jurynotat her...",
                key=f"notat_{navn}"
            )

            if st.button(f"💾 Lagre vurdering", key=f"lagre_{navn}"):
                ok = lagre_juryvurdering(navn, jury_score, jury_status, jury_notat)
                if ok:
                    st.success("Lagret til Google Sheets!")
                else:
                    # Fallback – lagre i session state
                    st.session_state.resultater[navn]["juryScore"] = jury_score
                    st.session_state.resultater[navn]["juryStatus"] = jury_status
                    st.session_state.resultater[navn]["juryNote"] = jury_notat
                    st.success("Lagret!")

    # Eksporter Topp 100
    st.markdown("---")
    if st.button("↓ Eksporter Topp 100 som CSV"):
        df_data = [{
            "Rang": i+1,
            "Navn": b.get("name"),
            "URL": b.get("url"),
            "Klasse": b.get("klasse"),
            "Bransje": b.get("bransje"),
            "Agentscore": b.get("total"),
            "Juryscore": b.get("juryScore", ""),
            "Jurystatus": b.get("juryStatus", ""),
            "Jurynotat": b.get("juryNote", ""),
        } for i, b in enumerate(alle_topp)]
        df = pd.DataFrame(df_data)
        csv = df.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            "Last ned CSV",
            data=csv.encode("utf-8-sig"),
            file_name="topp100_netthandelsprisen.csv",
            mime="text/csv"
        )


# ─────────────────────────────────────────────
# SIDE 4 – FINALE
# ─────────────────────────────────────────────
elif side == "🏆 Finale":
    st.header("🏆 Finale – Juryens endelige vurderinger")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå.")
        st.stop()

    r = st.session_state.resultater
    juryvurderinger = hent_juryvurderinger()

    finale_statuser = ["Semifinalist", "Finalist", "Vinner"]

    # Kombiner agent-data med juryvurderinger
    finalister = []
    for navn, butikk in r.items():
        lagret = juryvurderinger.get(navn, {})
        status = lagret.get("Status", butikk.get("juryStatus", "Ikke vurdert"))
        if status in finale_statuser:
            butikk["juryStatus"] = status
            butikk["juryScore"] = lagret.get("Juryscore", butikk.get("juryScore", 0))
            butikk["juryNote"] = lagret.get("Notat", butikk.get("juryNote", ""))
            finalister.append(butikk)

    if not finalister:
        st.info("Ingen butikker er merket som Semifinalist, Finalist eller Vinner ennå. Gå til Topp 100 for å sette jurystatus.")
        st.stop()

    # Sorter – Vinnere øverst, deretter Finalister, Semifinalister
    rang = {"Vinner": 0, "Finalist": 1, "Semifinalist": 2}
    finalister.sort(key=lambda x: (rang.get(x.get("juryStatus"), 3), -(x.get("juryScore") or 0)))

    # Filter
    vis_filter = st.radio("Vis", ["Alle", "Vinnere", "Finalister", "Semifinalister"], horizontal=True)
    if vis_filter == "Vinnere":
        finalister = [f for f in finalister if f.get("juryStatus") == "Vinner"]
    elif vis_filter == "Finalister":
        finalister = [f for f in finalister if f.get("juryStatus") == "Finalist"]
    elif vis_filter == "Semifinalister":
        finalister = [f for f in finalister if f.get("juryStatus") == "Semifinalist"]

    st.markdown(f"**{len(finalister)} butikker i finalen**")

    status_farger = {"Vinner": "#C8102E", "Finalist": "#B8860B", "Semifinalist": "#1B6B3A"}
    status_ikoner = {"Vinner": "🏆", "Finalist": "★", "Semifinalist": "◐"}

    for i, butikk in enumerate(finalister):
        status = butikk.get("juryStatus", "")
        farge = status_farger.get(status, "#666")
        ikon = status_ikoner.get(status, "")
        stjerner = "⭐" * int(butikk.get("juryScore") or 0)

        st.markdown(f"""
        <div style="background:white;border-radius:8px;padding:16px;margin-bottom:12px;border-left:4px solid {farge};box-shadow:0 1px 4px rgba(0,0,0,0.07)">
            <div style="display:flex;align-items:center;gap:12px;margin-bottom:8px">
                <span style="font-size:20px;font-weight:700;color:{farge}">{ikon} {i+1}</span>
                <div>
                    <div style="font-size:16px;font-weight:700">{butikk.get('name')}</div>
                    <a href="{butikk.get('url','#')}" target="_blank" style="font-size:12px;color:#C8102E">{butikk.get('url','')}</a>
                </div>
                <div style="margin-left:auto;text-align:right">
                    <div style="font-weight:700;color:{farge}">{status}</div>
                    <div style="font-size:18px">{stjerner}</div>
                </div>
            </div>
            <div style="display:flex;gap:16px;font-size:12px;color:#666">
                <span>Klasse: <strong>{butikk.get('klasse')}</strong></span>
                <span>Agentscore: <strong>{butikk.get('total')}</strong></span>
                <span>Bransje: <strong>{butikk.get('bransje','–')}</strong></span>
            </div>
            {f'<div style="margin-top:8px;font-size:13px;color:#666;font-style:italic">"{butikk.get("juryNote")}"</div>' if butikk.get("juryNote") else ""}
        </div>
        """, unsafe_allow_html=True)

    # Eksporter finale
    st.markdown("---")
    if st.button("↓ Eksporter finale som CSV"):
        df_data = [{
            "Jurystatus": b.get("juryStatus"),
            "Navn": b.get("name"),
            "URL": b.get("url"),
            "Klasse": b.get("klasse"),
            "Bransje": b.get("bransje"),
            "Agentscore": b.get("total"),
            "Juryscore": b.get("juryScore", ""),
            "Jurynotat": b.get("juryNote", ""),
        } for b in finalister]
        df = pd.DataFrame(df_data)
        csv = df.to_csv(index=False, sep=";", encoding="utf-8-sig")
        st.download_button(
            "Last ned finale CSV",
            data=csv.encode("utf-8-sig"),
            file_name="finale_netthandelsprisen.csv",
            mime="text/csv"
        )


# ─────────────────────────────────────────────
# SIDE 5 – LOGISTIKK
# ─────────────────────────────────────────────
elif side == "📦 Logistikk":
    st.header("📦 Logistikkrapport – Posten Bring")
    st.markdown("Kartlegging av logistikkpartnere blant scorede butikker.")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå.")
        st.stop()

    r = st.session_state.resultater
    med_logistikk = [v for v in r.values() if v.get("status") == "inn" and v.get("logistikk")]

    if not med_logistikk:
        st.info("Ingen logistikkdata tilgjengelig ennå.")
        st.stop()

    # Tell opp
    counts = {a: sum(1 for s in med_logistikk if s.get("logistikk", {}).get(a)) for a in LOGISTIKK_AKTORER}

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Logistikkpartnere blant scorede butikker**")
        df_logi = pd.DataFrame([
            {"Partner": k, "Antall butikker": v}
            for k, v in sorted(counts.items(), key=lambda x: x[1], reverse=True)
        ])
        st.bar_chart(df_logi.set_index("Partner"))

    with col2:
        st.markdown("**Status Posten/Bring**")
        pb_count = counts.get("Posten/Bring", 0)
        ikke_pb = len(med_logistikk) - pb_count
        st.metric("Bruker Posten/Bring", pb_count)
        st.metric("Bruker ikke Posten/Bring", ikke_pb)
        st.metric("Salgspotensial", ikke_pb, help="Butikker som ikke bruker Posten/Bring i dag")

    # Salgsmuligheter
    st.markdown("---")
    st.markdown("### Salgsmuligheter – bruker ikke Posten/Bring")

    prospects = [
        s for s in med_logistikk
        if not s.get("logistikk", {}).get("Posten/Bring")
    ]
    prospects.sort(key=lambda x: x.get("total", 0), reverse=True)

    if prospects:
        df_data = [{
            "Butikk": s.get("name"),
            "Klasse": s.get("klasse"),
            "Bransje": s.get("bransje", "–"),
            "Score": s.get("total"),
            "Bruker nå": ", ".join([a for a in LOGISTIKK_AKTORER if a != "Posten/Bring" and s.get("logistikk", {}).get(a)]) or "Ukjent"
        } for s in prospects]
        st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
# SIDE 6 – OM VERKTØYET
# ─────────────────────────────────────────────
elif side == "ℹ️ Om verktøyet":
    st.header("ℹ️ Om Netthandelsprisen – Jurysystem")

    st.markdown("""
    ### Scoringsmodell

    **4 kategorier – 25% vekt hver:**

    | Kategori | Vekt | Kriterier |
    |---|---|---|
    | Første inntrykk | 25% | Startside, bilder/film, produktinfo, søk |
    | Info, kundeservice og bærekraft | 25% | KLR (35%), Kundeservice (35%), Bærekraft (30%) |
    | Kassen / mersalg | 25% | Leveringsvalg, betaling, mersalg |
    | Markedsføring | 25% | SoMe (40%), Kundeklubb (30%), Nyhetsbrev (30%) |

    **Kjøpsvilkår, levering og retur (35% av IKS):**
    - Kjøpsvilkår: 20%
    - Leveringsinformasjon: 40%
    - Returløsning: 40%

    **Markedsføring – kundeklubb-scoring:**
    | Score | Betydning |
    |---|---|
    | 2 | Ingen kundeklubb |
    | 3 | Kun inngangsrabatt |
    | 4 | Poeng/rabatter/lojalitetsfordeler |
    | 5 | Full lojalitetspakke |

    ### KO-kriterier
    - ENK (Enkeltmannsforetak) – filtreres automatisk ut
    - Ingen URL funnet – kan ikke identifiseres
    - B2C og skandinavisk tilstedeværelse vises som advarsel – juryen bestemmer

    ### Datakilder
    - Org.form: Brønnøysundregisteret (velger AS over ENK ved konflikt)
    - Omsetning: Regnskapsregisteret (ekte regnskapstall)
    - Åpenhetsloven: Flagges med advarsel hvis mangler

    ### Versjon
    v5.0 · Juli 2026 · Posten Bring
    """)
