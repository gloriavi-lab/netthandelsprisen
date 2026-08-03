"""
NETTHANDELSPRISEN – POSTEN BRING
Streamlit-app v5 med full rik visning
"""

import streamlit as st
import json
import pandas as pd
import gspread
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

st.markdown("""
<style>
.stApp { background-color: #F5F4F2; }
.main-header {
    background: #212121; color: white; padding: 16px 28px;
    border-radius: 10px; margin-bottom: 24px;
    display: flex; align-items: center; gap: 16px;
}
.logo-badge {
    background: #C8102E; color: white; padding: 5px 12px;
    border-radius: 4px; font-weight: 700; font-size: 12px;
    text-transform: uppercase; letter-spacing: 0.5px;
}
.stat-box {
    background: white; border-radius: 8px; padding: 14px 16px;
    border-top: 3px solid #D8D6D2; cursor: pointer;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07); text-align: center;
    transition: all 0.15s;
}
.stat-box:hover { transform: translateY(-1px); box-shadow: 0 3px 10px rgba(0,0,0,0.1); }
.stat-label { font-size: 10px; color: #999; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }
.stat-value { font-size: 24px; font-weight: 700; }
.detail-panel {
    background: white; border-radius: 12px; padding: 24px 28px;
    margin-bottom: 20px; border-top: 4px solid #C8102E;
    box-shadow: 0 2px 16px rgba(0,0,0,0.08);
}
.score-pill {
    display: inline-block; border-radius: 4px; padding: 2px 8px;
    font-weight: 700; font-size: 13px; min-width: 32px; text-align: center;
}
.logi-icon {
    display: inline-block; background: #E6F4EC; color: #1B6B3A;
    border-radius: 4px; padding: 2px 6px; font-size: 11px;
    font-weight: 600; margin: 2px;
}
.warning-box {
    background: #FEF3E2; border-left: 3px solid #E8A020;
    border-radius: 0 8px 8px 0; padding: 10px 14px;
    font-size: 12px; color: #7A4800; margin-bottom: 10px;
}
.tech-ok { background: #E6F4EC; color: #1B6B3A; padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: 500; }
.tech-warn { background: #FEF3E2; color: #7A4800; padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: 500; }
.tech-bad { background: #FDECEA; color: #C8102E; padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: 500; }
.trust-badge { background: #E6F0FA; color: #0D4A8A; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; display: inline-block; margin: 2px; }
.cat-header { font-size: 13px; font-weight: 700; margin-bottom: 8px; padding-bottom: 6px; border-bottom: 1.5px solid #E8E6E2; display: flex; align-items: center; gap: 8px; }
.crit-row { display: flex; gap: 10px; padding: 8px 0; border-bottom: 1px solid #E8E6E2; align-items: flex-start; }
.crit-score { min-width: 36px; text-align: center; }
.crit-name { font-size: 12px; font-weight: 600; }
.crit-beg { font-size: 11px; color: #666; }
.crit-vekt { font-size: 10px; color: #999; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <span class="logo-badge">Posten Bring</span>
    <div>
        <div style="font-size:18px;font-weight:700">Netthandelsprisen – Jurysystem</div>
        <div style="font-size:13px;color:#aaa">Automatisk screening og scoringsagent</div>
    </div>
    <div style="margin-left:auto;background:#2a2a2a;color:#888;padding:3px 10px;border-radius:4px;font-size:12px">2026</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "resultater" not in st.session_state:
    st.session_state.resultater = {}
if "valgt_butikk" not in st.session_state:
    st.session_state.valgt_butikk = None
if "screening_filter" not in st.session_state:
    st.session_state.screening_filter = "alle"

# ─────────────────────────────────────────────
# GOOGLE SHEETS
# ─────────────────────────────────────────────
def koble_sheets():
    try:
        creds_info = st.secrets["gcp_service_account"]
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        gc = gspread.authorize(creds)
        sheet_id = st.secrets["google_sheets"]["sheet_id"]
        return gc.open_by_key(sheet_id).sheet1
    except Exception:
        return None

def lagre_jury(navn, score, status, notat):
    ws = koble_sheets()
    if not ws:
        return False
    try:
        data = ws.get_all_records()
        for i, rad in enumerate(data, 2):
            if rad.get("Navn") == navn:
                ws.update(f"D{i}:F{i}", [[score, status, notat]])
                return True
        ws.append_row([navn, "", "", score, status, notat])
        return True
    except Exception:
        return False

def hent_jury() -> dict:
    ws = koble_sheets()
    if not ws:
        return {}
    try:
        return {r["Navn"]: r for r in ws.get_all_records() if r.get("Navn")}
    except Exception:
        return {}

# ─────────────────────────────────────────────
# HJELPEFUNKSJONER
# ─────────────────────────────────────────────
def score_html(s, stor=False):
    if s is None:
        return '<span style="color:#999">–</span>'
    n = float(s)
    if n >= 4.5:
        c, bg = "#1B6B3A", "#E6F4EC"
    elif n >= 3.5:
        c, bg = "#0D4A8A", "#E6F0FA"
    elif n >= 2.5:
        c, bg = "#7A4800", "#FEF3E2"
    else:
        c, bg = "#C8102E", "#FDECEA"
    size = "16px" if stor else "13px"
    return f'<span style="background:{bg};color:{c};padding:3px 9px;border-radius:4px;font-weight:700;font-size:{size}">{n:.1f}</span>'

def klasse_html(k):
    m = {"Liten": ("#0D4A8A","#E6F0FA"), "Medium": ("#7A4800","#FEF3E2"), "Stor": ("#5B2D8E","#F0E8FA")}
    c, bg = m.get(k, ("#666","#eee"))
    return f'<span style="background:{bg};color:{c};padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">{k}</span>'

def status_html(s, enk=False):
    if enk:
        return '<span style="background:#FDECEA;color:#C8102E;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">⛔ ENK</span>'
    if s == "inn":
        return '<span style="background:#E6F4EC;color:#1B6B3A;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">✓ Videre</span>'
    if s == "ut":
        return '<span style="background:#FDECEA;color:#C8102E;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">✕ Ut</span>'
    return '<span style="background:#FEF3E2;color:#7A4800;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700">? Sjekk</span>'

def logi_html(logi):
    if not logi:
        return '<span style="color:#999;font-size:11px">–</span>'
    treff = [a for a in LOGISTIKK_AKTORER if logi.get(a)]
    if not treff:
        return '<span style="color:#999;font-size:11px">Ukjent</span>'
    kort = {"Posten/Bring":"P/B","PostNord":"PN","Helthjem":"HH","Instabox":"IB","Porterbuddy":"PB","DHL":"DHL","Budbee":"BB","UPS/FedEx":"UPS","Egne biler":"Egne"}
    return " ".join(f'<span class="logi-icon">{kort.get(a,a)}</span>' for a in treff)

def tech_html(tech):
    if not tech:
        return '<span style="color:#999;font-size:11px">Ikke sjekket</span>'
    def badge(v, l):
        cls = "tech-ok" if v == "ok" else "tech-warn" if v == "warn" else "tech-bad"
        ikon = "✓" if v == "ok" else "⚠" if v == "warn" else "✕"
        return f'<span class="{cls}">{ikon} {l}</span>'
    tp = tech.get("trustpilot", "Ikke funnet")
    tp_html = f'<span class="tech-ok">⭐ {tp}</span>' if tp != "Ikke funnet" else '<span class="tech-warn">⭐ Trustpilot ikke funnet</span>'
    return (badge(tech.get("mobil","warn"), "Mobil") + " " +
            badge(tech.get("ssl","warn"), "SSL") + " " +
            badge(tech.get("lastetid","warn"), "Lastetid") + " " + tp_html)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏆 Netthandelsprisen")
    st.markdown("---")

    side = st.radio(
        "Naviger",
        ["📂 Last opp og kjør", "📋 Screening", "⭐ Topp 100", "🏆 Finale", "📦 Logistikk", "ℹ️ Om verktøyet"],
        label_visibility="collapsed"
    )

    st.markdown("---")

    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if api_key:
        st.success("✅ API-nøkkel hentet automatisk")
    else:
        api_key = st.text_input("API-nøkkel", type="password")

    if st.session_state.resultater:
        r = st.session_state.resultater
        st.markdown("---")
        st.markdown("**Statistikk**")
        col1, col2 = st.columns(2)
        col1.metric("Totalt", len(r))
        col2.metric("Videre", sum(1 for v in r.values() if v.get("status")=="inn" and not v.get("enk")))
        col1.metric("Ut", sum(1 for v in r.values() if v.get("status")=="ut" and not v.get("enk")))
        col2.metric("ENK", sum(1 for v in r.values() if v.get("enk")))


# ─────────────────────────────────────────────
# DETALJPANEL – brukes på tvers av faner
# ─────────────────────────────────────────────
def vis_detaljpanel(butikk, juryvurderinger={}):
    navn = butikk.get("name","")
    lagret = juryvurderinger.get(navn, {})

    with st.container():
        st.markdown('<div class="detail-panel">', unsafe_allow_html=True)

        # Topp-seksjon
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f"## {navn}")
            if butikk.get("url"):
                st.markdown(f'<a href="{butikk["url"]}" target="_blank" style="background:#C8102E;color:#fff;padding:6px 14px;border-radius:6px;text-decoration:none;font-size:12px;font-weight:600">🌐 Besøk nettbutikk</a>&nbsp;&nbsp;<span style="font-size:11px;color:#999">{butikk["url"].replace("https://","").replace("http://","")}</span>', unsafe_allow_html=True)
            st.markdown("")
            # Meta-badges
            badges = status_html(butikk.get("status"), butikk.get("enk"))
            if butikk.get("klasse") and butikk["klasse"] != "-":
                badges += " " + klasse_html(butikk["klasse"])
            badges += f' <span style="background:#E8E6E2;color:#666;padding:2px 8px;border-radius:20px;font-size:11px">{butikk.get("orgform","–")}</span>'
            if butikk.get("bransje"):
                badges += f' <span style="background:#F0E8FA;color:#5B2D8E;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600">{butikk["bransje"]}</span>'
            badges += f' <span style="background:#E8E6E2;color:#666;padding:2px 8px;border-radius:20px;font-size:11px">{butikk.get("omsetning","–")}</span>'
            st.markdown(badges, unsafe_allow_html=True)
        with col2:
            if butikk.get("total") is not None:
                st.markdown(f'<div style="text-align:center"><div style="font-size:11px;color:#999;text-transform:uppercase;letter-spacing:0.5px;font-weight:600">Totalpoeng</div><div style="font-size:36px;font-weight:700">{butikk["total"]}</div></div>', unsafe_allow_html=True)

        # Advarsler
        for advarsel in butikk.get("advarsler", []):
            st.markdown(f'<div class="warning-box">🔍 {advarsel}</div>', unsafe_allow_html=True)

        ap = butikk.get("apenhetsloven", {})
        if ap and ap.get("palagt") and not ap.get("rapport_funnet"):
            st.markdown(f'<div class="warning-box">⚠️ <strong>Åpenhetsloven:</strong> Ser ut til å være rapporteringspliktig men ingen rapport funnet for 2024/2025. {ap.get("kommentar","")}</div>', unsafe_allow_html=True)

        # Juryverktøy
        st.markdown("---")
        st.markdown("**🎯 Juryverktøy**")
        jcol1, jcol2, jcol3 = st.columns([2, 2, 3])
        with jcol1:
            jury_score = st.select_slider(
                "Juryscore",
                options=[1, 2, 3, 4, 5],
                value=int(lagret.get("Juryscore", butikk.get("juryScore", 3))),
                format_func=lambda x: "⭐" * x,
                key=f"js_{navn}"
            )
        with jcol2:
            jury_status = st.selectbox(
                "Status",
                ["Ikke vurdert", "Kandidat", "Semifinalist", "Finalist", "Vinner"],
                index=["Ikke vurdert", "Kandidat", "Semifinalist", "Finalist", "Vinner"].index(
                    lagret.get("Status", butikk.get("juryStatus", "Ikke vurdert"))
                ),
                key=f"jstat_{navn}"
            )
        with jcol3:
            jury_notat = st.text_area(
                "Jurynotat",
                value=lagret.get("Notat", butikk.get("juryNote", "")),
                placeholder="Skriv jurynotat...",
                height=80,
                key=f"jnot_{navn}"
            )
        if st.button("💾 Lagre vurdering", key=f"jlagre_{navn}"):
            ok = lagre_jury(navn, jury_score, jury_status, jury_notat)
            st.session_state.resultater[navn]["juryScore"] = jury_score
            st.session_state.resultater[navn]["juryStatus"] = jury_status
            st.session_state.resultater[navn]["juryNote"] = jury_notat
            st.success("Lagret!" if ok else "Lagret lokalt (Google Sheets ikke koblet)")

        # Screeningresultat
        st.markdown("---")
        st.markdown(f'<div style="background:#F5F4F2;border-radius:8px;padding:10px 14px;border-left:3px solid #C8102E;font-size:12px;color:#666;margin-bottom:12px"><strong style="color:#999;font-size:10px;text-transform:uppercase;letter-spacing:0.5px">Screeningresultat</strong><br/>{butikk.get("screeningBegrunnelse","")}</div>', unsafe_allow_html=True)

        if butikk.get("kommentar"):
            st.markdown(f'<div style="background:#F5F4F2;border-radius:8px;padding:12px 14px;font-size:13px;color:#444;line-height:1.7;margin-bottom:16px">{butikk["kommentar"]}</div>', unsafe_allow_html=True)

        # Score-kort
        if butikk.get("total") is not None:
            st.markdown("**Kategoriscorer**")
            scol1, scol2, scol3, scol4, scol5 = st.columns(5)
            scol1.markdown(f'<div style="text-align:center"><div style="font-size:10px;color:#999;font-weight:600;text-transform:uppercase">Total</div>{score_html(butikk.get("total"), stor=True)}</div>', unsafe_allow_html=True)
            scol2.markdown(f'<div style="text-align:center"><div style="font-size:10px;color:#C8102E;font-weight:600;text-transform:uppercase">Inntrykk</div>{score_html(butikk.get("inntrykk"))}</div>', unsafe_allow_html=True)
            scol3.markdown(f'<div style="text-align:center"><div style="font-size:10px;color:#E87D3E;font-weight:600;text-transform:uppercase">Info/KS/BK</div>{score_html(butikk.get("iks"))}<div style="font-size:9px;color:#999;margin-top:3px">KLR:{butikk.get("iksDetalj",{}).get("klr","–")} KS:{butikk.get("iksDetalj",{}).get("kundeservice","–")} BK:{butikk.get("iksDetalj",{}).get("baerekraft","–")}</div></div>', unsafe_allow_html=True)
            scol4.markdown(f'<div style="text-align:center"><div style="font-size:10px;color:#0D4A8A;font-weight:600;text-transform:uppercase">Kassen/Mersalg</div>{score_html(butikk.get("kat3"))}<div style="font-size:9px;color:#999;margin-top:3px">K:{butikk.get("kat3Detalj",{}).get("kassen","–")} M:{butikk.get("kat3Detalj",{}).get("mersalg","–")} I:{butikk.get("kat3Detalj",{}).get("inspirasjon","–")}</div></div>', unsafe_allow_html=True)
            scol5.markdown(f'<div style="text-align:center"><div style="font-size:10px;color:#1B6B3A;font-weight:600;text-transform:uppercase">Markedsf.</div>{score_html(butikk.get("markedsforing"))}</div>', unsafe_allow_html=True)
            st.markdown("")

        # Teknisk kvalitet
        if butikk.get("tech"):
            st.markdown("**Teknisk kvalitet og mobilopplevelse**")
            st.markdown(tech_html(butikk.get("tech")), unsafe_allow_html=True)
            st.markdown("")

        # Tillit og trygghet
        if butikk.get("trust"):
            st.markdown("**Tillit og trygghet**")
            trust_html = " ".join(f'<span class="trust-badge">{t}</span>' for t in butikk["trust"])
            st.markdown(trust_html, unsafe_allow_html=True)
            st.markdown("")

        # Logistikkpartnere
        if butikk.get("logistikk"):
            st.markdown("**Logistikkpartnere**")
            cols = st.columns(len(LOGISTIKK_AKTORER))
            for j, aktør in enumerate(LOGISTIKK_AKTORER):
                har = butikk["logistikk"].get(aktør, False)
                bg = "#E6F4EC" if har else "#F5F4F2"
                farge = "#1B6B3A" if har else "#D8D6D2"
                cols[j].markdown(f'<div style="background:{bg};border-radius:6px;padding:6px;text-align:center;font-size:10px;font-weight:600;color:{farge}">{aktør}</div>', unsafe_allow_html=True)
            st.markdown("")

        # Kriteriegjennomgang
        if butikk.get("scoring"):
            st.markdown("---")
            st.markdown("**Kriteriegjennomgang**")

            scoring = butikk["scoring"]
            tab1, tab2, tab3, tab4 = st.tabs([
                "🔍 Første inntrykk", "📋 Info/KS/Bærekraft", "🛒 Kassen/Mersalg", "📣 Markedsføring"
            ])

            kat_farger = {"🔍 Første inntrykk": "#C8102E", "📋 Info/KS/Bærekraft": "#E87D3E", "🛒 Kassen/Mersalg": "#0D4A8A", "📣 Markedsføring": "#1B6B3A"}

            def vis_kriterier(tab, kriterier, farge):
                with tab:
                    for k in kriterier:
                        score = k.get("score", 0)
                        n = float(score)
                        if n >= 4.5: c, bg = "#1B6B3A", "#E6F4EC"
                        elif n >= 3.5: c, bg = "#0D4A8A", "#E6F0FA"
                        elif n >= 2.5: c, bg = "#7A4800", "#FEF3E2"
                        else: c, bg = "#C8102E", "#FDECEA"

                        st.markdown(f"""
                        <div class="crit-row">
                            <div class="crit-score"><span style="background:{bg};color:{c};padding:3px 8px;border-radius:4px;font-weight:700;font-size:13px">{score}</span></div>
                            <div>
                                <div class="crit-name">{k.get("navn","")}</div>
                                <div class="crit-beg">{k.get("begrunnelse","")}</div>
                                <div class="crit-vekt">{k.get("vekt","")}</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

            vis_kriterier(tab1, scoring.get("inntrykk", []), "#C8102E")
            vis_kriterier(tab2, scoring.get("iks", []), "#E87D3E")
            vis_kriterier(tab3, scoring.get("kassen", []), "#0D4A8A")
            vis_kriterier(tab4, scoring.get("markedsforing", []), "#1B6B3A")

        if st.button("✕ Lukk detaljer", key=f"lukk_{navn}"):
            st.session_state.valgt_butikk = None
            st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)


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
        2. Velg antall (10 for test, alle for produksjon)
        3. Trykk "Start screening"
        4. Agenten screener og scorer automatisk
        5. Se resultater i Screening- og Topp 100-fanene
        """)
        fil = st.file_uploader("Last opp Excel-fil (.xlsx)", type=["xlsx"])
        maks = st.selectbox("Antall butikker", [10, 25, 50, 100, 250, 500, "Alle"], index=0)
        start = st.button("🚀 Start screening", type="primary", disabled=not fil or not api_key)

    with col2:
        kostnader = {10:"~$0.50", 25:"~$1.25", 50:"~$2.50", 100:"~$5.00", 250:"~$12.50", 500:"~$25.00", "Alle":"~$40-50"}
        tider = {10:"3-5 min", 25:"8-12 min", 50:"15-20 min", 100:"30-40 min", 250:"1.5-2 t", 500:"3-4 t", "Alle":"6-8 t"}
        st.info(f"**Estimert kostnad:** {kostnader.get(maks,'?')}\n\n**Estimert tid:** {tider.get(maks,'?')}")

    if start and fil and api_key:
        maks_int = None if maks == "Alle" else int(maks)
        st.markdown("---")
        fremgang = st.empty()
        bar = st.progress(0)
        logg = st.empty()
        linjer = []

        def cb(melding, prosent):
            fremgang.markdown(f"**{melding}**")
            bar.progress(prosent / 100)
            linjer.append(melding)
            if len(linjer) > 8:
                linjer.pop(0)
            logg.code("\n".join(linjer))

        try:
            res = kjor_agent(fil, api_key, maks_int, cb)
            st.session_state.resultater = res
            bar.progress(1.0)
            st.success(f"✅ Ferdig! {len(res)} butikker behandlet.")
            st.balloons()
        except Exception as e:
            st.error(f"❌ Feil: {e}")


# ─────────────────────────────────────────────
# SIDE 2 – SCREENING
# ─────────────────────────────────────────────
elif side == "📋 Screening":
    st.header("📋 Screening – alle butikker")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå. Gå til 'Last opp og kjør'.")
        st.stop()

    r = st.session_state.resultater
    alle = list(r.values())

    # Statistikkbokser
    def stat_box(label, verdi, farge, filter_key):
        aktiv = st.session_state.screening_filter == filter_key
        border = f"border-top-color:{farge}" if aktiv else ""
        box_style = f"border-top-color:{farge};box-shadow:0 0 0 2px {farge}" if aktiv else ""
        return f'<div class="stat-box" style="{box_style}" onclick=""><div class="stat-label">{label}</div><div class="stat-value" style="color:{farge}">{verdi}</div></div>'

    scol = st.columns(8)
    filtre = [
        ("alle", "Totalt", len(alle), "#1A1A1A"),
        ("inn", "Går videre", sum(1 for s in alle if s.get("status")=="inn" and not s.get("enk")), "#1B6B3A"),
        ("ut", "Filtrert ut", sum(1 for s in alle if s.get("status")=="ut" and not s.get("enk")), "#C8102E"),
        ("usikker", "Krever sjekk", sum(1 for s in alle if s.get("status")=="usikker"), "#7A4800"),
        ("enk", "ENK – ut", sum(1 for s in alle if s.get("enk")), "#C8102E"),
        ("liten", "Liten", sum(1 for s in alle if s.get("klasse")=="Liten"), "#0D4A8A"),
        ("medium", "Medium", sum(1 for s in alle if s.get("klasse")=="Medium"), "#7A4800"),
        ("stor", "Stor", sum(1 for s in alle if s.get("klasse")=="Stor"), "#5B2D8E"),
    ]

    for i, (key, label, verdi, farge) in enumerate(filtre):
        with scol[i]:
            if st.button(f"{label}\n{verdi}", key=f"stat_{key}", use_container_width=True):
                st.session_state.screening_filter = key
                st.session_state.valgt_butikk = None
                st.rerun()

    # Søk og bransjefilter
    fcol1, fcol2, fcol3 = st.columns([2, 2, 3])
    with fcol1:
        klasse_f = st.selectbox("Klasse", ["Alle", "Liten", "Medium", "Stor"], key="skl")
    with fcol2:
        bransjer = ["Alle"] + sorted(set(b.get("bransje","Annet") for b in alle if b.get("bransje") and b.get("bransje") != "Annet"))
        bransje_f = st.selectbox("Bransje", bransjer, key="sbr")
    with fcol3:
        sok = st.text_input("🔍 Søk nettbutikk", placeholder="Skriv navn...", key="ssok")

    # Filtrer
    vis = alle
    cf = st.session_state.screening_filter
    if cf == "inn": vis = [s for s in vis if s.get("status")=="inn" and not s.get("enk")]
    elif cf == "ut": vis = [s for s in vis if s.get("status")=="ut" and not s.get("enk")]
    elif cf == "usikker": vis = [s for s in vis if s.get("status")=="usikker"]
    elif cf == "enk": vis = [s for s in vis if s.get("enk")]
    elif cf == "liten": vis = [s for s in vis if s.get("klasse")=="Liten"]
    elif cf == "medium": vis = [s for s in vis if s.get("klasse")=="Medium"]
    elif cf == "stor": vis = [s for s in vis if s.get("klasse")=="Stor"]
    if klasse_f != "Alle": vis = [s for s in vis if s.get("klasse")==klasse_f]
    if bransje_f != "Alle": vis = [s for s in vis if s.get("bransje")==bransje_f]
    if sok: vis = [s for s in vis if sok.lower() in s.get("name","").lower()]
    vis = sorted(vis, key=lambda x: x.get("total") or 0, reverse=True)

    st.markdown(f"**Viser {len(vis)} butikker**")

    # Detaljpanel
    juryvurderinger = hent_jury()
    if st.session_state.valgt_butikk and st.session_state.valgt_butikk in r:
        vis_detaljpanel(r[st.session_state.valgt_butikk], juryvurderinger)

    # Tabell
    st.markdown("---")
    for i, s in enumerate(vis):
        kol = st.columns([3, 2, 1, 1, 1, 1, 2, 3])
        with kol[0]:
            navn = s.get("name","")
            url = s.get("url","")
            url_tekst = url.replace("https://","").replace("http://","")[:30] if url else ""
            st.markdown(f'<div style="font-weight:600;font-size:13px">{navn}</div><div style="font-size:11px;color:#C8102E">{url_tekst}</div><div style="font-size:10px;color:#999">Klikk for detaljer ↓</div>', unsafe_allow_html=True)
        with kol[1]:
            st.markdown(f'<span style="background:#F0E8FA;color:#5B2D8E;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600">{s.get("bransje","–")}</span>', unsafe_allow_html=True)
        with kol[2]:
            st.markdown(f'<span style="font-size:11px;color:#666">{s.get("orgform","–")}</span>', unsafe_allow_html=True)
        with kol[3]:
            st.markdown(status_html(s.get("status"), s.get("enk")), unsafe_allow_html=True)
        with kol[4]:
            if s.get("klasse") and s["klasse"] != "-":
                st.markdown(klasse_html(s["klasse"]), unsafe_allow_html=True)
        with kol[5]:
            st.markdown(score_html(s.get("total")), unsafe_allow_html=True)
        with kol[6]:
            st.markdown(logi_html(s.get("logistikk")), unsafe_allow_html=True)
        with kol[7]:
            if st.button("Vis detaljer", key=f"det_{i}_{s.get('name','')}", use_container_width=True):
                st.session_state.valgt_butikk = s.get("name")
                st.rerun()
        st.markdown('<hr style="margin:4px 0;border-color:#E8E6E2">', unsafe_allow_html=True)

    # Eksporter
    if vis:
        st.markdown("---")
        if st.button("↓ Eksporter CSV"):
            df = pd.DataFrame([{
                "Nettbutikk": s.get("name"), "URL": s.get("url",""),
                "Bransje": s.get("bransje",""), "Org.form": s.get("orgform",""),
                "Status": s.get("status",""), "Klasse": s.get("klasse",""),
                "Omsetning": s.get("omsetning",""), "Total": s.get("total",""),
                "Inntrykk": s.get("inntrykk",""), "IKS": s.get("iks",""),
                "Kassen/Mersalg": s.get("kat3",""), "Markedsf.": s.get("markedsforing",""),
                "Begrunnelse": s.get("screeningBegrunnelse","")
            } for s in vis])
            st.download_button("Last ned", df.to_csv(index=False, sep=";").encode("utf-8-sig"),
                               "screening.csv", "text/csv")


# ─────────────────────────────────────────────
# SIDE 3 – TOPP 100
# ─────────────────────────────────────────────
elif side == "⭐ Topp 100":
    st.header("⭐ Topp 100 – Juryens arbeidsflate")
    st.markdown("Topp 33-34 per klasse basert på agentens score.")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå.")
        st.stop()

    r = st.session_state.resultater
    juryvurderinger = hent_jury()

    def hent_topp(klasse, n=34):
        return sorted(
            [v for v in r.values() if v.get("status")=="inn" and not v.get("enk")
             and v.get("klasse")==klasse and v.get("total") is not None],
            key=lambda x: x.get("total", 0), reverse=True
        )[:n]

    liten = hent_topp("Liten")
    medium = hent_topp("Medium")
    stor = hent_topp("Stor")
    alle_topp = liten + medium + stor

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Totalt i Topp 100", len(alle_topp))
    col2.metric("Liten", len(liten))
    col3.metric("Medium", len(medium))
    col4.metric("Stor", len(stor))

    klasse_tab = st.radio("Vis", ["Alle", "Liten", "Medium", "Stor"], horizontal=True)
    vis_liste = {"Liten": liten, "Medium": medium, "Stor": stor}.get(klasse_tab, alle_topp)

    # Detaljpanel
    if st.session_state.valgt_butikk and st.session_state.valgt_butikk in r:
        vis_detaljpanel(r[st.session_state.valgt_butikk], juryvurderinger)

    st.markdown("---")

    for i, butikk in enumerate(vis_liste):
        navn = butikk.get("name","")
        lagret = juryvurderinger.get(navn, {})
        js = lagret.get("Status", butikk.get("juryStatus","Ikke vurdert"))
        jsc = int(lagret.get("Juryscore", butikk.get("juryScore", 0)))
        farge_map = {"Finalist":"#B8860B","Vinner":"#C8102E","Semifinalist":"#1B6B3A"}
        border = farge_map.get(js, "#E8E6E2")

        with st.expander(f"**{i+1}.** {navn} — {score_html(butikk.get('total'))} {klasse_html(butikk.get('klasse',''))} {'⭐'*jsc if jsc else ''}", expanded=False):
            c1, c2, c3 = st.columns([3, 1, 1])
            with c1:
                if butikk.get("url"):
                    st.markdown(f'🌐 [{butikk["url"]}]({butikk["url"]})')
                st.caption(butikk.get("kommentar",""))
                st.markdown(logi_html(butikk.get("logistikk")), unsafe_allow_html=True)
            with c2:
                st.metric("Inntrykk", butikk.get("inntrykk","–"))
                st.metric("IKS", butikk.get("iks","–"))
            with c3:
                st.metric("Kassen", butikk.get("kat3","–"))
                st.metric("Markedsf.", butikk.get("markedsforing","–"))

            jc1, jc2, jc3 = st.columns([1, 1, 2])
            with jc1:
                ny_score = st.select_slider("Juryscore", [1,2,3,4,5], value=max(1,jsc), format_func=lambda x: "⭐"*x, key=f"ts_{navn}")
            with jc2:
                ny_status = st.selectbox("Status", ["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"],
                    index=["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"].index(js),
                    key=f"tstat_{navn}")
            with jc3:
                ny_notat = st.text_input("Notat", value=lagret.get("Notat",""), key=f"tnot_{navn}")

            col_a, col_b = st.columns([1, 3])
            with col_a:
                if st.button("💾 Lagre", key=f"tlagre_{navn}"):
                    lagre_jury(navn, ny_score, ny_status, ny_notat)
                    st.session_state.resultater[navn]["juryScore"] = ny_score
                    st.session_state.resultater[navn]["juryStatus"] = ny_status
                    st.success("Lagret!")
            with col_b:
                if st.button("🔍 Vis full detaljvisning", key=f"tdet_{navn}"):
                    st.session_state.valgt_butikk = navn
                    st.rerun()

    st.markdown("---")
    if st.button("↓ Eksporter Topp 100 CSV"):
        df = pd.DataFrame([{
            "Rang": i+1, "Navn": b.get("name"), "URL": b.get("url"),
            "Klasse": b.get("klasse"), "Bransje": b.get("bransje"),
            "Agentscore": b.get("total"), "Juryscore": b.get("juryScore",""),
            "Jurystatus": b.get("juryStatus",""), "Jurynotat": b.get("juryNote",""),
        } for i, b in enumerate(alle_topp)])
        st.download_button("Last ned", df.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           "topp100.csv", "text/csv")


# ─────────────────────────────────────────────
# SIDE 4 – FINALE
# ─────────────────────────────────────────────
elif side == "🏆 Finale":
    st.header("🏆 Finale – Juryens endelige vurderinger")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå.")
        st.stop()

    r = st.session_state.resultater
    juryvurderinger = hent_jury()

    finalister = []
    for navn, butikk in r.items():
        lagret = juryvurderinger.get(navn, {})
        status = lagret.get("Status", butikk.get("juryStatus","Ikke vurdert"))
        if status in ["Semifinalist","Finalist","Vinner"]:
            b = dict(butikk)
            b["juryStatus"] = status
            b["juryScore"] = int(lagret.get("Juryscore", butikk.get("juryScore",0)))
            b["juryNote"] = lagret.get("Notat", butikk.get("juryNote",""))
            finalister.append(b)

    if not finalister:
        st.info("Ingen butikker er merket som Semifinalist, Finalist eller Vinner ennå.")
        st.stop()

    rang = {"Vinner":0,"Finalist":1,"Semifinalist":2}
    finalister.sort(key=lambda x: (rang.get(x.get("juryStatus"),3), -(x.get("juryScore") or 0)))

    filter_val = st.radio("Vis", ["Alle","Vinnere","Finalister","Semifinalister"], horizontal=True)
    if filter_val == "Vinnere": finalister = [f for f in finalister if f.get("juryStatus")=="Vinner"]
    elif filter_val == "Finalister": finalister = [f for f in finalister if f.get("juryStatus")=="Finalist"]
    elif filter_val == "Semifinalister": finalister = [f for f in finalister if f.get("juryStatus")=="Semifinalist"]

    farge_map = {"Vinner":"#C8102E","Finalist":"#B8860B","Semifinalist":"#1B6B3A"}
    ikon_map = {"Vinner":"🏆","Finalist":"★","Semifinalist":"◐"}

    for i, b in enumerate(finalister):
        status = b.get("juryStatus","")
        farge = farge_map.get(status,"#666")
        ikon = ikon_map.get(status,"")
        stjerner = "⭐" * int(b.get("juryScore") or 0)

        st.markdown(f"""
        <div style="background:white;border-radius:10px;padding:18px;margin-bottom:12px;
                    border-left:5px solid {farge};box-shadow:0 1px 4px rgba(0,0,0,0.07)">
            <div style="display:flex;align-items:center;gap:14px">
                <span style="font-size:22px;font-weight:700;color:{farge}">{ikon} {i+1}</span>
                <div style="flex:1">
                    <div style="font-size:16px;font-weight:700">{b.get("name")}</div>
                    <div style="font-size:12px;color:#C8102E">{b.get("url","")}</div>
                </div>
                <div style="text-align:right">
                    <div style="font-weight:700;color:{farge};font-size:14px">{status}</div>
                    <div style="font-size:20px">{stjerner}</div>
                    <div>{score_html(b.get("total"))}</div>
                </div>
            </div>
            <div style="display:flex;gap:16px;font-size:12px;color:#666;margin-top:10px">
                <span>Klasse: <strong>{b.get("klasse")}</strong></span>
                <span>Bransje: <strong>{b.get("bransje","–")}</strong></span>
                <span>Inntrykk: <strong>{b.get("inntrykk","–")}</strong></span>
                <span>IKS: <strong>{b.get("iks","–")}</strong></span>
                <span>Kassen: <strong>{b.get("kat3","–")}</strong></span>
            </div>
            {f'<div style="margin-top:8px;font-size:13px;color:#666;font-style:italic">"{b.get("juryNote")}"</div>' if b.get("juryNote") else ""}
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    if st.button("↓ Eksporter finale CSV"):
        df = pd.DataFrame([{
            "Jurystatus": b.get("juryStatus"), "Navn": b.get("name"),
            "URL": b.get("url"), "Klasse": b.get("klasse"),
            "Agentscore": b.get("total"), "Juryscore": b.get("juryScore",""),
            "Jurynotat": b.get("juryNote",""),
        } for b in finalister])
        st.download_button("Last ned", df.to_csv(index=False, sep=";").encode("utf-8-sig"),
                           "finale.csv", "text/csv")


# ─────────────────────────────────────────────
# SIDE 5 – LOGISTIKK
# ─────────────────────────────────────────────
elif side == "📦 Logistikk":
    st.header("📦 Logistikkrapport – Posten Bring")

    if not st.session_state.resultater:
        st.info("Ingen resultater ennå.")
        st.stop()

    r = st.session_state.resultater
    med = [v for v in r.values() if v.get("status")=="inn" and v.get("logistikk")]

    if not med:
        st.info("Ingen logistikkdata tilgjengelig ennå.")
        st.stop()

    counts = {a: sum(1 for s in med if s.get("logistikk",{}).get(a)) for a in LOGISTIKK_AKTORER}

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Logistikkpartnere**")
        df_logi = pd.DataFrame([{"Partner": k, "Antall": v} for k, v in sorted(counts.items(), key=lambda x: x[1], reverse=True)])
        st.bar_chart(df_logi.set_index("Partner"))
    with c2:
        st.markdown("**Status Posten/Bring**")
        pb = counts.get("Posten/Bring", 0)
        st.metric("Bruker Posten/Bring", pb)
        st.metric("Bruker IKKE Posten/Bring", len(med) - pb)
        st.metric("Salgspotensial", len(med) - pb)

    st.markdown("---")
    st.markdown("### Salgsmuligheter")
    prospects = [s for s in med if not s.get("logistikk",{}).get("Posten/Bring")]
    prospects.sort(key=lambda x: x.get("total",0), reverse=True)
    if prospects:
        df = pd.DataFrame([{
            "Butikk": s.get("name"), "Klasse": s.get("klasse"),
            "Score": s.get("total"),
            "Bruker nå": ", ".join([a for a in LOGISTIKK_AKTORER if a!="Posten/Bring" and s.get("logistikk",{}).get(a)]) or "Ukjent"
        } for s in prospects])
        st.dataframe(df, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────
# SIDE 6 – OM VERKTØYET
# ─────────────────────────────────────────────
elif side == "ℹ️ Om verktøyet":
    st.header("ℹ️ Om Netthandelsprisen – Jurysystem")
    st.markdown("""
    ### Scoringsmodell v5

    | Kategori | Vekt | Kriterier |
    |---|---|---|
    | Første inntrykk | 25% | Startside (25%), Bilder/film (25%), Produktinfo (25%), Søk (25%) |
    | Info, kundeservice og bærekraft | 25% | KLR 35%, Kundeservice 35%, Bærekraft 30% |
    | Kassen, mersalg og inspirasjon | 25% | Kassen 50%, Mersalg 25%, Inspirasjon 25% |
    | Markedsføring og kundedialog | 25% | SoMe 40%, Kundeklubb 30% (dynamisk), Nyhetsbrev 30% |

    **Kjøpsvilkår, levering og retur (35% av IKS):**
    Kjøpsvilkår 20% · Levering 40% · Retur 40%

    **Kassen (50% av Kassen/Mersalg/Inspirasjon):**
    Innlogging 20% · Leveringsvalg 30% · Leveringspris 20% · Leveringstid 15% · Betaling 15%

    **Kundeklubb-scoring:**
    2 = Ingen kundeklubb · 3 = Kun inngangsrabatt · 4 = Poeng/rabatter · 5 = Full lojalitetspakke

    **NB:** Stor-klassen bedømmes strengere enn Medium og Liten.

    ### KO-kriterier
    - ENK bekreftet i Brønnøysund – filtreres automatisk ut
    - Ingen URL funnet – kan ikke identifiseres
    - B2C og skandinavisk tilstedeværelse vises som advarsel – juryen bestemmer

    ### Versjon 5.0 · August 2026 · Posten Bring
    """)
