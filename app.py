import streamlit as st
import json
import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

LOGISTIKK_AKTORER = [
    "Posten/Bring", "PostNord", "Helthjem", "Instabox",
    "Porterbuddy", "DHL", "Budbee", "UPS/FedEx", "Egne biler"
]

st.set_page_config(
    page_title="Netthandelsprisen – Posten Bring",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.stApp { background-color: #F5F4F2; }
section[data-testid="stSidebar"] > div { background-color: #FFFFFF !important; border-right: 1px solid #E8E6E2; }
.stTabs [data-baseweb="tab-panel"] { background-color: white; padding: 16px; border-radius: 0 0 8px 8px; }
div[data-testid="stExpander"] { background-color: white; border-radius: 8px; }
.main-header { background: #212121; color: white; padding: 16px 28px; border-radius: 10px; margin-bottom: 24px; display: flex; align-items: center; gap: 16px; }
.logo-badge { background: #C8102E; color: white; padding: 5px 12px; border-radius: 4px; font-weight: 700; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }
.detail-panel { background: white; border-radius: 12px; padding: 28px 32px; margin-bottom: 20px; border-top: 4px solid #C8102E; box-shadow: 0 2px 12px rgba(0,0,0,0.10); }
.jury-box { background: #F8F7F5; border-radius: 10px; padding: 16px 20px; margin: 16px 0; border-left: 4px solid #C8102E; }
.section-title { font-size: 11px; font-weight: 700; color: #999; text-transform: uppercase; letter-spacing: 0.6px; margin: 20px 0 10px 0; padding-bottom: 6px; border-bottom: 1.5px solid #E8E6E2; }
.warning-box { background: #FEF3E2; border-left: 3px solid #E8A020; border-radius: 0 8px 8px 0; padding: 10px 14px; font-size: 13px; color: #7A4800; margin: 10px 0; }
.tech-ok { background: #E6F4EC; color: #1B6B3A; padding: 6px 12px; border-radius: 8px; font-size: 13px; font-weight: 500; display: inline-block; margin: 3px; }
.tech-warn { background: #FEF3E2; color: #7A4800; padding: 6px 12px; border-radius: 8px; font-size: 13px; font-weight: 500; display: inline-block; margin: 3px; }
.tech-bad { background: #FDECEA; color: #C8102E; padding: 6px 12px; border-radius: 8px; font-size: 13px; font-weight: 500; display: inline-block; margin: 3px; }
.trust-badge { background: #E6F0FA; color: #0D4A8A; padding: 5px 10px; border-radius: 4px; font-size: 13px; font-weight: 500; display: inline-block; margin: 3px; }
.logi-box { border-radius: 8px; padding: 10px 8px; text-align: center; font-size: 12px; font-weight: 600; }
.logi-yes { background: #E6F4EC; color: #1B6B3A; }
.logi-no { background: #F5F4F2; color: #bbb; }
.crit-row { padding: 16px 0; border-bottom: 1px solid #E8E6E2; display: flex; gap: 16px; align-items: flex-start; }
.crit-name { font-size: 15px; font-weight: 700; color: #1A1A1A; margin-bottom: 6px; }
.crit-beg { font-size: 14px; color: #333; line-height: 1.7; }
.crit-vekt { font-size: 12px; color: #aaa; margin-top: 4px; font-style: italic; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <span class="logo-badge">Posten Bring</span>
    <div>
        <div style="font-size:19px;font-weight:700">Netthandelsprisen – Jurysystem</div>
        <div style="font-size:13px;color:#aaa">Automatisk screening og scoringsagent · 2026</div>
    </div>
</div>
""", unsafe_allow_html=True)

RESULTATER_FIL = "netthandelsprisen_resultater.json"

def lagre_lokalt(data):
    try:
        with open(RESULTATER_FIL, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def last_lokalt():
    try:
        if os.path.exists(RESULTATER_FIL):
            with open(RESULTATER_FIL, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

if "resultater" not in st.session_state:
    st.session_state.resultater = last_lokalt()
if "valgt_butikk" not in st.session_state:
    st.session_state.valgt_butikk = None
if "screening_filter" not in st.session_state:
    st.session_state.screening_filter = "alle"

def koble_sheets():
    try:
        creds_info = st.secrets["gcp_service_account"]
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_info, scopes=scope)
        gc = gspread.authorize(creds)
        return gc.open_by_key(st.secrets["google_sheets"]["sheet_id"]).sheet1
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

def hent_jury():
    ws = koble_sheets()
    if not ws:
        return {}
    try:
        return {r["Navn"]: r for r in ws.get_all_records() if r.get("Navn")}
    except Exception:
        return {}

def score_farge(s):
    if s is None: return "#999", "#eee"
    n = float(s)
    if n >= 4.5: return "#1B6B3A", "#E6F4EC"
    elif n >= 3.5: return "#0D4A8A", "#E6F0FA"
    elif n >= 2.5: return "#7A4800", "#FEF3E2"
    else: return "#C8102E", "#FDECEA"

def score_html(s):
    if s is None: return '<span style="color:#999">–</span>'
    c, bg = score_farge(s)
    return f'<span style="background:{bg};color:{c};padding:3px 10px;border-radius:5px;font-weight:700;font-size:14px">{float(s):.1f}</span>'

def klasse_html(k):
    m = {"Liten": ("#0D4A8A","#E6F0FA"), "Medium": ("#7A4800","#FEF3E2"), "Stor": ("#5B2D8E","#F0E8FA")}
    c, bg = m.get(k, ("#666","#eee"))
    return f'<span style="background:{bg};color:{c};padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700">{k}</span>'

def status_html(s, enk=False):
    if enk: return '<span style="background:#FDECEA;color:#C8102E;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700">⛔ ENK</span>'
    if s == "inn": return '<span style="background:#E6F4EC;color:#1B6B3A;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700">✓ Videre</span>'
    if s == "ut": return '<span style="background:#FDECEA;color:#C8102E;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700">✕ Ut</span>'
    return '<span style="background:#FEF3E2;color:#7A4800;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700">? Sjekk</span>'

with st.sidebar:
    st.markdown("### 🏆 Netthandelsprisen")
    st.markdown("---")
    side = st.radio("Naviger", ["📋 Screening", "⭐ Topp 100", "🏆 Finale", "📦 Logistikk", "ℹ️ Om verktøyet"], label_visibility="collapsed")
    st.markdown("---")
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if api_key:
        st.success("✅ API-nøkkel klar")
    st.markdown("---")
    st.markdown("**📂 Last opp resultater**")
    st.caption("Last opp resultater.json fra Colab")
    opplastet = st.file_uploader("resultater.json", type=["json"], key="json_upload", label_visibility="collapsed")
    if opplastet is not None:
        try:
            data = json.loads(opplastet.read().decode("utf-8"))
            if data and isinstance(data, dict):
                st.session_state.resultater = data
                lagre_lokalt(data)
                st.success(f"✅ {len(data)} butikker lastet inn!")
        except Exception as e:
            st.error(f"Feil: {e}")
    if st.session_state.resultater:
        r = st.session_state.resultater
        st.markdown("---")
        st.markdown("**Statistikk**")
        col1, col2 = st.columns(2)
        col1.metric("Totalt", len(r))
        col2.metric("Videre", sum(1 for v in r.values() if v.get("status")=="inn" and not v.get("enk")))
        col1.metric("Ut", sum(1 for v in r.values() if v.get("status")=="ut" and not v.get("enk")))
        col2.metric("ENK", sum(1 for v in r.values() if v.get("enk")))
        st.markdown("---")
        json_str = json.dumps(r, ensure_ascii=False, indent=2)
        st.download_button("↓ Last ned JSON", data=json_str.encode("utf-8"), file_name="netthandelsprisen_resultater.json", mime="application/json")
        if st.button("🗑️ Nullstill"):
            st.session_state.resultater = {}
            lagre_lokalt({})
            st.rerun()

def vis_detaljpanel(butikk, juryvurderinger={}):
    navn = butikk.get("name", "")
    lagret = juryvurderinger.get(navn, {})
    st.markdown('<div class="detail-panel">', unsafe_allow_html=True)
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown(f"## {navn}")
        if butikk.get("url"):
            st.markdown(f'<a href="{butikk["url"]}" target="_blank" style="background:#C8102E;color:#fff;padding:7px 16px;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600">🌐 Besøk nettbutikk</a>&nbsp;&nbsp;<span style="font-size:12px;color:#999">{butikk["url"].replace("https://","").replace("http://","")}</span>', unsafe_allow_html=True)
        st.markdown("")
        badges = status_html(butikk.get("status"), butikk.get("enk"))
        if butikk.get("klasse") and butikk["klasse"] not in ("-","Ukjent"):
            badges += " " + klasse_html(butikk["klasse"])
        badges += f' <span style="background:#E8E6E2;color:#666;padding:3px 10px;border-radius:20px;font-size:12px">{butikk.get("orgform","–")}</span>'
        if butikk.get("bransje"):
            badges += f' <span style="background:#F0E8FA;color:#5B2D8E;padding:3px 10px;border-radius:4px;font-size:12px;font-weight:600">{butikk["bransje"]}</span>'
        badges += f' <span style="background:#E8E6E2;color:#666;padding:3px 10px;border-radius:20px;font-size:12px">{butikk.get("omsetning","–")}</span>'
        st.markdown(badges, unsafe_allow_html=True)
    with col2:
        if butikk.get("total") is not None:
            c, bg = score_farge(butikk["total"])
            st.markdown(f'<div style="text-align:center;background:{bg};border-radius:10px;padding:16px"><div style="font-size:11px;color:#999;font-weight:600;text-transform:uppercase;letter-spacing:0.5px">Totalpoeng</div><div style="font-size:42px;font-weight:700;color:{c}">{butikk["total"]}</div></div>', unsafe_allow_html=True)

    for advarsel in butikk.get("advarsler", []):
        st.markdown(f'<div class="warning-box">🔍 {advarsel}</div>', unsafe_allow_html=True)
    ap = butikk.get("apenhetsloven", {})
    if ap and ap.get("palagt") and not ap.get("rapport_funnet"):
        st.markdown(f'<div class="warning-box">⚠️ Åpenhetsloven: Rapporteringspliktig men ingen rapport funnet.</div>', unsafe_allow_html=True)

    st.markdown('<div class="jury-box"><div style="font-size:11px;font-weight:700;color:#999;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:12px">🎯 Juryverktøy</div>', unsafe_allow_html=True)
    jcol1, jcol2, jcol3 = st.columns([2, 2, 3])
    with jcol1:
        jury_score = st.select_slider("Juryscore", options=[1,2,3,4,5], value=max(1, min(5, int(lagret.get("Juryscore", butikk.get("juryScore", 3)) or 3))), format_func=lambda x: "⭐"*x, key=f"js_{navn}")
    with jcol2:
        js_val = lagret.get("Status", butikk.get("juryStatus", "Ikke vurdert"))
        if js_val not in ["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"]:
            js_val = "Ikke vurdert"
        jury_status = st.selectbox("Status", ["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"], index=["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"].index(js_val), key=f"jstat_{navn}")
    with jcol3:
        jury_notat = st.text_area("Jurynotat", value=lagret.get("Notat", butikk.get("juryNote","")), placeholder="Skriv jurynotat her...", height=90, key=f"jnot_{navn}")
    if st.button("💾 Lagre vurdering", key=f"jlagre_{navn}"):
        lagre_jury(navn, jury_score, jury_status, jury_notat)
        st.session_state.resultater[navn]["juryScore"] = jury_score
        st.session_state.resultater[navn]["juryStatus"] = jury_status
        st.session_state.resultater[navn]["juryNote"] = jury_notat
        lagre_lokalt(st.session_state.resultater)
        st.success("✅ Lagret!")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(f'<div style="background:white;border:1.5px solid #E8E6E2;border-radius:8px;padding:12px 16px;border-left:3px solid #C8102E;font-size:13px;color:#666;margin:12px 0"><strong style="font-size:11px;color:#999;text-transform:uppercase;letter-spacing:0.5px">Screeningresultat</strong><br/>{butikk.get("screeningBegrunnelse","")}</div>', unsafe_allow_html=True)

    if butikk.get("kommentar"):
        st.markdown(f'<div style="background:white;border:1.5px solid #E8E6E2;border-radius:8px;padding:14px 18px;font-size:14px;color:#333;line-height:1.7;margin:12px 0">{butikk["kommentar"]}</div>', unsafe_allow_html=True)

    if butikk.get("total") is not None:
        st.markdown('<div class="section-title">Kategoriscorer</div>', unsafe_allow_html=True)
        scol = st.columns(5)
        kategorier = [
            ("Total", butikk.get("total"), "#212121"),
            ("Første inntrykk", butikk.get("inntrykk"), "#C8102E"),
            ("Info/KS/BK", butikk.get("iks"), "#E87D3E"),
            ("Kassen/Mersalg", butikk.get("kat3"), "#0D4A8A"),
            ("Markedsf.", butikk.get("markedsforing"), "#1B6B3A"),
        ]
        for i, (label, val, farge) in enumerate(kategorier):
            c, bg = score_farge(val)
            sub = ""
            if label == "Info/KS/BK" and butikk.get("iksDetalj"):
                d = butikk["iksDetalj"]
                sub = f'<div style="font-size:11px;color:#999;margin-top:4px">KLR:{d.get("klr","–")} KS:{d.get("kundeservice","–")} BK:{d.get("baerekraft","–")}</div>'
            if label == "Kassen/Mersalg" and butikk.get("kat3Detalj"):
                d = butikk["kat3Detalj"]
                sub = f'<div style="font-size:11px;color:#999;margin-top:4px">K:{d.get("kassen","–")} M:{d.get("mersalg","–")} I:{d.get("inspirasjon","–")}</div>'
            scol[i].markdown(
                f'<div style="background:white;border-radius:10px;padding:16px 12px;text-align:center;border-top:4px solid {farge};box-shadow:0 1px 4px rgba(0,0,0,0.08);min-height:110px">'
                f'<div style="font-size:11px;color:#999;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px">{label}</div>'
                f'<div style="font-size:32px;font-weight:700;color:{c if label!="Total" else "#212121"};line-height:1">{val if val is not None else "–"}</div>'
                f'{sub}</div>',
                unsafe_allow_html=True
            )

    st.markdown('<div class="section-title">Teknisk kvalitet og mobilopplevelse</div>', unsafe_allow_html=True)
    tech = butikk.get("tech") or {}
    if tech:
        def tbadge(v, l):
            cls = "tech-ok" if v == "ok" else "tech-warn" if v == "warn" else "tech-bad"
            ikon = "✓" if v == "ok" else "⚠" if v == "warn" else "✕"
            return f'<span class="{cls}">{ikon} {l}</span>'
        tp_raw = tech.get("trustpilot", "Ikke funnet")
        if isinstance(tp_raw, dict):
            tp_html = f'<span class="tech-ok">⭐ Trustpilot {tp_raw.get("score","")}/5 – {str(tp_raw.get("begrunnelse",""))[:60]}</span>'
        elif tp_raw and tp_raw != "Ikke funnet":
            tp_html = f'<span class="tech-ok">⭐ {tp_raw}</span>'
        else:
            tp_html = '<span class="tech-warn">⭐ Trustpilot ikke funnet</span>'
        st.markdown(tbadge(tech.get("mobil","warn"), "Mobilvennlig") + " " + tbadge(tech.get("ssl","warn"), "SSL") + " " + tbadge(tech.get("lastetid","warn"), "Lastetid") + " " + tp_html, unsafe_allow_html=True)
    else:
        st.caption("Ikke sjekket")

    trust = butikk.get("trust", [])
    if trust:
        st.markdown('<div class="section-title">Tillit og trygghet</div>', unsafe_allow_html=True)
        st.markdown(" ".join(f'<span class="trust-badge">{t}</span>' for t in trust), unsafe_allow_html=True)

    st.markdown('<div class="section-title">Logistikkpartnere</div>', unsafe_allow_html=True)
    logi = butikk.get("logistikk") or {}
    lcols = st.columns(len(LOGISTIKK_AKTORER))
    noen_funnet = False
    for j, aktør in enumerate(LOGISTIKK_AKTORER):
        val = logi.get(aktør, False)
        har = bool(val) and val not in (False, "false", 0, None, "")
        if har:
            noen_funnet = True
        css = "logi-yes" if har else "logi-no"
        ikon = "● " if har else ""
        lcols[j].markdown(f'<div class="logi-box {css}">{ikon}{aktør}</div>', unsafe_allow_html=True)
    if not noen_funnet:
        st.caption("ℹ️ Ingen logistikkpartnere identifisert denne kjøringen.")

    if butikk.get("scoring"):
        st.markdown('<div class="section-title">Kriteriegjennomgang</div>', unsafe_allow_html=True)
        scoring = butikk["scoring"]
        tab1, tab2, tab3, tab4 = st.tabs([
            "🔍 Første inntrykk", "📋 Info/KS/Bærekraft", "🛒 Kassen/Mersalg", "📣 Markedsføring"
        ])

        def vis_kriterier(tab, kriterier, kat_score, farge):
            with tab:
                if kat_score is not None:
                    c, bg = score_farge(kat_score)
                    st.markdown(
                        f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;padding-bottom:10px;border-bottom:2px solid {farge}">'
                        f'<span style="width:10px;height:10px;border-radius:50%;background:{farge};display:inline-block"></span>'
                        f'<span style="font-size:13px;color:#999">Vekt 25%</span>'
                        f'<span style="background:{bg};color:{c};padding:3px 10px;border-radius:5px;font-weight:700;font-size:15px">{kat_score}</span>'
                        f'</div>', unsafe_allow_html=True)
                for k in kriterier:
                    score = k.get("score", 0)
                    c, bg = score_farge(score)
                    st.markdown(f"""
                    <div class="crit-row">
                        <div style="min-width:52px;text-align:center">
                            <span style="background:{bg};color:{c};padding:8px 12px;border-radius:8px;font-weight:700;font-size:20px;display:inline-block;min-width:44px;text-align:center">{score}</span>
                        </div>
                        <div style="flex:1">
                            <div class="crit-name">{k.get("navn","")}</div>
                            <div class="crit-beg">{k.get("begrunnelse","")}</div>
                            <div class="crit-vekt">{k.get("vekt","")}</div>
                        </div>
                    </div>""", unsafe_allow_html=True)

        vis_kriterier(tab1, scoring.get("inntrykk", []), butikk.get("inntrykk"), "#C8102E")
        vis_kriterier(tab2, scoring.get("iks", []), butikk.get("iks"), "#E87D3E")
        vis_kriterier(tab3, scoring.get("kassen", []), butikk.get("kat3"), "#0D4A8A")
        vis_kriterier(tab4, scoring.get("markedsforing", []), butikk.get("markedsforing"), "#1B6B3A")

    if st.button("✕ Lukk detaljer", key=f"lukk_{navn}"):
        st.session_state.valgt_butikk = None
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


if side == "📋 Screening":
    st.header("📋 Screening – alle butikker")
    if not st.session_state.resultater:
        st.info("💡 Last opp resultater.json fra Colab i sidepanelet til venstre.")
        st.stop()
    r = st.session_state.resultater
    alle = list(r.values())
    filtre_def = [
        ("alle","Totalt",len(alle),"#1A1A1A"),
        ("inn","Går videre",sum(1 for s in alle if s.get("status")=="inn" and not s.get("enk")),"#1B6B3A"),
        ("ut","Filtrert ut",sum(1 for s in alle if s.get("status")=="ut" and not s.get("enk")),"#C8102E"),
        ("usikker","Krever sjekk",sum(1 for s in alle if s.get("status")=="usikker"),"#7A4800"),
        ("enk","ENK",sum(1 for s in alle if s.get("enk")),"#C8102E"),
        ("liten","Liten",sum(1 for s in alle if s.get("klasse")=="Liten"),"#0D4A8A"),
        ("medium","Medium",sum(1 for s in alle if s.get("klasse")=="Medium"),"#7A4800"),
        ("stor","Stor",sum(1 for s in alle if s.get("klasse")=="Stor"),"#5B2D8E"),
    ]
    cols = st.columns(8)
    for i, (key, label, verdi, farge) in enumerate(filtre_def):
        with cols[i]:
            if st.button(f"{label}\n{verdi}", key=f"stat_{key}", use_container_width=True):
                st.session_state.screening_filter = key
                st.session_state.valgt_butikk = None
                st.rerun()
    fcol1, fcol2, fcol3 = st.columns([2,2,3])
    with fcol1:
        klasse_f = st.selectbox("Klasse", ["Alle","Liten","Medium","Stor"])
    with fcol2:
        bransjer = ["Alle"] + sorted(set(b.get("bransje","Annet") for b in alle if b.get("bransje") and b.get("bransje") != "Annet"))
        bransje_f = st.selectbox("Bransje", bransjer)
    with fcol3:
        sok = st.text_input("🔍 Søk nettbutikk", placeholder="Skriv navn...")
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
    juryvurderinger = hent_jury()
    if st.session_state.valgt_butikk and st.session_state.valgt_butikk in r:
        vis_detaljpanel(r[st.session_state.valgt_butikk], juryvurderinger)
    st.markdown("---")
    for i, s in enumerate(vis):
        tcol = st.columns([3,2,1,1,1,1,2,2])
        navn = s.get("name","")
        url = s.get("url","")
        with tcol[0]:
            st.markdown(f'<div style="font-weight:700;font-size:16px">{navn}</div><div style="font-size:13px;color:#C8102E">{url.replace("https://","").replace("http://","")[:35] if url else ""}</div>', unsafe_allow_html=True)
        with tcol[1]:
            st.markdown(f'<span style="background:#F0E8FA;color:#5B2D8E;padding:3px 8px;border-radius:4px;font-size:11px;font-weight:600">{s.get("bransje","–")}</span>', unsafe_allow_html=True)
        with tcol[2]:
            st.markdown(f'<span style="font-size:12px;color:#666">{s.get("orgform","–")}</span>', unsafe_allow_html=True)
        with tcol[3]:
            st.markdown(status_html(s.get("status"), s.get("enk")), unsafe_allow_html=True)
        with tcol[4]:
            if s.get("klasse") and s["klasse"] not in ("-","Ukjent"):
                st.markdown(klasse_html(s["klasse"]), unsafe_allow_html=True)
        with tcol[5]:
            st.markdown(score_html(s.get("total")), unsafe_allow_html=True)
        with tcol[6]:
            logi = s.get("logistikk") or {}
            treff = [a for a in LOGISTIKK_AKTORER if bool(logi.get(a)) and logi.get(a) not in (False,"false",0,None,"")]
            kort = {"Posten/Bring":"P/B","PostNord":"PN","Helthjem":"HH","Instabox":"IB","Porterbuddy":"PB","DHL":"DHL","Budbee":"BB","UPS/FedEx":"UPS","Egne biler":"Egne"}
            if treff:
                st.markdown(" ".join(f'<span style="background:#E6F4EC;color:#1B6B3A;border-radius:4px;padding:2px 6px;font-size:11px;font-weight:600">{kort.get(a,a)}</span>' for a in treff), unsafe_allow_html=True)
            else:
                st.markdown('<span style="color:#ccc;font-size:11px">–</span>', unsafe_allow_html=True)
        with tcol[7]:
            label = "▼ Åpnet" if st.session_state.valgt_butikk == navn else "Vis detaljer"
            if st.button(label, key=f"det_{i}_{navn}", use_container_width=True):
                if st.session_state.valgt_butikk == navn:
                    st.session_state.valgt_butikk = None
                else:
                    st.session_state.valgt_butikk = navn
                st.rerun()
        st.markdown('<hr style="margin:4px 0;border-color:#D8D6D2;opacity:0.4">', unsafe_allow_html=True)
    if vis:
        st.markdown("---")
        df = pd.DataFrame([{"Nettbutikk": s.get("name"), "URL": s.get("url",""), "Bransje": s.get("bransje",""), "Status": s.get("status",""), "Klasse": s.get("klasse",""), "Total": s.get("total","")} for s in vis])
        st.download_button("↓ Eksporter CSV", df.to_csv(index=False, sep=";").encode("utf-8-sig"), "screening.csv", "text/csv")

elif side == "⭐ Topp 100":
    st.header("⭐ Topp 100 – Juryens arbeidsflate")
    if not st.session_state.resultater:
        st.info("Last opp resultater.json i sidepanelet.")
        st.stop()
    r = st.session_state.resultater
    juryvurderinger = hent_jury()
    def hent_topp(klasse, n=34):
        return sorted([v for v in r.values() if v.get("status")=="inn" and not v.get("enk") and v.get("klasse")==klasse and v.get("total") is not None], key=lambda x: x.get("total",0), reverse=True)[:n]
    liten = hent_topp("Liten")
    medium = hent_topp("Medium")
    stor = hent_topp("Stor")
    alle_topp = liten + medium + stor
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Totalt i Topp 100", len(alle_topp))
    col2.metric("Liten", len(liten))
    col3.metric("Medium", len(medium))
    col4.metric("Stor", len(stor))
    klasse_tab = st.radio("Vis", ["Alle","Liten","Medium","Stor"], horizontal=True)
    vis_liste = {"Liten":liten,"Medium":medium,"Stor":stor}.get(klasse_tab, alle_topp)
    if st.session_state.valgt_butikk and st.session_state.valgt_butikk in r:
        vis_detaljpanel(r[st.session_state.valgt_butikk], juryvurderinger)
    st.markdown("---")
    for i, butikk in enumerate(vis_liste):
        navn = butikk.get("name","")
        lagret = juryvurderinger.get(navn, {})
        js = lagret.get("Status", butikk.get("juryStatus","Ikke vurdert"))
        jsc = int(lagret.get("Juryscore", butikk.get("juryScore",0)) or 0)
        with st.expander(f"**{i+1}. {navn}** — {butikk.get('total','–')} | {butikk.get('klasse','')} | {butikk.get('bransje','–')} {'⭐'*jsc if jsc else ''}"):
            c1, c2, c3 = st.columns([3,1,1])
            with c1:
                if butikk.get("url"): st.markdown(f'🌐 [{butikk["url"]}]({butikk["url"]})')
                st.caption(butikk.get("kommentar",""))
            with c2:
                st.metric("Inntrykk", butikk.get("inntrykk","–"))
                st.metric("IKS", butikk.get("iks","–"))
            with c3:
                st.metric("Kassen", butikk.get("kat3","–"))
                st.metric("Markedsf.", butikk.get("markedsforing","–"))
            jc1, jc2, jc3 = st.columns([1,1,2])
            with jc1:
                ny_score = st.select_slider("Juryscore", [1,2,3,4,5], value=max(1,min(5,jsc or 3)), format_func=lambda x: "⭐"*x, key=f"ts_{navn}")
            with jc2:
                if js not in ["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"]: js = "Ikke vurdert"
                ny_status = st.selectbox("Status", ["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"], index=["Ikke vurdert","Kandidat","Semifinalist","Finalist","Vinner"].index(js), key=f"tstat_{navn}")
            with jc3:
                ny_notat = st.text_input("Notat", value=lagret.get("Notat",""), key=f"tnot_{navn}")
            ca, cb = st.columns([1,3])
            with ca:
                if st.button("💾 Lagre", key=f"tlagre_{navn}"):
                    lagre_jury(navn, ny_score, ny_status, ny_notat)
                    st.session_state.resultater[navn]["juryScore"] = ny_score
                    st.session_state.resultater[navn]["juryStatus"] = ny_status
                    st.success("Lagret!")
            with cb:
                if st.button("🔍 Vis full detalj", key=f"tdet_{navn}"):
                    st.session_state.valgt_butikk = navn
                    st.rerun()

elif side == "🏆 Finale":
    st.header("🏆 Finale – Juryens endelige vurderinger")
    if not st.session_state.resultater:
        st.info("Last opp resultater.json i sidepanelet.")
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
            b["juryScore"] = int(lagret.get("Juryscore", butikk.get("juryScore",0)) or 0)
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
        st.markdown(f'<div style="background:white;border-radius:12px;padding:20px 24px;margin-bottom:14px;border-left:5px solid {farge};box-shadow:0 1px 6px rgba(0,0,0,0.08)"><div style="display:flex;align-items:center;gap:16px"><span style="font-size:24px;font-weight:700;color:{farge}">{ikon} {i+1}</span><div style="flex:1"><div style="font-size:17px;font-weight:700">{b.get("name")}</div><div style="font-size:13px;color:#C8102E">{b.get("url","")}</div></div><div style="text-align:right"><div style="font-weight:700;color:{farge};font-size:15px">{status}</div><div style="font-size:22px">{stjerner}</div></div></div></div>', unsafe_allow_html=True)

elif side == "📦 Logistikk":
    st.header("📦 Logistikkrapport – Posten Bring")
    if not st.session_state.resultater:
        st.info("Last opp resultater.json i sidepanelet.")
        st.stop()
    r = st.session_state.resultater
    med = [v for v in r.values() if v.get("status")=="inn"]
    counts = {a: sum(1 for s in med if bool((s.get("logistikk") or {}).get(a)) and (s.get("logistikk") or {}).get(a) not in (False,"false",0,None,"")) for a in LOGISTIKK_AKTORER}
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Logistikkpartnere**")
        df_logi = pd.DataFrame([{"Partner": k, "Antall": v} for k, v in sorted(counts.items(), key=lambda x: x[1], reverse=True)])
        st.bar_chart(df_logi.set_index("Partner"))
    with c2:
        pb = counts.get("Posten/Bring", 0)
        st.metric("Bruker Posten/Bring", pb)
        st.metric("Bruker IKKE Posten/Bring", len(med)-pb)

elif side == "ℹ️ Om verktøyet":
    st.header("ℹ️ Om Netthandelsprisen – Jurysystem")
    st.markdown("""
    ### Scoringsmodell
    | Kategori | Vekt | Kriterier |
    |---|---|---|
    | Første inntrykk | 25% | Startside, bilder/film, produktinfo, søk |
    | Info, kundeservice og bærekraft | 25% | KLR 35%, Kundeservice 35%, Bærekraft 30% |
    | Kassen, mersalg og inspirasjon | 25% | Kassen 50%, Mersalg 25%, Inspirasjon 25% |
    | Markedsføring og kundedialog | 25% | SoMe 40%, Kundeklubb 30%, Nyhetsbrev 30% |

    **Kundeklubb:** 2=Ingen · 3=Kun inngangsrabatt · 4=Poeng/rabatter · 5=Full lojalitetspakke

    ### Versjon 7.0 · August 2026 · Posten Bring
    """)
