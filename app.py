import io
import time
import zipfile
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import requests
import streamlit as st
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook
from dotenv import load_dotenv
from google import genai

load_dotenv()

MODELI = ["gemini-3.5-flash-lite", "gemini-3.6-flash", "gemini-3.8-flash"]
GLAVA = {"User-Agent": "Mozilla/5.0"}
GLAVA_API = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
SLOVENSKI_CAS = ZoneInfo("Europe/Ljubljana")
NAJVEC_VRSTIC_EXCEL = 300  # iz vsakega lista preberemo največ toliko vrstic

NAVODILO = """Si pomočnik za javna naročila v Sloveniji.
Iz spodnje razpisne dokumentacije v slovenščini izpiši:

Pogoji za sodelovanje (alineje):
Merila za izbor ponudbe:
Kaj mora ponudnik oddati (alineje):
Finančna zavarovanja in garancije:
Rok dobave oziroma izvedbe:
Plačilni pogoji:

Če kakšnega podatka v dokumentaciji ni, napiši "ni navedeno". Ničesar si ne izmišljuj.
Odgovor začni neposredno s prvim naslovom, brez uvodnega stavka in brez zaključka.
Piši stvarno in neosebno, ne nagovarjaj bralca.

Dokumentacija:
"""


# ---------- Branje objave ----------

@st.cache_data(ttl=600)
def preberi_objavo(povezava):
    id_objave = povezava.strip().rstrip("/").split("/")[-1]
    if not id_objave.isdigit():
        return None
    url = "https://www.enarocanje.si/api/obrazec/objava/obrazecGet?id=" + id_objave
    odziv = requests.get(url, headers=GLAVA_API, timeout=30)
    odziv.raise_for_status()
    return odziv.json()


def lep_datum(besedilo):
    # Izpiše datum roka in doda, koliko dni je še do njega (ali da je že potekel)
    if not besedilo:
        return "ni navedeno"

    rok = datetime.fromisoformat(besedilo)
    if rok.tzinfo is not None:
        rok = rok.astimezone(SLOVENSKI_CAS).replace(tzinfo=None)
    zdaj = datetime.now(SLOVENSKI_CAS).replace(tzinfo=None)

    izpis = rok.strftime("%d. %m. %Y ob %H:%M")

    if rok < zdaj:
        return izpis + " — ⚠️ **poteklo**"

    dni = (rok.date() - zdaj.date()).days
    if dni == 0:
        return izpis + " — **danes**"
    if dni == 1:
        return izpis + " — še 1 dan"
    if dni == 2:
        return izpis + " — še 2 dneva"
    return izpis + " — še " + str(dni) + " dni"


# ---------- Iskanje razpisov ----------

@st.cache_data(ttl=600)
def poisci_razpise(beseda):
    danes = datetime.now()
    telo = {
        "page": 1,
        "idSifCpv": [],
        "idSifPostopekFaza": [],
        "podrejeniCpv": True,
        "datumDd": 3,
        "objavaDejanskaDatumOd": (danes - timedelta(days=90)).strftime("%Y-%m-%dT00:00:00.000Z"),
        "startRow": 0,
        "endRow": 20,
        "sortModel": [{"colId": "objavaDejanskaDatum", "sort": "desc"}],
        "datumPonudbaDo": None,
        "datumPonudbaOd": danes.strftime("%Y-%m-%dT00:00:00.000Z"),
        "naslov": beseda,
    }
    url = "https://www.enarocanje.si/api/obrazec/objava/obrazecGetGrid"
    odziv = requests.post(url, json=telo, headers=GLAVA_API, timeout=30)
    odziv.raise_for_status()
    return odziv.json()["data"]


def izberi_objavo(id_objave):
    st.session_state["povezava"] = "https://www.enarocanje.si/#/pregled-objav/" + str(id_objave)


def nazaj_na_iskanje():
    st.session_state["povezava"] = ""


# ---------- Branje dokumentov ----------

def preberi_pdf(datoteka):
    reader = PdfReader(datoteka)
    besedilo = ""
    for stran in reader.pages:
        besedilo += stran.extract_text() + "\n"
    return besedilo


def preberi_word(datoteka):
    dokument = Document(datoteka)
    besedilo = ""
    for odstavek in dokument.paragraphs:
        besedilo += odstavek.text + "\n"
    for tabela in dokument.tables:
        for vrstica in tabela.rows:
            for celica in vrstica.cells:
                besedilo += celica.text + " | "
            besedilo += "\n"
    return besedilo


def preberi_excel(datoteka):
    # data_only=True pomeni: preberi izračunane vrednosti, ne formul
    zvezek = load_workbook(datoteka, data_only=True, read_only=True)
    besedilo = ""
    for list_ in zvezek.worksheets:
        besedilo += "--- List: " + list_.title + " ---\n"
        stevilo_vrstic = 0
        for vrstica in list_.iter_rows(values_only=True):
            celice = [str(celica) for celica in vrstica if celica is not None and str(celica).strip() != ""]
            if not celice:
                continue  # prazno vrstico preskočimo
            besedilo += " | ".join(celice) + "\n"
            stevilo_vrstic += 1
            if stevilo_vrstic >= NAJVEC_VRSTIC_EXCEL:
                besedilo += "(list ima še več vrstic, ki niso prebrane)\n"
                break
    zvezek.close()
    return besedilo


def preberi_datoteko(ime, datoteka):
    # Izbere pravo branje glede na končnico. Če vrste datoteke ne poznamo, vrne None.
    malo = ime.lower()
    if malo.endswith(".docx"):
        return preberi_word(datoteka)
    if malo.endswith(".pdf"):
        return preberi_pdf(datoteka)
    if malo.endswith(".xlsx") or malo.endswith(".xlsm"):
        return preberi_excel(datoteka)
    return None


def preberi_zip(url):
    # Prenesi datoteko s portala
    odziv = requests.get(url, headers=GLAVA, timeout=60)
    odziv.raise_for_status()
    vsebina = odziv.content

    # Če dokumentacija ni ZIP, ampak kar en PDF
    if vsebina.startswith(b"%PDF"):
        return preberi_pdf(io.BytesIO(vsebina)), ["dokumentacija.pdf"], []

    # ZIP datoteke se vedno začnejo s črkama PK
    if not vsebina.startswith(b"PK"):
        return "", [], []

    arhiv = zipfile.ZipFile(io.BytesIO(vsebina))
    besedilo = ""
    prebrane = []
    neprebrane = []

    # Samo datoteke, brez map (imena map se končajo s /)
    imena = [ime for ime in arhiv.namelist() if not ime.endswith("/")]

    # Imena Word datotek brez končnice, da prepoznamo dvojnike v PDF
    wordi = {ime.lower()[:-5] for ime in imena if ime.lower().endswith(".docx")}

    for ime in imena:
        if ime.lower().endswith(".pdf") and ime.lower()[:-4] in wordi:
            continue  # isti dokument že imamo v Wordu, PDF preskočimo

        try:
            prebrano = preberi_datoteko(ime, io.BytesIO(arhiv.read(ime)))
        except Exception:
            prebrano = None  # datoteke se ni dalo prebrati

        if prebrano is None:
            neprebrane.append(ime)  # npr. XML, slike, stari .xls ali poškodovana datoteka
        else:
            besedilo += "\n\n=== " + ime + " ===\n" + prebrano
            prebrane.append(ime)

    return besedilo, prebrane, neprebrane


# ---------- AI ----------

def naredi_povzetek(besedilo):
    client = genai.Client(http_options={"timeout": 90000})
    sporocilo = st.empty()
    sporocilo.info("Priprava povzetka v teku ...")
    zadnja_napaka = None

    # Vsak model poskusimo dvakrat, po vrsti
    poskusi = []
    for model in MODELI:
        poskusi += [model, model]

    for stevilka, model in enumerate(poskusi):
        try:
            odgovor = client.models.generate_content(model=model, contents=NAVODILO + besedilo)
            sporocilo.empty()
            return odgovor.text
        except Exception as napaka:
            zadnja_napaka = napaka
            if stevilka < len(poskusi) - 1:
                time.sleep(10)  # po zadnjem poskusu ne čakamo več

    sporocilo.empty()
    st.caption("Tehnične podrobnosti: " + str(zadnja_napaka))
    return None


def pripravi_in_shrani(kljuc, besedilo, prebrane, neprebrane):
    # Naredi povzetek in ga shrani v "spomin" strani, da ob naslednjem kliku ne izgine
    povzetek = naredi_povzetek(besedilo)
    if povzetek:
        st.session_state["povzetki"][kljuc] = {
            "povzetek": povzetek,
            "prebrane": prebrane,
            "neprebrane": neprebrane,
        }
    else:
        # Povzetka ni, vseeno pa pokažemo, katere datoteke so bile prebrane
        if neprebrane:
            st.warning("Te datoteke niso bile prebrane, preglejte jih sami: " + ", ".join(neprebrane))
        if prebrane:
            st.write("**Prebrani dokumenti:**", ", ".join(prebrane))
        st.error("Povzetka trenutno ni mogoče pripraviti. Poskusite znova čez nekaj minut.")


def pokazi_povzetek(kljuc):
    shranjeno = st.session_state["povzetki"].get(kljuc)
    if shranjeno is None:
        return
    if shranjeno["neprebrane"]:
        st.warning("Te datoteke niso bile prebrane, preglejte jih sami: " + ", ".join(shranjeno["neprebrane"]))
    if shranjeno["prebrane"]:
        st.write("**Prebrani dokumenti:**", ", ".join(shranjeno["prebrane"]))
    st.markdown(shranjeno["povzetek"])


# ---------- Izgled strani ----------

st.title("📄 Povzetovalnik razpisov")

# Pripravi "spomin" strani
if "povezava" not in st.session_state:
    st.session_state["povezava"] = ""
if "beseda" not in st.session_state:
    st.session_state["beseda"] = ""
if "povzetki" not in st.session_state:
    st.session_state["povzetki"] = {}
st.session_state["beseda"] = st.session_state["beseda"]  # da si stran zapomni iskano besedo

if st.session_state["povezava"] == "":
    # Nič še ni izbrano: pokaži iskanje
    st.subheader("🔎 Iskanje odprtih razpisov")
    beseda = st.text_input("Ključna beseda v nazivu naročila (npr. vzdrževanje)", key="beseda")

    if beseda:
        try:
            zadetki = poisci_razpise(beseda)
        except Exception as napaka:
            st.error("Iskanje ni uspelo: " + str(napaka))
            zadetki = []

        if not zadetki:
            st.info("Ni odprtih razpisov s to besedo.")

        for zadetek in zadetki:
            levo, desno = st.columns([5, 1])
            objavljeno = datetime.fromisoformat(zadetek["objavaDejanskaDatum"][:10]).strftime("%d. %m. %Y")
            levo.write("**" + zadetek["naslov"] + "**  \n" + zadetek["narocnikNaziv"] + " · objavljeno " + objavljeno)
            desno.button("Povzemi", key=zadetek["idObrazec"], on_click=izberi_objavo, args=(zadetek["idObrazec"],))

    st.divider()
else:
    # Razpis je izbran: skrij seznam in ponudi pot nazaj
    st.button("← Nazaj na iskanje", on_click=nazaj_na_iskanje)

povezava = st.text_input("Povezava do objave na enarocanje.si", key="povezava")

if povezava:
    try:
        objava = preberi_objavo(povezava)
    except Exception as napaka:
        st.write("Podrobnosti napake:", str(napaka))
        objava = None

    if objava is None:
        st.error("Objave ni bilo mogoče prebrati. Preverite povezavo.")
    else:
        st.subheader(objava["naslov"])
        if objava["narocnik"]:
            st.write("**Naročnik:**", objava["narocnik"][0]["naziv"])
        st.write("**Rok za oddajo ponudb:**", lep_datum(objava["datumDoPonudba"]))
        st.write("**Rok za vprašanja:**", lep_datum(objava["datumDoVprasanja"]))

        url_dokumentacije = objava["razpisnaDokumentacijaURL"]

        if not url_dokumentacije:
            st.warning("Dokumentacija ni objavljena na portalu. Naložite jo ročno spodaj.")
        elif povezava not in st.session_state["povzetki"] and st.button("Pripravi povzetek dokumentacije"):
            with st.spinner("Prenos in branje dokumentacije v teku ..."):
                try:
                    besedilo, prebrane, neprebrane = preberi_zip(url_dokumentacije)
                except Exception as napaka:
                    st.write("Podrobnosti napake:", str(napaka))
                    besedilo, prebrane, neprebrane = "", [], []

            if not besedilo.strip():
                st.error("Dokumentacije ni bilo mogoče prebrati. Prenesite jo ročno in jo naložite spodaj.")
            else:
                pripravi_in_shrani(povezava, besedilo, prebrane, neprebrane)

        pokazi_povzetek(povezava)

st.divider()

with st.expander("Ročno nalaganje dokumentacije"):
    datoteka = st.file_uploader("Izberite PDF, Word ali Excel", type=["pdf", "docx", "xlsx"])
    if datoteka is not None:
        kljuc = "rocno:" + datoteka.name
        besedilo = preberi_datoteko(datoteka.name, datoteka)

        if not besedilo or not besedilo.strip():
            st.error("V dokumentu ni besedila – verjetno je skeniran.")
        elif kljuc not in st.session_state["povzetki"] and st.button("Pripravi povzetek naložene datoteke"):
            pripravi_in_shrani(kljuc, besedilo, [], [])

        pokazi_povzetek(kljuc)
