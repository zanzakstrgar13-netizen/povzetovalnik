import io
import time
import zipfile
from datetime import datetime

import requests
import streamlit as st
from pypdf import PdfReader
from docx import Document
from dotenv import load_dotenv
from google import genai

load_dotenv()

MODELI = ["gemini-3.5-flash-lite", "gemini-3.6-flash", "gemini-3.8-flash"]
GLAVA = {"User-Agent": "Mozilla/5.0"}

NAVODILO = """Si pomočnik za javna naročila v Sloveniji.
Iz spodnje razpisne dokumentacije v slovenščini izpiši:

Pogoji za sodelovanje (alineje):
Merila za izbor ponudbe:
Kaj mora ponudnik oddati (alineje):
Finančna zavarovanja in garancije:
Rok dobave oziroma izvedbe:
Plačilni pogoji:

Če kakšnega podatka v dokumentaciji ni, napiši "ni navedeno". Ničesar si ne izmišljuj.

Dokumentacija:
"""


# ---------- Branje objave ----------

def preberi_objavo(povezava):
    id_objave = povezava.strip().rstrip("/").split("/")[-1]
    if not id_objave.isdigit():
        return None
    url = "https://www.enarocanje.si/api/obrazec/objava/obrazecGet?id=" + id_objave
    glava = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
    odziv = requests.get(url, headers=glava, timeout=30)
    odziv.raise_for_status()
    return odziv.json()


def lep_datum(besedilo):
    if not besedilo:
        return "ni navedeno"
    return datetime.fromisoformat(besedilo).strftime("%d. %m. %Y ob %H:%M")


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


def preberi_zip(url):
    # Prenesi datoteko s portala
    odziv = requests.get(url, headers=GLAVA, timeout=60)
    odziv.raise_for_status()
    vsebina = odziv.content

    # Če dokumentacija ni ZIP, ampak kar en PDF
    if vsebina.startswith(b"%PDF"):
        return preberi_pdf(io.BytesIO(vsebina)), ["dokumentacija.pdf"]

    # ZIP datoteke se vedno začnejo s črkama PK
    if not vsebina.startswith(b"PK"):
        return "", []

    arhiv = zipfile.ZipFile(io.BytesIO(vsebina))
    besedilo = ""
    prebrane = []

    for ime in arhiv.namelist():
        malo = ime.lower()
        try:
            if malo.endswith(".docx"):
                besedilo += "\n\n=== " + ime + " ===\n" + preberi_word(io.BytesIO(arhiv.read(ime)))
                prebrane.append(ime)
            elif malo.endswith(".pdf"):
                besedilo += "\n\n=== " + ime + " ===\n" + preberi_pdf(io.BytesIO(arhiv.read(ime)))
                prebrane.append(ime)
        except Exception:
            pass  # če se ena datoteka ne da prebrati, jo preskočimo

    return besedilo, prebrane


# ---------- AI ----------

def naredi_povzetek(besedilo):
    client = genai.Client(http_options={"timeout": 90000})
    sporocilo = st.empty()
    zadnja_napaka = None

    for model in MODELI:
        for poskus in range(1, 3):
            sporocilo.info(f"AI bere dokumentacijo ... (model {model}, poskus {poskus}/2)")
            try:
                odgovor = client.models.generate_content(model=model, contents=NAVODILO + besedilo)
                sporocilo.empty()
                return odgovor.text
            except Exception as napaka:
                zadnja_napaka = napaka
                time.sleep(10)

    sporocilo.empty()
    st.write("Zadnja napaka:", str(zadnja_napaka))
    return None


# ---------- Izgled strani ----------

st.title("📄 Povzetovalnik razpisov")
povezava = st.text_input("Prilepi povezavo do objave z enarocanje.si")

if povezava:
    try:
        objava = preberi_objavo(povezava)
    except Exception as napaka:
        st.write("Podrobnosti napake:", str(napaka))
        objava = None

    if objava is None:
        st.error("Objave ni bilo mogoče prebrati. Preveri povezavo.")
    else:
        st.subheader(objava["naslov"])
        if objava["narocnik"]:
            st.write("**Naročnik:**", objava["narocnik"][0]["naziv"])
        st.write("**Rok za oddajo ponudb:**", lep_datum(objava["datumDoPonudba"]))
        st.write("**Rok za vprašanja:**", lep_datum(objava["datumDoVprasanja"]))

        url_dokumentacije = objava["razpisnaDokumentacijaURL"]

        if not url_dokumentacije:
            st.warning("Dokumentacija ni objavljena na portalu. Naloži jo ročno spodaj.")
        elif st.button("Naredi AI povzetek dokumentacije"):
            with st.spinner("Prenašam in berem dokumentacijo ..."):
                try:
                    besedilo, prebrane = preberi_zip(url_dokumentacije)
                except Exception as napaka:
                    st.write("Podrobnosti napake:", str(napaka))
                    besedilo, prebrane = "", []

            if not besedilo.strip():
                st.error("Dokumentacije ni bilo mogoče prebrati. Prenesi jo ročno in jo naloži spodaj.")
            else:
                st.write("**Prebrani dokumenti:**", ", ".join(prebrane))
                povzetek = naredi_povzetek(besedilo)
                if povzetek:
                    st.markdown(povzetek)
                else:
                    st.error("Vsi modeli so trenutno zasedeni. Poskusi čez nekaj minut.")

st.divider()

with st.expander("Ali naloži dokumentacijo ročno"):
    datoteka = st.file_uploader("Izberi PDF ali Word", type=["pdf", "docx"])
    if datoteka is not None:
        if datoteka.name.endswith(".docx"):
            besedilo = preberi_word(datoteka)
        else:
            besedilo = preberi_pdf(datoteka)

        if not besedilo.strip():
            st.error("V dokumentu ni besedila – verjetno je skeniran.")
        elif st.button("Naredi povzetek naložene datoteke"):
            povzetek = naredi_povzetek(besedilo)
            if povzetek:
                st.markdown(povzetek)
            else:
                st.error("Vsi modeli so trenutno zasedeni. Poskusi čez nekaj minut.")