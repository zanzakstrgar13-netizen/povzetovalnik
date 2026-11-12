import time
from datetime import datetime

import requests
import streamlit as st
from pypdf import PdfReader
from docx import Document
from dotenv import load_dotenv
from google import genai

load_dotenv()

MODELI = ["gemini-3.5-flash-lite", "gemini-3.6-flash", "gemini-3.8-flash"]

NAVODILO = """Si pomočnik za javna naročila v Sloveniji.
Iz spodnje razpisne dokumentacije v slovenščini izpiši:

Pogoji za sodelovanje (alineje):
Merila za izbor ponudbe:
Kaj mora ponudnik oddati (alineje):
Finančna zavarovanja in garancije:
Rok dobave oziroma izvedbe:
Plačilni pogoji:

Če kakšnega podatka v dokumentu ni, napiši "ni navedeno". Ničesar si ne izmišljuj.

Dokument:
"""


def preberi_objavo(povezava):
    # Iz povezave vzamemo številko na koncu, npr. 1230728
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


# ---------- Izgled strani ----------
st.title("📄 Povzetovalnik razpisov")

# 1. del: objava s portala
povezava = st.text_input("1. Prilepi povezavo do objave z enarocanje.si")

if povezava:
    try:
        objava = preberi_objavo(povezava)
    except Exception:
        objava = None

    if objava is None:
        st.error("Objave ni bilo mogoče prebrati. Preveri povezavo.")
    else:
        st.subheader(objava["naslov"])
        if objava["narocnik"]:
            st.write("**Naročnik:**", objava["narocnik"][0]["naziv"])
        st.write("**Rok za oddajo ponudb:**", lep_datum(objava["datumDoPonudba"]))
        st.write("**Rok za vprašanja:**", lep_datum(objava["datumDoVprasanja"]))
        if objava["razpisnaDokumentacijaURL"]:
            st.markdown("[⬇️ Prenesi razpisno dokumentacijo](" + objava["razpisnaDokumentacijaURL"] + ")")

st.divider()

# 2. del: dokumentacija in AI povzetek
datoteka = st.file_uploader("2. Naloži razpisno dokumentacijo (PDF ali Word)", type=["pdf", "docx"])

if datoteka is not None:
    if datoteka.name.endswith(".docx"):
        besedilo = preberi_word(datoteka)
    else:
        besedilo = preberi_pdf(datoteka)

    if len(besedilo.strip()) == 0:
        st.error("V dokumentu ni besedila – verjetno je skeniran.")
        st.stop()

    st.write("Dokumentacija:", len(besedilo), "znakov.")

    if st.button("Naredi AI povzetek dokumentacije"):
        odgovor = None
        client = genai.Client()
        sporocilo = st.empty()

        for model in MODELI:
            for poskus in range(1, 3):
                sporocilo.info(f"AI bere dokument ... (model {model}, poskus {poskus}/2)")
                try:
                    odgovor = client.models.generate_content(model=model, contents=NAVODILO + besedilo)
                    break
                except Exception:
                    time.sleep(10)
            if odgovor is not None:
                break

        sporocilo.empty()

        if odgovor is not None:
            st.markdown(odgovor.text)
        else:
            st.error("Vsi modeli so trenutno zasedeni. Poskusi čez nekaj minut.")