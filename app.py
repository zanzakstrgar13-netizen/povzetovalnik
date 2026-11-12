import time
import streamlit as st
from pypdf import PdfReader
from docx import Document
from dotenv import load_dotenv
from google import genai

load_dotenv()

MODELI = ["gemini-3.5-flash-lite", "gemini-3.6-flash", "gemini-3.8-flash"]

NAVODILO = """Si pomočnik za javna naročila v Sloveniji.
Spodaj sta dva vira o istem javnem naročilu: OBJAVA s portala in RAZPISNA DOKUMENTACIJA.
Iz obeh skupaj v slovenščini izpiši točno te podatke:

Naročnik:
Predmet naročila:
Ocenjena vrednost:
Rok za oddajo ponudb:
Rok za vprašanja ponudnikov:
Pogoji za sodelovanje (alineje):
Merila za izbor ponudbe:
Kaj mora ponudnik oddati (alineje):

Roke vzemi predvsem iz objave.
Če kakšnega podatka ni v nobenem viru, napiši "ni navedeno". Ničesar si ne izmišljuj.
"""


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


# Izgled strani
st.title("📄 Povzetovalnik razpisov")
st.write("Naloži razpisno dokumentacijo in prilepi objavo s portala. Dobiš povzetek najpomembnejših podatkov.")

datoteka = st.file_uploader("1. Izberi dokumentacijo (PDF ali Word)", type=["pdf", "docx"])
objava = st.text_area("2. Prilepi besedilo objave z enarocanje.si (neobvezno, a priporočeno zaradi rokov)", height=150)

if datoteka is not None:
    if datoteka.name.endswith(".docx"):
        besedilo = preberi_word(datoteka)
    else:
        besedilo = preberi_pdf(datoteka)

    if len(besedilo.strip()) == 0:
        st.error("V dokumentu ni besedila – verjetno je skeniran.")
        st.stop()

    st.write("Dokumentacija:", len(besedilo), "znakov. Objava:", len(objava), "znakov.")

    if st.button("Naredi povzetek"):
        vsebina = NAVODILO + "\n\nOBJAVA:\n" + objava + "\n\nRAZPISNA DOKUMENTACIJA:\n" + besedilo

        odgovor = None
        client = genai.Client()
        sporocilo = st.empty()

        for model in MODELI:
            for poskus in range(1, 3):
                sporocilo.info(f"AI bere dokument ... (model {model}, poskus {poskus}/2)")
                try:
                    odgovor = client.models.generate_content(model=model, contents=vsebina)
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