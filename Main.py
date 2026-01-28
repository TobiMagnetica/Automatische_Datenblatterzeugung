# datenblatt_app_cloud.py
import os
import math
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
import streamlit as st
from PyPDF2 import PdfMerger

# ============================================================
# GITHUB LINKS ZU DEN EXCEL-DATEIEN
# ============================================================
MASTERFILE_URL = "https://github.com/TobiMagnetica/Automatische_Datenblatterzeugung/raw/main/SEW_Masterfile.xlsx"
TEMPLATE_GETRIEBE_URL = "https://github.com/TobiMagnetica/Automatische_Datenblatterzeugung/raw/main/Datenblattvorlage_Getriebemotor.xlsx"
TEMPLATE_MOTOR_URL = "https://github.com/TobiMagnetica/Automatische_Datenblatterzeugung/raw/main/Datenblattvorlage_Motor.xlsx"

# Verzeichnisse für Zeichnungs-PDFs (lokal im Repo)
ZEICHNUNGS_PFADE = {
    "KSY": "STEP Dateien/KSY-Maßblätter",
    "KSY_B5": "STEP Dateien/KSY-Maßblätter B5",
    "KSY_Stecker": "STEP Dateien/KSY B14 mit Stecker",
    "KSG": "STEP Dateien/KSG-Maßblätter",
}

# ============================================================
# HILFSFUNKTIONEN
# ============================================================

def lese_wert_mit_merge_support(df, row, col, min_col=0):
    """Liest eine Zelle und geht nach links, falls leer (z.B. bei Merge-Zellen)."""
    while col >= min_col:
        value = df.iat[row, col]
        if pd.notna(value):
            return value
        col -= 1
    return None

def finde_spalte(df, suchwert):
    """Findet die Spalte eines Werts in Zeile 5 (Index 4)."""
    header = df.iloc[4, :].astype(str).str.strip()
    if str(suchwert) not in header.values:
        raise ValueError(f"Eintrag '{suchwert}' nicht gefunden")
    return header[header == str(suchwert)].index[0]

def erstelle_motor_string(motor, variante, b5, b5_string):
    return f"{motor} {variante} {b5_string}" if b5 else f"{motor} {variante}"

def erstelle_motornummer_float(zahlen):
    vordere = ''.join(str(z) for z in zahlen[:-1])
    letzte = str(zahlen[-1])
    return f"{vordere}.{letzte}"

def erstelle_zeichnungs_string(motor, baugroesse, polzahl, variante, b5=False, b5_string=None,
                               passfeder=False, passfeder_string="PF",
                               blockflansch=False, blockflansch_string="BF",
                               stecker=False):
    teile = []
    basis = f"{motor} {baugroesse}{polzahl}x {variante}"
    teile.append(basis)

    if motor.startswith("KSY"):
        if passfeder:
            teile.append(f"mit {passfeder_string}")
        elif stecker:
            teile.append("mit Stecker")
        elif b5:
            if b5_string == "B5":
                teile.append("mit B5")
            elif b5_string == "B14":
                teile.append("mit Stecker")
    elif motor.startswith("KSG"):
        if blockflansch:
            teile.append(f"mit {blockflansch_string}")

    return " ".join(teile)

def finde_pdf_mit_text(verzeichnis, textbaustein):
    """Sucht PDF-Datei im Verzeichnis nach Textbaustein."""
    if not os.path.exists(verzeichnis):
        return None
    for datei in os.listdir(verzeichnis):
        if datei.lower().endswith(".pdf") and textbaustein.lower() in datei.lower():
            return os.path.join(verzeichnis, datei)
    return None

def pdf_mergen(pdf1, pdf2, speicherpfad):
    merger = PdfMerger()
    merger.append(pdf1)
    if pdf2:
        merger.append(pdf2)
    merger.write(speicherpfad)
    merger.close()
    st.success(f"PDF erfolgreich erstellt: {speicherpfad}")

def create_pdf_from_dataframe(df, pdf_path):
    """Erzeugt PDF aus Pandas DataFrame über ReportLab."""
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    data = df.fillna("").astype(str).values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('FONTNAME', (0,0),(-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING',(0,0),(-1,0),12),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
    ]))
    doc.build([table])

# ============================================================
# STREAMLIT APP
# ============================================================
st.title("Automatische Datenblatt-Erzeugung (Cloud-Version)")

# ---------------- Auswahl ----------------
motor = st.selectbox("Motorgrundtyp:", ["KSY", "KSD", "KSG", "KTY"])
variante = st.selectbox("Variante:", ["HD"])
baugroesse = st.selectbox("Baugröße:", ["1","2","3","4","5","6","8"])
polzahl = st.selectbox("Polzahl:", ["2","4","6","8","10","12","14"])
paketlaenge = st.selectbox("Paketlänge in cm:", ["2","4","6","8","10","12","16"])
bemessungsdrehzahl = st.selectbox("Bemessungsdrehzahl in 100/*min:", ["25","30","40","45","50","55","60","80","90"])
schutzart = st.selectbox("Schutzart:", ["IP54","IP65","IP67","IP69K"])
betriebsart = st.selectbox("Betriebsart:", ["S1"])
isolationsklasse = st.selectbox("Isolationsklasse:", ["F","H"])
rotorlagegeber = st.selectbox("Rotorlagegeber:", ["R4","Rx"])

bremse = st.checkbox("Bremse?")
b5 = st.checkbox("Mit B5 Flansch?")
getriebe = st.checkbox("Mit Getriebe?")
datenblatt_pdf = st.checkbox("Datenblatt direkt als PDF speichern!")
passfeder = st.checkbox("Mit Passfeder?")
blockflansch = st.checkbox("Mit Blockflansch?")
stecker = st.checkbox("Mit Stecker?")

getriebeuebersetzung = st.selectbox("Getriebeübersetzung:", ["3","5","7","9","10","12","15","21","25","30","35","49","70","100"]) if getriebe else None

# ---------------- Button für Erzeugung ----------------
if st.button("Datenblatt erzeugen"):
    # ---------------- Masterfile laden ----------------
    sheet_name = f"{motor} {variante}" if not getriebe else f"{motor} {variante} - Getriebe"
    df_master = pd.read_excel(MASTERFILE_URL, sheet_name=sheet_name, header=None)

    # ---------------- Vorlage laden ----------------
    template_url = TEMPLATE_GETRIEBE_URL if getriebe else TEMPLATE_MOTOR_URL
    df_ziel = pd.read_excel(template_url, header=None)

    # ---------------- Motorstrings erstellen ----------------
    b5_string = "B5" if b5 else "B14"
    motor_string = erstelle_motor_string(motor, variante, b5, b5_string)
    zahlen = [baugroesse, polzahl, paketlaenge, bemessungsdrehzahl]
    spaltezahl_float = erstelle_motornummer_float(zahlen)
    Zeichnungs_string = erstelle_zeichnungs_string(motor, baugroesse, polzahl, variante,
                                                    b5, b5_string, passfeder, "PF",
                                                    blockflansch, "BF", stecker)

    # ---------------- Beispiel: Werte eintragen ----------------
    # Hier kannst du alle Zellen aus df_master einfügen wie in xlwings-Version
    df_ziel.iat[5,0] = "BOT"
    df_ziel.iat[19,0] = polzahl
    df_ziel.iat[24,0] = schutzart
    df_ziel.iat[25,0] = betriebsart
    df_ziel.iat[26,0] = isolationsklasse
    df_ziel.iat[0,0] = f"{motor} {spaltezahl_float} {variante}"

    # ---------------- PDF erzeugen ----------------
    pdf_path = f"Datenblatt_{motor_string}_{spaltezahl_float}.pdf"
    create_pdf_from_dataframe(df_ziel, pdf_path)
    st.success(f"Datenblatt PDF erstellt: {pdf_path}")

    # ---------------- Merge mit Zeichnungs-PDF ----------------
    key = "KSY_B5" if b5 else "KSY_Stecker" if motor=="KSY" else "KSG"
    pdf_drawing_verzeichnis = ZEICHNUNGS_PFADE[key]
    pdf_drawing_datei = finde_pdf_mit_text(pdf_drawing_verzeichnis, Zeichnungs_string)
    
    pdf_final_path = f"Datenblatt_{motor_string}_{spaltezahl_float}_V2.pdf"
    pdf_mergen(pdf_path, pdf_drawing_datei, pdf_final_path)
    
    # --- NEU: Download-Button ---
    with open(pdf_final_path, "rb") as f:
        st.download_button(
            label="Datenblatt herunterladen",
            data=f,
            file_name=os.path.basename(pdf_final_path),
            mime="application/pdf"
        )
