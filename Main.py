# datenblatt_app.py
import os
import math
import xlwings as xw
from pypdf import PdfWriter
import streamlit as st

# ============================================================
# GLOBALE KONSTANTEN
# ============================================================
HOME_VERZEICHNIS = r"C:\Users\kzt\Python_Code\Automatische Datenblatt Erzeugung V2"
STEP_BASIS = os.path.join(HOME_VERZEICHNIS, "STEP Dateien")
ZEICHNUNGS_PFADE = {
    "KSY": os.path.join(STEP_BASIS, "KSY-Maßblätter"),
    "KSY_B5": os.path.join(STEP_BASIS, "KSY-Maßblätter B5"),
    "KSY_Stecker": os.path.join(STEP_BASIS, "KSY B14 mit Stecker"),
    "KSG": os.path.join(STEP_BASIS, "KSG-Maßblätter"),
}
MASTERFILE_NAME = "SEW_Masterfile.xlsx"
ZIEL_SPALTE = "L"

# ============================================================
# HILFSFUNKTIONEN – EXCEL
# ============================================================
def lese_wert_mit_merge_support(ws, row, col, min_col=1):
    while col >= min_col:
        value = ws.range((row, col)).value
        if value not in (None, ""):
            return value
        col -= 1
    return None

def write_value(ws, row, value):
    ws.range(f"{ZIEL_SPALTE}{row}").value = value

def read_master_value(ws, row, col):
    return ws.range((row, col)).value

def finde_spalte(ws, suchwert):
    header = [str(v).strip() for v in ws.range("A5:XFD5").value]
    if str(suchwert) not in header:
        raise ValueError(f"Eintrag '{suchwert}' nicht in Zeile 5 gefunden")
    return header.index(str(suchwert)) + 1

def finde_pdf_mit_text(verzeichnis, textbaustein):
    for datei in os.listdir(verzeichnis):
        if datei.lower().endswith(".pdf") and textbaustein.lower() in datei.lower():
            return os.path.join(verzeichnis, datei)
    return None

# ============================================================
# SCHREIBLOGIK – MOTOR / GETRIEBE
# ============================================================
def write_common_motor_data(ws_master, ws_ziel, col, polzahl, schutzart, betriebsart, isolationsklasse):
    mapping = {6:8,8:16,17:12,18:15,20:17,21:18,25:21,26:22,23:23}
    for src, target in mapping.items():
        write_value(ws_ziel, target, read_master_value(ws_master, src, col))
    write_value(ws_ziel, 19, polzahl)
    write_value(ws_ziel, 24, schutzart)
    write_value(ws_ziel, 25, betriebsart)
    write_value(ws_ziel, 26, isolationsklasse)
    write_value(ws_ziel, 5, "BOT")
    spannung = read_master_value(ws_master, 6, col)
    if spannung == 400:
        write_value(ws_ziel, 20, round(480 * math.sqrt(2), 0))
    elif spannung == 230:
        write_value(ws_ziel, 20, round(230 * math.sqrt(2), 0))

def write_leistungsdaten(ws_master, ws_ziel, col, zeilen):
    ziel_mapping = {"n":9,"m":10,"i":13,"ms":11,"is":14}
    for key, src_row in zeilen.items():
        write_value(ws_ziel, ziel_mapping[key], read_master_value(ws_master, src_row, col))

# ============================================================
# PDF-ZUSAMMENFÜHRUNG
# ============================================================
def pdf_mergen(pdf1, pdf2, speicherpfad):
    merger = PdfWriter()
    merger.append(pdf1)
    merger.append(pdf2)
    merger.write(speicherpfad)
    merger.close()
    st.success(f"PDFs erfolgreich zusammengefügt: {speicherpfad}")

# ============================================================
# HAUPTVERARBEITUNG
# ============================================================
def weiterverarbeitung(motor, motorgetriebe_string, motorgetriebe_nummer,
                        variante, b5, getriebe, getriebeuebersetzung,
                        motor_string, spaltezahl_float, zahlen, polzahl,
                        schutzart, betriebsart, isolationsklasse,
                        datenblatt_pdf, bremse, rotorlagegeber, Zeichnungs_string):
    app = xw.App(visible=False)
    try:
        master_path = os.path.join(HOME_VERZEICHNIS, MASTERFILE_NAME)
        wb_master = xw.Book(master_path)
        ws_master = wb_master.sheets[motorgetriebe_string if getriebe else motor_string]

        ziel_path = os.path.join(HOME_VERZEICHNIS,
                                 "Datenblattvorlage_Getriebemotor.xlsx" if getriebe else "Datenblattvorlage_Motor.xlsx")
        wb_ziel = xw.Book(ziel_path)
        ws_ziel = wb_ziel.sheets[0]

        col = finde_spalte(ws_master, spaltezahl_float)
        write_common_motor_data(ws_master, ws_ziel, col, polzahl, schutzart, betriebsart, isolationsklasse)

        if not getriebe:
            varianten = {(False, "R4"): {"n":9,"m":11,"i":12,"ms":14,"is":15},
                         (True, "R4"): {"n":34,"m":36,"i":37,"ms":39,"is":40},
                         (False, "Rx"): {"n":47,"m":49,"i":50,"ms":52,"is":53}}
            write_leistungsdaten(ws_master, ws_ziel, col, varianten[(bremse, rotorlagegeber)])
        else:
            write_leistungsdaten(ws_master, ws_ziel, col, {"n":9,"m":11,"i":12,"ms":14,"is":15})
            write_value(ws_ziel, 30, getriebeuebersetzung)
            ws_master = wb_master.sheets[motor_string]
            col = finde_spalte(ws_master, motorgetriebe_nummer)
            write_value(ws_ziel, 29, lese_wert_mit_merge_support(ws_master, 33, col))
            write_value(ws_ziel, 31, read_master_value(ws_master, 26, col))
            write_value(ws_ziel, 32, read_master_value(ws_master, 24, col))
            write_value(ws_ziel, 34, read_master_value(ws_master, 28, col))

        bremse_label = "-MD" if bremse else ""
        ws_ziel.range("C6").value = f"{motor} {spaltezahl_float} {variante} {bremse_label} {rotorlagegeber} / {read_master_value(ws_master,6,col)}"

        wb_master.close()
        wb_ziel.save()

        if datenblatt_pdf:
            pdf_name = f"Datenblatt_{motor_string}_{spaltezahl_float}.pdf"
            pdf_name_with_path = os.path.join(HOME_VERZEICHNIS, pdf_name)
            ws_ziel.api.ExportAsFixedFormat(Type=0,Filename=pdf_name_with_path,Quality=0,IncludeDocProperties=True,IgnorePrintAreas=False,OpenAfterPublish=False)
        wb_ziel.close()
        return pdf_name_with_path

    finally:
        app.quit()

# ============================================================
# HILFSFUNKTIONEN – LOGIK
# ============================================================
def erstelle_motor_string(motor, variante, b5, b5_string):
    return f"{motor} {variante} {b5_string}" if b5 else f"{motor} {variante}"

def erstelle_motornummer_float(zahlen):
    vordere = ''.join(str(z) for z in zahlen[:-1])
    letzte = str(zahlen[-1])
    return f"{vordere}.{letzte}"

def erstelle_zeichnungs_string(motor, baugroesse, polzahl, variante, b5=False, b5_string=None, passfeder=False, passfeder_string="PF", blockflansch=False, blockflansch_string="BF", stecker=False):
    teile = []
    basis = f"{motor} {baugroesse}{polzahl}x {variante}"
    teile.append(basis)
    if motor.startswith("KSY"):
        if passfeder:
            teile.append(f"mit {passfeder_string}")
        elif stecker:
            teile.append(f"mit Stecker")
        elif b5:
            if b5_string == "B5":
                teile.append("mit B5")
            elif b5_string == "B14":
                teile.append("mit Stecker")
    elif motor.startswith("KSG"):
        if blockflansch:
            teile.append(f"mit {blockflansch_string}")
    return " ".join(teile)

# ============================================================
# STREAMLIT APP
# ============================================================
st.title("Automatische Datenblatt Erzeugung V2")

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

if st.button("Datenblatt erzeugen"):
    b5_string = "B5" if b5 else "B14"
    motor_string = erstelle_motor_string(motor, variante, b5, b5_string)
    zahlen = [baugroesse, polzahl, paketlaenge, bemessungsdrehzahl]
    spaltezahl_float = erstelle_motornummer_float(zahlen)
    Zeichnungs_string = erstelle_zeichnungs_string(motor, baugroesse, polzahl, variante, b5, b5_string, passfeder, "PF", blockflansch, "BF", stecker)

    motorgetriebe_string = "KSY HD - KSG" if motor=="KSG" else ""
    motorgetriebe_nummer = f"{motor} {spaltezahl_float} {getriebeuebersetzung}" if motor=="KSG" else ""

    pdf_path = weiterverarbeitung(motor, motorgetriebe_string, motorgetriebe_nummer,
                                  variante, b5, getriebe, getriebeuebersetzung,
                                  motor_string, spaltezahl_float, zahlen, polzahl,
                                  schutzart, betriebsart, isolationsklasse,
                                  datenblatt_pdf, bremse, rotorlagegeber, Zeichnungs_string)

    # Merge PDFs
    if motor=="KSY":
        key = "KSY_B5" if b5 else "KSY_Stecker"
    else:
        key = "KSG"
    pdf_drawing_verzeichnis = ZEICHNUNGS_PFADE[key]
    pdf_drawing_textbaustein = Zeichnungs_string
    pdf_drawing_datei = finde_pdf_mit_text(pdf_drawing_verzeichnis, pdf_drawing_textbaustein)

    pdf_final_name = f"Datenblatt_{motor_string}_{spaltezahl_float}_V2.pdf"
    pdf_final_path = os.path.join(HOME_VERZEICHNIS, pdf_final_name)
    pdf_mergen(pdf_path, pdf_drawing_datei, pdf_final_path)
