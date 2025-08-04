import os
from pathlib import Path
import re
import json
import tkinter as tk
from tkinter import ttk
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
from gg import get_titles_from_url
import warnings


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


Aufgaben = "aufgaben.txt"

# Arbeitsverzeichnis (hier: aktuelles Verzeichnis)
VERZEICHNIS = r"Z:\Gruppen\1_NEUE_GRUPPEN"

letzte_meldungen = {}

def finde_hendrik_ordner(verzeichnis):
    """
    Sucht im angegebenen Verzeichnis nach Ordnern, die 'hendrik' im Namen haben.
    Gibt eine Liste von Tupeln (Ordnername, Typ) zurück.
    Typ ist 'primär' oder 'sekundär' (wenn ein '+' im Namensteil steht).
    """
    ordner = []
    muster = re.compile(r"\([^\)]*hendrik", re.IGNORECASE)
    for name in os.listdir(verzeichnis):
        pfad = os.path.join(verzeichnis, name)
        if os.path.isdir(pfad):
            match = muster.search(name)
            if match:
                zwischen = name[match.start():match.end()]
                if "+" in zwischen:
                    ordner.append((name, "sekundär"))
                else:
                    ordner.append((name, "primär"))
    return ordner

def backup(Ordnername, i):
    regex = r"\(\s*[A-Za-z]{2,}\s*(\+|und)\s*[A-Za-z]{2,}\s*\)"
    treffer = re.search(regex, Ordnername)
    eintrag = ws["B6"].value 
    # Prüfung: Wert vorhanden & besteht aus mindestens 2 Wörtern
    if isinstance(eintrag, str) and len(eintrag.strip().split()) >= 2:
        eintrag = True
    else:
        eintrag = False  
    if treffer and eintrag:
        i[1] = datetime.now().date()
    elif treffer and not eintrag:
        return f"{Ordnername} Backup nicht in Excel eingetragen!"
    elif eintrag and not treffer:
        return f"{Ordnername} Backup nicht im Ordnernamen eingetragen!"

def termin_GGG(Ordnername, i):
        GGG_Termin = ws["B5"].value.date()
        # Prüfung, ob ein Wert vorhanden ist
        if pd.notna(GGG_Termin):
            i[1] = datetime.now().date()
        else:
            return f"{Ordnername} Termin für GGG vereinbaren!"
        
def text_warten(Ordnername, i):
    aktuellerpfad = Path(VERZEICHNIS) / Ordnername
    # Prüfen, ob mindestens ein Unterordner existiert
    unterordner = any(p.is_dir() for p in aktuellerpfad.iterdir())
    if unterordner:
        i[1] = datetime.now().date()
    elif not unterordner: 
        heute = datetime.now().date()
        zwei_wochen = timedelta(weeks=2)
        GGG_Termin = ws["B5"].value.date()
        if heute - GGG_Termin > zwei_wochen:
            return "Es sind mehr als 2 Wochen seit dem GGG-Termin vergangen, bitte Initiator:in bezüglich Beschreibung kontaktieren."
        else:
            return f"{Ordnername} warten auf Text für Homepage / Insta / Presse"

def zettel(Ordnername, i):
    a = input("Wurde der grüne Zettel Brigitte gegeben? (ja/nein): ").strip().lower()
    while a not in ["ja", "nein"]:
        print ("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der grüne Zettel Brigitte gegeben? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        return f"{Ordnername} grünen Zettel nicht Brigitte gegeben!"
        
def homepage(Ordnername, i):
    homepage = 0
    excel = 0
    Name = Ordnername.split(" (")[0]
    Gruppen_in_Gruendung = get_titles_from_url(os.getenv("url"))
    if Name in Gruppen_in_Gruendung:
        homepage = 1
    # ist etwas in Zelle B29 eingetragen?
    if pd.notna(df.iloc[27, 1]):
        excel = 1
    if homepage and excel or df.iloc[27, 1] == "-":
        i[1] = datetime.now().date()
    elif not homepage and excel:
        return f"Gruppe {Ordnername} nicht auf Homepage gefunden! Stimmen Ordnername und Gruppenname überein?"
    elif homepage and not excel:
        return f"Gruppe {Ordnername} auf Homepage gefunden aber kein Eintrag in Excel vorhanden!"
    else:
        return f"Gruppe {Ordnername} weder auf Homepage noch in Excel gefunden! Bitte prüfen!"
    

def instagram(Ordnername, i):
    # ist etwas in Zelle B30 eingetragen?
    if pd.notna(df.iloc[29, 1]):
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} kein Eintrag in Excel für Instagram vorhanden! (Datum oder '-')"

def presse(Ordnername, i):
    # ist etwas in Zelle B31 eingetragen?
    if pd.notna(df.iloc[30, 1]):
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} kein Eintrag in Excel für Pressemitteilung vorhanden! (Datum oder '-')"
            
def interessenten(Ordnername, i): 
    excel_namen = [
    "H4", "H18", "H32", "H46",
    "R4", "R18", "R32", "R46",
    "AB4", "AB18", "AB32", "AB46",
    "AL4", "AL18", "AL32", "AL46",
    "AV4", "AV18", "AV32", "AV46",
    "BF4", "BF18", "BF32", "BF46"
    ]
    anzahl = 0
    heute = datetime.now().date()
    GGG_Termin = ws["B5"].value.date()
    zwei_monate = timedelta(weeks=2)
    for interessent in excel_namen:
        if pd.notna(ws[interessent].value):
            anzahl += 1
    if anzahl >= 4:
        i[1] = datetime.now().date()
    if anzahl < 4 and heute - GGG_Termin > zwei_monate:
        return f"Bei Gruppe {Ordnername} ist das Gründungsgespräch mehr als 2 Monate her und es sind weniger als 4 ({anzahl} Personen) Interessent:innen auf der Liste. Initiator:in bezgl. weiterem Vorgehen kontaktieren."
    else:
        return f"Es gibt erst {anzahl}/4 Interessent:innen für ein erstes Treffen"
     

def erstesTreffen(Ordnername, i):   
    # Ist ein Datum in Zelle B21 eingetragen?
    if pd.notna(df.iloc[19, 1]):
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} Termin für erstes Treffen vereinbaren!"

def konferenzraum1(Ordnername, i):
    a = input(f"Wurde der Konferenzraum für das erste Treffen am {df.iloc[19, 1].strftime("%d.%m.%Y")} reserviert? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        return f"{Ordnername} Konferenzraum nicht reserviert!"
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()    

def infoTreffen1(Ordnername, i):
    excel_namen = [
    "H4", "H18", "H32", "H46",
    "R4", "R18", "R32", "R46",
    "AB4", "AB18", "AB32", "AB46",
    "AL4", "AL18", "AL32", "AL46",
    "AV4", "AV18", "AV32", "AV46",
    "BF4", "BF18", "BF32", "BF46",
    ]
    infoCheck = [
    "K4", "K18", "K32", "K46",
    "U4", "U18", "U32", "U46",
    "AE4", "AE18", "AE32", "AE46",
    "AO4", "AO18", "AO32", "AO46",
    "AY4", "AY18", "AY32", "AY46",
    "BI4", "BI18", "BI32", "BI46"
    ]
    nicht_informiert = []
    infoCount = 0
    for interessent in excel_namen:
        if ws[interessent].value is not None:
            if ws[infoCheck[infoCount]].value == False:
                nicht_informiert.append(ws[interessent].value)
        infoCount += 1
    if nicht_informiert == []:
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} wurden folgende Interessent:innen über den ersten Termin am {df.iloc[19, 1].strftime("%d.%m.%Y")} informiert?: \n-{'\n-'.join(nicht_informiert)} \nes ist kein Haken in der Interessiertenliste gesetzt!"
        
def anwesenheit1(Ordnername, i):
    # liegt das erste Treffen in der Vergangenheit?
    if pd.isna(df.iloc[19, 1]):
        return f"{Ordnername} Termin für erstes Treffen nicht eingetragen!"
    erstesTreffen = df.iloc[19, 1].date()
    heute = datetime.now().date()
    checks = False
    anzahl = False
    anwesenheiten = [
    "M4", "M18", "M32", "M46",
    "W4", "W18", "W32", "W46",
    "AG4", "AG18", "AG32", "AG46",
    "AQ4", "AQ18", "AQ32", "AQ46",
    "BA4", "BA18", "BA32", "BA46",
    "BK4", "BK18", "BK32", "BK46"
    ]
    if erstesTreffen < heute:
        # ist ein Wert in Zelle D21 eingetragen?
        if pd.notna(df.iloc[19, 3]):
            anzahl = True
        # ist mindestens eine Anwesenheit bei den Interessent:innen eingetragen?
        for interessent in anwesenheiten:
            if ws[interessent].value is not None:
                checks = True
                break
        if anzahl and checks:
            i[1] = datetime.now().date()
        elif anzahl and not checks:
            return f"{Ordnername} Anzahl in Zelle D21 eingetragen aber keine Anwesenheiten abgehakt!"
        elif not anzahl and checks:
            return f"{Ordnername} Anwesenheiten abgehakt aber keine Anzahl in Zelle D21 eingetragen!"
        else:
            return f"{Ordnername} weder Anzahl in Zelle D21 eingetragen noch Anwesenheiten abgehakt!"
    else:
        return f"{Ordnername} warten bis erstes Treffen am {df.iloc[19, 1].strftime("%d.%m.%Y")} stattgefunden hat!"

def zweitesTreffen(Ordnername, i):    
    # Ist ein Datum in Zelle B21 eingetragen?
    if pd.notna(df.iloc[20, 1]):
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} Termin für erstes Treffen vereinbaren!"

        
def konferenzraum2(Ordnername, i):
    a = input(f"Wurde der Konferenzraum für das zweite Treffen am {df.iloc[20, 1].strftime("%d.%m.%Y")} reserviert? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        return f"{Ordnername} Konferenzraum nicht reserviert!"
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()  
        
def infoTreffen2(Ordnername, i):
    excel_namen = [
    "H4", "H18", "H32", "H46",
    "R4", "R18", "R32", "R46",
    "AB4", "AB18", "AB32", "AB46",
    "AL4", "AL18", "AL32", "AL46",
    "AV4", "AV18", "AV32", "AV46",
    "BF4", "BF18", "BF32", "BF46",
    ]
    infoCheck = [
    "K6", "K20", "K34", "K48",
    "U6", "U20", "U34", "U48",
    "AE6", "AE20", "AE34", "AE48",
    "AO6", "AO20", "AO34", "AO48",
    "AY6", "AY20", "AY34", "AY48",
    "BI6", "BI20", "BI34", "BI48"
    ]
    nicht_informiert = []
    infoCount = 0
    for interessent in excel_namen:
        if ws[interessent].value is not None:
            if ws[infoCheck[infoCount]].value == False:
                nicht_informiert.append(ws[interessent].value)
        infoCount += 1
    if nicht_informiert == []:
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} wurden folgende Interessent:innen über den zweiten Termin am {df.iloc[20, 1].strftime("%d.%m.%Y")}  informiert?: \n-{'\n-'.join(nicht_informiert)} \nes ist kein Haken in der Interessiertenliste gesetzt!"
        
def anwesenheit2(Ordnername, i):
    # liegt das zweite Treffen in der Vergangenheit?
    if pd.isna(df.iloc[20, 1]):
        return f"{Ordnername} Termin für zweites Treffen nicht eingetragen!"
    zweitesTreffen = df.iloc[20, 1].date()
    heute = datetime.now().date()
    checks = False
    anzahl = False
    anwesenheiten = [
    "M6", "M20", "M34", "M48",
    "W6", "W20", "W34", "W48",
    "AG6", "AG20", "AG34", "AG48",
    "AQ6", "AQ20", "AQ34", "AQ48",
    "BA6", "BA20", "BA34", "BA48",
    "BK6", "BK20", "BK34", "BK48"
    ]
    if zweitesTreffen < heute:
        # ist ein Wert in Zelle D22 eingetragen?
        if pd.notna(df.iloc[20, 3]):
            anzahl = True
        # ist mindestens eine Anwesenheit bei den Interessent:innen eingetragen?
        for interessent in anwesenheiten:
            if ws[interessent].value is not None:
                checks = True
                break
        if anzahl and checks:
            i[1] = datetime.now().date()
        elif anzahl and not checks:
            return f"{Ordnername} Anzahl in Zelle D22 eingetragen aber keine Anwesenheiten abgehakt!"
        elif not anzahl and checks:
            return f"{Ordnername} Anwesenheiten abgehakt aber keine Anzahl in Zelle D22 eingetragen!"
        else:
            return f"{Ordnername} weder Anzahl in Zelle D22 eingetragen noch Anwesenheiten abgehakt!"
    else:
        return f"{Ordnername} warten bis zweites Treffen am {df.iloc[20, 1].strftime("%d.%m.%Y")} stattgefunden hat!"
      
def raumsuche(Ordnername, i):
    a = input("Hat die Gruppe einen eigenen Raum für weitere Treffen nach dem dritten Termin? (ja/nein): ").strip().lower()
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()   
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        return f"{Ordnername} Es muss ein Raum für weitere Treffen gefunden werden!"          
    
def drittesTreffen(Ordnername, i):
        # Ist ein Datum in Zelle B23 eingetragen?
    if pd.notna(df.iloc[21, 1]):
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} Termin für drittes Treffen vereinbaren!"

def konferenzraum3(Ordnername, i):
    a = input(f"Wurde der Konferenzraum für das dritte Treffen am {df.iloc[21, 1].strftime("%d.%m.%Y")} reserviert? (ja/nein): ").strip().lower()
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower() 
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        return f"{Ordnername} Konferenzraum nicht reserviert!" 

def infoTreffen3(Ordnername, i):
    excel_namen = [
    "H4", "H18", "H32", "H46",
    "R4", "R18", "R32", "R46",
    "AB4", "AB18", "AB32", "AB46",
    "AL4", "AL18", "AL32", "AL46",
    "AV4", "AV18", "AV32", "AV46",
    "BF4", "BF18", "BF32", "BF46",
    ]
    infoCheck = [
    "K8", "K22", "K36", "K50",
    "U8", "U22", "U36", "U50",
    "AE8", "AE22", "AE36", "AE50",
    "AO8", "AO22", "AO36", "AO50",
    "AY8", "AY22", "AY36", "AY50",
    "BI8", "BI22", "BI36", "BI50"
    ]
    nicht_informiert = []
    infoCount = 0
    for interessent in excel_namen:
        if ws[interessent].value is not None:
            if ws[infoCheck[infoCount]].value == False:
                nicht_informiert.append(ws[interessent].value)
        infoCount += 1
    if nicht_informiert == []:
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} wurden folgende Interessent:innen über den dritten Termin am {df.iloc[21, 1].strftime("%d.%m.%Y")} informiert?: \n-{'\n-'.join(nicht_informiert)} \nes ist kein Haken in der Interessiertenliste gesetzt!"

def fragebogen1(Ordnername, i):
    #Ist ein Wert in Zelle B56 eingetragen?
    if pd.notna(df.iloc[54, 1]):
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} Fragebogen an Initiator:in aushändigen!"

def anwesenheit3(Ordnername, i):
    # liegt das dritte Treffen in der Vergangenheit?
    if pd.isna(df.iloc[21, 1]):
        return f"{Ordnername} Termin für drittes Treffen nicht eingetragen!"
        return
    drittesTreffen = df.iloc[21, 1].date()
    heute = datetime.now().date()
    checks = False
    anzahl = False
    anwesenheiten = [
    "M8", "M22", "M36", "M50",
    "W8", "W22", "W36", "W50",
    "AG8", "AG22", "AG36", "AG50",
    "AQ8", "AQ22", "AQ36", "AQ50",
    "BA8", "BA22", "BA36", "BA50",
    "BK8", "BK22", "BK36", "BK50"
    ]
    if drittesTreffen < heute:
        # ist ein Wert in Zelle D23 eingetragen?
        if pd.notna(df.iloc[21, 3]):
            anzahl = True
        # ist mindestens eine Anwesenheit bei den Interessent:innen eingetragen?
        for interessent in anwesenheiten:
            if ws[interessent].value is not None:
                checks = True
                break
        if anzahl and checks:
            i[1] = datetime.now().date()
        elif anzahl and not checks:
            return f"{Ordnername} Anzahl in Zelle D23 eingetragen aber keine Anwesenheiten abgehakt!"
        elif not anzahl and checks:
            return f"{Ordnername} Anwesenheiten abgehakt aber keine Anzahl in Zelle D23 eingetragen!"
        else:
            return f"{Ordnername} weder Anzahl in Zelle D23 eingetragen noch Anwesenheiten abgehakt!"
    else:
        return f"{Ordnername} warten bis drittes Treffen am {df.iloc[21, 1].strftime("%d.%m.%Y")} stattgefunden hat!"

def fragebogen2(Ordnername, i):
    # Ist ein Wert in Zelle B57 eingetragen?
    if pd.notna(df.iloc[55, 1]):
        i[1] = datetime.now().date()
    else:
        return f"{Ordnername} Fragebogen von Initiator:in noch nicht zurückerhalten!"
        # Sind seit dem Aushändigen des Fragebogens mehr als 2 Wochen vergangen?
        if datetime.now().date() - df.iloc[54, 1].date() > timedelta(weeks=2):
            return "Es sind mehr als 2 Wochen seit dem Aushändigen des Fragebogens vergangen, bitte Initiator:in kontaktieren."

root = tk.Tk()
root.title("Aufgabenstatus")

frame = tk.Frame(root)
frame.pack(expand=True, fill='both')

baum_pro_gruppe = {}

def update_gui():
    for widget in frame.winfo_children():
        widget.destroy()

    hauptschleife()  # Aufgaben prüfen

    for Gruppe in Gruppen:
        gruppenname = Gruppe[0]
        Titel = gruppenname.split(" (")[0]
        MA = f"({gruppenname.split(" (")[1]}"
        todo_status_datei = os.path.join(VERZEICHNIS, gruppenname, "todo_status.json")
        
        if not os.path.exists(todo_status_datei):
            continue

        # Neuer Frame pro Gruppe (damit Titel + Tabelle zusammenbleiben)
        gruppen_frame = tk.Frame(frame, borderwidth=1, relief="solid", padx=5, pady=5)
        gruppen_frame.pack(side="left", expand=True, fill='both', padx=5, pady=5)

        label = tk.Label(gruppen_frame, text=f"{Titel}\n{MA}", font=("Arial", 12, "bold"))
        label.pack()

        tree = ttk.Treeview(gruppen_frame, columns=("Aufgabe", "Status"), show="headings", height=20)
        tree.heading("Aufgabe", text="Aufgabe")
        tree.heading("Status", text="Status")
        tree.column("Aufgabe", width=150)
        tree.column("Status", width=80)
        tree.pack(expand=True, fill='both')

        with open(todo_status_datei, 'r', encoding='utf-8') as f:
            daten = json.load(f)
            for eintrag in daten:
                status = "✅Erledigt" if eintrag[1] else "Offen"
                tree.insert("", "end", values=(eintrag[0], status))

        baum_pro_gruppe[gruppenname] = tree

        # === Rückmeldung anzeigen, falls vorhanden ===
        meldung = letzte_meldungen.get(gruppenname, "")
        if meldung:
            meldungs_label = tk.Label(
                gruppen_frame,
                text=meldung,
                fg="green",  # Farbe für Erfolgsmeldungen – kann dynamisch angepasst werden
                wraplength=180,
                justify="left"
            )
            meldungs_label.pack(pady=(5, 0))

    # Wiederhole nach 10 Minuten
    root.after(600000, update_gui)

todo_functions = {"Backup Mitarbeiter:in finden": backup,
                  "Mögliche Termine für GGG finden und mit Initiator:in vereinbaren": termin_GGG,
                  "Auf Text für Homepage / Social Media / Pressemitteilung warten": text_warten,
                  "Grünen Zettel Brigitte geben": zettel,
                  "Gruppe auf Homepage inserieren": homepage,
                  "Text an Sabine für Instagram senden": instagram,
                  "Pressemitteilung versenden": presse,
                  "Interessent:innen sammeln": interessenten,
                  "Termin für erstes Treffen vereinbaren": erstesTreffen,
                  "Konferenzraum Reservieren1": konferenzraum1,
                  "Interessent:innen informieren1": infoTreffen1,
                  "Anwesenheiten notieren1": anwesenheit1,
                  "Termin für zweites Treffen vereinbaren": zweitesTreffen,
                  "Konferenzraum Reservieren2": konferenzraum2,
                  "Interessen:innen informieren2": infoTreffen2,
                  "Anwesenheiten notieren2": anwesenheit2,
                  "Initiator:in bei Raumsuche unterstützen": raumsuche,
                  "Termin für drittes Treffen vereinbaren": drittesTreffen,
                  "Konferenzraum Reservieren3": konferenzraum3,
                  "Interessent:innen informieren3": infoTreffen3,
                  "Initiator:in Fragebogen zukommen lassen": fragebogen1,
                  "Anwesenheiten notieren3": anwesenheit3,
                  "Fragebogen zurückerhalten und in Datenbank einpflegen": fragebogen2
                }

Gruppen = finde_hendrik_ordner(VERZEICHNIS)

def hauptschleife():
    for Gruppe in Gruppen:
        gruppenname = Gruppe[0]
        print(gruppenname)
        todo_status_datei = os.path.join(VERZEICHNIS, gruppenname, "todo_status.json")

        if not os.path.exists(todo_status_datei):
            print(f"{gruppenname} todo_status.json nicht vorhanden, wird erstellt")
            with open(Aufgaben, 'r', encoding='utf-8') as f:
                zeilen = [zeile.strip() for zeile in f if zeile.strip()]
            daten = [[zeile, None] for zeile in zeilen]
            with open(todo_status_datei, 'w', encoding='utf-8') as f:
                json.dump(daten, f, indent=2, ensure_ascii=False)
        
        datei = list(Path(os.path.join(VERZEICHNIS, gruppenname)).glob('GG Verlauf*.xlsx'))
        if datei:
            pfad = datei[0]
            global df, wb, ws
            df = pd.read_excel(pfad)
            wb = load_workbook(pfad, data_only=True)
            ws = wb.active
        else:
            print(f"{gruppenname} keine Excel-Datei gefunden!")
            continue

        with open(todo_status_datei, 'r', encoding='utf-8') as f:
            daten = json.load(f)

        meldung = ""  # Rückgabewert der Funktion

        for i in daten:
            if i[1] is not None:
                continue
            else:
                meldung = todo_functions[i[0]](gruppenname, i)  # Funktion liefert Text zurück
                with open(todo_status_datei, 'w', encoding='utf-8') as f:
                    json.dump(daten, f, indent=2, ensure_ascii=False, default=str)
                break

        letzte_meldungen[gruppenname] = meldung  # Merken für die GUI
        
while __name__ == "__main__":
    # Starte initiales Update
    update_gui()

    # Starte Hauptloop
    root.mainloop()