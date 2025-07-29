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
VERZEICHNIS = "."

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



def backup(Ordnername):
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    if datei:
        regex = r"\(\s*[A-Za-z]{2,}\s*\+\s*[A-Za-z]{2,}\s*\)"
        treffer = re.search(regex, Ordnername)
        pfad = datei[0]
        df = pd.read_excel(pfad)
        eintrag = df.iloc[4, 1]
        # Prüfung: Wert vorhanden & besteht aus mindestens 2 Wörtern
        if isinstance(eintrag, str) and len(eintrag.strip().split()) >= 2:
            eintrag = True
        else:
            eintrag = False
                
        if treffer and eintrag:
            print(f"{Ordnername} Backup ist da!")
            i[1] = datetime.now().date()
        elif treffer and not eintrag:
            print(f"{Ordnername} Backup nicht in Excel eingetragen!")
        elif eintrag and not treffer:
            print(f"{Ordnername} Backup nicht im Ordnernamen eingetragen!")
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")
        

def termin_GGG(Ordnername):
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte
        global GGG_Termin
        GGG_Termin = df.iloc[3, 1].date()
        # Prüfung, ob ein Wert vorhanden ist
        if pd.notna(GGG_Termin):
            i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")
        
def text_warten(Ordnername):
    pfad = Path(VERZEICHNIS) / Ordnername
    # Prüfen, ob mindestens ein Unterordner existiert
    unterordner = any(p.is_dir() for p in pfad.iterdir())
    if unterordner:
        i[1] = datetime.now().date()
    elif not unterordner: 
        heute = datetime.now().date()
        
        if GGG_Termin is None:
            print(f"{Ordnername} Termin GGG nicht in Excel gefunden!")
        zwei_wochen = timedelta(weeks=2)
        if heute - GGG_Termin > zwei_wochen:
            print("Es sind mehr als 2 Wochen seit dem GGG-Termin vergangen, bitte Initiator:in kontaktieren.")

def zettel(Ordnername):
    a = input("Wurde der grüne Zettel Brigitte gegeben? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        print(f"{Ordnername} grünen Zettel nicht gegeben!")
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der grüne Zettel Brigitte gegeben? (ja/nein): ").strip().lower()
        
def homepage(Ordnername):
    homepage = 0
    excel = 0
    Name = Ordnername.split(" (")[0]
    Gruppen_in_Gruendung = get_titles_from_url(os.getenv("url"))
    if Name in Gruppen_in_Gruendung:
        homepage = 1
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte  
        # ist etwas in Zelle B29 eingetragen?  
        if pd.notna(df.iloc[29, 1]):
            excel = 1
        if homepage and excel:
            i[1] = datetime.now().date()
        elif not homepage and excel:
            print(f"Gruppe {Ordnername} nicht auf Homepage gefunden! Stimmen Ordnername und Gruppenname überein?")
        elif homepage and not excel:
            print(f"Gruppe {Ordnername} auf Homepage gefunden aber kein Eintrag in Excel vorhanden!")
        else:
            print(f"Gruppe {Ordnername} weder auf Homepage noch in Excel gefunden! Bitte prüfen!")
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")
    

def instagram(Ordnername):
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte
        # ist etwas in Zelle B30 eingetragen?
        if pd.notna(df.iloc[30, 1]):
            i[1] = datetime.now().date()
        else:
            print(f"{Ordnername} kein Eintrag in Excel für Instagram vorhanden!")
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")

def presse(Ordnername):
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte
        # ist etwas in Zelle B31 eingetragen?
        if pd.notna(df.iloc[31, 1]):
            i[1] = datetime.now().date()
        else:
            print(f"{Ordnername} kein Eintrag in Excel für Pressemitteilung vorhanden!")
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")
            
def interessenten(Ordnername): 
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    excel_namen = [
    # Spalte H (Index 7)
    (3, 7), (17, 7), (31, 7), (45, 7),
    # Spalte R (Index 17)
    (3, 17), (17, 17), (31, 17), (45, 17),
    # Spalte AB (Index 27)
    (3, 27), (17, 27), (31, 27), (45, 27),
    # Spalte AL (Index 37)
    (3, 37), (17, 37), (31, 37), (45, 37),
    # Spalte AV (Index 47)
    (3, 47), (17, 47), (31, 47), (45, 47),
    # Spalte BF (Index 57)
    (3, 57), (17, 57), (31, 57), (45, 57),
    ]
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte
        anzahl = 0
        heute = datetime.now().date()
        zwei_monate = timedelta(weeks=2)
        for interessent in excel_namen:
            if pd.notna(df.iloc[interessent]):
                anzahl += 1
        if anzahl >= 4:
            i[1] = datetime.now().date()
        if anzahl < 4 and heute - GGG_Termin > zwei_monate:
            print(f"Bei Gruppe {Ordnername} ist das Gründungsgespräch mehr als 2 Monate her und es sind weniger als 4 Interessent:innen auf der Liste. Initiator:in bezgl. weiterem Vorgehen kontaktieren.")
            
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")      

def erstesTreffen(Ordnername):
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte    
        # Ist ein Datum in Zelle B21 eingetragen?
        if pd.notna(df.iloc[20, 1]):
            i[1] = datetime.now().date()
        else:
            print(f"{Ordnername} Termin für erstes Treffen vereinbaren!")
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")   

def konferenzraum(Ordnername):
    a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        print(f"{Ordnername} Konferenzraum nicht reserviert!")
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()    

def infoTreffen1(Ordnername):
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    excel_namen = [
    # Spalte H
    "H4", "H18", "H32", "H46",
    # Spalte R
    "R4", "R18", "R32", "R46",
    # Spalte AB
    "AB4", "AB18", "AB32", "AB46",
    # Spalte AL
    "AL4", "AL18", "AL32", "AL46",
    # Spalte AV
    "AV4", "AV18", "AV32", "AV46",
    # Spalte BF
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
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte
        wb = load_workbook(pfad, data_only=True)
        ws = wb.active 
        nicht_informiert = []
        infoCount = 0
        for interessent in excel_namen:
            print(ws[interessent].value)
            if ws[interessent].value is not None:
                if ws[infoCheck[infoCount]].value == False:
                    nicht_informiert.append(ws[interessent].value)
                    print(f"{ws[interessent].value} nicht informiert!")
            print(ws[infoCheck[infoCount]].value)
            infoCount += 1
        if nicht_informiert == []:
            i[1] = datetime.now().date()
        else:
            print(f"{Ordnername} folgende Interessent:innen wurden über den ersten Termin NICHT informiert: {', '.join(nicht_informiert)}")
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")
        
def anwesenheit1(Ordnername):
    datei = list(Path(os.path.join(VERZEICHNIS, Ordnername)).glob('GG Verlauf*.xlsx'))
    if datei:
        pfad = datei[0]
        df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte
        # liegt das erste Treffen in der Vergangenheit?
        erstesTreffen = df.iloc[20, 1].date()
        heute = datetime.now().date()
        if erstesTreffen < heute:
            # ist ein Wert in Zelle D21 eingetragen?
            if pd.notna(df.iloc[20, 3]):
                i[1] = datetime.now().date()
            else:
                print(f"{Ordnername} Anwesenheiten für erstes Treffen nicht eingetragen!")
        else:
            print(f"{Ordnername} warten bis erstes Treffen stattgefunden hat!")
    else:
        print(f"{Ordnername} keine Excel-Datei gefunden!")
    
        
todo_functions = {"Backup Mitarbeiter:in finden": backup,
                  "Mögliche Termine für GGG finden und mit Initiator:in vereinbaren": termin_GGG,
                  "Auf Text für Homepage / Social Media / Pressemitteilung warten": text_warten,
                  "Grünen Zettel Brigitte geben": zettel,
                  "Gruppe auf Homepage inserieren": homepage,
                  "Text an Sabine für Instagram senden": instagram,
                  "Pressemitteilung versenden": presse,
                  "Interessent:innen sammeln": interessenten,
                  "Termin für erstes Treffen vereinbaren": erstesTreffen,
                  "Konferenzraum Reservieren": konferenzraum,
                  "Interessent:innen informieren1": infoTreffen1,
                  "Anwesenheiten notieren1": anwesenheit1,
                    }

Gruppen = finde_hendrik_ordner(VERZEICHNIS)

for Gruppe in Gruppen:
    # Ist eine "todo_status.json" Datei vorhanden?
    todo_status_datei = os.path.join(VERZEICHNIS, Gruppe[0], "todo_status.json")
    if os.path.exists(todo_status_datei):
        print(f"{Gruppe[0]} todo_status.json vorhanden")
    else:
        print(f"{Gruppe[0]} todo_status.json nicht vorhanden, wird erstellt")
        # Erstelle json Datei aus aufgaben.txt
        with open(Aufgaben, 'r', encoding='utf-8') as f:
            zeilen = [zeile.strip() for zeile in f if zeile.strip()]  
        
        daten = [[zeile, None] for zeile in zeilen]
        with open(todo_status_datei, 'w', encoding='utf-8') as f:
            json.dump(daten, f, indent=2, ensure_ascii=False)    
            
    with open(todo_status_datei, 'r', encoding='utf-8') as f:
        daten = json.load(f)
        
        for i in daten:
            if i[1] != None:
                continue
            else:
                todo_functions[i[0]](Gruppe[0])
                #break
    continue
                
            