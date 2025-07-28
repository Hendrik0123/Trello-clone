import os
from pathlib import Path
import re
import json
import tkinter as tk
from tkinter import ttk
import pandas as pd
from datetime import datetime, timedelta
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
        
    
        

todo_functions = {"Backup Mitarbeiter:in finden": backup,
                  "Mögliche Termine für GGG finden und mit Initiator:in vereinbaren": termin_GGG,
                  "Auf Text für Homepage / Social Media / Pressemitteilung warten": text_warten,
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
            if i[1] is not None:
                continue
            else:
                todo_functions[i[0]](Gruppe[0])
                
            