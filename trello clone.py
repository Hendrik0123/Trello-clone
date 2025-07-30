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
    regex = r"\(\s*[A-Za-z]{2,}\s*(\+|und)\s*[A-Za-z]{2,}\s*\)"
    treffer = re.search(regex, Ordnername)
    eintrag = df.iloc[5, 1]
    # Prüfung: Wert vorhanden & besteht aus mindestens 2 Wörtern
    if isinstance(eintrag, str) and len(eintrag.strip().split()) >= 2:
        eintrag = True
    else:
        eintrag = False  
    if treffer and eintrag:
        i[1] = datetime.now().date()
    elif treffer and not eintrag:
        print(f"{Ordnername} Backup nicht in Excel eingetragen!")
    elif eintrag and not treffer:
        print(f"{Ordnername} Backup nicht im Ordnernamen eingetragen!")        

def termin_GGG(Ordnername):
        GGG_Termin = ws["B5"].value.date()
        # Prüfung, ob ein Wert vorhanden ist
        if pd.notna(GGG_Termin):
            i[1] = datetime.now().date()
        
def text_warten(Ordnername):
    aktuellerpfad = Path(VERZEICHNIS) / Ordnername
    # Prüfen, ob mindestens ein Unterordner existiert
    unterordner = any(p.is_dir() for p in aktuellerpfad.iterdir())
    if unterordner:
        i[1] = datetime.now().date()
    elif not unterordner: 
        heute = datetime.now().date()
        GGG_Termin = ws["B5"].value.date()
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
    # ist etwas in Zelle B29 eingetragen?
    if pd.notna(df.iloc[27, 1]):
        excel = 1
    if homepage and excel or df.iloc[27, 1] == "-":
        i[1] = datetime.now().date()
    elif not homepage and excel:
        print(f"Gruppe {Ordnername} nicht auf Homepage gefunden! Stimmen Ordnername und Gruppenname überein?")
    elif homepage and not excel:
        print(f"Gruppe {Ordnername} auf Homepage gefunden aber kein Eintrag in Excel vorhanden!")
    else:
        print(f"Gruppe {Ordnername} weder auf Homepage noch in Excel gefunden! Bitte prüfen!")
    

def instagram(Ordnername):
    # ist etwas in Zelle B30 eingetragen?
    if pd.notna(df.iloc[29, 1]):
        i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} kein Eintrag in Excel für Instagram vorhanden! (Datum oder '-')")

def presse(Ordnername):
    # ist etwas in Zelle B31 eingetragen?
    print(f"Pressemitteilung {df.iloc[30, 0]}")
    if pd.notna(df.iloc[30, 1]):
        i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} kein Eintrag in Excel für Pressemitteilung vorhanden! (Datum oder '-')")
            
def interessenten(Ordnername): 
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
        print(f"Bei Gruppe {Ordnername} ist das Gründungsgespräch mehr als 2 Monate her und es sind weniger als 4 Interessent:innen auf der Liste. Initiator:in bezgl. weiterem Vorgehen kontaktieren.")
     

def erstesTreffen(Ordnername):   
    # Ist ein Datum in Zelle B21 eingetragen?
    if pd.notna(df.iloc[19, 1]):
        i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} Termin für erstes Treffen vereinbaren!") 

def konferenzraum1(Ordnername):
    a = input(f"Wurde der Konferenzraum für das erste Treffen am {df.iloc[19, 1].strftime("%d.%m.%Y")} reserviert? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        print(f"{Ordnername} Konferenzraum nicht reserviert!")
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()    

def infoTreffen1(Ordnername):
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
        print(f"{Ordnername} wurden folgende Interessent:innen über den ersten Termin am {df.iloc[19, 1].strftime("%d.%m.%Y")} informiert?: \n-{'\n-'.join(nicht_informiert)} \nes ist kein Haken in der Interessiertenliste gesetzt!")
        
def anwesenheit1(Ordnername):
    # liegt das erste Treffen in der Vergangenheit?
    if pd.isna(df.iloc[19, 1]):
        print(f"{Ordnername} Termin für erstes Treffen nicht eingetragen!")
        return
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
            print(f"{Ordnername} Anzahl in Zelle D21 eingetragen aber keine Anwesenheiten abgehakt!")
        elif not anzahl and checks:
            print(f"{Ordnername} Anwesenheiten abgehakt aber keine Anzahl in Zelle D21 eingetragen!")
        else:
            print(f"{Ordnername} weder Anzahl in Zelle D21 eingetragen noch Anwesenheiten abgehakt!")
    else:
        print(f"{Ordnername} warten bis erstes Treffen am {df.iloc[19, 1].strftime("%d.%m.%Y")} stattgefunden hat!")

def zweitesTreffen(Ordnername):    
    # Ist ein Datum in Zelle B21 eingetragen?
    if pd.notna(df.iloc[20, 1]):
        i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} Termin für erstes Treffen vereinbaren!")

        
def konferenzraum2(Ordnername):
    a = input(f"Wurde der Konferenzraum für das zweite Treffen am {df.iloc[20, 1].strftime("%d.%m.%Y")} reserviert? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        print(f"{Ordnername} Konferenzraum nicht reserviert!")
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()  
        
def infoTreffen2(Ordnername):
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
        print(f"{Ordnername} wurden folgende Interessent:innen über den zweiten Termin am {df.iloc[20, 1].strftime("%d.%m.%Y")}  informiert?: \n-{'\n-'.join(nicht_informiert)} \nes ist kein Haken in der Interessiertenliste gesetzt!")
        
def anwesenheit2(Ordnername):
    # liegt das zweite Treffen in der Vergangenheit?
    if pd.isna(df.iloc[20, 1]):
        print(f"{Ordnername} Termin für zweites Treffen nicht eingetragen!")
        return
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
            print(f"{Ordnername} Anzahl in Zelle D22 eingetragen aber keine Anwesenheiten abgehakt!")
        elif not anzahl and checks:
            print(f"{Ordnername} Anwesenheiten abgehakt aber keine Anzahl in Zelle D22 eingetragen!")
        else:
            print(f"{Ordnername} weder Anzahl in Zelle D22 eingetragen noch Anwesenheiten abgehakt!")
    else:
        print(f"{Ordnername} warten bis zweites Treffen am {df.iloc[20, 1].strftime("%d.%m.%Y")} stattgefunden hat!")
      
def raumsuche(Ordnername):
    a = input("Hat die Gruppe einen eigenen Raum für weitere Treffen nach dem dritten Termin? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        print(f"{Ordnername} Es muss ein Raum für weitere Treffen gefunden werden!")
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()             
    
def drittesTreffen(Ordnername):
        # Ist ein Datum in Zelle B23 eingetragen?
    if pd.notna(df.iloc[21, 1]):
        i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} Termin für drittes Treffen vereinbaren!")

def konferenzraum3(Ordnername):
    a = input(f"Wurde der Konferenzraum für das dritte Treffen am {df.iloc[21, 1].strftime("%d.%m.%Y")} reserviert? (ja/nein): ").strip().lower()
    if a == "ja":
        i[1] = datetime.now().date()
    elif a == "nein":
        print(f"{Ordnername} Konferenzraum nicht reserviert!")
    while a not in ["ja", "nein"]:
        print("Ungültige Eingabe. Bitte 'ja' oder 'nein' eingeben.")
        a = input("Wurde der Konferenzraum reserviert? (ja/nein): ").strip().lower()  

def infoTreffen3(Ordnername):
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
        print(f"{Ordnername} wurden folgende Interessent:innen über den dritten Termin am {df.iloc[21, 1].strftime("%d.%m.%Y")} informiert?: \n-{'\n-'.join(nicht_informiert)} \nes ist kein Haken in der Interessiertenliste gesetzt!")

def fragebogen1(Ordnername):
    #Ist ein Wert in Zelle B56 eingetragen?
    if pd.notna(df.iloc[54, 1]):
        i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} Fragebogen an Initiator:in aushändigen!")

def anwesenheit3(Ordnername):
    # liegt das dritte Treffen in der Vergangenheit?
    if pd.isna(df.iloc[21, 1]):
        print(f"{Ordnername} Termin für drittes Treffen nicht eingetragen!")
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
            print(f"{Ordnername} Anzahl in Zelle D23 eingetragen aber keine Anwesenheiten abgehakt!")
        elif not anzahl and checks:
            print(f"{Ordnername} Anwesenheiten abgehakt aber keine Anzahl in Zelle D23 eingetragen!")
        else:
            print(f"{Ordnername} weder Anzahl in Zelle D23 eingetragen noch Anwesenheiten abgehakt!")
    else:
        print(f"{Ordnername} warten bis drittes Treffen am {df.iloc[21, 1].strftime("%d.%m.%Y")} stattgefunden hat!")

def fragebogen2(Ordnername):
    # Ist ein Wert in Zelle B57 eingetragen?
    if pd.notna(df.iloc[55, 1]):
        i[1] = datetime.now().date()
    else:
        print(f"{Ordnername} Fragebogen von Initiator:in noch nicht zurückerhalten!")
        # Sind seit dem Aushändigen des Fragebogens mehr als 2 Wochen vergangen?
        if datetime.now().date() - df.iloc[54, 1].date() > timedelta(weeks=2):
            print("Es sind mehr als 2 Wochen seit dem Aushändigen des Fragebogens vergangen, bitte Initiator:in kontaktieren.")

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
        print(Gruppe)
        # Ist eine "todo_status.json" Datei vorhanden?
        todo_status_datei = os.path.join(VERZEICHNIS, Gruppe[0], "todo_status.json")
        print("Suche nach Datei:", todo_status_datei)
        print("Existiert Datei?", os.path.exists(todo_status_datei))
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
        
        datei = list(Path(os.path.join(VERZEICHNIS, Gruppe[0])).glob('GG Verlauf*.xlsx'))
        if datei:
            pfad = datei[0]   
            global df, wb, ws          
            df = pd.read_excel(pfad)  # Direkt übergeben – pandas versteht Path-Objekte
            wb = load_workbook(pfad, data_only=True)
            ws = wb.active
        else:
            print(f"{Gruppe[0]} keine Excel-Datei gefunden!")
            continue        
        
        with open(todo_status_datei, 'r', encoding='utf-8') as f:
            daten = json.load(f)
            
            for i in daten:
                if i[1] != None:
                    continue
                else:
                    todo_functions[i[0]](Gruppe[0])                
                    # JSON-Datei direkt nach der Änderung aktualisieren
                    with open(todo_status_datei, 'w', encoding='utf-8') as f:
                        json.dump(daten, f, indent=2, ensure_ascii=False, default=str)
                    break
            continue
        
while __name__ == "__main__":
    hauptschleife()
    print("Alle Aufgaben überprüft und aktualisiert.")
    
    # Optional: GUI zur Anzeige der Ergebnisse
    root = tk.Tk()
    root.title("Aufgabenstatus")
    
    tree = ttk.Treeview(root, columns=("Aufgabe", "Status"), show='headings')
    tree.heading("Aufgabe", text="Aufgabe")
    tree.heading("Status", text="Status")
    
    for Gruppe in Gruppen:
        todo_status_datei = os.path.join(VERZEICHNIS, Gruppe[0], "todo_status.json")
        with open(todo_status_datei, 'r', encoding='utf-8') as f:
            daten = json.load(f)
            for i in daten:
                status = "Erledigt" if i[1] else "Offen"
                tree.insert("", "end", values=(i[0], status))
    
    tree.pack(expand=True, fill='both')
    root.mainloop()