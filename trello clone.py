import os
import re
import json
import tkinter as tk
from tkinter import ttk
import pandas as pd
import datetime

# Name der Datei mit den Aufgaben
AUFGABEN_DATEI = "aufgaben.txt"
# Arbeitsverzeichnis (hier: aktuelles Verzeichnis)
VERZEICHNIS = "."

def lade_aufgaben(dateipfad):
    """Liest Aufgabenzeilen aus einer Datei und gibt sie als Liste zurück."""
    with open(dateipfad, "r", encoding="utf-8") as f:
        return [zeile.strip() for zeile in f if zeile.strip()]

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

def gui_anzeigen(aufgaben, ordner):
    """Zeigt die Aufgaben und Ordner in einer grafischen Oberfläche an."""
    root = tk.Tk()
    root.title("Aufgabenübersicht")
    root.geometry("915x650")

    # Canvas mit horizontalem Scrollbalken für viele Spalten
    canvas = tk.Canvas(root)
    scrollbar = ttk.Scrollbar(root, orient="horizontal", command=canvas.xview)
    frame = ttk.Frame(canvas)
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(xscrollcommand=scrollbar.set)
    canvas.pack(side="top", fill="both", expand=True)
    scrollbar.pack(side="bottom", fill="x")

    # Ermöglicht horizontales Scrollen mit dem Mausrad
    def _on_mousewheel(event):
        canvas.xview_scroll(int(1*(event.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def save_status(ordnername, pfad, status_vars):
        """
        Speichert den Status (erledigt/ignoriert) für jede Aufgabe im Ordner als JSON-Datei.
        """
        status = []
        now = datetime.datetime.now().isoformat()
        for sv in status_vars:
            done_ts = sv['done_ts']
            done_var = sv['done_var']
            ignored_var = sv['ignored_var']
            # Wenn erledigt angehakt, setze Zeitstempel, sonst entferne ihn
            if done_var.get():
                if done_ts is None:
                    done_ts = now
            else:
                done_ts = None
            status.append({
                "done": done_ts,
                "ignored": ignored_var.get()
            })
            sv['done_ts'] = done_ts
        with open(os.path.join(pfad, "todo_status.json"), "w", encoding="utf-8") as f:
            json.dump(status, f, indent=2)

    # Sortiere: primär zuerst, dann sekundär
    ordner = sorted(ordner, key=lambda x: 1 if x[1] == "sekundär" else 0)
    primär = [o for o in ordner if o[1] == "primär"]
    sekundär = [o for o in ordner if o[1] == "sekundär"]

    def count_done(o):
        """Zählt erledigte Aufgaben im Ordner o."""
        pfad = os.path.join(VERZEICHNIS, o[0])
        status_file = os.path.join(pfad, "todo_status.json")
        if os.path.exists(status_file):
            with open(status_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return sum(1 for item in data if item.get("done"))
        return 0

    # Sortiere Ordner nach Anzahl erledigter Aufgaben (aufsteigend)
    primär = sorted(primär, key=count_done)
    sekundär = sorted(sekundär, key=count_done)

    col_primär = 0
    col_sekundär = 0

    def erstelle_spalte(o, row, col):
        """
        Erstellt eine Spalte für einen Ordner mit Checkboxen für jede Aufgabe.
        """
        pfad = os.path.join(VERZEICHNIS, o[0])
        status_file = os.path.join(pfad, "todo_status.json")

        # Lade bisherigen Status oder initialisiere neu
        if os.path.exists(status_file):
            with open(status_file, "r", encoding="utf-8") as f:
                status = json.load(f)
        else:
            status = []

        # Passe Status-Liste an die Anzahl der Aufgaben an
        while len(status) < len(aufgaben):
            status.append({"done": None, "ignored": False})
        if len(status) > len(aufgaben):
            status = status[:len(aufgaben)]

        # Erstelle Rahmen für die Spalte
        col_frame = ttk.LabelFrame(frame, text=f"{o[0]} ({o[1]})", padding=10)
        col_frame.grid(row=row, column=col, padx=10, pady=10, sticky="n")

        status_vars = []
        for idx, aufgabe in enumerate(aufgaben):
            done_ts = status[idx].get("done")
            ignored = status[idx].get("ignored", False)
            done_var = tk.BooleanVar(value=(done_ts is not None))
            ignored_var = tk.BooleanVar(value=ignored)

            def make_cmd(i=idx):
                # Callback für "Erledigt"-Checkbox
                def cmd():
                    now = datetime.datetime.now().isoformat()
                    sv = status_vars[i]
                    if sv['done_var'].get():
                        if sv['done_ts'] is None:
                            sv['done_ts'] = now
                    else:
                        sv['done_ts'] = None
                    save_status(o[0], pfad, status_vars)
                return cmd

            # Layout: Aufgabe + zwei Checkboxen
            task_frame = ttk.Frame(col_frame)
            task_frame.pack(anchor="w", pady=2)

            label = ttk.Label(task_frame, text=aufgabe, width=40)
            label.pack(side="left")

            cb_done = ttk.Checkbutton(
                task_frame, text="Erledigt", variable=done_var, command=make_cmd()
            )
            cb_done.pack(side="left", padx=5)

            cb_ignored = ttk.Checkbutton(task_frame, text="Ignorieren", variable=ignored_var,
                command=lambda i=idx: save_status(o[0], pfad, status_vars))

            cb_ignored.pack(side="left", padx=5)

            status_vars.append({
                "done_var": done_var,
                "ignored_var": ignored_var,
                "done_ts": done_ts
            })
        return status_vars

    # Erstelle Spalten für primäre Ordner (oben)
    for o in primär:
        erstelle_spalte(o, 0, col_primär)
        col_primär += 1

    # Erstelle Spalten für sekundäre Ordner (unten)
    for o in sekundär:
        erstelle_spalte(o, 1, col_sekundär)
        col_sekundär += 1

    root.mainloop()

def main():
    """Hauptfunktion: lädt Aufgaben und Ordner, startet die GUI."""
    aufgaben = lade_aufgaben(AUFGABEN_DATEI)
    ordner = finde_hendrik_ordner(VERZEICHNIS)
    if not ordner:
        print("Keine passenden Ordner gefunden.")
        return
    gui_anzeigen(aufgaben, ordner)

if __name__ == "__main__":
    main()
