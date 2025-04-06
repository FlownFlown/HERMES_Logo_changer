import tkinter as tk
from tkinter import filedialog, messagebox
import os
import zipfile
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import threading
import logging

# Logging konfigurieren
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def dateien_bearbeiten_thread(vorlagen_ordner, logo_pfad, fortschritt_var, status_label, ergebnisse_text, ergebnisse_button):
    """
    Führt die Dateibearbeitung in einem separaten Thread aus, um die GUI nicht zu blockieren.

    Args:
        vorlagen_ordner (str): Pfad zum Ordner mit den Vorlagendateien (.dotx).
        logo_pfad (str): Pfad zur Logo-Datei (.png).
        fortschritt_var (tk.IntVar): Variable für den Fortschrittsbalken.
        status_label (ttk.Label): Label zur Anzeige des aktuellen Status.
        ergebnisse_text (tk.Text): Textfeld zur Anzeige der Ergebnisse.
        ergebnisse_button (ttk.Button): Button, um das Ergebnis-Fenster anzuzeigen.
    """
    logo_umbenannt = False  # Flag, ob das Logo erfolgreich umbenannt wurde
    neuer_logo_pfad = ""  # Pfad zum temporär umbenannten Logo

    try:
        # Logo in "image1.png" umbenennen, wie in den .dotx-Dateien erwartet
        neuer_logo_pfad = os.path.join(os.path.dirname(logo_pfad), "image1.png")

        # Prüfen, ob die temporäre Datei bereits existiert und ggf. überschreiben
        if os.path.exists(neuer_logo_pfad):
            logging.warning(f"Überschreibe existierende temporäre Logo-Datei: {neuer_logo_pfad}")
        os.rename(logo_pfad, neuer_logo_pfad)
        logo_umbenannt = True
        logging.info(f"Logo umbenannt zu: {neuer_logo_pfad}")

    except FileNotFoundError:
        messagebox.showerror("Fehler", f"Logo-Datei nicht gefunden: {logo_pfad}")
        logging.error(f"Logo-Datei nicht gefunden: {logo_pfad}")
        start_button.config(state=NORMAL)
        progress_bar.stop()
        status_label.config(text="Fehler: Logo nicht gefunden.")
        return

    except PermissionError as e:
        messagebox.showerror("Fehler", f"Keine Berechtigung zum Umbenennen des Logos: {e}")
        logging.error(f"Keine Berechtigung zum Umbenennen des Logos: {e}")
        start_button.config(state=NORMAL)
        progress_bar.stop()
        status_label.config(text="Fehler: Berechtigungsproblem.")
        return

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Umbenennen des Logos: {e}")
        logging.error(f"Fehler beim Umbenennen des Logos: {e}")
        # Versuch, das Logo zurückzubenennen, falls ein Fehler auftritt
        if neuer_logo_pfad and os.path.exists(neuer_logo_pfad) and not logo_umbenannt:
            try:
                os.rename(neuer_logo_pfad, logo_pfad)
                logging.info("Logo nach Fehler zurückbenannt.")
            except Exception as rename_back_error:
                logging.error(f"Konnte Logo nach Fehler nicht zurückbenennen: {rename_back_error}")
        start_button.config(state=NORMAL)
        progress_bar.stop()
        status_label.config(text="Fehler beim Logo-Umbenennen.")
        return

    dateien = [f for f in os.listdir(vorlagen_ordner) if f.endswith(".dotx")]
    gesamt_dateien = len(dateien)
    fortschritt_var.set(0)
    status_label.config(text="Bearbeitung gestartet...")
    ergebnisse_text.delete(1.0, tk.END)  # Ergebnisse zurücksetzen
    ergebnisse = []

    for i, dateiname in enumerate(dateien):
        vorlage_pfad = os.path.join(vorlagen_ordner, dateiname)
        fortschritt = int((i + 1) / gesamt_dateien * 100)
        status_label.config(text=f"Bearbeite: {dateiname} ({fortschritt}%)")
        fenster.update_idletasks()  # GUI aktualisieren, um den Status anzuzeigen
        try:
            bearbeite_dotx(vorlage_pfad, neuer_logo_pfad)
            fortschritt_var.set(fortschritt)
            ergebnisse.append(f"Erfolgreich bearbeitet: {dateiname}")
            logging.info(f"Erfolgreich bearbeitet: {dateiname}")
        except Exception as e:
            ergebnisse.append(f"Fehler beim Bearbeiten von {dateiname}: {e}")
            logging.error(f"Fehler beim Bearbeiten von {dateiname}: {e}")

    # Logo zurückbenennen, falls erfolgreich umbenannt
    if logo_umbenannt:
        try:
            # Sicherheitscheck, falls die Originaldatei unerwartet existiert
            if os.path.exists(logo_pfad):
                logging.warning(f"Original-Logo-Datei '{logo_pfad}' existiert unerwartet. Überschreibe...")
            os.rename(neuer_logo_pfad, logo_pfad)
            ergebnisse.append("Logo erfolgreich zurückbenannt.")
            logging.info("Logo erfolgreich zurückbenannt.")
        except Exception as e:
            ergebnisse.append(f"FEHLER beim Zurückbenennen des Logos: {e}")
            logging.error(f"FEHLER beim Zurückbenennen des Logos von {neuer_logo_pfad} zu {logo_pfad}: {e}")
            messagebox.showwarning("Warnung", f"Das Logo konnte nicht automatisch zurückbenannt werden!\nBitte benennen Sie '{neuer_logo_pfad}' manuell zurück zu '{os.path.basename(logo_pfad)}'.\nFehler: {e}")

    status_label.config(text="Bearbeitung abgeschlossen!")
    messagebox.showinfo("Fertig", "Logo-Ersetzung abgeschlossen!")

    # Ergebnisse im Textfeld anzeigen
    for ergebnis in ergebnisse:
        ergebnisse_text.insert(tk.END, ergebnis + "\n")

    progress_bar.stop()
    fortschritt_var.set(0)
    fenster.update_idletasks()
    ergebnisse_button.config(state=NORMAL)
    fenster.ergebnis_fenster.deiconify()  # Ergebnis-Fenster anzeigen

    start_button.config(state=NORMAL)  # Start-Button wieder aktivieren

def dateien_bearbeiten():
    """Startet die Dateibearbeitung in einem separaten Thread."""
    vorlagen_ordner = vorlagen_ordner_var.get()
    logo_pfad = logo_pfad_var.get()

    # Fehlerprüfung für die ausgewählten Pfade
    if not vorlagen_ordner or not logo_pfad:
        messagebox.showerror("Fehler", "Bitte beide Pfade auswählen.")
        return
    if not os.path.isdir(vorlagen_ordner):
        messagebox.showerror("Fehler", f"Vorlagenordner nicht gefunden:\n{vorlagen_ordner}")
        return
    if not os.path.isfile(logo_pfad):
        messagebox.showerror("Fehler", f"Logo-Datei nicht gefunden:\n{logo_pfad}")
        return

    start_button.config(state=DISABLED)
    ergebnisse_button.config(state=DISABLED)
    ergebnisse_text.delete(1.0, tk.END)  # Alte Ergebnisse löschen
    progress_bar.start()
    status_label.config(text="Starte Bearbeitung...")

    # Thread starten, um die Bearbeitung im Hintergrund durchzuführen
    threading.Thread(target=dateien_bearbeiten_thread, args=(vorlagen_ordner, logo_pfad, fortschritt_var, status_label, ergebnisse_text, ergebnisse_button), daemon=True).start()

def bearbeite_dotx(vorlage_pfad, neues_logo_pfad):
    """Ersetzt das Logo in einer einzelnen .dotx-Datei."""
    temp_pfad = vorlage_pfad + ".temp"

    try:
        # Prüfen, ob Quell- und Logo-Dateien existieren
        if not os.path.exists(vorlage_pfad):
            raise FileNotFoundError(f"Quelldatei nicht gefunden: {vorlage_pfad}")
        if not os.path.exists(neues_logo_pfad):
            raise FileNotFoundError(f"Temporäre Logo-Datei nicht gefunden: {neues_logo_pfad}")

        with zipfile.ZipFile(vorlage_pfad, 'r') as vorhandene_zip:
            # Prüfen, ob das Logo im Archiv existiert
            try:
                vorhandene_zip.getinfo("word/media/image1.png")
            except KeyError:
                logging.warning(f"'word/media/image1.png' nicht gefunden in: {os.path.basename(vorlage_pfad)}. Datei übersprungen.")
                raise FileNotFoundError(f"'word/media/image1.png' nicht in {os.path.basename(vorlage_pfad)}")

            with zipfile.ZipFile(temp_pfad, 'w', compression=zipfile.ZIP_DEFLATED) as neue_zip:
                for item in vorhandene_zip.infolist():
                    if item.filename != "word/media/image1.png":
                        try:
                            data = vorhandene_zip.read(item.filename)
                            neue_zip.writestr(item, data)
                        except Exception as read_write_error:
                            logging.error(f"Fehler beim Lesen/Schreiben von {item.filename} in {os.path.basename(vorlage_pfad)}: {read_write_error}")
                            raise  # Fehler weitergeben

                neue_zip.write(neues_logo_pfad, "word/media/image1.png")

        # Originaldatei löschen und temporäre Datei umbenennen
        try:
            os.remove(vorlage_pfad)
        except OSError as e:
            logging.error(f"Konnte Originaldatei nicht löschen: {vorlage_pfad} - {e}")
            if os.path.exists(temp_pfad):
                os.remove(temp_pfad)
            raise

        try:
            os.rename(temp_pfad, vorlage_pfad)
        except OSError as e:
            logging.error(f"Konnte Temp-Datei nicht umbenennen zu: {vorlage_pfad} - {e}")
            raise

    except FileNotFoundError as e:
        logging.error(f"Datei-Fehler bei {os.path.basename(vorlage_pfad)}: {e}")
        if os.path.exists(temp_pfad):
            try:
                os.remove(temp_pfad)
            except OSError:
                logging.warning(f"Konnte temporäre Datei nicht löschen: {temp_pfad}")
        raise

    except zipfile.BadZipFile:
        logging.error(f"Zip-Datei beschädigt: {os.path.basename(vorlage_pfad)}")
        if os.path.exists(temp_pfad):
            try:
                os.remove(temp_pfad)
            except OSError:
                logging.warning(f"Konnte temporäre Datei nicht löschen: {temp_pfad}")
        raise

    except Exception as e:
        logging.error(f"Unerwarteter Fehler bei {os.path.basename(vorlage_pfad)}: {e}")
        if os.path.exists(temp_pfad):
            try:
                os.remove(temp_pfad)
            except OSError:
                logging.warning(f"Konnte temporäre Datei nicht löschen: {temp_pfad}")
        raise

def waehle_ordner(var):
    """Öffnet einen Ordnerauswahldialog."""
    ordner = filedialog.askdirectory()
    if ordner:
        var.set(ordner)

def waehle_datei(var):
    """Öffnet einen Dateiauswahldialog."""
    datei = filedialog.askopenfilename(filetypes=[("PNG Dateien", "*.png")])
    if datei:
        var.set(datei)

def ergebnis_fenster_erstellen():
    """Erstellt das Fenster zur Anzeige der Ergebnisse."""
    if hasattr(fenster, 'ergebnis_fenster') and fenster.ergebnis_fenster.winfo_exists():
        return fenster.ergebnis_fenster

    ergebnis_fenster = ttk.Toplevel(fenster)
    ergebnis_fenster.title("Ergebnisse")
    ergebnis_fenster.geometry("600x400")
    ergebnis_fenster.withdraw()

    ergebnisse_label = ttk.Label(ergebnis_fenster, text="Bearbeitungsergebnisse:", font=("Arial", 12, "bold"))
    ergebnisse_label.pack(pady=10, padx=10, fill=X)

    ergebnis_fenster.ergebnisse_text_widget = tk.Text(ergebnis_fenster, height=15, wrap="word")
    ergebnis_fenster.ergebnisse_text_widget.pack(pady=10, padx=10, fill=BOTH, expand=True)

    scrollbar = ttk.Scrollbar(ergebnis_fenster.ergebnisse_text_widget, orient="vertical", command=ergebnis_fenster.ergebnisse_text_widget.yview)
    ergebnis_fenster.ergebnisse_text_widget.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=RIGHT, fill=Y)

    schliessen_button = ttk.Button(ergebnis_fenster, text="Schliessen", command=ergebnis_fenster.withdraw)
    schliessen_button.pack(pady=10)

    ergebnis_fenster.protocol("WM_DELETE_WINDOW", ergebnis_fenster.withdraw)
    fenster.ergebnis_fenster = ergebnis_fenster
    global ergebnisse_text
    ergebnisse_text = ergebnis_fenster.ergebnisse_text_widget

    return ergebnis_fenster

# --- GUI erstellen ---
fenster = ttk.Window(title="HERMES Logo changer v1.1", themename="flatly")
fenster.minsize(500, 280)

vorlagen_ordner_var = tk.StringVar()
logo_pfad_var = tk.StringVar()
fortschritt_var = tk.IntVar()

# --- Widgets erstellen ---

vorlagen_frame = ttk.Frame(fenster)
vorlagen_frame.pack(fill=X, padx=10, pady=(10, 5))
vorlagen_label = ttk.Label(vorlagen_frame, text="Vorlagenordner:", width=16, anchor='w')
vorlagen_label.pack(side=LEFT, padx=(0, 5))
vorlagen_entry = ttk.Entry(vorlagen_frame, textvariable=vorlagen_ordner_var)
vorlagen_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
vorlagen_button = ttk.Button(vorlagen_frame, text="Durchsuchen...", command=lambda: waehle_ordner(vorlagen_ordner_var))
vorlagen_button.pack(side=RIGHT)

logo_frame = ttk.Frame(fenster)
logo_frame.pack(fill=X, padx=10, pady=5)
logo_label = ttk.Label(logo_frame, text="Logo-Datei (.png):", width=16, anchor='w')
logo_label.pack(side=LEFT, padx=(0, 5))
logo_entry = ttk.Entry(logo_frame, textvariable=logo_pfad_var)
logo_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
logo_button = ttk.Button(logo_frame, text="Durchsuchen...", command=lambda: waehle_datei(logo_pfad_var))
logo_button.pack(side=RIGHT)

fortschritt_frame = ttk.Frame(fenster)
fortschritt_frame.pack(fill=X, padx=10, pady=5)
fortschritt_frame.columnconfigure(0, weight=1)
status_label = ttk.Label(fortschritt_frame, text="Bereit", anchor=W, width=70)
status_label.grid(row=0, column=0, sticky='w', padx=5, pady=(5, 0))
progress_bar = ttk.Progressbar(fortschritt_frame, variable=fortschritt_var, mode="determinate")
progress_bar.grid(row=1, column=0, sticky='ew', padx=5, pady=(0, 5))

button_frame = ttk.Frame(fenster)
button_frame.pack(pady=10)
start_button = ttk.Button(button_frame, text="Logos ersetzen", command=dateien_bearbeiten, bootstyle=PRIMARY)
start_button.pack(pady=(0, 5))
ergebnisse_button = ttk.Button(button_frame, text="Ergebnisse anzeigen", command=lambda: fenster.ergebnis_fenster.deiconify(), bootstyle=INFO, state=DISABLED)
ergebnisse_button.pack(pady=(5, 0))

ergebnis_fenster_erstellen()

fenster.mainloop()
