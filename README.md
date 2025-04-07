# HERMES_Logo_changer

## Beschreibung

Dieses Python-Skript verwendet `tkinter` für eine grafische Benutzeroberfläche (GUI) und ermöglicht das Ersetzen eines Logos in mehreren Microsoft Word-Vorlagendateien (.dotx) auf einmal.

### Hauptfunktionen

* **Einfache GUI:** Benutzerfreundliche Oberfläche zum Auswählen von Vorlagenordner und Logo-Datei.
* **Stapelverarbeitung:** Ersetzt das Logo in allen .dotx-Dateien eines angegebenen Ordners.
* **Fortschrittsanzeige:** Zeigt den Fortschritt der Bearbeitung mit einem Fortschrittsbalken und Statusmeldungen an.
* **Ergebnisprotokollierung:** Speichert Ergebnisse und Fehler in einem Textfeld zur Überprüfung.
* **Fehlerbehandlung:** Robuste Fehlerbehandlung für Dateioperationen und Benutzerfehler.
* **Threading:** Verwendet Threads, um die GUI während der Bearbeitung reaktionsfähig zu halten.

## Installation

Stelle sicher, dass Python 3 und `pip` installiert sind. Folge dann diesen Schritten:

1.  Klone das Repository:

    ```bash
    git clone https://github.com/FlownFlown/HERMES_Logo_changer.git
    cd HERMES_Logo_changer
    ```

2.  Installiere die erforderlichen Python-Pakete:

    ```bash
    pip install tkinter ttkbootstrap
    ```

## Verwendung

1.  Starte das Skript:

    ```bash
    python HERMES_Logo_change_V1.py
    ```

2.  Wähle im Fenster den Ordner aus, der deine .dotx-Vorlagendateien enthält.
3.  Wähle die .png-Datei aus, die als neues Logo verwendet werden soll.
4.  Klicke auf „Logos ersetzen“, um die Bearbeitung zu starten.
5.  Überwache den Fortschritt und die Ergebnisse im Fenster.

## Voraussetzungen

* Python 3.x
* `tkinter`
* `ttkbootstrap`

## Fehlerbehebung

* **Fehler beim Umbenennen des Logos**: Stelle sicher, dass die Logo-Datei nicht von einem anderen Programm verwendet wird und dass du Schreibrechte im Zielordner hast.
* **Fehler beim Bearbeiten von .dotx-Dateien**: Überprüfe, ob die .dotx-Dateien gültig sind und nicht beschädigt sind.

## Lizenz

Die Verwendung benötigt keine Lizenz.
