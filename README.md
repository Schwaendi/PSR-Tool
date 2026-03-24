# 📝 Custom PSR Tool (Problem Steps Recorder Alternative)

A lightweight, portable tool for automatically documenting workflows and clicks on your PC. 
With Microsoft phasing out the classic "Problem Steps Recorder" (PSR) in Windows, this tool offers the perfect, modern alternative – exporting your steps directly into an editable Word document!

Developed by **Sacha Schwendimann**.

## ✨ Key Features

* **Word Export instead of HTML:** Every click is saved live into a clean `.docx` document. No data loss if the system crashes!
* **Smart Click Markers:** Clicks are highlighted directly on the screenshot (🔴 Red = Left click, 🟠 Orange = Right click) along with a visible mouse cursor.
* **Active (Sub-)Window Detection:** Optionally captures only the specific sub-window or dialog box you clicked in, instead of your entire ultra-wide monitor.
* **⏱️ Long Click Mode:** Only records clicks held for more than 0.5 seconds (with audio feedback). Perfect for preventing accidental clicks from cluttering your guide.
* **🔢 Multi-Click Mode (SHIFT):** Hold the `SHIFT` key to collect multiple clicks. Release it, and the tool captures a single photo with your clicks neatly numbered (1, 2, 3...)!
* **Detail Magnifier (Optional):** Automatically inserts a zoomed-in crop of the exact click area next to the full screenshot.
* **Ninja Mode:** The control panel temporarily hides itself in a fraction of a second when taking a photo so it never blocks your screenshots.
* **Bilingual:** The user interface and generated Word documents can be switched between English and German with a single click.

## 🚀 How to Use

1. Download the `mein_psr.exe` from the Releases tab (no installation required).
2. Launch the program. The control panel will always stay on top.
3. Click on **"▶️ New Recording"** and perform your workflow.
4. Click on **"💾 Finish Recording"**. 
5. Done! Your finished Word document and the original images are now waiting for you in the automatically generated `screenshots` folder (right next to the `.exe`).

## 📜 Version History (Changelog)

* **v1.13 (Current):** * Added Multi-Click Mode (SHIFT) to number multiple steps in a single image.
  * Renamed main export folder to `screenshots`.
  * Completely overhauled and added scrollable help texts.
* **v1.12:** UX Update (Audio feedback now plays exactly at the 0.5s mark).
* **v1.11:** Limited audio feedback to the active long-click mode. Integrated version history.
* **v1.10:** Introduced long-click functionality, smart sub-window detection, and Ninja mode (invisible UI during capture).


DEUTSCH

# 📝 Custom PSR Tool (Problem Steps Recorder Alternative)

Ein leichtgewichtiges, portables Tool zur automatischen Dokumentation von Arbeitsschritten und Klicks am PC. 
Da Microsoft den klassischen "Problem Steps Recorder" (PSR) in Windows auslaufen lässt, bietet dieses Tool die perfekte, moderne Alternative – direkt exportiert als editierbares Word-Dokument!

Entwickelt von **Sacha Schwendimann**.

## ✨ Hauptfunktionen

* **Word-Export statt HTML:** Jeder Klick wird live in ein sauberes `.docx` Dokument geschrieben. Kein Datenverlust bei Abstürzen!
* **Intelligente Klick-Marker:** Klicks werden direkt auf dem Screenshot markiert (🔴 Rot = Linksklick, 🟠 Orange = Rechtsklick) und mit einem künstlichen Mauszeiger versehen.
* **(Unter-)Fenster Spion:** Das Tool fotografiert auf Wunsch nicht den ganzen Bildschirm, sondern intelligent nur das kleine Fenster oder Menü, in dem du gerade klickst.
* **⏱️ Langklick-Modus:** Zeichnet nur Klicks auf, die länger als 0.5 Sekunden gehalten werden (inklusive akustischem Feedback). Perfekt, um ungewollte Klicks beim Arbeiten aus dem Protokoll herauszuhalten.
* **🔢 Multi-Klick Modus (SHIFT):** Halte die `SHIFT`-Taste gedrückt, um mehrere Klicks zu sammeln. Lass sie los, und das Tool erstellt ein einzelnes Foto mit durchnummerierten Klicks (1, 2, 3...)!
* **Detail-Lupe (Optional):** Neben dem großen Screenshot wird automatisch ein herangezoomter Ausschnitt des Klickbereichs eingefügt.
* **Ninja-Modus:** Das Menüfenster versteckt sich für den Bruchteil einer Sekunde selbst, wenn ein Foto gemacht wird, damit es nie auf dem Screenshot stört.
* **Zweisprachig:** Die Benutzeroberfläche kann per Knopfdruck zwischen Deutsch und Englisch umgeschaltet werden (inkl. übersetzter Word-Dokumente).

## 🚀 Nutzung

1. Lade dir die `mein_psr.exe` aus den Releases herunter (keine Installation notwendig).
2. Starte das Programm. Das Menü bleibt immer im Vordergrund.
3. Klicke auf **"▶️ Neue Aufzeichnung"** und führe deine Arbeitsschritte aus.
4. Klicke auf **"💾 Anleitung abschließen"**. 
5. Fertig! Dein Word-Dokument und die Originalbilder liegen nun im automatisch erstellten Ordner `screenshots` (direkt neben dem Tool).

## 📜 Versionshistorie (Changelog)

* **v1.13 (Aktuell):** * Multi-Klick Modus (SHIFT) hinzugefügt, um Schritte in einem Bild zu nummerieren.
  * Hauptordner für Exporte in `screenshots` umbenannt.
  * Hilfe-Texte komplett überarbeitet und scrollbar gemacht.
* **v1.12:** UX-Update (Audio-Feedback ertönt exakt nach 0.5s Haltezeit).
* **v1.11:** Audio-Feedback auf den aktiven Langklick-Modus beschränkt. Versionshistorie integriert.
* **v1.10:** Langklick-Funktion, intelligente Unterfenster-Erkennung und Ninja-Modus (unsichtbares Menü) eingeführt.
