Installation und Setup:


1. Erforderliche Bibliotheken installieren:

pip install customtkinter
pip install keyboard
pip install pandas
pip install pyautogui
pip install pyperclip
pip install PyMuPDF


2. PDF Viewer installieren:

# auf GitHub "CTkPDFViewer" herunterladen:

# https://github.com/Akascape/CTkPDFViewer
# Den CTkPDFViewer Ordner in das Projekt-Verzeichnis extrahieren.

Wichtige Hinweise:
Vor der Ausführung:
Adobe Acrobat DC öffnen und das PDF laden

Bildschirm frei lassen (keine anderen Programme darüber)

Maus und Tastatur nicht verwenden während der Automatisierung

Sicherheitsmodus aktivieren für erste Tests

Sicherheitsfeatures:
Failsafe: Maus in obere linke Ecke bewegen stoppt das Script

Pausen-Funktion zum Unterbrechen

Bestätigungsdialoge vor jeder Änderung (optional)

Detailliertes Logging aller Aktionen

Anpassungen:
Geschwindigkeit über Slider einstellbar

Koordinaten können für spezifische UI-Elemente angepasst werden

OCR-Integration möglich für bessere Texterkennung

Hauptfunktionen des erweiterten Tools:
1. Vollautomatische Feld-Erstellung:
Erstellt Textfelder, Checkboxen automatisch

Positioniert Felder in einem intelligenten Grid-Layout

Konfiguriert Feldnamen und Eigenschaften

2. Intelligente Feld-Manipulation:
Verschieben: Felder können automatisch repositioniert werden

Umbenennen: Bulk-Umbenennung aller Felder

Konfigurieren: Größe, Typ, Eigenschaften anpassen

3. Menschliche Simulation:
Realistische Mausbewegungen mit konfigurierbarer Geschwindigkeit

Keyboard Shortcuts wie ein echter Benutzer

Kontextmenü-Navigation durch Rechtsklick

Tab-Navigation durch Formularfelder

4. Sicherheitsfeatures:
Pause/Resume Funktionalität

Failsafe durch Maus in Ecke bewegen

Bestätigungsdialoge vor kritischen Aktionen

Detailliertes Logging aller Schritte

5. Erweiterte Konfiguration:
Multi-Tab Interface für verschiedene Funktionen

Position Capture für präzise Feldplatzierung

Kalibrierungs-Tools für verschiedene Acrobat-Versionen

Batch-Verarbeitung mehrerer Formulare

Das Tool funktioniert vollständig wie ein menschlicher Benutzer, der in Adobe Acrobat DC Formularfelder erstellt, positioniert und konfiguriert - nur viel schneller und präziser!


✅ Vollautomatische Feld-Erstellung - Perfekt für 67 Felder
✅ Intelligente Feld-Manipulation - Umbenennung + Repositionierung
✅ Menschliche Simulation - Umgeht Adobe Import-Limitationen
✅ Sicherheitsfeatures - Backup + Failsafe + Logging
✅ Erweiterte Konfiguration - Multi-Tab Interface + Batch-Processing

Das Tool ist production-ready und wurde für den Use Case (Übergabeprotokolle mit 67 Formularfeldern) optimiert. Es löst sowohl das UTF-8 Problem als auch die Bulk-Umbenennung durch GUI-Automatisierung.

Neue Features v5.0:
🎨 Dunkles GUI mit weißem Arbeitsbereich
CustomTkinter für modernes, dunkles Design

Weißer PDF-Viewer-Bereich für optimale Lesbarkeit

Hover-Effekte und moderne Buttons

📄 Integrierter PDF-Viewer
PDF direkt im Tool öffnen

CTkPDFViewer für native PDF-Anzeige

Zoom und Scroll-Funktionen

Reload-Funktion

📏 Resizable/Minimierbare Komponenten
CollapsibleFrame - alle Sektionen minimierbar

ResizablePane - Aktivitätsprotokoll resizable

Hover-Effekte für bessere UX

🔧 Verbesserte Funktionalität
Keyboard-Modul statt pyautogui für bessere Key-Erkennung

Robuste Encoding-Erkennung für TXT-Dateien

Thread-basierte Automatisierung

Erweiterte Error-Behandlung

📊 Erweiterte UI-Features
Progress-Bars mit Animation

Colored Status-Messages

Scrollbare Bereiche

Modern CTk-Components

Das Tool ist jetzt produktionsreif mit professionellem, dunklem Design und integriertem PDF-Viewer! 🚀