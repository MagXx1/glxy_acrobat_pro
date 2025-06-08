import sys
import os

# Füge das src-Verzeichnis zum Python-Pfad hinzu
sys.path.insert(0, os.path.join(os.path.dirname(__file__)))

from gui.main_window import AcrobatFormAutomatorPro

def main():
    """Hauptfunktion - startet die Anwendung"""
    print("🚀 Starte Roboti x Acrobat x GLXY...")
    
    try:
        app = AcrobatFormAutomatorPro()
        app.mainloop()
    except Exception as e:
        print(f"❌ Fehler beim Starten der Anwendung: {e}")
        input("Drücken Sie Enter zum Beenden...")

if __name__ == "__main__":
    main()
