import pyautogui
import time
import subprocess
import os
from ..utils.logger import Logger

class AcrobatController:
    def __init__(self):
        self.logger = Logger()
        pyautogui.FAILSAFE = True
        
    def focus_acrobat(self):
        """Bringt Adobe Acrobat DC in den Vordergrund"""
        self.logger.log("Fokussiere Adobe Acrobat DC...", "INFO")
        
        try:
            window_titles = ["Adobe Acrobat", "Acrobat", "Adobe Acrobat Reader", "Adobe Reader"]
            
            focused = False
            for title in window_titles:
                windows = pyautogui.getWindowsWithTitle(title)
                if windows:
                    windows[0].activate()
                    time.sleep(1)
                    focused = True
                    self.logger.log(f"Adobe Acrobat DC fokussiert (Titel: {title})", "SUCCESS")
                    break
                    
            if not focused:
                self.logger.log("Adobe Acrobat DC Fenster nicht gefunden", "WARNING")
                return False
            return True
                
        except Exception as e:
            self.logger.log(f"Fehler beim Fokussieren: {str(e)}", "ERROR")
            return False
            
    def launch_acrobat(self, pdf_path=None):
        """Startet Adobe Acrobat DC"""
        self.logger.log("Starte Adobe Acrobat DC...", "INFO")
        
        try:
            if pdf_path:
                # Öffne PDF direkt
                os.startfile(pdf_path)
                self.logger.log(f"Öffne PDF: {os.path.basename(pdf_path)}", "SUCCESS")
            else:
                # Versuche Acrobat zu starten
                acrobat_paths = [
                    r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
                ]
                
                started = False
                for path in acrobat_paths:
                    if os.path.exists(path):
                        subprocess.Popen([path])
                        started = True
                        self.logger.log(f"Adobe Acrobat DC gestartet von: {path}", "SUCCESS")
                        break
                        
                if not started:
                    self.logger.log("Adobe Acrobat DC nicht gefunden. Bitte manuell starten.", "WARNING")
                    return False
                    
            time.sleep(3)  # Warte bis Acrobat geladen ist
            return True
            
        except Exception as e:
            self.logger.log(f"Fehler beim Starten von Adobe Acrobat DC: {str(e)}", "ERROR")
            return False
            
    def enter_form_mode(self, speed=1.0):
        """Aktiviert den Formular-Bearbeitungsmodus"""
        self.logger.log("Aktiviere Formular-Bearbeitungsmodus...", "INFO")
        
        try:
            # Stelle sicher, dass Acrobat fokussiert ist
            if not self.focus_acrobat():
                return False
                
            pyautogui.PAUSE = speed
            
            # Methode 1: Keyboard Shortcut
            self.logger.log("Versuche Tastenkombination Shift+Ctrl+7...", "DEBUG")
            pyautogui.hotkey('shift', 'ctrl', '7')
            time.sleep(3)
            
            self.logger.log("Formular-Modus sollte jetzt aktiv sein", "SUCCESS")
            return True
            
        except Exception as e:
            self.logger.log(f"Fehler beim Aktivieren des Formular-Modus: {str(e)}", "ERROR")
            return False
            
    def save_pdf(self):
        """Speichert das PDF"""
        self.logger.log("Speichere PDF...", "INFO")
        
        try:
            if not self.focus_acrobat():
                return False
                
            pyautogui.hotkey('ctrl', 's')
            time.sleep(2)
            
            # Falls "Speichern unter" Dialog erscheint
            pyautogui.press('enter')  # Bestätigen
            time.sleep(1)
            
            self.logger.log("PDF gespeichert", "SUCCESS")
            return True
            
        except Exception as e:
            self.logger.log(f"Fehler beim Speichern: {str(e)}", "ERROR")
            return False
