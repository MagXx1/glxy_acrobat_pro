import pyautogui
import time
import keyboard
import pyperclip
from ..utils.logger import Logger

class FieldOperations:
    def __init__(self):
        self.logger = Logger()
        self.field_positions = {}
        
    def create_field_at_position(self, x, y, field_type, field_name, display_name="", width=200, height=25):
        """Erstellt ein Feld an der angegebenen Position"""
        self.logger.log(f"Erstelle {field_type}: {field_name} bei ({x}, {y})", "INFO")
        
        try:
            # Wähle entsprechendes Tool
            if not self.select_field_tool(field_type):
                self.logger.log(f"Konnte Tool für {field_type} nicht auswählen", "WARNING")
                return False
                
            # Erstelle Feld durch Ziehen
            self.logger.log(f"Zeichne {field_type} mit Größe {width}x{height}", "DEBUG")
            
            pyautogui.moveTo(x, y, duration=0.3)
            time.sleep(0.2)
            
            pyautogui.mouseDown()
            pyautogui.dragTo(x + width, y + height, duration=0.8)
            pyautogui.mouseUp()
            
            time.sleep(1)
            
            # Konfiguriere Feld-Eigenschaften
            return self.configure_field_properties(field_name, display_name, field_type)
            
        except Exception as e:
            self.logger.log(f"Fehler beim Erstellen von Feld {field_name}: {str(e)}", "ERROR")
            return False
            
    def select_field_tool(self, field_type):
        """Wählt das entsprechende Tool basierend auf Feldtyp"""
        
        # Mapping von Feldtypen zu Keyboard Shortcuts
        tool_map = {
            'Textfeld': 't',
            'Text': 't',
            'Checkbox': 'c',
            'Kontrollkästchen': 'c',
            'Optionsfeld': 'r',
            'Radio Button': 'r',
            'Dropdown': 'd',
            'Listenfeld': 'd',
            'Signatur': 's',
            'Unterschrift': 's'
        }
        
        key = tool_map.get(field_type, 't')  # Standard: Textfeld
        
        try:
            # Versuche Keyboard Shortcut
            pyautogui.press(key)
            time.sleep(0.5)
            
            self.logger.log(f"Tool ausgewählt: {field_type} (Taste: {key})", "DEBUG")
            return True
            
        except Exception as e:
            self.logger.log(f"Fehler beim Auswählen des Tools: {str(e)}", "ERROR")
            return False
            
    def configure_field_properties(self, field_name, display_name, field_type):
        """Konfiguriert die Eigenschaften des erstellten Feldes"""
        
        try:
            time.sleep(0.5)
            
            # Öffne Properties-Dialog
            # Methode 1: Doppelklick (falls Feld noch selektiert)
            pyautogui.doubleClick()
            time.sleep(1)
            
            # Falls das nicht funktioniert, versuche Rechtsklick
            if not self._is_properties_dialog_open():
                pyautogui.rightClick()
                time.sleep(0.5)
                pyautogui.press('p')  # Properties
                time.sleep(1)
                
            # Feldname eingeben
            pyautogui.hotkey('ctrl', 'a')  # Alles markieren
            time.sleep(0.1)
            pyautogui.typewrite(field_name, interval=0.02)
            
            # Display Name (falls vorhanden und Feld verfügbar)
            if display_name and display_name.strip():
                pyautogui.press('tab')  # Nächstes Feld
                time.sleep(0.1)
                pyautogui.typewrite(display_name, interval=0.02)
                
            # Speichere Properties
            pyautogui.press('enter')  # OK Button
            time.sleep(0.5)
            
            self.logger.log(f"Feld-Eigenschaften konfiguriert: {field_name}", "SUCCESS")
            return True
            
        except Exception as e:
            self.logger.log(f"Fehler beim Konfigurieren der Eigenschaften: {str(e)}", "ERROR")
            # Versuche Dialog zu schließen
            pyautogui.press('esc')
            return False
            
    def _is_properties_dialog_open(self):
        """Prüft ob der Properties-Dialog geöffnet ist"""
        # Einfache Implementierung - könnte mit OCR verbessert werden
        return True  # Für jetzt annehmen, dass es funktioniert
        
    def find_and_rename_field(self, original_name, new_name):
        """Sucht und benennt ein spezifisches Feld um"""
        self.logger.log(f"Suche Feld: '{original_name}' → '{new_name}'", "INFO")
        
        # Tab-Navigation durch alle Felder
        for i in range(100):  # Maximal 100 Felder durchsuchen
            try:
                # Rechtsklick für Kontextmenü
                pyautogui.rightClick()
                time.sleep(0.3)
                
                # Properties öffnen
                pyautogui.press('p')
                time.sleep(0.8)
                
                # Prüfe Feldname
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.2)
                
                try:
                    current_name = pyperclip.paste().strip()
                except:
                    current_name = ""
                    
                if current_name == original_name:
                    # Gefunden! Umbenennen
                    self.logger.log(f"Feld gefunden: '{original_name}'", "SUCCESS")
                    
                    # Neuen Namen eingeben
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.typewrite(new_name, interval=0.02)
                    time.sleep(0.3)
                    
                    # Speichern
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    
                    return True
                else:
                    # Nicht gefunden, Dialog schließen und weiter
                    pyautogui.press('esc')
                    time.sleep(0.3)
                    
                # Nächstes Feld
                pyautogui.press('tab')
                time.sleep(0.2)
                
            except Exception as e:
                # Bei Fehler Dialog schließen und weiter
                pyautogui.press('esc')
                time.sleep(0.2)
                continue
                
        self.logger.log(f"Feld '{original_name}' nicht gefunden", "WARNING")
        return False
        
    def capture_position_with_keyboard(self):
        """Erfasst Position mit Keyboard-Modul"""
        self.logger.log("Warte auf Leertaste...", "INFO")
        
        try:
            # Verwende keyboard Modul für bessere Erkennung
            keyboard.wait('space')
            
            # Position erfassen
            x, y = pyautogui.position()
            self.logger.log(f"Position erfasst: ({x}, {y})", "SUCCESS")
            return (x, y)
            
        except Exception as e:
            self.logger.log(f"Fehler beim Erfassen der Position: {str(e)}", "ERROR")
            return None
