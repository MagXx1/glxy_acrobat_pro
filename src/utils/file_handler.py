import json
import os
import pandas as pd
from .logger import Logger

class SettingsManager:
    def __init__(self):
        self.settings_dir = "settings"
        self.settings_file = os.path.join(self.settings_dir, "app_settings.json")
        self.logger = Logger()
        
    def save_settings(self, settings):
        """Speichert Einstellungen in JSON-Datei"""
        try:
            os.makedirs(self.settings_dir, exist_ok=True)
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
            self.logger.log("Einstellungen gespeichert", "SUCCESS")
            return True
        except Exception as e:
            self.logger.log(f"Fehler beim Speichern der Einstellungen: {e}", "ERROR")
            return False
            
    def load_settings(self):
        """Lädt Einstellungen aus JSON-Datei"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                self.logger.log("Einstellungen geladen", "SUCCESS")
                return settings
            else:
                self.logger.log("Keine gespeicherten Einstellungen gefunden", "INFO")
                return {}
        except Exception as e:
            self.logger.log(f"Fehler beim Laden der Einstellungen: {e}", "ERROR")
            return {}

class FileHandler:
    def __init__(self):
        self.logger = Logger()
        
    def load_field_definitions(self, file_path):
        """Lädt Feld-Definitionen mit robuster Encoding-Erkennung"""
        if not file_path:
            self.logger.log("Keine Daten-Datei ausgewählt", "ERROR")
            return None
            
        self.logger.log("Lade Feld-Definitionen...", "INFO")
        
        try:
            encodings = ['utf-8', 'utf-16', 'windows-1252', 'latin1', 'iso-8859-1']
            df = None
            used_encoding = None
            
            for encoding in encodings:
                try:
                    if file_path.endswith('.json'):
                        with open(file_path, 'r', encoding=encoding) as f:
                            data = json.load(f)
                        df = pd.DataFrame(data)
                    elif file_path.endswith('.csv'):
                        for sep in [';', ',', '\t']:
                            try:
                                df = pd.read_csv(file_path, sep=sep, encoding=encoding)
                                if len(df.columns) > 1:
                                    break
                            except:
                                continue
                    else:  # TXT-Datei
                        for sep in ['\t', ';', ',']:
                            try:
                                df = pd.read_csv(file_path, sep=sep, encoding=encoding)
                                if len(df.columns) > 1:
                                    break
                            except:
                                continue
                    
                    if df is not None and len(df) > 0:
                        used_encoding = encoding
                        break
                        
                except (UnicodeDecodeError, pd.errors.EmptyDataError):
                    continue
                    
            if df is None:
                raise Exception("Konnte Datei mit keinem Encoding/Separator lesen")
                
            self.logger.log(f"Daten erfolgreich geladen mit {used_encoding} Encoding", "SUCCESS")
            
            # Validiere Spalten
            required_cols = ['original_name', 'new_name']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise Exception(f"Fehlende erforderliche Spalten: {missing_cols}")
                
            # Filtere gültige Einträge
            initial_count = len(df)
            df = df[df['new_name'].notna() & (df['new_name'].str.strip() != '')]
            valid_count = len(df)
            
            if valid_count != initial_count:
                self.logger.log(f"Gefiltert: {initial_count - valid_count} leere Einträge entfernt", "WARNING")
                
            self.logger.log(f"Gefunden: {valid_count} gültige Feld-Definitionen", "SUCCESS")
            return df
            
        except Exception as e:
            self.logger.log(f"Fehler beim Laden der Feld-Definitionen: {str(e)}", "ERROR")
            return None
