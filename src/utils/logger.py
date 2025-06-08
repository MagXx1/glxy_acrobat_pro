from datetime import datetime
import os

class Logger:
    def __init__(self):
        self.log_dir = "logs"
        self.ensure_log_dir()
    
    def ensure_log_dir(self):
        """Erstellt das Log-Verzeichnis falls es nicht existiert"""
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
    
    def log(self, message, level="INFO"):
        """Erstellt eine formatierte Log-Nachricht"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Level-Icons
        level_icons = {
            "INFO": "ℹ️",
            "SUCCESS": "✅", 
            "WARNING": "⚠️",
            "ERROR": "❌",
            "DEBUG": "🔍"
        }
        
        icon = level_icons.get(level, "ℹ️")
        formatted_message = f"[{timestamp}] {icon} {message}"
        
        # Auch in Datei schreiben
        self.write_to_file(formatted_message)
        
        return formatted_message
    
    def write_to_file(self, message):
        """Schreibt Logs in Datei"""
        try:
            log_file = os.path.join(self.log_dir, f"app_log_{datetime.now().strftime('%Y-%m-%d')}.txt")
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"{message}\n")
        except Exception:
            pass  # Ignoriere Log- Fehler
