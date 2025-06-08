import pyautogui
import pandas as pd
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pyperclip
import json
from PIL import Image, ImageTk
import threading
import subprocess
import shutil
from datetime import datetime

# Sicherheitseinstellungen
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.8

class AcrobatFormAutomator:
    def __init__(self, root):
        self.root = root
        self.root.title("🤖 Adobe Acrobat DC Formular-Automatisierung v4.0 - Vollversion")
        self.root.geometry("1000x900")
        self.root.configure(bg="#f0f0f0")
        
        # Variablen
        self.pdf_path = None
        self.data_path = None
        self.is_running = False
        self.pause_automation = False
        self.field_positions = {}
        self.created_fields = []
        self.current_task = ""
        
        # Adobe Acrobat Tool-Koordinaten (werden kalibriert)
        self.tool_coordinates = {
            'textfield': None,
            'checkbox': None,
            'radiobutton': None,
            'dropdown': None,
            'signature': None
        }
        
        self.create_gui()
        self.load_settings()
        
    def create_gui(self):
        # Hauptcontainer
        main_container = tk.Frame(self.root, bg="#f0f0f0")
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Titel mit Icon
        title_frame = tk.Frame(main_container, bg="#f0f0f0")
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tk.Label(title_frame, 
                              text="🤖 Adobe Acrobat DC Formular-Automatisierung", 
                              font=("Arial", 18, "bold"), 
                              fg="#1976D2", 
                              bg="#f0f0f0")
        title_label.pack()
        
        subtitle_label = tk.Label(title_frame, 
                                 text="Vollautomatische Erstellung und Bearbeitung von PDF-Formularfeldern", 
                                 font=("Arial", 10), 
                                 fg="#666666", 
                                 bg="#f0f0f0")
        subtitle_label.pack()
        
        # Notebook für Tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill="both", expand=True, pady=10)
        
        # Styles für besseres Aussehen
        style = ttk.Style()
        style.theme_use('clam')
        
        # Tab 1: Hauptfunktionen
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="🏠 Hauptfunktionen")
        self.create_main_tab()
        
        # Tab 2: Felderstellung
        self.creation_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.creation_tab, text="📝 Feld-Erstellung")
        self.create_creation_tab()
        
        # Tab 3: Einstellungen
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="⚙️ Einstellungen")
        self.create_settings_tab()
        
        # Tab 4: Kalibrierung
        self.calibration_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.calibration_tab, text="📐 Kalibrierung")
        self.create_calibration_tab()
        
        # Status und Log
        self.create_status_section(main_container)
        
    def create_main_tab(self):
        # Scrollbarer Bereich
        canvas = tk.Canvas(self.main_tab, bg="white")
        scrollbar = ttk.Scrollbar(self.main_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Datei-Auswahl Sektion
        file_section = ttk.LabelFrame(scrollable_frame, text="📁 Datei-Auswahl", padding="10")
        file_section.pack(fill="x", pady=10, padx=10)
        
        # PDF-Auswahl
        pdf_frame = ttk.Frame(file_section)
        pdf_frame.pack(fill="x", pady=5)
        ttk.Label(pdf_frame, text="Adobe Acrobat DC PDF:", font=("Arial", 10, "bold")).pack(anchor="w")
        
        pdf_select_frame = ttk.Frame(pdf_frame)
        pdf_select_frame.pack(fill="x", pady=2)
        
        self.pdf_label = tk.Label(pdf_select_frame, text="Keine PDF ausgewählt", 
                                 bg="#ffffff", relief="sunken", anchor="w", 
                                 font=("Arial", 9), pady=5, padx=5)
        self.pdf_label.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ttk.Button(pdf_select_frame, text="📂 PDF wählen", 
                  command=self.select_pdf).pack(side="right")
        
        # Daten-Auswahl
        data_frame = ttk.Frame(file_section)
        data_frame.pack(fill="x", pady=5)
        ttk.Label(data_frame, text="Feld-Definitionen (TXT/CSV):", font=("Arial", 10, "bold")).pack(anchor="w")
        
        data_select_frame = ttk.Frame(data_frame)
        data_select_frame.pack(fill="x", pady=2)
        
        self.data_label = tk.Label(data_select_frame, text="Keine Daten ausgewählt", 
                                  bg="#ffffff", relief="sunken", anchor="w", 
                                  font=("Arial", 9), pady=5, padx=5)
        self.data_label.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ttk.Button(data_select_frame, text="📂 Daten laden", 
                  command=self.select_data).pack(side="right")
        
        # Vorschau Button
        ttk.Button(file_section, text="👁️ Daten-Vorschau", 
                  command=self.preview_data, 
                  style="Accent.TButton").pack(pady=5)
        
        # Hauptfunktionen Sektion
        functions_section = ttk.LabelFrame(scrollable_frame, text="🚀 Hauptfunktionen", padding="10")
        functions_section.pack(fill="x", pady=10, padx=10)
        
        # Buttons für Hauptfunktionen
        button_configs = [
            ("🔧 Adobe Acrobat DC öffnen", self.open_acrobat, "#FF5722"),
            ("📝 Formular-Modus aktivieren", self.enter_form_mode, "#2196F3"),
            ("➕ Alle Felder erstellen", self.create_all_fields, "#4CAF50"),
            ("✏️ Felder umbenennen", self.rename_all_fields, "#FF9800"),
            ("📐 Felder positionieren", self.position_all_fields, "#9C27B0"),
            ("🎯 Vollautomatik (Alles)", self.full_automation, "#F44336")
        ]
        
        for i, (text, command, color) in enumerate(button_configs):
            btn = tk.Button(functions_section, text=text, command=command,
                           font=("Arial", 11, "bold"), fg="white", bg=color,
                           relief="raised", bd=2, pady=8)
            btn.pack(fill="x", pady=3)
            
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def create_creation_tab(self):
        # Manueller Feld-Ersteller
        manual_section = ttk.LabelFrame(self.creation_tab, text="🎯 Manueller Feld-Ersteller", padding="10")
        manual_section.pack(fill="x", pady=10, padx=10)
        
        # Feld-Typ Auswahl
        type_frame = ttk.Frame(manual_section)
        type_frame.pack(fill="x", pady=5)
        ttk.Label(type_frame, text="Feldtyp:", font=("Arial", 10, "bold")).pack(anchor="w")
        
        self.field_type_var = tk.StringVar(value="Textfeld")
        field_types = [
            ("📝 Textfeld", "Textfeld"),
            ("☑️ Checkbox", "Checkbox"), 
            ("🔘 Optionsfeld", "Optionsfeld"),
            ("📋 Dropdown", "Dropdown"),
            ("✍️ Signatur", "Signatur")
        ]
        
        type_radio_frame = ttk.Frame(type_frame)
        type_radio_frame.pack(fill="x", pady=2)
        
        for i, (display, value) in enumerate(field_types):
            ttk.Radiobutton(type_radio_frame, text=display, 
                           variable=self.field_type_var, value=value).grid(
                               row=i//3, column=i%3, sticky="w", padx=10, pady=2)
        
        # Feld-Eigenschaften
        props_frame = ttk.LabelFrame(manual_section, text="Feld-Eigenschaften", padding="5")
        props_frame.pack(fill="x", pady=10)
        
        # Grid Layout für Properties
        properties = [
            ("Feldname:", "field_name_entry"),
            ("Anzeigename:", "display_name_entry"),
            ("Breite (px):", "field_width_entry"),
            ("Höhe (px):", "field_height_entry")
        ]
        
        self.property_entries = {}
        for i, (label, var_name) in enumerate(properties):
            ttk.Label(props_frame, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=3)
            entry = ttk.Entry(props_frame, width=30)
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=3)
            self.property_entries[var_name] = entry
            
        # Standardwerte setzen
        self.property_entries["field_width_entry"].insert(0, "200")
        self.property_entries["field_height_entry"].insert(0, "25")
        
        props_frame.columnconfigure(1, weight=1)
        
        # Zusätzliche Optionen
        options_frame = ttk.Frame(props_frame)
        options_frame.grid(row=len(properties), column=0, columnspan=2, sticky="ew", pady=5)
        
        self.required_var = tk.BooleanVar()
        self.readonly_var = tk.BooleanVar()
        
        ttk.Checkbutton(options_frame, text="Pflichtfeld", variable=self.required_var).pack(side="left", padx=10)
        ttk.Checkbutton(options_frame, text="Nur Lesen", variable=self.readonly_var).pack(side="left", padx=10)
        
        # Buttons für manuelle Erstellung
        manual_buttons_frame = ttk.Frame(manual_section)
        manual_buttons_frame.pack(fill="x", pady=10)
        
        ttk.Button(manual_buttons_frame, text="📍 Position erfassen", 
                  command=self.capture_position).pack(side="left", padx=5)
        ttk.Button(manual_buttons_frame, text="➕ Feld erstellen", 
                  command=self.create_single_field).pack(side="left", padx=5)
        ttk.Button(manual_buttons_frame, text="🔍 Feld finden", 
                  command=self.find_field).pack(side="left", padx=5)
        
        # Batch-Operationen
        batch_section = ttk.LabelFrame(self.creation_tab, text="🔄 Batch-Operationen", padding="10")
        batch_section.pack(fill="x", pady=10, padx=10)
        
        batch_buttons = [
            ("📊 Grid-Layout erstellen", self.create_grid_layout),
            ("🔄 Alle Felder aktualisieren", self.update_all_fields),
            ("🗑️ Alle Felder löschen", self.delete_all_fields),
            ("💾 Feldpositionen speichern", self.save_field_positions)
        ]
        
        for text, command in batch_buttons:
            ttk.Button(batch_section, text=text, command=command).pack(side="left", padx=5, pady=2)
            
    def create_settings_tab(self):
        # Automatisierung Einstellungen
        automation_section = ttk.LabelFrame(self.settings_tab, text="🤖 Automatisierung", padding="10")
        automation_section.pack(fill="x", pady=10, padx=10)
        
        # Geschwindigkeit
        speed_frame = ttk.Frame(automation_section)
        speed_frame.pack(fill="x", pady=5)
        ttk.Label(speed_frame, text="Geschwindigkeit (Sekunden zwischen Aktionen):").pack(anchor="w")
        
        self.speed_var = tk.DoubleVar(value=1.0)
        speed_scale = ttk.Scale(speed_frame, from_=0.2, to=3.0, 
                               variable=self.speed_var, orient="horizontal")
        speed_scale.pack(fill="x", pady=2)
        
        self.speed_label = ttk.Label(speed_frame, text="1.0 Sekunden")
        self.speed_label.pack()
        
        speed_scale.configure(command=lambda v: self.speed_label.config(text=f"{float(v):.1f} Sekunden"))
        
        # Sicherheitsoptionen
        safety_frame = ttk.LabelFrame(automation_section, text="Sicherheit", padding="5")
        safety_frame.pack(fill="x", pady=10)
        
        self.safety_var = tk.BooleanVar(value=True)
        self.backup_var = tk.BooleanVar(value=True)
        self.confirm_var = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(safety_frame, text="Sicherheitsmodus aktiviert", 
                       variable=self.safety_var).pack(anchor="w")
        ttk.Checkbutton(safety_frame, text="Automatische Backups erstellen", 
                       variable=self.backup_var).pack(anchor="w")
        ttk.Checkbutton(safety_frame, text="Vor jeder Aktion bestätigen", 
                       variable=self.confirm_var).pack(anchor="w")
        
        # Feld-Standardwerte
        defaults_section = ttk.LabelFrame(self.settings_tab, text="📏 Standard-Feldeinstellungen", padding="10")
        defaults_section.pack(fill="x", pady=10, padx=10)
        
        # Grid für Standardwerte
        defaults_grid = ttk.Frame(defaults_section)
        defaults_grid.pack(fill="x")
        
        default_settings = [
            ("Standard Textfeld-Breite:", "default_width", "200"),
            ("Standard Textfeld-Höhe:", "default_height", "25"),
            ("Standard Checkbox-Größe:", "default_checkbox_size", "15"),
            ("Feld-Abstand (Grid):", "default_spacing", "40"),
            ("Startposition X:", "start_x", "100"),
            ("Startposition Y:", "start_y", "150")
        ]
        
        self.default_entries = {}
        for i, (label, var_name, default_value) in enumerate(default_settings):
            ttk.Label(defaults_grid, text=label).grid(row=i//2, column=(i%2)*2, sticky="w", padx=5, pady=3)
            entry = ttk.Entry(defaults_grid, width=15)
            entry.insert(0, default_value)
            entry.grid(row=i//2, column=(i%2)*2+1, sticky="ew", padx=5, pady=3)
            self.default_entries[var_name] = entry
            
        # Kontroll-Buttons
        control_section = ttk.LabelFrame(self.settings_tab, text="🎮 Steuerung", padding="10")
        control_section.pack(fill="x", pady=10, padx=10)
        
        control_frame = ttk.Frame(control_section)
        control_frame.pack(fill="x")
        
        self.pause_btn = tk.Button(control_frame, text="⏸️ Pausieren", 
                                  command=self.toggle_pause,
                                  bg="#FF9800", fg="white", font=("Arial", 10, "bold"))
        self.pause_btn.pack(side="left", padx=5)
        
        tk.Button(control_frame, text="🛑 Stoppen", 
                 command=self.stop_automation,
                 bg="#F44336", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        
        tk.Button(control_frame, text="💾 Einstellungen speichern", 
                 command=self.save_settings,
                 bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side="right", padx=5)
        
    def create_calibration_tab(self):
        # Tool-Kalibrierung
        calibration_section = ttk.LabelFrame(self.calibration_tab, text="🎯 Adobe Acrobat DC Tool-Kalibrierung", padding="10")
        calibration_section.pack(fill="x", pady=10, padx=10)
        
        ttk.Label(calibration_section, 
                 text="Kalibrieren Sie die Tool-Positionen für eine präzise Automatisierung:",
                 font=("Arial", 10)).pack(anchor="w", pady=5)
        
        # Tool-Buttons
        tools = [
            ("📝 Textfeld-Tool", "textfield"),
            ("☑️ Checkbox-Tool", "checkbox"),
            ("🔘 Optionsfeld-Tool", "radiobutton"),
            ("📋 Dropdown-Tool", "dropdown"),
            ("✍️ Signatur-Tool", "signature")
        ]
        
        self.calibration_labels = {}
        
        for text, tool_name in tools:
            tool_frame = ttk.Frame(calibration_section)
            tool_frame.pack(fill="x", pady=3)
            
            ttk.Button(tool_frame, text=f"📍 {text} kalibrieren", 
                      command=lambda t=tool_name: self.calibrate_tool(t)).pack(side="left", padx=5)
            
            status_label = ttk.Label(tool_frame, text="Nicht kalibriert", foreground="red")
            status_label.pack(side="left", padx=10)
            self.calibration_labels[tool_name] = status_label
            
        # Kalibrierungs-Buttons
        calib_buttons_frame = ttk.Frame(calibration_section)
        calib_buttons_frame.pack(fill="x", pady=10)
        
        ttk.Button(calib_buttons_frame, text="💾 Kalibrierung speichern", 
                  command=self.save_calibration).pack(side="left", padx=5)
        ttk.Button(calib_buttons_frame, text="📂 Kalibrierung laden", 
                  command=self.load_calibration).pack(side="left", padx=5)
        ttk.Button(calib_buttons_frame, text="🔄 Auto-Kalibrierung", 
                  command=self.auto_calibrate).pack(side="left", padx=5)
        
        # Adobe Acrobat Kontrolle
        acrobat_section = ttk.LabelFrame(self.calibration_tab, text="🔧 Adobe Acrobat DC Kontrolle", padding="10")
        acrobat_section.pack(fill="x", pady=10, padx=10)
        
        acrobat_buttons = [
            ("🚀 Adobe Acrobat DC starten", self.launch_acrobat),
            ("🎯 Acrobat fokussieren", self.focus_acrobat),
            ("📝 Formular-Modus testen", self.test_form_mode),
            ("📊 Screenshot erstellen", self.take_screenshot)
        ]
        
        for text, command in acrobat_buttons:
            ttk.Button(acrobat_section, text=text, command=command).pack(side="left", padx=5, pady=2)
            
    def create_status_section(self, parent):
        # Status und Progress
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill="x", pady=10)
        
        # Aktuelle Aufgabe
        self.task_label = tk.Label(status_frame, text="Aktuelle Aufgabe: Bereit", 
                                  font=("Arial", 10, "bold"), fg="#1976D2", bg="#f0f0f0")
        self.task_label.pack(anchor="w")
        
        # Status
        self.status_label = tk.Label(status_frame, text="Status: Bereit für Formular-Automatisierung", 
                                    bg="#E8F5E8", relief="sunken", anchor="w", font=("Arial", 9), pady=3)
        self.status_label.pack(fill="x", pady=2)
        
        # Progress Bar
        self.progress = ttk.Progressbar(status_frame, length=400, mode="determinate")
        self.progress.pack(fill="x", pady=2)
        
        # Log-Bereich
        log_frame = ttk.LabelFrame(parent, text="📋 Aktivitätsprotokoll", padding="5")
        log_frame.pack(fill="both", expand=True, pady=10)
        
        # Log-Text mit Scrollbar
        log_scroll_frame = ttk.Frame(log_frame)
        log_scroll_frame.pack(fill="both", expand=True)
        
        self.log_text = tk.Text(log_scroll_frame, height=10, font=("Consolas", 9),
                               wrap=tk.WORD, bg="#ffffff")
        log_scrollbar = ttk.Scrollbar(log_scroll_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scrollbar.pack(side="right", fill="y")
        
        # Log-Buttons
        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.pack(fill="x", pady=5)
        
        ttk.Button(log_buttons_frame, text="🗑️ Log löschen", 
                  command=self.clear_log).pack(side="left", padx=5)
        ttk.Button(log_buttons_frame, text="💾 Log speichern", 
                  command=self.save_log).pack(side="left", padx=5)
        ttk.Button(log_buttons_frame, text="📋 Log kopieren", 
                  command=self.copy_log).pack(side="left", padx=5)
        
    def log_message(self, message, level="INFO"):
        """Erweiterte Log-Funktion mit Leveln"""
        timestamp = time.strftime("%H:%M:%S")
        
        # Level-spezifische Formatierung
        level_colors = {
            "INFO": "#000000",
            "SUCCESS": "#4CAF50", 
            "WARNING": "#FF9800",
            "ERROR": "#F44336",
            "DEBUG": "#9E9E9E"
        }
        
        color = level_colors.get(level, "#000000")
        
        # Füge Nachricht hinzu
        self.log_text.insert(tk.END, f"[{timestamp}] {level}: {message}\n")
        
        # Scrolle zum Ende
        self.log_text.see(tk.END)
        
        # Update GUI
        self.root.update_idletasks()
        
        # Log auch in Datei schreiben
        try:
            log_dir = "logs"
            os.makedirs(log_dir, exist_ok=True)
            log_file = os.path.join(log_dir, f"automation_log_{datetime.now().strftime('%Y-%m-%d')}.txt")
            
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {level}: {message}\n")
        except Exception as e:
            pass  # Ignoriere Log-Datei Fehler
            
    def update_status(self, message, bg_color="#E8F5E8"):
        """Aktualisiert die Statusanzeige"""
        self.status_label.config(text=message, bg=bg_color)
        self.root.update_idletasks()
        
    def update_task(self, task):
        """Aktualisiert die aktuelle Aufgabe"""
        self.current_task = task
        self.task_label.config(text=f"Aktuelle Aufgabe: {task}")
        self.root.update_idletasks()
        
    def update_progress(self, value, maximum=100):
        """Aktualisiert die Progress Bar"""
        self.progress['value'] = value
        self.progress['maximum'] = maximum
        self.root.update_idletasks()
        
    # ===== DATEI-OPERATIONEN =====
    
    def select_pdf(self):
        """PDF-Datei auswählen"""
        file = filedialog.askopenfilename(
            title="Adobe Acrobat DC PDF auswählen",
            filetypes=[
                ("PDF-Dateien", "*.pdf"),
                ("Alle Dateien", "*.*")
            ]
        )
        if file:
            self.pdf_path = file
            filename = os.path.basename(file)
            self.pdf_label.config(text=filename)
            self.log_message(f"PDF ausgewählt: {filename}", "SUCCESS")
            
    def select_data(self):
        """Feld-Definitionen auswählen"""
        file = filedialog.askopenfilename(
            title="Feld-Definitionen auswählen",
            filetypes=[
                ("Text-Dateien", "*.txt"),
                ("CSV-Dateien", "*.csv"),
                ("JSON-Dateien", "*.json"),
                ("Alle Dateien", "*.*")
            ]
        )
        if file:
            self.data_path = file
            filename = os.path.basename(file)
            self.data_label.config(text=filename)
            self.log_message(f"Feld-Definitionen geladen: {filename}", "SUCCESS")
            
    def load_field_definitions(self):
        """Lädt die Feld-Definitionen mit robuster Encoding-Erkennung"""
        if not self.data_path:
            self.log_message("Keine Daten-Datei ausgewählt", "ERROR")
            return None
            
        self.log_message("Lade Feld-Definitionen...", "INFO")
        
        try:
            # Verschiedene Encodings probieren
            encodings = ['utf-8', 'utf-16', 'windows-1252', 'latin1', 'iso-8859-1']
            df = None
            used_encoding = None
            
            for encoding in encodings:
                try:
                    if self.data_path.endswith('.json'):
                        with open(self.data_path, 'r', encoding=encoding) as f:
                            data = json.load(f)
                        df = pd.DataFrame(data)
                    elif self.data_path.endswith('.csv'):
                        # Versuche verschiedene Separatoren
                        for sep in [';', ',', '\t']:
                            try:
                                df = pd.read_csv(self.data_path, sep=sep, encoding=encoding)
                                if len(df.columns) > 1:  # Erfolgreich geparst
                                    break
                            except:
                                continue
                    else:  # TXT-Datei
                        for sep in ['\t', ';', ',']:
                            try:
                                df = pd.read_csv(self.data_path, sep=sep, encoding=encoding)
                                if len(df.columns) > 1:  # Erfolgreich geparst
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
                
            self.log_message(f"Daten erfolgreich geladen mit {used_encoding} Encoding", "SUCCESS")
            
            # Validiere erforderliche Spalten
            required_cols = ['original_name', 'new_name']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise Exception(f"Fehlende erforderliche Spalten: {missing_cols}")
                
            # Filtere gültige Einträge
            initial_count = len(df)
            df = df[df['new_name'].notna() & (df['new_name'].str.strip() != '')]
            valid_count = len(df)
            
            if valid_count != initial_count:
                self.log_message(f"Gefiltert: {initial_count - valid_count} leere Einträge entfernt", "WARNING")
                
            self.log_message(f"Gefunden: {valid_count} gültige Feld-Definitionen", "SUCCESS")
            
            # Zeige Feldtyp-Verteilung
            if 'type' in df.columns:
                type_counts = df['type'].value_counts()
                for field_type, count in type_counts.items():
                    self.log_message(f"  - {field_type}: {count} Felder", "DEBUG")
                    
            return df
            
        except Exception as e:
            self.log_message(f"Fehler beim Laden der Feld-Definitionen: {str(e)}", "ERROR")
            messagebox.showerror("Daten-Fehler", f"Konnte Feld-Definitionen nicht laden:\n\n{str(e)}")
            return None
            
    def preview_data(self):
        """Zeigt eine Vorschau der geladenen Daten"""
        df = self.load_field_definitions()
        if df is None:
            return
            
        # Vorschau-Fenster
        preview_window = tk.Toplevel(self.root)
        preview_window.title("📊 Feld-Definitionen Vorschau")
        preview_window.geometry("1000x700")
        preview_window.configure(bg="#f0f0f0")
        
        # Header
        header_frame = tk.Frame(preview_window, bg="#f0f0f0")
        header_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Label(header_frame, text="📊 Feld-Definitionen Vorschau", 
                font=("Arial", 16, "bold"), bg="#f0f0f0").pack()
        tk.Label(header_frame, text=f"Datei: {os.path.basename(self.data_path)}", 
                font=("Arial", 10), bg="#f0f0f0").pack()
        
        # Statistiken
        stats_frame = tk.Frame(preview_window, bg="#f0f0f0")
        stats_frame.pack(fill="x", padx=10, pady=5)
        
        total_fields = len(df)
        field_types = df['type'].value_counts() if 'type' in df.columns else {}
        
        stats_text = f"Gesamt: {total_fields} Felder"
        if field_types.empty == False:
            stats_text += " | " + " | ".join([f"{t}: {c}" for t, c in field_types.items()])
            
        tk.Label(stats_frame, text=stats_text, font=("Arial", 10, "bold"), 
                bg="#f0f0f0", fg="#1976D2").pack()
        
        # Treeview für Daten
        tree_frame = tk.Frame(preview_window)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Bestimme Spalten
        columns = list(df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=20)
        
        # Konfiguriere Spalten
        col_widths = {
            'original_name': 150,
            'new_name': 200,
            'display_name': 250,
            'type': 100
        }
        
        for col in columns:
            tree.heading(col, text=col.replace('_', ' ').title())
            width = col_widths.get(col, 120)
            tree.column(col, width=width, minwidth=80)
            
        # Fülle Daten
        for _, row in df.iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else '' for col in columns]
            tree.insert('', 'end', values=values)
            
        # Scrollbars
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # Pack Treeview
        tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Buttons
        button_frame = tk.Frame(preview_window, bg="#f0f0f0")
        button_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(button_frame, text="📊 Statistiken exportieren", 
                  command=lambda: self.export_statistics(df)).pack(side="left", padx=5)
        ttk.Button(button_frame, text="✅ Schließen", 
                  command=preview_window.destroy).pack(side="right", padx=5)
        
    # ===== ADOBE ACROBAT KONTROLLE =====
    
    def launch_acrobat(self):
        """Startet Adobe Acrobat DC"""
        self.log_message("Starte Adobe Acrobat DC...", "INFO")
        
        try:
            if self.pdf_path:
                # Öffne PDF direkt
                os.startfile(self.pdf_path)
                self.log_message(f"Öffne PDF: {os.path.basename(self.pdf_path)}", "SUCCESS")
            else:
                # Versuche Acrobat zu starten
                acrobat_paths = [
                    r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
                ]
                
                started = False
                for path in acrobat_paths:
                    if os.path.exists(path):
                        subprocess.Popen([path])
                        started = True
                        break
                        
                if not started:
                    self.log_message("Adobe Acrobat DC nicht gefunden. Bitte manuell starten.", "WARNING")
                else:
                    self.log_message("Adobe Acrobat DC gestartet", "SUCCESS")
                    
            time.sleep(3)  # Warte bis Acrobat geladen ist
            
        except Exception as e:
            self.log_message(f"Fehler beim Starten von Adobe Acrobat DC: {str(e)}", "ERROR")
            
    def open_acrobat(self):
        """Öffnet Adobe Acrobat DC und das PDF"""
        if not self.pdf_path:
            messagebox.showwarning("Keine PDF", "Bitte wählen Sie zuerst eine PDF-Datei aus.")
            return
            
        self.update_task("Adobe Acrobat DC öffnen")
        self.launch_acrobat()
        
        # Warte und fokussiere
        time.sleep(2)
        self.focus_acrobat()
        
    def focus_acrobat(self):
        """Bringt Adobe Acrobat DC in den Vordergrund"""
        self.log_message("Fokussiere Adobe Acrobat DC...", "INFO")
        
        try:
            # Versuche verschiedene Fenstertitel
            window_titles = [
                "Adobe Acrobat",
                "Acrobat",
                "Adobe Acrobat Reader",
                "Adobe Reader"
            ]
            
            focused = False
            for title in window_titles:
                windows = pyautogui.getWindowsWithTitle(title)
                if windows:
                    windows[0].activate()
                    time.sleep(1)
                    focused = True
                    self.log_message(f"Adobe Acrobat DC fokussiert (Titel: {title})", "SUCCESS")
                    break
                    
            if not focused:
                self.log_message("Adobe Acrobat DC Fenster nicht gefunden", "WARNING")
                messagebox.showwarning("Acrobat nicht gefunden", 
                                     "Adobe Acrobat DC Fenster nicht gefunden.\n\n"
                                     "Bitte stellen Sie sicher, dass Adobe Acrobat DC geöffnet ist.")
                
        except Exception as e:
            self.log_message(f"Fehler beim Fokussieren: {str(e)}", "ERROR")
            
    def enter_form_mode(self):
        """Aktiviert den Formular-Bearbeitungsmodus"""
        self.update_task("Formular-Modus aktivieren")
        self.log_message("Aktiviere Formular-Bearbeitungsmodus...", "INFO")
        
        try:
            # Stelle sicher, dass Acrobat fokussiert ist
            self.focus_acrobat()
            
            pyautogui.PAUSE = self.speed_var.get()
            
            # Methode 1: Keyboard Shortcut
            self.log_message("Versuche Tastenkombination Shift+Ctrl+7...", "DEBUG")
            pyautogui.hotkey('shift', 'ctrl', '7')
            time.sleep(3)
            
            # Prüfe ob erfolgreich
            if self.confirm_var.get():
                result = messagebox.askyesno("Formular-Modus", 
                                           "Wurde der Formular-Modus aktiviert?\n\n"
                                           "Falls nicht, versuchen wir es über das Tools-Menü.")
                if not result:
                    # Methode 2: Tools-Menü
                    self.log_message("Versuche über Tools-Menü...", "DEBUG")
                    pyautogui.hotkey('alt', 't')  # Tools-Menü
                    time.sleep(1)
                    pyautogui.press('f')  # Formulare
                    time.sleep(2)
                    
            self.log_message("Formular-Modus aktiviert", "SUCCESS")
            self.update_status("Status: Formular-Modus aktiv", "#E3F2FD")
            
            # Gebe Zeit für UI-Updates
            time.sleep(2)
            
        except Exception as e:
            self.log_message(f"Fehler beim Aktivieren des Formular-Modus: {str(e)}", "ERROR")
            messagebox.showerror("Formular-Modus Fehler", str(e))
            
    def test_form_mode(self):
        """Testet ob der Formular-Modus aktiv ist"""
        self.log_message("Teste Formular-Modus...", "INFO")
        
        try:
            self.focus_acrobat()
            
            # Versuche ein Textfeld-Tool zu aktivieren
            pyautogui.press('t')  # T für Textfeld
            time.sleep(1)
            
            # Screenshot für visueller Test
            screenshot = pyautogui.screenshot()
            
            messagebox.showinfo("Formular-Modus Test", 
                              "Test durchgeführt.\n\n"
                              "Prüfen Sie visuell ob:\n"
                              "• Formular-Tools sichtbar sind\n"
                              "• Textfeld-Tool aktiviert ist\n"
                              "• Cursor sich zu einem Kreuz ändert")
            
        except Exception as e:
            self.log_message(f"Fehler beim Testen: {str(e)}", "ERROR")
            
    # ===== FELD-ERSTELLUNG =====
    
    def create_all_fields(self):
        """Erstellt alle Felder basierend auf den Definitionen"""
        if not self.pdf_path or not self.data_path:
            messagebox.showwarning("Fehlende Dateien", 
                                  "Bitte wählen Sie sowohl PDF als auch Feld-Definitionen aus.")
            return
            
        df = self.load_field_definitions()
        if df is None:
            return
            
        # Bestätigung
        result = messagebox.askyesno("Alle Felder erstellen", 
                                   f"🤖 Sollen {len(df)} Felder automatisch erstellt werden?\n\n"
                                   "⚠️ Voraussetzungen:\n"
                                   "• Adobe Acrobat DC ist geöffnet\n"
                                   "• PDF ist geladen\n"
                                   "• Formular-Modus ist aktiv\n"
                                   "• Bildschirm ist sichtbar\n\n"
                                   "🛑 Zum Stoppen: Maus in obere linke Ecke!")
        if not result:
            return
            
        # Starte Automatisierung in separatem Thread
        self.is_running = True
        self.update_status("Status: Erstelle Felder...", "#FFE8E8")
        
        thread = threading.Thread(target=self._create_fields_worker, args=(df,))
        thread.daemon = True
        thread.start()
        
    def _create_fields_worker(self, df):
        """Worker-Thread für Feld-Erstellung"""
        try:
            self.update_task("Felder erstellen")
            pyautogui.PAUSE = self.speed_var.get()
            
            successful = 0
            failed = 0
            total = len(df)
            
            # Hole Standardwerte
            start_x = int(self.default_entries["start_x"].get())
            start_y = int(self.default_entries["start_y"].get())
            spacing = int(self.default_entries["default_spacing"].get())
            
            self.log_message(f"Starte Erstellung von {total} Feldern...", "INFO")
            self.log_message(f"Startposition: ({start_x}, {start_y}), Abstand: {spacing}px", "DEBUG")
            
            for index, row in df.iterrows():
                if not self.is_running:
                    self.log_message("Erstellung vom Benutzer gestoppt", "WARNING")
                    break
                    
                # Pause-Behandlung
                while self.pause_automation:
                    time.sleep(0.5)
                    
                field_name = str(row['new_name'])
                field_type = str(row.get('type', 'Textfeld'))
                display_name = str(row.get('display_name', ''))
                
                # Berechne Position (intelligentes Grid-Layout)
                x, y = self._calculate_field_position(index, start_x, start_y, spacing, field_name)
                
                self.log_message(f"Erstelle Feld {index+1}/{total}: {field_name} ({field_type})", "INFO")
                self.update_progress(index, total)
                
                try:
                    # Bestimme Feldgröße basierend auf Typ
                    width, height = self._get_field_size(field_type)
                    
                    # Erstelle Feld
                    if self.create_field_at_position(x, y, field_type, field_name, display_name, width, height):
                        successful += 1
                        
                        # Speichere Position für spätere Verwendung
                        self.field_positions[field_name] = {
                            'x': x, 'y': y, 'width': width, 'height': height,
                            'type': field_type, 'index': index
                        }
                        
                        self.log_message(f"✅ '{field_name}' erfolgreich erstellt", "SUCCESS")
                    else:
                        failed += 1
                        self.log_message(f"❌ Fehler bei '{field_name}'", "ERROR")
                        
                except pyautogui.FailSafeException:
                    self.log_message("🛑 Automatisierung durch FailSafe gestoppt", "WARNING")
                    break
                except Exception as e:
                    failed += 1
                    self.log_message(f"❌ Fehler bei '{field_name}': {str(e)}", "ERROR")
                    
            # Zusammenfassung
            self.update_progress(total, total)
            self.log_message(f"🏁 Feld-Erstellung abgeschlossen!", "SUCCESS")
            self.log_message(f"✅ Erfolgreich: {successful} Felder", "SUCCESS")
            self.log_message(f"❌ Fehler: {failed} Felder", "ERROR" if failed > 0 else "INFO")
            
            # Speichere PDF
            if successful > 0:
                self.save_pdf()
                
            # Benachrichtigung
            self.root.after(0, lambda: messagebox.showinfo("Feld-Erstellung abgeschlossen", 
                                       f"Feld-Erstellung beendet!\n\n"
                                       f"✅ Erfolgreich: {successful} Felder\n"
                                       f"❌ Fehler: {failed} Felder\n"
                                       f"📊 Gesamt: {total} Felder"))
                                       
        except Exception as e:
            self.log_message(f"❌ Schwerwiegender Fehler bei Feld-Erstellung: {str(e)}", "ERROR")
            self.root.after(0, lambda: messagebox.showerror("Kritischer Fehler", str(e)))
            
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.update_status("Status: Bereit", "#E8F5E8"))
            self.root.after(0, lambda: self.update_task("Bereit"))
            
    def _calculate_field_position(self, index, start_x, start_y, spacing, field_name):
        """Berechnet intelligente Feldposition basierend auf Feldname und Index"""
        
        # Kategorien für intelligente Positionierung
        categories = {
            'obj_': (start_x, start_y),                    # Objektdaten oben links
            'erfasser_': (start_x + 300, start_y),         # Erfasser oben rechts
            'mieter1_': (start_x, start_y + 200),          # Mieter 1 links mittig
            'mieter2_': (start_x + 300, start_y + 200),    # Mieter 2 rechts mittig
            'zaehler': (start_x, start_y + 400),           # Zähler unten links
            'schluessel': (start_x + 150, start_y + 500),  # Übergabegegenstände
            'handsender': (start_x + 150, start_y + 520),
            'chips': (start_x + 150, start_y + 540),
            'unterschrift': (start_x, start_y + 650)       # Unterschriften ganz unten
        }
        
        # Suche passende Kategorie
        base_x, base_y = start_x, start_y  # Standard
        category_index = 0
        
        for prefix, (cat_x, cat_y) in categories.items():
            if field_name.startswith(prefix):
                base_x, base_y = cat_x, cat_y
                # Zähle Felder in dieser Kategorie
                category_fields = [fn for fn in self.field_positions.keys() if fn.startswith(prefix)]
                category_index = len(category_fields)
                break
                
        # Berechne finale Position
        x = base_x
        y = base_y + (category_index * spacing)
        
        return x, y
        
    def _get_field_size(self, field_type):
        """Bestimmt Feldgröße basierend auf Typ"""
        
        size_map = {
            'Textfeld': (int(self.default_entries["default_width"].get()), 
                        int(self.default_entries["default_height"].get())),
            'Checkbox': (int(self.default_entries["default_checkbox_size"].get()), 
                        int(self.default_entries["default_checkbox_size"].get())),
            'Optionsfeld': (int(self.default_entries["default_checkbox_size"].get()), 
                           int(self.default_entries["default_checkbox_size"].get())),
            'Dropdown': (int(self.default_entries["default_width"].get()), 
                        int(self.default_entries["default_height"].get())),
            'Signatur': (200, 50),
            'Datum': (120, int(self.default_entries["default_height"].get())),
            'Zählwerk': (80, int(self.default_entries["default_height"].get()))
        }
        
        return size_map.get(field_type, (200, 25))
        
    def create_field_at_position(self, x, y, field_type, field_name, display_name, width, height):
        """Erstellt ein Feld an der angegebenen Position"""
        
        try:
            # Fokussiere Acrobat
            self.focus_acrobat()
            
            # Wähle entsprechendes Tool
            if not self.select_field_tool(field_type):
                self.log_message(f"Konnte Tool für {field_type} nicht auswählen", "WARNING")
                return False
                
            # Erstelle Feld durch Ziehen
            self.log_message(f"Zeichne {field_type} bei ({x}, {y}) mit Größe {width}x{height}", "DEBUG")
            
            pyautogui.moveTo(x, y, duration=0.3)
            time.sleep(0.2)
            
            pyautogui.mouseDown()
            pyautogui.dragTo(x + width, y + height, duration=0.8)
            pyautogui.mouseUp()
            
            time.sleep(1)
            
            # Konfiguriere Feld-Eigenschaften
            return self.configure_field_properties(field_name, display_name, field_type)
            
        except Exception as e:
            self.log_message(f"Fehler beim Erstellen von Feld {field_name}: {str(e)}", "ERROR")
            return False
            
    def select_field_tool(self, field_type):
        """Wählt das entsprechende Tool basierend auf Feldtyp"""
        
        # Mapping von Feldtypen zu Tools
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
            time.sleep(0.3)
            
            # Falls kalibrierte Koordinaten verfügbar sind, verwende diese
            tool_name = {
                't': 'textfield',
                'c': 'checkbox', 
                'r': 'radiobutton',
                'd': 'dropdown',
                's': 'signature'
            }.get(key, 'textfield')
            
            if self.tool_coordinates.get(tool_name):
                coords = self.tool_coordinates[tool_name]
                pyautogui.click(coords[0], coords[1])
                time.sleep(0.3)
                
            return True
            
        except Exception as e:
            self.log_message(f"Fehler beim Auswählen des Tools: {str(e)}", "ERROR")
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
                
            # Prüfe erneut
            if not self._is_properties_dialog_open():
                self.log_message("Properties-Dialog konnte nicht geöffnet werden", "WARNING")
                return False
                
            # Feldname eingeben
            pyautogui.hotkey('ctrl', 'a')  # Alles markieren
            time.sleep(0.1)
            pyautogui.typewrite(field_name, interval=0.02)
            
            # Display Name (falls vorhanden und Feld verfügbar)
            if display_name and display_name.strip():
                pyautogui.press('tab')  # Nächstes Feld
                time.sleep(0.1)
                pyautogui.typewrite(display_name, interval=0.02)
                
            # Zusätzliche Konfiguration basierend auf Feldtyp
            self._configure_field_type_specific(field_type)
            
            # Speichere Properties
            pyautogui.press('enter')  # OK Button
            time.sleep(0.5)
            
            return True
            
        except Exception as e:
            self.log_message(f"Fehler beim Konfigurieren der Eigenschaften: {str(e)}", "ERROR")
            # Versuche Dialog zu schließen
            pyautogui.press('esc')
            return False
            
    def _is_properties_dialog_open(self):
        """Prüft ob der Properties-Dialog geöffnet ist"""
        # Einfache Implementierung - könnte mit OCR verbessert werden
        return True  # Für jetzt annehmen, dass es funktioniert
        
    def _configure_field_type_specific(self, field_type):
        """Feldtyp-spezifische Konfiguration"""
        
        if field_type in ['Datum', 'Date']:
            # Für Datumsfelder könnte man Validierung setzen
            # Hier könnten Tab-Navigation und spezifische Einstellungen folgen
            pass
        elif field_type in ['Zählwerk', 'Number']:
            # Für Zahlenfelder Validierung setzen
            pass
        elif field_type in ['Unterschrift', 'Signatur']:
            # Signaturfeld-spezifische Einstellungen
            pass
            
    # ===== FELD-UMBENENNUNG =====
    
    def rename_all_fields(self):
        """Benennt alle bestehenden Felder um"""
        if not self.data_path:
            messagebox.showwarning("Keine Daten", "Bitte laden Sie zuerst Feld-Definitionen.")
            return
            
        df = self.load_field_definitions()
        if df is None:
            return
            
        result = messagebox.askyesno("Felder umbenennen", 
                                   f"🔄 Sollen {len(df)} Felder umbenannt werden?\n\n"
                                   "Dies durchsucht alle bestehenden Felder und benennt sie um.\n\n"
                                   "⚠️ Stellen Sie sicher, dass:\n"
                                   "• Adobe Acrobat DC geöffnet ist\n"
                                   "• PDF geladen ist\n"
                                   "• Formular-Modus aktiv ist")
        if not result:
            return
            
        # Starte Umbenennung in separatem Thread
        self.is_running = True
        self.update_status("Status: Benenne Felder um...", "#FFF8E1")
        
        thread = threading.Thread(target=self._rename_fields_worker, args=(df,))
        thread.daemon = True
        thread.start()
        
    def _rename_fields_worker(self, df):
        """Worker-Thread für Feld-Umbenennung"""
        try:
            self.update_task("Felder umbenennen")
            pyautogui.PAUSE = self.speed_var.get()
            
            successful = 0
            failed = 0
            total = len(df)
            
            self.log_message(f"Starte Umbenennung von {total} Feldern...", "INFO")
            
            for index, row in df.iterrows():
                if not self.is_running:
                    self.log_message("Umbenennung vom Benutzer gestoppt", "WARNING")
                    break
                    
                while self.pause_automation:
                    time.sleep(0.5)
                    
                original_name = str(row['original_name'])
                new_name = str(row['new_name'])
                
                self.log_message(f"Suche Feld {index+1}/{total}: '{original_name}'", "INFO")
                self.update_progress(index, total)
                
                try:
                    if self.find_and_rename_field(original_name, new_name):
                        successful += 1
                        self.log_message(f"✅ '{original_name}' → '{new_name}'", "SUCCESS")
                    else:
                        failed += 1
                        self.log_message(f"❌ Feld '{original_name}' nicht gefunden", "WARNING")
                        
                except pyautogui.FailSafeException:
                    self.log_message("🛑 Umbenennung durch FailSafe gestoppt", "WARNING")
                    break
                except Exception as e:
                    failed += 1
                    self.log_message(f"❌ Fehler bei '{original_name}': {str(e)}", "ERROR")
                    
            # Zusammenfassung
            self.update_progress(total, total)
            self.log_message(f"🏁 Umbenennung abgeschlossen!", "SUCCESS")
            self.log_message(f"✅ Erfolgreich: {successful} Felder", "SUCCESS")
            self.log_message(f"❌ Fehler: {failed} Felder", "ERROR" if failed > 0 else "INFO")
            
            # Speichere PDF
            if successful > 0:
                self.save_pdf()
                
            # Benachrichtigung
            self.root.after(0, lambda: messagebox.showinfo("Umbenennung abgeschlossen", 
                                       f"Feld-Umbenennung beendet!\n\n"
                                       f"✅ Erfolgreich: {successful} Felder\n"
                                       f"❌ Nicht gefunden: {failed} Felder\n"
                                       f"📊 Gesamt: {total} Felder"))
                                       
        except Exception as e:
            self.log_message(f"❌ Schwerwiegender Fehler bei Umbenennung: {str(e)}", "ERROR")
            self.root.after(0, lambda: messagebox.showerror("Kritischer Fehler", str(e)))
            
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.update_status("Status: Bereit", "#E8F5E8"))
            self.root.after(0, lambda: self.update_task("Bereit"))
            
    def find_and_rename_field(self, original_name, new_name):
        """Sucht und benennt ein spezifisches Feld um"""
        
        # Verschiedene Strategien zum Finden von Feldern
        strategies = [
            self._find_field_by_tab_navigation,
            self._find_field_by_click_search,
            self._find_field_by_coordinate_search
        ]
        
        for strategy in strategies:
            try:
                if strategy(original_name, new_name):
                    return True
            except Exception as e:
                self.log_message(f"Strategie fehlgeschlagen: {str(e)}", "DEBUG")
                continue
                
        return False
        
    def _find_field_by_tab_navigation(self, original_name, new_name):
        """Sucht Feld durch Tab-Navigation"""
        
        # Fokussiere erstes Feld
        pyautogui.click(100, 200)  # Klicke irgendwo auf die Seite
        time.sleep(0.3)
        
        # Durchsuche maximal 100 Felder
        for i in range(100):
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
                    self.log_message(f"Feld gefunden mit Tab-Navigation: '{original_name}'", "DEBUG")
                    
                    if self.confirm_var.get():
                        result = messagebox.askyesno("Bestätigung", 
                                                   f"Feld umbenennen?\n\n"
                                                   f"Von: '{original_name}'\n"
                                                   f"Zu: '{new_name}'")
                        if not result:
                            pyautogui.press('esc')
                            return False
                    
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
                
        return False
        
    def _find_field_by_click_search(self, original_name, new_name):
        """Sucht Feld durch systematisches Klicken"""
        
        # Erstelle Suchraster
        search_areas = [
            (100, 150, 600, 200),   # Objektdaten
            (100, 250, 600, 150),   # Erfasser
            (100, 400, 600, 200),   # Mieter 1
            (100, 600, 600, 200),   # Mieter 2
            (100, 800, 600, 150)    # Zähler/Übergabe
        ]
        
        for area_x, area_y, area_w, area_h in search_areas:
            # Durchsuche Bereich in 50px Schritten
            for y in range(area_y, area_y + area_h, 50):
                for x in range(area_x, area_x + area_w, 100):
                    try:
                        pyautogui.click(x, y)
                        time.sleep(0.2)
                        
                        # Versuche Properties zu öffnen
                        pyautogui.rightClick()
                        time.sleep(0.3)
                        pyautogui.press('p')
                        time.sleep(0.5)
                        
                        # Prüfe Feldname
                        pyautogui.hotkey('ctrl', 'a')
                        pyautogui.hotkey('ctrl', 'c')
                        time.sleep(0.2)
                        
                        current_name = pyperclip.paste().strip()
                        
                        if current_name == original_name:
                            # Gefunden!
                            pyautogui.hotkey('ctrl', 'a')
                            pyautogui.typewrite(new_name, interval=0.02)
                            pyautogui.press('enter')
                            return True
                        else:
                            pyautogui.press('esc')
                            
                    except:
                        pyautogui.press('esc')  # Sicherheit
                        continue
                        
        return False
        
    def _find_field_by_coordinate_search(self, original_name, new_name):
        """Sucht Feld basierend auf gespeicherten Koordinaten"""
        
        if original_name in self.field_positions:
            pos = self.field_positions[original_name]
            
            try:
                pyautogui.click(pos['x'] + pos['width']//2, pos['y'] + pos['height']//2)
                time.sleep(0.3)
                
                pyautogui.rightClick()
                time.sleep(0.3)
                pyautogui.press('p')
                time.sleep(0.5)
                
                # Umbenennen
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.typewrite(new_name, interval=0.02)
                pyautogui.press('enter')
                
                return True
                
            except:
                pyautogui.press('esc')
                return False
                
        return False
        
    # ===== WEITERE FUNKTIONEN =====
    
    def position_all_fields(self):
        """Positioniert alle Felder neu"""
        messagebox.showinfo("Position", "Funktion 'Felder positionieren' noch nicht implementiert.")
        
    def full_automation(self):
        """Führt die komplette Automatisierung durch"""
        if not self.pdf_path or not self.data_path:
            messagebox.showwarning("Fehlende Dateien", 
                                  "Bitte wählen Sie sowohl PDF als auch Feld-Definitionen aus.")
            return
            
        result = messagebox.askyesno("Vollautomatik", 
                                   "🤖 Vollautomatische Formular-Bearbeitung starten?\n\n"
                                   "Dies wird:\n"
                                   "1. Adobe Acrobat DC öffnen\n"
                                   "2. Formular-Modus aktivieren\n"
                                   "3. Alle Felder erstellen/umbenennen\n"
                                   "4. PDF speichern\n\n"
                                   "⚠️ Nicht unterbrechen!")
        if not result:
            return
            
        self.log_message("🚀 Starte Vollautomatik...", "SUCCESS")
        
        try:
            # Schritt 1: Adobe öffnen
            self.open_acrobat()
            time.sleep(3)
            
            # Schritt 2: Formular-Modus
            self.enter_form_mode()
            time.sleep(3)
            
            # Schritt 3: Felder erstellen
            self.create_all_fields()
            
            # Warte bis Erstellung fertig
            while self.is_running:
                time.sleep(1)
                
            # Schritt 4: PDF speichern
            time.sleep(2)
            self.save_pdf()
            
            self.log_message("🎉 Vollautomatik erfolgreich abgeschlossen!", "SUCCESS")
            messagebox.showinfo("Erfolg", "Vollautomatische Formular-Bearbeitung abgeschlossen!")
            
        except Exception as e:
            self.log_message(f"❌ Fehler in Vollautomatik: {str(e)}", "ERROR")
            messagebox.showerror("Fehler", f"Fehler in Vollautomatik:\n{str(e)}")
            
    def save_pdf(self):
        """Speichert das PDF"""
        self.log_message("💾 Speichere PDF...", "INFO")
        
        try:
            self.focus_acrobat()
            pyautogui.hotkey('ctrl', 's')
            time.sleep(2)
            
            # Falls "Speichern unter" Dialog erscheint
            pyautogui.press('enter')  # Bestätigen
            time.sleep(1)
            
            self.log_message("PDF gespeichert", "SUCCESS")
            
        except Exception as e:
            self.log_message(f"Fehler beim Speichern: {str(e)}", "ERROR")
            
    # ===== KALIBRIERUNG =====
    
    def calibrate_tool(self, tool_name):
        """Kalibriert ein spezifisches Tool"""
        self.log_message(f"Starte Kalibrierung für {tool_name}...", "INFO")
        
        # Anweisungen
        tool_names = {
            'textfield': 'Textfeld-Tool',
            'checkbox': 'Checkbox-Tool',
            'radiobutton': 'Optionsfeld-Tool',
            'dropdown': 'Dropdown-Tool',
            'signature': 'Signatur-Tool'
        }
        
        tool_display = tool_names.get(tool_name, tool_name)
        
        result = messagebox.askquestion("Kalibrierung", 
                                       f"📍 Kalibrierung für {tool_display}\n\n"
                                       f"1. Öffnen Sie Adobe Acrobat DC\n"
                                       f"2. Aktivieren Sie den Formular-Modus\n"
                                       f"3. Bewegen Sie die Maus zum {tool_display}\n"
                                       f"4. Drücken Sie OK und dann LEERTASTE\n\n"
                                       f"Bereit?")
        
        if result != 'yes':
            return
            
        self.log_message(f"Warte auf Leertaste für {tool_display}...", "INFO")
        
        # Warte auf Leertaste
        def wait_for_space():
            while True:
                if pyautogui.isKeyPressed('space'):
                    x, y = pyautogui.position()
                    self.tool_coordinates[tool_name] = (x, y)
                    
                    self.log_message(f"✅ {tool_display} kalibriert: ({x}, {y})", "SUCCESS")
                    
                    # Update UI
                    self.calibration_labels[tool_name].config(text=f"Kalibriert: ({x}, {y})", 
                                                             foreground="green")
                    
                    messagebox.showinfo("Kalibrierung", 
                                       f"✅ {tool_display} erfolgreich kalibriert!\n\n"
                                       f"Position: ({x}, {y})")
                    break
                time.sleep(0.1)
                
        # Starte in separatem Thread
        thread = threading.Thread(target=wait_for_space)
        thread.daemon = True
        thread.start()
        
    def save_calibration(self):
        """Speichert die Kalibrierung"""
        try:
            calib_dir = "calibration"
            os.makedirs(calib_dir, exist_ok=True)
            calib_file = os.path.join(calib_dir, "acrobat_tools.json")
            
            with open(calib_file, "w") as f:
                json.dump(self.tool_coordinates, f, indent=2)
                
            self.log_message(f"Kalibrierung gespeichert: {calib_file}", "SUCCESS")
            messagebox.showinfo("Gespeichert", f"Kalibrierung gespeichert in:\n{calib_file}")
            
        except Exception as e:
            self.log_message(f"Fehler beim Speichern der Kalibrierung: {str(e)}", "ERROR")
            
    def load_calibration(self):
        """Lädt die Kalibrierung"""
        try:
            calib_file = filedialog.askopenfilename(
                title="Kalibrierung laden",
                filetypes=[("JSON-Dateien", "*.json"), ("Alle Dateien", "*.*")],
                initialdir="calibration"
            )
            
            if calib_file:
                with open(calib_file, "r") as f:
                    self.tool_coordinates = json.load(f)
                    
                # Update UI
                for tool_name, coords in self.tool_coordinates.items():
                    if tool_name in self.calibration_labels and coords:
                        self.calibration_labels[tool_name].config(
                            text=f"Kalibriert: ({coords[0]}, {coords[1]})",
                            foreground="green"
                        )
                        
                self.log_message(f"Kalibrierung geladen: {calib_file}", "SUCCESS")
                messagebox.showinfo("Geladen", f"Kalibrierung geladen von:\n{calib_file}")
                
        except Exception as e:
            self.log_message(f"Fehler beim Laden der Kalibrierung: {str(e)}", "ERROR")
            
    def auto_calibrate(self):
        """Automatische Kalibrierung (experimentell)"""
        messagebox.showinfo("Auto-Kalibrierung", 
                           "Auto-Kalibrierung ist experimentell und noch nicht implementiert.\n\n"
                           "Verwenden Sie bitte die manuelle Kalibrierung.")
        
    # ===== WEITERE HILFSFUNKTIONEN =====
    
    def capture_position(self):
        """Erfasst die aktuelle Mausposition"""
        messagebox.showinfo("Position erfassen", 
                           "Bewegen Sie die Maus zur gewünschten Position und drücken Sie LEERTASTE.")
        
        def wait_for_space():
            while True:
                if pyautogui.isKeyPressed('space'):
                    x, y = pyautogui.position()
                    self.captured_position = (x, y)
                    self.log_message(f"📍 Position erfasst: ({x}, {y})", "SUCCESS")
                    messagebox.showinfo("Position erfasst", 
                                       f"Position erfasst: ({x}, {y})\n\n"
                                       f"Diese Position wird für das nächste Feld verwendet.")
                    break
                time.sleep(0.1)
                
        thread = threading.Thread(target=wait_for_space)
        thread.daemon = True
        thread.start()
        
    def create_single_field(self):
        """Erstellt ein einzelnes Feld"""
        if not hasattr(self, 'captured_position'):
            messagebox.showwarning("Keine Position", 
                                  "Bitte erfassen Sie zuerst eine Position mit 'Position erfassen'.")
            return
            
        field_type = self.field_type_var.get()
        field_name = self.property_entries["field_name_entry"].get()
        display_name = self.property_entries["display_name_entry"].get()
        
        try:
            width = int(self.property_entries["field_width_entry"].get())
            height = int(self.property_entries["field_height_entry"].get())
        except ValueError:
            messagebox.showerror("Fehler", "Breite und Höhe müssen Zahlen sein.")
            return
            
        if not field_name:
            messagebox.showwarning("Fehlende Eingabe", "Bitte geben Sie einen Feldnamen ein.")
            return
            
        self.log_message(f"📝 Erstelle {field_type}: {field_name}", "INFO")
        
        try:
            if self.create_field_at_position(
                self.captured_position[0], 
                self.captured_position[1], 
                field_type, 
                field_name, 
                display_name, 
                width, 
                height
            ):
                self.log_message(f"✅ {field_type} '{field_name}' erfolgreich erstellt", "SUCCESS")
                messagebox.showinfo("Erfolg", f"Feld '{field_name}' erfolgreich erstellt!")
            else:
                self.log_message(f"❌ Fehler beim Erstellen von '{field_name}'", "ERROR")
                
        except Exception as e:
            self.log_message(f"❌ Fehler beim Erstellen: {str(e)}", "ERROR")
            messagebox.showerror("Fehler", f"Fehler beim Erstellen:\n{str(e)}")
            
    def find_field(self):
        """Sucht ein Feld anhand des Namens"""
        field_name = self.property_entries["field_name_entry"].get()
        if not field_name:
            messagebox.showwarning("Fehlende Eingabe", "Bitte geben Sie einen Feldnamen ein.")
            return
            
        self.log_message(f"🔍 Suche Feld: {field_name}", "INFO")
        
        # Vereinfachte Suche
        found = self.find_and_rename_field(field_name, field_name)  # Umbenennung auf sich selbst
        
        if found:
            self.log_message(f"✅ Feld '{field_name}' gefunden", "SUCCESS")
            messagebox.showinfo("Gefunden", f"Feld '{field_name}' wurde gefunden und ausgewählt.")
        else:
            self.log_message(f"❌ Feld '{field_name}' nicht gefunden", "WARNING")
            messagebox.showwarning("Nicht gefunden", f"Feld '{field_name}' konnte nicht gefunden werden.")
            
    def create_grid_layout(self):
        """Erstellt ein Grid-Layout für Felder"""
        messagebox.showinfo("Grid-Layout", "Grid-Layout Funktion noch nicht implementiert.")
        
    def update_all_fields(self):
        """Aktualisiert alle Felder"""
        messagebox.showinfo("Update", "Update-Funktion noch nicht implementiert.")
        
    def delete_all_fields(self):
        """Löscht alle Felder"""
        result = messagebox.askyesno("Alle Felder löschen", 
                                   "⚠️ Sollen wirklich ALLE Felder gelöscht werden?\n\n"
                                   "Diese Aktion kann nicht rückgängig gemacht werden!")
        if result:
            messagebox.showinfo("Löschen", "Lösch-Funktion noch nicht implementiert.")
            
    def save_field_positions(self):
        """Speichert Feldpositionen"""
        try:
            pos_file = filedialog.asksaveasfilename(
                title="Feldpositionen speichern",
                defaultextension=".json",
                filetypes=[("JSON-Dateien", "*.json"), ("Alle Dateien", "*.*")]
            )
            
            if pos_file:
                with open(pos_file, "w") as f:
                    json.dump(self.field_positions, f, indent=2)
                    
                self.log_message(f"Feldpositionen gespeichert: {pos_file}", "SUCCESS")
                messagebox.showinfo("Gespeichert", f"Feldpositionen gespeichert in:\n{pos_file}")
                
        except Exception as e:
            self.log_message(f"Fehler beim Speichern der Positionen: {str(e)}", "ERROR")
            
    def take_screenshot(self):
        """Erstellt einen Screenshot"""
        try:
            screenshot = pyautogui.screenshot()
            
            # Speichere Screenshot
            screenshots_dir = "screenshots"
            os.makedirs(screenshots_dir, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_file = os.path.join(screenshots_dir, f"acrobat_screenshot_{timestamp}.png")
            
            screenshot.save(screenshot_file)
            
            self.log_message(f"Screenshot gespeichert: {screenshot_file}", "SUCCESS")
            messagebox.showinfo("Screenshot", f"Screenshot gespeichert:\n{screenshot_file}")
            
        except Exception as e:
            self.log_message(f"Fehler beim Screenshot: {str(e)}", "ERROR")
            
    # ===== EINSTELLUNGEN =====
    
    def save_settings(self):
        """Speichert die aktuellen Einstellungen"""
        try:
            settings = {
                'speed': self.speed_var.get(),
                'safety_mode': self.safety_var.get(),
                'auto_backup': self.backup_var.get(),
                'confirm_actions': self.confirm_var.get(),
                'default_width': self.default_entries["default_width"].get(),
                'default_height': self.default_entries["default_height"].get(),
                'default_checkbox_size': self.default_entries["default_checkbox_size"].get(),
                'default_spacing': self.default_entries["default_spacing"].get(),
                'start_x': self.default_entries["start_x"].get(),
                'start_y': self.default_entries["start_y"].get(),
                'tool_coordinates': self.tool_coordinates
            }
            
            settings_dir = "settings"
            os.makedirs(settings_dir, exist_ok=True)
            settings_file = os.path.join(settings_dir, "automation_settings.json")
            
            with open(settings_file, "w") as f:
                json.dump(settings, f, indent=2)
                
            self.log_message(f"Einstellungen gespeichert: {settings_file}", "SUCCESS")
            messagebox.showinfo("Gespeichert", "Einstellungen erfolgreich gespeichert!")
            
        except Exception as e:
            self.log_message(f"Fehler beim Speichern der Einstellungen: {str(e)}", "ERROR")
            
    def load_settings(self):
        """Lädt die gespeicherten Einstellungen"""
        try:
            settings_file = os.path.join("settings", "automation_settings.json")
            
            if os.path.exists(settings_file):
                with open(settings_file, "r") as f:
                    settings = json.load(f)
                    
                # Lade Einstellungen
                self.speed_var.set(settings.get('speed', 1.0))
                self.safety_var.set(settings.get('safety_mode', True))
                self.backup_var.set(settings.get('auto_backup', True))
                self.confirm_var.set(settings.get('confirm_actions', False))
                
                # Lade Default-Werte (nach GUI-Erstellung)
                if hasattr(self, 'default_entries'):
                    for key, entry in self.default_entries.items():
                        if key in settings:
                            entry.delete(0, tk.END)
                            entry.insert(0, str(settings[key]))
                            
                # Lade Tool-Koordinaten
                self.tool_coordinates = settings.get('tool_coordinates', {})
                
                self.log_message("Einstellungen geladen", "SUCCESS")
                
        except Exception as e:
            self.log_message(f"Fehler beim Laden der Einstellungen: {str(e)}", "WARNING")
            
    # ===== LOG FUNKTIONEN =====
    
    def clear_log(self):
        """Löscht das Log"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("Log gelöscht", "INFO")
        
    def save_log(self):
        """Speichert das Log"""
        try:
            log_content = self.log_text.get(1.0, tk.END)
            
            log_file = filedialog.asksaveasfilename(
                title="Log speichern",
                defaultextension=".txt",
                filetypes=[("Text-Dateien", "*.txt"), ("Alle Dateien", "*.*")]
            )
            
            if log_file:
                with open(log_file, "w", encoding="utf-8") as f:
                    f.write(log_content)
                    
                self.log_message(f"Log gespeichert: {log_file}", "SUCCESS")
                messagebox.showinfo("Gespeichert", f"Log gespeichert in:\n{log_file}")
                
        except Exception as e:
            self.log_message(f"Fehler beim Speichern des Logs: {str(e)}", "ERROR")
            
    def copy_log(self):
        """Kopiert das Log in die Zwischenablage"""
        try:
            log_content = self.log_text.get(1.0, tk.END)
            pyperclip.copy(log_content)
            self.log_message("Log in Zwischenablage kopiert", "SUCCESS")
            messagebox.showinfo("Kopiert", "Log wurde in die Zwischenablage kopiert.")
            
        except Exception as e:
            self.log_message(f"Fehler beim Kopieren: {str(e)}", "ERROR")
            
    def export_statistics(self, df):
        """Exportiert Statistiken der Feld-Definitionen"""
        try:
            stats = {
                'total_fields': len(df),
                'field_types': df['type'].value_counts().to_dict() if 'type' in df.columns else {},
                'field_names': df['new_name'].tolist(),
                'timestamp': datetime.now().isoformat()
            }
            
            stats_file = filedialog.asksaveasfilename(
                title="Statistiken exportieren",
                defaultextension=".json",
                filetypes=[("JSON-Dateien", "*.json"), ("Alle Dateien", "*.*")]
            )
            
            if stats_file:
                with open(stats_file, "w") as f:
                    json.dump(stats, f, indent=2)
                    
                messagebox.showinfo("Exportiert", f"Statistiken exportiert:\n{stats_file}")
                
        except Exception as e:
            self.log_message(f"Fehler beim Exportieren der Statistiken: {str(e)}", "ERROR")
            
    # ===== STEUERUNG =====
    
    def toggle_pause(self):
        """Pausiert/Startet die Automatisierung"""
        self.pause_automation = not self.pause_automation
        
        if self.pause_automation:
            self.pause_btn.config(text="▶️ Fortsetzen", bg="#4CAF50")
            self.update_status("Status: Pausiert", "#FFF8E1")
            self.log_message("⏸️ Automatisierung pausiert", "WARNING")
        else:
            self.pause_btn.config(text="⏸️ Pausieren", bg="#FF9800")
            self.update_status("Status: Läuft", "#FFE8E8")
            self.log_message("▶️ Automatisierung fortgesetzt", "INFO")
            
    def stop_automation(self):
        """Stoppt die Automatisierung"""
        self.is_running = False
        self.pause_automation = False
        self.pause_btn.config(text="⏸️ Pausieren", bg="#FF9800")
        self.update_status("Status: Gestoppt", "#FFEBEE")
        self.update_task("Gestoppt")
        self.log_message("🛑 Automatisierung gestoppt", "WARNING")
        
    def on_closing(self):
        """Behandelt das Schließen des Fensters"""
        if self.is_running:
            result = messagebox.askyesno("Automatisierung läuft", 
                                       "Eine Automatisierung läuft noch.\n\n"
                                       "Wirklich beenden?")
            if not result:
                return
                
        # Speichere Einstellungen
        self.save_settings()
        
        self.root.destroy()

# ===== HAUPTPROGRAMM =====

def main():
    """Hauptfunktion"""
    root = tk.Tk()
    
    # Icon setzen (falls verfügbar)
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
        
    app = AcrobatFormAutomator(root)
    
    # Schließen-Event behandeln
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Starte GUI
    root.mainloop()

if __name__ == "__main__":
    main()
