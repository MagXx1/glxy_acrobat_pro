import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import time
import pyautogui
import os
import pyperclip
import json
import threading
import subprocess
import shutil
from datetime import datetime
import keyboard

# Versuche CTkPDFViewer zu importieren (optional)
try:
    from CTkPDFViewer import *
    PDF_VIEWER_AVAILABLE = True
except ImportError:
    PDF_VIEWER_AVAILABLE = False

# Lokale Imports
from .components.collapsible_frame import CollapsibleFrame
from .components.resizable_pane import ResizablePane
from ..automation.acrobat_controller import AcrobatController
from ..automation.field_operations import FieldOperations
from ..utils.logger import Logger
from ..utils.file_handler import SettingsManager, FileHandler

# Sicherheitseinstellungen
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.8

# CustomTkinter Konfiguration
ctk.set_appearance_mode("dark")  # Dunkles Theme
ctk.set_default_color_theme("blue")  # Blaues Farbschema

class AcrobatFormAutomatorPro(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("🤖 Adobe Acrobat GLXY Roboti PRO v5.0")
        self.geometry("1400x1000")
        self.minsize(1200, 800)
        
        # Initialisiere Komponenten
        self.logger = Logger()
        self.settings_manager = SettingsManager()
        self.file_handler = FileHandler()
        self.acrobat_controller = AcrobatController()
        self.field_operations = FieldOperations()
        
        # Variablen
        self.pdf_path = None
        self.data_path = None
        self.is_running = False
        self.pause_automation = False
        self.field_positions = {}
        self.created_fields = []
        self.current_task = ""
        self.captured_position = None
        
        # Tool-Koordinaten für Kalibrierung
        self.tool_coordinates = {
            'textfield': None,
            'checkbox': None,
            'radiobutton': None,
            'dropdown': None,
            'signature': None
        }
        
        self.create_gui()
        self.load_settings()
        
        # Initial-Log
        self.log_message("🚀 ROBOTI ADOBE ROBOTI GLXY PRO GESTARTET", "SUCCESS")
        self.log_message("Bereit für den Roboti GLXY Roboti Formular - Weltmeister?", "INFO")
        
    def create_gui(self):
        # Hauptlayout mit zwei Spalten
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)
        
        # Linke Seite: Steuerung
        self.left_panel = ctk.CTkFrame(self, corner_radius=10)
        self.left_panel.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        
        # Rechte Seite: PDF-Viewer und Arbeitsbereich
        self.right_panel = ctk.CTkFrame(self, corner_radius=10, fg_color="white")
        self.right_panel.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        
        self.create_left_panel()
        self.create_right_panel()
        
    def create_left_panel(self):
        # Scrollbarer Bereich für linkes Panel
        self.left_scrollable = ctk.CTkScrollableFrame(self.left_panel, width=400)
        self.left_scrollable.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Titel
        title_label = ctk.CTkLabel(
            self.left_scrollable,
            text="🤖 GLXY Form Roboti PRO",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Erstelle alle Sektionen
        self.create_file_section()
        self.create_main_functions()
        self.create_field_creation()
        self.create_settings_section()
        self.create_calibration_section()
        self.create_status_section()
        
    def create_file_section(self):
        # Datei-Auswahl Frame (minimierbar)
        file_frame = CollapsibleFrame(self.left_scrollable, title="📁 Datei-Auswahl")
        file_frame.pack(fill="x", pady=10)
        
        content = file_frame.get_content_frame()
        
        # PDF-Auswahl
        pdf_label = ctk.CTkLabel(content, text="Adobe Acrobat DC PDF:", 
                                font=ctk.CTkFont(weight="bold"))
        pdf_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        pdf_frame = ctk.CTkFrame(content)
        pdf_frame.pack(fill="x", padx=10, pady=5)
        
        self.pdf_label = ctk.CTkLabel(pdf_frame, text="Keine PDF ausgewählt", 
                                     anchor="w", font=ctk.CTkFont(size=10))
        self.pdf_label.pack(side="left", fill="x", expand=True, padx=10, pady=5)
        
        pdf_btn = ctk.CTkButton(pdf_frame, text="📂 PDF wählen", width=100,
                               command=self.select_pdf)
        pdf_btn.pack(side="right", padx=10, pady=5)
        
        # Daten-Auswahl
        data_label = ctk.CTkLabel(content, text="Feld-Definitionen:", 
                                 font=ctk.CTkFont(weight="bold"))
        data_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        data_frame = ctk.CTkFrame(content)
        data_frame.pack(fill="x", padx=10, pady=5)
        
        self.data_label = ctk.CTkLabel(data_frame, text="Keine Daten ausgewählt", 
                                      anchor="w", font=ctk.CTkFont(size=10))
        self.data_label.pack(side="left", fill="x", expand=True, padx=10, pady=5)
        
        data_btn = ctk.CTkButton(data_frame, text="📂 Daten laden", width=100,
                                command=self.select_data)
        data_btn.pack(side="right", padx=10, pady=5)
        
        # Vorschau Button
        preview_btn = ctk.CTkButton(content, text="👁️ Daten-Vorschau", 
                                   command=self.preview_data,
                                   hover_color="#1f538d")
        preview_btn.pack(pady=10)
        
    def create_main_functions(self):
        # Hauptfunktionen Frame
        main_frame = CollapsibleFrame(self.left_scrollable, title="🚀 Hauptfunktionen")
        main_frame.pack(fill="x", pady=10)
        
        content = main_frame.get_content_frame()
        
        # Button-Konfigurationen
        buttons = [
            ("🔧 Adobe Acrobat DC öffnen", self.open_acrobat, "#2196F3"),
            ("📝 Formular-Modus aktivieren", self.enter_form_mode, "#4CAF50"),
            ("➕ Alle Felder erstellen", self.create_all_fields, "#FF9800"),
            ("✏️ Felder umbenennen", self.rename_all_fields, "#9C27B0"),
            ("🎯 Magie - Vollautomatik", self.full_automation, "#F44336")
        ]
        
        for text, command, color in buttons:
            btn = ctk.CTkButton(
                content,
                text=text,
                command=command,
                height=40,
                font=ctk.CTkFont(size=12, weight="bold"),
                hover_color=color
            )
            btn.pack(fill="x", padx=10, pady=5)
            
    def create_field_creation(self):
        # Feld-Erstellung Frame
        field_frame = CollapsibleFrame(self.left_scrollable, title="📝 Feld-Erstellung")
        field_frame.pack(fill="x", pady=10)
        
        content = field_frame.get_content_frame()
        
        # Feldtyp-Auswahl
        type_label = ctk.CTkLabel(content, text="Feldtyp:", font=ctk.CTkFont(weight="bold"))
        type_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.field_type_var = ctk.StringVar(value="Textfeld")
        type_menu = ctk.CTkOptionMenu(
            content,
            variable=self.field_type_var,
            values=["Textfeld", "Checkbox", "Optionsfeld", "Dropdown", "Signatur"]
        )
        type_menu.pack(fill="x", padx=10, pady=5)
        
        # Feld-Eigenschaften
        props_frame = ctk.CTkFrame(content)
        props_frame.pack(fill="x", padx=10, pady=10)
        
        # Eingabefelder
        ctk.CTkLabel(props_frame, text="Feldname:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.field_name_entry = ctk.CTkEntry(props_frame, placeholder_text="Feldname eingeben")
        self.field_name_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        ctk.CTkLabel(props_frame, text="Anzeigename:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.display_name_entry = ctk.CTkEntry(props_frame, placeholder_text="Anzeigename")
        self.display_name_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        props_frame.columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = ctk.CTkFrame(content)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        capture_btn = ctk.CTkButton(btn_frame, text="📍 Position erfassen", 
                                   command=self.capture_position)
        capture_btn.pack(side="left", padx=5)
        
        create_btn = ctk.CTkButton(btn_frame, text="➕ Feld erstellen",
                                  command=self.create_single_field)
        create_btn.pack(side="right", padx=5)
        
    def create_settings_section(self):
        # Einstellungen Frame
        settings_frame = CollapsibleFrame(self.left_scrollable, title="⚙️ Einstellungen")
        settings_frame.pack(fill="x", pady=10)
        
        content = settings_frame.get_content_frame()
        
        # Geschwindigkeit
        speed_label = ctk.CTkLabel(content, text="RobotiRoboti-Geschwindigkeit:", 
                                  font=ctk.CTkFont(weight="bold"))
        speed_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.speed_var = ctk.DoubleVar(value=1.0)
        speed_slider = ctk.CTkSlider(content, from_=0.2, to=3.0, variable=self.speed_var)
        speed_slider.pack(fill="x", padx=10, pady=5)
        
        self.speed_label = ctk.CTkLabel(content, text="1.0 Sekunden")
        self.speed_label.pack(padx=10, pady=5)
        
        speed_slider.configure(command=self.update_speed_label)
        
        # Sicherheitsoptionen
        safety_frame = ctk.CTkFrame(content)
        safety_frame.pack(fill="x", padx=10, pady=10)
        
        self.safety_var = ctk.BooleanVar(value=True)
        safety_check = ctk.CTkCheckBox(safety_frame, text="Sicherheitsmodus", 
                                      variable=self.safety_var)
        safety_check.pack(anchor="w", padx=10, pady=5)
        
        self.backup_var = ctk.BooleanVar(value=True)
        backup_check = ctk.CTkCheckBox(safety_frame, text="Auto-Backup", 
                                      variable=self.backup_var)
        backup_check.pack(anchor="w", padx=10, pady=5)
        
        # Kontroll-Buttons
        control_frame = ctk.CTkFrame(content)
        control_frame.pack(fill="x", padx=10, pady=10)
        
        self.pause_btn = ctk.CTkButton(control_frame, text="⏸️ Pausieren", 
                                      command=self.toggle_pause)
        self.pause_btn.pack(side="left", padx=5)
        
        stop_btn = ctk.CTkButton(control_frame, text="🛑 Stoppen", 
                                command=self.stop_automation,
                                hover_color="#d32f2f")
        stop_btn.pack(side="right", padx=5)
        
    def create_calibration_section(self):
        # Kalibrierung Frame
        calib_frame = CollapsibleFrame(self.left_scrollable, title="📐 Tool-Kalibrierung")
        calib_frame.pack(fill="x", pady=10)
        calib_frame.toggle_frame()  # Standardmäßig eingeklappt
        
        content = calib_frame.get_content_frame()
        
        calib_label = ctk.CTkLabel(content, 
                                  text="Kalibrieren Sie die Adobe Acrobat DC Tools für präzise Automatisierung",
                                  wraplength=350)
        calib_label.pack(padx=10, pady=10)
        
        # Tool-Buttons
        tools = [
            ("📝 Textfeld-Tool", "textfield"),
            ("☑️ Checkbox-Tool", "checkbox"),
            ("🔘 Optionsfeld-Tool", "radiobutton")
        ]
        
        self.calibration_labels = {}
        
        for text, tool_name in tools:
            tool_frame = ctk.CTkFrame(content)
            tool_frame.pack(fill="x", padx=10, pady=5)
            
            calib_btn = ctk.CTkButton(tool_frame, text=f"📍 {text}", width=200,
                                     command=lambda t=tool_name: self.calibrate_tool(t))
            calib_btn.pack(side="left", padx=5, pady=5)
            
            status_label = ctk.CTkLabel(tool_frame, text="Nicht kalibriert", 
                                       text_color="#ff6b6b")
            status_label.pack(side="right", padx=5, pady=5)
            self.calibration_labels[tool_name] = status_label
            
    def create_status_section(self):
        # Status Frame (minimierbar und resizable)
        status_frame = CollapsibleFrame(self.left_scrollable, title="📊 Status & Protokoll")
        status_frame.pack(fill="x", pady=10)
        
        content = status_frame.get_content_frame()
        
        # Aktuelle Aufgabe
        self.task_label = ctk.CTkLabel(content, text="Aktuelle Aufgabe: Bereit", 
                                      font=ctk.CTkFont(weight="bold"))
        self.task_label.pack(anchor="w", padx=10, pady=5)
        
        # Status
        self.status_label = ctk.CTkLabel(content, text="Status: Bereit für Automatisierung",
                                        fg_color="#4CAF50", corner_radius=5)
        self.status_label.pack(fill="x", padx=10, pady=5)
        
        # Progress Bar
        self.progress = ctk.CTkProgressBar(content)
        self.progress.pack(fill="x", padx=10, pady=5)
        self.progress.set(0)
        
        # Log-Bereich (Resizable)
        log_pane = ResizablePane(content, min_height=150)
        log_pane.pack(fill="both", expand=True, padx=10, pady=10)
        
        log_label = ctk.CTkLabel(log_pane, text="📋 Aktivitätsprotokoll", 
                                font=ctk.CTkFont(weight="bold"))
        log_label.pack(anchor="w", padx=10, pady=5)
        
        # Log-Text mit Scrollbar
        self.log_text = ctk.CTkTextbox(log_pane, height=150, font=ctk.CTkFont(family="Consolas"))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Log-Buttons
        log_btn_frame = ctk.CTkFrame(log_pane)
        log_btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        clear_btn = ctk.CTkButton(log_btn_frame, text="🗑️ Löschen", width=80,
                                 command=self.clear_log)
        clear_btn.pack(side="left", padx=5)
        
        save_btn = ctk.CTkButton(log_btn_frame, text="💾 Speichern", width=80,
                                command=self.save_log)
        save_btn.pack(side="right", padx=5)
        
    def create_right_panel(self):
        # Rechtes Panel mit PDF-Viewer und Arbeitsbereich
        self.right_panel.grid_rowconfigure(0, weight=1)
        self.right_panel.grid_columnconfigure(0, weight=1)
        
        # Header für rechtes Panel
        header_frame = ctk.CTkFrame(self.right_panel, height=50, fg_color="#2b2b2b")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        header_frame.grid_propagate(False)
        
        header_label = ctk.CTkLabel(header_frame, text="📄 PDF-Arbeitsbereich", 
                                   font=ctk.CTkFont(size=16, weight="bold"),
                                   text_color="white")
        header_label.pack(side="left", padx=20, pady=15)
        
        # PDF-Steuerung
        pdf_controls = ctk.CTkFrame(header_frame, fg_color="transparent")
        pdf_controls.pack(side="right", padx=20, pady=10)
        
        open_pdf_btn = ctk.CTkButton(pdf_controls, text="📂 PDF öffnen", 
                                    command=self.open_pdf_in_viewer)
        open_pdf_btn.pack(side="left", padx=5)
        
        reload_btn = ctk.CTkButton(pdf_controls, text="🔄 Neu laden",
                                  command=self.reload_pdf)
        reload_btn.pack(side="left", padx=5)
        
        # PDF-Viewer Bereich
        self.pdf_viewer_frame = ctk.CTkFrame(self.right_panel, fg_color="white")
        self.pdf_viewer_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.right_panel.grid_rowconfigure(1, weight=1)
        
        # Placeholder für PDF-Viewer
        self.create_pdf_placeholder()
        
    def create_pdf_placeholder(self):
        """Erstellt einen Placeholder für den PDF-Viewer"""
        placeholder_frame = ctk.CTkFrame(self.pdf_viewer_frame, fg_color="#f8f9fa")
        placeholder_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        icon_label = ctk.CTkLabel(placeholder_frame, text="📄", 
                                 font=ctk.CTkFont(size=80))
        icon_label.pack(pady=(50, 20))
        
        text_label = ctk.CTkLabel(placeholder_frame, 
                                 text="PDF-Dokument hier anzeigen\n\nWählen Sie ein PDF-Dokument aus und klicken Sie auf 'PDF öffnen'",
                                 font=ctk.CTkFont(size=14),
                                 text_color="#666666")
        text_label.pack(pady=20)
        
        self.placeholder_frame = placeholder_frame
        
    def open_pdf_in_viewer(self):
        """Öffnet PDF im integrierten Viewer"""
        if not self.pdf_path:
            messagebox.showwarning("Keine PDF", "Bitte wählen Sie zuerst eine PDF-Datei aus.")
            return
            
        try:
            # Entferne Placeholder
            if hasattr(self, 'placeholder_frame'):
                self.placeholder_frame.destroy()
                
            if PDF_VIEWER_AVAILABLE:
                # Erstelle PDF-Viewer
                self.pdf_viewer = CTkPDFViewer(
                    self.pdf_viewer_frame,
                    file=self.pdf_path,
                    page_width=800,
                    page_height=600
                )
                self.pdf_viewer.pack(fill="both", expand=True, padx=10, pady=10)
                
                self.log_message(f"PDF geöffnet: {os.path.basename(self.pdf_path)}", "SUCCESS")
            else:
                # Fallback: Zeige Info dass PDF-Viewer nicht verfügbar ist
                info_frame = ctk.CTkFrame(self.pdf_viewer_frame, fg_color="#f0f0f0")
                info_frame.pack(fill="both", expand=True, padx=20, pady=20)
                
                info_label = ctk.CTkLabel(info_frame, 
                                         text="📄 PDF-Viewer nicht verfügbar\n\nInstallieren Sie CTkPDFViewer für die Vorschau",
                                         font=ctk.CTkFont(size=14),
                                         text_color="#666666")
                info_label.pack(expand=True)
                
                self.log_message("PDF-Viewer nicht verfügbar - CTkPDFViewer installieren. Ansonsten Pech gehabt.", "WARNING")
            
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen der PDF: {str(e)}", "ERROR")
            messagebox.showerror("PDF-Fehler", f"Konnte PDF nicht öffnen:\n{str(e)}")
            
    def reload_pdf(self):
        """Lädt die PDF neu"""
        if hasattr(self, 'pdf_viewer'):
            try:
                self.pdf_viewer.destroy()
                self.open_pdf_in_viewer()
            except Exception as e:
                self.log_message(f"Fehler beim Neuladen: {str(e)}", "ERROR")
                
    # ===== HAUPTFUNKTIONEN =====
    
    def select_pdf(self):
        """PDF-Datei auswählen"""
        file = filedialog.askopenfilename(
            title="Adobe Acrobat DC PDF auswählen",
            filetypes=[("PDF-Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        if file:
            self.pdf_path = file
            filename = os.path.basename(file)
            self.pdf_label.configure(text=filename)
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
            self.data_label.configure(text=filename)
            self.log_message(f"Feld-Definitionen geladen: {filename}", "SUCCESS")
            
    def preview_data(self):
        """Zeigt eine Vorschau der geladenen Daten"""
        df = self.file_handler.load_field_definitions(self.data_path)
        if df is None:
            return
            
        # Vorschau-Fenster
        preview_window = ctk.CTkToplevel(self)
        preview_window.title("📊 Feld-Definitionen Vorschau")
        preview_window.geometry("900x600")
        
        # Header
        header_label = ctk.CTkLabel(preview_window, text="📊 Feld-Definitionen Vorschau", 
                                   font=ctk.CTkFont(size=18, weight="bold"))
        header_label.pack(pady=20)
        
        # Statistiken
        total_fields = len(df)
        field_types = df['type'].value_counts() if 'type' in df.columns else {}
        
        stats_text = f"Gesamt: {total_fields} Felder"
        if not field_types.empty:
            stats_text += " | " + " | ".join([f"{t}: {c}" for t, c in field_types.items()])
            
        stats_label = ctk.CTkLabel(preview_window, text=stats_text, 
                                  font=ctk.CTkFont(size=12, weight="bold"))
        stats_label.pack(pady=10)
        
        # Schließen Button
        close_btn = ctk.CTkButton(preview_window, text="✅ Schließen", 
                                 command=preview_window.destroy)
        close_btn.pack(pady=20)
        
    def open_acrobat(self):
        """Öffnet Adobe Acrobat DC und das PDF"""
        if not self.pdf_path:
            messagebox.showwarning("Keine PDF", "Bitte wählen Sie zuerst eine PDF-Datei aus.")
            return
            
        self.update_task("Adobe Acrobat DC öffnen")
        success = self.acrobat_controller.launch_acrobat(self.pdf_path)
        
        if success:
            self.log_message("Adobe Acrobat DC geöffnet", "SUCCESS")
            time.sleep(2)
            self.acrobat_controller.focus_acrobat()
        else:
            self.log_message("Fehler beim Öffnen von Adobe Acrobat DC", "ERROR")
        
    def enter_form_mode(self):
        """Aktiviert den Formular-Bearbeitungsmodus"""
        self.update_task("Formular-Modus aktivieren")
        success = self.acrobat_controller.enter_form_mode(self.speed_var.get())
        
        if success:
            self.update_status("Status: Formular-Modus aktiv", "#2196F3")
            self.log_message("Formular-Modus aktiviert", "SUCCESS")
        else:
            self.log_message("Fehler beim Aktivieren des Formular-Modus", "ERROR")
            
    def create_all_fields(self):
        """Erstellt alle Felder basierend auf den Definitionen"""
        if not self.pdf_path or not self.data_path:
            messagebox.showwarning("Fehlende Dateien", 
                                  "Bitte wählen Sie sowohl PDF als auch Feld-Definitionen aus.")
            return
            
        df = self.file_handler.load_field_definitions(self.data_path)
        if df is None:
            return
            
        result = messagebox.askyesno("Alle Felder erstellen", 
                                   f"🤖 Sollen {len(df)} Felder automatisch erstellt werden?\n\n"
                                   "⚠️ Voraussetzungen:\n"
                                   "• Adobe Acrobat DC ist geöffnet\n"
                                   "• PDF ist geladen\n"
                                   "• Formular-Modus ist aktiv")
        if not result:
            return
            
        # Starte Automatisierung in separatem Thread
        self.is_running = True
        self.update_status("Status: Erstelle Felder...", "#FF9800")
        
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
            
            self.log_message(f"Starte Erstellung von {total} Feldern...", "INFO")
            
            for index, row in df.iterrows():
                if not self.is_running:
                    break
                    
                while self.pause_automation:
                    time.sleep(0.5)
                    
                field_name = str(row['new_name'])
                field_type = str(row.get('type', 'Textfeld'))
                display_name = str(row.get('display_name', ''))
                
                self.log_message(f"Erstelle Feld {index+1}/{total}: {field_name}", "INFO")
                self.update_progress(index / total)
                
                try:
                    # Berechne Position (einfaches Grid)
                    x = 100
                    y = 150 + (index * 40)
                    
                    # Erstelle Feld
                    success = self.field_operations.create_field_at_position(
                        x, y, field_type, field_name, display_name, 200, 25
                    )
                    
                    if success:
                        successful += 1
                        self.log_message(f"✅ '{field_name}' erfolgreich erstellt", "SUCCESS")
                    else:
                        failed += 1
                        self.log_message(f"❌ Fehler bei '{field_name}'", "ERROR")
                        
                except Exception as e:
                    failed += 1
                    self.log_message(f"❌ Fehler bei '{field_name}': {str(e)}", "ERROR")
                    
            # Zusammenfassung
            self.update_progress(1.0)
            self.log_message(f"🏁 Feld-Erstellung abgeschlossen!", "SUCCESS")
            self.log_message(f"✅ Erfolgreich: {successful} Felder 👌", "SUCCESS")
            self.log_message(f"❌ Fehler: {failed} Felder", "ERROR" if failed > 0 else "INFO")
            
            # Benachrichtigung
            self.after(0, lambda: messagebox.showinfo("Feld-Erstellung abgeschlossen", 
                                       f"✅ Erfolgreich: {successful} Felder\n"
                                       f"❌ Fehler: {failed} Felder"))
                                       
        except Exception as e:
            self.log_message(f"❌ Schwerwiegender Fehler: {str(e)}", "ERROR")
            
        finally:
            self.is_running = False
            self.after(0, lambda: self.update_status("Status: Bereit", "#4CAF50"))
            self.after(0, lambda: self.update_task("Bereit"))
            
    def rename_all_fields(self):
        """Benennt alle bestehenden Felder um"""
        if not self.data_path:
            messagebox.showwarning("Keine Daten", "Bitte laden Sie zuerst Feld-Definitionen.")
            return
            
        df = self.file_handler.load_field_definitions(self.data_path)
        if df is None:
            return
            
        result = messagebox.askyesno("Felder umbenennen", 
                                   f"🔄 Sollen {len(df)} Felder umbenannt werden?")
        if not result:
            return
            
        # Starte Umbenennung in separatem Thread
        self.is_running = True
        self.update_status("Status: Benenne Felder um...", "#9C27B0")
        
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
                    break
                    
                while self.pause_automation:
                    time.sleep(0.5)
                    
                original_name = str(row['original_name'])
                new_name = str(row['new_name'])
                
                self.log_message(f"Suche Feld {index+1}/{total}: '{original_name}'", "INFO")
                self.update_progress(index / total)
                
                try:
                    if self.field_operations.find_and_rename_field(original_name, new_name):
                        successful += 1
                        self.log_message(f"✅ '{original_name}' → '{new_name}'", "SUCCESS")
                    else:
                        failed += 1
                        self.log_message(f"❌ Feld '{original_name}' nicht gefunden", "WARNING")
                        
                except Exception as e:
                    failed += 1
                    self.log_message(f"❌ Fehler bei '{original_name}': {str(e)}", "ERROR")
                    
            # Zusammenfassung
            self.update_progress(1.0)
            self.log_message(f"🏁 Umbenennung abgeschlossen!", "SUCCESS")
            self.log_message(f"✅ Erfolgreich: {successful} Felder", "SUCCESS")
            self.log_message(f"❌ Fehler: {failed} Felder", "ERROR" if failed > 0 else "INFO")
            
            # Benachrichtigung
            self.after(0, lambda: messagebox.showinfo("Umbenennung abgeschlossen", 
                                       f"✅ Erfolgreich: {successful} Felder\n"
                                       f"❌ Nicht gefunden: {failed} Felder"))
                                       
        except Exception as e:
            self.log_message(f"❌ Schwerwiegender Fehler: {str(e)}", "ERROR")
            
        finally:
            self.is_running = False
            self.after(0, lambda: self.update_status("Status: Bereit", "#4CAF50"))
            self.after(0, lambda: self.update_task("Bereit"))
            
    def full_automation(self):
        """Führt die komplette Automatisierung durch"""
        if not self.pdf_path or not self.data_path:
            messagebox.showwarning("Fehlende Dateien", 
                                  "Bitte wählen Sie sowohl PDF als auch Feld-Definitionen aus.")
            return
            
        result = messagebox.askyesno("Vollautomatik", 
                                   "🤖 Vollautomatische Formular-Bearbeitung starten?\n\n"
                                   "Dies wird alle Schritte automatisch durchführen.")
        if result:
            self.log_message("🚀 Starte die Magie -  Vollautomatik...", "SUCCESS")
            
            # Führe Schritte nacheinander aus
            self.open_acrobat()
            time.sleep(3)
            self.enter_form_mode()
            time.sleep(2)
            self.create_all_fields()
            
    def capture_position(self):
        """Erfasst die aktuelle Mausposition mit verbesserter Keyboard-Erkennung"""
        messagebox.showinfo("Position erfassen", 
                           "Bewegen Sie die Maus zur gewünschten Position und drücken Sie LEERTASTE.")
        
        def wait_for_space():
            try:
                self.log_message("Warte auf Leertaste...", "INFO")
                position = self.field_operations.capture_position_with_keyboard()
                
                if position:
                    self.captured_position = position
                    self.after(0, lambda: messagebox.showinfo("Position erfasst", 
                                               f"Position erfasst: {position}"))
                    self.log_message(f"Position erfasst: {position}", "SUCCESS")
                    
            except Exception as e:
                self.log_message(f"Fehler beim Erfassen: {str(e)}", "ERROR")
                
        thread = threading.Thread(target=wait_for_space)
        thread.daemon = True
        thread.start()
        
    def create_single_field(self):
        """Erstellt ein einzelnes Feld"""
        if not hasattr(self, 'captured_position') or not self.captured_position:
            messagebox.showwarning("Keine Position", 
                                  "Bitte erfassen Sie zuerst eine Position.")
            return
            
        field_type = self.field_type_var.get()
        field_name = self.field_name_entry.get()
        display_name = self.display_name_entry.get()
        
        if not field_name:
            messagebox.showwarning("Fehlende Eingabe", "Bitte geben Sie einen Feldnamen ein.")
            return
            
        self.log_message(f"📝 Erstelle {field_type}: {field_name}", "INFO")
        
        try:
            success = self.field_operations.create_field_at_position(
                self.captured_position[0], 
                self.captured_position[1], 
                field_type, 
                field_name, 
                display_name, 
                200, 
                25
            )
            
            if success:
                self.log_message(f"✅ '{field_name}' erfolgreich erstellt", "SUCCESS")
                messagebox.showinfo("Erfolg", f"Feld '{field_name}' erfolgreich erstellt!")
            else:
                self.log_message(f"❌ Fehler beim Erstellen von '{field_name}'", "ERROR")
                
        except Exception as e:
            self.log_message(f"❌ Fehler: {str(e)}", "ERROR")
            messagebox.showerror("Fehler", f"Fehler beim Erstellen:\n{str(e)}")
        
    def calibrate_tool(self, tool_name):
        """Kalibriert ein spezifisches Tool mit verbesserter Keyboard-Erkennung"""
        tool_names = {
            'textfield': 'Textfeld-Tool',
            'checkbox': 'Checkbox-Tool',
            'radiobutton': 'Optionsfeld-Tool'
        }
        
        tool_display = tool_names.get(tool_name, tool_name)
        
        result = messagebox.askquestion("Kalibrierung", 
                                       f"📍 Kalibrierung für {tool_display}\n\n"
                                       f"1. Öffnen Sie Adobe Acrobat DC\n"
                                       f"2. Aktivieren Sie den Formular-Modus\n"
                                       f"3. Bewegen Sie die Maus zum {tool_display}\n"
                                       f"4. Drücken Sie LEERTASTE\n\n"
                                       f"Bereit?")
        
        if result != 'yes':
            return
            
        def wait_for_calibration():
            try:
                self.log_message(f"Warte auf Leertaste für {tool_display}...", "INFO")
                keyboard.wait('space')
                
                x, y = pyautogui.position()
                self.tool_coordinates[tool_name] = (x, y)
                
                self.log_message(f"✅ {tool_display} kalibriert: ({x}, {y})", "SUCCESS")
                
                # Update UI
                self.after(0, lambda: self.calibration_labels[tool_name].configure(
                    text=f"Kalibriert: ({x}, {y})", text_color="#4CAF50"))
                
                self.after(0, lambda: messagebox.showinfo("Kalibrierung", 
                                           f"✅ {tool_display} erfolgreich kalibriert!\n"
                                           f"Position: ({x}, {y})"))
                
            except Exception as e:
                self.log_message(f"Fehler bei Kalibrierung: {str(e)}", "ERROR")
                
        thread = threading.Thread(target=wait_for_calibration)
        thread.daemon = True
        thread.start()
        
    # ===== HILFSFUNKTIONEN =====
    
    def log_message(self, message, level="INFO"):
        """Erweiterte Log-Funktion mit Leveln und Farben"""
        formatted_message = self.logger.log(message, level)
        
        # Füge zu GUI-Log hinzu
        self.log_text.insert("end", formatted_message + "\n")
        self.log_text.see("end")
        
        # Aktualisiere UI
        self.update_idletasks()
        
    def update_status(self, message, color="#4CAF50"):
        """Aktualisiert die Statusanzeige"""
        self.status_label.configure(text=message, fg_color=color)
        
    def update_task(self, task):
        """Aktualisiert die aktuelle Aufgabe"""
        self.current_task = task
        self.task_label.configure(text=f"Aktuelle Aufgabe: {task}")
        
    def update_progress(self, value):
        """Aktualisiert die Progress Bar"""
        self.progress.set(value)
        
    def update_speed_label(self, value):
        """Aktualisiert das Geschwindigkeits-Label"""
        self.speed_label.configure(text=f"{float(value):.1f} Sekunden")
        
    def toggle_pause(self):
        """Pausiert/Startet die Automatisierung"""
        self.pause_automation = not self.pause_automation
        
        if self.pause_automation:
            self.pause_btn.configure(text="▶️ Fortsetzen")
            self.update_status("Status: Pausiert", "#FF9800")
            self.log_message("⏸️ Automatisierung pausiert", "WARNING")
        else:
            self.pause_btn.configure(text="⏸️ Pausieren")
            self.update_status("Status: Läuft", "#2196F3")
            self.log_message("▶️ Automatisierung fortgesetzt", "INFO")
            
    def stop_automation(self):
        """Stoppt die Automatisierung"""
        self.is_running = False
        self.pause_automation = False
        self.pause_btn.configure(text="⏸️ Pausieren")
        self.update_status("Status: Gestoppt", "#F44336")
        self.update_task("Gestoppt")
        self.log_message("🛑 Automatisierung gestoppt", "WARNING")
        
    def clear_log(self):
        """Löscht das Log"""
        self.log_text.delete("0.0", "end")
        self.log_message("Log gelöscht", "INFO")
        
    def save_log(self):
        """Speichert das Log"""
        try:
            log_content = self.log_text.get("0.0", "end")
            
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
            
    def save_settings(self):
        """Speichert die aktuellen Einstellungen"""
        try:
            settings = {
                'speed': self.speed_var.get(),
                'safety_mode': self.safety_var.get(),
                'auto_backup': self.backup_var.get(),
                'tool_coordinates': self.tool_coordinates,
                'window_geometry': self.geometry()
            }
            
            self.settings_manager.save_settings(settings)
            
        except Exception as e:
            self.log_message(f"Fehler beim Speichern der Einstellungen: {str(e)}", "ERROR")
            
    def load_settings(self):
        """Lädt die gespeicherten Einstellungen"""
        try:
            settings = self.settings_manager.load_settings()
            
            # Lade Einstellungen
            if 'speed' in settings:
                self.speed_var.set(settings['speed'])
            if 'safety_mode' in settings:
                self.safety_var.set(settings['safety_mode'])
            if 'auto_backup' in settings:
                self.backup_var.set(settings['auto_backup'])
            if 'tool_coordinates' in settings:
                self.tool_coordinates = settings['tool_coordinates']
            if 'window_geometry' in settings:
                self.geometry(settings['window_geometry'])
                
        except Exception as e:
            self.log_message(f"Hinweis: Keine gespeicherten Einstellungen gefunden", "DEBUG")
            
    def on_closing(self):
        """Behandelt das Schließen des Fensters"""
        if self.is_running:
            result = messagebox.askyesno("Automatisierung läuft", 
                                       "Eine Automatisierung läuft noch.\n\n"
                                       "´Wollens wirklich beenden? Muss ich böse werden?")
            if not result:
                return
                
        # Speichere Einstellungen
        self.save_settings()
        self.destroy()
