import customtkinter as ctk

class CollapsibleFrame(ctk.CTkFrame):
    """Minimierbare/Erweiterbare Frame-Komponente"""
    
    def __init__(self, parent, title="", **kwargs):
        super().__init__(parent, **kwargs)
        
        self.title = title
        self.is_expanded = True
        
        # Header mit Titel und Toggle-Button
        self.header_frame = ctk.CTkFrame(self, height=40)
        self.header_frame.pack(fill="x", padx=5, pady=2)
        self.header_frame.pack_propagate(False)
        
        # Toggle Button
        self.toggle_btn = ctk.CTkButton(
            self.header_frame,
            text="▼",
            width=30,
            height=30,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.toggle_frame
        )
        self.toggle_btn.pack(side="left", padx=5, pady=5)
        
        # Titel Label
        self.title_label = ctk.CTkLabel(
            self.header_frame,
            text=title,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.title_label.pack(side="left", padx=10, pady=5)
        
        # Content Frame
        self.content_frame = ctk.CTkFrame(self)
        self.content_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        
    def toggle_frame(self):
        """Toggle zwischen erweitert und minimiert"""
        if self.is_expanded:
            self.content_frame.pack_forget()
            self.toggle_btn.configure(text="▶")
            self.is_expanded = False
        else:
            self.content_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))
            self.toggle_btn.configure(text="▼")
            self.is_expanded = True
    
    def get_content_frame(self):
        """Gibt den Content-Frame zurück für weitere Widgets"""
        return self.content_frame
