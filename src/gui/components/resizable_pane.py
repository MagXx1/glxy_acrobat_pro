import customtkinter as ctk

class ResizablePane(ctk.CTkFrame):
    """Resizable Pane mit Hover-Effekt"""
    
    def __init__(self, parent, min_height=100, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.min_height = min_height
        self.is_resizing = False
        self.start_y = 0
        self.start_height = 0
        
        # Resize Handle
        self.resize_handle = ctk.CTkFrame(self, height=5, cursor="sb_v_double_arrow")
        self.resize_handle.pack(side="bottom", fill="x")
        
        # Bind Events
        self.resize_handle.bind("<Button-1>", self.start_resize)
        self.resize_handle.bind("<B1-Motion>", self.do_resize)
        self.resize_handle.bind("<ButtonRelease-1>", self.end_resize)
        self.resize_handle.bind("<Enter>", self.on_hover_enter)
        self.resize_handle.bind("<Leave>", self.on_hover_leave)
        
    def start_resize(self, event):
        self.is_resizing = True
        self.start_y = event.y_root
        self.start_height = self.winfo_height()
        
    def do_resize(self, event):
        if self.is_resizing:
            delta_y = event.y_root - self.start_y
            new_height = max(self.min_height, self.start_height + delta_y)
            self.configure(height=new_height)
            
    def end_resize(self, event):
        self.is_resizing = False
        
    def on_hover_enter(self, event):
        self.resize_handle.configure(fg_color="#4a9eff")
        
    def on_hover_leave(self, event):
        self.resize_handle.configure(fg_color="#2b2b2b")
