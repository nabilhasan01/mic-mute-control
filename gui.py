import tkinter as tk
from tkinter import messagebox, ttk
import keyboard
import threading

class MicMuteGUI:
    def __init__(self, root, toggle_mute, hide_window, start_hotkey_capture, 
                 update_overlay_position, update_overlay_size, update_margin):
        self.root = root
        self.root.title("Microphone Mute Control")
        self.root.geometry("244x320")
        self.root.resizable(False, True)
        self.root.configure(bg="#f0f0f0")
        
        self.toggle_mute = toggle_mute
        self.hide_window = hide_window
        self.start_hotkey_capture_callback = start_hotkey_capture
        self.update_overlay_position = update_overlay_position
        self.update_overlay_size = update_overlay_size
        self.update_margin = update_margin
        
        self.setup_styles()
        self.create_widgets()
        self.is_capturing_hotkey = False
    
    def setup_styles(self):
        style = ttk.Style()
        style.configure("TButton", padding=5, font=("Helvetica", 10))
        style.configure("TLabel", background="#f0f0f0", font=("Helvetica", 10))
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TCombobox", font=("Helvetica", 10))
    
    def create_widgets(self):
        self.title_label = ttk.Label(self.root, text="Microphone Control", font=("Helvetica", 12, "bold"))
        self.title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        self.label = ttk.Label(self.root, text="Status: Unknown")
        self.label.grid(row=1, column=0, columnspan=2, pady=5)
        
        self.toggle_button = ttk.Button(self.root, text="Toggle Mute", command=self.toggle_mute)
        self.toggle_button.grid(row=2, column=0, columnspan=2, pady=5)
        
        self.minimize_button = ttk.Button(self.root, text="Minimize to Tray", command=self.hide_window)
        self.minimize_button.grid(row=3, column=0, columnspan=2, pady=5)
        
        self.hotkey_label = ttk.Label(self.root, text="Hotkey:")
        self.hotkey_label.grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.label_hotkey = ttk.Label(self.root, text="ctrl+alt+m")
        self.label_hotkey.grid(row=4, column=1, sticky="w", padx=5)
        self.set_hotkey_button = ttk.Button(self.root, text="Set Hotkey", command=self.start_hotkey_capture)
        self.set_hotkey_button.grid(row=5, column=0, columnspan=2, pady=5)
        
        self.overlay_frame = ttk.LabelFrame(self.root, text="Overlay Settings", padding=5)
        self.overlay_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)
        
        self.position_label = ttk.Label(self.overlay_frame, text="Position")
        self.position_label.grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.position_var = tk.StringVar(value="Top Mid")
        positions = ["Top Left", "Top Mid", "Top Right", "Middle Left", "Middle Right", 
                     "Bottom Left", "Bottom Mid", "Bottom Right"]
        self.position_menu = ttk.Combobox(self.overlay_frame, textvariable=self.position_var, 
                                        values=positions, state="readonly", width=15)
        self.position_menu.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        self.position_menu.bind("<<ComboboxSelected>>", lambda e: self.update_overlay_position(self.position_var.get()))
        
        self.size_label = ttk.Label(self.overlay_frame, text="Size")
        self.size_label.grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.size_var = tk.StringVar(value="48x48")
        sizes = ["32x32", "48x48", "64x64"]
        self.size_menu = ttk.Combobox(self.overlay_frame, textvariable=self.size_var, 
                                     values=sizes, state="readonly", width=15)
        self.size_menu.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        self.size_menu.bind("<<ComboboxSelected>>", lambda e: self.update_overlay_size(self.size_var.get()))
        
        self.margin_label = ttk.Label(self.overlay_frame, text="Margin")
        self.margin_label.grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.margin_var = tk.StringVar(value="10")
        self.margin_entry = ttk.Entry(self.overlay_frame, textvariable=self.margin_var, width=5)
        self.margin_entry.grid(row=2, column=1, sticky="w", padx=(5, 0), pady=2)
        self.margin_button = ttk.Button(self.overlay_frame, text="Set", command=self.update_margin)
        self.margin_button.grid(row=2, column=1, sticky="w", padx=(50, 5), pady=2)
    
    def start_hotkey_capture(self):
        if self.is_capturing_hotkey:
            return
        self.is_capturing_hotkey = True
        self.set_hotkey_button.config(text="Press Keys...", state="disabled")
        self.root.update()
        try:
            threading.Thread(target=self.start_hotkey_capture_callback, daemon=True).start()
        except Exception as e:
            self.is_capturing_hotkey = False
            self.set_hotkey_button.config(text="Set Hotkey", state="normal")
            messagebox.showerror("Error", f"Failed to start hotkey capture: {str(e)}")
    
    def set_hotkey(self, new_hotkey, callback):
        try:
            if not new_hotkey:
                raise ValueError("No hotkey captured")
            keyboard.add_hotkey(new_hotkey, callback)
            self.label_hotkey.config(text=f"{new_hotkey}")
            print(f"Captured and set hotkey: {new_hotkey}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set hotkey: {str(e)}")
        finally:
            self.is_capturing_hotkey = False
            self.set_hotkey_button.config(text="Set Hotkey", state="normal")
            self.root.update()
    
    def remove_hotkey(self, hotkey):
        try:
            keyboard.remove_hotkey(hotkey)
        except:
            pass
    
    def show_window(self):
        self.root.deiconify()
    
    def hide_window(self):
        self.root.withdraw()
    
    def update_status(self, status):
        self.label.config(text=f"Status: {status}")