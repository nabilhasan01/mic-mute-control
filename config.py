import json
import os
import tkinter as tk

class ConfigManager:
    def __init__(self):
        self.position_var = tk.StringVar(value="Top Mid")
        self.size_var = tk.StringVar(value="48x48")
        self.margin_var = tk.StringVar(value="10")
        self.load_config()
    
    def load_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "config.json")
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    self.position_var.set(config.get("overlay_position", "Top Mid"))
                    self.size_var.set(config.get("overlay_size", "48x48"))
                    self.margin_var.set(str(config.get("overlay_margin", 10)))
                    print(f"Loaded config: position={self.position_var.get()}, size={self.size_var.get()}, margin={self.margin_var.get()}")
        except Exception as e:
            print(f"Error loading config: {str(e)}")
    
    def save_config(self, overlay_manager):
        config_path = os.path.join(os.path.dirname(__file__), "config.json")
        try:
            config = {
                "overlay_position": overlay_manager.position_var.get(),
                "overlay_size": overlay_manager.size_var.get(),
                "overlay_margin": int(overlay_manager.margin_var.get())
            }
            with open(config_path, 'w') as f:
                json.dump(config, f)
            print(f"Saved config: position={overlay_manager.position_var.get()}, size={overlay_manager.size_var.get()}, margin={overlay_manager.margin_var.get()}")
        except Exception as e:
            print(f"Error saving config: {str(e)}")