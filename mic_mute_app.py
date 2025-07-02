import tkinter as tk
from gui import MicMuteGUI
from system_tray import SystemTray
from overlay import OverlayManager
from config import ConfigManager
from audio import AudioManager
import pythoncom

class MicMuteApp:
    def __init__(self, root):
        self.root = root
        pythoncom.CoInitialize()
        
        self.audio_manager = AudioManager()
        if not self.audio_manager.volume:
            self.root.quit()
            return
        
        self.config_manager = ConfigManager()
        self.overlay_manager = OverlayManager(self.root, self.audio_manager, self.config_manager)
        self.system_tray = SystemTray(self.toggle_mute, self.show_window, self.exit_app)
        self.gui = MicMuteGUI(self.root, self.toggle_mute, self.hide_window, 
                             self.start_hotkey_capture, self.update_overlay_position, 
                             self.update_overlay_size, self.update_margin)
        
        self.current_hotkey = None
        self.set_initial_hotkey()
        self.root.withdraw()
        self.update_status()
        self.audio_manager.poll_mute_state(self.update_status)
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)
    
    def set_initial_hotkey(self):
        try:
            self.current_hotkey = "ctrl+alt+m"
            self.gui.set_hotkey(self.current_hotkey, self.toggle_mute)
            print(f"Initial hotkey set: {self.current_hotkey}")
        except Exception as e:
            print(f"Error setting initial hotkey: {str(e)}")
    
    def start_hotkey_capture(self):
        self.gui.start_hotkey_capture(self.toggle_mute, self.set_hotkey)
    
    def set_hotkey(self, new_hotkey):
        self.current_hotkey = new_hotkey
    
    def toggle_mute(self):
        self.audio_manager.toggle_mute()
        self.update_status()
    
    def update_status(self):
        self.audio_manager.update_status(self.gui, self.system_tray, self.overlay_manager)
    
    def show_window(self):
        self.gui.show_window()
    
    def hide_window(self):
        self.gui.hide_window()
    
    def update_overlay_position(self, position):
        self.overlay_manager.update_position(position)
        self.config_manager.save_config(self.overlay_manager)
    
    def update_overlay_size(self, size):
        self.overlay_manager.update_size(size)
        self.config_manager.save_config(self.overlay_manager)
    
    def update_margin(self):
        self.overlay_manager.update_margin()
        self.config_manager.save_config(self.overlay_manager)
    
    def exit_app(self):
        try:
            if self.current_hotkey:
                self.gui.remove_hotkey(self.current_hotkey)
        except:
            pass
        self.system_tray.stop()
        self.overlay_manager.destroy()
        pythoncom.CoUninitialize()
        self.root.destroy()

def main():
    root = tk.Tk()
    app = MicMuteApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()