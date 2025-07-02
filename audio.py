from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
from tkinter import messagebox

class AudioManager:
    def __init__(self):
        self.volume = None
        self.last_mute_state = None
        try:
            devices = AudioUtilities.GetMicrophone()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume = cast(interface, POINTER(IAudioEndpointVolume))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize audio device: {str(e)}")
    
    def toggle_mute(self):
        if self.volume:
            try:
                current_mute = self.volume.GetMute()
                new_mute = 1 if current_mute == 0 else 0
                self.volume.SetMute(new_mute, None)
                print(f"Microphone toggled to: {'Muted' if new_mute else 'Unmuted'}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to toggle mute: {str(e)}")
    
    def poll_mute_state(self, update_status_callback):
        if self.volume:
            try:
                current_mute = self.volume.GetMute()
                if hasattr(self, 'last_mute_state') and current_mute != self.last_mute_state:
                    print(f"External mute change detected: {'Muted' if current_mute else 'Unmuted'}")
                    update_status_callback()
                self.last_mute_state = current_mute
            except Exception as e:
                print(f"Error polling mute state: {str(e)}")
        if hasattr(self, 'root'):
            self.root.after(500, lambda: self.poll_mute_state(update_status_callback))
    
    def update_status(self, gui, system_tray, overlay_manager):
        if self.volume:
            try:
                mute_state = self.volume.GetMute()
                status = "Muted" if mute_state else "Unmuted"
                gui.update_status(status)
                system_tray.update_status(mute_state)
                overlay_manager.update_status(mute_state)
            except Exception as e:
                gui.update_status("Error")
                messagebox.showerror("Error", f"Failed to get mute status: {str(e)}")