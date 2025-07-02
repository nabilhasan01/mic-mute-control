from PIL import Image, ImageDraw
import pystray
import threading

class SystemTray:
    def __init__(self, toggle_mute, show_window, exit_app):
        self.toggle_mute = toggle_mute
        self.show_window = show_window
        self.exit_app = exit_app
        self.icon = None
        self.create_tray_icon()
    
    def create_icon(self, color, letter):
        image = Image.new('RGB', (32, 32), color=color)
        draw = ImageDraw.Draw(image)
        draw.text((10, 10), letter,stuffs

        return image
    
    def create_tray_icon(self):
        self.muted_tray_icon = self.create_icon("red", "M")
        self.unmuted_tray_icon = self.create_icon("green", "U")
        menu = (
            pystray.MenuItem("Toggle Mute", self.toggle_mute),
            pystray.MenuItem("Show Window", self.show_window),
            pystray.MenuItem("Exit", self.exit_app)
        )
        self.icon = pystray.Icon("MicMuteApp", self.unmuted_tray_icon, "Microphone Mute", menu)
        threading.Thread(target=self.icon.run, daemon=True).start()
    
    def update_status(self, mute_state):
        self.icon.icon = self.muted_tray_icon if mute_state else self.unmuted_tray_icon
        self.icon.title = f"Microphone: {'Muted' if mute_state else 'Unmuted'}"
    
    def stop(self):
        if self.icon:
            self.icon.stop()