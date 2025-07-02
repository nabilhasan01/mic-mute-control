import tkinter as tk
from tkinter import messagebox
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import pythoncom
import keyboard
import pystray
from PIL import Image, ImageDraw, ImageTk, ImageFilter
import threading
import io
import cairosvg

class MicMuteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Microphone Mute Control")
        self.root.geometry("300x200")  # Height for GUI elements
        
        # Initialize COM for Windows API
        pythoncom.CoInitialize()
        
        # Get default audio recording device
        self.device = None
        self.volume = None
        try:
            devices = AudioUtilities.GetMicrophone()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume = cast(interface, POINTER(IAudioEndpointVolume))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize audio device: {str(e)}")
            self.root.quit()
        
        # GUI Elements
        self.label = tk.Label(root, text="Microphone Status: Unknown")
        self.label.pack(pady=10)
        
        self.toggle_button = tk.Button(root, text="Toggle Mute", command=self.toggle_mute)
        self.toggle_button.pack(pady=10)
        
        # Hotkey Setup
        self.current_hotkey = "ctrl+alt+m"  # Default hotkey
        try:
            keyboard.add_hotkey(self.current_hotkey, self.toggle_mute)
            print(f"Initial hotkey set: {self.current_hotkey}")  # Debug output
        except Exception as e:
            messagebox.showwarning("Warning", f"Failed to set up hotkey: {str(e)}")
        
        # Hotkey Modification UI
        self.label_hotkey = tk.Label(root, text=f"Current Hotkey: {self.current_hotkey}")
        self.label_hotkey.pack(pady=5)
        
        self.set_hotkey_button = tk.Button(root, text="Set Hotkey (Press Keys)", command=self.start_hotkey_capture)
        self.set_hotkey_button.pack(pady=5)
        
        self.is_capturing_hotkey = False
        
        # System Tray Setup
        self.icon = None
        self.create_tray_icon()
        
        # Overlay Setup
        self.overlay = None
        self.create_overlay()
        
        # Start minimized to tray
        self.root.withdraw()  # Hide the main window initially
        self.update_status()
        
        # Start polling for external mute changes
        self.poll_mute_state()
        
        # Ensure cleanup on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_tray_icon(self):
        # Create simple icons for muted and unmuted states
        self.muted_tray_icon = self.create_icon("red", "M")  # Red for muted
        self.unmuted_tray_icon = self.create_icon("green", "U")  # Green for unmuted
        
        # System tray menu (for right-click)
        menu = (
            pystray.MenuItem("Toggle Mute", self.toggle_mute),
            pystray.MenuItem("Show Window", self.show_window),
            pystray.MenuItem("Exit", self.exit_app)
        )
        # Set up tray icon
        self.icon = pystray.Icon("MicMuteApp", self.unmuted_tray_icon, "Microphone Mute", menu)
        
        # Run tray icon in a separate thread
        threading.Thread(target=self.icon.run, daemon=True).start()
    
    def create_icon(self, color, letter):
        # Create a simple 32x32 icon with a colored background and letter
        image = Image.new('RGB', (32, 32), color=color)
        draw = ImageDraw.Draw(image)
        draw.text((10, 10), letter, fill="white")
        return image
    
    def create_overlay(self):
        print("Creating overlay window")  # Debug output
        # Define SVG for red mic with white slash (48x48)
        svg_code = """
        <svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
            <rect x="0" y="0" width="48" height="48" fill="none"/>
            <path d="M24 8 A6 6 0 0 1 30 14 V22 A6 6 0 0 1 24 28 A6 6 0 0 1 18 22 V14 A6 6 0 0 1 24 8 M24 28 V34 M21 34 H27" fill="red"/>
            <path d="M12 12 L36 36" stroke="white" stroke-width="3"/>
        </svg>
        """
        try:
            # Convert SVG to PNG
            png_data = io.BytesIO()
            cairosvg.svg2png(bytestring=svg_code.encode('utf-8'), write_to=png_data, output_width=48, output_height=48)
            muted_image = Image.open(png_data)
            print(f"SVG converted to PNG, mode: {muted_image.mode}")  # Debug output
            if muted_image.mode != 'RGBA':
                muted_image = muted_image.convert('RGBA')  # Ensure RGBA for transparency
            self.muted_overlay_icon = ImageTk.PhotoImage(muted_image)
            print("Created 48x48 vector muted icon without shadow")  # Debug output
        except Exception as e:
            print(f"Error creating vector muted icon: {str(e)}")
            self.muted_overlay_icon = self.create_overlay_icon("red", muted=True)
        
        # Create an always-on-top, semi-transparent overlay window
        self.overlay = tk.Toplevel(self.root)
        self.overlay.overrideredirect(True)  # Remove window borders
        self.overlay.attributes('-topmost', True)  # Always on top
        self.overlay.attributes('-alpha', 0.7)  # Semi-transparent
        self.overlay.attributes('-transparentcolor', 'black')  # Set black as transparent color
        
        # Create a label to display the icon with transparent background
        self.overlay_label = tk.Label(self.overlay, borderwidth=0, bg="black")
        self.overlay_label.pack()
        
        # Position the overlay in the top-middle of the screen
        try:
            screen_width = self.root.winfo_screenwidth()
            x_position = (screen_width - 48) // 2  # Center for 48x48 icon
            self.overlay.geometry(f"48x48+{x_position}+10")
            print(f"Overlay positioned at {x_position},10 for screen width {screen_width}")  # Debug output
            self.overlay.update_idletasks()  # Force window update
            self.overlay.update()  # Force redraw
            print("Overlay updated after creation")  # Debug output
        except Exception as e:
            print(f"Error setting overlay geometry: {str(e)}")
    
    def create_overlay_icon(self, color, muted):
        # Fallback: Create a 48x48 icon with a simple microphone shape, no shadow
        image = Image.new('RGBA', (48, 48), (0, 0, 0, 0))  # Transparent background
        draw = ImageDraw.Draw(image)
        draw.rectangle((15, 8, 33, 30), fill=color)  # Mic body, scaled for 48x48
        draw.rectangle((21, 30, 27, 38), fill=color)  # Mic stand
        if muted:
            draw.line((8, 8, 40, 40), fill="white", width=3)  # Slash for muted
        return ImageTk.PhotoImage(image)
    
    def start_hotkey_capture(self):
        if self.is_capturing_hotkey:
            return  # Prevent multiple captures
        self.is_capturing_hotkey = True
        self.set_hotkey_button.config(text="Press Keys...", state="disabled")
        self.root.update()  # Update GUI to show button state
        try:
            # Run hotkey capture in a separate thread to avoid blocking GUI
            threading.Thread(target=self.capture_hotkey, daemon=True).start()
        except Exception as e:
            self.is_capturing_hotkey = False
            self.set_hotkey_button.config(text="Set Hotkey (Press Keys)", state="normal")
            messagebox.showerror("Error", f"Failed to start hotkey capture: {str(e)}")
    
    def capture_hotkey(self):
        try:
            new_hotkey = keyboard.read_hotkey(suppress=False)
            if not new_hotkey:
                raise ValueError("No hotkey captured")
            # Remove the current hotkey
            if self.current_hotkey:
                keyboard.remove_hotkey(self.current_hotkey)
            # Set the new hotkey
            keyboard.add_hotkey(new_hotkey, self.toggle_mute)
            self.current_hotkey = new_hotkey
            self.label_hotkey.config(text=f"Current Hotkey: {new_hotkey}")
            print(f"Captured and set hotkey: {new_hotkey}")  # Debug output
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set hotkey '{new_hotkey}': {str(e)}")
            # Revert to the previous hotkey if it fails
            if self.current_hotkey:
                try:
                    keyboard.add_hotkey(self.current_hotkey, self.toggle_mute)
                except:
                    pass
        finally:
            self.is_capturing_hotkey = False
            self.set_hotkey_button.config(text="Set Hotkey (Press Keys)", state="normal")
            self.root.update()  # Update GUI to reset button state
    
    def poll_mute_state(self):
        # Periodically check the microphone mute state
        if self.volume:
            try:
                current_mute = self.volume.GetMute()
                if hasattr(self, 'last_mute_state') and current_mute != self.last_mute_state:
                    print(f"External mute change detected: {'Muted' if current_mute else 'Unmuted'}")  # Debug output
                    self.update_status()
                self.last_mute_state = current_mute
            except Exception as e:
                print(f"Error polling mute state: {str(e)}")
        # Schedule the next poll (every 500ms)
        self.root.after(500, self.poll_mute_state)
    
    def toggle_mute(self):
        if self.volume:
            try:
                current_mute = self.volume.GetMute()
                new_mute = 1 if current_mute == 0 else 0
                self.volume.SetMute(new_mute, None)
                self.update_status()
                print(f"Microphone toggled to: {'Muted' if new_mute else 'Unmuted'}")  # Debug output
            except Exception as e:
                messagebox.showerror("Error", f"Failed to toggle mute: {str(e)}")
    
    def update_status(self):
        if self.volume:
            try:
                mute_state = self.volume.GetMute()
                status = "Muted" if mute_state else "Unmuted"
                self.label.config(text=f"Microphone Status: {status}")
                # Update tray icon
                if self.icon:
                    self.icon.icon = self.muted_tray_icon if mute_state else self.unmuted_tray_icon
                    self.icon.title = f"Microphone: {status}"
                # Update overlay: show only when muted
                if self.overlay:
                    if mute_state:
                        self.overlay_label.config(image=self.muted_overlay_icon)
                        self.overlay.deiconify()  # Show overlay
                        self.overlay.update_idletasks()  # Force window update
                        self.overlay.update()  # Force redraw
                        print(f"Overlay shown: {status}")  # Debug output
                    else:
                        self.overlay.withdraw()  # Hide overlay
                        print(f"Overlay hidden: {status}")  # Debug output
            except Exception as e:
                self.label.config(text="Microphone Status: Error")
                messagebox.showerror("Error", f"Failed to get mute status: {str(e)}")
    
    def show_window(self):
        self.root.deiconify()  # Show the main window
    
    def hide_window(self):
        self.root.withdraw()  # Hide the main window
    
    def on_closing(self):
        self.exit_app()
    
    def exit_app(self):
        try:
            if self.current_hotkey:
                keyboard.remove_hotkey(self.current_hotkey)
        except:
            pass
        if self.icon:
            self.icon.stop()
        if self.overlay:
            self.overlay.destroy()
        pythoncom.CoUninitialize()
        self.root.destroy()
    
    def __del__(self):
        try:
            if self.current_hotkey:
                keyboard.remove_hotkey(self.current_hotkey)
        except:
            pass
        if self.icon:
            self.icon.stop()
        if self.overlay:
            self.overlay.destroy()
        pythoncom.CoUninitialize()

def main():
    root = tk.Tk()
    app = MicMuteApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()