import tkinter as tk
from tkinter import messagebox, ttk
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import pythoncom
import keyboard
import pystray
from PIL import Image, ImageDraw, ImageTk
import threading
import io
import cairosvg
import json
import os

class MicMuteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Microphone Mute Control")
        self.root.geometry("244x350")  # Increased height to accommodate overlay settings
        self.root.resizable(False, True)
        self.root.configure(bg="#f0f0f0")
        
        pythoncom.CoInitialize()
        
        self.device = None
        self.volume = None
        self.initialize_audio_device()
        
        style = ttk.Style()
        style.configure("TButton", padding=5, font=("Helvetica", 10))
        style.configure("TLabel", background="#f0f0f0", font=("Helvetica", 10))
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TCombobox", font=("Helvetica", 10))
        
        self.title_label = ttk.Label(root, text="Microphone Control", font=("Helvetica", 12, "bold"))
        self.title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        self.label = ttk.Label(root, text="Status: Unknown")
        self.label.grid(row=1, column=0, columnspan=2, pady=5)
        
        self.toggle_button = ttk.Button(root, text="Toggle Mute", command=self.toggle_mute)
        self.toggle_button.grid(row=2, column=0, columnspan=2, pady=5)
        
        self.refresh_button = ttk.Button(root, text="Refresh Device", command=self.refresh_device)
        self.refresh_button.grid(row=3, column=0, padx=(10, 5), pady=5, sticky="e")
        
        self.minimize_button = ttk.Button(root, text="Minimize to Tray", command=self.hide_window)
        self.minimize_button.grid(row=3, column=1, padx=(5, 10), pady=5, sticky="w")
        
        self.hotkey_label = ttk.Label(root, text="Hotkey:")
        self.hotkey_label.grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.current_hotkey = None  # Initialize as None
        self.label_hotkey = ttk.Label(root, text="ctrl+alt+m")
        self.label_hotkey.grid(row=4, column=1, sticky="w", padx=5)
        self.set_hotkey_button = ttk.Button(root, text="Set Hotkey", command=self.start_hotkey_capture)
        self.set_hotkey_button.grid(row=5, column=0, columnspan=2, pady=5)
        
        self.overlay_frame = ttk.LabelFrame(root, text="Overlay Settings", padding=5)
        self.overlay_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)
        
        self.position_label = ttk.Label(self.overlay_frame, text="Position")
        self.position_label.grid(row=0, column=0, sticky="e", padx=5, pady=2)
        self.position_var = tk.StringVar(value="Top Left")
        positions = ["Top Left", "Top Mid", "Top Right", "Middle Left", "Middle Right", 
                     "Bottom Left", "Bottom Mid", "Bottom Right"]
        self.position_menu = ttk.Combobox(self.overlay_frame, textvariable=self.position_var, values=positions, state="readonly", width=15)
        self.position_menu.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        self.position_menu.bind("<<ComboboxSelected>>", lambda e: self.update_overlay_position(self.position_var.get()))
        
        self.size_label = ttk.Label(self.overlay_frame, text="Size")
        self.size_label.grid(row=1, column=0, sticky="e", padx=5, pady=2)
        self.size_var = tk.StringVar(value="32x32")
        sizes = ["32x32", "48x48", "64x64"]
        self.size_menu = ttk.Combobox(self.overlay_frame, textvariable=self.size_var, values=sizes, state="readonly", width=15)
        self.size_menu.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        self.size_menu.bind("<<ComboboxSelected>>", lambda e: self.update_overlay_size(self.size_var.get()))
        
        self.margin_label = ttk.Label(self.overlay_frame, text="Margin")
        self.margin_label.grid(row=2, column=0, sticky="e", padx=5, pady=2)
        self.margin_var = tk.StringVar(value="0")
        self.margin_entry = ttk.Entry(self.overlay_frame, textvariable=self.margin_var, width=5)
        self.margin_entry.grid(row=2, column=1, sticky="w", padx=(5, 0), pady=2)
        self.margin_button = ttk.Button(self.overlay_frame, text="Set", command=self.update_margin)
        self.margin_button.grid(row=2, column=1, sticky="w", padx=(50, 5), pady=2)
        
        self.load_config()
        self.is_capturing_hotkey = False
        
        self.icon = None
        self.create_tray_icon()
        
        self.overlay = None
        self.create_overlay()
        
        # Set initial hotkey after GUI setup
        try:
            self.current_hotkey = "ctrl+alt+m"
            keyboard.add_hotkey(self.current_hotkey, self.toggle_mute)
            print(f"Initial hotkey set: {self.current_hotkey}")
        except Exception as e:
            print(f"Error setting initial hotkey: {str(e)}")
        
        self.root.withdraw()
        self.update_status()
        self.poll_mute_state()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def initialize_audio_device(self):
        try:
            devices = AudioUtilities.GetMicrophone()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume = cast(interface, POINTER(IAudioEndpointVolume))
            print("Audio device initialized successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize audio device: {str(e)}")
            self.volume = None
    
    def refresh_device(self):
        try:
            self.initialize_audio_device()
            self.update_status()
            print("Microphone device refreshed")
            messagebox.showinfo("Success", "Microphone device refreshed successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh audio device: {str(e)}")
    
    def load_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "config.json")
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    self.position_var.set(config.get("overlay_position", "Top Left"))
                    self.size_var.set(config.get("overlay_size", "32x32"))
                    self.margin_var.set(str(config.get("overlay_margin", 0)))
                    print(f"Loaded config: position={self.position_var.get()}, size={self.size_var.get()}, margin={self.margin_var.get()}")
        except Exception as e:
            print(f"Error loading config: {str(e)}")
    
    def save_config(self):
        config_path = os.path.join(os.path.dirname(__file__), "config.json")
        try:
            config = {
                "overlay_position": self.position_var.get(),
                "overlay_size": self.size_var.get(),
                "overlay_margin": int(self.margin_var.get())
            }
            with open(config_path, 'w') as f:
                json.dump(config, f)
            print(f"Saved config: position={self.position_var.get()}, size={self.size_var.get()}, margin={self.margin_var.get()}")
        except Exception as e:
            print(f"Error saving config: {str(e)}")
    
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
    
    def create_icon(self, color, letter):
        image = Image.new('RGB', (32, 32), color=color)
        draw = ImageDraw.Draw(image)
        draw.text((10, 10), letter, fill="white")
        return image
    
    def create_overlay(self):
        print("Creating overlay window")
        svg_code = """
        <svg width="48" height="48" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <rect x="0" y="0" width="24" height="24" fill="none"/>
            <path d="M12,20a9,9,0,0,1-7-3.37,1,1,0,0,1,1.56-1.26,7,7,0,0,0,10.92,0A1,1,0,0,1,19,16.63,9,9,0,0,1,12,20Z" style="fill:#ff0000"/>
            <path d="M12,2A5,5,0,0,0,7,7v4a5,5,0,0,0,10,0V7A5,5,0,0,0,12,2Z" style="fill:#ff0000"/>
            <path d="M12,22a1,1,0,0,1-1-1V19a1,1,0,0,1,2,0v2A1,1,0,0,1,12,22Z" style="fill:#ff0000"/>
            <path d="M6 6 L18 18" stroke="white" stroke-width="1.5"/>
        </svg>
        """
        try:
            icon_size = int(self.size_var.get().split('x')[0])
            png_data = io.BytesIO()
            cairosvg.svg2png(bytestring=svg_code.encode('utf-8'), write_to=png_data, output_width=icon_size, output_height=icon_size)
            muted_image = Image.open(png_data)
            print(f"SVG converted to PNG, mode: {muted_image.mode}, size: {icon_size}x{icon_size}")
            if muted_image.mode != 'RGBA':
                muted_image = muted_image.convert('RGBA')
            self.muted_overlay_icon = ImageTk.PhotoImage(muted_image)
            print(f"Created {icon_size}x{icon_size} vector muted icon with new SVG")
        except Exception as e:
            print(f"Error creating vector muted icon: {str(e)}")
            self.muted_overlay_icon = self.create_overlay_icon("red", muted=True)
        
        self.overlay = tk.Toplevel(self.root)
        self.overlay.overrideredirect(True)
        self.overlay.attributes('-topmost', True)
        self.overlay.attributes('-alpha', 0.7)
        self.overlay.attributes('-transparentcolor', 'black')
        self.overlay_label = tk.Label(self.overlay, borderwidth=0, bg="black")
        self.overlay_label.pack()
        if self.volume and self.volume.GetMute():
            self.overlay_label.config(image=self.muted_overlay_icon)
            self.overlay.deiconify()
        else:
            self.overlay.withdraw()
        self.update_overlay_position(self.position_var.get())
    
    def update_overlay_position(self, position):
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            icon_size = int(self.size_var.get().split('x')[0])
            margin = int(self.margin_var.get())
            if position == "Top Left":
                x_position, y_position = margin, margin
            elif position == "Top Mid":
                x_position, y_position = (screen_width - icon_size) // 2, margin
            elif position == "Top Right":
                x_position, y_position = screen_width - icon_size - margin, margin
            elif position == "Middle Left":
                x_position, y_position = margin, (screen_height - icon_size) // 2
            elif position == "Middle Right":
                x_position, y_position = screen_width - icon_size - margin, (screen_height - icon_size) // 2
            elif position == "Bottom Left":
                x_position, y_position = margin, screen_height - icon_size - margin
            elif position == "Bottom Mid":
                x_position, y_position = (screen_width - icon_size) // 2, screen_height - icon_size - margin
            elif position == "Bottom Right":
                x_position, y_position = screen_width - icon_size - margin, screen_height - icon_size - margin
            else:
                x_position, y_position = (screen_width - icon_size) // 2, margin
            self.overlay.geometry(f"{icon_size}x{icon_size}+{x_position}+{y_position}")
            print(f"Overlay positioned at {x_position},{y_position} for screen {screen_width}x{screen_height}, position: {position}, size: {icon_size}x{icon_size}, margin: {margin}")
            self.overlay.update_idletasks()
            self.overlay.update()
            self.save_config()
        except Exception as e:
            print(f"Error setting overlay position: {str(e)}")
    
    def update_overlay_size(self, size):
        try:
            print(f"Overlay size set to: {size}")
            was_muted = self.volume.GetMute() if self.volume else False
            self.overlay.destroy()
            self.create_overlay()
            if was_muted:
                self.overlay_label.config(image=self.muted_overlay_icon)
                self.overlay.deiconify()
            self.save_config()  # Fixed typo from self.save_dist() to self.save_config()
        except Exception as e:
            print(f"Error updating overlay size: {str(e)}")
            messagebox.showerror("Error", f"Failed to update size: {str(e)}")
    
    def update_margin(self):
        try:
            margin_str = self.margin_var.get().strip()
            margin = int(margin_str) if margin_str else 0
            if 0 <= margin <= 50:
                self.margin_var.set(str(margin))
                print(f"Margin set to: {margin}")
                self.update_overlay_position(self.position_var.get())
            else:
                messagebox.showerror("Error", "Margin must be between 0 and 50")
                self.margin_var.set("0")
                self.update_overlay_position(self.position_var.get())
        except ValueError:
            messagebox.showerror("Error", "Margin must be a number")
            self.margin_var.set("0")
            self.update_overlay_position(self.position_var.get())
    
    def create_overlay_icon(self, color, muted):
        icon_size = int(self.size_var.get().split('x')[0])
        image = Image.new('RGBA', (icon_size, icon_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        scale = icon_size / 48
        draw.ellipse((6*scale, 2*scale, 18*scale, 10*scale), fill=color)
        draw.rectangle((11*scale, 10*scale, 13*scale, 14*scale), fill=color)
        draw.arc((6*scale, 8*scale, 18*scale, 14*scale), start=0, end=180, fill=color)
        if muted:
            draw.line((4*scale, 4*scale, 20*scale, 20*scale), fill="white", width=int(2*scale))
        return ImageTk.PhotoImage(image)
    
    def start_hotkey_capture(self):
        if self.is_capturing_hotkey:
            return
        self.is_capturing_hotkey = True
        self.set_hotkey_button.config(text="Press Keys...", state="disabled")
        self.root.update()
        try:
            threading.Thread(target=self.capture_hotkey, daemon=True).start()
        except Exception as e:
            self.is_capturing_hotkey = False
            self.set_hotkey_button.config(text="Set Hotkey", state="normal")
            messagebox.showerror("Error", f"Failed to start hotkey capture: {str(e)}")
    
    def capture_hotkey(self):
        try:
            new_hotkey = keyboard.read_hotkey(suppress=False)
            if not new_hotkey:
                raise ValueError("No hotkey captured")
            if self.current_hotkey:
                keyboard.remove_hotkey(self.current_hotkey)
            keyboard.add_hotkey(new_hotkey, self.toggle_mute)
            self.current_hotkey = new_hotkey
            self.label_hotkey.config(text=f"{new_hotkey}")
            print(f"Captured and set hotkey: {new_hotkey}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set hotkey: {str(e)}")
            if self.current_hotkey:
                try:
                    keyboard.add_hotkey(self.current_hotkey, self.toggle_mute)
                except:
                    pass
        finally:
            self.is_capturing_hotkey = False
            self.set_hotkey_button.config(text="Set Hotkey", state="normal")
            self.root.update()
    
    def poll_mute_state(self):
        if self.volume:
            try:
                current_mute = self.volume.GetMute()
                if hasattr(self, 'last_mute_state') and current_mute != self.last_mute_state:
                    print(f"External mute change detected: {'Muted' if current_mute else 'Unmuted'}")
                    self.update_status()
                self.last_mute_state = current_mute
            except Exception as e:
                print(f"Error polling mute state: {str(e)}")
        self.root.after(500, self.poll_mute_state)
    
    def toggle_mute(self):
        if self.volume:
            try:
                current_mute = self.volume.GetMute()
                new_mute = 1 if current_mute == 0 else 0
                self.volume.SetMute(new_mute, None)
                self.update_status()
                print(f"Microphone toggled to: {'Muted' if new_mute else 'Unmuted'}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to toggle mute: {str(e)}")
    
    def update_status(self):
        if self.volume:
            try:
                mute_state = self.volume.GetMute()
                status = "Muted" if mute_state else "Unmuted"
                self.label.config(text=f"Status: {status}")
                if self.icon:
                    self.icon.icon = self.muted_tray_icon if mute_state else self.unmuted_tray_icon
                    self.icon.title = f"Microphone: {status}"
                if self.overlay:
                    if mute_state:
                        self.overlay_label.config(image=self.muted_overlay_icon)
                        self.overlay.deiconify()
                        self.update_overlay_position(self.position_var.get())
                        print(f"Overlay shown: {status}")
                    else:
                        self.overlay.withdraw()
                        print(f"Overlay hidden: {status}")
            except Exception as e:
                self.label.config(text="Status: Error")
                messagebox.showerror("Error", f"Failed to get mute status: {str(e)}")
    
    def show_window(self):
        self.root.deiconify()
    
    def hide_window(self):
        self.root.withdraw()
    
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