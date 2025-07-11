import tkinter as tk
from tkinter import messagebox, ttk, filedialog
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
import pygame
import sys
import winreg

class MicMuteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Microphone Mute Control")
        self.root.geometry("410x900")
        self.root.resizable(False, False)
        self.root.configure(bg="#ffffff")
        
        pythoncom.CoInitialize()
        
        self.device = None
        self.volume = None
        self.initialize_audio_device()
        
        # Configure modern ttk style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", padding=8, font=("Helvetica", 10), background="#4CAF50", foreground="white")
        style.map("TButton", background=[('active', '#45a049')])
        style.configure("TLabel", background="#ffffff", font=("Helvetica", 10))
        style.configure("TFrame", background="#ffffff")
        style.configure("TCombobox", font=("Helvetica", 10), padding=5)
        style.configure("TEntry", padding=5)
        
        # Main container
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Title
        self.title_label = ttk.Label(main_frame, text="Microphone Mute Control", font=("Helvetica", 14, "bold"))
        self.title_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        # Status
        self.label = ttk.Label(main_frame, text="Status: Unknown")
        self.label.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Toggle and Minimize buttons
        self.toggle_button = ttk.Button(main_frame, text="Toggle Mute", command=self.toggle_mute)
        self.toggle_button.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=5)
        self.refresh_button = ttk.Button(button_frame, text="Refresh Device", command=self.refresh_device)
        self.refresh_button.grid(row=0, column=0, padx=(0, 5), sticky="e")
        self.minimize_button = ttk.Button(button_frame, text="Minimize to Tray", command=self.hide_window)
        self.minimize_button.grid(row=0, column=1, padx=(5, 0), sticky="w")
        
        # Hotkey settings
        hotkey_frame = ttk.LabelFrame(main_frame, text="Hotkey Settings", padding=10)
        hotkey_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=10)
        
        self.hotkey_label = ttk.Label(hotkey_frame, text="Hotkey:")
        self.hotkey_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.current_hotkey = None
        self.label_hotkey = ttk.Label(hotkey_frame, text="ctrl+alt+m")
        self.label_hotkey.grid(row=0, column=1, sticky="w", padx=5)
        self.set_hotkey_button = ttk.Button(hotkey_frame, text="Set Hotkey", command=self.start_hotkey_capture)
        self.set_hotkey_button.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Overlay settings
        self.overlay_frame = ttk.LabelFrame(main_frame, text="Overlay Settings", padding=10)
        self.overlay_frame.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=10)
        
        self.position_label = ttk.Label(self.overlay_frame, text="Position")
        self.position_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.position_var = tk.StringVar(value="Top Mid")
        positions = ["Top Left", "Top Mid", "Top Right", "Middle Left", "Middle Right", 
                     "Bottom Left", "Bottom Mid", "Bottom Right"]
        self.position_menu = ttk.Combobox(self.overlay_frame, textvariable=self.position_var, values=positions, state="readonly", width=15)
        self.position_menu.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.position_menu.bind("<<ComboboxSelected>>", lambda e: self.update_overlay_position(self.position_var.get()))
        
        self.size_label = ttk.Label(self.overlay_frame, text="Size")
        self.size_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.size_var = tk.StringVar(value="48x48")
        sizes = ["32x32", "48x48", "64x64"]
        self.size_menu = ttk.Combobox(self.overlay_frame, textvariable=self.size_var, values=sizes, state="readonly", width=15)
        self.size_menu.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.size_menu.bind("<<ComboboxSelected>>", lambda e: self.update_overlay_size(self.size_var.get()))
        
        self.margin_label = ttk.Label(self.overlay_frame, text="Margin")
        self.margin_label.grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.margin_var = tk.StringVar(value="10")
        self.margin_entry = ttk.Entry(self.overlay_frame, textvariable=self.margin_var, width=5)
        self.margin_entry.grid(row=2, column=1, sticky="w", padx=(5, 0), pady=5)
        self.margin_button = ttk.Button(self.overlay_frame, text="Set", command=self.update_margin)
        self.margin_button.grid(row=2, column=1, sticky="w", padx=(50, 5), pady=5)
        
        self.opacity_label = ttk.Label(self.overlay_frame, text="Opacity")
        self.opacity_label.grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.opacity_var = tk.DoubleVar(value=0.7)
        self.opacity_scale = ttk.Scale(self.overlay_frame, from_=0.1, to=1.0, orient="horizontal", variable=self.opacity_var, command=self.update_opacity)
        self.opacity_scale.grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        self.opacity_value_label = ttk.Label(self.overlay_frame, text="0.7")
        self.opacity_value_label.grid(row=3, column=2, sticky="w", padx=5)
        
        # Sound settings
        self.sound_frame = ttk.LabelFrame(main_frame, text="Sound Settings", padding=10)
        self.sound_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=10)
        
        self.mute_sound_label = ttk.Label(self.sound_frame, text="Mute Sound:")
        self.mute_sound_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.mute_sound_var = tk.StringVar(value="")
        self.mute_sound_entry = ttk.Entry(self.sound_frame, textvariable=self.mute_sound_var, width=20)
        self.mute_sound_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.mute_browse_button = ttk.Button(self.sound_frame, text="Browse", command=self.browse_mute_sound)
        self.mute_browse_button.grid(row=0, column=2, padx=5, pady=5)
        
        self.unmute_sound_label = ttk.Label(self.sound_frame, text="Unmute Sound:")
        self.unmute_sound_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.unmute_sound_var = tk.StringVar(value="")
        self.unmute_sound_entry = ttk.Entry(self.sound_frame, textvariable=self.unmute_sound_var, width=20)
        self.unmute_sound_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.unmute_browse_button = ttk.Button(self.sound_frame, text="Browse", command=self.browse_unmute_sound)
        self.unmute_browse_button.grid(row=1, column=2, padx=5, pady=5)
        
        self.apply_sound_button = ttk.Button(self.sound_frame, text="Apply", command=self.apply_sounds)
        self.apply_sound_button.grid(row=2, column=1, columnspan=2, pady=10, sticky="e")
        
        # Startup settings
        self.startup_frame = ttk.LabelFrame(main_frame, text="Startup Settings", padding=10)
        self.startup_frame.grid(row=7, column=0, columnspan=2, sticky="nsew", pady=10)
        
        self.start_minimized_var = tk.BooleanVar(value=False)
        self.start_minimized_check = ttk.Checkbutton(
            self.startup_frame, 
            text="Start Minimized to Tray", 
            variable=self.start_minimized_var, 
            command=self.save_config
        )
        self.start_minimized_check.grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        
        self.start_with_windows_var = tk.BooleanVar(value=False)
        self.start_with_windows_check = ttk.Checkbutton(
            self.startup_frame, 
            text="Start with Windows", 
            variable=self.start_with_windows_var, 
            command=self.toggle_windows_startup
        )
        self.start_with_windows_check.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        
        self.load_config()
        self.is_capturing_hotkey = False
        
        self.icon = None
        self.create_tray_icon()
        
        self.overlay = None
        self.create_overlay()
        
        pygame.mixer.init()
        self.mute_sound = None
        self.unmute_sound = None
        self.last_mute_state = None
        
        try:
            self.current_hotkey = "ctrl+alt+m"
            keyboard.add_hotkey(self.current_hotkey, self.toggle_mute)
            print(f"Initial hotkey set: {self.current_hotkey}")
        except Exception as e:
            print(f"Error setting initial hotkey: {str(e)}")
        
        # Only minimize if start_minimized_var is True
        if not self.start_minimized_var.get():
            self.root.deiconify()
        else:
            self.root.withdraw()
        
        self.update_status()
        self.poll_mute_state()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def get_resource_path(self, relative_path, writable=False):
        """Get the path for a resource, handling both script and executable cases."""
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(os.path.dirname(__file__))

        if writable:
            config_dir = os.path.join(os.path.expanduser("~"), ".mic_mute_app")
            os.makedirs(config_dir, exist_ok=True)
            return os.path.join(config_dir, os.path.basename(relative_path))
        else:
            return os.path.join(base_path, relative_path)
    
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
        config_path = self.get_resource_path("config.json", writable=True)
        default_mute_sound = self.get_resource_path(os.path.join("resource", "_mute.wav"))
        default_unmute_sound = self.get_resource_path(os.path.join("resource", "_unmute.wav"))
        
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    self.position_var.set(config.get("overlay_position", "Top Mid"))
                    self.size_var.set(config.get("overlay_size", "48x48"))
                    self.margin_var.set(str(config.get("overlay_margin", 10)))
                    self.opacity_var.set(config.get("overlay_opacity", 0.7))
                    self.mute_sound_var.set(config.get("mute_sound_file", default_mute_sound))
                    self.unmute_sound_var.set(config.get("unmute_sound_file", default_unmute_sound))
                    self.start_minimized_var.set(config.get("start_minimized", False))
                    self.start_with_windows_var.set(config.get("start_with_windows", False))
                    self.opacity_value_label.config(text=f"{self.opacity_var.get():.1f}")
                    print(f"Loaded config from {config_path}: position={self.position_var.get()}, size={self.size_var.get()}, margin={self.margin_var.get()}, opacity={self.opacity_var.get()}, mute_sound={self.mute_sound_var.get()}, unmute_sound={self.unmute_sound_var.get()}, start_minimized={self.start_minimized_var.get()}, start_with_windows={self.start_with_windows_var.get()}")
            else:
                bundled_config_path = self.get_resource_path("config.json")
                if os.path.exists(bundled_config_path):
                    with open(bundled_config_path, 'r') as f:
                        config = json.load(f)
                        self.position_var.set(config.get("overlay_position", "Top Mid"))
                        self.size_var.set(config.get("overlay_size", "48x48"))
                        self.margin_var.set(str(config.get("overlay_margin", 10)))
                        self.opacity_var.set(config.get("overlay_opacity", 0.7))
                        self.mute_sound_var.set(config.get("mute_sound_file", default_mute_sound))
                        self.unmute_sound_var.set(config.get("unmute_sound_file", default_unmute_sound))
                        self.start_minimized_var.set(config.get("start_minimized", False))
                        self.start_with_windows_var.set(config.get("start_with_windows", False))
                        self.opacity_value_label.config(text=f"{self.opacity_var.get():.1f}")
                        print(f"Loaded bundled config from {bundled_config_path}: position AREA=Top Mid, size={self.size_var.get()}, margin={self.margin_var.get()}, opacity={self.opacity_var.get()}, mute_sound={self.mute_sound_var.get()}, unmute_sound={self.unmute_sound_var.get()}, start_minimized={self.start_minimized_var.get()}, start_with_windows={self.start_with_windows_var.get()}")
                else:
                    self.mute_sound_var.set(default_mute_sound)
                    self.unmute_sound_var.set(default_unmute_sound)
                    self.start_minimized_var.set(False)
                    self.start_with_windows_var.set(False)
                    self.save_config()
            self.toggle_windows_startup()
        except Exception as e:
            print(f"Error loading config: {str(e)}")
            self.mute_sound_var.set(default_mute_sound)
            self.unmute_sound_var.set(default_unmute_sound)
            self.start_minimized_var.set(False)
            self.start_with_windows_var.set(False)
            self.save_config()
    
    def save_config(self):
        config_path = self.get_resource_path("config.json", writable=True)
        try:
            config = {
                "overlay_position": self.position_var.get(),
                "overlay_size": self.size_var.get(),
                "overlay_margin": int(self.margin_var.get()),
                "overlay_opacity": float(self.opacity_var.get()),
                "mute_sound_file": self.mute_sound_var.get(),
                "unmute_sound_file": self.unmute_sound_var.get(),
                "start_minimized": self.start_minimized_var.get(),
                "start_with_windows": self.start_with_windows_var.get()
            }
            with open(config_path, 'w') as f:
                json.dump(config, f)
            print(f"Saved config to {config_path}: position={self.position_var.get()}, size={self.size_var.get()}, margin={self.margin_var.get()}, opacity={self.opacity_var.get()}, mute_sound={self.mute_sound_var.get()}, unmute_sound={self.unmute_sound_var.get()}, start_minimized={self.start_minimized_var.get()}, start_with_windows={self.start_with_windows_var.get()}")
        except Exception as e:
            print(f"Error saving config: {str(e)}")
    
    def toggle_windows_startup(self):
        try:
            key = winreg.HKEY_CURRENT_USER
            sub_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
            app_name = "MicMuteApp"
            executable_path = os.path.abspath(sys.executable if hasattr(sys, '_MEIPASS') else __file__)
            
            with winreg.OpenKey(key, sub_key, 0, winreg.KEY_SET_VALUE) as reg_key:
                if self.start_with_windows_var.get():
                    winreg.SetValueEx(reg_key, app_name, 0, winreg.REG_SZ, f'"{executable_path}"')
                    print(f"Added to Windows startup: {executable_path}")
                else:
                    try:
                        winreg.DeleteValue(reg_key, app_name)
                        print("Removed from Windows startup")
                    except FileNotFoundError:
                        print("App was not in Windows startup")
        except Exception as e:
            print(f"Error toggling Windows startup: {str(e)}")
            messagebox.showerror("Error", f"Failed to toggle Windows startup: {str(e)}")
        self.save_config()
    
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
        draw.text((10, 10), letter, fill="white", font_size=16)
        return image
    
    def create_overlay(self):
        print("Creating overlay window")
        svg_code = """
        <svg fill="red" width="64" height="64" viewBox="-0.24 0 1.52 1.52" xmlns="http://www.w3.org/2000/svg" class="cf-icon-svg"><path d="M0.933 0.633v0.105a0.421 0.421 0 0 1 -0.121 0.296 0.416 0.416 0 0 1 -0.131 0.090 0.4 0.4 0 0 1 -0.116 0.031v0.152h0.185a0.044 0.044 0 0 1 0 0.089H0.291a0.044 0.044 0 0 1 0 -0.089h0.185v-0.152a0.4 0.4 0 0 1 -0.116 -0.031 0.416 0.416 0 0 1 -0.131 -0.090 0.421 0.421 0 0 1 -0.121 -0.296v-0.105a0.044 0.044 0 1 1 0.089 0v0.105a0.33 0.33 0 0 0 0.096 0.233 0.319 0.319 0 0 0 0.458 0 0.33 0.33 0 0 0 0.096 -0.233v-0.105a0.044 0.044 0 1 1 0.089 0zM0.302 0.83a0.232 0.232 0 0 1 -0.019 -0.092V0.379A0.232 0.232 0 0 1 0.302 0.286a0.24 0.24 0 0 1 0.127 -0.127 0.232 0.232 0 0 1 0.093 -0.019 0.232 0.232 0 0 1 0.092 0.019 0.238 0.238 0 0 1 0.143 0.22l-0.001 0.359a0.237 0.237 0 0 1 -0.068 0.167 0.24 0.24 0 0 1 -0.075 0.051 0.232 0.232 0 0 1 -0.092 0.019 0.232 0.232 0 0 1 -0.093 -0.019A0.237 0.237 0 0 1 0.302 0.83"/></svg>
        """
        try:
            icon_size = int(self.size_var.get().split('x')[0])
            png_data = io.BytesIO()
            cairosvg.svg2png(bytestring=svg_code.encode('utf-8'), write_to=png_data, output_width=icon_size, output_height=icon_size, background_color='transparent')
            muted_image = Image.open(png_data)
            print(f"SVG converted to PNG, mode: {muted_image.mode}, size: {icon_size}x{icon_size}")
            if muted_image.mode != 'RGBA':
                muted_image = muted_image.convert('RGBA')
            self.muted_overlay_icon = ImageTk.PhotoImage(muted_image)
            print(f"Created {icon_size}x{icon_size} vector muted icon with transparent background")
        except Exception as e:
            print(f"Error creating vector muted icon: {str(e)}")
            self.muted_overlay_icon = self.create_overlay_icon("red", muted=True)
        
        self.overlay = tk.Toplevel(self.root)
        self.overlay.overrideredirect(True)
        self.overlay.attributes('-topmost', True)
        self.overlay.attributes('-alpha', self.opacity_var.get())
        self.overlay.attributes('-transparentcolor', '#000000')
        self.overlay.configure(bg='#000000')
        self.overlay_label = tk.Label(self.overlay, borderwidth=0, bg='#000000')
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
            self.save_config()
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
    
    def update_opacity(self, value):
        try:
            opacity = float(value)
            self.opacity_var.set(opacity)
            self.opacity_value_label.config(text=f"{opacity:.1f}")
            if self.overlay:
                self.overlay.attributes('-alpha', opacity)
                print(f"Overlay opacity set to: {opacity}")
            self.save_config()
        except Exception as e:
            print(f"Error updating overlay opacity: {str(e)}")
            messagebox.showerror("Error", f"Failed to update opacity: {str(e)}")
    
    def create_overlay_icon(self, color, muted):
        icon_size = int(self.size_var.get().split('x')[0])
        image = Image.new('RGBA', (icon_size, icon_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        scale = icon_size / 48
        draw.ellipse((6*scale, 2*scale, 18*scale, 10*scale), fill=color)
        draw.line((11*scale, 10*scale, 13*scale, 14*scale), fill=color)
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
        self.root.after(100, self.poll_mute_state)
    
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
    
    def play_sound(self, is_muted):
        sound = self.mute_sound if is_muted else self.unmute_sound
        sound_path = self.mute_sound_var.get().strip() if is_muted else self.unmute_sound_var.get().strip()

        if sound is None and sound_path and os.path.exists(sound_path):
            try:
                sound = pygame.mixer.Sound(sound_path)
                if is_muted:
                    self.mute_sound = sound
                    print(f"Mute sound loaded for playback from: {sound_path}")
                else:
                    self.unmute_sound = sound
                    print(f"Unmute sound loaded for playback from: {sound_path}")
            except Exception as e:
                print(f"Error loading {'mute' if is_muted else 'unmute'} sound for playback: {str(e)}")
                sound = None

        if sound:
            try:
                pygame.mixer.Sound.play(sound)
            except Exception as e:
                print(f"Error playing {'mute' if is_muted else 'unmute'} sound: {str(e)}")
                if is_muted:
                    self.mute_sound = None
                else:
                    self.unmute_sound = None
        else:
            print(f"No valid {'mute' if is_muted else 'unmute'} sound, skipping playback")
    
    def browse_mute_sound(self):
        file_path = filedialog.askopenfilename(filetypes=[("WAV files", "*.wav")])
        if file_path:
            self.mute_sound_var.set(file_path)
    
    def browse_unmute_sound(self):
        file_path = filedialog.askopenfilename(filetypes=[("WAV files", "*.wav")])
        if file_path:
            self.unmute_sound_var.set(file_path)
    
    def apply_sounds(self):
        self.mute_sound = None
        self.unmute_sound = None
        mute_sound_path = self.mute_sound_var.get().strip()
        unmute_sound_path = self.unmute_sound_var.get().strip()
        
        if mute_sound_path and os.path.exists(mute_sound_path):
            try:
                self.mute_sound = pygame.mixer.Sound(mute_sound_path)
                print(f"Mute sound loaded from: {mute_sound_path}")
            except Exception as e:
                print(f"Failed to load mute sound file: {str(e)}")
                self.mute_sound = None
        
        if unmute_sound_path and os.path.exists(unmute_sound_path):
            try:
                self.unmute_sound = pygame.mixer.Sound(unmute_sound_path)
                print(f"Unmute sound loaded from: {unmute_sound_path}")
            except Exception as e:
                print(f"Failed to load unmute sound file: {str(e)}")
                self.unmute_sound = None
        
        self.save_config()
    
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
                if hasattr(self, 'last_mute_state') and mute_state != self.last_mute_state:
                    self.play_sound(mute_state)
                self.last_mute_state = mute_state
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
        pygame.mixer.quit()
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
        pygame.mixer.quit()
        pythoncom.CoUninitialize()

def main():
    root = tk.Tk()
    app = MicMuteApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()