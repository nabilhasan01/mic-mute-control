import tkinter as tk
from PIL import Image, ImageDraw, ImageTk
import io
import cairosvg

class OverlayManager:
    def __init__(self, root, audio_manager, config_manager):
        self.root = root
        self.audio_manager = audio_manager
        self.config_manager = config_manager
        self.overlay = None
        self.muted_overlay_icon = None
        self.overlay_label = None
        self.position_var = config_manager.position_var
        self.size_var = config_manager.size_var
        self.margin_var = config_manager.margin_var
        self.create_overlay()
    
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
            self.muted_overlay_icon = self.create_fallback_icon()
        
        self.overlay = tk.Toplevel(self.root)
        self.overlay.overrideredirect(True)
        self.overlay.attributes('-topmost', True)
        self.overlay.attributes('-alpha', 0.7)
        self.overlay.attributes('-transparentcolor', 'black')
        self.overlay_label = tk.Label(self.overlay, borderwidth=0, bg="black")
        self.overlay_label.pack()
        if self.audio_manager.volume and self.audio_manager.volume.GetMute():
            self.overlay_label.config(image=self.muted_overlay_icon)
            self.overlay.deiconify()
        else:
            self.overlay.withdraw()
        self.update_position(self.position_var.get())
    
    def create_fallback_icon(self):
        icon_size = int(self.size_var.get().split('x')[0])
        image = Image.new('RGBA', (icon_size, icon_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        scale = icon_size / 48
        draw.ellipse((6*scale, 2*scale, 18*scale, 10*scale), fill="red")
        draw.rectangle((11*scale, 10*scale, 13*scale, 14*scale), fill="red")
        draw.arc((6*scale, 8*scale, 18*scale, 14*scale), start=0, end=180, fill="red")
        draw.line((4*scale, 4*scale, 20*scale, 20*scale), fill="white", width=int(2*scale))
        return ImageTk.PhotoImage(image)
    
    def update_position(self, position):
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
        except Exception as e:
            print(f"Error setting overlay position: {str(e)}")
    
    def update_size(self, size):
        try:
            print(f"Overlay size set to: {size}")
            was_muted = self.audio_manager.volume.GetMute() if self.audio_manager.volume else False
            self.overlay.destroy()
            self.create_overlay()
            if was_muted:
                self.overlay_label.config(image=self.muted_overlay_icon)
                self.overlay.deiconify()
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
                self.update_position(self.position_var.get())
            else:
                messagebox.showerror("Error", "Margin must be between 0 and 50")
                self.margin_var.set("0")
                self.update_position(self.position_var.get())
        except ValueError:
            messagebox.showerror("Error", "Margin must be a number")
            self.margin_var.set("0")
            self.update_position(self.position_var.get())
    
    def update_status(self, mute_state):
        if mute_state:
            self.overlay_label.config(image=self.muted_overlay_icon)
            self.overlay.deiconify()
            self.update_position(self.position_var.get())
            print(f"Overlay shown: Muted")
        else:
            self.overlay.withdraw()
            print(f"Overlay hidden: Unmuted")
    
    def destroy(self):
        if self.overlay:
            self.overlay.destroy()