from PyQt6.QtCore import Qt, QTimer, QRectF, pyqtSignal, QObject, QThread, QMutex, QMutexLocker
import time
import pythoncom
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import keyboard
from keyboard import is_pressed
import pygame
import sys
import os
import json
import winreg
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QComboBox, QSlider, QLineEdit, QPushButton, QSystemTrayIcon, 
                             QMenu, QFileDialog, QMessageBox)
from PyQt6.QtGui import QIcon, QPainter, QImage, QPixmap
from PyQt6.QtSvg import QSvgRenderer

class HotkeyWorker(QObject):
    hotkey_captured = pyqtSignal(str)

    def run(self):
        try:
            new_hotkey = keyboard.read_hotkey(suppress=False)
            if new_hotkey:
                self.hotkey_captured.emit(new_hotkey)
        except Exception as e:
            print(f"[HotkeyWorker] Error: {e}")

class OverlayWidget(QWidget):
    def __init__(self, svg_code, size, opacity, position, margin, screen_size):
        super().__init__(None)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.renderer = QSvgRenderer(svg_code.encode('utf-8'))
        self.size = int(size.split('x')[0])
        self.setFixedSize(self.size, self.size)
        self.setWindowOpacity(opacity)
        self.update_position(position, margin, screen_size)

    def paintEvent(self, event):
        painter = QPainter(self)
        self.renderer.render(painter, QRectF(self.rect()))

    def update_position(self, position, margin, screen_size):
        screen_width, screen_height = screen_size
        if position == "Top Left":
            x, y = margin, margin
        elif position == "Top Mid":
            x, y = (screen_width - self.size) // 2, margin
        elif position == "Top Right":
            x, y = screen_width - self.size - margin, margin
        elif position == "Middle Left":
            x, y = margin, (screen_height - self.size) // 2
        elif position == "Middle Right":
            x, y = screen_width - self.size - margin, (screen_height - self.size) // 2
        elif position == "Bottom Left":
            x, y = margin, screen_height - self.size - margin
        elif position == "Bottom Mid":
            x, y = (screen_width - self.size) // 2, screen_height - self.size - margin
        elif position == "Bottom Right":
            x, y = screen_width - self.size - margin, screen_height - self.size - margin
        else:
            x, y = (screen_width - self.size) // 2, margin
        self.move(x, y)

class MicMuteApp(QMainWindow):
    trigger_toggle_mute = pyqtSignal()

    def __init__(self):
        super().__init__()
        icon_path = self.get_resource_path(os.path.join("resource", "icon.ico"))
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"[ERROR] Window icon file {icon_path} not found")

        self.setWindowTitle("Microphone Mute Control")
        self.setFixedWidth(400)
        self.setMinimumHeight(600)

        # Initialize COM
        pythoncom.CoInitialize()

        self.device = None
        self.volume = None
        self.initialize_audio_device()

        # Initialize mutex and toggle state
        self.mute_lock = QMutex()
        self.is_toggling = False
        self.is_toggle_pending = False

        # Initialize debounce timer
        self.debounce_timer = QTimer()
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.timeout.connect(self.process_pending_toggle)
        self.debounce_interval = 100  # ms

        # Connect signals
        self.trigger_toggle_mute.connect(self.queue_toggle, Qt.ConnectionType.QueuedConnection)
        self.setup_ui()
        self.setup_tray_icon()
        self.setup_overlay()

        pygame.mixer.init()
        self.mute_sound = None
        self.unmute_sound = None
        self.last_mute_state = None

        # Setup hotkey
        self.current_hotkey = "ctrl+alt+m"
        try:
            def check_hotkey(event):
                keys = self.current_hotkey.split('+')
                all_pressed = all(is_pressed(key.strip()) for key in keys)
                if all_pressed and not self.is_toggling:
                    self.trigger_toggle_mute.emit()
            
            self.hotkey_hook = keyboard.hook(check_hotkey, suppress=False)
            print(f"Initial hotkey hook set: {self.current_hotkey}")
        except Exception as e:
            print(f"Error setting initial hotkey hook: {str(e)}")

        self.load_config()
        self.update_status()
        self.setup_polling()

        if not self.start_minimized_check.isChecked():
            self.show()
        else:
            self.hide()

    def queue_toggle(self):
        """Queue a toggle request with debouncing."""
        if self.is_toggling or self.debounce_timer.isActive():
            print("[INFO] Toggle request ignored: already toggling or debouncing")
            return
        self.is_toggle_pending = True
        self.debounce_timer.start(self.debounce_interval)
        print("[INFO] Toggle request queued")

    def process_pending_toggle(self):
        """Process a queued toggle request."""
        if self.is_toggle_pending and not self.is_toggling:
            self.is_toggle_pending = False
            self.toggle_mute()

    def toggle_mute(self):
        """Toggle microphone mute state with robust error handling and retries."""
        if self.is_toggling:
            print("[INFO] toggle_mute skipped: already in progress")
            return

        self.is_toggling = True
        with QMutexLocker(self.mute_lock):
            try:
                if not self.volume:
                    self.initialize_audio_device()
                if self.volume:
                    # Retry COM operation up to 3 times
                    for attempt in range(3):
                        try:
                            current_mute = self.volume.GetMute()
                            new_mute = 1 if current_mute == 0 else 0
                            self.volume.SetMute(new_mute, None)
                            self.update_status()
                            print(f"[INFO] Toggled: {'Muted' if new_mute else 'Unmuted'}")
                            break
                        except Exception as e:
                            print(f"[ERROR] Toggle attempt {attempt + 1} failed: {str(e)}")
                            self.volume = None
                            self.initialize_audio_device()
                            if attempt == 2:
                                print("[ERROR] All toggle attempts failed")
                                self.status_label.setText("Status: Error (Toggle failed)")
                else:
                    print("[ERROR] No audio device available")
                    self.status_label.setText("Status: Error (No audio device)")
            except Exception as e:
                print(f"[ERROR] Toggle failed: {str(e)}")
                self.volume = None
                self.status_label.setText("Status: Error")
            finally:
                self.is_toggling = False

    def initialize_audio_device(self):
        """Initialize audio device with retry logic."""
        for attempt in range(3):
            try:
                devices = AudioUtilities.GetMicrophone()
                if not devices:
                    print("[ERROR] No microphone device found")
                    return
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                self.volume = cast(interface, POINTER(IAudioEndpointVolume))
                print("[INFO] Audio device initialized successfully")
                return
            except Exception as e:
                print(f"[ERROR] Audio device initialization attempt {attempt + 1} failed: {str(e)}")
                time.sleep(0.2)  # Brief delay before retry
        self.volume = None
        print("[ERROR] Failed to initialize audio device after retries")

    def update_status(self):
        """Update status label, tray icon, overlay, and play sound if needed."""
        if not self.volume:
            self.status_label.setText("Status: Error (No audio device)")
            if self.overlay:
                self.overlay.hide()
            return

        try:
            mute_state = self.volume.GetMute()
            status = "Muted" if mute_state else "Unmuted"
            self.status_label.setText(f"Status: {status}")
            self.tray_icon.setIcon(self.muted_tray_icon if mute_state else self.unmuted_tray_icon)
            self.tray_icon.setToolTip(f"Microphone: {status}")

            # Update overlay only if state changed
            if hasattr(self, 'last_mute_state') and mute_state != self.last_mute_state:
                if self.overlay:
                    if mute_state:
                        self.overlay.show()
                        print(f"[INFO] Overlay shown: {status}")
                    else:
                        self.overlay.hide()
                        print(f"[INFO] Overlay hidden: {status}")
                self.play_sound(mute_state)

            self.last_mute_state = mute_state
        except Exception as e:
            print(f"[ERROR] Failed to get mute status: {str(e)}")
            self.volume = None
            self.status_label.setText("Status: Error")

    def play_sound(self, is_muted):
        """Play sound for mute/unmute with robust error handling."""
        if not pygame.mixer.get_init():
            try:
                pygame.mixer.init()
                print("[INFO] Pygame mixer reinitialized")
            except Exception as e:
                print(f"[ERROR] Failed to reinitialize pygame mixer: {str(e)}")
                return

        sound = self.mute_sound if is_muted else self.unmute_sound
        sound_path = self.mute_sound_edit.text().strip() if is_muted else self.unmute_sound_edit.text().strip()

        if sound is None and sound_path and os.path.exists(sound_path):
            try:
                sound = pygame.mixer.Sound(sound_path)
                if is_muted:
                    self.mute_sound = sound
                    print(f"Mute sound loaded: {sound_path}")
                else:
                    self.unmute_sound = sound
                    print(f"Unmute sound loaded: {sound_path}")
            except Exception as e:
                print(f"[ERROR] Failed to load {'mute' if is_muted else 'unmute'} sound: {str(e)}")
                return

        if sound:
            try:
                if pygame.mixer.get_busy():
                    pygame.mixer.stop()
                pygame.mixer.Sound.play(sound)
                print(f"[INFO] Playing {'mute' if is_muted else 'unmute'} sound")
            except Exception as e:
                print(f"[ERROR] Failed to play {'mute' if is_muted else 'unmute'} sound: {str(e)}")
                if is_muted:
                    self.mute_sound = None
                else:
                    self.unmute_sound = None
        else:
            print(f"[INFO] No valid {'mute' if is_muted else 'unmute'} sound, skipping playback")

    def poll_mute_state(self):
        """Poll mute state periodically, avoiding conflicts with toggle."""
        if self.is_toggling:
            print("[INFO] Polling skipped: toggle in progress")
            self.timer.start(100)
            return

        with QMutexLocker(self.mute_lock):
            if not self.volume:
                self.initialize_audio_device()
            if self.volume:
                try:
                    mute_state = self.volume.GetMute()
                    if mute_state != self.last_mute_state:
                        print(f"[INFO] External change: {'Muted' if mute_state else 'Unmuted'}")
                        self.update_status()
                    self.last_mute_state = mute_state
                except Exception as e:
                    print(f"[ERROR] Polling failed: {str(e)}")
                    self.volume = None
            self.timer.start(100)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Title
        title_label = QLabel("Microphone Mute Control")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title_label)

        # Status
        self.status_label = QLabel("Status: Unknown")
        layout.addWidget(self.status_label)

        # Toggle and Minimize buttons
        button_layout = QHBoxLayout()
        self.toggle_button = QPushButton("Toggle Mute")
        self.toggle_button.clicked.connect(self.queue_toggle)  # Route through debouncing
        button_layout.addWidget(self.toggle_button)
        self.refresh_button = QPushButton("Refresh Device")
        self.refresh_button.clicked.connect(self.refresh_device)
        button_layout.addWidget(self.refresh_button)
        self.minimize_button = QPushButton("Minimize to Tray")
        self.minimize_button.clicked.connect(self.hide)
        button_layout.addWidget(self.minimize_button)
        layout.addLayout(button_layout)

        # Hotkey settings
        hotkey_frame = QVBoxLayout()
        hotkey_label = QLabel("Hotkey Settings")
        hotkey_label.setStyleSheet("font-weight: bold;")
        hotkey_frame.addWidget(hotkey_label)
        self.hotkey_display = QLabel("ctrl+alt+m")
        hotkey_frame.addWidget(self.hotkey_display)
        self.set_hotkey_button = QPushButton("Set Hotkey")
        self.set_hotkey_button.clicked.connect(self.start_hotkey_capture)
        hotkey_frame.addWidget(self.set_hotkey_button)
        layout.addLayout(hotkey_frame)

        # Overlay settings
        overlay_frame = QVBoxLayout()
        overlay_label = QLabel("Overlay Settings")
        overlay_label.setStyleSheet("font-weight: bold;")
        overlay_frame.addWidget(overlay_label)

        position_layout = QHBoxLayout()
        position_layout.addWidget(QLabel("Position:"))
        self.position_combo = QComboBox()
        self.position_combo.addItems(["Top Left", "Top Mid", "Top Right", "Middle Left", "Middle Right", 
                                    "Bottom Left", "Bottom Mid", "Bottom Right"])
        self.position_combo.setCurrentText("Top Mid")
        self.position_combo.currentTextChanged.connect(self.update_overlay_position)
        position_layout.addWidget(self.position_combo)
        overlay_frame.addLayout(position_layout)

        size_layout = QHBoxLayout()
        size_layout.addWidget(QLabel("Size:"))
        self.size_combo = QComboBox()
        self.size_combo.addItems(["32x32", "48x48", "64x64"])
        self.size_combo.setCurrentText("48x48")
        self.size_combo.currentTextChanged.connect(self.update_overlay_size)
        size_layout.addWidget(self.size_combo)
        overlay_frame.addLayout(size_layout)

        margin_layout = QHBoxLayout()
        margin_layout.addWidget(QLabel("Margin:"))
        self.margin_edit = QLineEdit("10")
        self.margin_edit.setFixedWidth(50)
        margin_layout.addWidget(self.margin_edit)
        margin_button = QPushButton("Set")
        margin_button.clicked.connect(self.update_margin)
        margin_layout.addWidget(margin_button)
        overlay_frame.addLayout(margin_layout)

        opacity_layout = QHBoxLayout()
        opacity_layout.addWidget(QLabel("Opacity:"))
        self.opacity_slider = QSlider(Qt.Orientation.Horizontal)
        self.opacity_slider.setMinimum(10)
        self.opacity_slider.setMaximum(100)
        self.opacity_slider.setValue(70)
        self.opacity_slider.valueChanged.connect(self.update_opacity)
        opacity_layout.addWidget(self.opacity_slider)
        self.opacity_label = QLabel("0.7")
        opacity_layout.addWidget(self.opacity_label)
        overlay_frame.addLayout(opacity_layout)

        layout.addLayout(overlay_frame)

        # Sound settings
        sound_frame = QVBoxLayout()
        sound_label = QLabel("Sound Settings")
        sound_label.setStyleSheet("font-weight: bold;")
        sound_frame.addWidget(sound_label)

        mute_sound_layout = QHBoxLayout()
        mute_sound_layout.addWidget(QLabel("Mute Sound:"))
        self.mute_sound_edit = QLineEdit()
        mute_sound_layout.addWidget(self.mute_sound_edit)
        mute_browse_button = QPushButton("Browse")
        mute_browse_button.clicked.connect(self.browse_mute_sound)
        mute_sound_layout.addWidget(mute_browse_button)
        sound_frame.addLayout(mute_sound_layout)

        unmute_sound_layout = QHBoxLayout()
        unmute_sound_layout.addWidget(QLabel("Unmute Sound:"))
        self.unmute_sound_edit = QLineEdit()
        unmute_sound_layout.addWidget(self.unmute_sound_edit)
        unmute_browse_button = QPushButton("Browse")
        unmute_browse_button.clicked.connect(self.browse_unmute_sound)
        unmute_sound_layout.addWidget(unmute_browse_button)
        sound_frame.addLayout(unmute_sound_layout)

        apply_sound_button = QPushButton("Apply")
        apply_sound_button.clicked.connect(self.apply_sounds)
        sound_frame.addWidget(apply_sound_button)
        layout.addLayout(sound_frame)

        # Startup settings
        startup_frame = QVBoxLayout()
        startup_label = QLabel("Startup Settings")
        startup_label.setStyleSheet("font-weight: bold;")
        startup_frame.addWidget(startup_label)
        self.start_minimized_check = QPushButton("Start Minimized to Tray")
        self.start_minimized_check.setCheckable(True)
        self.start_minimized_check.clicked.connect(self.save_config)
        startup_frame.addWidget(self.start_minimized_check)
        self.start_with_windows_check = QPushButton("Start with Windows")
        self.start_with_windows_check.setCheckable(True)
        self.start_with_windows_check.clicked.connect(self.toggle_windows_startup)
        startup_frame.addWidget(self.start_with_windows_check)
        layout.addLayout(startup_frame)

        layout.addStretch()

    def setup_tray_icon(self):
        self.muted_tray_icon = self.create_tray_icon("mute_icon.ico")
        self.unmuted_tray_icon = self.create_tray_icon("icon.ico")
        self.tray_icon = QSystemTrayIcon(self.unmuted_tray_icon, self)
        menu = QMenu()
        menu.addAction("Toggle Mute", self.queue_toggle)
        menu.addAction("Show Window", self.show)
        menu.addAction("Exit", self.exit_app)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    def create_tray_icon(self, icon_filename):
        icon_path = self.get_resource_path(os.path.join("resource", icon_filename))
        if os.path.exists(icon_path):
            return QIcon(icon_path)
        else:
            print(f"[ERROR] Icon file {icon_path} not found, using fallback icon")
            # Fallback to a default icon (e.g., empty QIcon or a generated one)
            image = QImage(32, 32, QImage.Format.Format_RGB32)
            image.fill(Qt.GlobalColor.gray)
            return QIcon(QPixmap.fromImage(image))

    def setup_overlay(self):
        self.svg_code = """
        <svg fill="red" width="64" height="64" viewBox="-0.24 0 1.52 1.52" xmlns="http://www.w3.org/2000/svg" class="cf-icon-svg"><path d="M0.933 0.633v0.105a0.421 0.421 0 0 1 -0.121 0.296 0.416 0.416 0 0 1 -0.131 0.090 0.4 0.4 0 0 1 -0.116 0.031v0.152h0.185a0.044 0.044 0 0 1 0 0.089H0.291a0.044 0.044 0 0 1 0 -0.089h0.185v-0.152a0.4 0.4 0 0 1 -0.116 -0.031 0.416 0.416 0 0 1 -0.131 -0.090 0.421 0.421 0 0 1 -0.121 -0.296v-0.105a0.044 0.044 0 1 1 0.089 0v0.105a0.33 0.33 0 0 0 0.096 0.233 0.319 0.319 0 0 0 0.458 0 0.33 0.33 0 0 0 0.096 -0.233v-0.105a0.044 0.044 0 1 1 0.089 0zM0.302 0.83a0.232 0.232 0 0 1 -0.019 -0.092V0.379A0.232 0.232 0 0 1 0.302 0.286a0.24 0.24 0 0 1 0.127 -0.127 0.232 0.232 0 0 1 0.093 -0.019 0.232 0.232 0 0 1 0.092 0.019 0.238 0.238 0 0 1 0.143 0.220l-0.001 0.359a0.237 0.237 0 0 1 -0.068 0.167 0.24 0.24 0 0 1 -0.075 0.051 0.232 0.232 0 0 1 -0.092 0.019 0.232 0.232 0 0 1 -0.093 -0.019A0.237 0.237 0 0 1 0.302 0.83"/>
        </svg>
        """
        self.overlay = None
        self.update_overlay()

    def update_overlay(self):
        if self.overlay:
            self.overlay.close()
            self.overlay = None
        screen = QApplication.primaryScreen()
        screen_size = (screen.size().width(), screen.size().height())
        try:
            margin = int(self.margin_edit.text() or 0)
        except ValueError:
            margin = 0
        self.overlay = OverlayWidget(
            self.svg_code,
            self.size_combo.currentText(),
            self.opacity_slider.value() / 100.0,
            self.position_combo.currentText(),
            margin,
            screen_size
        )
        if self.volume and self.volume.GetMute():
            self.overlay.show()
            print("Overlay shown: Muted")
        else:
            self.overlay.hide()
            print("Overlay hidden: Unmuted")

    def load_config(self):
        config_path = self.get_resource_path("config.json", writable=True)
        default_mute_sound = self.get_resource_path(os.path.join("resource", "_mute.wav"))
        default_unmute_sound = self.get_resource_path(os.path.join("resource", "_unmute.wav"))
        default_hotkey = "ctrl+alt+m"
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    self.position_combo.setCurrentText(config.get("overlay_position", "Top Mid"))
                    self.size_combo.setCurrentText(config.get("overlay_size", "48x48"))
                    self.margin_edit.setText(str(config.get("overlay_margin", 10)))
                    self.opacity_slider.setValue(int(config.get("overlay_opacity", 0.7) * 100))
                    self.opacity_label.setText(f"{self.opacity_slider.value() / 100:.1f}")
                    self.mute_sound_edit.setText(config.get("mute_sound_file", default_mute_sound))
                    self.unmute_sound_edit.setText(config.get("unmute_sound_file", default_unmute_sound))
                    self.start_minimized_check.setChecked(config.get("start_minimized", False))
                    self.start_with_windows_check.setChecked(config.get("start_with_windows", False))
                    loaded_hotkey = config.get("hotkey", default_hotkey)
                    if loaded_hotkey != self.current_hotkey:
                        try:
                            if hasattr(self, 'hotkey_hook'):
                                keyboard.unhook(self.hotkey_hook)
                            def check_hotkey(event):
                                keys = loaded_hotkey.split('+')
                                all_pressed = all(is_pressed(key.strip()) for key in keys)
                                if all_pressed and not self.is_toggling:
                                    self.trigger_toggle_mute.emit()
                            self.hotkey_hook = keyboard.hook(check_hotkey, suppress=False)
                            self.current_hotkey = loaded_hotkey
                            self.hotkey_display.setText(loaded_hotkey)
                            print(f"Loaded hotkey hook: {loaded_hotkey}")
                        except Exception as e:
                            print(f"Error setting loaded hotkey hook '{loaded_hotkey}': {str(e)}")
                            if hasattr(self, 'hotkey_hook'):
                                keyboard.unhook(self.hotkey_hook)
                            def check_default_hotkey(event):
                                keys = default_hotkey.split('+')
                                all_pressed = all(is_pressed(key.strip()) for key in keys)
                                if all_pressed and not self.is_toggling:
                                    self.trigger_toggle_mute.emit()
                            self.hotkey_hook = keyboard.hook(check_default_hotkey, suppress=False)
                            self.current_hotkey = default_hotkey
                            self.hotkey_display.setText(default_hotkey)
                            print(f"Fell back to default hotkey hook: {default_hotkey}")
                    print(f"Loaded config from {config_path}")
            else:
                bundled_config_path = self.get_resource_path("config.json")
                if os.path.exists(bundled_config_path):
                    with open(bundled_config_path, 'r') as f:
                        config = json.load(f)
                        self.position_combo.setCurrentText(config.get("overlay_position", "Top Mid"))
                        self.size_combo.setCurrentText(config.get("overlay_size", "48x48"))
                        self.margin_edit.setText(str(config.get("overlay_margin", 10)))
                        self.opacity_slider.setValue(int(config.get("overlay_opacity", 0.7) * 100))
                        self.opacity_label.setText(f"{self.opacity_slider.value() / 100:.1f}")
                        self.mute_sound_edit.setText(config.get("mute_sound_file", default_mute_sound))
                        self.unmute_sound_edit.setText(config.get("unmute_sound_file", default_unmute_sound))
                        self.start_minimized_check.setChecked(config.get("start_minimized", False))
                        self.start_with_windows_check.setChecked(config.get("start_with_windows", False))
                        loaded_hotkey = config.get("hotkey", default_hotkey)
                        if loaded_hotkey != self.current_hotkey:
                            try:
                                if hasattr(self, 'hotkey_hook'):
                                    keyboard.unhook(self.hotkey_hook)
                                def check_hotkey(event):
                                    keys = loaded_hotkey.split('+')
                                    all_pressed = all(is_pressed(key.strip()) for key in keys)
                                    if all_pressed and not self.is_toggling:
                                        self.trigger_toggle_mute.emit()
                                self.hotkey_hook = keyboard.hook(check_hotkey, suppress=False)
                                self.current_hotkey = loaded_hotkey
                                self.hotkey_display.setText(loaded_hotkey)
                                print(f"Loaded hotkey from bundled config: {loaded_hotkey}")
                            except Exception as e:
                                print(f"Error setting bundled hotkey hook '{loaded_hotkey}': {str(e)}")
                                if hasattr(self, 'hotkey_hook'):
                                    keyboard.unhook(self.hotkey_hook)
                                def check_default_hotkey(event):
                                    keys = default_hotkey.split('+')
                                    all_pressed = all(is_pressed(key.strip()) for key in keys)
                                    if all_pressed and not self.is_toggling:
                                        self.trigger_toggle_mute.emit()
                                self.hotkey_hook = keyboard.hook(check_default_hotkey, suppress=False)
                                self.current_hotkey = default_hotkey
                                self.hotkey_display.setText(default_hotkey)
                                print(f"Fell back to default hotkey hook: {default_hotkey}")
                    print(f"Loaded bundled config from {bundled_config_path}")
                else:
                    self.mute_sound_edit.setText(default_mute_sound)
                    self.unmute_sound_edit.setText(default_unmute_sound)
                    self.start_minimized_check.setChecked(False)
                    self.start_with_windows_check.setChecked(False)
                    self.current_hotkey = default_hotkey
                    self.hotkey_display.setText(default_hotkey)
                    try:
                        if hasattr(self, 'hotkey_hook'):
                            keyboard.unhook(self.hotkey_hook)
                        def check_default_hotkey(event):
                            keys = default_hotkey.split('+')
                            all_pressed = all(is_pressed(key.strip()) for key in keys)
                            if all_pressed and not self.is_toggling:
                                self.trigger_toggle_mute.emit()
                        self.hotkey_hook = keyboard.hook(check_default_hotkey, suppress=False)
                        print(f"Set default hotkey hook: {self.current_hotkey}")
                    except Exception as e:
                        print(f"Error setting default hotkey hook: {str(e)}")
                    self.save_config()
            self.toggle_windows_startup()
            self.update_overlay()
        except Exception as e:
            print(f"Error loading config: {str(e)}")
            self.mute_sound_edit.setText(default_mute_sound)
            self.unmute_sound_edit.setText(default_unmute_sound)
            self.start_minimized_check.setChecked(False)
            self.start_with_windows_check.setChecked(False)
            self.current_hotkey = default_hotkey
            self.hotkey_display.setText(default_hotkey)
            try:
                if hasattr(self, 'hotkey_hook'):
                    keyboard.unhook(self.hotkey_hook)
                def check_default_hotkey(event):
                    keys = default_hotkey.split('+')
                    all_pressed = all(is_pressed(key.strip()) for key in keys)
                    if all_pressed and not self.is_toggling:
                        self.trigger_toggle_mute.emit()
                self.hotkey_hook = keyboard.hook(check_default_hotkey, suppress=False)
                print(f"Set default hotkey hook after error: {self.current_hotkey}")
            except Exception as e:
                print(f"Error setting default hotkey hook after config load failure: {str(e)}")
            self.save_config()

    def save_config(self):
        config_path = self.get_resource_path("config.json", writable=True)
        try:
            config = {
                "overlay_position": self.position_combo.currentText(),
                "overlay_size": self.size_combo.currentText(),
                "overlay_margin": int(self.margin_edit.text() or 0),
                "overlay_opacity": self.opacity_slider.value() / 100.0,
                "mute_sound_file": self.mute_sound_edit.text(),
                "unmute_sound_file": self.unmute_sound_edit.text(),
                "start_minimized": self.start_minimized_check.isChecked(),
                "start_with_windows": self.start_with_windows_check.isChecked(),
                "hotkey": self.current_hotkey
            }
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=4)
            print(f"Saved config to {config_path}")
        except Exception as e:
            print(f"Error saving config: {str(e)}")

    def toggle_windows_startup(self):
        try:
            key = winreg.HKEY_CURRENT_USER
            sub_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
            app_name = "MicMuteApp"
            executable_path = os.path.abspath(sys.executable if hasattr(sys, '_MEIPASS') else __file__)
            with winreg.OpenKey(key, sub_key, 0, winreg.KEY_SET_VALUE) as reg_key:
                if self.start_with_windows_check.isChecked():
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
            QMessageBox.critical(self, "Error", f"Failed to toggle Windows startup: {str(e)}")
        self.save_config()

    def update_overlay_position(self, position):
        self.update_overlay()
        self.save_config()

    def update_overlay_size(self, size):
        self.update_overlay()
        self.save_config()

    def update_margin(self):
        try:
            margin = int(self.margin_edit.text() or 0)
            if 0 <= margin <= 50:
                self.update_overlay()
            else:
                QMessageBox.critical(self, "Error", "Margin must be between 0 and 50")
                self.margin_edit.setText("0")
                self.update_overlay()
        except ValueError:
            QMessageBox.critical(self, "Error", "Margin must be a number")
            self.margin_edit.setText("0")
            self.update_overlay()
        self.save_config()

    def update_opacity(self, value):
        try:
            opacity = value / 100.0
            self.opacity_label.setText(f"{opacity:.1f}")
            if self.overlay:
                self.overlay.setWindowOpacity(opacity)
                print(f"Overlay opacity set to: {opacity}")
            self.save_config()
        except Exception as e:
            print(f"Error updating overlay opacity: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to update opacity: {str(e)}")

    def start_hotkey_capture(self):
        if hasattr(self, 'is_capturing_hotkey') and self.is_capturing_hotkey:
            return

        self.is_capturing_hotkey = True
        self.set_hotkey_button.setText("Press Keys...")
        self.set_hotkey_button.setEnabled(False)

        self.hotkey_worker = HotkeyWorker()
        self.hotkey_thread = QThread()
        self.hotkey_worker.moveToThread(self.hotkey_thread)
        self.hotkey_worker.hotkey_captured.connect(self.apply_captured_hotkey)
        self.hotkey_thread.started.connect(self.hotkey_worker.run)
        self.hotkey_thread.start()

    def apply_captured_hotkey(self, new_hotkey):
        try:
            if hasattr(self, 'hotkey_hook'):
                keyboard.unhook(self.hotkey_hook)
            
            def check_hotkey(event):
                keys = new_hotkey.split('+')
                all_pressed = all(is_pressed(key.strip()) for key in keys)
                if all_pressed and not self.is_toggling:
                    self.trigger_toggle_mute.emit()
            
            self.hotkey_hook = keyboard.hook(check_hotkey, suppress=False)
            self.current_hotkey = new_hotkey
            self.hotkey_display.setText(new_hotkey)
            self.save_config()
            print(f"[HotkeyWorker] New hotkey hook set: {new_hotkey}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to set hotkey: {str(e)}")
        finally:
            self.is_capturing_hotkey = False
            self.set_hotkey_button.setText("Set Hotkey")
            self.set_hotkey_button.setEnabled(True)
            self.hotkey_thread.quit()
            self.hotkey_thread.wait()

    def setup_polling(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.poll_mute_state)
        self.timer.start(100)

    def refresh_device(self):
        try:
            self.initialize_audio_device()
            self.update_status()
            print("Microphone device refreshed")
            QMessageBox.information(self, "Success", "Microphone device refreshed successfully")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to refresh audio device: {str(e)}")

    def browse_mute_sound(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Mute Sound", "", "WAV files (*.wav)")
        if file_path:
            self.mute_sound_edit.setText(file_path)

    def browse_unmute_sound(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Unmute Sound", "", "WAV files (*.wav)")
        if file_path:
            self.unmute_sound_edit.setText(file_path)

    def apply_sounds(self):
        self.mute_sound = None
        self.unmute_sound = None
        mute_sound_path = self.mute_sound_edit.text().strip()
        unmute_sound_path = self.unmute_sound_edit.text().strip()
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

    def get_resource_path(self, relative_path, writable=False):
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(os.path.dirname(__file__))
        if writable:
            config_dir = os.path.join(os.path.expanduser("~"), ".mic_mute_app")
            os.makedirs(config_dir, exist_ok=True)
            return os.path.join(config_dir, os.path.basename(relative_path))
        return os.path.join(base_path, relative_path)

    def exit_app(self):
        try:
            if hasattr(self, 'hotkey_hook'):
                keyboard.unhook(self.hotkey_hook)
        except Exception as e:
            print(f"Error removing hotkey hook: {str(e)}")
        if self.overlay:
            self.overlay.close()
            self.overlay = None
        self.tray_icon.hide()
        pygame.mixer.quit()
        pythoncom.CoUninitialize()
        QApplication.quit()

    def closeEvent(self, event):
        try:
            if hasattr(self, 'hotkey_hook'):
                keyboard.unhook(self.hotkey_hook)
        except Exception as e:
            print(f"Error removing hotkey hook: {str(e)}")
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MicMuteApp()
    sys.exit(app.exec())