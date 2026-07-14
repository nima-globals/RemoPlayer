import sys
import os
import json
import urllib.request
import ctypes
import winreg
from pathlib import Path
import vlc
from win11toast import toast
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, QSlider, QFrame,
                             QGraphicsDropShadowEffect, QStackedWidget, QDialog, QTabWidget,
                             QLineEdit, QFormLayout, QFileDialog, QMessageBox, QListWidgetItem,
                             QTextEdit, QScrollArea, QCheckBox, QSystemTrayIcon, QMenu, QSizeGrip)
from PyQt6.QtCore import Qt, QTimer, QUrl, QSize, QThread, QObject, pyqtSignal
from PyQt6.QtGui import QFont, QColor, QDesktopServices, QPixmap, QPainter, QPainterPath, QIcon


def resource_path(relative):
    """Resolve bundled resource path (PyInstaller onefile extracts to _MEIPASS)."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


APP_ID = 'nima.remoplayer.1'
APP_NAME = 'RemoPlayer'


def register_app_identity(icon_path):
    """Register the AppUserModelID so Windows shows the app name and icon
    in the media flyout (SMTC) and toast notifications instead of
    'Unknown app' / 'python'."""
    try:
        key_path = rf"Software\Classes\AppUserModelId\{APP_ID}"
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
        if icon_path and os.path.exists(icon_path):
            winreg.SetValueEx(key, "IconUri", 0, winreg.REG_SZ, icon_path)
        winreg.SetValueEx(key, "IconBackgroundColor", 0, winreg.REG_SZ, "FF0F0F15")
        winreg.CloseKey(key)
    except Exception:
        pass
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass
    ensure_start_menu_shortcut(icon_path)


def ensure_start_menu_shortcut(icon_path):
    """Windows media flyout (SMTC) shows 'Unknown app' for unpackaged apps
    unless a Start Menu shortcut carries the AppUserModelID property —
    the registry key alone only covers toasts."""
    try:
        import pythoncom
        from win32com.shell import shell
        from win32com.propsys import propsys, pscon

        programs = os.path.join(os.getenv('APPDATA'),
                                'Microsoft', 'Windows', 'Start Menu', 'Programs')
        lnk_path = os.path.join(programs, f'{APP_NAME}.lnk')

        if getattr(sys, 'frozen', False):
            target, args = sys.executable, ''
        else:
            target = sys.executable.replace('python.exe', 'pythonw.exe')
            if not os.path.exists(target):
                target = sys.executable
            args = f'"{os.path.abspath(sys.argv[0])}"'

        # skip rewrite if shortcut already points at the same target
        link = pythoncom.CoCreateInstance(
            shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
            shell.IID_IShellLink)
        link.SetPath(target)
        link.SetArguments(args)
        link.SetWorkingDirectory(os.path.dirname(os.path.abspath(sys.argv[0])))
        if icon_path and os.path.exists(icon_path):
            link.SetIconLocation(icon_path, 0)

        store = link.QueryInterface(pscon.IID_IPropertyStore)
        store.SetValue(pscon.PKEY_AppUserModel_ID,
                       propsys.PROPVARIANTType(APP_ID))
        store.Commit()

        link.QueryInterface(pythoncom.IID_IPersistFile).Save(lnk_path, 0)
    except Exception:
        pass
import threading
from tinytag import TinyTag

try:
    from winrt.windows.media import (
        SystemMediaTransportControlsButton,
        MediaPlaybackStatus,
        MediaPlaybackType,
    )
    from winrt.windows.media.playback import MediaPlayer as WinRtMediaPlayer
    from winrt.windows.storage.streams import RandomAccessStreamReference
    from winrt.windows.foundation import Uri
    SMTC_AVAILABLE = True
except ImportError:
    SMTC_AVAILABLE = False


class PlaybackState:
    PLAYING = "playing"
    PAUSED = "paused"
    STOPPED = "stopped"


class SmtcManager(QObject):
    """Windows System Media Transport Controls (media flyout above notification center).

    Uses a hidden WinRT MediaPlayer to own an SMTC session, since VLC does not
    integrate with Windows media controls on its own. Button presses arrive on
    a WinRT thread, so they are forwarded to the UI thread via Qt signals.
    """
    play_pressed = pyqtSignal()
    pause_pressed = pyqtSignal()
    stop_pressed = pyqtSignal()
    next_pressed = pyqtSignal()
    prev_pressed = pyqtSignal()

    _STATUS_MAP = None

    def __init__(self, parent=None):
        super().__init__(parent)
        self._mp = WinRtMediaPlayer()
        self._mp.command_manager.is_enabled = False
        self._smtc = self._mp.system_media_transport_controls
        self._smtc.is_enabled = True
        self._smtc.is_play_enabled = True
        self._smtc.is_pause_enabled = True
        self._smtc.is_stop_enabled = True
        self._smtc.is_next_enabled = False
        self._smtc.is_previous_enabled = False
        self._smtc.add_button_pressed(self._on_button_pressed)
        SmtcManager._STATUS_MAP = {
            PlaybackState.PLAYING: MediaPlaybackStatus.PLAYING,
            PlaybackState.PAUSED: MediaPlaybackStatus.PAUSED,
            PlaybackState.STOPPED: MediaPlaybackStatus.STOPPED,
        }

    def _on_button_pressed(self, sender, args):
        button = args.button
        if button == SystemMediaTransportControlsButton.PLAY:
            self.play_pressed.emit()
        elif button == SystemMediaTransportControlsButton.PAUSE:
            self.pause_pressed.emit()
        elif button == SystemMediaTransportControlsButton.STOP:
            self.stop_pressed.emit()
        elif button == SystemMediaTransportControlsButton.NEXT:
            self.next_pressed.emit()
        elif button == SystemMediaTransportControlsButton.PREVIOUS:
            self.prev_pressed.emit()

    def set_nav_enabled(self, enabled):
        self._smtc.is_next_enabled = enabled
        self._smtc.is_previous_enabled = enabled

    def update(self, state, title, artist, album_artist="", thumbnail_path=None):
        self._smtc.playback_status = self._STATUS_MAP[state]
        updater = self._smtc.display_updater
        updater.type = MediaPlaybackType.MUSIC
        updater.music_properties.title = title or ""
        updater.music_properties.artist = artist or ""
        updater.music_properties.album_artist = album_artist or ""
        if thumbnail_path and os.path.exists(thumbnail_path):
            uri = Uri(Path(thumbnail_path).as_uri())
            updater.thumbnail = RandomAccessStreamReference.create_from_uri(uri)
        else:
            updater.thumbnail = None
        updater.update()

class MusicMetadataLoader(QThread):
    metadata_loaded = pyqtSignal(int, dict)  
    finished_all = pyqtSignal()
    
    def __init__(self, music_files, vlc_instance):
        super().__init__()
        self.music_files = music_files
        self.vlc_instance = vlc_instance
        self.should_stop = False
        
    def run(self):
        for idx, filepath in enumerate(self.music_files):
            if self.should_stop:
                break
            
            filename = os.path.splitext(os.path.basename(filepath))[0]
            
            try:
                tag = TinyTag.get(filepath, image=True)
                metadata = {
                    'title': tag.title if tag.title else filename,
                    'artist': tag.artist if tag.artist else "",
                    'album': tag.album if tag.album else "",
                    'duration': int(tag.duration * 1000) if tag.duration else 0,
                    'image': tag.get_image()
                }
            except Exception:
                metadata = {
                    'title': filename,
                    'artist': '',
                    'album': '',
                    'duration': 0,
                    'image': None
                }
                
            self.metadata_loaded.emit(idx, metadata)
            self.msleep(5)
                
        self.finished_all.emit()
    
    def stop(self):
        self.should_stop = True

class ConfigManager:
    def __init__(self):
        self.appdata_dir = os.path.join(os.getenv('APPDATA'), 'RemoPlayer')
        os.makedirs(self.appdata_dir, exist_ok=True)
        self.config_file = os.path.join(self.appdata_dir, 'config.json')
        self.default_music_dir = os.path.join(os.path.expanduser('~'), 'Music')
        base = "http://212.80.8.200/listen"
        self.default_stations = {
            "🎤 رپ‌فا": {"url": f"{base}/rapfa/radio.mp3", "desc": "رپ و هیپ‌هاپ فارسی"},
            "🔥 پاپ‌فا": {"url": f"{base}/popularfa/radio.mp3", "desc": "هیت‌های پاپ فارسی"},
            "🎧 ریمیکس‌فا": {"url": f"{base}/remixfa/radio.mp3", "desc": "ریمیکس‌های داغ فارسی"},
            "😢 سَد": {"url": f"{base}/sad/radio.mp3", "desc": "وقتی ناراحتی"},
            "🕯 خاص": {"url": f"{base}/khaaz/radio.mp3", "desc": "خاطره‌انگیز"},
            "🎻 کلاسیک": {"url": f"{base}/classic/radio.mp3", "desc": "قدیمی ولی محبوب"},
            "💔 دپ": {"url": f"{base}/dep/radio.mp3", "desc": "برای دلی که شکسته"},
            "😈 فونک": {"url": f"{base}/phonk/radio.mp3", "desc": "امروزی‌پسند و هایپ"},
            "☕ لوفای": {"url": f"{base}/lofi/radio.mp3", "desc": "چیل و آروم"},
            "🤘 متال": {"url": f"{base}/metal/radio.mp3", "desc": "سنگین و پرانرژی"},
            "🎤 هیپ‌هاپ": {"url": f"{base}/hiphop/radio.mp3", "desc": "بیت و فلوی بین‌المللی"},
            "🌟 پاپ": {"url": f"{base}/pop/radio.mp3", "desc": "پاپ روز دنیا"},
            "🪆 روسی": {"url": f"{base}/russian/radio.mp3", "desc": "حال‌وهوای موزیک روسی"},
        }
        self.default_config = {
            "stations": self.default_stations,
            "music_folder": self.default_music_dir,
            "proxy_enabled": False,
            "custom_proxy": "",
            "notifications_enabled": True,
            "run_on_startup": False,
            "close_to_tray": False,
            "stations_version": 2
        }
        self.config = self.load_config()

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                return self.migrate_config(cfg)
            except Exception:
                return json.loads(json.dumps(self.default_config))
        return json.loads(json.dumps(self.default_config))

    def migrate_config(self, cfg):
        # old format: station value was a plain URL string
        stations = cfg.get("stations", {})
        for name, val in list(stations.items()):
            if isinstance(val, str):
                stations[name] = {"url": val, "desc": "ایستگاه رادیویی آنلاین"}
        # merge new default stations once (bump stations_version to re-merge)
        if cfg.get("stations_version", 0) < 2:
            existing_urls = {v["url"] for v in stations.values()}
            for name, val in self.default_stations.items():
                if val["url"] not in existing_urls:
                    stations[name] = dict(val)
            cfg["stations_version"] = 2
        cfg["stations"] = stations
        for key, val in self.default_config.items():
            cfg.setdefault(key, val)
        return cfg

    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

class MediaItemWidget(QWidget):
    def __init__(self, title, subtitle, extra_info, icon_text="🎵", image_data=None, rtl=True):
        super().__init__()
        self.setStyleSheet("background: transparent; border: none;")
        # AlignAbsolute: Qt mirrors AlignRight to visual left inside RTL widgets
        self._align = ((Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignAbsolute) if rtl
                       else Qt.AlignmentFlag.AlignLeft) | Qt.AlignmentFlag.AlignVCenter
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft if rtl else Qt.LayoutDirection.LeftToRight)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(12)
        
        self.img_label = QLabel()
        self.img_label.setFixedSize(48, 48)
        self.img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        if image_data:
            pix = QPixmap()
            pix.loadFromData(image_data)
            pix = pix.scaled(48, 48, Qt.AspectRatioMode.KeepAspectRatioByExpanding, Qt.TransformationMode.SmoothTransformation)
            self.img_label.setPixmap(pix)
            self.img_label.setStyleSheet("border-radius: 8px; background: #111118;")
        else:
            self.img_label.setText(icon_text)
            self.img_label.setStyleSheet("font-size: 20px; color: #00f0ff; background: #111118; border-radius: 8px;")
        
        text_layout = QVBoxLayout()
        text_layout.setSpacing(4)
        text_layout.setContentsMargins(0, 0, 0, 0)
        
        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet("color: #ffffff; font-weight: bold; font-size: 13px; background: transparent;")
        self.title_lbl.setAlignment(self._align)

        # consistent order: time/extra first, then artist/album info
        sub_text = f"⏱ {extra_info}  •  {subtitle}" if extra_info else subtitle
        self.sub_lbl = QLabel(sub_text)
        self.sub_lbl.setStyleSheet("color: #8b8b9c; font-size: 11px; background: transparent;")
        self.sub_lbl.setAlignment(self._align)

        text_layout.addWidget(self.title_lbl)
        text_layout.addWidget(self.sub_lbl)
        
        layout.addWidget(self.img_label)
        layout.addLayout(text_layout, stretch=1)

class SettingsDialog(QDialog):
    def __init__(self, parent=None, config_manager=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.setWindowTitle("⚙️ تنظیمات RemoPlayer")
        self.resize(760, 640)
        self.setMinimumSize(700, 560)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        
        self.setStyleSheet("""
            QDialog { 
                background: #0b0b10;
                color: #ffffff; 
                font-family: 'Segoe UI', Tahoma, sans-serif;
            }
            QTabWidget::pane { 
                border: 1px solid #1f1f2e; 
                border-radius: 12px; 
                background: #12121a;
                padding: 15px;
            }
            QTabBar::tab { 
                background: #161622;
                color: #8b8b9c; 
                padding: 10px 20px; 
                border-top-left-radius: 8px; 
                border-top-right-radius: 8px;
                margin-right: 4px;
                font-weight: bold;
                font-size: 13px;
                border: 1px solid #1f1f2e;
                border-bottom: none;
            }
            QTabBar::tab:selected { 
                background: #12121a;
                color: #00f0ff;
                border-bottom: 2px solid #00f0ff;
                font-size: 13px;
            }
            QTabBar::tab:hover:!selected { 
                background: #1f1f30; 
                color: #ffffff;
            }
            QLineEdit { 
                background: #1c1c28;
                border: 1px solid #2a2a3e; 
                color: white; 
                padding: 10px 14px; 
                border-radius: 8px;
                font-size: 13px;
                selection-background-color: #ff007f;
            }
            QLineEdit:focus { 
                border: 1px solid #00f0ff; 
                background: #222232;
            }
            QPushButton { 
                background: #1c1c28;
                color: white; 
                padding: 10px 18px; 
                border-radius: 8px; 
                border: 1px solid #2a2a3e;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover { 
                background: #252538;
                border: 1px solid #3a3a56;
            }
            QPushButton:pressed { 
                background: #12121a;
            }
            QListWidget { 
                background: #161622;
                color: white; 
                border: 1px solid #1f1f2e; 
                border-radius: 10px;
                padding: 5px;
                font-size: 13px;
            }
            QListWidget::item {
                padding: 10px;
                border-radius: 6px;
                margin: 2px 0px;
                background: #1c1c28;
            }
            QListWidget::item:hover {
                background: #252538;
            }
            QListWidget::item:selected {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff007f, stop:1 #7000ff);
                color: white;
            }
            QScrollBar:vertical {
                border: none;
                background: #12121a;
                width: 8px;
                margin: 0px 0px 0px 0px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #2a2a3e;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #00f0ff;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QLabel { 
                color: #ffffff; 
                font-size: 13px;
            }
            QCheckBox {
                color: #ffffff;
                font-size: 13px;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 1px solid #2a2a3e;
                background: #1c1c28;
            }
            QCheckBox::indicator:checked {
                background: #00f0ff;
                border: 1px solid #00f0ff;
            }
        """)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        header_layout = QHBoxLayout()
        header_title = QLabel("⚙️ مدیریت تنظیمات سیستم")
        header_title.setStyleSheet("font-size: 18px; font-weight: bold; color: #00f0ff;")
        header_layout.addWidget(header_title)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)
        
        self.tabs = QTabWidget()
        
        self.radio_tab = QWidget()
        self.player_tab = QWidget()
        self.proxy_tab = QWidget()
        self.about_tab = QWidget()
        
        self.tabs.addTab(self.radio_tab, "📻 ایستگاه‌های رادیو")
        self.tabs.addTab(self.player_tab, "🎵 موزیک پلیر و سیستم")
        self.tabs.addTab(self.proxy_tab, "🌐 تنظیمات پروکسی")
        self.tabs.addTab(self.about_tab, "ℹ️ درباره برنامه")
        
        self.setup_radio_tab()
        self.setup_player_tab()
        self.setup_proxy_tab()
        self.setup_about_tab()
        
        main_layout.addWidget(self.tabs)
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        
        cancel_btn = QPushButton("❌ انصراف")
        cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel_btn.setStyleSheet("""
            QPushButton {
                background: #251820;
                border: 1px solid #5c1e30;
                color: #ff5580;
                padding: 10px 22px;
            }
            QPushButton:hover {
                background: #ff2a5f;
                color: white;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        
        save_btn = QPushButton("💾 ذخیره تغییرات")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.setStyleSheet("""
            QPushButton {
                background: #122528;
                border: 1px solid #1e5c5a;
                color: #00f0ff;
                padding: 10px 22px;
            }
            QPushButton:hover {
                background: #00f0ff;
                color: black;
            }
        """)
        save_btn.clicked.connect(self.save_and_close)
        
        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        main_layout.addLayout(btn_layout)

    def setup_radio_tab(self):
        layout = QVBoxLayout(self.radio_tab)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(12)
        
        self.radio_list_widget = QListWidget()
        self.refresh_radio_list()
        layout.addWidget(self.radio_list_widget, stretch=1)
        
        action_section = QFrame()
        action_section.setStyleSheet("""
            QFrame {
                background: #161622;
                border: 1px solid #2a2a3e;
                border-radius: 10px;
                padding: 12px;
            }
            QLabel { color: #ff007f; font-weight: bold; }
        """)
        action_layout = QVBoxLayout(action_section)
        action_layout.setSpacing(10)
        
        action_title = QLabel("➕ افزودن ایستگاه جدید یا مدیریت:")
        action_layout.addWidget(action_title)
        
        form_layout = QHBoxLayout()
        form_layout.setSpacing(8)

        self.r_name = QLineEdit()
        self.r_name.setPlaceholderText("نام رادیو (مثال: رپ‌فا)")

        self.r_url = QLineEdit()
        self.r_url.setPlaceholderText("لینک استریم (http://...)")

        self.r_desc = QLineEdit()
        self.r_desc.setPlaceholderText("توضیح کوتاه (مثال: چیل و آروم)")

        add_btn = QPushButton("افزودن ➕")
        add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff007f, stop:1 #7000ff);
                border: none;
                color: white;
                padding: 10px 15px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff3399, stop:1 #8833ff);
            }
        """)
        add_btn.clicked.connect(self.add_radio)
        
        form_layout.addWidget(self.r_name, stretch=2)
        form_layout.addWidget(self.r_url, stretch=4)
        form_layout.addWidget(add_btn, stretch=1)
        action_layout.addLayout(form_layout)
        action_layout.addWidget(self.r_desc)
        
        del_btn = QPushButton("🗑️ حذف ایستگاه انتخاب شده از لیست فوق")
        del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        del_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: 1px dashed #5c1e30;
                color: #ff5580;
                font-size: 12px;
            }
            QPushButton:hover {
                background: #5c1e30;
                color: white;
                border-style: solid;
            }
        """)
        del_btn.clicked.connect(self.del_radio)
        action_layout.addWidget(del_btn)
        
        layout.addWidget(action_section)

    def refresh_radio_list(self):
        self.radio_list_widget.clear()
        for name, station in self.config_manager.config["stations"].items():
            desc = station.get("desc", "")
            self.radio_list_widget.addItem(f"{name}  —  {desc}  |  {station['url']}")

    def add_radio(self):
        name = self.r_name.text().strip()
        url = self.r_url.text().strip()
        desc = self.r_desc.text().strip()
        if name and url:
            self.config_manager.config["stations"][name] = {
                "url": url, "desc": desc or "ایستگاه رادیویی آنلاین"}
            self.refresh_radio_list()
            self.r_name.clear()
            self.r_url.clear()
            self.r_desc.clear()

    def del_radio(self):
        selected = self.radio_list_widget.currentItem()
        if selected:
            name = selected.text().split("  —  ")[0]
            if name in self.config_manager.config["stations"]:
                del self.config_manager.config["stations"][name]
                self.refresh_radio_list()

    def setup_player_tab(self):
        layout = QVBoxLayout(self.player_tab)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(15)
        
        folder_card = QFrame()
        folder_card.setStyleSheet("""
            QFrame {
                background: #161622;
                border: 1px solid #2a2a3e;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        folder_layout = QVBoxLayout(folder_card)
        folder_layout.setSpacing(10)
        
        folder_title = QLabel("📁 مسیر پوشه پیش‌فرض موسیقی:")
        folder_title.setStyleSheet("font-weight: bold; color: #00f0ff;")
        
        path_layout = QHBoxLayout()
        path_layout.setSpacing(8)
        
        self.folder_label = QLabel(self.config_manager.config['music_folder'])
        self.folder_label.setWordWrap(True)
        self.folder_label.setStyleSheet("""
            background: #0b0b10; 
            padding: 12px; 
            border-radius: 6px; 
            color: #d0d0d5;
            font-family: 'Consolas', monospace;
            border: 1px solid #1f1f2e;
            font-size: 12px;
        """)
        
        change_btn = QPushButton("🔍 تغییر پوشه")
        change_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        change_btn.clicked.connect(self.change_folder)
        
        path_layout.addWidget(self.folder_label, stretch=4)
        path_layout.addWidget(change_btn, stretch=1)
        
        folder_layout.addWidget(folder_title)
        folder_layout.addLayout(path_layout)
        layout.addWidget(folder_card)
        
        sys_card = QFrame()
        sys_card.setStyleSheet("""
            QFrame {
                background: #161622;
                border: 1px solid #2a2a3e;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        sys_layout = QVBoxLayout(sys_card)
        sys_layout.setSpacing(12)
        
        sys_title = QLabel("⚙️ تنظیمات رفتار سیستم:")
        sys_title.setStyleSheet("font-weight: bold; color: #ff007f;")
        sys_layout.addWidget(sys_title)
        
        self.notif_checkbox = QCheckBox("🔔 نمایش اعلان ویندوز هنگام تغییر آهنگ رادیو")
        self.notif_checkbox.setChecked(self.config_manager.config.get("notifications_enabled", True))
        sys_layout.addWidget(self.notif_checkbox)
        
        self.startup_checkbox = QCheckBox("🚀 اجرای خودکار برنامه هنگام روشن شدن ویندوز (Startup)")
        self.startup_checkbox.setChecked(self.config_manager.config.get("run_on_startup", False))
        sys_layout.addWidget(self.startup_checkbox)

        self.tray_checkbox = QCheckBox("📥 با بستن پنجره، برنامه به سینی سیستم (Tray) برود و پخش ادامه یابد")
        self.tray_checkbox.setChecked(self.config_manager.config.get("close_to_tray", False))
        sys_layout.addWidget(self.tray_checkbox)
        
        layout.addWidget(sys_card)
        layout.addStretch()

    def change_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "انتخاب پوشه موزیک")
        if folder:
            self.config_manager.config['music_folder'] = folder
            self.folder_label.setText(folder)

    def setup_proxy_tab(self):
        layout = QVBoxLayout(self.proxy_tab)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(15)
        
        proxy_card = QFrame()
        proxy_card.setStyleSheet("""
            QFrame {
                background: #161622;
                border: 1px solid #2a2a3e;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        proxy_layout = QVBoxLayout(proxy_card)
        proxy_layout.setSpacing(10)
        
        proxy_title = QLabel("🌐 تنظیم و آدرس‌دهی پروکسی (Proxy Server):")
        proxy_title.setStyleSheet("font-weight: bold; color: #00f0ff;")
        
        self.proxy_input = QLineEdit()
        self.proxy_input.setText(self.config_manager.config.get("custom_proxy", ""))
        self.proxy_input.setPlaceholderText("مانند: http://127.0.0.1:8080")
        
        info_hint = QLabel("💡 نکته: جهت غیرفعال‌سازی و استفاده از اینترنت مستقیم، کادر بالا را کاملاً خالی بگذارید.")
        info_hint.setStyleSheet("color: #8b8b9c; font-size: 12px;")
        
        proxy_layout.addWidget(proxy_title)
        proxy_layout.addWidget(self.proxy_input)
        proxy_layout.addWidget(info_hint)
        layout.addWidget(proxy_card)
        
        guide_card = QFrame()
        guide_card.setStyleSheet("""
            QFrame {
                background: #0e0e16;
                border-radius: 8px;
                padding: 12px;
                border-left: 4px solid #7000ff;
            }
            QLabel { color: #8b8b9c; font-size: 12px; line-height: 1.5; }
        """)
        guide_layout = QVBoxLayout(guide_card)
        guide_text = QLabel(
            "📚 پروتکل‌های مجاز ساختار آدرس‌دهی:\n"
            "• پروتکل HTTP معمولی:  http://ip:port\n"
            "• پروتکل ساکس ۵:  socks5://ip:port\n"
            "• پروکسی‌های دارای یوزرنیم:  http://user:pass@ip:port"
        )
        guide_layout.addWidget(guide_text)
        layout.addWidget(guide_card)
        
        layout.addStretch()

    def setup_about_tab(self):
        layout = QVBoxLayout(self.about_tab)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(15)
        
        banner = QFrame()
        banner.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1a0011, stop:0.5 #0c001a, stop:1 #001a1a);
                border: 1px solid #2a2a3e;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        banner_layout = QVBoxLayout(banner)
        banner_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        title_lbl = QLabel("✨ RemoPlayer")
        title_lbl.setStyleSheet("font-size: 24px; font-weight: bold; color: #ffffff; padding: 4px; background: transparent; border: none;")
        title_lbl.setMinimumHeight(48)
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        version_lbl = QLabel("نسخه ۱.۰")
        version_lbl.setStyleSheet("font-size: 13px; color: #00f0ff; font-weight: bold; padding: 4px; background: transparent; border: none;")
        version_lbl.setMinimumHeight(30)
        version_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        banner_layout.setSpacing(6)
        banner_layout.addWidget(title_lbl)
        banner_layout.addWidget(version_lbl)
        layout.addWidget(banner)
        
        desc_lbl = QLabel(
            "پلیر نیتیو فوق مدرن و سبک ویندوز توسعه داده شده با فریمورک قدرتمند PyQt6 و هسته LibVLC.\n"
            "دارای قابلیت مدیریت داینامیک ایستگاه‌های رادیویی جهانی، واکشی متادیتا آهنگ آنلاین،\n"
            "یکپارچگی عمیق با سیستم کنترل ویندوز (SMTC) و لایه پروکسی اختصاصی داخلی."
        )
        desc_lbl.setWordWrap(True)
        desc_lbl.setStyleSheet("color: #b5b5c0; font-size: 13px; line-height: 1.6; padding: 5px;")
        layout.addWidget(desc_lbl)
        
        links_layout = QHBoxLayout()
        links_layout.setSpacing(10)
        
        gh_btn = QPushButton("🐱 گیت‌هاب پروژه")
        gh_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        gh_btn.setStyleSheet("QPushButton { background: #1e1e2a; border-color: #33334d; } QPushButton:hover { background: #2b2b3d; }")
        gh_btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/nima-globals")))
        
        site_btn = QPushButton("🌐 وبسایت نیما ابراهیمی")
        site_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        site_btn.setStyleSheet("QPushButton { background: #1e1e2a; border-color: #33334d; } QPushButton:hover { background: #2b2b3d; }")
        site_btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://nimaebrahimi.ir")))
        
        links_layout.addWidget(gh_btn)
        links_layout.addWidget(site_btn)
        layout.addLayout(links_layout)
        
        layout.addStretch()
        
        copyright_lbl = QLabel("© 2024-2026 Nima Ebrahimi. All rights reserved.")
        copyright_lbl.setStyleSheet("color: #4a4a5a; font-size: 11px;")
        layout.addWidget(copyright_lbl, alignment=Qt.AlignmentFlag.AlignCenter)

    def save_and_close(self):
        self.config_manager.config["custom_proxy"] = self.proxy_input.text().strip()
        self.config_manager.config["notifications_enabled"] = self.notif_checkbox.isChecked()
        self.config_manager.config["close_to_tray"] = self.tray_checkbox.isChecked()
        
        old_startup = self.config_manager.config.get("run_on_startup", False)
        new_startup = self.startup_checkbox.isChecked()
        self.config_manager.config["run_on_startup"] = new_startup
        
        if old_startup != new_startup:
            RemoPlayer.update_windows_startup(new_startup)
            
        self.config_manager.save_config()
        self.accept()

class RemoPlayer(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("RemoPlayer")
        self.resize(1000, 730)
        self.setMinimumSize(900, 650)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self._drag_pos = None
        
        self.config_mgr = ConfigManager()
        self.apply_proxy_settings()
        
        self.current_mode = "radio"
        self.session_seconds = 0
        self.total_seconds = 0
        self.bitrate_kbps = 64
        self.last_played_title = ""
        self.current_media_name = "چیزی انتخاب نشده"
        self.music_files = []
        self.is_playing = False
        self.visualizer_frame = 0
        self.metadata_loader = None
        self.music_metadata_cache = {}
        self.history_list = []
        self.current_music_url = ""
        self.current_art_data = None
        self.last_recorded_finish = ""
        self.playing_source = None      # 'radio' | 'music' | None — what is actually playing
        self.current_music_index = -1
        self.shuffle_enabled = False
        self.repeat_mode = "off"        # 'off' | 'all' | 'one'
        self.music_loaded = False       # music library scanned once, cached after
        self.force_quit = False
        self.play_order_history = []    # for prev while shuffling
        self.tray = None
        
        self.instance = vlc.Instance("--network-caching=400 --no-video --avcodec-fast") 
        self.player = self.instance.media_player_new()
        self.player.video_set_callbacks(None, None, None, None)

        self.smtc = None
        if SMTC_AVAILABLE:
            try:
                self.smtc = SmtcManager(self)
                self.smtc.play_pressed.connect(self.play_media)
                self.smtc.pause_pressed.connect(self.pause_media)
                self.smtc.stop_pressed.connect(self.stop_media)
                self.smtc.next_pressed.connect(self.next_track)
                self.smtc.prev_pressed.connect(self.prev_track)
                self.smtc.update(PlaybackState.STOPPED, "RemoPlayer", "آماده به پخش")
            except Exception:
                self.smtc = None

        self.meta_timer = QTimer(self)
        self.meta_timer.timeout.connect(self.update_metadata)
        self.stats_timer = QTimer(self)
        self.stats_timer.timeout.connect(self.update_stats)
        
        self.progress_timer = QTimer(self)
        self.progress_timer.timeout.connect(self.update_music_progress)
        
        self.visualizer_timer = QTimer(self)
        self.visualizer_timer.timeout.connect(self.update_visualizer)
        self.visualizer_timer.start(300)

        self.setup_ui()
        self.apply_styles()
        self.load_radio_list()
        # preload music library in the background so the music tab opens instantly
        QTimer.singleShot(600, lambda: self.load_music_list() if not self.music_loaded else None)

    @staticmethod
    def update_windows_startup(enabled):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            if enabled:
                cmd = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
                winreg.SetValueEx(key, "RemoPlayer", 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, "RemoPlayer")
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
        except Exception:
            pass

    def apply_proxy_settings(self):
        custom_proxy = self.config_mgr.config.get("custom_proxy", "")
        if custom_proxy:
            os.environ['http_proxy'] = custom_proxy
            os.environ['https_proxy'] = custom_proxy
        else:
            os.environ['http_proxy'] = ''
            os.environ['https_proxy'] = ''
            os.environ['ALL_PROXY'] = ''

    def create_glow_effect(self, color_hex, blur_radius, offset_y=0):
        effect = QGraphicsDropShadowEffect(self)
        effect.setBlurRadius(blur_radius)
        effect.setColor(QColor(color_hex))
        effect.setOffset(0, offset_y)
        return effect

    def create_circular_pixmap(self, image_data, size=160):
        pixmap = QPixmap()
        pixmap.loadFromData(image_data)
        pixmap = pixmap.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatioByExpanding, Qt.TransformationMode.SmoothTransformation)
        target = QPixmap(size, size)
        target.fill(Qt.GlobalColor.transparent)
        painter = QPainter(target)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        path = QPainterPath()
        path.addEllipse(0, 0, size, size)
        painter.setClipPath(path)
        x = (size - pixmap.width()) // 2
        y = (size - pixmap.height()) // 2
        painter.drawPixmap(x, y, pixmap)
        painter.end()
        return target

    def setup_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)

        self.app_container = QFrame()
        self.app_container.setObjectName("appContainer")
        self.app_container.setGraphicsEffect(self.create_glow_effect("#000000", 30, 5))
        container_layout = QVBoxLayout(self.app_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)

        # custom titlebar: macOS-style circular window controls, top-right
        self.titlebar = QFrame()
        self.titlebar.setObjectName("titleBar")
        self.titlebar.setFixedHeight(42)
        titlebar_layout = QHBoxLayout(self.titlebar)
        titlebar_layout.setContentsMargins(18, 10, 14, 0)

        tb_title = QLabel("✦ RemoPlayer")
        tb_title.setStyleSheet("color: #55556b; font-size: 12px; font-weight: bold; background: transparent;")

        def make_dot(color, hover_glyph, tip, slot):
            btn = QPushButton(hover_glyph)
            btn.setFixedSize(16, 16)
            btn.setToolTip(tip)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: {color}; border-radius: 8px; border: none;
                    color: transparent; font-size: 9px; font-weight: bold; padding: 0;
                }}
                QPushButton:hover {{ color: rgba(0,0,0,160); }}
            """)
            btn.clicked.connect(slot)
            return btn

        self.min_dot = make_dot("#febc2e", "–", "کوچک کردن", self.showMinimized)
        self.max_dot = make_dot("#28c840", "⤢", "بزرگ کردن", self.toggle_maximize)
        self.close_dot = make_dot("#ff5f57", "✕", "بستن", self.close)

        titlebar_layout.addWidget(tb_title)
        titlebar_layout.addStretch()
        titlebar_layout.addWidget(self.min_dot)
        titlebar_layout.addSpacing(8)
        titlebar_layout.addWidget(self.max_dot)
        titlebar_layout.addSpacing(8)
        titlebar_layout.addWidget(self.close_dot)
        container_layout.addWidget(self.titlebar)

        top_content = QFrame()
        top_layout = QHBoxLayout(top_content)
        top_layout.setContentsMargins(20, 5, 20, 20)
        top_layout.setSpacing(20)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(320)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(15, 20, 15, 20)
        
        logo_layout = QHBoxLayout()
        logo_label = QLabel("RemoPlayer")
        logo_label.setObjectName("logo")
        logo_label.setStyleSheet("font-size: 18px")
        
        settings_btn = QPushButton("⚙️")
        settings_btn.setObjectName("settingsBtn")
        settings_btn.setFixedSize(35, 35)
        settings_btn.clicked.connect(self.open_settings)
        
        logo_layout.addWidget(logo_label)
        logo_layout.addStretch()
        logo_layout.addWidget(settings_btn)
        
        mode_layout = QHBoxLayout()
        self.btn_radio_mode = QPushButton("رادیو")
        self.btn_radio_mode.setObjectName("modeBtnActive")
        self.btn_radio_mode.clicked.connect(lambda: self.switch_mode("radio"))
        
        self.btn_music_mode = QPushButton("موزیک")
        self.btn_music_mode.setObjectName("modeBtn")
        self.btn_music_mode.clicked.connect(lambda: self.switch_mode("music"))
        
        self.btn_history_mode = QPushButton("تاریخچه")
        self.btn_history_mode.setObjectName("modeBtn")
        self.btn_history_mode.clicked.connect(lambda: self.switch_mode("history"))
        
        mode_layout.addWidget(self.btn_radio_mode)
        mode_layout.addWidget(self.btn_music_mode)
        mode_layout.addWidget(self.btn_history_mode)
        
        self.folder_info_label = QLabel()
        self.folder_info_label.setObjectName("folderInfo")
        self.folder_info_label.setWordWrap(True)
        self.folder_info_label.hide()

        self.refresh_btn = QPushButton("🔄 بروزرسانی لیست موزیک")
        self.refresh_btn.setObjectName("modeBtn")
        self.refresh_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.refresh_btn.clicked.connect(self.force_refresh_music)
        self.refresh_btn.hide()
        
        # one persistent list per tab — no rebuild (and no lag) on tab switch
        self.radio_list = QListWidget()
        self.radio_list.setObjectName("mediaList")
        self.radio_list.itemClicked.connect(self.handle_radio_selection)

        self.music_list = QListWidget()
        self.music_list.setObjectName("mediaList")
        self.music_list.itemClicked.connect(self.handle_music_selection)

        self.history_listw = QListWidget()
        self.history_listw.setObjectName("mediaList")

        self.list_stack = QStackedWidget()
        self.list_stack.setStyleSheet("background: transparent;")
        self.list_stack.addWidget(self.radio_list)
        self.list_stack.addWidget(self.music_list)
        self.list_stack.addWidget(self.history_listw)

        sidebar_layout.addLayout(logo_layout)
        sidebar_layout.addSpacing(15)
        sidebar_layout.addLayout(mode_layout)
        sidebar_layout.addSpacing(15)
        sidebar_layout.addWidget(self.folder_info_label)
        sidebar_layout.addWidget(self.refresh_btn)
        sidebar_layout.addWidget(self.list_stack)
        
        main_panel = QFrame()
        main_panel.setObjectName("mainPanel")
        panel_layout = QVBoxLayout(main_panel)
        panel_layout.setContentsMargins(20, 20, 20, 20)

        cover_container = QWidget()
        cover_container.setStyleSheet("background: transparent;")
        cover_container_layout = QVBoxLayout(cover_container)
        cover_container_layout.setContentsMargins(0, 0, 0, 0)
        
        cover_frame = QFrame()
        cover_frame.setStyleSheet("""
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                stop:0 rgba(255,0,127,0.1), stop:0.5 rgba(0,240,255,0.1), stop:1 rgba(112,0,255,0.1));
            border-radius: 80px;
            padding: 10px;
        """)
        cover_frame_layout = QVBoxLayout(cover_frame)
        cover_frame_layout.setContentsMargins(0, 0, 0, 0)
        
        self.cover_art = QLabel("🎧")
        self.cover_art.setObjectName("coverArt")
        self.cover_art.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cover_art.setGraphicsEffect(self.create_glow_effect("#ff007f", 50, 0))
        
        cover_frame_layout.addWidget(self.cover_art, alignment=Qt.AlignmentFlag.AlignCenter)
        cover_container_layout.addWidget(cover_frame, alignment=Qt.AlignmentFlag.AlignCenter)

        self.title_label = QLabel("یک آیتم انتخاب کن...")
        self.title_label.setObjectName("titleLabel")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setWordWrap(True)

        self.subtitle_label = QLabel("آماده به کار")
        self.subtitle_label.setObjectName("subtitleLabel")
        self.subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.music_progress_layout = QHBoxLayout()
        self.music_progress_layout.setSpacing(12)
        
        self.music_time_lbl = QLabel("00:00")
        self.music_time_lbl.setStyleSheet("color:#00f0ff; font-weight:bold; font-size: 13px; min-width: 45px;")
        
        self.music_slider = QSlider(Qt.Orientation.Horizontal)
        self.music_slider.setRange(0, 1000)
        self.music_slider.setStyleSheet("""
            QSlider::groove:horizontal { 
                border-radius: 4px; 
                height: 8px; 
                background: #1f1f2e; 
            }
            QSlider::sub-page:horizontal { 
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                    stop:0 #00f0ff, stop:0.5 #0088ff, stop:1 #7000ff); 
                border-radius: 4px; 
            }
            QSlider::handle:horizontal { 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                    stop:0 #ffffff, stop:1 #00f0ff); 
                width: 18px; 
                height: 18px; 
                margin: -5px 0; 
                border-radius: 9px; 
                border: 2px solid #00f0ff;
            }
            QSlider::handle:horizontal:hover {
                background: #00f0ff;
                width: 20px;
                height: 20px;
                margin: -6px 0;
                border-radius: 10px;
            }
        """)
        self.music_slider.sliderMoved.connect(self.seek_music)
        
        self.music_duration_lbl = QLabel("00:00")
        self.music_duration_lbl.setStyleSheet("color:#ff007f; font-weight:bold; font-size: 13px; min-width: 45px;")
        
        self.music_progress_layout.addWidget(self.music_time_lbl)
        self.music_progress_layout.addWidget(self.music_slider)
        self.music_progress_layout.addWidget(self.music_duration_lbl)
        
        self.progress_container = QWidget()
        self.progress_container.setLayout(self.music_progress_layout)
        self.progress_container.hide()

        controls_layout = QHBoxLayout()
        controls_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        controls_layout.setSpacing(18)

        self.shuffle_btn = QPushButton("🔀")
        self.shuffle_btn.setObjectName("sideBtn")
        self.shuffle_btn.setFixedSize(46, 46)
        self.shuffle_btn.setToolTip("پخش تصادفی")
        self.shuffle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.shuffle_btn.clicked.connect(self.toggle_shuffle)

        self.prev_btn = QPushButton("⏮")
        self.prev_btn.setObjectName("navBtn")
        self.prev_btn.setFixedSize(56, 56)
        self.prev_btn.setToolTip("آهنگ قبلی")
        self.prev_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.prev_btn.clicked.connect(self.prev_track)

        self.play_btn = QPushButton("▶")
        self.play_btn.setObjectName("playBtn")
        self.play_btn.setFixedSize(76, 76)
        self.play_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.play_btn.clicked.connect(self.toggle_play_pause)
        self.play_btn.setGraphicsEffect(self.create_glow_effect("#00f0ff", 35, 0))

        self.next_btn = QPushButton("⏭")
        self.next_btn.setObjectName("navBtn")
        self.next_btn.setFixedSize(56, 56)
        self.next_btn.setToolTip("آهنگ بعدی")
        self.next_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.next_btn.clicked.connect(self.next_track)

        self.repeat_btn = QPushButton("🔁")
        self.repeat_btn.setObjectName("sideBtn")
        self.repeat_btn.setFixedSize(46, 46)
        self.repeat_btn.setToolTip("تکرار: خاموش")
        self.repeat_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.repeat_btn.clicked.connect(self.cycle_repeat)

        controls_layout.addWidget(self.shuffle_btn)
        controls_layout.addWidget(self.prev_btn)
        controls_layout.addWidget(self.play_btn)
        controls_layout.addWidget(self.next_btn)
        controls_layout.addWidget(self.repeat_btn)

        vol_layout = QHBoxLayout()
        vol_icon = QLabel("🔊")
        vol_icon.setStyleSheet("color: #00f0ff; font-size: 22px; background: transparent;")
        
        self.volume_slider = QSlider(Qt.Orientation.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setValue(80)
        self.volume_slider.setFixedWidth(280)
        self.volume_slider.setStyleSheet("""
            QSlider::groove:horizontal { 
                border-radius: 4px; 
                height: 8px; 
                background: #1f1f2e; 
            }
            QSlider::sub-page:horizontal { 
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                    stop:0 #ff007f, stop:1 #ff5580); 
                border-radius: 4px; 
            }
            QSlider::handle:horizontal { 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                    stop:0 #ffffff, stop:1 #ff007f); 
                width: 18px; 
                height: 18px; 
                margin: -5px 0; 
                border-radius: 9px; 
                border: 2px solid #ff007f;
            }
            QSlider::handle:horizontal:hover {
                background: #ff007f;
                width: 20px;
                height: 20px;
                margin: -6px 0;
                border-radius: 10px;
            }
        """)
        self.volume_slider.valueChanged.connect(self.set_volume)
        self.player.audio_set_volume(80)
        
        self.volume_label = QLabel("80%")
        self.volume_label.setStyleSheet("color: #8b8b9c; font-weight: bold; font-size: 12px; min-width: 35px; background: transparent;")
        self.volume_slider.valueChanged.connect(lambda v: self.volume_label.setText(f"{v}%"))
        
        vol_layout.addStretch()
        vol_layout.addWidget(vol_icon)
        vol_layout.addSpacing(10)
        vol_layout.addWidget(self.volume_slider)
        vol_layout.addSpacing(8)
        vol_layout.addWidget(self.volume_label)
        vol_layout.addStretch()

        panel_layout.addStretch(1)
        panel_layout.addWidget(cover_container)
        panel_layout.addSpacing(20)
        panel_layout.addWidget(self.title_label)
        panel_layout.addSpacing(8)
        panel_layout.addWidget(self.subtitle_label)
        panel_layout.addSpacing(20)
        panel_layout.addWidget(self.progress_container)
        panel_layout.addSpacing(20)
        panel_layout.addLayout(controls_layout)
        panel_layout.addSpacing(25)
        panel_layout.addLayout(vol_layout)
        panel_layout.addStretch(1)

        top_layout.addWidget(sidebar) 
        top_layout.addWidget(main_panel) 
        container_layout.addWidget(top_content)

        info_bar = QFrame()
        info_bar.setObjectName("infoBar")
        info_bar.setFixedHeight(50)
        info_layout = QHBoxLayout(info_bar)
        info_layout.setContentsMargins(25, 0, 25, 0)
        
        self.duration_label = QLabel("⏱ تب فعلی: 00:00")
        self.duration_label.setObjectName("infoText")
        self.duration_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        
        self.stats_label = QLabel("📊 حجم مصرفی: 0.00 MB  |  کل زمان: 00:00:00")
        self.stats_label.setObjectName("infoText")
        self.stats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        copyright_label = QLabel("<a href='https://nimaebrahimi.ir' style='color:#55556b;text-decoration:none;'>Code & UI by nimaebrahimi.ir</a>")
        copyright_label.setOpenExternalLinks(True)
        copyright_label.setObjectName("copyrightText")
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        info_layout.addWidget(self.duration_label, stretch=1)
        info_layout.addWidget(self.stats_label, stretch=2)
        info_layout.addWidget(copyright_label, stretch=1)

        grip = QSizeGrip(info_bar)
        grip.setStyleSheet("background: transparent; width: 14px; height: 14px;")
        info_layout.addWidget(grip, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)
        
        container_layout.addWidget(info_bar)
        main_layout.addWidget(self.app_container)

    def apply_styles(self):
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI', Tahoma, sans-serif; }
            #appContainer { background: #0f0f15; border-radius: 20px; border: 1px solid #1f1f2e; }
            #titleBar { background: transparent; border-top-left-radius: 20px; border-top-right-radius: 20px; }
            #sidebar { background-color: #16161f; border-radius: 15px; }
            #mainPanel { background-color: transparent; }
            #logo { font-size: 20px; font-weight: 900; color: #ffffff; letter-spacing: 1px; background: transparent; }
            #settingsBtn { background: #2a2a3e; border-radius: 17px; font-size: 16px; border: none; }
            #settingsBtn:hover { background: #3a3a4e; }
            #modeBtn { background: #1f1f2e; color: #8b8b9c; border-radius: 10px; padding: 8px; font-weight: bold; border: none; font-size: 12px; }
            #modeBtn:hover { background: #2a2a3e; color: white; }
            #modeBtnActive { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff007f, stop:1 #7000ff); color: white; border-radius: 10px; padding: 8px; font-weight: bold; border: none; font-size: 12px; }
            #mediaList { background-color: transparent; border: none; color: #8b8b9c; font-size: 13px; font-weight: bold; outline: none; }
            #mediaList::item { padding: 0px; margin-bottom: 8px; border-radius: 8px; background-color: #1c1c28; }
            #mediaList::item:hover { background-color: #242436; color: white; }
            #mediaList::item:selected { background: #2a2a3e; color: #00f0ff; border-left: 4px solid #00f0ff; }
            #folderInfo { color: #8b8b9c; font-size: 11px; padding: 8px; background: #1c1c28; border-radius: 6px; margin-bottom: 8px; }
            #coverArt { font-size: 70px; background-color: #16161f; border-radius: 70px; min-width: 140px; max-width: 140px; min-height: 140px; max-height: 140px; border: 4px solid #ff007f; }
            #titleLabel { font-size: 22px; font-weight: bold; color: #ffffff; background: transparent; }
            #subtitleLabel { font-size: 14px; color: #00f0ff; margin-top: 4px; font-weight: bold; background: transparent; }
            QPushButton { background-color: #1f1f2e; color: white; border-radius: 28px; font-size: 20px; border: none; }
            QPushButton:hover { background-color: #2a2a3e; }
            #playBtn { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #00f0ff, stop:1 #0055ff); font-size: 28px; border-radius: 38px; }
            #playBtn:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1affff, stop:1 #0077ff); }
            #navBtn { background: #1c1c28; font-size: 20px; border-radius: 28px; border: 1px solid #2a2a3e; }
            #navBtn:hover { background: #2a2a3e; border: 1px solid #00f0ff; }
            #navBtn:disabled { color: #3a3a4a; border: 1px solid #1f1f2e; }
            #sideBtn { background: transparent; font-size: 17px; border-radius: 23px; border: 1px solid #2a2a3e; color: #8b8b9c; }
            #sideBtn:hover { background: #1c1c28; }
            #sideBtnActive { background: rgba(0,240,255,0.12); font-size: 17px; border-radius: 23px; border: 1px solid #00f0ff; color: #00f0ff; }
            #sideBtn:disabled, #navBtn:disabled { color: #3a3a4a; }
            #infoBar { background-color: #111118; border-bottom-left-radius: 20px; border-bottom-right-radius: 20px; border-top: 1px solid #1f1f2e; }
            #infoText { color: #7a7a8f; font-size: 12px; font-weight: bold; background: transparent; }
            #copyrightText { font-size: 12px; font-weight: bold; background: transparent; }
            QSlider::groove:horizontal { border-radius: 3px; height: 6px; background: #1f1f2e; }
            QSlider::sub-page:horizontal { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00f0ff, stop:1 #7000ff); border-radius: 3px; }
            QSlider::handle:horizontal { background: white; width: 16px; height: 16px; margin: -5px 0; border-radius: 8px; border: 2px solid #00f0ff; }
            QSlider::handle:horizontal:hover { background: #00f0ff; width: 18px; height: 18px; margin: -6px 0; border-radius: 9px; }
        """)

    def open_settings(self):
        old_folder = self.config_mgr.config.get("music_folder", "")
        dialog = SettingsDialog(self, self.config_mgr)
        if dialog.exec():
            self.apply_proxy_settings()
            self.load_radio_list()
            if self.config_mgr.config.get("music_folder", "") != old_folder:
                self.load_music_list()  # folder changed → re-scan

    def toggle_play_pause(self):
        if self.is_playing:
            self.pause_media()
        else:
            if self.player.get_media():
                self.play_media()
            elif self.current_mode == "radio" and self.radio_list.currentItem():
                self.handle_radio_selection(self.radio_list.currentItem())
            elif self.current_mode == "music" and self.music_list.currentItem():
                self.handle_music_selection(self.music_list.currentItem())

    def switch_mode(self, mode):
        self.current_mode = mode
        self.btn_radio_mode.setObjectName("modeBtnActive" if mode == "radio" else "modeBtn")
        self.btn_music_mode.setObjectName("modeBtnActive" if mode == "music" else "modeBtn")
        self.btn_history_mode.setObjectName("modeBtnActive" if mode == "history" else "modeBtn")
        
        # progress bar belongs to actual music playback, not to the music tab
        self.progress_container.setVisible(self.playing_source == "music")

        # keep cover art of whatever is actually playing; only reset when idle
        self.refresh_btn.setVisible(mode == "music")
        if mode == "radio":
            self.folder_info_label.hide()
            if not self.is_playing:
                self.cover_art.clear()
                self.cover_art.setText("🎧")
                self.cover_art.setGraphicsEffect(self.create_glow_effect("#ff007f", 40, 0))
                self.cover_art.setStyleSheet("border: 3px solid #ff007f;")
            self.list_stack.setCurrentWidget(self.radio_list)
        elif mode == "music":
            if not self.is_playing:
                self.cover_art.clear()
                self.cover_art.setText("🎵")
                self.cover_art.setGraphicsEffect(self.create_glow_effect("#00f0ff", 40, 0))
                self.cover_art.setStyleSheet("border: 3px solid #00f0ff;")

            folder_path = self.config_mgr.config.get("music_folder", "")
            short_path = folder_path.split(os.sep)[-1] if folder_path else "نامشخص"
            self.folder_info_label.setText(f"📁 {short_path}")
            self.folder_info_label.show()
            self.list_stack.setCurrentWidget(self.music_list)
            if not self.music_loaded:
                self.load_music_list()
        elif mode == "history":
            self.folder_info_label.hide()
            if not self.is_playing:
                self.cover_art.clear()
                self.cover_art.setText("📜")
                self.cover_art.setGraphicsEffect(self.create_glow_effect("#7000ff", 40, 0))
                self.cover_art.setStyleSheet("border: 3px solid #7000ff;")
            self.list_stack.setCurrentWidget(self.history_listw)
            self.load_history_list()

        self.setStyleSheet(self.styleSheet())

    def load_radio_list(self):
        self.radio_list.clear()
        for name, station in self.config_mgr.config["stations"].items():
            item = QListWidgetItem()
            item.setSizeHint(QSize(0, 64))
            self.radio_list.addItem(item)
            desc = station.get("desc") or "ایستگاه رادیویی آنلاین"
            widget = MediaItemWidget(name, desc, "پخش زنده", icon_text="📻")
            self.radio_list.setItemWidget(item, widget)

    def load_music_list(self):
        if self.metadata_loader and self.metadata_loader.isRunning():
            self.metadata_loader.stop()
            self.metadata_loader.wait()

        self.music_list.clear()
        self.music_files = []
        self.music_metadata_cache = {}
        folder = self.config_mgr.config.get("music_folder", "")

        if os.path.exists(folder):
            extensions = ('.mp3', '.wav', '.flac', '.ogg')
            for f in os.listdir(folder):
                if f.lower().endswith(extensions):
                    full_path = os.path.join(folder, f)
                    self.music_files.append(full_path)

                    item = QListWidgetItem()
                    item.setSizeHint(QSize(0, 64))
                    self.music_list.addItem(item)

                    filename = os.path.splitext(f)[0]
                    if len(filename) > 30:
                        filename = filename[:30] + "..."

                    widget = MediaItemWidget(filename, "در حال بارگذاری اطلاعات...", "", icon_text="⏳", rtl=False)
                    self.music_list.setItemWidget(item, widget)

        if not self.music_files:
            item = QListWidgetItem()
            item.setSizeHint(QSize(0, 64))
            self.music_list.addItem(item)
            widget = MediaItemWidget("هیچ موزیکی یافت نشد!", "", "")
            self.music_list.setItemWidget(item, widget)
        else:
            self.metadata_loader = MusicMetadataLoader(self.music_files, self.instance)
            self.metadata_loader.metadata_loaded.connect(self.on_metadata_loaded)
            self.metadata_loader.finished_all.connect(self.on_metadata_finished)
            self.metadata_loader.start()
        self.music_loaded = True

    def force_refresh_music(self):
        self.load_music_list()
    
    def on_metadata_loaded(self, idx, metadata):
        self.music_metadata_cache[idx] = metadata
        self.set_music_item_widget(idx, metadata)

    def set_music_item_widget(self, idx, metadata):
        item = self.music_list.item(idx)
        if not item:
            return
        title = metadata['title']
        artist = metadata['artist'] or "نامشخص"
        album = metadata['album']
        duration_ms = metadata['duration']
        image_data = metadata.get('image')

        duration_str = self.format_time(duration_ms // 1000) if duration_ms > 0 else "؟؟:؟؟"
        if len(title) > 30:
            title = title[:30] + "..."

        subtitle = f"👤 {artist}"
        if album:
            subtitle += f"  •  💿 {album}"

        widget = MediaItemWidget(title, subtitle, duration_str, icon_text="🎵", image_data=image_data, rtl=False)
        self.music_list.setItemWidget(item, widget)
    
    def on_metadata_finished(self):
        pass

    def load_history_list(self):
        self.history_listw.clear()
        if not self.history_list:
            item = QListWidgetItem()
            item.setSizeHint(QSize(0, 64))
            self.history_listw.addItem(item)
            widget = MediaItemWidget("تاریخچه پخش خالی است", "آهنگی پخش نشده", "", icon_text="📜")
            self.history_listw.setItemWidget(item, widget)
            return

        for data in reversed(self.history_list):
            item = QListWidgetItem()
            item.setSizeHint(QSize(0, 64))
            self.history_listw.addItem(item)
            widget = MediaItemWidget(data['title'], data['subtitle'], data['extra'], icon_text=data['icon'], image_data=data.get('image_data'), rtl=(data['icon'] == '📻'))
            self.history_listw.setItemWidget(item, widget)

    def update_visualizer(self):
        if self.is_playing:
            if self.current_mode == "music" and not self.cover_art.pixmap().isNull():
                self.cover_art.setGraphicsEffect(self.create_glow_effect("#00f0ff", 40 + (self.visualizer_frame % 2) * 10, 0))
                self.visualizer_frame = (self.visualizer_frame + 1) % 8
                return
                
            emojis = ["🎵", "🎶", "🎧", "🎼", "🎹", "🎸", "🎺", "🎻"]
            self.visualizer_frame = (self.visualizer_frame + 1) % len(emojis)
            
            if self.current_mode == "radio":
                base_emoji = emojis[self.visualizer_frame]
                glow_color = "#ff007f"
            else:
                base_emoji = emojis[self.visualizer_frame]
                glow_color = "#00f0ff"
            
            self.cover_art.clear()
            self.cover_art.setText(base_emoji)
            self.cover_art.setGraphicsEffect(self.create_glow_effect(glow_color, 50, 0))
        else:
            if self.current_mode == "radio":
                self.cover_art.clear()
                self.cover_art.setText("🎧")
                self.cover_art.setGraphicsEffect(self.create_glow_effect("#ff007f", 50, 0))
            elif self.current_mode == "music":
                if self.cover_art.pixmap().isNull():
                    self.cover_art.clear()
                    self.cover_art.setText("🎵")
                self.cover_art.setGraphicsEffect(self.create_glow_effect("#00f0ff", 50, 0))
            else:
                self.cover_art.clear()
                self.cover_art.setText("📜")
                self.cover_art.setGraphicsEffect(self.create_glow_effect("#7000ff", 50, 0))
    
    def format_time(self, seconds):
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def update_stats(self):
        if self.player.is_playing():
            self.session_seconds += 1
            self.total_seconds += 1
            if self.playing_source == "radio":
                mb_used = (self.total_seconds * (self.bitrate_kbps / 8)) / 1024
                self.stats_label.setText(f"📊 حجم مصرفی: {mb_used:.2f} MB  |  کل زمان: {self.format_time(self.total_seconds)}")
            else:
                self.stats_label.setText(f"🎵 پخش لوکال  |  کل زمان استفاده: {self.format_time(self.total_seconds)}")
            self.duration_label.setText(f"⏱ تب فعلی: {self.format_time(self.session_seconds)}")

    def update_music_progress(self):
        # auto-advance when VLC reaches the end of a local track
        if (self.playing_source == "music" and self.is_playing
                and self.player.get_state() == vlc.State.Ended):
            QTimer.singleShot(0, self.handle_track_end)
            return
        if self.player.is_playing():
            length = self.player.get_length()
            time = self.player.get_time()
            if length > 0:
                self.music_slider.setValue(int((time / length) * 1000))
                self.music_time_lbl.setText(self.format_time(time // 1000))
                self.music_duration_lbl.setText(self.format_time(length // 1000))
                
                if time >= length - 1200:
                    if self.last_recorded_finish != self.current_music_url:
                        self.last_recorded_finish = self.current_music_url
                        entry = {
                            'title': self.current_media_name,
                            'subtitle': self.subtitle_label.text(),
                            'extra': self.music_duration_lbl.text(),
                            'icon': '🎵',
                            'image_data': self.current_art_data
                        }
                        if not self.history_list or self.history_list[-1]['title'] != entry['title']:
                            self.history_list.append(entry)
                            if self.current_mode == "history":
                                self.load_history_list()

    def seek_music(self, position):
        length = self.player.get_length()
        if length > 0:
            new_time = int(length * (position / 1000.0))
            self.player.set_time(new_time)

    def handle_radio_selection(self, item):
        idx = self.radio_list.row(item)
        stations_keys = list(self.config_mgr.config["stations"].keys())
        if idx >= len(stations_keys):
            return
        self.current_media_name = stations_keys[idx]
        url = self.config_mgr.config["stations"][self.current_media_name]["url"]
        self.title_label.setText("در حال اتصال...")
        self.subtitle_label.setText(f"📡 {self.current_media_name}")
        self.playing_source = "radio"
        self.progress_container.hide()
        self.start_playback(url)

    def handle_music_selection(self, item):
        idx = self.music_list.row(item)
        if idx < len(self.music_files):
            self.play_music_at(idx)

    def play_music_at(self, idx):
        if not (0 <= idx < len(self.music_files)):
            return
        self.current_music_index = idx
        self.play_order_history.append(idx)
        if len(self.play_order_history) > 100:
            self.play_order_history.pop(0)
        url = self.music_files[idx]
        self.current_music_url = url

        if idx in self.music_metadata_cache:
            metadata = self.music_metadata_cache[idx]
            self.current_media_name = metadata['title']
            self.title_label.setText(metadata['title'])
            self.current_art_data = metadata.get('image')

            artist = metadata['artist'] or "نامشخص"
            if metadata['album']:
                self.subtitle_label.setText(f"👤 {artist}  •  💿 {metadata['album']}")
            else:
                self.subtitle_label.setText(f"👤 {artist}")

            if self.current_art_data:
                circular_pixmap = self.create_circular_pixmap(self.current_art_data)
                self.cover_art.setPixmap(circular_pixmap)
                self.cover_art.setStyleSheet("border: 3px solid #00f0ff;")
            else:
                self.cover_art.clear()
                self.cover_art.setText("🎵")
                self.cover_art.setStyleSheet("border: 3px solid #00f0ff;")
        else:
            name = os.path.splitext(os.path.basename(url))[0]
            self.current_media_name = name
            self.title_label.setText(name)
            self.subtitle_label.setText("🎵 پخش لوکال")
            self.current_art_data = None
            self.cover_art.clear()
            self.cover_art.setText("🎵")

        self.playing_source = "music"
        self.progress_container.show()
        if self.music_list.currentRow() != idx:
            self.music_list.setCurrentRow(idx)
        self.start_playback(url)

    def start_playback(self, url):
        self.session_seconds = 0
        self.duration_label.setText("⏱ تب فعلی: 00:00")
        self.last_played_title = ""
        try:
            media = self.instance.media_new(url)
            self.player.set_media(media)
            self.play_media()
        except Exception as e:
            self.title_label.setText("❌ خطا در بارگذاری")
            self.subtitle_label.setText(str(e)[:50])

    def play_media(self):
        if self.player.get_media():
            self.player.play()
            self.is_playing = True
            self.play_btn.setText("⏸")

            if self.playing_source == "radio":
                self.subtitle_label.setText(f"🎶 در حال پخش: {self.current_media_name}")
                self.meta_timer.start(2000)

            if self.smtc:
                self.smtc.set_nav_enabled(self.playing_source == "music")
            self.update_smtc_system(PlaybackState.PLAYING)
            self.stats_timer.start(1000)
            self.progress_timer.start(500)

    def pause_media(self):
        self.player.pause()
        self.is_playing = False
        self.play_btn.setText("▶")
        if not self.subtitle_label.text().startswith("⏸"):
            self.subtitle_label.setText(f"⏸ متوقف شد - {self.subtitle_label.text()}")
        self.update_smtc_system(PlaybackState.PAUSED)
        self.meta_timer.stop()
        self.stats_timer.stop()
        self.progress_timer.stop()

    def stop_media(self):
        self.player.stop()
        self.is_playing = False
        self.playing_source = None
        self.play_btn.setText("▶")
        self.subtitle_label.setText("⏹ متوقف شد")
        self.title_label.setText("پخش متوقف شد")
        self.music_slider.setValue(0)
        self.music_time_lbl.setText("00:00")
        self.progress_container.hide()
        self.update_smtc_system(PlaybackState.STOPPED)
        self.meta_timer.stop()
        self.stats_timer.stop()
        self.progress_timer.stop()

    def toggle_shuffle(self):
        self.shuffle_enabled = not self.shuffle_enabled
        self.shuffle_btn.setObjectName("sideBtnActive" if self.shuffle_enabled else "sideBtn")
        self.shuffle_btn.setToolTip("پخش تصادفی: " + ("روشن" if self.shuffle_enabled else "خاموش"))
        self.setStyleSheet(self.styleSheet())

    def cycle_repeat(self):
        modes = {"off": ("all", "🔁", "تکرار: همه", "sideBtnActive"),
                 "all": ("one", "🔂", "تکرار: یک آهنگ", "sideBtnActive"),
                 "one": ("off", "🔁", "تکرار: خاموش", "sideBtn")}
        self.repeat_mode, icon, tip, style = modes[self.repeat_mode]
        self.repeat_btn.setText(icon)
        self.repeat_btn.setToolTip(tip)
        self.repeat_btn.setObjectName(style)
        self.setStyleSheet(self.styleSheet())

    def pick_next_index(self, direction=1):
        n = len(self.music_files)
        if n == 0:
            return -1
        if self.shuffle_enabled and n > 1:
            import random
            choices = [i for i in range(n) if i != self.current_music_index]
            return random.choice(choices)
        return (self.current_music_index + direction) % n

    def next_track(self):
        if self.playing_source != "music" or not self.music_files:
            return
        self.play_music_at(self.pick_next_index(1))

    def prev_track(self):
        if self.playing_source != "music" or not self.music_files:
            return
        # restart track if >3s in, like standard players
        if self.player.get_time() > 3000:
            self.player.set_time(0)
            return
        if self.shuffle_enabled and len(self.play_order_history) >= 2:
            self.play_order_history.pop()          # current
            prev_idx = self.play_order_history.pop()
            self.play_music_at(prev_idx)
            return
        self.play_music_at(self.pick_next_index(-1))

    def handle_track_end(self):
        if self.playing_source != "music":
            return
        if self.repeat_mode == "one":
            self.play_music_at(self.current_music_index)
        elif self.repeat_mode == "all" or self.shuffle_enabled:
            self.play_music_at(self.pick_next_index(1))
        else:
            if self.current_music_index < len(self.music_files) - 1:
                self.play_music_at(self.current_music_index + 1)
            else:
                self.stop_media()

    def update_smtc_system(self, status):
        if not self.smtc:
            return
        try:
            title = self.title_label.text()
            thumbnail_path = None
            if self.playing_source != "music":
                artist = self.current_media_name
                album_artist = "RemoPlayer Radio"
            else:
                artist_text = self.subtitle_label.text()
                if "👤" in artist_text:
                    artist = artist_text.split("👤")[1].split("•")[0].strip()
                else:
                    artist = "نامشخص"
                album_artist = "RemoPlayer"
                thumbnail_path = self._export_smtc_thumbnail()
            self.smtc.update(status, title, artist, album_artist, thumbnail_path)
        except Exception:
            pass

    def _export_smtc_thumbnail(self):
        """Write current cover art to a temp file so SMTC can show it."""
        if not self.current_art_data:
            return None
        path = os.path.join(self.config_mgr.appdata_dir, 'smtc_cover.jpg')
        try:
            with open(path, 'wb') as f:
                f.write(self.current_art_data)
            return path
        except Exception:
            return None

    def update_metadata(self):
        if self.current_mode != "radio":
            return
            
        media = self.player.get_media()
        if not media:
            return
        
        try:
            now_playing = media.get_meta(vlc.Meta.NowPlaying)
            
            if now_playing and now_playing != self.last_played_title:
                self.last_played_title = now_playing
                self.title_label.setText(now_playing)
                self.update_smtc_system(PlaybackState.PLAYING)
                
                entry = {
                    'title': now_playing,
                    'subtitle': self.current_media_name,
                    'extra': 'رادیو آنلاین',
                    'icon': '📻',
                    'image_data': None
                }
                if not self.history_list or self.history_list[-1]['title'] != now_playing:
                    self.history_list.append(entry)
                    if self.current_mode == "history":
                        self.load_history_list()
                
                if self.config_mgr.config.get("notifications_enabled", True):
                    icon_file = resource_path(os.path.join('assets', 'icon.png'))
                    def show_toast_bg(station_name, song_title):
                        try:
                            kwargs = {"duration": "short", "app_id": APP_ID}
                            if os.path.exists(icon_file):
                                kwargs["icon"] = {"src": Path(icon_file).as_uri(),
                                                  "placement": "appLogoOverride"}
                            toast(f"📻 {station_name}", song_title, **kwargs)
                        except Exception:
                            pass
                    threading.Thread(target=show_toast_bg, args=(self.current_media_name, now_playing), daemon=True).start()
        except:
            pass

    def set_volume(self, value):
        self.player.audio_set_volume(int(value))

    def toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def mousePressEvent(self, event):
        # drag window from the titlebar strip (frameless window)
        if (event.button() == Qt.MouseButton.LeftButton
                and event.position().y() < 60 and not self.isMaximized()):
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_pos is not None and event.buttons() & Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_pos = None
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.position().y() < 60:
            self.toggle_maximize()
        super().mouseDoubleClickEvent(event)

    def setup_tray(self, icon):
        self.tray = QSystemTrayIcon(icon, self)
        self.tray.setToolTip("RemoPlayer")
        menu = QMenu()
        show_action = menu.addAction("نمایش پلیر")
        show_action.triggered.connect(self.show_from_tray)
        toggle_action = menu.addAction("پخش/توقف")
        toggle_action.triggered.connect(self.toggle_play_pause)
        menu.addSeparator()
        quit_action = menu.addAction("خروج کامل")
        quit_action.triggered.connect(self.quit_app)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(
            lambda reason: self.show_from_tray()
            if reason == QSystemTrayIcon.ActivationReason.Trigger else None)
        self.tray.show()

    def show_from_tray(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def quit_app(self):
        self.force_quit = True
        self.close()

    def closeEvent(self, event):
        if self.config_mgr.config.get("close_to_tray", False) and not self.force_quit:
            event.ignore()
            self.hide()
            if self.tray:
                self.tray.showMessage("RemoPlayer", "برنامه در سینی سیستم در حال اجراست",
                                      QSystemTrayIcon.MessageIcon.Information, 2500)
            return
        if self.metadata_loader and self.metadata_loader.isRunning():
            self.metadata_loader.stop()
            self.metadata_loader.wait()
        self.player.stop()
        if self.tray:
            self.tray.hide()
        event.accept()
        QApplication.instance().quit()

if __name__ == "__main__":
    icon_file = resource_path(os.path.join('assets', 'icon.ico'))
    register_app_identity(icon_file)
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app_icon = QIcon(icon_file) if os.path.exists(icon_file) else QIcon()
    app.setWindowIcon(app_icon)
    app.setQuitOnLastWindowClosed(False)  # tray mode keeps app alive
    window = RemoPlayer()
    window.setup_tray(app_icon)
    window.show()
    sys.exit(app.exec())