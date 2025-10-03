import os, sys, time, logging, pathlib, threading, subprocess, psutil, mss, json, comtypes
import sounddevice as sd
from functools import partial
from datetime import datetime, timedelta
from functools import wraps
from typing import Dict, Any, List
from contextlib import contextmanager
from PIL import Image, ImageDraw
from flask import Flask, render_template, request, redirect, url_for, session, jsonify
from pystray import Icon, MenuItem, Menu
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

class Config:
    if getattr(sys, "frozen", False):  # Running as PyInstaller .exe
        BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        APP_DATA_DIR = os.path.dirname(sys.executable)
    else: # Running as script
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        APP_DATA_DIR = BASE_DIR

    STATIC_DIR = os.path.join(BASE_DIR, "static")
    APP_DATA_DIR = os.path.join(APP_DATA_DIR, "static")
    CONFIG_FILE = os.path.join(APP_DATA_DIR, "config.json")
    LOG_FILE = os.path.join(APP_DATA_DIR, "pc_remote.log")

    DEFAULTS = {
        "USERNAME": "admin",
        "PASSWORD": "password",
        "SECRET_KEY": "your-secret-key",
        "HOST": "0.0.0.0",
        "PORT": 5000,
        "DEBUG": False,
        "LOGGING_ENABLED": True,
        "SESSION_TIMEOUT_MINUTES": 30,
        "MAX_REQUESTS_PER_MINUTE": 120,
        "PLAYBACK_DEVICE_1": sd.query_devices(kind='output')['name'].split(' ')[0],
        "PLAYBACK_DEVICE_2": "HeadPhone",
        "RECORDING_DEVICE_1": sd.query_devices(kind='input')['name'].split(' ')[0],
    }
    def __init__(self):
        os.makedirs(self.APP_DATA_DIR, exist_ok=True)
        if not os.path.exists(self.CONFIG_FILE):
            with open(self.CONFIG_FILE, "w") as f: json.dump(self.DEFAULTS, f, indent=4)
        with open(self.CONFIG_FILE, "r") as f: config_data = json.load(f)
        for key, value in self.DEFAULTS.items(): setattr(self, key, config_data.get(key, value))

@contextmanager
def audio_context():
    try:
        comtypes.CoInitialize()
        yield
    finally:
        comtypes.CoUninitialize()

class PCRemoteControl:
    def __init__(self):
        self.config = Config()
        self._setup_logging()
        self.nircmd = os.path.join(self.config.STATIC_DIR, "nircmd.exe")
        self.running_apps_cache: Dict[str, Any] = {}
        self.last_cache_update: datetime = datetime.min
        self.cache_ttl = timedelta(seconds=2)
        self.app_enabled = True
        self.request_counts: Dict[str, List[datetime]] = {}
        self.modifier_key_timer: threading.Timer = None
        self.active_modifier: str = None
        self.active_modifiers: set = set()
        self.pictures_dir = pathlib.Path.home() / "Pictures" / "Remote Screenshots"
        self.apps = self._define_apps()
        self.flask_app = self._create_flask_app()

    def _setup_logging(self):
        if self.config.LOGGING_ENABLED:
            log_level = logging.DEBUG if self.config.DEBUG else logging.INFO
            logging.basicConfig(level=log_level, format="%(asctime)s - %(levelname)s - %(message)s", handlers=[logging.FileHandler(self.config.LOG_FILE, encoding="utf-8"), logging.StreamHandler(sys.stdout)])
        self.logger = logging.getLogger(__name__)

    def _define_apps(self) -> Dict[str, Dict[str, str]]:
        return {
        "discord": {"exe": "Discord.exe", "process_name": "discord.exe", "cmd": f'"{pathlib.Path.home() / "AppData/Local/Discord/Update.exe"}" --processStart Discord.exe', "closable": True},
        "steam": {"exe": "steam.exe", "process_name": "steam.exe", "cmd": str(pathlib.Path("C:/Program Files (x86)/Steam/Steam.exe")), "closable": True},
        "npp": {"exe": "notepad++.exe", "process_name": "notepad++.exe", "cmd": "start notepad++", "closable": True},
        "chrome": {"exe": "chrome.exe", "process_name": "chrome.exe", "cmd": "start chrome", "closable": True},
        "mediaplayer": {"exe": "mediaplayer.exe", "process_name": "mediaplayer.exe", "cmd": "start mediaplayer", "closable": True},
        "task_manager": {"exe": "Taskmgr.exe", "process_name": "taskmgr.exe", "cmd": "taskmgr", "closable": False},
        }

    def _update_running_apps_cache(self, force: bool = False):
        if not force and datetime.now() - self.last_cache_update < self.cache_ttl: return
        try:
            self.running_apps_cache = {p.info['name'].lower() for p in psutil.process_iter(['name']) if p.info['name']}
            self.last_cache_update = datetime.now()
        except Exception: pass
    # ----------------------- Rate Limiting -----------------------
    def _rate_limit_check(self, client_ip: str) -> bool:
        now = datetime.now()
        self.request_counts.setdefault(client_ip, [])
        self.request_counts[client_ip] = [t for t in self.request_counts[client_ip] if t > now - timedelta(minutes=1)]
        if len(self.request_counts[client_ip]) >= self.config.MAX_REQUESTS_PER_MINUTE:
            return False
        self.request_counts[client_ip].append(now)
        return True
    # ----------------------- Rate Limiting -----------------------
    # ----------------------- Flask Routes -----------------------
    def _create_flask_app(self) -> Flask:
        app = Flask(__name__,
                    # template_folder=self.config.TEMPLATE_DIR,
                    static_folder=self.config.STATIC_DIR)
        app.secret_key = self.config.SECRET_KEY

        @app.before_request
        def before_request():
            self.logger.debug(f"Request: {request.method} {request.path}")

            if not self.app_enabled and request.endpoint not in ("static",):
                return jsonify(error="App is disabled from system tray"), 503

            if not self._rate_limit_check(request.remote_addr):
                self.logger.warning(f"Rate limit exceeded for {request.remote_addr}")
                return jsonify(error="Rate limit exceeded"), 429

            if 'logged_in' in session and 'last_activity' in session:
                if datetime.now() - datetime.fromisoformat(session['last_activity']) > timedelta(minutes=self.config.SESSION_TIMEOUT):
                    session.clear()
                    return redirect(url_for('login'))
                session['last_activity'] = datetime.now().isoformat()
        
        def login_required(f):
            @wraps(f)
            def decorated_function(*args, **kwargs):
                if not session.get("logged_in"):
                    return redirect(url_for("login"))
                return f(*args, **kwargs)
            return decorated_function

        @app.route("/")
        @login_required
        def index(): return render_template("index.html")

        @app.route("/login", methods=["GET", "POST"])
        def login():
            if request.method == "POST" and request.form.get("username") == self.config.USERNAME and request.form.get("password") == self.config.PASSWORD:
                session["logged_in"] = True
                return redirect(url_for("index"))
            return render_template("login.html")

        @app.route("/logout")
        def logout():
            session.clear()
            return redirect(url_for("login"))

        @app.route("/api/status/all")
        @login_required
        def get_all_status():
            self._update_running_apps_cache(force=True)
            app_statuses = {}
            for app_key, app_config in self.apps.items():
                process_name = app_config.get('process_name', app_config['exe']).lower()
                app_statuses[app_key] = process_name in self.running_apps_cache

            return jsonify({
                "apps": app_statuses,
                "volume": self._get_volume(),
                "muted_state": self._get_mute_states()
            })

        @app.route("/api/action/<action_name>", methods=["POST"])
        @login_required
        def handle_action(action_name):
            actions = {
                'media_play_pause':   partial(self._create_simple_response,    f'"{self.nircmd}" sendkeypress 0xB3',                                  "Media play/pause toggled"),
                'media_next':         partial(self._create_simple_response,    f'"{self.nircmd}" sendkeypress 0xB0',                                  "Skipped to next track"),
                'media_previous':     partial(self._create_simple_response,    f'"{self.nircmd}" sendkeypress 0xB1',                                  "Skipped to previous track"),
                'undo':               partial(self._create_simple_response,    f'"{self.nircmd}" sendkeypress leftctrl+z',                               "Undo (Ctrl+Z)"),
                'redo':               partial(self._create_simple_response,    f'"{self.nircmd}" sendkeypress leftctrl+y',                               "Redo (Ctrl+Y)"),
                'sleep':              partial(self._create_simple_response,    f'"{self.nircmd}" standby',                                            "System sleep initiated"),
                'hard_sleep':         partial(self._create_simple_response,    'rundll32.exe powrprof.dll,SetSuspendState 0,1,0',                     "System hard sleep initiated"),
                'shutdown':           partial(self._create_simple_response,    'shutdown /s /t 5',                                                    "System shutdown initiated"),
                'restart':            partial(self._create_simple_response,    'shutdown /r /t 5',                                                    "System restart initiated"),
                'lock':               partial(self._create_simple_response,    'rundll32.exe user32.dll,LockWorkStation',                             "Workstation locked"),
                'mute_toggle_sound':  partial(self._create_simple_response,    f'"{self.nircmd}" mutesysvolume 2',                                    "System volume mute toggled"),
                'mute_toggle_mic':    partial(self._create_simple_response,    f'"{self.nircmd}" mutesysvolume 2 "{self.config.RECORDING_DEVICE_1}"', "Microphone mute toggled"),
                  
                'arrow_left':         partial(self._handle_standard_key_press, "left",                                                                "Left"),
                'arrow_up':           partial(self._handle_standard_key_press, "up",                                                                  "Up"),
                'arrow_right':        partial(self._handle_standard_key_press, "right",                                                               "Right"),
                'arrow_down':         partial(self._handle_standard_key_press, "down",                                                                "Down"),
  
                'press_enter':        partial(self._handle_standard_key_press, "enter",                                                               "Enter"),
                'press_space':        partial(self._handle_standard_key_press, "spc",                                                                 "Space"),
                'press_esc':          partial(self._handle_standard_key_press, "esc",                                                                 "Esc"),
                'press_backspace':    partial(self._handle_standard_key_press, "backspace", "Backspace"),
                'press_win':          partial(self._handle_standard_key_press, "lwin",                                                                "Window"),
                'press_tab':          partial(self._handle_standard_key_press, "tab",                                                                 "Tab"),
                'press_del':          partial(self._handle_standard_key_press, "delete",                                                                 "Delete"),
                'press_f4':           partial(self._handle_standard_key_press, "f4",                                                                  "F4"),
                'press_f5':           partial(self._handle_standard_key_press, "f5",                                                                  "F5"),
  
                'press_alt':          partial(self._handle_modifier_press,     "alt"),
                'press_ctrl':         partial(self._handle_modifier_press,     "ctrl"),
                'press_shift':        partial(self._handle_modifier_press,     "shift"),

                'screenshot':       self._take_screenshot,
                'audio_device_toggle': self._toggle_audio_device,
            }
            if action_name in actions:
                response_data = actions[action_name]()
                response_data["state"] = self._get_mute_states()
                return jsonify(response_data)
        
            return jsonify({"success": False, "message": "Invalid action"})

        @app.route("/api/app/<app_name>/toggle", methods=["POST"])
        @login_required
        def app_toggle(app_name):
            app_config = self.apps.get(app_name)
            if not app_config: return jsonify({"success": False, "message": "App not configured"})

            display_name = app_name.title()
            process_name = app_config.get('process_name', app_config['exe']).lower()
            is_closable = app_config.get("closable", True)
        
            self._update_running_apps_cache(force=True)
            is_running = process_name in self.running_apps_cache
            if is_running:
                if not is_closable:
                    return jsonify({
                        "success": True, 
                        "message": f"cannot be close {display_name}",
                        "running": True
                    })

                success = self._execute_command(f'taskkill /F /IM "{app_config["exe"]}"')
                message = f"{display_name} closed." if success else f"Failed to close {display_name}."
            else:
                try:
                    subprocess.Popen(app_config["cmd"], shell=True)
                    success = True
                    message = f"{display_name} started."
                except Exception as e:
                    success = False
                    message = f"Failed to start {display_name}: {str(e)}"
            time.sleep(0.5)
            self._update_running_apps_cache(force=True)
            return jsonify({
                "success": success,
                "message": message,
                "running": not is_running if success else is_running
            })

        @app.route("/api/volume/<int:value>", methods=["POST"])
        @login_required
        def set_volume(value: int):
            return jsonify(self._set_volume(value))

        return app
    # ----------------------- Flask Routes -----------------------
    # ----------------------- Executaion -----------------------
    def _create_simple_response(self, command: str, success_message: str) -> Dict[str, Any]:
        success = self._execute_command(command)
        return {
            "success": success,
            "message": success_message if success else f"Failed to execute: {command.split()[0]}"
        }

    def _execute_command(self, command: str) -> bool:
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                timeout=10,
                startupinfo=startupinfo,
                check=False
            )
            return result.returncode == 0
        except Exception:
            return False
        
    def _handle_arrow(self, key_code: str, message: str) -> Dict[str, Any]:
        return self._create_simple_response(f'"{self.nircmd}" sendkeypress {key_code}', message)
    # ----------------------- Executaion -----------------------
    # ----------------------- Alt Ctrl Shift Tab -----------------------
    def _reset_alt_tab_timer_if_active(self):
        if self.alt_tab_timer and self.alt_tab_timer.is_alive():
            self.alt_tab_timer.cancel()
    
            def cleanup_action():
                self._execute_command(f'"{self.nircmd}" sendkey 0xA4 up')  # Alt up
                self.alt_tab_timer = None
                self.logger.info("Alt-Tab session timed out after an interaction.")
    
            self.alt_tab_timer = threading.Timer(2.1, cleanup_action)
            self.alt_tab_timer.start()
            self.logger.info("Alt-Tab timer reset.")
            
    def _clear_modifier_state(self):
        MODIFIER_MAP = {"alt": "alt", "ctrl": "leftctrl", "shift": "leftshift"}
        for modifier in self.active_modifiers:
            self._execute_command(f'"{self.nircmd}" sendkey {MODIFIER_MAP[modifier]} up')

        if self.modifier_key_timer and self.modifier_key_timer.is_alive():
            self.modifier_key_timer.cancel()

        self.active_modifiers.clear()
        self.modifier_key_timer = None
        self.logger.info("All modifier keys have been released.")
        
    def _handle_modifier_press(self, modifier: str) -> Dict[str, Any]:
        MODIFIER_MAP = {"alt": "alt", "ctrl": "leftctrl", "shift": "leftshift"}

        if self.modifier_key_timer and self.modifier_key_timer.is_alive():
            self.modifier_key_timer.cancel()

        if modifier in self.active_modifiers:
            self.active_modifiers.remove(modifier)
            self._execute_command(f'"{self.nircmd}" sendkey {MODIFIER_MAP[modifier]} up')
            message = f"{modifier.capitalize()} released."
        else:
            self.active_modifiers.add(modifier)
            self._execute_command(f'"{self.nircmd}" sendkey {MODIFIER_MAP[modifier]} down')
            message = f"{modifier.capitalize()} is active."

        if self.active_modifiers:
            self.modifier_key_timer = threading.Timer(4.1, self._clear_modifier_state)
            self.modifier_key_timer.start()
        return {"success": True, "message": message, "activeModifiers": list(self.active_modifiers)}

    def _handle_standard_key_press(self, key_command: str, message: str) -> Dict[str, Any]:
        time.sleep(0.1)
        if self.active_modifier:
            MODIFIER_MAP = {"alt": "alt", "ctrl": "leftctrl", "shift": "leftshift"}
            modifier_keys = [MODIFIER_MAP[mod] for mod in self.active_modifiers]
            combo = "+".join(modifier_keys)
            
            full_command = f'"{self.nircmd}" sendkeypress {combo}+{key_command}'
            success_message = f"Sent {'+'.join([m.capitalize() for m in self.active_modifiers])}+{message}"
            self._clear_modifier_state()
        else:
            full_command = f'"{self.nircmd}" sendkeypress {key_command}'
            success_message = f"Sent {message}"
    
        success = self._execute_command(full_command)
        return {
            "success": success,
            "message": success_message if success else "Command failed",
            "activeModifiers": list(self.active_modifiers)
        }
    # ----------------------- Alt Ctrl Shift Tab -----------------------
    # ----------------------- Audio Device -----------------------
    def _get_volume(self) -> Dict[str, Any]:
        try:
            with audio_context():
                speakers = AudioUtilities.GetSpeakers()
                interface = speakers.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
                volume = interface.QueryInterface(IAudioEndpointVolume)
                return {"volume": round(volume.GetMasterVolumeLevelScalar() * 100)}
        except Exception as e:
            self.logger.error(f"Failed to get volume: {e}")
            return {"volume": 50}

    def _set_volume(self, value: int):
        try:
            with audio_context():
                value = max(0, min(100, value))
                speakers = AudioUtilities.GetSpeakers()
                interface = speakers.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
                volume = interface.QueryInterface(IAudioEndpointVolume)
                volume.SetMasterVolumeLevelScalar(value / 100.0, None)
            return {"success": True, "message": f"Volume set to {value}%"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def _get_mute_states(self) -> Dict[str, Any]:
        speaker_muted = False
        mic_muted = False
        device_name = "Unknown"
        is_headphone_active = False

        try:
            sd._terminate()
            sd._initialize()
            device_name = sd.query_devices(kind='output')['name'].split(' ')[0]
            
            headphone_name_from_config = self.config.PLAYBACK_DEVICE_2
            if headphone_name_from_config.lower() in device_name.lower():
                is_headphone_active = True

        except Exception as e:
            self.logger.error(f"Could not get device name using sounddevice: {e}")

        try:
            with audio_context():
                try:
                    speakers = AudioUtilities.GetSpeakers()
                    interface = speakers.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
                    volume_speaker = interface.QueryInterface(IAudioEndpointVolume)
                    speaker_muted = bool(volume_speaker.GetMute())
                except Exception as e:
                    self.logger.error(f"Could not get speaker mute status using pycaw: {e}")

                try:
                    mic = AudioUtilities.GetMicrophone()
                    interface_mic = mic.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
                    volume_mic = interface_mic.QueryInterface(IAudioEndpointVolume)
                    mic_muted = bool(volume_mic.GetMute())
                except Exception as e:
                    self.logger.error(f"Could not get microphone mute status using pycaw: {e}")
        except Exception as e:
            self.logger.error(f"A general pycaw audio context error occurred: {e}")
        return {
            "speaker_muted": speaker_muted,
            "mic_muted": mic_muted,
            "is_headphone_active": is_headphone_active,
        }

    def _toggle_audio_device(self) -> Dict[str, Any]:
        try:
            sd._terminate()
            sd._initialize()
            current_device = sd.query_devices(kind='output')['name'].split(' ')[0]
            target_device = self.config.PLAYBACK_DEVICE_2 if current_device == self.config.PLAYBACK_DEVICE_1 else self.config.PLAYBACK_DEVICE_1
            
            success = self._execute_command(f'"{self.nircmd}" setdefaultsounddevice "{target_device}" 1')
            
            if success:
                return {"success": True, "message": f"Switched to {target_device}"}
            else:
                return {"success": False, "message": "Failed to switch audio device"}
        except Exception as e:
            self.logger.error(f"Error toggling audio device: {e}")
            return {"success": False, "message": str(e)}
    # ----------------------- Audio Device Toggle -----------------------
    # ----------------------- Screen Shot -----------------------
    def _take_screenshot(self) -> Dict[str, Any]:
        try:
            filepath = self.pictures_dir / f"screenshot_{int(time.time())}.png"
            self.pictures_dir.mkdir(parents=True, exist_ok=True)
            with mss.mss() as sct:
                monitor = sct.monitors[0]
                img = Image.frombytes("RGB", sct.grab(monitor).size, sct.grab(monitor).rgb)
                img.save(filepath)
            return {"success": True, "message": f"Screenshot saved to {filepath.name}"}
        except Exception as e:
            return {"success": False, "message": str(e)}
    # ----------------------- Screen Shot -----------------------
    # ----------------------- System Tray -----------------------
    def _create_tray_image(self, enabled: bool = True) -> Image.Image:
        img = Image.new('RGB', (64, 64), (0, 128, 255) if enabled else (128, 128, 128))
        draw = ImageDraw.Draw(img)
        draw.rectangle([12, 16, 52, 38], fill=(255, 255, 255))
        draw.rectangle([28, 38, 36, 46], fill=(255, 255, 255))
        draw.rectangle([20, 46, 44, 50], fill=(255, 255, 255))
        return img

    def _toggle_app_enabled(self, icon):
        self.app_enabled = not self.app_enabled
        icon.icon = self._create_tray_image(self.app_enabled)
        self.logger.info(f"App {'enabled' if self.app_enabled else 'disabled'} from system tray")

    def _quit_app(self, icon):
        self.logger.info("Shutting down from system tray")
        icon.stop()
    # ----------------------- System Tray -----------------------
    # ----------------------- Run -----------------------
    def run_flask(self):
        self.flask_app.run(host=self.config.HOST, port=self.config.PORT, debug=self.config.DEBUG, threaded=True)

    def run(self):
        try:
            menu = Menu(
                MenuItem(lambda item: "Disable App" if self.app_enabled else "Enable App", self._toggle_app_enabled),
                MenuItem('Quit', self._quit_app)
            )
            self.icon = Icon("PC Remote Control", self._create_tray_image(True), "PC Remote Control", menu)
            threading.Thread(target=self.run_flask, daemon=True).start()
            self.logger.info(f"PC Remote Control started on {self.config.HOST}:{self.config.PORT}")
            self.icon.run()
        except KeyboardInterrupt:
            self.logger.info("Shutting down...")
        finally:
            self.logger.info("Application stopped")
    # ----------------------- Run -----------------------

if __name__ == "__main__":
    PCRemoteControl().run()