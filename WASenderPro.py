import sys
import os
import platform
import subprocess
import requests
import zipfile
import tarfile
import stat
import time
import random
import pygame
import xlsxwriter
import json
import urllib.parse
import logging
import re
import pandas as pd
import phonenumbers
import openpyxl 

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QFileDialog, QWidget,
    QMessageBox, QFrame, QMenuBar, QMenu, QAction,
    QColorDialog, QFontDialog, QInputDialog, QProgressBar
)
from PyQt5.QtGui import QFont, QColor, QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    WebDriverException, 
    TimeoutException, 
    SessionNotCreatedException, 
    NoSuchElementException
)

# ------------------- Configuration -------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# ------------------- Dependency Installer (الكود الأصلي الكامل) -------------------
class DependencyInstaller:
    def __init__(self):
        self.system = platform.system().lower()
        self.arch = platform.machine().lower()
        self.driver_dir = os.path.join(os.getcwd(), "drivers")
        os.makedirs(self.driver_dir, exist_ok=True)
        self.browser_paths = self._detect_browsers()
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": USER_AGENT})

    def _detect_browsers(self):
        browsers = {
            "chrome": self._get_browser_path("chrome"),
            "firefox": self._get_browser_path("firefox"),
            "brave": self._get_browser_path("brave"),
            "edge": self._get_browser_path("edge")
        }
        return {k: v for k, v in browsers.items() if v and os.path.exists(v)}

    def _get_browser_path(self, browser_name):
        if self.system == "windows":
            paths = {
                "chrome": os.path.join(os.getenv("ProgramFiles"), "Google", "Chrome", "Application", "chrome.exe"),
                "firefox": os.path.join(os.getenv("ProgramFiles"), "Mozilla Firefox", "firefox.exe"),
                "brave": os.path.join(os.getenv("ProgramFiles"), "BraveSoftware", "Brave-Browser", "Application", "brave.exe"),
                "edge": os.path.join(os.getenv("ProgramFiles(x86)"), "Microsoft", "Edge", "Application", "msedge.exe")
            }
        elif self.system == "linux":
            paths = {
                "chrome": "/usr/bin/google-chrome",
                "firefox": "/usr/bin/firefox",
                "brave": "/usr/bin/brave-browser",
                "edge": "/usr/bin/microsoft-edge"
            }
        elif self.system == "darwin":
            paths = {
                "chrome": "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "firefox": "/Applications/Firefox.app/Contents/MacOS/firefox",
                "brave": "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
                "edge": "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
            }
        return paths.get(browser_name)

    def _get_chrome_version(self):
        chrome_path = self.browser_paths.get("chrome")
        if not chrome_path: return None
        try:
            if self.system == "windows":
                try:
                    import winreg
                    reg_path = r'SOFTWARE\Google\Chrome\BLBeacon'
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                        version, _ = winreg.QueryValueEx(key, 'version')
                        return version
                except Exception:
                    command = f'(Get-Item "{chrome_path}").VersionInfo.FileVersion'
                    result = subprocess.run(["powershell", "-Command", command], 
                                          capture_output=True, text=True, check=True)
                    return result.stdout.strip()
            elif self.system == "linux":
                result = subprocess.run([chrome_path, "--version"], capture_output=True, text=True, check=True)
                return result.stdout.strip().split()[-1]
            elif self.system == "darwin":
                plist_path = os.path.join(os.path.dirname(chrome_path), '..', 'Info.plist')
                with open(os.path.abspath(plist_path), 'rb') as f:
                    content = f.read().decode('utf-8', errors='ignore')
                    match = re.search(r'<key>CFBundleShortVersionString</key>\s*<string>([\d.]+)</string>', content)
                    if match: return match.group(1)
                return None
        except Exception as e:
            logging.error(f"Error getting Chrome version: {e}")
            return None

    def _get_chrome_platform(self):
        if self.system == "windows":
            return "win64" if self.arch == "amd64" else "win32"
        elif self.system == "linux":
            return "linux64" if self.arch == "x86_64" else "linux32"
        elif self.system == "darwin":
            return "mac-arm64" if self.arch == "arm64" else "mac-x64"
        else: return None

    def is_python_package_installed(self, package_name):
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "show", package_name],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except subprocess.CalledProcessError:
            return False

    def install_python_packages(self):
        packages = ["selenium", "pygame", "xlsxwriter", "PyQt5", "phonenumbers", "pandas", "openpyxl", "requests"]
        for package in packages:
            if not self.is_python_package_installed(package):
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                    logging.info(f"Successfully installed {package}.")
                except subprocess.CalledProcessError as e:
                    logging.error(f"Failed to install {package}: {e}")

    def _download_file(self, url, destination):
        try:
            response = self.session.get(url, stream=True, timeout=30)
            response.raise_for_status()
            with open(destination, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
        except Exception as e:
            logging.error(f"Download failed: {e}")
            return False

    def _extract_archive(self, file_path, target_dir):
        try:
            if file_path.endswith(".zip"):
                with zipfile.ZipFile(file_path, "r") as zip_ref:
                    zip_ref.extractall(target_dir)
            elif file_path.endswith(".tar.gz"):
                with tarfile.open(file_path, "r:gz") as tar_ref:
                    tar_ref.extractall(target_dir)
            return True
        except Exception as e:
            logging.error(f"Extraction failed: {e}")
            return False

    def _install_driver(self, driver_name, download_url, file_pattern):
        driver_path = os.path.join(self.driver_dir, driver_name)
        if os.path.exists(driver_path):
            logging.info(f"{driver_name} already installed")
            return True

        try:
            temp_file = os.path.join(self.driver_dir, f"temp_{driver_name}.zip")
            if not self._download_file(download_url, temp_file):
                return False

            if not self._extract_archive(temp_file, self.driver_dir):
                return False

            for root, dirs, files in os.walk(self.driver_dir):
                for file in files:
                    if file.lower().startswith(file_pattern):
                        if os.path.exists(driver_path):
                            os.remove(driver_path)
                        os.rename(os.path.join(root, file), driver_path)
                        break

            if self.system != "windows":
                os.chmod(driver_path, stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

            if os.path.exists(temp_file):
                os.remove(temp_file)
            logging.info(f"{driver_name} installed successfully")
            return True
        except Exception as e:
            logging.error(f"Installation failed: {e}")
            if os.path.exists(temp_file):
                os.remove(temp_file)
            return False

    def install_chromedriver(self):
        driver_name = "chromedriver.exe" if self.system == "windows" else "chromedriver"
        driver_path = os.path.join(self.driver_dir, driver_name)
        
        if os.path.exists(driver_path):
            # Optional: check if it works, but for now assume yes
            return True

        chrome_version = self._get_chrome_version()
        if not chrome_version: return False
        
        # HYBRID LOGIC: If version is Canary/Dev (high number), skip manual download
        try:
            major_version = int(chrome_version.split('.')[0])
            if major_version > 135:
                logging.warning("Bleeding edge Chrome detected. Using Selenium Manager fallback.")
                return True 
        except ValueError:
            pass
        
        platform = self._get_chrome_platform()
        if not platform: return False

        driver_url = f"https://storage.googleapis.com/chrome-for-testing-public/{chrome_version}/{platform}/chromedriver-{platform}.zip"
        return self._install_driver(driver_name, driver_url, "chromedriver")

    def install_geckodriver(self):
        if "firefox" not in self.browser_paths: return
        try:
            response = self.session.get("https://api.github.com/repos/mozilla/geckodriver/releases/latest")
            version = response.json()["tag_name"]
            os_map = {"windows": "win64", "linux": "linux64", "darwin": "macos"}
            extension = "zip" if self.system == "windows" else "tar.gz"
            driver_url = f"https://github.com/mozilla/geckodriver/releases/download/{version}/geckodriver-{version}-{os_map[self.system]}.{extension}"
            return self._install_driver("geckodriver.exe" if self.system == "windows" else "geckodriver", driver_url, "geckodriver")
        except Exception: return False

    def install_edgedriver(self):
        if "edge" not in self.browser_paths: return
        try:
            driver_url = "https://msedgedriver.azureedge.net/LATEST_STABLE"
            response = self.session.get(driver_url)
            version = response.text.strip()
            os_map = {"windows": "win64", "linux": "linux64", "darwin": "mac64"}
            driver_url = f"https://msedgedriver.azureedge.net/{version}/msedgedriver_{os_map[self.system]}.zip"
            return self._install_driver("msedgedriver.exe" if self.system == "windows" else "msedgedriver", driver_url, "msedgedriver")
        except Exception: return False

    def install_all(self):
        logging.info("Checking Python packages...")
        self.install_python_packages()
        logging.info("Checking Drivers (Chrome, Firefox, Edge)...")
        self.install_chromedriver()
        self.install_geckodriver()
        self.install_edgedriver()
        logging.info("Dependency check completed.")


# ------------------- Thread-Safe Signal Container -------------------
class ThreadSignals(QObject):
    update_sent = pyqtSignal(dict)
    finished = pyqtSignal()
    error_occurred = pyqtSignal(str)
    login_required = pyqtSignal()
    progress_update = pyqtSignal(int)


# ------------------- Sending Thread (Updated Logic) -------------------
class SendingThread(QThread):
    def __init__(self, parent, numbers, message, attached_file, browser, delay, driver_dir):
        super().__init__()
        self.parent = parent
        self.numbers = numbers
        self.message = message
        self.attached_file = attached_file
        self.browser = browser
        self.delay = delay
        self.driver_dir = driver_dir
        self.signals = ThreadSignals()
        self.driver = None
        self.results = []
        self.supported_files = ('.jpg', '.jpeg', '.png', '.pdf', '.docx', '.txt', '.zip')
        self.retry_count = 3

    def _validate_file(self):
        if self.attached_file:
            if not os.path.exists(self.attached_file):
                raise FileNotFoundError("Attached file not found")
        return True

    def _retry_operation(self, operation, max_retries=3):
        for attempt in range(max_retries):
            try:
                return operation()
            except WebDriverException as e:
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                else:
                    raise

    def run(self):
        try:
            self._validate_file()

            # Determine driver path
            driver_name = {
                "Chrome": "chromedriver",
                "Brave": "chromedriver",
                "Firefox": "geckodriver",
                "Edge": "msedgedriver"
            }[self.browser]
            
            driver_path = os.path.join(self.driver_dir, driver_name + (".exe" if os.name == "nt" else ""))
            
            # Fallback Logic: If manual driver missing, assume Selenium Manager
            if self.browser in ["Chrome", "Brave"] and not os.path.exists(driver_path):
                 driver_path = None
            
            options = self._get_browser_options()
            
            try:
                self.driver = self._create_driver(driver_path, options)
            except SessionNotCreatedException as e:
                # Last resort fallback for Chrome
                if driver_path and "Chrome" in self.browser:
                    logging.warning("Session failed. Trying Selenium Manager fallback.")
                    self.driver = self._create_driver(None, options)
                else:
                    raise e
            
            self._retry_operation(lambda: self.driver.get("https://web.whatsapp.com"))

            if self._check_login_required():
                self.signals.login_required.emit()
                WebDriverWait(self.driver, 300).until(EC.presence_of_element_located((By.ID, "side")))

            WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "side")))

            for index, number in enumerate(self.numbers):
                # --- Pause Logic ---
                while self.parent.is_paused:
                    time.sleep(1)
                    if not self.parent.is_sending: break
                
                if not self.parent.is_sending: break
                
                # --- Anti-ban Safety ---
                if index > 0 and index % 50 == 0:
                    time.sleep(random.uniform(60, 120))

                result = {"number": number, "status": "Failed", "reason": ""}
                try:
                    self._process_number(number, index)
                    result["status"] = "Sent"
                    result["reason"] = "Success"
                except Exception as e:
                    if self._check_if_sent_strict():
                        result["status"] = "Sent"
                        result["reason"] = "Success (Verified)"
                    elif "Invalid" in str(e):
                        result["status"] = "Invalid"
                        result["reason"] = "Number not on WhatsApp"
                    else:
                        result["reason"] = str(e)
                        logging.error(f"Error sending to {number}: {e}")
                        try: self.driver.save_screenshot(f"error_{number}_{time.time()}.png")
                        except: pass
                finally:
                    self._update_progress(index, number, result)
                    safe_delay = (self.delay / 1000.0) + random.uniform(0.5, 1.5)
                    time.sleep(safe_delay)

            self.signals.finished.emit()
        except Exception as e:
            self.signals.error_occurred.emit(str(e))
        finally:
            if self.driver:
                try: self.driver.quit()
                except: pass

    def _get_browser_options(self):
        options_map = {
            "Chrome": webdriver.ChromeOptions,
            "Brave": webdriver.ChromeOptions,
            "Firefox": webdriver.FirefoxOptions,
            "Edge": webdriver.EdgeOptions
        }
        options = options_map[self.browser]()
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        if self.browser in ["Chrome", "Brave", "Edge"]:
            if self.browser == "Chrome":
                 options.add_argument("--remote-debugging-pipe")
            
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--log-level=3")

            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            
            # Original profile path
            options.add_argument(f"user-data-dir={os.path.join(os.getcwd(), 'chrome_profile')}")

            if self.browser == "Brave":
                options.binary_location = self.parent.installer.browser_paths.get("brave")

        elif self.browser == "Firefox":
            firefox_profile = webdriver.FirefoxProfile()
            firefox_profile.set_preference("dom.webdriver.enabled", False)
            options.profile = firefox_profile

        return options

    def _create_driver(self, driver_path, options):
        if driver_path and os.path.exists(driver_path):
            service = Service(executable_path=driver_path)
        else:
            service = None

        driver_map = {
            "Chrome": webdriver.Chrome,
            "Brave": webdriver.Chrome,
            "Firefox": webdriver.Firefox,
            "Edge": webdriver.Edge
        }
        
        if service:
            driver = driver_map[self.browser](service=service, options=options)
        else:
            driver = driver_map[self.browser](options=options)
            
        driver.set_window_size(1280, 900)
        return driver

    def _check_login_required(self):
        try:
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//div[@data-testid="qrcode"]'))
            )
            return True
        except: return False

    def _process_number(self, number, index):
        encoded_number = urllib.parse.quote(''.join(filter(str.isdigit, str(number))), safe='')
        encoded_message = urllib.parse.quote(self.message)
        url = f"https://web.whatsapp.com/send?phone={encoded_number}&text={encoded_message}"
        
        self.driver.get(url)
        
        try: WebDriverWait(self.driver, 3).until(EC.alert_is_present()).accept()
        except: pass
        
        # Wait for Footer Message Box (Precision Targeting)
        try:
            WebDriverWait(self.driver, 40).until(
                lambda d: d.find_elements(By.XPATH, '//footer//div[@contenteditable="true"]') or 
                          d.find_elements(By.XPATH, '//span[@data-icon="send"]') or
                          d.find_elements(By.XPATH, '//div[contains(text(), "url is invalid")]')
            )
        except TimeoutException:
            if self._check_if_sent_strict(): return
            raise Exception("Chat load timeout")

        if self.driver.find_elements(By.XPATH, '//div[contains(text(), "url is invalid")]'):
            raise Exception("Invalid Phone Number")

        if self.attached_file:
            self._handle_attachments()

        # === SEND STRATEGY: Focus -> Enter -> JS -> Manual ===
        self._target_and_send()
        
        self._verify_delivery()

    def _target_and_send(self):
        """
        Attempts to send using multiple strategies.
        Allows manual user intervention.
        """
        try:
            # 1. Target Footer Input specifically
            input_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//footer//div[@contenteditable="true"][@role="textbox"]'))
            )
            input_box.click() 
            time.sleep(0.5)
            
            # 2. Press ENTER
            input_box.send_keys(Keys.ENTER)
            time.sleep(1)
            
            # 3. Check Immediate Success
            if self._check_if_sent_strict():
                logging.info("Sent via Enter")
                return

            # 4. Fallback: JS Click
            if self.driver.find_elements(By.XPATH, '//span[@data-icon="send"]'):
                send_btn = self.driver.find_element(By.XPATH, '//span[@data-icon="send"]')
                parent = send_btn.find_element(By.XPATH, "./ancestor::button")
                self.driver.execute_script("arguments[0].click();", parent)
                
                time.sleep(0.5)
                if self._check_if_sent_strict():
                    return

            # 5. Manual Polling (Waiting for user)
            logging.info("Waiting for manual send...")
            for _ in range(10): # 5 seconds to press manually
                if self._check_if_sent_strict():
                    logging.info("User sent manually")
                    return
                time.sleep(0.5)

        except Exception as e:
            logging.warning(f"Send sequence warning: {e}")

    def _check_if_sent_strict(self):
        """Checks for actual message bubble."""
        try:
            # Check if 'message-out' bubble exists (Proof of sent message)
            if self.driver.find_elements(By.XPATH, '//div[contains(@class, "message-out")]'):
                return True
            # Double checks
            if self.driver.find_elements(By.XPATH, '//span[@data-icon="msg-dblcheck"] | //span[@data-icon="msg-check"]'):
                return True
            return False
        except: return False

    def _handle_attachments(self):
        try:
            attach_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@title="Attach"] | //span[@data-icon="clip"]'))
            )
            attach_btn.click()
            
            file_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
            )
            file_input.send_keys(self.attached_file)
            
            send_preview = WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            send_preview.click()
            time.sleep(2)
        except Exception as e:
            logging.error(f"Attachment error: {e}")

    def _verify_delivery(self):
        try:
            if self.driver.find_elements(By.XPATH, '//span[@data-icon="alert-error"]'):
                raise Exception("Message failed (Red Alert)")
        except: pass

    def _update_progress(self, index, number, result):
        self.results.append(result)
        self.signals.update_sent.emit({
            "sent": index + 1,
            "total": len(self.numbers),
            "current": number
        })
        self.signals.progress_update.emit(int((index + 1) / len(self.numbers) * 100))


# ------------------- Main Window -------------------
class WhatsAppSenderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings_file = "settings.json"
        self.installer = DependencyInstaller()
        self.driver_dir = self.installer.driver_dir
        self.load_settings()
        self.setWindowTitle("WhatsApp Message Sender")
        self.setGeometry(300, 200, 900, 600)
        self.sent_count = 0
        self.remaining_numbers = []
        self.is_sending = False
        self.is_paused = False
        self.attached_file = None
        pygame.mixer.init()
        
        self.installer.install_python_packages()
        
        # Initialize GUI
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout(main_widget)
        
        self.initUI()
        self.retranslate_ui()
        self.update_numbers_count()

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, "r") as f:
                settings = json.load(f)
                self.language = settings.get("language", "English")
                self.theme = settings.get("theme", "Light")
                self.browser = settings.get("browser", "Chrome")
                self.default_delay = settings.get("delay", 2000)
        else:
            self.language = "English"
            self.theme = "Light"
            self.browser = "Chrome"
            self.default_delay = 2000

    def save_settings(self):
        settings = {
            "language": self.language,
            "theme": self.theme,
            "browser": self.browser,
            "delay": self.default_delay
        }
        with open(self.settings_file, "w") as f:
            json.dump(settings, f)

    def initUI(self):
        # Menu Bar
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)

        settings_menu = QMenu("Settings", self)
        self.settings_menu = settings_menu
        menu_bar.addMenu(settings_menu)

        self.language_menu = QMenu("Language", self)
        settings_menu.addMenu(self.language_menu)
        self.language_menu.addAction(QAction("English", self, triggered=lambda: self.set_language("English")))
        self.language_menu.addAction(QAction("العربية", self, triggered=lambda: self.set_language("Arabic")))

        self.theme_menu = QMenu("Theme", self)
        settings_menu.addMenu(self.theme_menu)
        self.theme_menu.addAction(QAction("Light Mode", self, triggered=lambda: self.set_theme("Light")))
        self.theme_menu.addAction(QAction("Dark Mode", self, triggered=lambda: self.set_theme("Dark")))

        self.browser_menu = QMenu("Browser", self)
        settings_menu.addMenu(self.browser_menu)
        for b in ["Chrome", "Firefox", "Brave", "Edge"]:
             self.browser_menu.addAction(QAction(b, self, triggered=lambda checked, b=b: self.set_browser(b)))

        self.delay_action = QAction("Set Message Delay", self)
        self.delay_action.triggered.connect(self.set_message_delay)
        settings_menu.addAction(self.delay_action)

        # Phone
        phone_frame = QFrame()
        phone_frame.setStyleSheet("border: 1px solid gray; padding: 10px;")
        phone_layout = QVBoxLayout()
        phone_frame.setLayout(phone_layout)
        self.numbers_label = QLabel("Phone Numbers:")
        self.numbers_label.setFont(QFont("Arial", 12))
        self.numbers_label.setAlignment(Qt.AlignCenter)
        phone_layout.addWidget(self.numbers_label)
        self.numbers_input = QTextEdit()
        self.numbers_input.setFont(QFont("Arial", 11))
        self.numbers_input.textChanged.connect(self.update_numbers_count)
        phone_layout.addWidget(self.numbers_input)
        self.main_layout.addWidget(phone_frame)

        # Message
        message_frame = QFrame()
        message_frame.setStyleSheet("border: 1px solid gray; padding: 10px;")
        message_layout = QVBoxLayout()
        message_frame.setLayout(message_layout)
        self.message_label = QLabel("Message:")
        self.message_label.setFont(QFont("Arial", 12))
        self.message_label.setAlignment(Qt.AlignCenter)
        message_layout.addWidget(self.message_label)
        self.message_input = QTextEdit()
        self.message_input.setFont(QFont("Arial", 11))
        message_layout.addWidget(self.message_input)
        
        fmt_layout = QHBoxLayout()
        self.bold_btn = QPushButton("Bold")
        self.bold_btn.clicked.connect(lambda: self.format_text("bold"))
        self.italic_btn = QPushButton("Italic")
        self.italic_btn.clicked.connect(lambda: self.format_text("italic"))
        self.color_btn = QPushButton("Color")
        self.color_btn.clicked.connect(self.change_text_color)
        self.font_btn = QPushButton("Font Size")
        self.font_btn.clicked.connect(self.change_font_size)
        fmt_layout.addWidget(self.bold_btn); fmt_layout.addWidget(self.italic_btn); fmt_layout.addWidget(self.color_btn); fmt_layout.addWidget(self.font_btn)
        message_layout.addLayout(fmt_layout)
        self.main_layout.addWidget(message_frame)

        # Buttons
        buttons_layout = QHBoxLayout()
        self.import_button = QPushButton("Import Numbers")
        self.import_button.clicked.connect(self.import_numbers)
        self.send_button = QPushButton("Send Messages")
        self.send_button.clicked.connect(self.start_sending)
        self.stop_button = QPushButton("Stop Sending")
        self.stop_button.clicked.connect(self.stop_sending)
        self.resume_button = QPushButton("Resume Sending")
        self.resume_button.clicked.connect(self.resume_sending)
        self.attach_button = QPushButton("Attach File")
        self.attach_button.clicked.connect(self.attach_file)
        self.export_button = QPushButton("Export Report")
        self.export_button.clicked.connect(self.export_report)
        buttons_layout.addWidget(self.import_button); buttons_layout.addWidget(self.send_button); buttons_layout.addWidget(self.stop_button)
        buttons_layout.addWidget(self.resume_button); buttons_layout.addWidget(self.attach_button); buttons_layout.addWidget(self.export_button)
        self.main_layout.addLayout(buttons_layout)

        # Progress
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.main_layout.addWidget(self.progress_bar)

        # Stats
        stats_layout = QHBoxLayout()
        self.total_numbers_label = QLabel("Total: 0")
        self.sent_numbers_label = QLabel("Sent: 0")
        self.remaining_numbers_label = QLabel("Remaining: 0")
        stats_layout.addWidget(self.total_numbers_label); stats_layout.addWidget(self.sent_numbers_label); stats_layout.addWidget(self.remaining_numbers_label)
        self.main_layout.addLayout(stats_layout)

    # --- Logic ---

    def update_numbers_count(self):
        numbers = self.numbers_input.toPlainText().strip().split("\n")
        valid_numbers = [num for num in numbers if num.strip()]
        self.remaining_numbers = valid_numbers
        
        is_ar = self.language == "Arabic"
        self.total_numbers_label.setText(f"{'العدد الكلي' if is_ar else 'Total'}: {len(valid_numbers)}")
        self.remaining_numbers_label.setText(f"{'المتبقي' if is_ar else 'Remaining'}: {len(valid_numbers) - self.sent_count}")
        self.sent_numbers_label.setText(f"{'تم الإرسال' if is_ar else 'Sent'}: {self.sent_count}")

    def update_stats_text(self):
        self.update_numbers_count()

    def import_numbers(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Import", "", "Files (*.csv *.xlsx *.xls *.txt);;All Files (*)", options=options)
        if file_path:
            try:
                extracted_numbers = []
                pattern = r'\+?\d{7,15}'
                if file_path.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path, dtype=str, header=None)
                    extracted_numbers = re.findall(pattern, " ".join(df.values.flatten().astype(str)))
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path, dtype=str, header=None)
                    extracted_numbers = re.findall(pattern, " ".join(df.values.flatten().astype(str)))
                else:
                    with open(file_path, "r", encoding='utf-8', errors='ignore') as f:
                        extracted_numbers = re.findall(pattern, f.read())
                
                if extracted_numbers:
                    self.numbers_input.setPlainText("\n".join(list(set(extracted_numbers))))
                    self.update_numbers_count()
                    QMessageBox.information(self, "Success", f"Imported {len(set(extracted_numbers))} numbers")
                else: QMessageBox.warning(self, "Warning", "No numbers found")
            except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def start_sending(self):
        try:
            if not self.remaining_numbers:
                return QMessageBox.warning(self, "Warning", "Please enter numbers.")
            if not self.message_input.toPlainText().strip() and not self.attached_file:
                return QMessageBox.warning(self, "Warning", "Please enter a message.")
            
            if os.path.exists("start_sound.mp3"): 
                try: pygame.mixer.music.load("start_sound.mp3"); pygame.mixer.music.play()
                except: pass

            self.is_sending = True
            self.is_paused = False
            self.sent_count = 0
            self.progress_bar.setValue(0)

            self.sending_thread = SendingThread(
                self, self.remaining_numbers.copy(), self.message_input.toPlainText(),
                self.attached_file, self.browser, self.default_delay, self.installer.driver_dir
            )
            self.sending_thread.signals.update_sent.connect(self.update_sent_count)
            self.sending_thread.signals.finished.connect(self.sending_finished)
            self.sending_thread.signals.error_occurred.connect(self.show_error)
            self.sending_thread.signals.login_required.connect(self.show_login_required)
            self.sending_thread.signals.progress_update.connect(self.progress_bar.setValue)
            self.sending_thread.start()
        except Exception as e:
            QMessageBox.critical(self, "Critical Error", f"Failed to start: {str(e)}")

    def stop_sending(self):
        if self.is_sending:
            self.is_paused = True
            QMessageBox.information(self, "Paused", "Sending Paused")

    def resume_sending(self):
        if self.is_sending:
            self.is_paused = False
            QMessageBox.information(self, "Resumed", "Sending Resumed")

    def attach_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Attach", "")
        if path:
            self.attached_file = path
            self.attach_button.setText(f"📎 {os.path.basename(path)}")

    def export_report(self):
        if not hasattr(self, 'sending_thread') or not self.sending_thread.results:
            return QMessageBox.warning(self, "Warning", "No Data")
        path, _ = QFileDialog.getSaveFileName(self, "Export", "", "Excel (*.xlsx)")
        if path:
            try:
                wb = xlsxwriter.Workbook(path)
                ws = wb.add_worksheet()
                headers = ["Number", "Status", "Reason"]
                for i, h in enumerate(headers): ws.write(0, i, h)
                for i, r in enumerate(self.sending_thread.results, 1):
                    ws.write(i, 0, r['number']); ws.write(i, 1, r['status']); ws.write(i, 2, r['reason'])
                wb.close()
                QMessageBox.information(self, "Success", "Exported")
            except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def play_sound(self, f): pass
    def format_text(self, s):
        c = self.message_input.textCursor()
        if c.hasSelection(): 
            t = c.selectedText()
            c.insertText(f"*{t}*" if s=="bold" else f"_{t}_")
    def change_text_color(self):
        c = QColorDialog.getColor()
        if c.isValid(): self.message_input.setTextColor(c)
    def change_font_size(self):
        f, ok = QFontDialog.getFont()
        if ok: self.message_input.setFont(f)
    def set_language(self, l):
        self.language = l; self.save_settings(); self.retranslate_ui()
    def set_theme(self, t):
        self.theme = t; self.save_settings()
        self.setStyleSheet("background-color: #333; color: white; QTextEdit { background-color: #444; }" if t=="Dark" else "")
    def set_browser(self, b):
        self.browser = b; self.save_settings(); QMessageBox.information(self, "Browser", f"Set to {b}")
    def set_message_delay(self):
        v, ok = QInputDialog.getInt(self, "Delay", "MS:", self.default_delay, 500, 60000)
        if ok: self.default_delay = v; self.save_settings()
    def update_sent_count(self, d):
        self.sent_count = d['sent']; self.update_numbers_count()
    def sending_finished(self):
        self.is_sending = False; self.is_paused = False; QMessageBox.information(self, "Done", "Finished")
    def show_error(self, m):
        self.is_sending = False; QMessageBox.critical(self, "Error", m)
    def show_login_required(self):
        QMessageBox.warning(self, "Login", "Scan QR")

    def retranslate_ui(self):
        is_ar = self.language == "Arabic"
        self.numbers_label.setText("أرقام الهواتف:" if is_ar else "Phone Numbers:")
        self.message_label.setText("الرسالة:" if is_ar else "Message:")
        self.import_button.setText("استيراد" if is_ar else "Import Numbers")
        self.send_button.setText("إرسال" if is_ar else "Send Messages")
        self.stop_button.setText("إيقاف مؤقت" if is_ar else "Stop Sending")
        self.resume_button.setText("استئناف" if is_ar else "Resume Sending")
        self.attach_button.setText("إرفاق" if is_ar else "Attach File")
        self.export_button.setText("تصدير" if is_ar else "Export Report")
        self.bold_btn.setText("عريض" if is_ar else "Bold")
        self.italic_btn.setText("مائل" if is_ar else "Italic")
        self.color_btn.setText("لون" if is_ar else "Color")
        self.font_btn.setText("حجم" if is_ar else "Font Size")
        self.settings_menu.setTitle("إعدادات" if is_ar else "Settings")
        self.language_menu.setTitle("اللغة" if is_ar else "Language")
        self.theme_menu.setTitle("المظهر" if is_ar else "Theme")
        self.browser_menu.setTitle("المتصفح" if is_ar else "Browser")
        self.delay_action.setText("تأخير" if is_ar else "Set Message Delay")
        
        if is_ar: self.setLayoutDirection(Qt.RightToLeft)
        else: self.setLayoutDirection(Qt.LeftToRight)
        self.update_numbers_count()

    def closeEvent(self, event):
        self.save_settings()
        if hasattr(self, 'sending_thread'): self.sending_thread.quit()
        event.accept()

if __name__ == "__main__":
    installer = DependencyInstaller()
    installer.install_chromedriver()
    installer.install_geckodriver()
    installer.install_edgedriver()
    app = QApplication(sys.argv)
    window = WhatsAppSenderApp()
    window.show()

    sys.exit(app.exec_())
