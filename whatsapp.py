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
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QFileDialog, QWidget,
    QMessageBox, QFrame, QMenuBar, QMenu, QAction,
    QColorDialog, QFontDialog, QInputDialog, QProgressBar
)
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException
import phonenumbers

# ------------------- Configuration -------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# ------------------- Dependency Installer -------------------
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
        """Retrieves the installed Chrome version."""
        chrome_path = self.browser_paths.get("chrome")
        if not chrome_path:
            logging.error("Chrome browser path not found.")
            return None

        try:
            if self.system == "windows":
                # Try registry first
                try:
                    import winreg
                    reg_path = r'SOFTWARE\Google\Chrome\BLBeacon'
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                        version, _ = winreg.QueryValueEx(key, 'version')
                        return version
                except Exception as e:
                    logging.warning(f"Registry read failed: {e}. Trying file version...")
                    # Fallback to PowerShell command
                    command = f'(Get-Item "{chrome_path}").VersionInfo.FileVersion'
                    result = subprocess.run(["powershell", "-Command", command], 
                                          capture_output=True, text=True, check=True)
                    return result.stdout.strip()
            elif self.system == "linux":
                result = subprocess.run([chrome_path, "--version"], 
                                      capture_output=True, text=True, check=True)
                return result.stdout.strip().split()[-1]
            elif self.system == "darwin":
                # Check Info.plist
                plist_path = os.path.join(os.path.dirname(chrome_path), '..', 'Info.plist')
                plist_path = os.path.abspath(plist_path)
                with open(plist_path, 'rb') as f:
                    content = f.read().decode('utf-8', errors='ignore')
                    match = re.search(r'<key>CFBundleShortVersionString</key>\s*<string>([\d.]+)</string>', content)
                    if match:
                        return match.group(1)
                # Fallback to mdls
                result = subprocess.run(['mdls', '-name', 'kMDItemVersion', chrome_path], 
                                      capture_output=True, text=True, check=True)
                return result.stdout.split('"')[1]
            else:
                return None
        except Exception as e:
            logging.error(f"Error getting Chrome version: {e}")
            return None

    def _get_chrome_platform(self):
        """Determines platform string for ChromeDriver download."""
        if self.system == "windows":
            return "win64" if self.arch == "amd64" else "win32"
        elif self.system == "linux":
            return "linux64" if self.arch == "x86_64" else "linux32"
        elif self.system == "darwin":
            return "mac-arm64" if self.arch == "arm64" else "mac-x64"
        else:
            return None

    def is_python_package_installed(self, package_name):
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "show", package_name],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except subprocess.CalledProcessError:
            return False

    def install_python_packages(self):
        packages = ["selenium", "pygame", "xlsxwriter", "PyQt5", "phonenumbers"]
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
            logging.info(f"Downloaded {os.path.basename(destination)}")
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

            # Handle nested directories in archives
            for root, dirs, files in os.walk(self.driver_dir):
                for file in files:
                    if file.lower().startswith(file_pattern):
                        os.rename(os.path.join(root, file), driver_path)
                        break

            if self.system != "windows":
                os.chmod(driver_path, stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

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
        
        # Remove existing driver if outdated
        if os.path.exists(driver_path):
            try:
                version_output = subprocess.check_output([driver_path, "--version"]).decode()
                if "ChromeDriver 1" in version_output:
                    os.remove(driver_path)
            except:
                pass

        if os.path.exists(driver_path):
            logging.info("ChromeDriver already installed")
            return True

        chrome_version = self._get_chrome_version()
        if not chrome_version:
            logging.error("Could not detect Chrome version.")
            return False

        platform = self._get_chrome_platform()
        if not platform:
            logging.error("Unsupported platform.")
            return False

        driver_url = f"https://storage.googleapis.com/chrome-for-testing-public/{chrome_version}/{platform}/chromedriver-{platform}.zip"
        logging.info(f"Downloading ChromeDriver {chrome_version} for {platform}")

        return self._install_driver(driver_name, driver_url, "chromedriver")

    def install_geckodriver(self):
        if "firefox" not in self.browser_paths:
            logging.warning("Firefox not found, skipping GeckoDriver installation")
            return

        try:
            response = self.session.get(
                "https://api.github.com/repos/mozilla/geckodriver/releases/latest"
            )
            response.raise_for_status()
            version = response.json()["tag_name"]

            os_map = {
                "windows": "win64",
                "linux": "linux64",
                "darwin": "macos"
            }
            extension = "zip" if self.system == "windows" else "tar.gz"
            driver_url = f"https://github.com/mozilla/geckodriver/releases/download/{version}/geckodriver-{version}-{os_map[self.system]}.{extension}"

            return self._install_driver(
                "geckodriver.exe" if self.system == "windows" else "geckodriver",
                driver_url,
                "geckodriver"
            )
        except Exception as e:
            logging.error(f"GeckoDriver installation failed: {e}")
            return False

    def install_edgedriver(self):
        if "edge" not in self.browser_paths:
            logging.warning("Edge not found, skipping EdgeDriver installation")
            return

        try:
            # Use direct latest stable version URL
            driver_url = "https://msedgedriver.azureedge.net/LATEST_STABLE"
            response = self.session.get(driver_url)
            response.raise_for_status()
            version = response.text.strip()

            os_map = {
                "windows": "win64",
                "linux": "linux64",
                "darwin": "mac64"
            }
            driver_url = f"https://msedgedriver.azureedge.net/{version}/edgedriver_{os_map[self.system]}.zip"

            return self._install_driver(
                "msedgedriver.exe" if self.system == "windows" else "msedgedriver",
                driver_url,
                "msedgedriver"
            )
        except Exception as e:
            logging.error(f"EdgeDriver installation failed: {e}")
            return False

    def install_all(self):
        logging.info("Checking Python packages...")
        self.install_python_packages()
        logging.info("Checking ChromeDriver...")
        self.install_chromedriver()
        logging.info("Checking GeckoDriver...")
        self.install_geckodriver()
        logging.info("Checking EdgeDriver...")
        self.install_edgedriver()
        logging.info("Dependency check completed")


## ------------------- Thread-Safe Signal Container -------------------
class ThreadSignals(QObject):
    update_sent = pyqtSignal(dict)
    finished = pyqtSignal()
    error_occurred = pyqtSignal(str)
    login_required = pyqtSignal()
    progress_update = pyqtSignal(int)


# ------------------- Sending Thread -------------------
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
            if not self.attached_file.lower().endswith(self.supported_files):
                raise ValueError(f"Unsupported file type: {os.path.splitext(self.attached_file)[1]}")
        return True

    def _validate_numbers(self):
        for number in self.numbers:
            try:
                parsed_number = phonenumbers.parse(number, None)
                if not phonenumbers.is_valid_number(parsed_number):
                    raise ValueError(f"Invalid phone number: {number}")
            except phonenumbers.phonenumberutil.NumberParseException:
                raise ValueError(f"Invalid phone number format: {number}")

    def _ensure_element_ready(self, locator, timeout=15):
        element = WebDriverWait(self.driver, timeout).until(
            EC.visibility_of_element_located(locator)
        )
        WebDriverWait(self.driver, timeout).until(
            lambda d: element.is_enabled()
        )
        return element

    def _safe_clear_input(self, element):
        for _ in range(3):
            element.send_keys(Keys.BACKSPACE)
        self.driver.execute_script("arguments[0].value = '';", element)
        element.send_keys(' ')
        element.send_keys(Keys.BACKSPACE)
        time.sleep(0.5)

    def _handle_popups(self):
        try:
            # Handle "Your computer is..." notification
            notification = WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Your computer is")]'))
            )
            close_btn = notification.find_element(By.XPATH, './following-sibling::div')
            close_btn.click()
            time.sleep(1)
        except:
            pass

        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//div[@role="dialog"]'))
            )
            close_buttons = self.driver.find_elements(
                By.XPATH, '//div[@role="dialog"]//button[@aria-label="Close"]'
            )
            if close_buttons:
                close_buttons[0].click()
                time.sleep(1)
        except:
            pass

    def _retry_operation(self, operation, max_retries=3):
        for attempt in range(max_retries):
            try:
                return operation()
            except WebDriverException as e:
                if attempt < max_retries - 1:
                    sleep_time = 2 ** attempt
                    logging.warning(f"Retrying in {sleep_time}s... ({str(e)})")
                    time.sleep(sleep_time)
                else:
                    raise

    def run(self):
        try:
            self._validate_file()
            self._validate_numbers()

            driver_name = {
                "Chrome": "chromedriver",
                "Brave": "chromedriver",
                "Firefox": "geckodriver",
                "Edge": "msedgedriver"
            }[self.browser]
            
            driver_path = os.path.join(self.driver_dir, driver_name + (".exe" if os.name == "nt" else ""))
            
            if not os.path.exists(driver_path):
                raise FileNotFoundError(f"Driver not found: {driver_path}")

            options = self._get_browser_options()
            self.driver = self._create_driver(driver_path, options)
            
            self._retry_operation(
                lambda: self.driver.get("https://web.whatsapp.com")
            )

            if self._check_login_required():
                self.signals.login_required.emit()
                return

            WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "side")))

            for index, number in enumerate(self.numbers):
                if not self.parent.is_sending:
                    break

                result = {"number": number, "status": "Failed", "reason": ""}
                try:
                    self._process_number(number, index)
                    result["status"] = "Success"
                except Exception as e:
                    result["reason"] = str(e)
                    logging.error(f"Error sending to {number}: {e}")
                    self.driver.save_screenshot(f"error_{number}_{time.time()}.png")
                    self.signals.error_occurred.emit(str(e))
                finally:
                    self._update_progress(index, number, result)

            self.signals.finished.emit()
        except Exception as e:
            self.signals.error_occurred.emit(str(e))
        finally:
            if self.driver:
                self.driver.quit()

    def _get_browser_options(self):
        options_map = {
            "Chrome": webdriver.ChromeOptions,
            "Brave": webdriver.ChromeOptions,
            "Firefox": webdriver.FirefoxOptions,
            "Edge": webdriver.EdgeOptions
        }
        options = options_map[self.browser]()

        # إعدادات مشتركة للمتصفحات
        options.add_argument("--disable-blink-features=AutomationControlled")

        # إعدادات خاصة بمتصفحات Chromium (Chrome, Brave, Edge)
        if self.browser in ["Chrome", "Brave", "Edge"]:
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            if self.browser == "Brave":
                options.binary_location = self.parent.installer.browser_paths.get("brave")
            if self.browser in ["Chrome", "Brave"]:
                options.add_argument(f"user-data-dir={os.path.join(os.getcwd(), 'chrome_profile')}")

        # إعدادات خاصة بـ Firefox
        elif self.browser == "Firefox":
            # إعدادات Firefox
            firefox_profile = webdriver.FirefoxProfile()
            firefox_profile.set_preference("dom.webdriver.enabled", False)
            firefox_profile.set_preference("useAutomationExtension", False)
            options.profile = firefox_profile

        return options

    def _create_driver(self, driver_path, options):
        service = Service(executable_path=driver_path)
        driver_map = {
            "Chrome": webdriver.Chrome,
            "Brave": webdriver.Chrome,
            "Firefox": webdriver.Firefox,
            "Edge": webdriver.Edge
        }
        driver = driver_map[self.browser](service=service, options=options)
        driver.set_window_size(1440, 900)  # Force window size
        return driver

    def _check_login_required(self):
        try:
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//div[@data-testid="qrcode"]'))
            )
            return True
        except:
            return False

    def _process_number(self, number, index):
        encoded_number = urllib.parse.quote(number, safe='')
        self._retry_operation(
            lambda: self.driver.get(f"https://web.whatsapp.com/send?phone={encoded_number}")
        )
        
        # Wait for initial load
        time.sleep(2)
        
        # Handle "Use WhatsApp Web" popup
        try:
            continue_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@role="button" and contains(text(), "use WhatsApp Web")]'))
            )
            continue_button.click()
            time.sleep(1)
        except:
            pass

        self._handle_popups()
        self._wait_for_chat_load()
        
        # Additional stability check
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        
        self._send_message()
        self._handle_attachments()
        self._send_with_retry()
        self._verify_delivery()
        time.sleep(random.uniform(2, 5))

    def _wait_for_chat_load(self):
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//div[@data-testid="conversation-panel-body"]'))
        )
        # New message box locator
        WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "copyable-text") and @role="textbox"]'))
        )

    def _send_message(self):
        # Updated message box selector
        message_box = self._ensure_element_ready((By.XPATH, '//div[contains(@class, "copyable-text") and @role="textbox"]'))
        self._safe_clear_input(message_box)
        message_box.send_keys(self.message)

    def _handle_attachments(self):
        if self.attached_file:
            self._retry_operation(self._attach_file)

    def _attach_file(self):
        attachment_button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@title="Attach"]'))
        )
        attachment_button.click()
        
        file_input = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@accept="*"]'))
        )
        file_input.send_keys(self.attached_file)
        
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@data-testid="media-attach-preview"]'))
        )
        time.sleep(2)

    def _send_with_retry(self):
        for attempt in range(self.retry_count):
            try:
                # Updated send button selector
                send_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[contains(@data-testid,"send") and @aria-label="Send"]'))
                )
                send_button.click()
                return
            except:
                if attempt == self.retry_count - 1:
                    raise
                time.sleep(1)

    def _verify_delivery(self):
        try:
            WebDriverWait(self.driver, 15).until(
                lambda d: d.find_element(By.XPATH, '//span[@data-testid="msg-time"]') and 
                        d.find_element(By.XPATH, '//span[@data-icon="msg-dblcheck"]')
            )
        except Exception as e:
            # Check for error message
            error_msg = self.driver.find_elements(By.XPATH, '//div[contains(text(), "couldn\'t send")]')
            if error_msg:
                raise Exception("Message failed to send: " + error_msg[0].text)
            # Fallback verification
            if not self.driver.find_elements(By.XPATH, '//div[@data-testid="msg-container"]'):
                raise Exception("Message verification failed")

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
        self.attached_file = None
        pygame.mixer.init()
        self.initUI()
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

        # Settings Menu
        settings_menu = QMenu("Settings", self)
        menu_bar.addMenu(settings_menu)

        # Language Menu
        language_menu = QMenu("Language", self)
        settings_menu.addMenu(language_menu)

        language_menu.addAction(QAction("English", self, triggered=lambda: self.set_language("English")))
        language_menu.addAction(QAction("Arabic", self, triggered=lambda: self.set_language("Arabic")))

        # Theme Menu
        theme_menu = QMenu("Theme", self)
        settings_menu.addMenu(theme_menu)

        theme_menu.addAction(QAction("Light Mode", self, triggered=lambda: self.set_theme("Light")))
        theme_menu.addAction(QAction("Dark Mode", self, triggered=lambda: self.set_theme("Dark")))

        # Browser Selection
        browser_menu = QMenu("Browser", self)
        settings_menu.addMenu(browser_menu)

        browser_menu.addAction(QAction("Chrome", self, triggered=lambda: self.set_browser("Chrome")))
        browser_menu.addAction(QAction("Firefox", self, triggered=lambda: self.set_browser("Firefox")))
        browser_menu.addAction(QAction("Brave", self, triggered=lambda: self.set_browser("Brave")))
        browser_menu.addAction(QAction("Edge", self, triggered=lambda: self.set_browser("Edge")))

        # Delay Setting
        delay_action = QAction("Set Message Delay", self)
        delay_action.triggered.connect(self.set_message_delay)
        settings_menu.addAction(delay_action)

        # Main Layout
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()

        # Phone Numbers Section
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
        self.numbers_input.setPlaceholderText("Enter phone numbers (one per line) or import from file...")
        self.numbers_input.textChanged.connect(self.update_numbers_count)
        phone_layout.addWidget(self.numbers_input)

        main_layout.addWidget(phone_frame)

        # Message Section
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
        self.message_input.setPlaceholderText("Enter your message here...")
        message_layout.addWidget(self.message_input)

        # Formatting Buttons
        formatting_buttons_layout = QHBoxLayout()

        bold_button = QPushButton("Bold")
        bold_button.clicked.connect(lambda: self.format_text("bold"))
        formatting_buttons_layout.addWidget(bold_button)

        italic_button = QPushButton("Italic")
        italic_button.clicked.connect(lambda: self.format_text("italic"))
        formatting_buttons_layout.addWidget(italic_button)

        color_button = QPushButton("Color")
        color_button.clicked.connect(self.change_text_color)
        formatting_buttons_layout.addWidget(color_button)

        font_button = QPushButton("Font Size")
        font_button.clicked.connect(self.change_font_size)
        formatting_buttons_layout.addWidget(font_button)

        message_layout.addLayout(formatting_buttons_layout)
        main_layout.addWidget(message_frame)

        # Control Buttons
        buttons_layout = QHBoxLayout()

        self.import_button = QPushButton("Import Numbers")
        self.import_button.clicked.connect(self.import_numbers)
        buttons_layout.addWidget(self.import_button)

        self.send_button = QPushButton("Send Messages")
        self.send_button.clicked.connect(self.start_sending)
        buttons_layout.addWidget(self.send_button)

        self.stop_button = QPushButton("Stop Sending")
        self.stop_button.clicked.connect(self.stop_sending)
        buttons_layout.addWidget(self.stop_button)

        self.resume_button = QPushButton("Resume Sending")
        self.resume_button.clicked.connect(self.resume_sending)
        buttons_layout.addWidget(self.resume_button)

        self.attach_button = QPushButton("Attach File")
        self.attach_button.clicked.connect(self.attach_file)
        buttons_layout.addWidget(self.attach_button)

        self.export_button = QPushButton("Export Report")
        self.export_button.clicked.connect(self.export_report)
        buttons_layout.addWidget(self.export_button)

        main_layout.addLayout(buttons_layout)

        # Progress Bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.progress_bar)

        # Statistics Section
        stats_layout = QHBoxLayout()

        self.total_numbers_label = QLabel("Total Numbers: 0")
        self.total_numbers_label.setFont(QFont("Arial", 10))
        stats_layout.addWidget(self.total_numbers_label)

        self.sent_numbers_label = QLabel("Sent: 0")
        self.sent_numbers_label.setFont(QFont("Arial", 10))
        stats_layout.addWidget(self.sent_numbers_label)

        self.remaining_numbers_label = QLabel("Remaining: 0")
        self.remaining_numbers_label.setFont(QFont("Arial", 10))
        stats_layout.addWidget(self.remaining_numbers_label)

        main_layout.addLayout(stats_layout)
        main_widget.setLayout(main_layout)

    def update_numbers_count(self):
        numbers = self.numbers_input.toPlainText().strip().split("\n")
        valid_numbers = [num for num in numbers if num.strip()]
        self.remaining_numbers = valid_numbers
        self.total_numbers_label.setText(f"Total Numbers: {len(valid_numbers)}")
        self.remaining_numbers_label.setText(f"Remaining: {len(valid_numbers)}")

    def import_numbers(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import Numbers", "", "Supported Files (*.csv *.xlsx *.txt);;All Files (*)", options=options
        )
        if file_path:
            try:
                with open(file_path, "r") as file:
                    numbers = file.read()
                    self.numbers_input.setPlainText(numbers)
                    self.update_numbers_count()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to import numbers: {e}")

    def start_sending(self):
        if not self.remaining_numbers:
            QMessageBox.warning(self, "No Numbers", "Please enter or import phone numbers.")
            return

        if not self.message_input.toPlainText().strip():
            QMessageBox.warning(self, "No Message", "Please enter a message.")
            return

        self.play_sound("start_sound.mp3")
        self.is_sending = True
        self.sent_count = 0
        self.progress_bar.setValue(0)

        self.sending_thread = SendingThread(
            self,
            self.remaining_numbers.copy(),
            self.message_input.toPlainText(),
            self.attached_file,
            self.browser,
            self.default_delay,
            self.driver_dir
        )
        self.sending_thread.signals.update_sent.connect(self.update_sent_count)
        self.sending_thread.signals.finished.connect(self.sending_finished)
        self.sending_thread.signals.error_occurred.connect(self.show_error)
        self.sending_thread.signals.login_required.connect(self.show_login_required)
        self.sending_thread.signals.progress_update.connect(self.progress_bar.setValue)
        self.sending_thread.start()

    def stop_sending(self):
        self.is_sending = False
        QMessageBox.warning(self, "Sending Stopped", "Message sending has been stopped!")

    def resume_sending(self):
        if not self.is_sending:
            self.is_sending = True
            QMessageBox.information(self, "Sending Resumed", "Message sending has resumed!")

    def attach_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Attach File", "", "Supported Files (*.jpg *.png *.pdf *.zip *.docx);;All Files (*)", options=options
        )
        if file_path:
            self.attached_file = file_path
            QMessageBox.information(self, "File Attached", f"File attached successfully: {file_path}")

    def export_report(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Report", "", "Excel Files (*.xlsx)")
        if file_path and hasattr(self, 'sending_thread'):
            workbook = xlsxwriter.Workbook(file_path)
            worksheet = workbook.add_worksheet()
            
            headers = ["Phone Number", "Status", "Reason"]
            for col, header in enumerate(headers):
                worksheet.write(0, col, header)
            
            for row, result in enumerate(self.sending_thread.results, start=1):
                worksheet.write(row, 0, result["number"])
                worksheet.write(row, 1, result["status"])
                worksheet.write(row, 2, result["reason"])
            
            workbook.close()
            QMessageBox.information(self, "Report Exported", "Report has been exported successfully!")

    def play_sound(self, sound_file):
        if os.path.exists(sound_file):
            try:
                pygame.mixer.music.load(sound_file)
                pygame.mixer.music.play()
            except Exception as e:
                print(f"Error playing sound: {e}")
        else:
            print(f"Sound file not found: {sound_file}")

    def format_text(self, style):
        cursor = self.message_input.textCursor()
        if style == "bold":
            cursor.insertText(f"*{cursor.selectedText()}*")
        elif style == "italic":
            cursor.insertText(f"_{cursor.selectedText()}_")

    def change_text_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.message_input.setTextColor(color)

    def change_font_size(self):
        font, ok = QFontDialog.getFont()
        if ok:
            self.message_input.setFont(font)

    def set_language(self, language):
        self.language = language
        self.retranslate_ui()
        QMessageBox.information(self, self.tr("Language Changed"), f"{self.tr('Language set to')} {language}!")
        self.save_settings()

    def retranslate_ui(self):
        if self.language == "Arabic":
            self.setWindowTitle("مرسل رسائل الواتساب")
            self.numbers_label.setText("أرقام الهواتف:")
            self.numbers_input.setPlaceholderText("أدخل أرقام الهواتف (رقم في كل سطر) أو استورد من ملف...")
            self.message_label.setText("الرسالة:")
            self.message_input.setPlaceholderText("أدخل رسالتك هنا...")
            self.import_button.setText("استيراد الأرقام")
            self.send_button.setText("إرسال الرسائل")
            self.stop_button.setText("إيقاف الإرسال")
            self.resume_button.setText("استئناف الإرسال")
            self.attach_button.setText("إرفاق ملف")
            self.export_button.setText("تصدير التقرير")
        else:
            self.setWindowTitle("WhatsApp Message Sender")
            self.numbers_label.setText("Phone Numbers:")
            self.numbers_input.setPlaceholderText("Enter phone numbers (one per line) or import from file...")
            self.message_label.setText("Message:")
            self.message_input.setPlaceholderText("Enter your message here...")
            self.import_button.setText("Import Numbers")
            self.send_button.setText("Send Messages")
            self.stop_button.setText("Stop Sending")
            self.resume_button.setText("Resume Sending")
            self.attach_button.setText("Attach File")
            self.export_button.setText("Export Report")

    def set_theme(self, theme):
        self.theme = theme
        if theme == "Light":
            self.setStyleSheet("")
        else:
            self.setStyleSheet("""
                background-color: #2E2E2E;
                color: white;
                QLabel, QPushButton, QTextEdit, QFrame {
                    color: white;
                }
                QTextEdit {
                    background-color: #3E3E3E;
                }
                QPushButton {
                    background-color: #505050;
                    border: 1px solid #606060;
                }
                QPushButton:hover {
                    background-color: #606060;
                }
            """)
        self.save_settings()

    def set_browser(self, browser):
        self.browser = browser
        QMessageBox.information(self, "Browser Changed", f"Browser set to {browser}!")
        self.save_settings()

    def set_message_delay(self):
        delay, ok = QInputDialog.getInt(self, "Set Message Delay", "Enter delay in milliseconds:", self.default_delay, 500, 10000)
        if ok:
            self.default_delay = delay
            QMessageBox.information(self, "Delay Set", f"Message delay set to {self.default_delay} ms")
            self.save_settings()

    def update_sent_count(self):
        self.sent_count += 1
        self.sent_numbers_label.setText(f"Sent: {self.sent_count}")
        self.remaining_numbers_label.setText(f"Remaining: {len(self.remaining_numbers) - self.sent_count}")

    def sending_finished(self):
        QMessageBox.information(self, "Sending Finished", "All messages have been sent!")
        self.is_sending = False

    def show_error(self, error_msg):
        QMessageBox.critical(self, "Error", f"Failed to send messages: {error_msg}")
        self.is_sending = False

    def show_login_required(self):
        QMessageBox.warning(self, "Login Required", "Please scan the QR code to log in to WhatsApp Web.")
        self.is_sending = False

    def closeEvent(self, event):
        self.save_settings()
        if hasattr(self, 'sending_thread') and self.sending_thread.isRunning():
            self.is_sending = False
            self.sending_thread.quit()
            self.sending_thread.wait(5000)
        event.accept()

if __name__ == "__main__":
    # Check drivers before creating the application
    installer = DependencyInstaller()
    required_drivers = {
        "Chrome": "chromedriver.exe" if os.name == "nt" else "chromedriver",
        "Firefox": "geckodriver.exe" if os.name == "nt" else "geckodriver",
        "Edge": "msedgedriver.exe" if os.name == "nt" else "msedgedriver"
    }
    
    # Install missing drivers
    installer.install_chromedriver()
    installer.install_geckodriver()
    installer.install_edgedriver()
    
    # Proceed to GUI
    app = QApplication(sys.argv)
    window = WhatsAppSenderApp()
    window.show()
    sys.exit(app.exec_())