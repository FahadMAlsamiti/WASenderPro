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
import shutil
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
    QColorDialog, QFontDialog, QInputDialog, QProgressBar,
    QDialog, QFormLayout, QSizePolicy
)
from PyQt5.QtGui import QFont, QColor, QCursor
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject, QUrl

# مكتبات السيلينيوم
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

# ------------------- Dependency Installer -------------------
class DependencyInstaller:
    def __init__(self):
        self.system = platform.system().lower()
        self.arch = platform.machine().lower()
        self.driver_dir = os.path.join(os.getcwd(), "drivers")
        
        if not os.path.exists(self.driver_dir):
            try: os.makedirs(self.driver_dir)
            except OSError: pass

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

    def is_python_package_installed(self, package_name):
        try:
            check_name = "python-docx" if package_name == "docx" else package_name
            subprocess.check_call([sys.executable, "-m", "pip", "show", check_name],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except subprocess.CalledProcessError:
            return False

    def install_python_packages(self):
        packages = ["selenium", "pygame", "xlsxwriter", "PyQt5", "phonenumbers", "pandas", "openpyxl", "requests", "python-docx"]
        for package in packages:
            if not self.is_python_package_installed(package):
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                    logging.info(f"Successfully installed {package}.")
                except subprocess.CalledProcessError: pass

    # Drivers handled by Selenium Manager
    def install_chromedriver(self): return True
    def install_geckodriver(self): return True
    def install_edgedriver(self): return True
    def install_all(self): self.install_python_packages()


# ------------------- Thread Signals -------------------
class ThreadSignals(QObject):
    update_sent = pyqtSignal(dict)
    finished = pyqtSignal()
    error_occurred = pyqtSignal(str)
    login_required = pyqtSignal()
    progress_update = pyqtSignal(int)


# ------------------- Sending Thread -------------------
class SendingThread(QThread):
    def __init__(self, parent, numbers, message, attached_file, browser, delay, installer):
        super().__init__()
        self.parent = parent
        self.numbers = numbers
        self.message = message
        self.attached_file = attached_file
        self.browser = browser
        self.delay = delay
        self.installer = installer
        self.signals = ThreadSignals()
        self.driver = None
        self.results = []

    def _validate_file(self):
        if self.attached_file and not os.path.exists(self.attached_file):
            raise FileNotFoundError("Attached file not found")
        return True

    def run(self):
        try:
            self._validate_file()

            # === Browser Setup ===
            if self.browser == "Chrome":
                options = webdriver.ChromeOptions()
                options.add_argument("--remote-debugging-pipe")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-gpu")
                options.add_argument("--log-level=3")
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                options.add_experimental_option("useAutomationExtension", False)
                profile_path = os.path.join(os.getcwd(), 'chrome_profile')
                options.add_argument(f"user-data-dir={profile_path}")
                self.driver = webdriver.Chrome(options=options)

            elif self.browser == "Firefox":
                options = webdriver.FirefoxOptions()
                # Firefox needs extra care for automation flags
                options.set_preference("dom.webdriver.enabled", False)
                options.set_preference("useAutomationExtension", False)
                profile_dir = os.path.join(os.getcwd(), 'firefox_profile')
                if not os.path.exists(profile_dir): os.makedirs(profile_dir)
                options.add_argument("-profile")
                options.add_argument(profile_dir)
                self.driver = webdriver.Firefox(options=options)

            elif self.browser == "Edge":
                options = webdriver.EdgeOptions()
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                options.add_experimental_option("useAutomationExtension", False)
                profile_path = os.path.join(os.getcwd(), 'edge_profile')
                options.add_argument(f"user-data-dir={profile_path}")
                self.driver = webdriver.Edge(options=options)

            elif self.browser == "Brave":
                options = webdriver.ChromeOptions()
                options.binary_location = self.installer.browser_paths.get("brave")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-gpu")
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                profile_path = os.path.join(os.getcwd(), 'brave_profile')
                options.add_argument(f"user-data-dir={profile_path}")
                self.driver = webdriver.Chrome(options=options)

            # Maximize window to ensure elements are visible and interactive
            self.driver.maximize_window()
            self.driver.get("https://web.whatsapp.com")

            # === FIX FOR FIREFOX FOCUS: FORCE CLICK BODY ===
            try:
                body = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                body.click() # Wake up the window
            except: pass

            # Login Wait
            try:
                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@data-testid="qrcode"]')))
                self.signals.login_required.emit()
                WebDriverWait(self.driver, 300).until_not(EC.presence_of_element_located((By.XPATH, '//div[@data-testid="qrcode"]')))
            except: pass

            WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "side")))

            # Processing Loop
            for index, number in enumerate(self.numbers):
                while self.parent.is_paused:
                    time.sleep(1)
                    if not self.parent.is_sending: break
                if not self.parent.is_sending: break
                
                if index > 0 and index % 50 == 0:
                    time.sleep(random.uniform(60, 120))

                result = {"number": number, "status": "Failed", "reason": ""}
                try:
                    self._process_number(number)
                    result["status"] = "Sent"
                    result["reason"] = "Success"
                except Exception as e:
                    if self._check_if_sent_strict():
                        result["status"] = "Sent"
                        result["reason"] = "Success (Verified)"
                    elif "Invalid" in str(e):
                        result["status"] = "Invalid"
                        result["reason"] = "Invalid Number"
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

    def _process_number(self, number):
        clean_number = ''.join(filter(str.isdigit, str(number)))
        encoded_message = urllib.parse.quote(self.message)
        url = f"https://web.whatsapp.com/send?phone={clean_number}&text={encoded_message}"
        
        self.driver.get(url)
        try: WebDriverWait(self.driver, 3).until(EC.alert_is_present()).accept()
        except: pass
        
        # Wait for Footer Input or Send Button
        try:
            WebDriverWait(self.driver, 40).until(
                lambda d: d.find_elements(By.XPATH, '//footer//div[@contenteditable="true"]') or 
                          d.find_elements(By.XPATH, '//span[@data-icon="send"]') or
                          d.find_elements(By.XPATH, '//div[contains(text(), "url is invalid")]')
            )
        except TimeoutException:
            if self._check_if_sent_strict(): return
            raise Exception("Chat load timeout")

        if self.driver.find_elements(By.XPATH, '//div[contains(text(), "url is invalid")]') or \
           self.driver.find_elements(By.XPATH, '//div[contains(text(), "phone number shared via url is invalid")]'):
            raise Exception("Invalid Phone Number")

        if self.attached_file:
            self._handle_attachments()

        self._target_and_send()
        self._verify_delivery()

    def _target_and_send(self):
        try:
            # 1. Find Input Box
            input_box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//footer//div[@contenteditable="true"][@role="textbox"]'))
            )
            
            # === Firefox Fix: Use ActionChains to Move and Click (Simulate Mouse) ===
            action = ActionChains(self.driver)
            action.move_to_element(input_box).click().perform()
            time.sleep(0.5)
            
            # 2. Enter
            input_box.send_keys(Keys.ENTER)
            time.sleep(1)
            
            if self._check_if_sent_strict(): return

            # 3. Fallback: JS Click
            if self.driver.find_elements(By.XPATH, '//span[@data-icon="send"]'):
                send_btn = self.driver.find_element(By.XPATH, '//span[@data-icon="send"]')
                parent = send_btn.find_element(By.XPATH, "./ancestor::button")
                self.driver.execute_script("arguments[0].click();", parent)
                time.sleep(0.5)
                if self._check_if_sent_strict(): return

            # 4. Wait for Manual Send
            for _ in range(10):
                if self._check_if_sent_strict(): return
                time.sleep(0.5)

        except Exception as e:
            logging.warning(f"Send warning: {e}")

    def _check_if_sent_strict(self):
        try:
            if self.driver.find_elements(By.XPATH, '//div[contains(@class, "message-out")]'): return True
            if self.driver.find_elements(By.XPATH, '//span[@data-icon="msg-dblcheck"] | //span[@data-icon="msg-check"]'): return True
            return False
        except: return False

    def _handle_attachments(self):
        try:
            attach_btn = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@title="Attach"] | //span[@data-icon="clip"]')))
            # Use JS Click for attachment button to be safe across browsers
            self.driver.execute_script("arguments[0].click();", attach_btn)
            
            file_input = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@type="file"]')))
            file_input.send_keys(self.attached_file)
            
            send_preview = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]')))
            self.driver.execute_script("arguments[0].click();", send_preview)
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


# ------------------- About Dialog -------------------
class AboutDialog(QDialog):
    def __init__(self, language, parent=None):
        super().__init__(parent)
        self.language = language
        self.initUI()

    def initUI(self):
        is_ar = self.language == "Arabic"
        self.setWindowTitle("About" if not is_ar else "حول البرنامج")
        self.setFixedSize(450, 350)
        self.setStyleSheet("""
            QDialog { background-color: #f0f2f5; }
            QLabel { color: #128C7E; }
            QLabel#title { font-size: 22px; font-weight: bold; color: #075E54; margin-bottom: 10px; }
            QLabel#desc { color: #333; font-size: 12px; margin-bottom: 15px; }
            QLabel#link { color: #007acc; text-decoration: none; font-weight: bold; }
            QLabel#copy { color: #888; font-size: 10px; margin-top: 20px; }
            QFrame { background-color: #ccc; }
        """)

        layout = QVBoxLayout()
        
        title_lbl = QLabel("WhatsApp Sender Pro")
        title_lbl.setObjectName("title")
        title_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_lbl)

        desc_text = (
            "A powerful automation tool designed to send bulk WhatsApp messages efficiently and securely. "
            "Supports multiple browsers, smart number import, and detailed reporting."
            if not is_ar else
            "أداة أتمتة قوية مصممة لإرسال رسائل واتساب جماعية بكفاءة وأمان. "
            "تدعم متصفحات متعددة، استيراد ذكي للأرقام، وتقارير مفصلة."
        )
        desc_lbl = QLabel(desc_text)
        desc_lbl.setObjectName("desc")
        desc_lbl.setWordWrap(True)
        desc_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc_lbl)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)

        form_layout = QFormLayout()
        form_layout.setSpacing(10)
        
        def create_link_label(text, url):
            lbl = QLabel(f'<a href="{url}" style="color: #007acc; text-decoration: none;">{text}</a>')
            lbl.setOpenExternalLinks(True)
            lbl.setObjectName("link")
            return lbl

        dev_key = QLabel("Developer:" if not is_ar else "المطور:")
        dev_key.setFont(QFont("Arial", 10, QFont.Bold))
        dev_val = QLabel("Fahad M Alsamiti" if not is_ar else "فهد منصور الصامطي")
        dev_val.setFont(QFont("Arial", 10))

        github_key = QLabel("GitHub:")
        github_key.setFont(QFont("Arial", 10, QFont.Bold))
        github_val = create_link_label("FahadMAlsamiti", "https://github.com/FahadMAlsamiti")

        x_key = QLabel("X Platform:")
        x_key.setFont(QFont("Arial", 10, QFont.Bold))
        x_val = create_link_label("@FahadAlsamiti", "https://x.com/FahadAlsamiti")

        form_layout.addRow(dev_key, dev_val)
        form_layout.addRow(github_key, github_val)
        form_layout.addRow(x_key, x_val)
        
        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        layout.addWidget(form_widget)

        copy_lbl = QLabel("© 2025 Fahad Alsamiti. All Rights Reserved." if not is_ar else "© 2025 فهد الصامطي. جميع الحقوق محفوظة.")
        copy_lbl.setObjectName("copy")
        copy_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(copy_lbl)

        self.setLayout(layout)
        if is_ar: self.setLayoutDirection(Qt.RightToLeft)


# ------------------- Main Window -------------------
class WhatsAppSenderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings_file = "settings.json"
        self.installer = DependencyInstaller()
        self.driver_dir = self.installer.driver_dir
        self.load_settings()
        
        self.setWindowTitle("WhatsApp Message Sender")
        self.setGeometry(300, 200, 950, 650)
        
        self.sent_count = 0
        self.remaining_numbers = []
        self.is_sending = False
        self.is_paused = False
        self.attached_file = None
        
        pygame.mixer.init()
        self.installer.install_python_packages()
        
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
        
        self.about_action = QAction("About", self)
        self.about_action.triggered.connect(self.show_about)
        settings_menu.addAction(self.about_action)

        # Phone Frame
        phone_frame = QFrame()
        phone_frame.setStyleSheet("border: 1px solid gray; padding: 10px;")
        phone_layout = QVBoxLayout()
        phone_frame.setLayout(phone_layout)
        self.numbers_label = QLabel("Phone Numbers:")
        self.numbers_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.numbers_label.setAlignment(Qt.AlignCenter)
        phone_layout.addWidget(self.numbers_label)
        self.numbers_input = QTextEdit()
        self.numbers_input.setFont(QFont("Arial", 11))
        self.numbers_input.textChanged.connect(self.update_numbers_count)
        phone_layout.addWidget(self.numbers_input)
        self.main_layout.addWidget(phone_frame)

        # Message Frame
        message_frame = QFrame()
        message_frame.setStyleSheet("border: 1px solid gray; padding: 10px;")
        message_layout = QVBoxLayout()
        message_frame.setLayout(message_layout)
        self.message_label = QLabel("Message:")
        self.message_label.setFont(QFont("Arial", 12, QFont.Bold))
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

    def show_about(self):
        dlg = AboutDialog(self.language, self)
        dlg.exec_()

    def update_numbers_count(self):
        numbers = self.numbers_input.toPlainText().strip().split("\n")
        valid_numbers = [num for num in numbers if num.strip()]
        self.remaining_numbers = valid_numbers
        self.update_stats_text()

    def update_stats_text(self):
        is_ar = self.language == "Arabic"
        total = len(self.remaining_numbers) + self.sent_count
        self.total_numbers_label.setText(f"{'العدد الكلي' if is_ar else 'Total'}: {total}")
        self.remaining_numbers_label.setText(f"{'المتبقي' if is_ar else 'Remaining'}: {len(self.remaining_numbers)}")
        self.sent_numbers_label.setText(f"{'تم الإرسال' if is_ar else 'Sent'}: {self.sent_count}")

    def import_numbers(self):
        options = QFileDialog.Options()
        # Filter for CSV, Excel, Word, and VCard
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import Numbers", "", 
            "Files (*.csv *.xlsx *.xls *.txt *.docx *.doc *.vcf);;All Files (*)", 
            options=options
        )
        if file_path:
            try:
                extracted_numbers = []
                # Strict Regex for International Numbers
                pattern = r'\+\d{7,15}'
                
                text_content = ""
                
                if file_path.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path, dtype=str, header=None)
                    text_content = " ".join(df.values.flatten().astype(str))
                elif file_path.endswith('.csv'):
                    df = pd.read_csv(file_path, dtype=str, header=None)
                    text_content = " ".join(df.values.flatten().astype(str))
                elif file_path.endswith('.docx'):
                    try:
                        import docx
                        doc = docx.Document(file_path)
                        # Extract text from paragraphs
                        full_text = [p.text for p in doc.paragraphs]
                        # Extract text from tables
                        for table in doc.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    full_text.append(cell.text)
                        text_content = " ".join(full_text)
                    except ImportError:
                        raise Exception("python-docx not installed")
                elif file_path.endswith('.doc') or file_path.endswith('.vcf'):
                    with open(file_path, "rb") as f:
                        text_content = f.read().decode('utf-8', errors='ignore') + " " + f.read().decode('latin-1', errors='ignore')
                else:
                    with open(file_path, "r", encoding='utf-8', errors='ignore') as f:
                        text_content = f.read()
                
                extracted_numbers = re.findall(pattern, text_content)
                unique_nums = list(set(extracted_numbers))
                
                if unique_nums:
                    current = self.numbers_input.toPlainText()
                    new_text = (current + "\n" + "\n".join(unique_nums)).strip()
                    self.numbers_input.setPlainText(new_text)
                    self.update_numbers_count()
                    msg = f"تم استيراد {len(unique_nums)} رقم دولي" if self.language == "Arabic" else f"Imported {len(unique_nums)} international numbers"
                    QMessageBox.information(self, "Success", msg)
                else:
                    msg = "لم يتم العثور على أرقام دولية (+)" if self.language == "Arabic" else "No international numbers (+) found"
                    QMessageBox.warning(self, "Warning", msg)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def start_sending(self):
        try:
            is_ar = self.language == "Arabic"
            if not self.remaining_numbers:
                msg = "الرجاء إدخال الأرقام" if is_ar else "Please enter numbers."
                return QMessageBox.warning(self, "Warning", msg)
            if not self.message_input.toPlainText().strip() and not self.attached_file:
                msg = "الرجاء إدخال رسالة أو مرفق" if is_ar else "Please enter a message or file."
                return QMessageBox.warning(self, "Warning", msg)
            
            if os.path.exists("start_sound.mp3"): 
                try: pygame.mixer.music.load("start_sound.mp3"); pygame.mixer.music.play()
                except: pass
            
            self.is_sending = True
            self.is_paused = False
            self.sent_count = 0
            self.progress_bar.setValue(0)
            
            self.sending_thread = SendingThread(
                self, self.remaining_numbers.copy(), self.message_input.toPlainText(),
                self.attached_file, self.browser, self.default_delay, self.installer
            )
            self.sending_thread.signals.update_sent.connect(self.update_sent_count)
            self.sending_thread.signals.finished.connect(self.sending_finished)
            self.sending_thread.signals.error_occurred.connect(self.show_error)
            self.sending_thread.signals.login_required.connect(self.show_login_required)
            self.sending_thread.signals.progress_update.connect(self.progress_bar.setValue)
            self.sending_thread.start()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def stop_sending(self):
        if self.is_sending:
            self.is_paused = True
            msg = "تم الإيقاف المؤقت" if self.language == "Arabic" else "Sending Paused"
            QMessageBox.information(self, "Paused", msg)

    def resume_sending(self):
        if self.is_sending:
            self.is_paused = False
            msg = "تم الاستئناف" if self.language == "Arabic" else "Sending Resumed"
            QMessageBox.information(self, "Resumed", msg)

    def attach_file(self):
        options = QFileDialog.Options()
        path, _ = QFileDialog.getOpenFileName(self, "Attach", "", "All Files (*.*)", options=options)
        if path:
            self.attached_file = path
            self.attach_button.setText(f"📎 {os.path.basename(path)}")

    def export_report(self):
        is_ar = self.language == "Arabic"
        if not hasattr(self, 'sending_thread') or not self.sending_thread.results:
            msg = "لا توجد بيانات" if is_ar else "No Data"
            return QMessageBox.warning(self, "Warning", msg)
        
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
                msg = "تم التصدير بنجاح" if is_ar else "Exported Successfully"
                QMessageBox.information(self, "Success", msg)
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
        self.browser = b; self.save_settings()
        msg = f"تم تعيين {b}" if self.language == "Arabic" else f"Set to {b}"
        QMessageBox.information(self, "Browser", msg)
    def set_message_delay(self):
        is_ar = self.language == "Arabic"
        title = "تأخير" if is_ar else "Delay"
        lbl = "مللي ثانية:" if is_ar else "MS:"
        v, ok = QInputDialog.getInt(self, title, lbl, self.default_delay, 500, 60000)
        if ok: self.default_delay = v; self.save_settings()
    def update_sent_count(self, d):
        self.sent_count = d['sent']; self.update_numbers_count()
    
    def sending_finished(self):
        self.is_sending = False; self.is_paused = False
        msg = "اكتمل الإرسال" if self.language == "Arabic" else "Sending Finished"
        QMessageBox.information(self, "Done", msg)
    
    def show_error(self, m):
        self.is_sending = False
        QMessageBox.critical(self, "Error", m)
    
    def show_login_required(self):
        msg = "امسح الرمز" if self.language == "Arabic" else "Scan QR"
        QMessageBox.warning(self, "Login", msg)

    def retranslate_ui(self):
        is_ar = self.language == "Arabic"
        self.numbers_label.setText("أرقام الهواتف:" if is_ar else "Phone Numbers:")
        self.message_label.setText("الرسالة:" if is_ar else "Message:")
        self.import_button.setText("استيراد أرقام" if is_ar else "Import Numbers")
        self.send_button.setText("إرسال الرسائل" if is_ar else "Send Messages")
        self.stop_button.setText("إيقاف مؤقت" if is_ar else "Stop Sending")
        self.resume_button.setText("استئناف" if is_ar else "Resume Sending")
        self.attach_button.setText("إرفاق ملف" if is_ar else "Attach File")
        self.export_button.setText("تصدير تقرير" if is_ar else "Export Report")
        self.bold_btn.setText("عريض" if is_ar else "Bold")
        self.italic_btn.setText("مائل" if is_ar else "Italic")
        self.color_btn.setText("لون" if is_ar else "Color")
        self.font_btn.setText("حجم" if is_ar else "Font Size")
        self.settings_menu.setTitle("إعدادات" if is_ar else "Settings")
        self.language_menu.setTitle("اللغة" if is_ar else "Language")
        self.theme_menu.setTitle("المظهر" if is_ar else "Theme")
        self.browser_menu.setTitle("المتصفح" if is_ar else "Browser")
        self.delay_action.setText("تأخير" if is_ar else "Set Message Delay")
        self.about_action.setText("حول البرنامج" if is_ar else "About")
        
        if is_ar: self.setLayoutDirection(Qt.RightToLeft)
        else: self.setLayoutDirection(Qt.LeftToRight)
        self.update_numbers_count()

    def closeEvent(self, event):
        self.save_settings()
        if hasattr(self, 'sending_thread'): self.sending_thread.quit()
        event.accept()

if __name__ == "__main__":
    # NO BLOCKING INSTALLS HERE
    app = QApplication(sys.argv)
    window = WhatsAppSenderApp()
    window.show()
    sys.exit(app.exec_())