from threading import Thread
import asyncio
import time
import os
import base64
import sys
from pathlib import Path
import tkinter as tk
from tkinter import simpledialog, messagebox
from tkinter.scrolledtext import ScrolledText
from enum import Enum
import traceback

import aiohttp
from selenium_profiles.webdriver import Chrome
from selenium_profiles.profiles import profiles
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    ElementNotInteractableException,
)
from selenium import webdriver
from pystray import Icon, MenuItem, Menu
from PIL import Image
import yaml
import win32com.client
import pythoncom
import winreg
import subprocess


def show_alert(title, message):
    """Display an alert message box with a title and message."""
    TIME_TO_WAIT = 2000  # in milliseconds
    root = tk.Tk()
    root.withdraw()
    try:
        root.after(TIME_TO_WAIT, root.destroy)
        messagebox.showerror(title, message, master=root)
    except tk.TclError:
        pass


class LogStatus(Enum):
    NOT_LOGGED_IN = "Not Logged in to SRMIST Portal."
    LOGIN_WAITING = "Not connected to SRMIST Wi-Fi."

    LOGIN_FAILED = "Login failed in SRMIST Portal."
    LOGIN_PENDING = "Login pending in SRMIST Portal."
    LOGIN_SUCCESS = "Login successful in SRMIST Portal."


def is_windows_11():
    if sys.platform != "win32":
        return False  # Not a Windows OS

    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        )
        build_number, _ = winreg.QueryValueEx(key, "CurrentBuildNumber")
        winreg.CloseKey(key)

        return int(build_number) >= 22000
    except Exception:
        return False


def restart():
    try:
        if sys.argv[0].endswith(".py"):
            script_name = sys.argv[0]
        else:
            script_name = "main.py"

        subprocess.Popen(["pythonw", script_name])
        print(f"Restarted the script using pythonw {script_name}")
    except Exception as e:
        print(f"Failed to restart the script: {e}")
        sys.exit(1)


def restart_and_exit():
    restart()
    sys.exit(0)


def install_and_restart():
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "blinker==1.7.0"]
        )
        print("Successfully installed blinker 1.7.0.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install blinker 1.7.0: {e}")
        sys.exit(1)

    restart()
    sys.exit(0)


config = None
previous_login_url = None
first_run = True
profile = profiles.Windows()
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")
try:
    driver = Chrome(
        profile,
        options=options,
        uc_driver=False,
    )
except ModuleNotFoundError:
    install_and_restart()

IS_WIN_11 = is_windows_11()
login_status = LogStatus.NOT_LOGGED_IN
last_login_time = 0
log_text = []


def log_message(message):
    """Append log messages to the log text."""
    log_text.append(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}")


def create_image_from_file():
    """Load the tray icon from a 'logo.png' file in the current directory."""
    logo_path = os.path.join(os.getcwd(), "resources", "logo.png")
    if os.path.exists(logo_path):
        return Image.open(logo_path)
    else:
        raise FileNotFoundError(f"Logo file not found: {logo_path}")


async def fetch_url(session, url, timeout):
    try:
        async with session.get(url, timeout=timeout) as response:
            return url, response.status
    except asyncio.TimeoutError:
        return url, None
    except Exception:
        return url, None


async def get_faster_url(urls, timeout=5):
    async with aiohttp.ClientSession() as session:
        tasks = [asyncio.create_task(fetch_url(session, url, timeout)) for url in urls]
        done, _ = await asyncio.wait(tasks, return_when=asyncio.FIRST_COMPLETED)

        for task in done:
            url, status = task.result()
            if status == 200:
                return url
            elif status == 401:
                return None

        return None


async def get_login_url():
    url = await get_faster_url(
        [
            "https://iac.srmist.edu.in/Connect/PortalMain",
            "https://iach.srmist.edu.in/Connect/PortalMain",
        ]
    )
    return url


def seconds_to_hms(seconds):
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    return f"{int(hours):02}:{int(minutes):02}:{int(secs):02}"


lock = asyncio.Event()


async def run_every_n_mins(interval_mins):
    interval_seconds = interval_mins * 60
    while True:
        start_time = time.time()
        status = await login()
        elapsed_time = time.time() - start_time
        sleep_time = max(1, interval_seconds - elapsed_time)
        if status:
            log_message(
                f"Performing login on {status} task took {seconds_to_hms(elapsed_time)}s, next login-check after {seconds_to_hms(sleep_time)}s."
            )

        try:
            await asyncio.wait_for(lock.wait(), timeout=sleep_time)
            lock.clear()
            log_message("User initiated refresh, performing a login-check.")
        except asyncio.TimeoutError:
            pass  # Timeout without interruption; continue normally


def refresh_login_status():
    lock.set()


def get_traceback(error):
    etype = type(error)
    trace = error.__traceback__
    lines = traceback.format_exception(etype, error, trace)
    traceback_text = "".join(lines)
    return traceback_text


async def login(*, retry_count=1):
    global config, login_status, previous_login_url, last_login_time, first_run
    if retry_count > 5:
        log_message("Retry counts exceeded, exiting...")
        login_status = LogStatus.LOGIN_FAILED
        time.sleep(120)
        return

    preferred_url = await get_login_url()
    if (
        previous_login_url
        and preferred_url == previous_login_url
        and login_status == LogStatus.LOGIN_SUCCESS
        and int(time.time()) - last_login_time <= (12 * 60 * 60)
    ):
        return
    if preferred_url:
        login_status = LogStatus.LOGIN_PENDING
        update_menu(icon)
        driver.get(preferred_url)
        try:
            try:
                WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.ID, "UserCheck_Logoff_Button"))
                )
                last_login_time = int(time.time())
                previous_login_url = preferred_url
                login_status = LogStatus.LOGIN_SUCCESS
                update_menu(icon)
                if not first_run:
                    driver.close()
                first_run = False
                log_message("Already logged in, skipping login!")

                return preferred_url
            except TimeoutException:
                pass
            username_div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.ID, "LoginUserPassword_auth_username")
                )
            )
            password_div = driver.find_element(By.ID, "LoginUserPassword_auth_password")
            username_div.send_keys(config["credentials"]["username"])
            try:
                password_div.send_keys(
                    base64.b64decode(config["credentials"]["password"]).decode("utf-8")
                )
            except UnicodeDecodeError as e:
                log_message(get_traceback(e))
                log_message("Invalid base64 credentials, skipping login...")
                show_alert(
                    "Warning!", "Invalid base64 credentials, unsuccessful login!"
                )
                login_status = LogStatus.LOGIN_FAILED
                return
            login_button = driver.find_element(By.ID, "UserCheck_Login_Button")
            login_button.click()
            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.ID, "nac_expiration_time"))
                )
                last_login_time = int(time.time())
                previous_login_url = preferred_url
                login_status = LogStatus.LOGIN_SUCCESS
                update_menu(icon)
                if not first_run:
                    driver.close()
                first_run = False

                return preferred_url
            except TimeoutException as e:
                log_message(get_traceback(e))
                log_message("Invalid credentials, skipping login...")
                show_alert("Warning!", "Invalid login credentials, unsuccessful login!")
                login_status = LogStatus.LOGIN_FAILED
                username, password = ask_for_refreshed_credentials()
                if username and password:
                    save_credentials(username, password)
                    config = yaml.safe_load(open("config.yml"))
                else:
                    await login(retry_count=retry_count + 1)
        except TimeoutException as e:
            log_message(get_traceback(e))
            log_message("Login page took too long to load, retrying login.")
            login_status = LogStatus.LOGIN_FAILED
            await asyncio.sleep(1)
            await login(retry_count=retry_count + 1)
        except NoSuchElementException as e:
            log_message(get_traceback(e))
            log_message("Invalid state, no such elements found. Retrying...")
            login_status = LogStatus.LOGIN_FAILED
            await asyncio.sleep(1)
            await login(retry_count=retry_count + 1)
        except ElementNotInteractableException as e:
            log_message(get_traceback(e))
            log_message("Invalid state, elements aren't interactable. Retrying...")
            login_status = LogStatus.LOGIN_FAILED
            await asyncio.sleep(1)
            await login(retry_count=retry_count + 1)
        except KeyError as e:
            log_message(get_traceback(e))
            show_alert(
                "Error!",
                "Invalid config.yml (must contain username, password), exiting...",
            )
            login_status = LogStatus.LOGIN_FAILED
            sys.exit(-1)
    else:
        if login_status != LogStatus.LOGIN_WAITING:
            log_message("Not connected to SRMIST Wi-Fi, skipping login attempt.")
            show_alert(
                "Warning!", "Not connected to SRMIST Wi-Fi, skipping login attempt."
            )
        login_status = LogStatus.LOGIN_WAITING
        update_menu(icon)


async def logout():
    """Perform logout operation."""
    global login_status
    try:
        preferred_url = await get_login_url()
        if preferred_url:
            driver.get(preferred_url)
            if login_status != LogStatus.LOGIN_SUCCESS:
                show_alert("Warning!", "Not logged in, skipping logout.")
                log_message("Already logged out, skipping logout.")
                login_status = LogStatus.NOT_LOGGED_IN
                update_menu(icon)
                return
            logout_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "UserCheck_Logoff_Button"))
            )
            logout_button.click()
            log_message("Successfully logged out.")
            login_status = LogStatus.NOT_LOGGED_IN
            update_menu(icon)
    except TimeoutException:
        show_alert("Error!", "Error occured while trying to logout, try again later.")
        log_message(
            "Invalid state, logout button not found while status is LOGIN_SUCCESS."
        )
        login_status = LogStatus.NOT_LOGGED_IN
        update_menu(icon)


def show_logs():
    """Display logs in a separate window and auto-update if new logs are added."""
    if hasattr(show_logs, "log_window") and show_logs.log_window:
        try:
            if show_logs.log_window.winfo_exists():
                show_logs.log_window.lift()
                return
        except tk.TclError:
            show_logs.log_window = None

    root = tk.Tk()
    root.title("Logs")
    root.geometry("600x400")
    root.attributes("-topmost", True)
    root.after(0, lambda: root.focus_force())

    show_logs.log_window = root

    text_widget = ScrolledText(root, wrap=tk.WORD, height=20, width=60)
    text_widget.pack(expand=True, fill=tk.BOTH)

    displayed_logs = [len(log_text)]

    def update_logs():
        """Update the log window with new entries if available."""
        if len(log_text) > displayed_logs[0]:
            new_logs = log_text[displayed_logs[0] :]
            for log in new_logs:
                text_widget.configure(state=tk.NORMAL)
                text_widget.insert(tk.END, log + "\n")
                text_widget.configure(state=tk.DISABLED)

            displayed_logs[0] = len(log_text)
            text_widget.see(tk.END)

        root.after(500, update_logs)

    text_widget.configure(state=tk.NORMAL)
    for log in log_text:
        text_widget.insert(tk.END, log + "\n")
    text_widget.configure(state=tk.DISABLED)

    update_logs()

    def on_close():
        show_logs.log_window = None
        root.destroy()

    if IS_WIN_11:
        root.after(5000, on_close)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


def save_credentials(username, password):
    """Save username and password to config.yml."""
    try:
        config_data = {}
        try:
            with open("config.yml", "r") as f:
                config_data = yaml.safe_load(f) or {}
        except FileNotFoundError:
            pass

        config_data["credentials"] = {
            "username": username,
            "password": base64.b64encode(password.encode("utf-8")).decode("utf-8"),
        }

        config_data["interval_mins"] = 0.5
        with open("config.yml", "w") as f:
            yaml.safe_dump(config_data, f)

    except Exception as e:
        show_alert("Error", f"Failed to save credentials: {str(e)}")


def show_message(title, message):
    """Display an message box with a title and message."""
    TIME_TO_WAIT = 3000  # in milliseconds
    root = tk.Tk()
    root.withdraw()
    try:
        root.after(TIME_TO_WAIT, root.destroy)
        messagebox.showinfo(title, message, master=root)
    except tk.TclError:
        pass


def ask_for_credentials():
    """Ask the user for username and password."""
    root = tk.Tk()
    root.withdraw()
    show_message("Setup", "Detected first startup, continue to enter your credentials.")
    username = simpledialog.askstring("Login", "Enter your username:")
    if username is None:
        show_alert("Error", "Username is required.")
        return None, None

    password = simpledialog.askstring("Login", "Enter your password:", show="*")
    if password is None:
        show_alert("Error", "Password is required.")
        return None, None

    return username, password


def ask_for_refreshed_credentials():
    """Ask the user for username and password."""
    root = tk.Tk()
    root.withdraw()
    username = simpledialog.askstring("Login", "Enter your username:")
    if username is None:
        show_alert("Error", "Username is required.")
        return None, None

    password = simpledialog.askstring("Login", "Enter your password:", show="*")
    if password is None:
        show_alert("Error", "Password is required.")
        return None, None

    return username, password


def save_autostart_shortcut():
    pythoncom.CoInitialize()
    current_file = Path(__file__).resolve()
    startup_folder = (
        os.getenv("APPDATA") + r"\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"
    )

    shortcut_path = os.path.join(startup_folder, "SRMAutoLogin.lnk")
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)

    pythonw_path = sys.executable.replace("python.exe", "pythonw.exe")
    shortcut.TargetPath = pythonw_path
    shortcut.Arguments = str(current_file)

    shortcut.WorkingDirectory = str(current_file.parent)
    shortcut.Description = "SRM WI-FI autologin software."
    shortcut.save()
    pythoncom.CoUninitialize()


def start_loop():
    global config
    try:
        config = yaml.safe_load(open("config.yml"))
    except FileNotFoundError:
        log_message("No username and password found in config, asking user...")
        username, password = ask_for_credentials()
        if username and password:
            save_credentials(username, password)
            config = yaml.safe_load(open("config.yml"))
        else:
            sys.exit(-1)
            return
    try:
        n_mins = float(config["interval_mins"])
    except ValueError:
        show_alert("Warning!", "Invalid interval_mins found in config, must be an int.")
        log_message("Invalid interval_mins found in config, defaulting to 0.5 mins.")
        n_mins = 0.5
    except KeyError:
        log_message("No interval_mins found in config, defaulting to 0.5 mins.")
        n_mins = 0.5

    try:
        config["credentials"]["username"]
        config["credentials"]["password"]
    except KeyError:
        log_message("No username and password found in config, asking user...")
        username, password = ask_for_credentials()
        if username and password:
            save_credentials(username, password)
            config = yaml.safe_load(open("config.yml"))
        else:
            sys.exit(-1)
            return

    save_autostart_shortcut()
    asyncio.run(run_every_n_mins(n_mins))


def start_app():
    """Start the asyncio loop."""
    Thread(target=start_loop, daemon=True).start()


def run_async_coro(coro):
    def wrapper(icon, item):
        asyncio.run(coro())

    return wrapper


def update_login_status(icon):
    show_message("Status", login_status.value)


def update_menu(icon):
    """Update the tray menu dynamically based on login status."""
    if login_status == LogStatus.LOGIN_SUCCESS:
        icon.menu = Menu(
            MenuItem("Show Logs", show_logs),
            MenuItem("Show Status", update_login_status),
            MenuItem("Logout Wi-Fi", run_async_coro(logout)),
            MenuItem("Restart", restart_and_exit),
            MenuItem("Quit", lambda: icon.stop()),
        )
    elif (
        login_status == LogStatus.NOT_LOGGED_IN
        or login_status == LogStatus.LOGIN_FAILED
    ):
        icon.menu = Menu(
            MenuItem("Show Logs", show_logs),
            MenuItem("Show Status", update_login_status),
            MenuItem("Login Wi-Fi", run_async_coro(login)),
            MenuItem("Restart", restart_and_exit),
            MenuItem("Quit", lambda: icon.stop()),
        )
    else:
        icon.menu = Menu(
            MenuItem("Show Logs", show_logs),
            MenuItem("Show Status", update_login_status),
            MenuItem("Refresh Status", refresh_login_status),
            MenuItem("Restart", restart_and_exit),
            MenuItem("Quit", lambda: icon.stop()),
        )
    icon.update_menu()


if __name__ == "__main__":
    icon = Icon(
        "Login App",
        create_image_from_file(),
    )
    update_menu(icon)

    start_app()
    icon.run()
