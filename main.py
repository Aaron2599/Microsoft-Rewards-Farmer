import pickle
import ctypes
import os
import random
import shutil
import subprocess
import time
import traceback
from datetime import datetime
from pathlib import Path
from patchright.sync_api import sync_playwright, BrowserContext


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint),
        ("dwTime", ctypes.c_ulong)
    ]


def get_idle_duration():
    """
    Returns the user's idle time in seconds on Windows.
    """
    lastInputInfo = LASTINPUTINFO()
    lastInputInfo.cbSize = ctypes.sizeof(lastInputInfo)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lastInputInfo)):
        # GetTickCount() returns milliseconds since system startup
        millis = ctypes.windll.kernel32.GetTickCount() - lastInputInfo.dwTime
        return millis / 1000.0
    return 0.0  # Return 0 if unable to get info


def taskkill_edge():
    """Kills all running Microsoft Edge processes using the taskkill command."""
    try:
        # The /IM flag specifies the image name (executable name)
        # The /F flag forcefully terminates the process
        command = ["taskkill", "/F", "/IM", "msedge.exe"]

        print(f"Attempting to kill Microsoft Edge processes with command: {' '.join(command)}")

        result = subprocess.run(command, capture_output=True, text=True, check=True)

        print("Microsoft Edge processes terminated successfully.")
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)

    except subprocess.CalledProcessError as e:
        print(f"Error terminating Microsoft Edge processes: {e}")
        print("STDOUT:", e.stdout)
        print("STDERR:", e.stderr)
    except FileNotFoundError:
        print("Error: 'taskkill' command not found. This should not happen on Windows.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def log_to_file(message: str,
                filename: str = Path(os.getenv("userprofile")) / "documents" / "microsoft_point_farmer.log"):
    """
    Logs a message with a timestamp to a specified file.

    This function is designed for simplicity and directness,
    making it highly readable for basic logging needs.

    Args:
        message (str): The string message to be logged.
        filename (str): The name of the log file. Defaults to "application.log".
                        The file will be created if it doesn't exist,
                        and messages will be appended if it does.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"

    try:
        # Ensure the directory exists if the filename includes a path
        log_dir = os.path.dirname(filename)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)  # exist_ok=True prevents error if dir already exists

        # Open the file in append mode ('a') and write the log entry
        with open(filename, 'a', encoding='utf-8') as f:
            f.write(log_entry)
        # print(f"Logged: {message}") # Optional: for immediate console feedback
    except IOError as e:
        print(f"ERROR: Could not write to log file '{filename}'. Reason: {e}")
    except Exception as e:
        print(f"An unexpected error occurred while logging: {e}")


def get_rewards_page(browser: BrowserContext):
    page = browser.new_page()
    page.goto("https://rewards.bing.com/")

    # CLOSE ANY POPUPS
    try:
        page.click(
            "//div/div[2]/div[4]/mee-rewards-user-status-banner-item/mee-rewards-user-status-banner-streak/mee-rewards-pop-up/div/div[2]/div[1]/span",
            timeout=1)
    except Exception:
        pass

    return page


def get_points(browser: BrowserContext):
    page = get_rewards_page(browser)

    time.sleep(1.5)

    try:
        points = page.locator(
            "//div/div/div/div/div[2]/div[1]/mee-rewards-user-status-banner-item/mee-rewards-user-status-banner-balance/div/div/div/div/div/div/p/mee-rewards-counter-animation/span").text_content()
        return points
    except Exception:
        return -1


def complete_quests(browser: BrowserContext):
    page = get_rewards_page(browser)
    tasks = page.locator("//main/div/ui-view/mee-rewards-dashboard/main/div//a").all()
    for task in tasks:
        href = task.get_attribute("href")
        if href and "rewards." not in href and "search" in href:
            try:
                task.click()
                w = page.viewport_size["width"]
                h = page.viewport_size["height"]
                page.mouse.move(random.randint(0, w), random.randint(0, h))
                time.sleep(3)
                page.evaluate(f"window.scrollBy(0, {random.randint(0, 1000)});")
                time.sleep(2)
            except Exception:
                continue
            finally:
                time.sleep(1)
    page.close()


def complete_daily_search(browser: BrowserContext):
    with open('queries.pkl', 'rb') as file:
        queries = pickle.load(file)

    for i in range(32):
        page = browser.new_page()
        page.goto("https://bing.com")
        search_bar = page.locator("//textarea")
        search_bar.fill(random.choice(queries))
        time.sleep(0.5)
        page.keyboard.down("Enter")
        time.sleep(2)
        page.close()


def sync_browser_data():
    data_dir = Path(os.getenv('LOCALAPPDATA')) / "Microsoft" / "Edge" / "User Data"
    copy_dir = "./.msedge"
    try:
        shutil.copytree(data_dir, copy_dir, dirs_exist_ok=True)
    except:
        pass


def main():
    sync_browser_data()

    if not os.path.exists("./FirstRun"):
        try:
            command = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                       os.path.abspath("./installed.html")]
            subprocess.run(command, check=True, shell=False)
            with open("./FirstRun", 'w') as file:
                pass
        except:
            pass
        finally:
            log_to_file("Successfully Installed")

    last_run = -1
    print(datetime.now().day)

    while True:

        if get_idle_duration() > 400:
            increment = 1
        elif get_idle_duration() > 200:
            increment = 10
        else:
            increment = 100

        if get_idle_duration() > 500 and last_run != datetime.now().day:

            sync_browser_data()

            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.launch_persistent_context("./.msedge", channel="msedge", headless=False)

                    start_points = get_points(browser)

                    complete_quests(browser)

                    complete_daily_search(browser)

                    end_points = get_points(browser)

                    browser.close()

                if start_points != end_points:
                    log_to_file(f"points: {end_points}")
                    log_to_file(f"collected:{end_points - start_points}")
                    last_run = datetime.now().day
                else:
                    log_to_file(f"No change in points trying again soon")
                    time.sleep(2 * 60 * 60)

            except Exception as e:
                t = traceback.format_exc()
                print(t)
                log_to_file(f"ERROR: {t}")

        time.sleep(increment)


if __name__ == '__main__':
    main()
