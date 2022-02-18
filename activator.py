import os
import sys
import pywinauto
import requests
from bs4 import BeautifulSoup
import pyautogui as gui
import winreg as reg
import ctypes
import subprocess


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def start():
    disk = os.environ['windir'][0]
    """This function searches for instances of Kaspersky on a computer."""
    files = [print("Kaspersky found!") for foldername, subfolder, filename in
             os.walk(fr"{disk}:\Program Files (x86)\Kaspersky Lab") if "avp.exe" in filename]
    if files:
        uninstall(disk)
    elif not files:
        try:
            deleting_certificates(disk)
        except FileNotFoundError:
            download(disk)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def uninstall(disk):
    """This function uninstalls Kaspersky Software if it's found in Program Files Directory"""
    Data = str(subprocess.check_output(
        ["wmic", "product", "where", "name like 'Kaspersky Internet Security'", "get", "IdentifyingNumber"])).split(
        "\\r\\r\\n")[1]
    app = pywinauto.Application(backend="uia").start(
        disk + r":\Windows\SysWOW64\MsiExec.exe /i" + Data + "REMOVE=ALL")
    try:
        wizzard = app['Kaspersky Internet Security']
        wizzard.NextButton.wait("enabled")
        wizzard.NextButton.click()
        wizzard.CheckBox.click()
        wizzard.Button2.click()
        wizzard.Button0.click()
        app['Kaspersky Internet Security'].child_window(title='Kaspersky Internet Security').wait('visible',
                                                                                                  timeout=200)
        wizzard.Button1.click()
    except:
        # The name of the occurring error is undefined. That's why we will use only except.
        # This error appears, when the user - didn't turn off the 'self-defense' feature,
        # so it makes the click impossible.
        print('Looks like you didn\'t turn off the self-defense feature in Kaspersky')


def deleting_certificates(disk):
    iteration_number = 0
    sub_keys = []
    with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\SystemCertificates\SPC\Certificates', 0,
                     reg.KEY_ALL_ACCESS) as sub_folder:
        while True:
            try:
                sub_keys.append(reg.EnumKey(sub_folder, iteration_number))
                iteration_number += 1
            except WindowsError:
                break
            except OSError:
                break
        for sub_key in sub_keys:
            reg.DeleteKey(sub_folder, sub_key)
            print(f'The subkey: {sub_key} have been deleted!')
    try:
        download(disk)
    except requests.exceptions.ConnectionError:
        print('No internet connection found!')
        gui.sleep(5)


def download(disk):
    """This function downloads and prepares the installer for a further installation"""
    os.makedirs(disk + r':\Kis_Activator', exist_ok=True)
    request = requests.get('https://www.comss.ru/download/page.php?id=1874',
                           headers={
                               'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                                             'Chrome/81.0.4044.138 Safari/537.36 OPR/68.0.3618.206', 'accept': '*/*'})
    try:
        request.raise_for_status()
    except requests.exceptions.HTTPError as err:
        print(f'{err} occured')
    soup = BeautifulSoup(request.text, 'html.parser')
    document = soup.select('a')[60]['href']
    request_file = requests.get(document)
    try:
        with open(disk + r':\Kis_Activator\KasperskyInstaller.exe', 'wb') as installer:
            installer.write(request_file.content)
    except FileNotFoundError:
        with open(disk + r':\Kis_Activator\KasperskyInstaller.exe', 'wb') as installer:
            installer.write(request_file.content)
    finally:
        print(fr'Starting to install. If you want to update your Installer Version, please delete the folder '
              f'{disk}:\Kis_Activator ')
        install(disk)


def install(disk):
    """This function installs the prepared (from download() function) Kaspersky Installer file"""
    picture = None
    try:
        pywinauto.Application(backend="uia").start(
            disk + r':\Kis_Activator\KasperskyInstaller.exe')
    except pywinauto.findwindows.ElementNotFoundError:
        print('The Kaspersky installer is launched already')
    except pywinauto.application.AppStartError:
        print('Something wrong with the Application Start.')
    while picture is None:
        # This cycle is made for waiting an appearance of 'Continue' button.
        # Further While cycles are made for the same thing.
        try:
            gui.click()
            gui.click(gui.center(gui.locateOnScreen(resource_path(r'images\Continue1920x1080.png'))))
            break
        except:
            picture = None
    gui.sleep(3)
    gui.click(gui.center(gui.locateOnScreen(resource_path(r'images\prinyat.png'))))
    while picture is None:
        try:
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\Check1920x1080.png"))))
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\Check1920x1080.png"))))
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\skip.png"))))
            gui.sleep(3)
            gui.click(gui.center(gui.locateOnScreen(resource_path(r'images\prinyat2.png'))))
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\Check1920x1080.png"))))
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\Install.png"))))
            break
        except TypeError:
            picture = None
    while picture is None:
        try:
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\anti-advertisement.PNG"))))
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\Look features.PNG"))))
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\Apply.PNG"))))
            gui.sleep(3)
            gui.click(gui.center(gui.locateOnScreen(resource_path(r"images\Done.PNG"))))
            gui.sleep(3)
            break
        except TypeError:
            picture = None
            gui.sleep(10)


if __name__ == "__main__":
    print('Starting the program...')
    if not is_admin():
        # Re-run the program with admin rights
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    else:
        start()
print('The program have been finished')
gui.sleep(5)
