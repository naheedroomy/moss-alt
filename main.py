import subprocess
import time
import platform, socket, re, uuid, json, psutil, logging
import os
import pythoncom
import wmi
from datetime import datetime
import pyscreenshot as ImageGrab
import schedule
import pyautogui
import zipfile
import os
import shutil
import time
import PySimpleGUI as sg
import threading
import pyminizip

count = 1
status = True
from mss import mss
from checksumdir import dirhash


import os

DESKTOP_PATH = os.path.expanduser("~\Desktop")
datetime_str = str(datetime.now())
datetime_str_mod = datetime_str.replace(':', '-')
datetime_str_mod_2 = datetime_str_mod.replace('.', '-')
print(DESKTOP_PATH)

def create_folder():
    global datetime_str_mod_2
    mydir = f"{DESKTOP_PATH}\\mossalt\\{datetime_str_mod_2}"
    CHECK_FOLDER = os.path.isdir(mydir)

    if not CHECK_FOLDER:
        os.makedirs(mydir)



def take_screenshot2():
    global count
    global datetime_str_mod_2
    final_dir = f"{DESKTOP_PATH}\\mossalt\\{datetime_str_mod_2}\\"
    final_dir_2 = final_dir + str(count) + ".jpg"
    with mss() as sct:
        sct.shot(mon=-1, output=final_dir_2)
    count += 1


def get_system_info():
    pythoncom.CoInitialize()
    global datetime_str_mod_2
    computer = wmi.WMI()
    gpu_info = computer.Win32_VideoController()[0].name
    info = {}
    try:

        info['platform'] = platform.system()
        info['platform-release'] = platform.release()
        info['platform-version'] = platform.version()
        info['architecture'] = platform.machine()
        info['hostname'] = socket.gethostname()
        info['ip-address'] = socket.gethostbyname(socket.gethostname())
        info['mac-address'] = ':'.join(re.findall('..', '%012x' % uuid.getnode()))
        info['processor'] = platform.processor()
        info['ram'] = str(round(psutil.virtual_memory().total / (1024.0 ** 3))) + " GB"
        info['gpu'] = gpu_info
        final_dir = f"{DESKTOP_PATH}\\mossalt\\{datetime_str_mod_2}\\"
        final_dir_2 = final_dir + "system_info.json"
        with open(final_dir_2, "w") as f:
            f.write(json.dumps(info))
    except Exception as e:
        logging.exception(e)


def get_processes():
    pythoncom.CoInitialize()
    global datetime_str_mod_2
    final_dir = f"{DESKTOP_PATH}\\mossalt\\{datetime_str_mod_2}\\"
    final_dir_2 = final_dir + "processes.txt"
    output = os.popen('wmic process get description, processid').read()
    now = datetime.now().strftime("%H:%M:%S")

    #     write output to file
    with open(final_dir_2, "a") as f:
        f.write("Process snapshot taken at :" + now + "\n")
        f.write(output)


def create_zip_with_password():
    global datetime_str_mod_2
    directory_name = f"{DESKTOP_PATH}\\mossalt\\{datetime_str_mod_2}"
    zip_created_directory = f"{DESKTOP_PATH}\\mossalt\\"
    directory = directory_name
    md5hash = dirhash(directory, 'md5')
    # write md5 value to a textfile in the directory
    with open(f"{directory_name}\\md5.txt", "w") as f:
        f.write(md5hash)

    # use pyminizip to create a zip file with password
    pyminizip.compress(f"{directory_name}\\md5.txt", f"{directory_name}\\checksum.zip", "noneshallpass", 0)

    name_of_archive = f"{zip_created_directory}\\MOSS-{datetime_str_mod_2}"

    shutil.make_archive(name_of_archive, 'zip', directory_name)
    # directory = directory_name
    # md5hash = dirhash(directory, 'md5')
    # print(md5hash)




def schedule_job():
    global status
    create_folder()
    get_system_info()
    schedule.every(2).seconds.do(get_processes)
    schedule.every(2).seconds.do(take_screenshot2)

    while status:
        schedule.run_pending()
        time.sleep(1)


def use_threading():
    pythoncom.CoInitialize()
    thread1 = threading.Thread(target=schedule_job)

    thread1.start()


def stop_job():
    global datetime_str_mod_2
    global status
    status = False
    create_zip_with_password()
    time.sleep(5)
    mydir = f"{DESKTOP_PATH}\\mossalt\\{datetime_str_mod_2}"
    shutil.rmtree(mydir)


def main():
    layout = [[sg.Button('Start'), sg.Button('Stop', disabled=True), sg.Exit(disabled=False)]]
    window = sg.Window('Window Title', layout)

    while True:
        event, values = window.read(timeout=10)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Start':
            window['Start'].update(disabled=True)
            window['Stop'].update(disabled=False)
            # disable the exit button
            window['Exit'].update(disabled=True)
            use_threading()
        elif event == 'Stop':
            stop_job()
            window['Stop'].update(disabled=True)
            window.refresh()
            time.sleep(7)
            window['Exit'].update(disabled=False)
        elif event == '-FUNCTION COMPLETED-':
            sg.popup('Your function completed!')
    window.close()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
