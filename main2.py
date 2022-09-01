import subprocess
import time
import platform, socket, re, uuid, json, psutil, logging
import wmi
from datetime import datetime
import pyscreenshot as ImageGrab
import schedule
import pyautogui
import zipfile
import os
import shutil
import time

count = 1
from mss import mss

import os

# You should change 'test' to your preferred folder.
datetime_str = str(datetime.now())
datetime_str_mod = datetime_str.replace(':', '-')
datetime_str_mod_2 = datetime_str_mod.replace('.', '-')
mydir = f"C:\\Users\\Administrator\\Desktop\\test\\{datetime_str_mod_2}"
CHECK_FOLDER = os.path.isdir(mydir)

if not CHECK_FOLDER:
    os.makedirs(mydir)
    # print("created folder : ", MYDIR)


# def take_screenshot():
#     global count
#     datetime_str = datetime.now()
#     time = datetime_str.time()
#     time_str = time.strftime("%H:%M:%S")
#     time_str = time_str.replace(":", "-")
#     image_name = str(time) + ".png"
#     screenshot = ImageGrab.grab()
#     dir_output = f"C:\\Users\\Administrator\\Desktop\\test\\{datetime_str_mod_2}\\"
#     filepath2 = dir_output + str(count) + ".png"
#
#     screenshot.save(filepath2)
#     count += 1
#
#     return filepath2



def take_screenshot2():
    global count
    final_dir = dir_output = f"C:\\Users\\Administrator\\Desktop\\test\\{datetime_str_mod_2}\\"
    final_dir_2 = final_dir + str(count) + ".jpg"
    with mss() as sct:
        sct.shot(mon=-1, output=final_dir_2)
    count += 1


def get_system_info():
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
        final_dir = f"C:\\Users\\Administrator\\Desktop\\test\\{datetime_str_mod_2}\\"
        final_dir_2 = final_dir + "system_info.json"
        with open(final_dir_2, "w") as f:
            f.write(json.dumps(info))
    except Exception as e:
        logging.exception(e)

def get_processes():
    final_dir = f"C:\\Users\\Administrator\\Desktop\\test\\{datetime_str_mod_2}\\"
    final_dir_2 = final_dir + "processes.txt"
    output = os.popen('wmic process get description, processid').read()
    now = datetime.now().strftime("%H:%M:%S")

#     write output to file
    with open(final_dir_2, "a") as f:

        f.write("Process snapshot taken at :" + now + "\n")
        f.write(output)

def create_zip_with_password():
    dir_output = f"C:\\Users\\Administrator\\Desktop\\test\\{datetime_str_mod_2}"
    zip_file_name = dir_output + ".zip"
    password = "test"
    zip_file = zipfile.ZipFile(zip_file_name, "w", zipfile.ZIP_DEFLATED)
    for folder, subfolders, files in os.walk(dir_output):
        for file in files:
            zip_file.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder, file), dir_output),
                           compress_type=zipfile.ZIP_DEFLATED)
    zip_file.setpassword(password.encode("utf-8"))
    zip_file.close()
    shutil.rmtree(dir_output)

def main2():
    create_zip_with_password()


#
# Press the green button in the gutter to run the script.
if __name__ == '__main2__':
    main2()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
