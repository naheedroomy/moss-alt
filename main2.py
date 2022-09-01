import shutil
import os
from datetime import datetime

DESKTOP_PATH = os.path.expanduser("~\Desktop")
datetime_str = str(datetime.now())
datetime_str_mod = datetime_str.replace(':', '-')
datetime_str_mod_2 = datetime_str_mod.replace('.', '-')


directory_name = f"{DESKTOP_PATH}\\mossalt\\{datetime_str_mod_2}"
name_of_archive = f"{directory_name}\\MOSS-{datetime_str_mod_2}"
shutil.make_archive("MOSS-", 'zip', directory_name)