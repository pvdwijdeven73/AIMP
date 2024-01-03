import os
from fnmatch import fnmatch
from datetime import datetime

MAX_DAYS_OLD = 2 * 365
set_time = datetime(
    year=2024, month=1, day=1, hour=12, minute=34, second=56
).timestamp()

root = r"C:\Users\p.vandewijdeven\OneDrive - Shell\PY\AIMP\Projects\CEOD\Displays\CDA"


def show_files():
    pattern = "*.*"
    total = 0
    for path, subdirs, files in os.walk(root):
        for name in files:
            if fnmatch(name, pattern):
                fname = os.path.join(path, name)
                ftime = datetime.fromtimestamp(os.path.getmtime(fname)) - datetime.now()
                if -1 * ftime.days > MAX_DAYS_OLD:
                    total += 1
                    print(fname, ftime.days)

    print(f"{total} files found")


def change_files():
    pattern = "*.*"
    for path, subdirs, files in os.walk(root):
        for name in files:
            if fnmatch(name, pattern):
                fname = os.path.join(path, name)
                ftime = datetime.fromtimestamp(os.path.getmtime(fname)) - datetime.now()
                if -1 * ftime.days > MAX_DAYS_OLD:
                    os.utime(fname, (set_time, set_time))
                    print(fname)


show_files()
change_files()
