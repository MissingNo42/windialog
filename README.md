# windialog

An API to use the new Windows file & folder Dialogue Explorer


# Description

[![License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)](https://en.wikipedia.org/wiki/MIT_License)

windialog is as single file and standalone module allowing, without any external dependency, the full use of
the native Windows Dialogue Explorer, for **open file(s)** (OPEN), **open folder(s)**, and **save file** (SAVE AS).
All useful features are implemented. (It is like, but more complete than, tkinter.filedialog)

It is written in **pure python / ctypes** avoiding any dependency and supports all useful
features and more. It only depend on ctypes, already installed.

Only Python 3.8 and later are supported (but can be adapted to Python 3 in general)

*(Only tested on CPython 3.9.4 - 64 bits)*


# Supported OS

Because the future is now and Windows XP is dead, only the new API is supported so
ONLY **WINDOWS VISTA AND LATER** ARE SUPPORTED, 32 and 64 bits.

*(Only tested on Windows 10 Pro 21H1 - 64-bits)*


# Installation

You can just copy the file, but [PIP](https://pypi.org/project/windialog/) is also supported
```cmd
pip install windialog
```


# Usage

It provides 3 main functions:
-   getfile: open an Explorer Dialogue for file(s).
-   getsave: open an Explorer Dialogue for save a file.
-   getdir: open an Explorer Dialogue for folder(s), in NewStyle Mode (same interface as getfile() or tkinter.filedialog.askdirectory(), not as the old Win32 one).

```py
from windialog import *

# open an Explorer Dialogue for a file.

f = getfile()
# f is [] if the user cancel, else ["c:\\file\\chose\\by\\user"]


f = getfile(multi = True)
# f can now contains multiple files paths e.g. ["c:\\file1", "c:\\file2", ...]



from tkinter import Tk()
wd = Tk()


# You can also attach the Explorer to a window by passing its Window Handle (HWND) as first argument
# indeed, it work with any GUI lib, not only tkinter (see your GUI lib doc to get HWND)

f = getfile(wd.winfo_id())



# You can also custom many things (exhaustive example)

hwnd = 0 # No window attachment

filetypes = (("Word Documents", "*.docx;*.doc"), # when "Word Documents" is selected in explorer, 
             ("Excel Sheets", "*.xlsx"),         # only .docx and .doc files are visible. ('*' is important)
             ("All Files", "*.*"))

title   = "Select a document" # explorer window title (in title bar)
initdir = "c:\\users\\john\\documents" # default folder where start the explorer
ok_text = "Confirm Please" # text on the OK button
file_text = "Selected Document" # text next to the text-field

f = getfile(hwnd, *filetypes, title = title, initdir = initdir, multi = False, ok_text = ok_text, file_text = file_text)


# A last argument is also available: '__fos_flags' it is only need in extremely special cases (= never)
# and take a binary combination (using |) of flags from windialog.FOS.
# e.g.: getfile(..., __fos_flags = FOS.PICKFOLDERS) is an alias of getdir(...)



# getdir works as getfile, but *filestypes is not supported (logic), and 'file_text' is here 'dir_text'

d = getdir(hwnd, title = title, initdir = initdir, multi = True, ok_text = ok_text, dir_text = file_text)
# d contains [] if the user cancel, else ["c:\\selected\\dir"] or ["c:\\dir1", "c:\\dir2", ...] if multi = True



# getsave works as getfile too, but not supports 'multi' so it return a str and not a list

s = getsave(hwnd, *filetypes, title = title, initdir = initdir, ok_text = ok_text, file_text = file_text)
s # take care that s is an empty string ('') if user cancel, else it is a path str ('c:\\path\\to\\save\\file')


# for more details:

help(getfile)
help(getdir)
help(getsave)
```


# Changelog

## [1.0.2] - First release

First release


# Copyright

Copyright (c) 2021 - MissingNo42

Permission is hereby granted, free of charge, to any person obtaining a copy
of this file, to deal it without any restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
sell copies of the file.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
