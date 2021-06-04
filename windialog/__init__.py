from ctypes import c_buffer, c_void_p, pointer, byref, sizeof, cast, WINFUNCTYPE, POINTER, ARRAY, HRESULT, windll
from ctypes.wintypes import DWORD, HWND, UINT, LPWSTR

from typing import Tuple, List


__all__ = ('FOS', 'getdir', 'getfile', 'getsave')

__version__ = "1.0.2"
__author__ = "MissingNo42"
__doc__ = """see windialog on PyPI for doc and license"""


ClsidFileOpenDialog = c_buffer(b'\x9c\x5a\x1c\xdc\x8a\xe8\xde\x4d\xa5\xa1\x60\xf8\x2a\x20\xae\xf7', 16)
ClsidFileSaveDialog = c_buffer(b'\xf3\xe2\xb4\xc0\x21\xba\x73\x47\x8d\xba\x33\x5e\xc9\x46\xeb\x8b', 16)
IIDIFileOpenDialog  = c_buffer(b'\x88\x72\x7c\xd5\xad\xd4\x68\x47\xbe\x02\x9d\x96\x95\x32\xd9\x60', 16)
IIDIFileSaveDialog  = c_buffer(b'\x23\xcd\xbc\x84\xde\x5f\xdb\x4c\xae\xa4\xaf\x64\xb8\x3d\x78\xab', 16)
IIDIShellItem       = c_buffer(b'\x1e\x6d\x82\x43\x18\xe7\xee\x42\xbc\x55\xa1\xe2\x61\xc3\x7b\xfe', 16)

CLSCTX_INPROC_SERVER = 0x1
SIGDN_FILESYSPATH    = DWORD(0x80058000)

class FOS():
    """see https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/ne-shobjidl_core-_fileopendialogoptions"""
    
    OVERWRITEPROMPT	     = 0x2
    STRICTFILETYPES	     = 0x4
    NOCHANGEDIR	             = 0x8
    PICKFOLDERS	             = 0x20
    FORCEFILESYSTEM	     = 0x40
    ALLNONSTORAGEITEMS	     = 0x80
    NOVALIDATE	             = 0x100
    ALLOWMULTISELECT	     = 0x200
    PATHMUSTEXIST	     = 0x800
    FILEMUSTEXIST	     = 0x1000
    CREATEPROMPT	     = 0x2000
    SHAREAWARE	             = 0x4000
    NOREADONLYRETURN	     = 0x8000
    NOTESTFILECREATE	     = 0x10000
    HIDEMRUPLACES	     = 0x20000
    HIDEPINNEDPLACES 	     = 0x40000
    NODEREFERENCELINKS       = 0x100000
    OKBUTTONNEEDSINTERACTION = 0x200000
    DONTADDTORECENT	     = 0x2000000
    FORCESHOWHIDDEN	     = 0x10000000
    DEFAULTNOMINIMODE	     = 0x20000000
    FORCEPREVIEWPANEON	     = 0x40000000
    SUPPORTSTREAMABLEITEMS   = 0x80000000


ole32 = windll.ole32
CoCreateInstance = ole32.CoCreateInstance
CoTaskMemFree = ole32.CoTaskMemFree
CoInitialize = ole32.CoInitialize

PSIZE = sizeof(c_mem_p := POINTER(c_void_p))


def is_null_ptr(p) -> bool:
    """[Internal] Return if the passed pointer is NULL"""
    
    return not cast(p, c_void_p).value


def _free(COM: c_mem_p, DIR: c_mem_p, mult: c_mem_p):
    """[Internal] Free interfaces' memory"""
    
    if not is_null_ptr(mult): cast(mult.contents.value + 2 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p))).contents(mult)
    if not is_null_ptr(DIR):  cast( DIR.contents.value + 2 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p))).contents(DIR)
    if not is_null_ptr(COM):  cast( COM.contents.value + 2 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p))).contents(COM)


def _extformat(ext: Tuple[Tuple[str, str], ...]):
    """[Internal] Try to correct erroneous extension format, avoiding unspecified/unwanted behaviors"""
    
    for n, e in ext:
        x = []
        for i in e.split(";"):
            if not i or i in '*.*': i = "*.*"
            if i.endswith("."): i += "*"
            if i.startswith("."): i = "*" + i
            if i.startswith("*."): x.append(i)
            elif i.startswith("*"): x.append("*." + i[1:])
            else: x.append("*." + i)
        yield (n, ";".join(x))


def _extcheck(path: str, ext: str):
    """[Internal] Check if a path's extension is in 'ext'"""

    return path.endswith(tuple(ext.replace("*.*", "").replace("*", "").split(";")))

    
def _getpaths(windowId: int = 0, title: str = None, initdir: str = None, multi: bool = False,
              ok_text: str = None, path_text: str = None, addflags: int = 0, *filetypes: Tuple[str, str], savemode: bool = False) -> List[str]:
    """[Internal] A pure ctypes Windows COM API (Windows Vista and later) which almost fully control the Path Explorer Interface"""
    
    flags = DWORD()
    COM, DIR, item, mult = c_mem_p(), c_mem_p(), c_mem_p(), c_mem_p() # pointer on struct containing single pointer
        
    try:
        if CoInitialize(None) < 0: raise
        if CoCreateInstance(byref(ClsidFileSaveDialog if savemode else ClsidFileOpenDialog), None, CLSCTX_INPROC_SERVER,
                            byref(IIDIFileSaveDialog  if savemode else IIDIFileOpenDialog), byref(COM)) < 0 or is_null_ptr(COM): raise

        # COM.content.value points to a struct containing only pointer, thus, factors before PSIZE are functions indexes (PSIZE = size (in bytes) of a pointer)
        Show        = cast(COM.contents.value +  3 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, HWND))).contents
        SetFileType = cast(COM.contents.value +  4 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, UINT, POINTER(LPWSTR * 2)))).contents
        SetFileIdx  = cast(COM.contents.value +  5 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, UINT))).contents
        GetFileIdx  = cast(COM.contents.value +  6 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, POINTER(UINT)))).contents
        SetOptions  = cast(COM.contents.value +  9 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, DWORD))).contents
        GetOptions  = cast(COM.contents.value + 10 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, POINTER(DWORD)))).contents
        SetFolder   = cast(COM.contents.value + 12 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, c_mem_p))).contents
        SetTitle    = cast(COM.contents.value + 17 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, LPWSTR))).contents
        SetOkBtnTxt = cast(COM.contents.value + 18 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, LPWSTR))).contents
        SetFnLabel  = cast(COM.contents.value + 19 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, LPWSTR))).contents
        GetResult   = cast(COM.contents.value + 20 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, POINTER(c_mem_p)))).contents
        GetResults  = cast(COM.contents.value + 27 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, POINTER(c_mem_p)))).contents

        # Config flags
        GetOptions(COM, byref(flags))

        flags.value |= FOS.FORCEFILESYSTEM | FOS.PATHMUSTEXIST | FOS.FILEMUSTEXIST | (bool(multi) and FOS.ALLOWMULTISELECT) | addflags

        if SetOptions(COM, flags) < 0: raise

        # Set file types
        if filetypes:
            SetFileType(COM, len(filetypes), ARRAY(LPWSTR * 2, len(filetypes))(*[ARRAY(LPWSTR, 2)(*i) for i in _extformat(filetypes)]))
            SetFileIdx(COM, 1)

        # Other optional config
        if title: SetTitle(COM, LPWSTR(title))
        if ok_text: SetOkBtnTxt(COM, LPWSTR(ok_text))
        if path_text: SetFnLabel(COM, LPWSTR(path_text))
        if initdir and windll.shell32.SHCreateItemFromParsingName(LPWSTR(initdir), None, byref(IIDIShellItem), byref(DIR)) >= 0: SetFolder(COM, DIR)

        # Show Explorer
        if Show(COM, HWND(windowId)) < 0: raise
        if windowId: windll.user32.EnableWindow(HWND(windowId), 1)


        paths = [] # Final list
        path = LPWSTR()
        
        if savemode: # Get selected path
            if GetResult(COM, byref(item)) < 0: raise
            
            GetName = cast(item.contents.value + 5 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, DWORD, POINTER(LPWSTR)))).contents
            Release = cast(item.contents.value + 2 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p))).contents

            if GetName(item, SIGDN_FILESYSPATH, byref(path)) < (Release(item) and 0): raise # Store path in 'path' and Free 'item' memory

            paths.append(path.value)
            CoTaskMemFree(path) # Free memory

            
            GetFileIdx(COM, byref(n := UINT()))
            n = n.value - 1
            
            return [paths[0] + n.replace("*.*", "").replace("*", "").split(";")[0]] if len(filetypes) > n >= 0 and not _extcheck(paths[0], n := filetypes[n][1]) else paths
                                                                                                                           
        # Get selected path(s)
        if GetResults(COM, byref(mult)) < 0: raise

        GetCount   = cast(mult.contents.value + 7 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, POINTER(DWORD)))).contents
        GetItemAt  = cast(mult.contents.value + 8 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, DWORD, POINTER(c_mem_p)))).contents
        GetName    = None
        Release    = None            
        
        pathsnum = DWORD()

        if GetCount(mult, byref(pathsnum)) < 0 or not pathsnum.value: raise # get results number

        for i in range(pathsnum.value):
            if GetItemAt(mult, i, byref(item)) < 0: break # Init 'item' from 'mult'
            
            if not GetName: # Get it once, same for all 'item'
                GetName = cast(item.contents.value + 5 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p, DWORD, POINTER(LPWSTR)))).contents
                Release = cast(item.contents.value + 2 * PSIZE, POINTER(WINFUNCTYPE(HRESULT, c_mem_p))).contents
                
            if GetName(item, SIGDN_FILESYSPATH, byref(path)) < (Release(item) and 0): break # Store path in 'path' and Free 'item' memory
            paths.append(path.value)
            CoTaskMemFree(path) # Free memory

        return paths
    
    except OSError as e: return []
    finally: _free(COM, DIR, mult) # Free memory, on success or fail
        

        
def getdir(windowId: int = 0, *, title: str = None, initdir: str = None, multi: bool = False, ok_text: str = None, dir_text: str = None, __fos_flags: int = 0) -> List[str]:
    """[Windows Vista and later only]

    getdir(): open the Windows' Directory Explorer and retrieves selected directories as a 'list' of 'str'.
    The 'list' returned contains from 0 to 1 or more 'str' (if 'multi' = False, or True).
    It contains 0 'str' only if the user cancels or if, for any reason, the Windows COM API is not available or denied (never in practice),
    in particular, this function never crashes.

    --- positional arguments ---
    windowId (int): the window identifier (hwnd) to attach the Explorer, no attachment (0) by default.

    --- keyword-only arguments ---
    title    (str): the text displayed in the title bar of the Explorer, if None (by default), Windows displays the standard translated text.
    initdir  (str): the path to the initial directory to explore, if None (by default) Windows chooses the last directory explored or the 'Computer' root.
    multi   (bool): if True, the selection of multiple directories is allowed, or denied otherwise (by default).
    ok_text  (str): a text to replace the Ok-Button's text.
    dir_text (str): a text to replace the Text-Field's text.


    /!\\ ONLY IF NECESSARY
    __fos_flags (int):  additional flags from FOS class, FOS.FORCEFILESYSTEM | FOS.PATHMUSTEXIST | FOS.FILEMUSTEXIST | FOS.PICKFOLDERS are already set.
    """
    
    return _getpaths(windowId, title, initdir, multi, ok_text, dir_text, FOS.PICKFOLDERS | __fos_flags)

    
def getfile(windowId: int = 0, *filetypes: Tuple[str, str], title: str = None, initdir: str = None, multi: bool = False, ok_text: str = None, file_text: str = None, __fos_flags: int = 0) -> List[str]:
    """[Windows Vista and later only]

    getfile(): open the Windows' Open File Explorer and retrieves selected files as a 'list' of 'str'.
    The 'list' returned contains from 0 to 1 or more 'str' (if 'multi' = False, or True).
    It contains 0 'str' only if the user cancels or if, for any reason, the Windows COM API is not available or denied (never in practice),
    in particular, this function never crashes.

    --- positional arguments ---
    windowId (int): the window identifier (hwnd) to attach the Explorer, no attachment (0) by default.
    *filetypes    : filters used to restrict to some file types, each filter must be a couple of 'str', the first 'str' is the displayed type name, the
    second contains the file type extension(s) started by a star '*': e.g. ("All Files", "*.*"), ("Word Document", "*.docx"), ("JPG Image", "*.jpg;*.jpeg").

    --- keyword-only arguments ---
    title    (str): the text displayed in the title bar of the Explorer, if None (by default), Windows displays the standard translated text.
    initdir  (str): the path to the initial directory to explore, if None (by default) Windows chooses the last directory explored or the 'Computer' root.
    multi   (bool): if True, the selection of multiple files is allowed, or denied otherwise (by default).
    ok_text  (str): a text to replace the Ok-Button's text.
    dir_text (str): a text to replace the Text-Field's text.


    /!\\ ONLY IF NECESSARY
    __fos_flags (int):  additional flags from FOS class, FOS.FORCEFILESYSTEM | FOS.PATHMUSTEXIST | FOS.FILEMUSTEXIST are already set.
    """
    
    return _getpaths(windowId, title, initdir, multi, ok_text, file_text, __fos_flags, *filetypes)


def getsave(windowId: int = 0, *filetypes: Tuple[str, str], title: str = None, initdir: str = None, ok_text: str = None, file_text: str = None, __fos_flags: int = 0) -> str:
    """[Windows Vista and later only]

    getfile(): open the Windows' Save File Explorer and retrieves the selected file as a single 'str'.
    An empty 'str' ('') is returned if the user cancels or if, for any reason, the Windows COM API is not available or denied (never in practice),
    in particular, this function never crashes.

    --- positional arguments ---
    windowId (int): the window identifier (hwnd) to attach the Explorer, no attachment (0) by default.
    *filetypes    : filters used to restrict to some file types, each filter must be a couple of 'str', the first 'str' is the displayed type name, the
    second contains the file type extension(s) started by a star '*': e.g. ("All Files", "*.*"), ("Word Document", "*.docx"), ("JPG Image", "*.jpg;*.jpeg").

    --- keyword-only arguments ---
    title    (str): the text displayed in the title bar of the Explorer, if None (by default), Windows displays the standard translated text.
    initdir  (str): the path to the initial directory to explore, if None (by default) Windows chooses the last directory explored or the 'Computer' root.
    ok_text  (str): a text to replace the Ok-Button's text.
    dir_text (str): a text to replace the Text-Field's text.


    /!\\ ONLY IF NECESSARY
    __fos_flags (int):  additional flags from FOS class, FOS.FORCEFILESYSTEM | FOS.PATHMUSTEXIST | FOS.FILEMUSTEXIST and FOS.STRICTFILETYPES (if filetypes) are already set.
    """
    
    return n[0] if (n := _getpaths(windowId, title, initdir, False, ok_text, file_text, (bool(filetypes) and FOS.STRICTFILETYPES) | __fos_flags, *filetypes, savemode = True)) else ''
    







