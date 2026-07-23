import ctypes, ctypes.util, sys

x11 = ctypes.CDLL(ctypes.util.find_library("X11") or "libX11.so.6")
x11.XOpenDisplay.restype = ctypes.c_void_p
x11.XOpenDisplay.argtypes = [ctypes.c_char_p]
x11.XMoveWindow.argtypes = [ctypes.c_void_p, ctypes.c_ulong, ctypes.c_int, ctypes.c_int]
x11.XFlush.argtypes = [ctypes.c_void_p]
x11.XCloseDisplay.argtypes = [ctypes.c_void_p]

d = x11.XOpenDisplay(None)
if not d:
    print("XOpenDisplay FAILED"); sys.exit(1)

win = 0x600081   # 統一 OCR 辨識工具
# 移到左主螢幕左上角 (+0+0)，確保可見
x11.XMoveWindow(d, win, 0, 0)
x11.XFlush(d)
x11.XCloseDisplay(d)
print("MOVED window 0x%X to +0+0" % win)
