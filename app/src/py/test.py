from pywinauto import Desktop

windows = Desktop(backend="uia").windows()
for win in windows:
    print(win.window_text())
