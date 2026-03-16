import psutil
import requests
import socket
import time
import win32gui
import win32process

SERVER = "http://Server/update"

computer = socket.gethostname()


def enum_windows():
    apps = []

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):

            title = win32gui.GetWindowText(hwnd)
            # print("title",title)

            if title.startswith("Tekla Structures -") and title != "Tekla Structures - Sign in":

                _, pid = win32process.GetWindowThreadProcessId(hwnd)

                try:
                    process = psutil.Process(pid)
                    app_name = process.name()
                    # print("aaa-------",app_name)
                    if win32gui.IsIconic(hwnd):
                        state = "minimized"
                    else:
                        state = "active"

                    apps.append({
                        "application": app_name,
                        "title": title,
                        "state": state
                    })

                except:
                    pass

    win32gui.EnumWindows(callback, None)

    return apps


while True:
    time.sleep(5)
    try:

        app_list = enum_windows()

        for app in app_list:

            data = {
                "computer": computer,
                "application": app["application"],
                "title": app["title"],
                "state": app["state"]
            }

            requests.post(SERVER, json=data, timeout=3)

    except:
        pass

