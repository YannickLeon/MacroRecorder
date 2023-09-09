import time
import win32api, win32gui, win32ui, win32con
import argparse

from action import action

mouseEventDownDict = {
    1: win32con.WM_LBUTTONDOWN,
    2: win32con.WM_RBUTTONDOWN,
    3: win32con.WM_MBUTTONDOWN
}

mouseEventUpDict = {
    1: win32con.WM_LBUTTONUP,
    2: win32con.WM_RBUTTONUP,
    3: win32con.WM_MBUTTONUP
}

def recordNewMacro(fileName, window):
    f = open(fileName, "w")

    keyBoardState = [0 for i in range(256)] # there are 256 key codes
    lastState = win32api.GetKeyboardState()

    start = time.time()

    while not ((win32api.GetKeyState(0xA4) == -128 or win32api.GetKeyState(0xA4) == -127) and win32api.GetKeyState(0x53)):
        offset = window.GetWindowRect()[:2]
        current = win32api.GetKeyboardState()
        if current != lastState:
            for i in range(0,256):
                if current[i:i+1] != b"\x00" and current[i:i+1] != b"\x01" and not keyBoardState[i]:
                    keyBoardState[i] = time.time()
                elif (current[i:i+1] == b"\x00" or current[i:i+1] == b"\x01") and keyBoardState[i] != 0:
                    if i < 3:
                        f.write(f"{i};{keyBoardState[i] - start};{time.time() - start};{win32api.GetCursorPos()[0] - offset[0]};{win32api.GetCursorPos()[1] - offset[1]};\n")
                    else:
                        f.write(f"{i};{keyBoardState[i] - start};{time.time() - start};0;0;\n")
                    keyBoardState[i] = 0
        lastState = win32api.GetKeyboardState()

    f.close()

def processLines(lines):
    actions = []
    for line in lines:
        substr = ""
        data = [] # stucture: [keycode, pressed, released, mouse_x, mouse_y]
        for char in line:
            if char == ";":
                data.append(substr)
                substr = ""
            else:
                substr += char
        actions.append(action(int(data[0]), float(data[1]), float(data[2]), int(data[3]), int(data[4])))
    return actions

def resetActions(actions):
    for action in actions:
        action.state = 0

def playMacro(fileName, window, repetitions = 1):
    print("[*] Reading input file")
    f = open(fileName, "r")
    lines = f.readlines()
    f.close()
    print("[*] Processing lines")
    actions = processLines(lines)
    print("[*] Running macro...")
    counter = 0

    while counter != repetitions:
        completionCount = 0
        start = time.time() - 0.1 # 100 ms puffer für resetActions
        resetActions(actions)
        while completionCount < len(actions):
            if ((win32api.GetKeyState(0xA4) == -128 or win32api.GetKeyState(0xA4) == -127) and win32api.GetKeyState(0x53)):
                print("[*] Terminating macro ...")
                exit()
            for action in actions:
                if start + action.pressed < time.time() and action.state == 0:
                    if int(action.keycode) > 6:
                        window.SendMessage(win32con.WM_KEYDOWN, action.keycode, 0)
                        print(f"pressed key {action.keycode}")
                    else:
                        x_str = ("0000" + hex(action.x)[2:])[-4:]
                        y_str = ("0000" + hex(action.y)[2:])[-4:]
                        pos = int(f"{y_str + x_str}", 16)
                        window.SendMessage(mouseEventDownDict[action.keycode], action.keycode, pos)
                    action.state = 1
                if start + action.released < time.time() and action.state == 1:
                    if int(action.keycode) > 6:
                        window.SendMessage(win32con.WM_KEYUP, action.keycode, 0)
                        print(f"released key {action.keycode}")
                    else:
                        x_str = ("0000" + hex(action.x)[2:])[-4:]
                        y_str = ("0000" + hex(action.y)[2:])[-4:]
                        pos = int(f"{y_str + x_str}", 16)
                        window.SendMessage(mouseEventUpDict[action.keycode], action.keycode, pos)
                    completionCount += 1
                    action.state = 2
        counter += 1
    print(f"Macro completed after {repetitions} repetitions.")

parser = argparse.ArgumentParser(
    description="This MacroRecorder uses the windows api to send fake input to any window. This means it can run completely in the background :)",
    epilog="If you want to cancel a macro just press Alt+Q!"
)
parser.add_argument("mode", choices=["record", "play"], help="Choose the mode in which to run the script.")
parser.add_argument("file", type=str, help="The name of the file used to save or read the macro.")
parser.add_argument("window", type=str, help="The name of the window affected by the macro.")
parser.add_argument("-r", type=int, help="The number of repetitions when playing a macro. -1 = unlimited repetitions")
args = parser.parse_args()
if args.mode == "play" and args.r == None:
    parser.error("Please specify the number of repetitions when using play-mode!")

hwnd = win32gui.FindWindow(None, args.window)
window = win32ui.CreateWindowFromHandle(hwnd)

if args.mode == "record":
    recordNewMacro(args.file, window)
elif args.mode == "play":
    playMacro(args.file, window, args.r)