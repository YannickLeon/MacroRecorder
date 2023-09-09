import sys
import win32api, win32gui, win32con
import time
from action import action

mouseEventDownDict = {
    1: win32con.MOUSEEVENTF_LEFTDOWN,
    2: win32con.MOUSEEVENTF_RIGHTDOWN,
    3: win32con.MOUSEEVENTF_MIDDLEDOWN
}

mouseEventUpDict = {
    1: win32con.MOUSEEVENTF_LEFTUP,
    2: win32con.MOUSEEVENTF_RIGHTUP,
    3: win32con.MOUSEEVENTF_MIDDLEUP
}

def recordNewMacro(fileName):
    f = open(fileName + ".macro", "w")

    keyBoardState = [0 for i in range(256)] # there are 256 key codes
    lastState = win32api.GetKeyboardState()

    start = time.time()

    while not ((win32api.GetKeyState(0xA4) == -128 or win32api.GetKeyState(0xA4) == -127) and win32api.GetKeyState(0x53)):
        current = win32api.GetKeyboardState()
        if current != lastState:
            for i in range(0,256):
                if current[i:i+1] != b"\x00" and current[i:i+1] != b"\x01" and not keyBoardState[i]:
                    keyBoardState[i] = time.time()
                elif (current[i:i+1] == b"\x00" or current[i:i+1] == b"\x01") and keyBoardState[i] != 0:
                    if i < 3:
                        f.write(f"{i};{keyBoardState[i] - start};{time.time() - start};{win32api.GetCursorPos()[0]};{win32api.GetCursorPos()[1]};\n")
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

def playMacro(fileName, repetitions, delay=0):
    print("[*] Reading input file")
    f = open(fileName, "r")
    lines = f.readlines()
    f.close()
    print("[*] Processing lines")
    actions = processLines(lines)
    time.sleep(delay)
    print("[*] Running macro...")
    counter = 0

    while str(counter) != repetitions:
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
                        win32api.keybd_event(action.keycode, action.keycode, 0, 0)
                        print(f"{chr(action.keycode)} pressed")
                    else:
                        win32api.SetCursorPos((action.x, action.y))
                        win32api.mouse_event(mouseEventDownDict[action.keycode], 0, 0, 0, 0)
                        print(f"Mouse pressed at {action.x}:{action.y}")
                    action.state = 1
                if start + action.released < time.time() and action.state == 1:
                    if int(action.keycode) > 6:
                        win32api.keybd_event(action.keycode, action.keycode, 2, 0)
                        print(f"{chr(action.keycode)} released")
                    else:
                        win32api.SetCursorPos((action.x, action.y))
                        win32api.mouse_event(mouseEventUpDict[action.keycode], 0, 0, 0, 0)
                        print(f"Mouse released at {action.x}:{action.y}")
                    completionCount += 1
                    action.state = 2
        counter += 1
    print(f"Macro completed after {repetitions} repetitions.")

if len(sys.argv) < 3:
    print("Usage: python macroRecorder <mode> <fileName> <repetitions> <initialDelay>")
    exit()

if sys.argv[1] == "-r":
    recordNewMacro(sys.argv[2])
if sys.argv[1] == "-p":
    playMacro(sys.argv[2], sys.argv[3])