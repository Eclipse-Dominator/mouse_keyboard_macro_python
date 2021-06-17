from re import findall
import win32con
from win32api import mouse_event, keybd_event, GetKeyState
from win32gui import GetForegroundWindow, GetWindowText
from time import sleep
from random import uniform
from gc import collect

DEFAULT_CLICK_TIME = 0.03


def nullfunction():
    return


def currentWindowName(): return GetWindowText(GetForegroundWindow()).upper()


def parseTime(string):
    data = [x.strip() for x in string.split("-")]
    if len(data) == 2:
        return f"uniform({float(data[0])},{float(data[1])})"
    elif len(data) == 1:
        return float(data[0])
    raise "error"


mouse_switcher = {
    "LEFT": {
        "UP": win32con.MOUSEEVENTF_LEFTUP,
        "DOWN": win32con.MOUSEEVENTF_LEFTDOWN,
        "VK": 0x01
    },
    "RIGHT": {
        "UP": win32con.MOUSEEVENTF_RIGHTUP,
        "DOWN": win32con.MOUSEEVENTF_RIGHTDOWN,
        "VK": 0x02
    },
    "MIDDLE": {
        "UP": win32con.MOUSEEVENTF_MIDDLEUP,
        "DOWN": win32con.MOUSEEVENTF_MIDDLEDOWN,
        "VK": 0x04
    }
}

key_switcher = {
    "SPACE": win32con.VK_SPACE,
    "LSHIFT": win32con.VK_LSHIFT,
    "RSHIFT": win32con.VK_RSHIFT,
    "SHIFT": win32con.VK_SHIFT,
    "LCTRL": win32con.VK_LCONTROL,
    "RCTRL": win32con.VK_RCONTROL,
    "CTRL": win32con.VK_CONTROL,
    "ESC": win32con.VK_ESCAPE,
    "END": win32con.VK_END,
    "RETURN": win32con.VK_RETURN,
    "TAB": win32con.VK_TAB,
    "CAPS": win32con.VK_CAPITAL,
    "UP": win32con.VK_UP,
    "DOWN": win32con.VK_DOWN,
    "LEFT": win32con.VK_LEFT,
    "RIGHT": win32con.VK_RIGHT,
    "X1": win32con.VK_XBUTTON1,
    "X2": win32con.VK_XBUTTON2,
}

dictionarySwitcher = {}


def getKey(input):
    if len(input) != 2:
        raise "invalid bind"
    if input[0] == "KEY":
        return key_switcher.get(input[1]) or ord(input[1])
    elif input[0] == "MOUSE":
        mouseobj = mouse_switcher.get(input[1])
        if mouseobj:
            return mouseobj["VK"]
        else:
            return key_switcher[input[1]]
    raise "invalid bind"


with open("macros.setting", 'r') as f:
    BINDINGFUNCTION = False
    currentfunction = ""
    tempFunc = None
    mouseUpDown = [nullfunction, nullfunction]  # KeyUp, KeyDown
    keyBind = None
    keyUp = False
    for line in f.readlines():
        command = findall(r'".+?"|\S+', line.strip().upper())
        commandLength = len(command)
        if commandLength == 0 or command[0][0] == "#":  # comment
            continue
        if BINDINGFUNCTION:
            if command[0] == "TIME":
                currentfunction += f"\tsleep({parseTime(command[1])})\n"
            elif command[0] == "MOUSE" and commandLength == 3:
                if command[2] == "CLICK":
                    currentfunction += f"\tmouse_event({mouse_switcher[command[1]]['DOWN']},0,0)\n"
                    currentfunction += f"\tsleep({DEFAULT_CLICK_TIME})\n"
                    currentfunction += f"\tmouse_event({mouse_switcher[command[1]]['UP']},0,0)\n"
                else:
                    currentfunction += f"\tmouse_event({mouse_switcher[command[1]][command[2]]},0,0)\n"
            elif command[0] == "KEY" and commandLength == 3:
                keystatus = ""
                if command[2] == "UP":
                    keystatus = win32con.KEYEVENTF_KEYUP
                elif command[2] == "DOWN":
                    keystatus = 0
                elif command[2] == "CLICK":
                    currentfunction += f"\tkeybd_event({key_switcher.get(command[1]) or ord(command[1])},0,0,0)\n"
                    currentfunction += f"\tsleep({DEFAULT_CLICK_TIME})\n"
                    currentfunction += f"\tkeybd_event({key_switcher.get(command[1]) or ord(command[1])},0,{win32con.KEYEVENTF_KEYUP},0)\n"
                    continue
                else:
                    raise "unknown status"
                currentfunction += f"\tkeybd_event({key_switcher.get(command[1]) or ord(command[1])},0,{keystatus},0)\n"
            elif command[0] == "ENDBIND":
                currentfunction += "\treturn\n"
                exec(currentfunction, globals())
                currentfunction = ""
                mouseUpDown = dictionarySwitcher.get(keyBind) or mouseUpDown
                if keyUp:
                    mouseUpDown[0] = tempFunc
                else:
                    mouseUpDown[1] = tempFunc
                dictionarySwitcher[keyBind] = mouseUpDown
                BINDINGFUNCTION = False
                tempFunc = None
                keyUp = False
                mouseUpDown = [nullfunction, nullfunction]
        elif command[0][:4] == "BIND":  # command can be BIND, BINDUP, BINDDOWN
            if BINDINGFUNCTION:
                raise "Invalid Code!"
            BINDINGFUNCTION = True
            keyBind = getKey(command[1:3])
            currentfunction += f"def tempFunc():\n"
            if commandLength == 4:  # if window name filter is present
                currentfunction += "\tif '{}' not in currentWindowName(): return\n".format(
                    command[3].strip('\"'))
            print(command[0][4:])
            if command[0][4:] == "UP":
                keyUp = True
            elif command[0][4:] == "DOWN" or command[0][4:] == "":
                keyUp = False
            else:
                raise "Invalid Code!"

        elif command[0] == "SET":
            if command[1] == "DEFAULT_CLICK_TIME":
                DEFAULT_CLICK_TIME = parseTime(command[2])
            else:
                raise "unknown variable"
        else:
            raise "unknown command"

del key_switcher
del mouse_switcher
del BINDINGFUNCTION
del currentfunction
del tempFunc
del keyBind
del command
del commandLength
del f
del line
del mouseUpDown
del keyUp
collect()
print(dictionarySwitcher)

clickedState = [-127, -128]
trackerwithstate = [[key, GetKeyState(key)] for key in dictionarySwitcher]


def checkEvent(keycode, prevState, onPress=nullfunction, onRelease=nullfunction):
    new_state = GetKeyState(keycode)
    if new_state != prevState:
        if new_state in clickedState:
            onPress()
        else:
            onRelease()

    return new_state


while 1:
    for tracked in trackerwithstate:
        tracked[1] = checkEvent(tracked[0], tracked[1], onRelease=dictionarySwitcher[tracked[0]][0],
                                onPress=dictionarySwitcher[tracked[0]][1])
    sleep(0.001)
