import win32api
import psutil
import os
import re
import screen_brightness_control as sbc

DEVICE = win32api.EnumDisplayDevices()

POWER_CFG_PATH = 'C:\\Windows\\System32\\powercfg.exe'


def getRefreshRate() -> int:
    devMode = win32api.EnumDisplaySettings(DEVICE.DeviceName, -1)

    displayFrequency = int(getattr(devMode, 'DisplayFrequency'))

    return displayFrequency


def changeRefreshRate(RR: int):
    # if (device.DeviceName != '\\\\.\\DISPLAY1'):
    #     print('Device is not Laptop\'s Display | DEVICE NAME : ' + (device.DeviceName))
    try:
        currRR = getRefreshRate()

        if (RR == currRR):
            print(f'Already on {currRR} Hz ')
            return

        devMode = win32api.EnumDisplaySettings(DEVICE.DeviceName, -1)

        devMode.DisplayFrequency = RR

        win32api.ChangeDisplaySettings(devMode, 0)
    except:
        print('Failed to change screen mode to {RR}Hz')

# PSGUID -> Power Scheme GUID


def getCurrentPowerScheme() -> str:

    return os.popen(f'cmd /c {POWER_CFG_PATH} /GETACTIVESCHEME').read()


def getGUID(PSStr: str) -> str:
    return re.sub(r'\s+|Power Scheme GUID:|\(.*\)', '', PSStr)


def changePowerSettings(PSGUID: str):
    PSGUID = PSGUID.strip()
    currScheme = getCurrentPowerScheme()
    currGUID = getGUID(currScheme)

    print("Current Power Scheme -> " + currScheme)

    if (currGUID == PSGUID):
        print('Power Scheme is already active')
        return

    retCode = os.system(f'cmd /c {POWER_CFG_PATH} /S {PSGUID}')

    if (retCode == 0):
        print("Power Config Changed Successfully")
    else:
        print(f'Failed to change power scheme | RETURN CODE = {retCode} ')


PSGUIDs = {
    'Ultimate Performance': '093a72af-9a90-4ab7-87e1-00baf09bdd0d',
    'Balance': '381b4222-f694-41f0-9685-ff5bb260df2e',
    'Power Saver': '6eba6547-5af8-4144-a5cd-43bd2ba5bb0e'
}


def getCurrentBrightness() -> int:
    return sbc.get_brightness(display=0)[0]


def setBrightness(intensity: int):

    currBrightness = getCurrentBrightness()

    if (currBrightness == intensity):
        print(f"Brightness is already on {currBrightness}")
        return

    sbc.set_brightness(intensity, display=0)


def main():

    onCharger = psutil.sensors_battery().power_plugged

    currBrightness = getCurrentBrightness() + 1

    if (onCharger):
        changeRefreshRate(144)
        changePowerSettings(PSGUIDs['Ultimate Performance'])
    else:
        changeRefreshRate(60)
        changePowerSettings(PSGUIDs['Balance'])

    setBrightness(currBrightness)


main()
