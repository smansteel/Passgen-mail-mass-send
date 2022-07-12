import time
import mouse, keyboard, pandas as pd, pyautogui

pyautogui.FAILSAFE= True

file_name = """C:/Users/USERNAME/Documents/NOFILE.xlsx"""  # path to file + file name
sheet = "Sheet1"  # sheet name or sheet number or list of sheet numbers and names
df = pd.ExcelFile(file_name).parse(sheet)
i= 0 # Start index
ok = False


def is_company(i):
    name_content = df.iat[i, 1][2:]
    if str(name_content).split("@")[1] != "domain.com":
        company =False
    else:
        company = True
    return company

def userselect(i):
    while not is_company(i):
        i+=1
    name_content = df.iat[i, 1][2:]
    name = str(name_content).split("@")[0]
    return i



while True:

    print(pyautogui.position())
    if keyboard.is_pressed("a"):
        exit()
    if keyboard.is_pressed("e") or ok:
        i+=1
        ok = True

        name_content = df.iat[i, 1][2:]
        mdp = df.iat[i, 2]
        print(i)
        name = str(name_content).split("@")[0]
        print(name)

        pyautogui.moveTo(x=550, y=200)
        mouse.click()
        time.sleep(0.25)
        pyautogui.moveTo(x=700, y=330)
        time.sleep(0.25)
        mouse.click()
        pyautogui.hotkey('ctrl','a')
        time.sleep(0.25)
        keyboard.write("\b")
        keyboard.write(name)
        time.sleep(0.25)
        pyautogui.moveTo(x=720, y=180)
        time.sleep(0.5)
        mouse.click()
        pyautogui.moveTo(x=387, y=265)
        time.sleep(1)
        mouse.double_click()
        pyautogui.moveTo(x=1294, y=789)
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(0.1)
        mouse.click()
        time.sleep(1)
        pyautogui.moveTo(x=666, y=366)
        time.sleep(0.5)
        mouse.click()
        keyboard.write(mdp)
        time.sleep(0.5)
        pyautogui.moveTo(x=666, y=480)
        time.sleep(0.5)
        mouse.click()
        keyboard.write(mdp)
        time.sleep(0.5)
        pyautogui.moveTo(x=800, y=180)
        time.sleep(0.1)
        mouse.click()

        time.sleep(1)



