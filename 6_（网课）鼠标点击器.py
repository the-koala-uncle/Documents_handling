#! python3 
import pyautogui,time,random
for e in range(50):
    pyautogui.click(1673,1033)

    pyautogui.click(87,1033)
    time.sleep(2)
    pyautogui.click(520,525)
    time.sleep(1)
    pyautogui.click(520,585)
    time.sleep(1)
    pyautogui.click(650,835)
    time.sleep(1)
    pyautogui.click(750,835)
    time.sleep(1)
##    pyautogui.click(930,825)
##    time.sleep(1)
##    pyautogui.click(1060,825)
##    time.sleep(1)
##    pyautogui.click(960,795)
##    time.sleep(1)
##    pyautogui.click(1060,795)
    time_num=random.choice([360,420,460,445,500])
    for e in range(time_num,0,-1):
        time.sleep(1)
        if e==12:
            pyautogui.click(1673,1033)
            print('即将开始下一次点击···')
            time.sleep(2)
            pyautogui.click(87,1033)
        print('距离下一次点击···'+str(e))
