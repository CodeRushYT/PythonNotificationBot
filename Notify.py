import datetime
import schedule
import time
import pandas as pd
import subprocess
import pyautogui
import tkinter as tk

df = pd.read_excel(r"C:\Users\YOUR DIRECTORY\Desktop\Python Zoom\TimeTable.xlsx")

def OpenMeeting(ID):
    #Opens the Zoomsdfyes
    print("okie dokie")
    subprocess.Popen(r'C:\Users\YOUR DIRECTORY\AppData\Roaming\Zoom\bin\Zoom.exe')
    time.sleep(1)

    #Join Button
    b = pyautogui.locateCenterOnScreen(r'C:\Users\YOUR DIRECTORY\Desktop\Python Zoom\NotificationThing\JoinButton.png', confidence=0.9)
    pyautogui.click(b)
    time.sleep(0.5)

    #TypeWriter
    pyautogui.typewrite(str(ID))

    #Next Join Button
    c = pyautogui.locateCenterOnScreen(r'C:\Users\YOUR DIRECTORY\Desktop\Python Zoom\NotificationThing\NextJoinButton.png', confidence=0.6)
    pyautogui.click(c)
    time.sleep(2)

def popUpWindow(ID):
    popup = tk.Tk()
    popup.wm_title("test")
    B = tk.Button(popup, text ="Join Class", command =lambda: OpenMeeting(ID))
    B.pack()
    popup.mainloop()
    print("yeet")



schedule.every().monday.at("13:08").do(popUpWindow, "3524235")

for x in range(14):
    if(x % 2) == 0:
        schedule.every().monday.at(str(df.at[x,0])).do(popUpWindow, df.at[x+1,0])


for x in range(7):
    if(x % 2) == 0:
        schedule.every().tuesday.at(str(df.at[x,1])).do(popUpWindow, (df.at[x+1,1]))
        #print(df.at[x,1])
for x in range(6):
    if(x % 2) == 0:
        schedule.every().wednesday.at(str(df.at[x,2])).do(popUpWindow, (df.at[x+1, 2]))
        print(df.at[x,2])


for x in range(8):
    if(x % 2) == 0:
        schedule.every().thursday.at(str(df.at[x,3])).do(popUpWindow, (df.at[x+1,3]))
        print(df.at[x,3])

for x in range(6):
    if(x % 2) == 0:
        schedule.every().friday.at(str(df.at[x,4])).do(popUpWindow, df.at[x+1,4])


while True:
    schedule.run_pending()
    time.sleep(1)
