from ast import List
from PIL import Image
import pandas as pd
from pandas import DataFrame, Series
import pyautogui
import pyperclip
import time
import tkinter as tk
from enum import Enum

SCRIPT_NAME = "C:/Users/Tony Kam/Desktop/Auto_Game/Commands.xlsx"

class Type(Enum):
    山河 = 1
    活动刷分 = 2
    善灵推图 = 3
    善灵刷属性 = 4
    浊雾之窟 = 5

class CommandType:
    SINGLE_CLICK = 1
    DOUBLE_CLICK = 2
    RIGHT_CLICK = 3
    INPUT = 4
    WAIT = 5
    SCROLL = 6

class Main:
    cType = -1
    def __init__(self):
        pass
    
    def button_click(self, choice: int, win: tk.Tk):
        self.cType = choice
        win.destroy()
        
    def get_col(self, df: DataFrame, col_index: int) -> list:
        col: Series = df.iloc[:, col_index]
        return col.tolist()

    def loads_script(self):
        df: DataFrame = pd.read_excel(SCRIPT_NAME)
        commandType: Series = self.get_col(df, 0)
        commands: Series = self.get_col(df, 1)
        contents: Series = self.get_col(df, 2)
        jump_list: Series = self.get_col(df, 3)
        id: Series = self.get_col(df, 4)
        return commandType, commands, contents, jump_list, id

    def get_pos(self, img: str):
        image = Image.open(img)
        while True:
            try:
                if pyautogui.locateCenterOnScreen(image, confidence=0.8) is None:
                    print ('waiting ...') 
                elif pyautogui.locateCenterOnScreen(image, confidence=0.8) is not None:
                    print(pyautogui.locateCenterOnScreen(image, confidence=0.8))
                    return pyautogui.locateCenterOnScreen(image, confidence=0.8)
            except Exception as e:
                # print("Exception: ", e.__class__)
                print("Waiting...")
                time.sleep(0.1)
    
    def main(self):
        commandType, commands, contents, jump_list, id = self.loads_script()
        i = 0
        intervalTime = 0.2
        duration = 0.2
        loopCount = 100
        window = tk.Tk()
        window.geometry('300x300')
        
        for a in range (5):
            btn = tk.Button(window, text=f"{a+1}: {Type(a+1).name}", bd='5', command=lambda a=a: self.button_click(a+1, window))
            btn.pack(side='top')
        window.mainloop()
        tempCommandsList = []
        tempContentsList = []
        tempJumpList = []
        for x in commandType:
            if (x == int(self.cType)): # get all the commands for that type
                tempCommandsList.append(commands[commandType.index(x)])
                tempContentsList.append(contents[commandType.index(x)])
                tempJumpList.append(jump_list[commandType.index(x)])
        # Add all the right commands to the temp list and assign the value to the lists
        commands = tempCommandsList
        contents = tempContentsList
        jump_list = tempJumpList
        while loopCount > 0:
            while i < len(commands):
                if (pd.isna(id[i])):
                    loopCount -= 1
                print(f"{i}: Loop Count: {loopCount}, Command: {commands[i]}, Content: {contents[i]}, Jump to: {jump_list[i]}")
                match commands[i]:
                    case CommandType.SINGLE_CLICK:
                        while self.get_pos(contents[i]) is not None:
                            x, y = self.get_pos(contents[i])
                            pyautogui.click(x, y, interval=intervalTime, duration=duration)
                            break
                    case CommandType.DOUBLE_CLICK:
                        while self.get_pos(contents[i]) is not None:
                            x, y = self.get_pos(contents[i])
                            pyautogui.click(x, y, interval=intervalTime, duration=duration, clicks=2)
                            break
                    case CommandType.RIGHT_CLICK:
                        while self.get_pos(contents[i]) is not None:
                            x, y = self.get_pos(contents[i])
                            pyautogui.click(x, y, interval=intervalTime, duration=duration, button='right')
                            break
                    case CommandType.INPUT:
                        pyperclip.copy(contents[i])
                        pyautogui.hotkey('ctrl', 'v')
                    case CommandType.WAIT:
                        pass
                    case CommandType.SCROLL:
                        # pyautogui.scroll(contents[i])
                        print("scroll")
                if pd.isna(jump_list[i]):
                    i += 1
                else:
                    i = int(jump_list[i] - 1)
        
main = Main()
main.main()