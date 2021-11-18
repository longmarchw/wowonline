from cv2 import CAP_FFMPEG
import pyautogui
import win32gui
import win32api
import win32con
from threading import Timer
import time
import os
import tkinter as tk
import tkinter.messagebox as tkmsgbox

class holdwow_gui():
    cancel_lineup_timer = False
    cancel_jump_timer = False
    tim_lineup = None
    tim_jump = None
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    def set_init_window(self):
        self.init_window_name.title("WOW online") 
        self.init_window_name.geometry('490x90+600+400') 
        self.init_window_name.resizable(0, 0)
        self.check_box_var = tk.IntVar()
        self.radio_jump = tk.Radiobutton(self.init_window_name, text=u"排队+挂机", variable = self.check_box_var, value=1, command=self.on_radiobox_changed)
        self.radio_jump.grid(row=0, column=0, padx=5, pady=10)
        self.radio_lineup = tk.Radiobutton(self.init_window_name, text=u"挂机", variable = self.check_box_var, value=2, command=self.on_radiobox_changed)
        self.radio_lineup.grid(row=0, column=1, padx=5, pady=10)

        self.button_start = tk.Button(self.init_window_name, width=32,height=1, text=u"开始", command=self.starthandle)
        self.button_start.grid(row = 1, column = 0, padx=5, pady=5)
        self.button_stop = tk.Button(self.init_window_name, width=32,height=1, text=u"停止", command=self.stophandle)
        self.button_stop.grid( row = 1, column = 1, padx=5, pady=5)

    def on_radiobox_changed(self):
        return 0
    
    def starthandle(self):
        global cancel_lineup_timer
        global cancel_jump_timer
        if(self.find_window(u"魔兽世界")):
            if self.check_box_var.get()==1:
                self.init_window_name.title(u"WOW online - 排队+挂机...")
                cancel_lineup_timer = False
                self.lineup_handle()
            elif self.check_box_var.get()==2:
                self.init_window_name.title(u"WOW online - 挂机...")
                cancel_jump_timer = False
                self.jump_handle()
            else:
                self.init_window_name.title(u"WOW online")
        else:
            self.init_window_name.title(u"WOW online - no WOW windows found")
            return 0
    
    def stophandle(self):
        global cancel_lineup_timer
        global cancel_jump_timer
        if self.check_box_var.get()==1:
            cancel_lineup_timer =True
            cancel_jump_timer = True
        elif self.check_box_var.get()==2:
            cancel_jump_timer = True
        else:
            return 0
        self.init_window_name.title(u"WOW online")
        
    def find_specify_picture(self, targetpicture):
        character = pyautogui.locateOnScreen(targetpicture, confidence=0.5)#grayscale=True)
        if (character is not None):
            imcenter = pyautogui.center(character)
            print(imcenter)
        else:
            print("none posion find")
        return (imcenter if character is not None else None)

    def find_window(self, targettitle) :
        hwnd=win32gui.FindWindow("GxWindowClass", targettitle)
        if(hwnd):
            #win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 600,300,800,800, win32con.SWP_SHOWWINDOW)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd) 
        return (True if hwnd > 0 else False)

    def lineup_handle(self):
        global tim_lineup
        global cancel_lineup_timer
        if(self.find_window(u"魔兽世界")):
            pyautogui.PAUSE = 1
            clickpos1 = self.find_specify_picture(os.getcwd() + "\\wow_login_in.png")#进入游戏
            if( clickpos1 is not None):#排队完成
                pyautogui.click(clickpos1)
                tim_lineup = None
                cancel_lineup_timer = True
            else:#排队中
                clickpos2 = self.find_specify_picture(os.getcwd() + "\\wow_disconnect.png")#断开连接
                if( clickpos2 is not None):
                    pyautogui.click(clickpos2)
                time.sleep(2)
                clickpos3 = self.find_specify_picture(os.getcwd() + "\\wow_reconnect.png")#重新连接
                if( clickpos3 is not None):
                    pyautogui.click(clickpos3)
                time.sleep(2)
                clickpos4 = self.find_specify_picture(os.getcwd() + "\\wow_disconnect_too.png")#继续断开
                if( clickpos4 is not None):
                    pyautogui.click(clickpos4)
                time.sleep(2)
                tim_lineup = Timer(30, self.lineup_handle)#启动定时
                cancel_lineup_timer = False
                tim_lineup.start()
        if(cancel_lineup_timer == True):#停止排队监测
            if(tim_lineup is not None):
                tim_lineup.cancel()
            time.sleep(10) #开启挂机
            self.jump_handle()
                
    def jump_handle(self):
        global tim_jump
        global cancel_jump_timer
        if(self.find_window(u"魔兽世界")):
            pyautogui.PAUSE = 1
            pyautogui.press("space")
            time.sleep(10)
            if(win32api.GetKeyState(win32con.VK_CAPITAL)==0):
                pyautogui.press("capslock")
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(1)
            pyautogui.typewrite('/CAMP',0.25)
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(30)
            pyautogui.press("enter")
            time.sleep(10)
            tim_jump = Timer(60, self.jump_handle)
            tim_jump.start()
            print(u"tim_jump start") 
        if(cancel_jump_timer == True):
            if(tim_jump is not None):
                tim_jump.cancel()
            print("tim_jump cancel")
    
def on_closing():
    global main_window
    if tkmsgbox.askokcancel(u"退出", u"确定退出程序吗？"):
        main_window.destroy()

def main_dlg():
    global main_window 
    mainfrm = holdwow_gui(main_window)
    mainfrm.set_init_window()
    main_window.protocol("WM_DELETE_WINDOW", on_closing)
    main_window.mainloop()  

if __name__ == "__main__":
    pyautogui.FAILSAFE = False
    main_window = tk.Tk()
    main_dlg()
    #lineup_handle()
    print("wow online exit")