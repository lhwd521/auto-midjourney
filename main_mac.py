import time
import pandas as pd
import pyautogui
import pyperclip
from tkinter import filedialog, Tk, Label, Entry, Button, StringVar
from pynput import mouse, keyboard

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("midjourney自动产出器 by 老陆 vx:laolu2045")
        self.root.geometry("900x500")  # adjust the size of the window

        self.excel_path = StringVar()
        self.wait_time_after_enter = StringVar()
        self.wait_time_after_paste = StringVar()
        self.cmd_sum = StringVar()
        self.wait_time_after_onece = StringVar()
        self.click_position = None
        self.listener = None
        self.running = False

        Label(root, text="Excel路径").grid(row=0)
        Label(root, text="发送命令间隔（秒）【推荐2】").grid(row=1)
        Label(root, text="粘贴后等待时间（秒）【推荐2】").grid(row=2)
        Label(root, text="一次发几条【推荐10】").grid(row=3)
        Label(root, text="等待时间（秒）【推荐800】").grid(row=4)
        Label(root, text="Excel只读取第1列从上到下的数据").grid(row=5)
        Label(root, text="按ESC键终止程序").grid(row=6)

        Entry(root, textvariable=self.excel_path).grid(row=0, column=1)
        Entry(root, textvariable=self.wait_time_after_enter).grid(row=1, column=1)
        Entry(root, textvariable=self.wait_time_after_paste).grid(row=2, column=1)
        Entry(root, textvariable=self.cmd_sum).grid(row=3, column=1)
        Entry(root, textvariable=self.wait_time_after_onece).grid(row=4, column=1)

        Button(root, text="浏览", command=self.browse_file).grid(row=0, column=2)
        self.start_button = Button(root, text="开始", command=self.start)
        self.start_button.grid(row=5, column=1)
        self.set_position_button = Button(root, text="确定鼠标位置", command=self.set_click_position)
        self.set_position_button.grid(row=6, column=1)

    def browse_file(self):
        self.excel_path.set(filedialog.askopenfilename())

    def set_click_position(self):
        self.set_position_button.config(text="请点击任意位置")
        self.root.config(cursor="cross")
        self.listener = mouse.Listener(on_click=self.on_click)
        self.listener.start()

    def on_click(self, x, y, button, pressed):
        if button== mouse.Button.left and pressed:
            self.click_position = (x, y)
            self.listener.stop()
            self.set_position_button.config(text="确定鼠标位置")
            self.root.config(cursor="")

    def start(self):
        self.running = True
        self.start_button.config(state="disabled")
        keyboard_listener = keyboard.Listener(on_press=self.on_key_press)
        keyboard_listener.start()

        df = pd.read_excel(self.excel_path.get(), header=None, usecols=[0])  # read the first row as data
        wait_time_after_enter = int(self.wait_time_after_enter.get())
        wait_time_after_paste = int(self.wait_time_after_paste.get())
        cmd_sum = int(self.cmd_sum.get())
        wait_time_after_onece = int(self.wait_time_after_onece.get())

        try:
            for i, row in df.iterrows():
                if not self.running:
                    break
                for cell in row:
                    pyperclip.copy(str(cell))  # copy the data to the clipboard
                    if self.click_position:
                        pyautogui.click(self.click_position)
                    else:
                        pyautogui.click()
                    pyautogui.hotkey('command', 'v')
                    time.sleep(wait_time_after_paste)
                    pyautogui.press('enter')
                    time.sleep(wait_time_after_enter)
                    pyautogui.press('enter')

                if (i + 1) % cmd_sum == 0:
                    time.sleep(wait_time_after_onece)  # wait for 10 minutes
        finally:
            self.running = False
            self.start_button.config(state="normal")

    def on_key_press(self, key):
        if key == keyboard.Key.esc:
            self.running = False
         
root = Tk()
app = App(root)
root.mainloop()