import cv2
import numpy as np
import pygetwindow as gw
import pyautogui
import keyboard
import ctypes
import os
import time
from tkinter import Tk, Label, Button, Entry, Text, Scrollbar, END
import os
from elevate import elevate

elevate(show_console=False)
class MyApp:
    pyautogui.PAUSE = 0.1
    # Configure mouse movement speed
    MOUSE_SPEED = 15

    last_key_time = 0  # To remember the last time a key was pressed
    # Initialize an empty list to hold the hooked keys.
    hooked_keys = []

    # ... (other imports and variables)
    initial_mouse_position = None

    def __init__(self, root):

        self.root = root
        # Dark theme styles
        root.configure(bg='#2E2E2E')
        root.title("keyboard-splitter-bar-handler")

        self.scrollbar = Scrollbar(root, bg='#555555')
        self.scrollbar.pack(side='right', fill='y')

        self.textarea = Text(root, wrap='word', yscrollcommand=self.scrollbar.set, bg='#555555', fg='#FFFFFF')
        self.textarea.pack()

        keyboard.add_hotkey('ctrl+F10', self.search_splitter_bars)

        # keyboard.wait()


    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    mouse_is_pressed = False  # Flag to keep track of mouse state

    def debug_print(self, text):
        self.textarea.insert(END, text)
        self.textarea.see(END)

    def end_mouse_drag(self):
        pyautogui.mouseUp()
        self.mouse_is_pressed = False
        pyautogui.moveTo(self.initial_mouse_position[0], self.initial_mouse_position[1])
        keyboard.unhook_all()
        keyboard.add_hotkey('ctrl+F10', self.search_splitter_bars)
        self.hooked_keys = []

    def mouse_move(self, e):
        current_time = time.time() * 1000  # Get the current time in milliseconds

        if current_time - self.last_key_time < 10:  # 200 milliseconds threshold
            return  # Skip this key press

        last_key_time = current_time  # Update last_key_time to the current time

        shift_pressed = keyboard.is_pressed('shift')

        debug_info = f"Key {e.name} - {e.scan_code} is pressed." + "\n"
        self.debug_print(debug_info)
        speed = self.MOUSE_SPEED

        if e.name.upper() == e.name or shift_pressed:
            speed *= 4  # Double the speed

        if e.event_type == 'down':  # Only act on the 'down' event
            if e.name.lower() == 'h' or e.name.lower() == 'a' or e.scan_code == 75:
                pyautogui.moveRel(-speed, 0, 0, pyautogui.linear, False, True)
            elif e.name.lower() == 'j' or e.name.lower() == 's' or e.scan_code == 80:
                pyautogui.move(0, speed, 0, pyautogui.linear, False, True)
            elif e.name.lower() == 'k' or e.name.lower() == 'w' or e.scan_code == 72:
                pyautogui.moveRel(0, -speed, 0, pyautogui.linear, False, True)
            elif e.name.lower() == 'l' or e.name.lower() == 'd' or e.scan_code == 77:
                pyautogui.moveRel(speed, 0, 0, pyautogui.linear, False, True)
            elif e.name.lower() == 'esc':
                self.end_mouse_drag()

    # Function to capture the screenshot and find the splitter bars
    def search_splitter_bars(self):
        active_window = gw.getActiveWindow()
        print(active_window.title)

        # Take a screenshot of the active window
        screenshot = pyautogui.screenshot(
            region=(active_window.left, active_window.top, active_window.width, active_window.height))
        screenshot_np = np.array(screenshot)
        gray_screenshot = cv2.cvtColor(screenshot_np, cv2.COLOR_BGR2GRAY)

        subfolders = [f.name for f in os.scandir("data") if f.is_dir()]
        matching_subfolders = [sf for sf in subfolders if sf.lower() in active_window.title.lower()]

        # Remember the initial mouse position
        self.initial_mouse_position = pyautogui.position()

        for folder in matching_subfolders:
            folder_path = os.path.join("data", folder)
            #print("checking folder " + folder_path);
            for filename in os.listdir(folder_path):
                if filename.endswith(".png"):
                    template = cv2.imread(os.path.join(folder_path, filename), 0)

                    if gray_screenshot.shape[0] < template.shape[0] or gray_screenshot.shape[1] < template.shape[1]:
                        print(f"Screenshot for '{active_window.title}' is smaller than template. Skipping.")
                        continue

                    self.debug_print("Match found: " + filename + "\n")
                    res = cv2.matchTemplate(gray_screenshot, template, cv2.TM_CCOEFF_NORMED)
                    threshold = 0.8
                    loc = np.where(res >= threshold)

                    for pt in zip(*loc[::-1]):
                        screen_x = pt[0] + active_window.left
                        screen_y = pt[1] + active_window.top
                        target_x = screen_x + (template.shape[1] // 2)
                        target_y = screen_y + (template.shape[0] // 2)

                        pyautogui.moveTo(target_x, target_y)
                        # Once the mouse is moved to the desired position and pressed down:
                        pyautogui.mouseDown()
                        mouse_is_pressed = True
                        self.hooked_keys.append(keyboard.hook_key('h', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('j', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('k', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('l', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('H', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('J', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('K', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('L', self.mouse_move, suppress=True))

                        self.hooked_keys.append(keyboard.hook_key('w', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('a', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('s', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('d', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('W', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('A', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('S', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('D', self.mouse_move, suppress=True))

                        self.hooked_keys.append(keyboard.hook_key('left', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('right', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('up', self.mouse_move, suppress=True))
                        self.hooked_keys.append(keyboard.hook_key('down', self.mouse_move, suppress=True))

                        self.hooked_keys.append(keyboard.hook_key('ESC', self.mouse_move, suppress=True))
                        # keyboard.add_hotkey('ctrl+F10', search_splitter_bars)
                        return

        print("No match found")
        self.debug_print("No match found" + "\n")


root = Tk()
app = MyApp(root)
root.mainloop()
