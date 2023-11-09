from typing import Optional
import cv2
import numpy as np
import psutil
import pygetwindow as gw
import pyautogui
import keyboard
import ctypes
import time
import win32gui
import win32process
from PIL import ImageGrab
from tkinter import Tk, Text, Scrollbar, END, simpledialog, Entry, Button, Label
import os
from elevate import elevate
import queue
from PyHotKey import Key, keyboard_manager as manager
import win32con
import win32event
import win32api
import sys
import winerror

main_hotkey = "alt+shift+W"
screenshot_hotkey = "ctrl+alt+F9"
data_folder = "data"

elevate(show_console=False)

# Capture the handle of the last active window before the Tkinter window is created
last_active_window = win32gui.GetForegroundWindow()

mutex = win32event.CreateMutex(None, 1, 'fast-keyboard-splitter-bar-handler')
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    mutex = None
    print("Another instance of this application is already running.")
    sys.exit(1)


class MyApp:
    pyautogui.PAUSE = 0.1
    # Configure mouse movement speed
    MOUSE_SPEED = 15

    last_key_time = 0  # To remember the last time a key was pressed
    # Initialize an empty list to hold the hooked keys.
    hooked_keys = []

    # ... (other imports and variables)
    initial_mouse_position = None
    screenshot = None
    last_window_title = None
    last_proccess_name = None

    def __init__(self, root):

        self.root = root
        self.q = queue.Queue()
        # Dark theme styles
        root.configure(bg='#2E2E2E')
        root.title("keyboard-splitter-bar-handler")

        self.scrollbar = Scrollbar(root, bg='#555555')
        self.scrollbar.pack(side='right', fill='y')

        self.textarea = Text(root, wrap='word', yscrollcommand=self.scrollbar.set, bg='#555555', fg='#FFFFFF')
        self.textarea.pack()
        '''
        self.text_box = Entry(root)
        self.text_box.pack(pady=20)

        self.button = Button(root, text="Submit", command=self.on_submit)
        self.button.pack(pady=10)

        self.label = Label(root, text="Waiting for message...")
        self.label.pack(pady=20)
        '''

        self.setup_hotkeys()

        # Minimize the window on startup
        self.root.iconify()

        self.root.after(100, self.focus_on_last_window)

    def focus_on_last_window(self):
        global last_active_window
        win32gui.ShowWindow(last_active_window, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(last_active_window)

    '''
    def on_submit(self):
        received_data = self.text_box.get()
        self.label.config(text=f"Received: {received_data}")
    '''

    def get_next_filename(self, folder_path, base_filename):
        counter = 1
        while True:
            filename = f"{base_filename}_{counter}.png"
            if not os.path.exists(os.path.join(folder_path, filename)):
                return filename
            counter += 1

    def check_queue(self):
        try:
            task = self.q.get_nowait()
            if task == "askstring":
                self.debug_print("yes")
                subfolder_name: Optional[str] = simpledialog.askstring("Input",
                                                                       "Enter the subfolder name to save screenshot in:",
                                                                       initialvalue=self.last_window_title + " - " + self.last_proccess_name)
                self.save_screenshot(subfolder_name)
                return
        except queue.Empty:
            pass

        root.after(100, self.check_queue)

    def active_window_process_name(self):
        try:
            pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
            return psutil.Process(pid[-1]).name()
        except:
            pass

    def take_screenshot(self):
        root.after(100, self.check_queue)
        self.debug_print("Taking screenshot" + "\n")

        # Get mouse position
        x, y = pyautogui.position()
        pyautogui.move(50, 50)
        time.sleep(1)  # wait for 5 seconds
        # Define area around the mouse position (200x100 px)
        left = x - 50
        top = y - 50
        right = x + 50
        bottom = y + 50

        # Take screenshot
        self.screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
        pyautogui.move(-50, -50)
        # Find the window title of the window the screenshot was taken of
        window_list = gw.getWindowsAt(x, y)
        window_title = window_list[0].title if window_list else "Unknown Window"
        self.last_window_title = window_title
        self.debug_print(f"Window title: {window_title}" + "\n")

        self.last_proccess_name = self.active_window_process_name()
        self.q.put("askstring")

    def save_screenshot(self, subfolder_name):
        if not subfolder_name:
            return
        self.debug_print(f"Subfolder name: {subfolder_name}" + "\n")
        # Create dataand subfolder if not exists
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)

        full_subfolder_path = os.path.join(data_folder, subfolder_name)
        self.debug_print(f"Saving screenshot to {full_subfolder_path}" + "\n")
        if not os.path.exists(full_subfolder_path):
            os.makedirs(full_subfolder_path)

        # Check if screenshot.png already exists
        if os.path.exists(os.path.join(full_subfolder_path, "screenshot.png")):
            filename = self.get_next_filename(full_subfolder_path, "screenshot")
        else:
            filename = "screenshot.png"

        # Save the image
        self.screenshot.save(os.path.join(full_subfolder_path, filename))

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
        self.setup_hotkeys()
        self.hooked_keys = []

    def mouse_move(self, e):
        # pyautogui.mouseDown()
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

    def setup_hotkeys(self):
        manager.suppress = True
        id1 = manager.register_hotkey([Key.shift_l, Key.alt_l, 'w'], None, self.search_splitter_bars)
        if -1 == id1:
            print('Already registered!')
        elif 0 == id1:
            print('Invalid parameters!')
        else:
            print('Hotkey id: {}'.format(id1))

        id2 = manager.register_hotkey([Key.shift_l, Key.alt_l, 'e'], None, self.take_screenshot)
        if -1 == id2:
            print('Already registered!')
        elif 0 == id2:
            print('Invalid parameters!')
        else:
            print('Hotkey id: {}'.format(id2))
        # keyboard.add_hotkey(main_hotkey, self.search_splitter_bars, suppress=True)
        # keyboard.add_hotkey(screenshot_hotkey, self.take_screenshot, suppress=True)

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
        matching_subfolders_from_title = [sf for sf in subfolders if sf.lower() in active_window.title.lower()]

        process_name = self.active_window_process_name()
        matching_subfolders_from_process = [sf for sf in subfolders if sf.lower() in process_name.lower()]

        matching_subfolders = list(set(matching_subfolders_from_title + matching_subfolders_from_process))

        if not matching_subfolders:
            self.debug_print(f"No matching subfolders found for '{active_window.title}' or '{process_name}'.")
            return True
        # Remember the initial mouse position
        self.initial_mouse_position = pyautogui.position()

        for folder in matching_subfolders:
            folder_path = os.path.join("data", folder)
            # print("checking folder " + folder_path);
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

                    border_limit = 50

                    for pt in zip(*loc[::-1]):
                        # Check if the match is within the border limit
                        if (pt[0] > border_limit and
                            pt[1] > border_limit and
                            pt[0] + template.shape[1] < active_window.width - border_limit and
                            pt[1] + template.shape[0] < active_window.height - border_limit):
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
                            return True

        self.debug_print("No match found" + "\n")
        return True


root = Tk()
app = MyApp(root)
root.mainloop()

# Release the mutex when application exits
if mutex:
    win32api.CloseHandle(mutex)
