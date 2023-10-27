import cv2
import numpy as np
import pygetwindow as gw
import pyautogui
import keyboard
import ctypes
import sys
import os
import time

# Configure mouse movement speed
MOUSE_SPEED = 15

last_key_time = 0  # To remember the last time a key was pressed
# Initialize an empty list to hold the hooked keys.
hooked_keys = []

# ... (other imports and variables)
initial_mouse_position = None

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


mouse_is_pressed = False  # Flag to keep track of mouse state


# Function to capture the screenshot and find the splitter bars
def search_splitter_bars():
    global hooked_keys, mouse_is_pressed, initial_mouse_position

    def move_mouse(e):
        global last_key_time
        current_time = time.time() * 1000  # Get the current time in milliseconds

        if current_time - last_key_time < 200:  # 200 milliseconds threshold
            return  # Skip this key press

        last_key_time = current_time  # Update last_key_time to the current time

        print(f"Key {e.name} is pressed.")
        speed = MOUSE_SPEED

        if e.name.upper() == e.name:  # Check if the key is uppercase
            speed *= 2  # Double the speed

        if e.event_type == 'down':  # Only act on the 'down' event
            if e.name.lower() == 'h':
                pyautogui.moveRel(-speed, 0)
            elif e.name.lower() == 'j':
                pyautogui.moveRel(0, speed)
            elif e.name.lower() == 'k':
                pyautogui.moveRel(0, -speed)
            elif e.name.lower() == 'l':
                pyautogui.moveRel(speed, 0)

    if mouse_is_pressed:  # Release the mouse and stop moving it with hjkl
        pyautogui.mouseUp()
        mouse_is_pressed = False
        pyautogui.moveTo(initial_mouse_position[0], initial_mouse_position[1])
        keyboard.unhook_all()
        keyboard.add_hotkey('ctrl+F10', search_splitter_bars)
        '''
        for key in hooked_keys:
            try:
                keyboard.unhook_key(key)
            except Exception as e:
                print(f"Error while unhooking key {key}: {e}")
        '''

        hooked_keys = []
        return

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
    initial_mouse_position = pyautogui.position()

    for folder in matching_subfolders:
        folder_path = os.path.join("data", folder)
        for filename in os.listdir(folder_path):
            if filename.endswith(".png"):
                template = cv2.imread(os.path.join(folder_path, filename), 0)

                if gray_screenshot.shape[0] < template.shape[0] or gray_screenshot.shape[1] < template.shape[1]:
                    print(f"Screenshot for '{active_window.title}' is smaller than template. Skipping.")
                    continue

                res = cv2.matchTemplate(gray_screenshot, template, cv2.TM_CCOEFF_NORMED)
                threshold = 0.8
                loc = np.where(res >= threshold)

                for pt in zip(*loc[::-1]):
                    screen_x = pt[0] + active_window.left
                    screen_y = pt[1] + active_window.top
                    target_x = screen_x + 5
                    target_y = screen_y
                    pyautogui.moveTo(target_x, target_y)
                    # Once the mouse is moved to the desired position and pressed down:
                    pyautogui.mouseDown()
                    mouse_is_pressed = True
                    hooked_keys.append(keyboard.hook_key('h', move_mouse, suppress=True))
                    hooked_keys.append(keyboard.hook_key('j', move_mouse, suppress=True))
                    hooked_keys.append(keyboard.hook_key('k', move_mouse, suppress=True))
                    hooked_keys.append(keyboard.hook_key('l', move_mouse, suppress=True))
                    hooked_keys.append(keyboard.hook_key('H', move_mouse, suppress=True))
                    hooked_keys.append(keyboard.hook_key('J', move_mouse, suppress=True))
                    hooked_keys.append(keyboard.hook_key('K', move_mouse, suppress=True))
                    hooked_keys.append(keyboard.hook_key('L', move_mouse, suppress=True))
                    # keyboard.add_hotkey('ctrl+F10', search_splitter_bars)
                    break


if is_admin():
    print("The script is running with administrative privileges.")
else:
    print("The script is not running with administrative privileges. Requesting admin rights...")
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

print("waiting")
# Wait for 'ctrl+F10' to be pressed
keyboard.add_hotkey('ctrl+F10', search_splitter_bars)

keyboard.wait()
