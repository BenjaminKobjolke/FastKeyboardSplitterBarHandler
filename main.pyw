"""
FastKeyboardSplitterBarHandler - Keyboard-controlled splitter bar dragging

This tool allows you to drag vertical splitter bars using keyboard shortcuts.

Detection Modes:
- 'template': Uses reference images from the data folder (traditional method)
- 'auto': Uses computer vision to automatically detect vertical bars
- 'both': Combines both methods for maximum coverage (default)

The automatic detection uses:
1. Canny edge detection + Hough line transform for detecting vertical lines
2. Gradient analysis to find vertical edges (splitter bars)
3. Morphological operations to enhance vertical structures

Change DETECTION_MODE in MyApp class to switch modes.
"""

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
import tkinter as tk
import math

# set to true when working on this project
is_debug = True

main_hotkey = "alt+shift+W"
screenshot_hotkey = "ctrl+alt+F9"
data_folder = "data"

'''
if not is_debug:
    elevate(show_console=False)

# Capture the handle of the last active window before the Tkinter window is created

mutex = win32event.CreateMutex(None, 1, 'fast-keyboard-splitter-bar-handler')
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    mutex = None
    print("Another instance of this application is already running.")
    sys.exit(1)
'''

class MyApp:
    print("hi")
    last_active_window = None
    mouse_is_pressed = False  # Flag to keep track of mouse state
    splitterbar_coordinates = []
    overlay = None
    canvas = None
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

    MATCH_THRESHOLD = 0.8
    BORDER_LIMIT = 30

    # Automatic detection mode: 'template', 'auto', or 'both'
    DETECTION_MODE = 'both'

    def __init__(self, root):
        self.last_active_window = win32gui.GetForegroundWindow()

        self.root = root
        self.q = queue.Queue()
        # Dark theme styles
        root.configure(bg='#2E2E2E')
        root.title("FastKeyboardSplitterBarHandler")

        self.scrollbar = Scrollbar(root, bg='#555555')
        self.scrollbar.pack(side='right', fill='y')

        self.textarea = Text(root, wrap='word', yscrollcommand=self.scrollbar.set, bg='#555555', fg='#FFFFFF')
        self.textarea.pack()

        self.setup_hotkeys()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Minimize the window on startup
        '''
        if not is_debug:
            self.root.iconify()
            self.root.after(100, self.focus_on_last_window)
        '''

    def on_close(self):
        self.debug_print('Stopping hotkey handlers...')
        manager.unregister_all_hotkeys()
        self.root.destroy()
        exit(0)

    def focus_on_last_window(self):
        #win32gui.ShowWindow(self.last_active_window, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(self.last_active_window)

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
        self.debug_print("Taking screenshot" + "\n")

        # Get mouse position
        x, y = pyautogui.position()
        pyautogui.move(50, 50)
        time.sleep(2)  # wait for 5 seconds
        # Define area around the mouse position (200x100 px)
        left = x - 25
        top = y - 25
        right = x + 25
        bottom = y + 25

        # Take screenshot
        self.screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
        pyautogui.move(-50, -50)
        # Find the window title of the window the screenshot was taken of
        window_list = gw.getWindowsAt(x, y)
        window_title = window_list[0].title if window_list else "Unknown Window"
        self.last_window_title = window_title
        self.debug_print(f"Window title: {window_title}" + "\n")

        self.last_proccess_name = self.active_window_process_name()
        root.after(100, self.check_queue)
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

    def debug_print(self, text):
        self.textarea.insert(END, text + "\n")
        self.textarea.see(END)

    def end_mouse_drag(self):
        pyautogui.mouseUp()
        time.sleep(1)
        self.mouse_is_pressed = False
        pyautogui.moveTo(self.initial_mouse_position[0], self.initial_mouse_position[1])
        keyboard.unhook_all()
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
        
        print(e.name.lower())
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
            elif e.name.lower() == 'enter':
                self.end_mouse_drag()                

    def setup_hotkeys(self):
        manager.suppress = True
        id1 = manager.register_hotkey([Key.shift_l, Key.alt_l, 'w'], None, self.main_hotkey_pressed)
        if -1 == id1:
            self.debug_print('Already registered!')
        elif 0 == id1:
            self.debug_print('Invalid parameters!')
        else:
            self.debug_print('Hotkey id: {}'.format(id1))

        id2 = manager.register_hotkey([Key.shift_l, Key.alt_l, 'e'], None, self.take_screenshot)
        if -1 == id2:
            self.debug_print('Already registered!')
        elif 0 == id2:
            self.debug_print('Invalid parameters!')
        else:
            self.debug_print('Hotkey id: {}'.format(id2))

    def main_hotkey_pressed(self):
        self.debug_print("main_hotkey_pressed")
        self.last_active_window = win32gui.GetForegroundWindow()
        self.splitterbar_coordinates = []
        self.search_splitter_bars()

    def create_overlay(self):
        self.debug_print("create_overlay")
        # Create the overlay window
        self.overlay = tk.Tk()
        self.overlay.overrideredirect(True)
        self.overlay.attributes('-topmost', True)
        self.overlay.attributes('-transparentcolor', 'green')
        self.canvas = tk.Canvas(self.overlay, bg='green', highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.overlay.withdraw()  # Start with the overlay hidden

    def show_overlay(self):
        self.debug_print("show_overlay")
        if self.overlay is None:
            self.create_overlay()
        self.schedule_overlay_update()

    def schedule_overlay_update(self):
        # Schedule the overlay update in the main Tkinter loop
        self.root.after(10, self.update_overlay)  # Adjust the time as needed

    def update_overlay(self):
        try:
            # Get the active window dimensions and position
            active_window = gw.getActiveWindow()
            if not active_window:
                self.debug_print("No active window found.")
                return

            win_x, win_y, win_width, win_height = active_window.left, active_window.top, active_window.width, active_window.height
            if self.overlay:
                #p rint(f"Setting geometry: {win_width}x{win_height}+{win_x}+{win_y}")
                self.overlay.geometry(f"{win_width}x{win_height}+{win_x}+{win_y}")
                # Clear the canvas of old drawings
                self.canvas.delete("all")

        except Exception as e:
            print(f"Error in update_overlay: {e}")
            self.schedule_overlay_update()  # Reschedule the next update

        if self.overlay:
            circle_diameter = 50 * 1.1
            radius = circle_diameter / 2
            self.canvas.delete("all")
            self.hooked_keys.append(keyboard.hook_key('ESC', self.overlay_esc_pressed, suppress=True))

            for i, coordinate in enumerate(self.splitterbar_coordinates, start=1):
                circle_x, circle_y = coordinate[2], coordinate[3]

                # Get detection method and choose color
                detection_method = coordinate[4] if len(coordinate) > 4 else 'unknown'
                if detection_method == 'template':
                    fill_color = 'blue'
                elif detection_method == 'auto' or detection_method == 'gradient':
                    fill_color = 'red'
                else:
                    fill_color = 'yellow'  # unknown

                self.debug_print(f"Drawing circle {i} at {circle_x}, {circle_y} ({detection_method})")
                print(f"Drawing circle {i} at {circle_x}, {circle_y}")
                if i < 10:
                    counter_text = str(i)
                    hook_key = str(i)
                else:
                    counter_text = chr(87 + i)  # 87 + 10 = 97 ('a')
                    hook_key = counter_text

                self.canvas.create_oval(circle_x - radius, circle_y - radius, circle_x + radius, circle_y + radius,
                                        fill=fill_color)
                self.canvas.create_text(circle_x, circle_y, text=counter_text, font=("Arial", int(radius), "bold"),
                                        fill="white")
                self.hooked_keys.append(keyboard.hook_key(hook_key, self.overlay_keyboard_pressed, suppress=True))

        self.debug_print("show overlay")
        self.overlay.deiconify()  # Show the overlay

    def overlay_esc_pressed(self, e):
        self.debug_print("overlay_esc_pressed")
        keyboard.unhook_all()
        self.root.after(100, self.hide_overlay)

    def overlay_keyboard_pressed(self, e):
        #self.debug_print("overlay_digit_pressed")
        keyboard.unhook_all()

        debug_info = f"Key {e.name} - {e.scan_code} is pressed." + "\n"
        #print(debug_info)
        if e.name.isdigit():
            # For digits, just convert to int and subtract 1
            array_index = int(e.name) - 1
        elif e.name.isalpha() and len(e.name) == 1:
            # For letters, convert to lowercase, find its ASCII, and adjust the index
            array_index = ord(e.name.lower()) - ord('a') + 9  # 'a' corresponds to the 10th element
        else:
            print(f"Invalid key pressed: {e.name}")
            return
        coordinate = self.splitterbar_coordinates[array_index]
        #print(f"Moving mouse to {coordinate[0]}, {coordinate[1]}")

        pyautogui.moveTo(coordinate[0], coordinate[1])
        # Once the mouse is moved to the desired position and pressed down:
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
        
        self.hooked_keys.append(keyboard.hook_key('enter', self.mouse_move, suppress=True))        

        self.root.after(10, self.enable_mouse_drag)

    def enable_mouse_drag(self):
        self.hide_overlay()
        self.focus_on_last_window()
        pyautogui.mouseDown()

        pyautogui.mouseDown()
        self.mouse_is_pressed = True

    def hide_overlay(self):
        if self.overlay:
            try:
                self.canvas.delete("all")
                self.overlay.withdraw()
            except Exception as e:
                print(f"Error in update_overlay: {e}")
                self.root.after(10, self.hide_overlay)

    def close_overlay(self):
        self.overlay.destroy()

    def get_matching_subfolders(self, active_window):
        subfolders = [f.name for f in os.scandir("data") if f.is_dir()]
        matching_subfolders_from_title = [sf for sf in subfolders if sf.lower() in active_window.title.lower()]

        process_name = self.active_window_process_name()
        matching_subfolders_from_process = [sf for sf in subfolders if sf.lower() in process_name.lower()]

        matching_subfolders = list(set(matching_subfolders_from_title + matching_subfolders_from_process))
        if not matching_subfolders:
            self.debug_print(f"No matching subfolders found for '{active_window.title}' or '{process_name}'.")
            return None
        return matching_subfolders

    def search_splitter_bars(self):
        self.debug_print("search_splitter_bars")
        active_window = gw.getActiveWindow()
        gray_screenshot = self.get_screenshot(active_window)

        # Remember the initial mouse position
        self.initial_mouse_position = pyautogui.position()

        all_coordinates = []

        # Template matching mode
        if self.DETECTION_MODE in ('template', 'both'):
            matching_subfolders = self.get_matching_subfolders(active_window)

            if matching_subfolders is not None:
                self.debug_print("Running template matching detection")
                template_matches = self.find_splitter_bars(active_window, gray_screenshot, matching_subfolders)
                if template_matches:
                    all_coordinates.extend(template_matches)
            else:
                self.debug_print(f"No matching subfolders found for '{active_window.title}'.")

        # Automatic detection mode
        if self.DETECTION_MODE in ('auto', 'both'):
            self.debug_print("Running automatic detection")

            # Try edge/line detection method
            auto_matches = self.detect_vertical_bars_automatically(active_window, gray_screenshot)
            if auto_matches:
                all_coordinates.extend(auto_matches)

            # Try gradient-based detection method
            gradient_matches = self.detect_vertical_bars_by_gradient(active_window, gray_screenshot)
            if gradient_matches:
                all_coordinates.extend(gradient_matches)

        # Remove duplicates from combined results
        if all_coordinates:
            all_coordinates = self.filter_by_proximity(all_coordinates, proximity=20)
            self.splitterbar_coordinates = all_coordinates
            self.debug_print(f"Total found: {len(self.splitterbar_coordinates)} splitter bars.")
            self.show_overlay()
            return True

        self.debug_print("No vertical bars found" + "\n")
        return False

    def get_screenshot(self, active_window):
        screenshot = pyautogui.screenshot(
            region=(active_window.left, active_window.top, active_window.width, active_window.height))
        screenshot_np = np.array(screenshot)
        return cv2.cvtColor(screenshot_np, cv2.COLOR_BGR2GRAY)

    def find_splitter_bars(self, active_window, gray_screenshot, matching_subfolders):
        self.debug_print("find_splitter_bars")
        all_matches = []
        filenames = self.get_filenames_for_matching_subfolders(matching_subfolders)

        for filename in filenames:
            self.debug_print(f"Checking {filename}")
            template = cv2.imread(filename, 0)

            if gray_screenshot.shape[0] < template.shape[0] or gray_screenshot.shape[1] < template.shape[1]:
                self.debug_print(f"Screenshot for '{active_window.title}' is smaller than template. Skipping.")
                continue

            res = cv2.matchTemplate(gray_screenshot, template, cv2.TM_CCOEFF_NORMED)
            loc = np.where(res >= 0.8)
            file_matches = []

            max_iterations = 10
            iterations = 0
            for pt in zip(*loc[::-1]):
                if self.is_within_border_limit(pt, active_window, template.shape, 50):
                    target_x, target_y, target_relative_x, target_relative_y = self.calculate_coordinates(pt,
                                                                                                          active_window,
                                                                                                          template.shape)
                    file_matches.append((target_x, target_y, target_relative_x, target_relative_y, 'template'))

                iterations += 1
                if iterations >= max_iterations:
                    break

            # if file_matches is 0, then no matches were found
            if len(file_matches) == 0:
                continue

            # Apply proximity check to matches from this file
            filtered_matches = self.filter_by_proximity(file_matches, proximity=20)
            all_matches.extend(filtered_matches)

        self.debug_print(f"Total matches found: {len(all_matches)}")
        return all_matches

    def filter_by_proximity(self, coordinates_list, proximity):
        filtered = []
        for new_coords in coordinates_list:
            if not self.is_too_close_to_existing(filtered, new_coords, proximity):
                filtered.append(new_coords)
        return filtered

    def is_too_close_to_existing(self, existing_coords, new_coords, proximity):
        for coords in existing_coords:
            if self.calculate_distance(coords[0], coords[1], new_coords[0], new_coords[1]) < proximity:
                return True
        return False

    def calculate_distance(self, x1, y1, x2, y2):
        return math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)

    def is_within_border_limit(self, pt, window, template_size, border_limit):
        return (pt[0] > border_limit and
                pt[1] > border_limit and
                pt[0] + template_size[1] < window.width - border_limit and
                pt[1] + template_size[0] < window.height - border_limit)

    '''
        pt = coordinates of the match in screen coordinates
        window = active window
        template_size = size of the screenshot image
    '''
    def calculate_coordinates(self, pt, window, template_size):
        '''
            screen_x, screen_y = absolute cooridnates of the target on the screen
        '''
        print(f"pt: {pt}")
        screen_x = pt[0] + window.left
        #print(f"screen_x: {screen_x}")

        screen_y = pt[1] + window.top
        print(f"screen_y: {screen_y}")
        print(window)
        target_x = screen_x + (template_size[0] // 2)
        target_y = screen_y + (template_size[1] // 2)
        target_relative_x = pt[0] + (template_size[0] // 2)
        target_relative_y = pt[1] + (template_size[1] // 2)
        print("target_relative_y: " + str(target_relative_y))

        return target_x, target_y, target_relative_x, target_relative_y

    def is_close_to_existing(self, coordinates, new_x, new_y, proximity=100):
        for x, y, _, _ in coordinates:
            if abs(new_x - x) < proximity or abs(new_y - y) < proximity:
                return True
        return False

    def get_filenames_for_matching_subfolders(self, matching_subfolders):
        filenames = []
        for folder in matching_subfolders:
            folder_path = os.path.join("data", folder)
            # self.debug_print("checking folder " + folder_path);
            for filename in os.listdir(folder_path):
                if filename.endswith(".png"):
                    filenames.append(os.path.join(folder_path, filename))

        return filenames

    def detect_vertical_bars_automatically(self, active_window, gray_screenshot):
        """
        Automatically detect vertical bars (splitter bars) using computer vision techniques.
        Returns a list of coordinates in the same format as find_splitter_bars.
        """
        self.debug_print("Running automatic vertical bar detection")

        # Get screenshot dimensions
        height, width = gray_screenshot.shape
        cv2.imwrite('debug_screenshot.png', gray_screenshot)

        # Apply Gaussian blur to reduce noise
        blurred = cv2.GaussianBlur(gray_screenshot, (5, 5), 0)

        # Use Canny edge detection
        edges = cv2.Canny(blurred, 20, 100)
        cv2.imwrite('debug_edges.png', edges)

        # Create a kernel that emphasizes vertical structures
        vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 60))

        # Apply morphological operations to enhance vertical lines
        vertical_lines = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, vertical_kernel)

        # Dilate to make lines thicker and easier to detect
        vertical_lines = cv2.dilate(vertical_lines, vertical_kernel, iterations=2)
        cv2.imwrite('debug_vertical_lines.png', vertical_lines)

        # Use HoughLinesP to detect line segments
        lines = cv2.HoughLinesP(vertical_lines, 1, np.pi/180, threshold=15,
                                minLineLength=height*0.05, maxLineGap=200)

        detected_bars = []

        if lines is not None:
            self.debug_print(f"Found {len(lines)} potential vertical lines")

            # Debug counters
            angle_rejected = 0
            border_rejected = 0
            passed_angle = 0

            # Process detected lines
            for line in lines:
                x1, y1, x2, y2 = line[0]

                # Ensure y1 < y2 for valid array slicing
                if y1 > y2:
                    y1, y2 = y2, y1
                    x1, x2 = x2, x1

                # Check if line is mostly vertical (angle close to 90 degrees)
                angle = abs(np.arctan2(y2 - y1, x2 - x1) * 180 / np.pi)

                if 75 <= angle <= 105:  # Nearly vertical
                    passed_angle += 1
                    # Calculate center point of the line
                    center_x = (x1 + x2) // 2
                    center_y = (y1 + y2) // 2
                    self.debug_print(f"Vertical line found: x1={x1}, x2={x2}, y1={y1}, y2={y2}, center_x={center_x}, angle={angle:.1f}°")

                    # Skip if too close to borders
                    if (center_x < self.BORDER_LIMIT or
                        center_x > width - self.BORDER_LIMIT or
                        center_y < self.BORDER_LIMIT or
                        center_y > height - self.BORDER_LIMIT):
                        border_rejected += 1
                        continue

                    # Check if this line looks like a splitter bar (darker vertical strip)
                    # Sample a small region around the line
                    sample_width = 10
                    x_start = max(0, center_x - sample_width)
                    x_end = min(width, center_x + sample_width)

                    region = gray_screenshot[max(0, y1):min(height, y2), x_start:x_end]

                    if region.size > 0:
                        # Calculate statistics
                        mean_brightness = np.mean(region)
                        std_brightness = np.std(region)

                        # Debug: show values for each detected line
                        self.debug_print(f"Testing line at x={center_x}: brightness={mean_brightness:.1f}, std={std_brightness:.1f}")

                        # Splitter bars are often darker and have low variance
                        # Adjust these thresholds based on testing
                        if mean_brightness < 180 and std_brightness < 100:
                            # Convert to absolute screen coordinates
                            screen_x = center_x + active_window.left
                            screen_y = center_y + active_window.top

                            detected_bars.append((screen_x, screen_y, center_x, center_y, 'auto'))
                    else:
                        self.debug_print(f"Empty region for line at x={center_x}, y1={y1}, y2={y2}")
                else:
                    angle_rejected += 1

            # Debug output
            self.debug_print(f"Angle filter: {passed_angle} passed, {angle_rejected} rejected (not 75-105°)")
            self.debug_print(f"Border filter: {border_rejected} rejected (within {self.BORDER_LIMIT}px of edge)")
            self.debug_print(f"Before brightness filter: {len(detected_bars)} bars")

            # Filter by proximity to avoid duplicates
            filtered_bars = self.filter_by_proximity(detected_bars, proximity=20)
            self.debug_print(f"After filtering: {len(filtered_bars)} vertical bars detected")

            return filtered_bars
        else:
            self.debug_print("No vertical lines detected")
            return []

    def detect_vertical_bars_by_gradient(self, active_window, gray_screenshot):
        """
        Alternative method using gradient analysis to find vertical splitter bars.
        Splitter bars typically have strong horizontal gradients (dark-light-dark pattern).
        """
        self.debug_print("Running gradient-based vertical bar detection")

        height, width = gray_screenshot.shape

        # Calculate horizontal gradient (Sobel)
        gradient_x = cv2.Sobel(gray_screenshot, cv2.CV_64F, 1, 0, ksize=3)
        gradient_x = np.abs(gradient_x)

        # Normalize to 0-255
        gradient_x = np.uint8(255 * gradient_x / np.max(gradient_x))

        # Threshold to find strong vertical edges
        _, thresh = cv2.threshold(gradient_x, 100, 255, cv2.THRESH_BINARY)

        # Find contours
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        detected_bars = []

        for contour in contours:
            # Get bounding rectangle
            x, y, w, h = cv2.boundingRect(contour)

            # Check if contour is tall and thin (like a vertical bar)
            aspect_ratio = h / (w + 1)

            if aspect_ratio > 3 and h > height * 0.2 and w < 30:
                # Calculate center
                center_x = x + w // 2
                center_y = y + h // 2

                # Skip if too close to borders
                if (center_x < self.BORDER_LIMIT or
                    center_x > width - self.BORDER_LIMIT or
                    center_y < self.BORDER_LIMIT or
                    center_y > height - self.BORDER_LIMIT):
                    continue

                # Convert to absolute screen coordinates
                screen_x = center_x + active_window.left
                screen_y = center_y + active_window.top

                detected_bars.append((screen_x, screen_y, center_x, center_y, 'gradient'))

        # Filter by proximity
        filtered_bars = self.filter_by_proximity(detected_bars, proximity=20)
        self.debug_print(f"Gradient detection found {len(filtered_bars)} vertical bars")

        return filtered_bars

root = Tk()
app = MyApp(root)
app.show_overlay()
root.mainloop()

# Release the mutex when application exits
if mutex:
    win32api.CloseHandle(mutex)
