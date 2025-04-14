import pyautogui
import time
from pynput import mouse

# Function to handle mouse click events
def on_click(x, y, button, pressed):
    if button == mouse.Button.left and pressed:
        print(f'\nLeft click at: ({x}, {y})')

# Set up the mouse listener
listener = mouse.Listener(on_click=on_click)
listener.start()

# Real-time display of mouse position
while True:
    time.sleep(1)
    x, y = pyautogui.position()
    print(f'鼠标位置: ({x}, {y})', end='\r')