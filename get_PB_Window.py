import pygetwindow as gw
import psutil
import sys

# Function to check if Power Broker (brow.exe) is running
def is_power_broker_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'brow.exe':
            return True
    return False

def locate_PB_Window():
    # Use a partial title to match the window
    partial_title = "Power Broker"
    windows = gw.getWindowsWithTitle(partial_title)

    if windows:
        for window in windows:
            # Ensure the window is active and not minimized
            if window.isActive or not window.isMinimized:
                return window  # Return the first matching window
    print(f"No active window with title containing '{partial_title}' found.")
    sys.exit()

# Alert user if Power Broker is not running
def check_power_broker():
    if not is_power_broker_running():
        print("Power Broker is not running. Please open the Power Broker application and try again.")
        sys.exit()  # Exit the script if Power Broker is not running
    else:
        print("Power Broker is running. Proceeding with the process...")

# Example usage
if __name__ == "__main__":
    check_power_broker()  # Ensure Power Broker is running
    window = locate_PB_Window()  # Retrieve the window object
    print(f"Power Broker window position: ({window.left}, {window.top})")
