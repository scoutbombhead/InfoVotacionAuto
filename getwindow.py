import pyautogui
import pygetwindow as gw
import time
import configparser
import os


def save_coordinates_to_ini(x, y, width, height, filename="config.ini"):
    """Save coordinates to an INI configuration file"""
    config = configparser.ConfigParser()
    
    # Create section if it doesn't exist
    config['SCREEN_REGION'] = {
        'x': str(x),
        'y': str(y),
        'width': str(width),
        'height': str(height)
    }
    
    # Write to file
    with open(filename, 'w') as configfile:
        config.write(configfile)
    
    print(f"\n✓ Coordinates saved to '{filename}'")


def discover_coordinates():
    """Interactive tool to find window and region coordinates"""

    print("=== Window Coordinate Discovery Tool ===")
    print("1. First, we'll find your application window")

    # Get window by title
    title = " "
    windows = gw.getWindowsWithTitle(title)

    if not windows:
        print("Window not found!")
        return

    window = windows[0]
    print(f"\nFound window: '{window.title}'")
    print(f"Window coordinates: left={window.left}, top={window.top}")
    print(f"Window size: {window.width} x {window.height}")

    # Activate the window
    window.activate()
    time.sleep(1)

    print("\n2. Now, move your mouse to the TOP-LEFT corner of the result area")
    print("   and press Enter...")
    input()
    x1, y1 = pyautogui.position()

    print(f"   Top-left: ({x1}, {y1})")

    print("\n3. Move your mouse to the BOTTOM-RIGHT corner of the result area")
    print("   and press Enter...")
    input()
    x2, y2 = pyautogui.position()

    # Calculate region
    width = x2 - x1
    height = y2 - y1

    print(f"\n=== RESULTS ===")
    print(f"Full window region: ({window.left}, {window.top}, {window.width}, {window.height})")
    print(f"Result area region: ({x1}, {y1}, {width}, {height})")
    print(f"\nUse this in your code:")
    print(f"result_region = ({x1}, {y1}, {width}, {height})")
    
    # Save to INI file
    save_coordinates_to_ini(x1, y1, width, height)
    print("\n✓ Configuration saved! The main application will now use these coordinates.")


if __name__ == "__main__":
    if __name__ == "__main__":
        discover_coordinates()