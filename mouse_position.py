import pyautogui

def get_mouse_position():
    # Get the current mouse position
    x, y = pyautogui.position()
    
    # Print the coordinates
    print(f"Current mouse position: X = {x}, Y = {y}")





# Example usage
get_mouse_position()
