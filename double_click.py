# -*- coding: utf-8 -*-
"""
Double Click
"""


import pyautogui
import time

def double_click(x=None, y=None, delay=1):
    """
    Perform a double-click at the specified coordinates.
    If no coordinates are provided, double-click at the current mouse position.
    Adds a delay after the double-click.
    
    :param x: X coordinate (optional)
    :param y: Y coordinate (optional)
    :param delay: Delay in seconds after the double-click (default is 1 second)
    """
    if x is not None and y is not None:
        # Move to the specified coordinates and double-click
        pyautogui.moveTo(x, y)
    # Perform the double-click
    pyautogui.doubleClick()
    # Add a delay
    time.sleep(delay)



# Example usage: Double-click at specific coordinates (200, 200) with a 3-second delay
double_click(1635, 31, delay=2)
