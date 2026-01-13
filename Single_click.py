# -*- coding: utf-8 -*-
"""
Created on Mon Oct 28 13:10:23 2024

single_click

@author: basanel
"""



import pyautogui
import time

def single_click(x=None, y=None, delay=1):
    """
    Perform a single click at the specified coordinates.
    If no coordinates are provided, click at the current mouse position.
    Adds a delay after the click.
    
    :param x: X coordinate (optional)
    :param y: Y coordinate (optional)
    :param delay: Delay in seconds after the click (default is 1 second)
    """
    if x is not None and y is not None:
        # Move to the specified coordinates and click
        pyautogui.moveTo(x, y)
    # Perform the single click
    pyautogui.click()
    # Add a delay
    time.sleep(delay)

# Example usage: Click at the current mouse position with a 2-second delay
single_click(delay=2)

# Example usage: Click at specific coordinates (200, 200) with a 3-second delay
single_click(200, 200, delay=3)
