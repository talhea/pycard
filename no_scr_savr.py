import pyautogui
import random, time

while True:
    scrn_width, scrn_height = pyautogui.size()
    
    x = random.randint(1, scrn_width)
    y = random.randint(1, scrn_height)
    
    pyautogui.moveTo(x, y, duration=2+)
    
    time.sleep(300)