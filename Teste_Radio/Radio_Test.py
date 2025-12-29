### Libraries ###

import      pyautogui
import      pyscreeze
import      pandas
import      time

### Variables ###




### Test Cases ###

class Elster():
    def click():
        
        img_barra = pyscreeze.locateOnScreen('New_position.png')
        x,y = pyautogui.center(img_barra)
        pyautogui.click(x,y, clicks=2, interval=0.25)
        time.sleep(2)
        img_samplat = pyscreeze.locateOnScreen('clickpc.png')
        x,y = pyautogui.center(img_samplat)
        pyautogui.click(x,y, clicks=2, interval=0.25)
        pyautogui.moveTo(x,y)
        time.sleep(2)
        
    click()
Elster()
