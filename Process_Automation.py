#!/usr/bin/env python
# coding: utf-8

## Systems and Process Automation with PYTHON
# 
### Problem:
# 
# Every day, our systems update the previous day's sales. My daily work, as analista, it's send a email to board of directors, as soon as i start working. With the billing and quantity of products sold on the day before.
# 
# E-mail of board directors: marcoslucena_1998@hotmail.com 
# it's place where sales from the previous day are made available: https://onedrive.live.com/?id=A5C162D13D80ADB2%21282&cid=A5C162D13D80ADB2 
# 
# to solve, we are going to use pyautogui, it's a mouse and keyboard command automation library. 

# # install modulos 
# 
# ### !pip install PyAutoGUI
# ### !pip install Pyperclip

# In[1]:

# step by step for resolve this problem
import pyautogui
import pyperclip
import time
import pandas as pd

# step one: log into the campany's Sistem (Link Drive).

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://onedrive.live.com/?id=A5C162D13D80ADB2%21282&cid=A5C162D13D80ADB2")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

    # Load the System (wait 5 seconds)   
time.sleep(50)
    
# step two: browse the system (until the export folder).
pyautogui.click(x=973, y=274)
time.sleep(2)

# step three: Download the Sales Base.
pyautogui.click(x=508, y=134)
time.sleep(2)

# In[3]:

# step four: Import Sales base into Python.

table = pd.read_excel(r"C:\Users\Pedro\Downloads\vendas - Dez.xlsx")
display(table)

# step five: Calculate billing e and quantity of products sold.
quantity_of_products = table["Quantidade"].sum()
billing = table["Valor Final"].sum()

# In[59]:

# step seven: send email to board of directors.
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://outlook.live.com/mail/0/inbox/id/AQQkADAwATMwMAItNDZlMy0xZmMxLTAwAi0wMAoAEAAsQomdquRfR5SpU3J%2F0Nq4")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# delay 5  sec
time.sleep(5)

# click New Mensseger
pyautogui.click(x=177, y=139)
time.sleep(10)

# In[60]:

# write email
time.sleep(5)

pyperclip.copy("marcoslucena_1998@hotmail.com")
pyautogui.hotkey("ctrl", "v")
time.sleep(4)

pyautogui.press("tab")
time.sleep(2)

# write title
pyautogui.write("Sales Report")
time.sleep(2)
pyautogui.press("tab")

text = f"""
Good Morning,

The billing of yesterday was R$ {billing}
The quantity of products was {quantity_of_products}

Mr. Lucena
"""

pyautogui.write(text)
time.sleep(5)

pyautogui.click(x=675, y=654)


# In[61]:

# Show us the required position
time.sleep(5)
pyautogui.position()