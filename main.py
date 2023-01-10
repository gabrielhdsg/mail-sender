import pyautogui as pa
import time
import pyperclip as pc
import pandas as pd

#time between each command
pa.PAUSE = 1

pa.alert("The program will run. After click on OK please do not interfere.")

#open browser
pa.press("winright")
pa.write("chrome")
pa.press("enter")

#search file in Google Drive
link = "-"
pc.copy(link)
pa.hotkey('ctrl', 'v')
pa.press("enter")
time.sleep(10)

#download excel file
pa.hotkey('alt', 'f')
pa.press("d")
pa.press("enter")
time.sleep(5)

#getting file info
target = pd.read_excel(r'C:/Users/mrbie/Downloads/Sales-Dec.xlsx')
amount_products = target['Quantidade'].sum()
turnover = target['Valor Final'].sum()

#opening e-mail
pa.hotkey('ctrl', 't')
pa.write("mail.google.com")
pa.press("enter")
time.sleep(5)
pa.press("c")

#writing e-mail subject and addressee
pa.write("-")
pa.press("tab")
pa.press("tab")
subject = "Sales Report December"
pc.copy(subject)
pa.hotkey('ctrl', 'v')

#writing e'mail message
pa.press("tab")
text = f""" 
Good morning,

December turnover was: R${turnover:,.2f}
The amount of products sold was: {amount_products:,}
"""
pc.copy(text)
pa.hotkey('ctrl', 'v')
pa.hotkey('ctrl', 'enter')