from tkinter import *
from interface import Interface
from hardcore import StoreList

# Screen setup
root = Tk()
root.title('e-Mail Models')
root.iconbitmap('icon/icon.ico')

# Global variable
unit_picker = ''

shop_list = StoreList()
#print(shop_list.store_head())

screen = Interface(root, shop_list.store_head())
root.mainloop()




