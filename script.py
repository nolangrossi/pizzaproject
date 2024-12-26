import openpyxl
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

path = filedialog.askopenfilename(
    title="Select an Excel File",
    filetypes=[("Excel Files", "*.xls *.xlsx")]
)

menu = {
    "Cheesy Bread" : '',
    "Mozz Cheese Bread" : '',
    "Pepperoni & Cheese" : '',
    "Ultimate 3 Cheese" : '',
    "Calzone" : '',
    "Paulzone" : '',
    "Catering Wings" : '',
    "Wings" : '',
    "Dessert" : '',
    "Big Double Chocolate Cake" : '',
    "Giant Cookie" : '',
    "Oreo Cheesecake" : '',
    "Red Velvet Cake" : '',
    "Strawberry Cheesecake" : '',
    "Jumbo Wings (8pc)" : '',
    "Drinks" : '',
    "Drinks Totals" : '',
    "Pizza" : '',
    "Pizza Totals" : '',
    "Slice" : '',
    "2 Slices" : '',
    "Pizza Slice" : '',
    "Subs" : '',
    "Ham" : '',
    "Italian" : '',
    "Italian Bake Sandwich" : '',
    "Turkey Club" : '',
    "Jumbo Wings (16)" : '', 
    "Jumbo Wings (6)" : '',
    "Jumbo Wings (8)" : '',
    "Bread" : '',
    "Bacon" : '',
    "Green Pepper" : '', 
    "Half Bacon" : '',
    "Hamburger" : '',
    "Burger" : '',
    "Onions" : '',
    "Cheese" : '',
    "Ground Beef" : '', 
    "Mushrooms" : '',
    "Sausage" : '',
    "1/2 Bacon" : '',
    "1/2 Green Peppers" : '',
    "1/2 Ham" : '', 
    "1/2 Hamburger" : '',
    "1/2 Mushrooms" : '',
    "1/2 Onions" : '',
    "1/2 Pepper Rings" : '', 
    "1/2 Pepperoni" : '', 
    "1/2 Sausage" : '', 
    "Chicken" : '', 
    "Pepper Rings" : '', 
    "Pepperoni" : '', 
    "Pineapple" : ''
}
menu_name_list = list(menu.keys())
menu_value_list = list(menu.values())

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active


num_rows = sheet_obj.max_row

for i in range(num_rows, 0, -1):
    cell_obj = sheet_obj.cell(row=i, column=2)
    if cell_obj.value not in menu_name_list: 
        sheet_obj.delete_rows(i)
        print(f"Row {i} deleted")
num_rows = sheet_obj.max_row
for i in range(num_rows, 0, -1):
    cell_obj = sheet_obj.cell(row=i, column=8)
    cell_obj.value = menu_value_list[i]
    print(menu_value_list[i+1])
print(menu_name_list)
print(menu_value_list)

wb_obj.save(path)
print("Workbook saved successfully.")
