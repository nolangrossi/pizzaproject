import openpyxl
import tkinter as tk
from tkinter import filedialog

# Initialize Tkinter and hide the root window
root = tk.Tk()
root.withdraw()

# Prompt the user to select an Excel file
path = filedialog.askopenfilename(
    title="Select an Excel File",
    filetypes=[("Excel Files", "*.xls *.xlsx")]  # Filter for Excel files
)

# Define the menu list
menu = [
    "Cheesy Bread", "Mozz Cheese Bread", "Pepperoni & Cheese", "Ultimate 3 Cheese",
    "Calzone", "Paulzone", "Catering Wings", "Wings", "Dessert",
    "Big Double Chocolate Cake", "Giant Cookie", "Oreo Cheesecake", "Red Velvet Cake",
    "Strawberry Cheesecake", "Jumbo Wings (8pc)", "Drinks", "Drinks Totals",
    "Pizza", "Pizza Totals", "Slice", "2 Slices", "Pizza Slice", "Subs", "Ham",
    "Italian", "Italian Bake Sandwich", "Turkey Club", "Jumbo Wings (16)", 
    "Jumbo Wings (6)", "Jumbo Wings (8)", "Bread", "Bacon", "Green Pepper", 
    "Half Bacon", "Hamburger", "Burger", "Onions", "Cheese", "Ground Beef", 
    "Mushrooms", "Sausage", "1/2 Bacon", "1/2 Green Peppers", "1/2 Ham", 
    "1/2 Hamburger", "1/2 Mushrooms", "1/2 Onions", "1/2 Pepper Rings", 
    "1/2 Pepperoni", "1/2 Sausage", "Chicken", "Pepper Rings", "Pepperoni", 
    "Pineapple"
]

# Load the workbook and get the active sheet
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

# Get the total number of rows in the sheet
num_rows = sheet_obj.max_row

# Iterate through the rows in reverse order
for i in range(num_rows, 0, -1):
    cell_obj = sheet_obj.cell(row=i, column=2)  # Check column 2
    if cell_obj.value not in menu:  # Check if the value is not in the menu
        sheet_obj.delete_rows(i)  # Delete the row
        print(f"Row {i} deleted")

# Save the workbook after processing
wb_obj.save(path)
print("Workbook saved successfully.")
