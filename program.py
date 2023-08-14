import tkinter as tk
from tkinter import ttk
from tkinter import font
import pandas as pd

form_data = []

def submit():
    product_name = product_name_entry.get()
    sku = sku_entry.get()
    product_brand = product_brand_entry.get()
    product_source = product_source_entry.get()
    product_categories = product_categories_entry.get()
    product_type = product_type_entry.get()
    description = description_entry.get(1.0, 'end-1c')
    colors = [var.get() for var in color_vars if var.get()]  # Get checked colors
    default_color = default_color_var.get()
    price = price_entry.get()
    downloadable = downloadable_var.get()
    sizes = [var.get() for var in size_vars if var.get()]

    row_data = {
        'Product Name': product_name,
        'SKU': sku,
        'Product Brand': product_brand,
        'Product Source': product_source,
        'Product Categories': product_categories,
        'Product Type': product_type,
        'Description': description,
        'Colors': colors,
        'Default Color': default_color,
        'Price': price,
        'Downloadable': downloadable,
        'Sizes': sizes
    }

    form_data.append(row_data)

    # clear/reset all fields
    product_name_entry.delete(0, 'end')
    sku_entry.delete(0, 'end')
    product_brand_entry.delete(0, 'end')
    product_source_entry.delete(0, 'end')
    product_categories_entry.delete(0, 'end')
    product_type_entry.delete(0, 'end')
    description_entry.delete(1.0, 'end')
    for var in color_vars:
        var.set("")
    default_color_var.set("")
    price_entry.delete(0, 'end')
    downloadable_var.set("")
    for var in size_vars:
        var.set("")




def export_to_excel():
    # Open the existing data
    try:
        df_existing = pd.read_excel("form_data.xlsx")
    except FileNotFoundError:
        df_existing = pd.DataFrame()

    # Create a DataFrame from the new data
    df_new = pd.DataFrame(form_data)

    # Append the new data to the existing data
    df = pd.concat([df_existing, df_new], ignore_index=True)

    # Write the data back to the Excel file
    df.to_excel("form_data.xlsx", index=False)
    form_data.clear()  # Clear form_data after exporting
    print("Data exported to 'form_data.xlsx'.")



def select_colors():
    top = tk.Toplevel(root)
    top.title("Select Colors")

    for i, color in enumerate(colors):
        if len(color_vars) < len(colors):
            var = tk.StringVar(value="")
            color_vars.append(var)
        else:
            var = color_vars[i]

        cb = ttk.Checkbutton(top, text=color, variable=var, onvalue=color, offvalue="")
        cb.grid(row=i//4, column=i%4)

def select_sizes():
    top = tk.Toplevel(root)
    top.title("Select Sizes")

    for i, size in enumerate(sizes):
        if len(size_vars) < len(sizes):
            var = tk.StringVar(value="")
            size_vars.append(var)
        else:
            var = size_vars[i]

        cb = ttk.Checkbutton(top, text=size, variable=var, onvalue=size, offvalue="")
        cb.grid(row=i//2, column=i%2)


root = tk.Tk()
root.geometry('800x800')
root.title('Product Input')
root.columnconfigure((0, 1), weight=1)

default_font = font.nametofont("TkDefaultFont")
default_font.configure(size=12)

# Create a frame for product info
product_frame = ttk.LabelFrame(root, text="Product Info")
product_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
product_frame.columnconfigure((0, 1), weight=1)

# Product Name Input
product_name_label = ttk.Label(product_frame, text="Product Name:")
product_name_label.grid(row=0, column=0, padx=5, pady=5, sticky='w')
product_name_entry = ttk.Entry(product_frame)
product_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

# SKU Input
sku_label = ttk.Label(product_frame, text="SKU:")
sku_label.grid(row=1, column=0, padx=5, pady=5, sticky='w')
sku_entry = ttk.Entry(product_frame)
sku_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

# Product Brand Input
product_brand_label = ttk.Label(product_frame, text="Product Brand:")
product_brand_label.grid(row=2, column=0, padx=5, pady=5, sticky='w')
product_brand_entry = ttk.Entry(product_frame)
product_brand_entry.grid(row=2, column=1, padx=5, pady=5, sticky='ew')

# Product Source Input
product_source_label = ttk.Label(product_frame, text="Product Source:")
product_source_label.grid(row=3, column=0, padx=5, pady=5, sticky='w')
product_source_entry = ttk.Entry(product_frame)
product_source_entry.grid(row=3, column=1, padx=5, pady=5, sticky='ew')

# Product Categories Input
product_categories_label = ttk.Label(product_frame, text="Product Categories:")
product_categories_label.grid(row=4, column=0, padx=5, pady=5, sticky='w')
product_categories_entry = ttk.Entry(product_frame)
product_categories_entry.grid(row=4, column=1, padx=5, pady=5, sticky='ew')

# Product Type Input
product_type_label = ttk.Label(product_frame, text="Product Type:")
product_type_label.grid(row=5, column=0, padx=5, pady=5, sticky='w')
product_type_entry = ttk.Entry(product_frame)
product_type_entry.grid(row=5, column=1, padx=5, pady=5, sticky='ew')

# Description Input
description_label = ttk.Label(product_frame, text="Description:")
description_label.grid(row=6, column=0, padx=5, pady=5, sticky='w')
description_entry = tk.Text(product_frame, height=5, width=40)
description_entry.grid(row=6, column=1, padx=5, pady=5, sticky='ew')

# Colors Dropdown
colors = ["Black", "Charcoal", "Dark Heather", "Navy", "Royal", "Anthracite", "Starlight", "Ash", "Sport Grey", "Althletic Grey", "Black Forest", "Royal Heather", "True Royal", "True Navy", "True Royal Heather", "Dark Silver HEather", "Diesel Grey", "Blacktop"]
color_vars = []
color_button = ttk.Button(root, text="Select Colors", command=select_colors)
color_button.grid(row=7, column=0)

# Default Color Dropdown
default_color_label = ttk.Label(root, text="Default Color:")
default_color_label.grid(row=8, column=0)
default_color_var = tk.StringVar()
default_color_dropdown = ttk.OptionMenu(root, default_color_var, *colors)
default_color_dropdown.grid(row=8, column=1)

# Price Input
price_label = ttk.Label(root, text="Price:")
price_label.grid(row=9, column=0)
price_entry = ttk.Entry(root)
price_entry.grid(row=9, column=1)

# Downloadable Dropdown
downloadable_label = ttk.Label(root, text="Downloadable:")
downloadable_label.grid(row=10, column=0)
downloadable_var = tk.StringVar()
downloadable_dropdown = ttk.OptionMenu(root, downloadable_var, "No", "Yes", "No")
downloadable_dropdown.grid(row=10, column=1)

# Sizes Dropdown
sizes = ["Youth-XS", "Youth-S", "Youth-M", "Youth-L", "Youth-XL", 
         "Ladies-XS", "Ladies-S", "Ladies-M", "Ladies-L", "Ladies-XL", 
         "Ladies-2XL", "Ladies-3XL", "Ladies-4XL", "Adult-XS", "Adult-S", 
         "Adult-M", "Adult-L", "Adult-XL", "Adult-2XL", "Adult-3XL", 
         "Adult-4XL", "Adult-5XL", "Adult-6XL"]
size_vars = []
size_button = ttk.Button(root, text="Select Sizes", command=select_sizes)
size_button.grid(row=11, column=0)

# Create a frame for buttons
button_frame = ttk.Frame(root)
button_frame.grid(row=12, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
button_frame.columnconfigure((0, 1), weight=1)

# Submit Button
submit_button = ttk.Button(button_frame, text="Submit", command=submit)
submit_button.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

# Export to Excel Button
export_button = ttk.Button(button_frame, text="Export to Excel", command=export_to_excel)
export_button.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

root.mainloop()