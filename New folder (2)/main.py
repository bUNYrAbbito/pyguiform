import tkinter as tk
from openpyxl import Workbook

def save_data():
    name = name_entry.get()
    address = address_entry.get()
    contact_no = contact_entry.get()
    email = email_entry.get()

    # Create a new Excel workbook or load an existing one
    workbook = Workbook()
    sheet = workbook.active

    # Append the data to the Excel sheet
    sheet.append([name, address, contact_no, email])

    # Save the data to an Excel file
    workbook.save("data.xlsx")

    # Clear the entry fields
    name_entry.delete(0, tk.END)
    address_entry.delete(0, tk.END)
    contact_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)

# Create the main window
root = tk.Tk()
root.title("Data Entry Form")

# Create labels and entry fields
name_label = tk.Label(root, text="Name:")
name_label.pack()
name_entry = tk.Entry(root)
name_entry.pack()

address_label = tk.Label(root, text="Address:")
address_label.pack()
address_entry = tk.Entry(root)
address_entry.pack()

contact_label = tk.Label(root, text="Contact No:")
contact_label.pack()
contact_entry = tk.Entry(root)
contact_entry.pack()

email_label = tk.Label(root, text="Email:")
email_label.pack()
email_entry = tk.Entry(root)
email_entry.pack()

# Create a button to save data
save_button = tk.Button(root, text="Save Data", command=save_data)
save_button.pack()

# Start the Tkinter main loop
root.mainloop()
