import tkinter as tk
from tkinter import messagebox


def delete_selected_items():
    selected_items = []
    for index, [item, var] in enumerate(items):
        if var.get():
            selected_items.append(item)
    
    if selected_items:
        confirmed = messagebox.askyesno("Confirmation", "Are you sure you want to delete the selected items?")
        if confirmed:
            items[:] = [[item, var] for item, var in items if item not in selected_items]
            update_checkbuttons()
    else:
        messagebox.showinfo("No Selection", "Please select items to delete.")


def update_checkbuttons():
    for cb, _ in checkbuttons:
        cb.destroy()
    checkbuttons.clear()
    
    for index, [item, var] in enumerate(items):
        cb = tk.Checkbutton(root, text=item, variable=var)
        cb.pack(anchor='w')
        checkbuttons.append([cb, var])


root = tk.Tk()
root.title("Delete Selected Items")

items = [["Item 1", tk.BooleanVar()], ["Item 2", tk.BooleanVar()], ["Item 3", tk.BooleanVar()],
         ["Item 4", tk.BooleanVar()], ["Item 5", tk.BooleanVar()]]

checkbuttons = []

for item, var in items:
    cb = tk.Checkbutton(root, text=item, variable=var)
    cb.pack(anchor='w')
    checkbuttons.append([cb, var])

delete_button = tk.Button(root, text="Delete Selected Items", command=delete_selected_items)
delete_button.pack()

root.mainloop()