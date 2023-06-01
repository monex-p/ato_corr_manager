import tkinter as tk
from tkinter.ttk import *
# from ttkwidgets import CheckboxTreeview
import pandas as pd

excel_path = "S:\\McCann Financial Group\\ATO -Correspondence\\2023\\ATO Correspondence Master 2023.xlsm"
sheet_name = "May2023"

def import_data(file_path, sh_name):
    try:
        data_frame = pd.read_excel(file_path, sh_name)
        return data_frame.values.tolist()
    except Exception as e:
        print("Error occurred while importing data:", e)
        return []


def create_table():
    root = tk.Tk()
    root.title("ATO Correspondence Manager")

    # Import data from Excel file
    data = import_data(excel_path, sheet_name)

    # Create a frame to hold the treeview and scrollbars
    frame = Frame(root)
    frame.pack(fill="both", expand=True)

    # Create a treeview widget
    treeview = Treeview(frame)
    treeview["columns"] = ("No","Name", "Client ID", "Subject", "Channel", 
                           "Issue Date", "Doc ID", "Importance Level", "ID Category",
                           "Partner","Manager","Email","Attended","Resolved")

    # Define column headings
    for col in treeview["columns"]:
        treeview.heading(col,text=col)
        treeview.column(col, stretch=tk.NO)

    # Add data to the table
    for client_data in data:
        treeview.insert("", "end", values=client_data)

    # # Configure column widths
    # for column in treeview["columns"]:
    #     treeview.column(column, width=1, stretch=False)  # Set initial width to 1

    # def update_column_widths():
    #     treeview.update()
    #     for column in treeview["columns"]:
    #         max_width = max(treeview.bbox(item, column)["width"] 
    #                         for item in treeview.get_children())
    #         treeview.column(column, width=max_width)

    # # Call update_column_widths when the treeview is resized
    # treeview.bind("<Configure>", lambda event: update_column_widths())

    # Adjust default column width to 0 (hide column)
    treeview.column("#0", width=0, stretch=tk.NO)
    treeview.column("No", width=40)
    treeview.column("Client ID", width=130)
    treeview.column("Channel", width=100)
    treeview.column("Issue Date", width=100)
    treeview.column("Doc ID", width=130)
    treeview.column("Importance Level", width=100)
    treeview.column("ID Category", width=100)
    treeview.column("Partner", width=130)
    treeview.column("Manager", width=130)
    treeview.column("Attended", width=100)
    # treeview.column("Resolved", width=100)

    # Add scrollbars to the table
    y_scrollbar = Scrollbar(frame, orient="vertical", command=treeview.yview)
    y_scrollbar.pack(side="right", fill="y")
    treeview.configure(yscrollcommand=y_scrollbar.set)

    x_scrollbar = Scrollbar(root, orient="horizontal", command=treeview.xview)
    x_scrollbar.pack(side="bottom", fill="x")
    treeview.configure(xscrollcommand=x_scrollbar.set)

    # Display the table
    treeview.pack(expand=True, fill="both")
    root.mainloop()

# Call the function to create and display the table
create_table()






