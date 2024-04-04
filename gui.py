import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage, ttk
import pandas as pd
import os, sys
from extractor import process_excel_files
from PIL import Image, ImageTk
import logging

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def load_excel_file(which_file):
    file_path = filedialog.askopenfilename(title=f"Select {which_file} Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        if which_file == "Dataset":
            global dataset_file_path
            dataset_file_path = file_path
            load_sheet_names(which_file, file_path, dataset_sheet_menu, dataset_sheet_name)
        elif which_file == "Employee List":
            global employee_file_path
            employee_file_path = file_path
            load_sheet_names(which_file, file_path, employee_sheet_menu, employee_sheet_name)
    else:
        messagebox.showinfo("Info", f"No {which_file} Excel file selected.")

def load_sheet_names(which_file, file_path, menu, variable):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        variable.set(sheet_names[0])  
        update_menu(menu, sheet_names, variable)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load {which_file} file: {str(e)}")

def update_menu(menu, sheet_names, variable):
    menu['menu'].delete(0, 'end')
    for name in sheet_names:
        menu['menu'].add_command(label=name, command=lambda value=name: variable.set(value))

def select_output_file():
    global output_file_path
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_file_path:
        output_file_name.set(output_file_path.split('/')[-1])
    else:
        messagebox.showinfo("Info", "No output file selected.")

def process_files():
    if not dataset_file_path or not dataset_sheet_name.get() or not employee_file_path or not employee_sheet_name.get() or not output_file_path:
        messagebox.showwarning("Warning", "Please select all required files, sheet names, and output file before processing.")
        return
    try:
        process_excel_files(dataset_file_path, dataset_sheet_name.get(), employee_file_path, employee_sheet_name.get(), output_file_path)
        messagebox.showinfo("Success", "Data processing completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during processing: {e}")
def launch_gui():
    global dataset_file_path, employee_file_path, output_file_path
    global dataset_sheet_name, employee_sheet_name, output_file_name
    global dataset_sheet_menu, employee_sheet_menu

    root = tk.Tk()
    root.title("EXTRACTOR")

    ico_path = resource_path("fav.ico")  
    icon = Image.open(ico_path)
    photo = ImageTk.PhotoImage(icon)
    root.tk.call('wm', 'iconphoto', root._w, photo)


    root.geometry("618x450")

    bg_color = "#57585a"
    text_color = "#ffffff" 
    root.configure(background=bg_color)

    # Ensure logo_image has global scope to prevent garbage-collection
    global logo_image
    logo_image = ImageTk.PhotoImage(Image.open(resource_path("logo.png")))

    logo_label = tk.Label(root, image=logo_image, bg=bg_color)
    logo_label.pack(pady=10)
    ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=5, pady=10)

    dataset_file_path = ""
    employee_file_path = ""
    output_file_path = ""
    dataset_sheet_name = tk.StringVar(root)
    employee_sheet_name = tk.StringVar(root)
    output_file_name = tk.StringVar(root)

    # Dataset file selection frame
    frame_dataset = tk.Frame(root, bg=bg_color)
    instruction_dataset = tk.Label(frame_dataset, text="Selectionner le fichier Excel extracté du GTE:", bg=bg_color, fg=text_color)
    instruction_dataset.pack(side=tk.LEFT, padx=(10, 5))
    button_dataset = tk.Button(frame_dataset, text="Parcourir", command=lambda: load_excel_file("Dataset"),fg=bg_color)
    button_dataset.pack(side=tk.LEFT, padx=5)
    dataset_sheet_menu = tk.OptionMenu(frame_dataset, dataset_sheet_name, "No sheets available")
    dataset_sheet_menu.config( fg=text_color)
    dataset_sheet_menu.pack(side=tk.LEFT, padx=5)
    frame_dataset.pack(pady=5, fill=tk.X)

    # Employee list file selection frame
    frame_employee = tk.Frame(root, bg=bg_color)
    instruction_employee = tk.Label(frame_employee, text="Selectionner le fichier Excel contenant la liste des employés :", bg=bg_color, fg=text_color)
    instruction_employee.pack(side=tk.LEFT, padx=(10, 5))
    button_employee = tk.Button(frame_employee, text="Parcourir", command=lambda: load_excel_file("Employee List"), fg=bg_color)
    button_employee.pack(side=tk.LEFT, padx=5)
    employee_sheet_menu = tk.OptionMenu(frame_employee, employee_sheet_name, "No sheets available")
    employee_sheet_menu.config( fg=text_color)
    employee_sheet_menu.pack(side=tk.LEFT, padx=5)
    frame_employee.pack(pady=5, fill=tk.X)

    # Separator before Output File Selection
    ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=5, pady=10)

    # Output file selection frame
    frame_output = tk.Frame(root, bg=bg_color)
    instruction_output = tk.Label(frame_output, text="Selectionner le nom et le chemin du fichier désiré :", bg=bg_color,fg=text_color)
    instruction_output.pack(side=tk.LEFT, padx=(10, 5))
    button_output = tk.Button(frame_output, text="Parcourir", command=select_output_file,  fg=bg_color)
    button_output.pack(side=tk.LEFT, padx=5)
    output_file_entry = tk.Entry(frame_output, textvariable=output_file_name, state='readonly')
    output_file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    frame_output.pack(pady=5, fill=tk.X)

    # "EXTRACT" button
    extract_button = tk.Button(root, text="Démarrer", command=process_files, height=2, width=20, font=("Helvetica", 12))
    extract_button.pack(pady=20)

    root.mainloop()



if __name__ == "__main__":
    launch_gui()
