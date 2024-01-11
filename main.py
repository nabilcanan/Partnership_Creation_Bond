import os
import tkinter as tk
from tkinter import ttk
from compare_files import compare_files
from compare_contract_file import compare_contract_file

# Create main window
window = tk.Tk()
window.title("Comparing NeoTech Contract Files")
window.configure(bg="white")
window.geometry("930x550")

# Create a canvas and a vertical scrollbar
canvas = tk.Canvas(window)
scrollbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
canvas.configure(bg="white", yscrollcommand=scrollbar.set)

canvas.configure(yscrollcommand=scrollbar.set)
inner_frame = tk.Frame(canvas, bg="white")
canvas.create_window((window.winfo_width() / 2, 0), window=inner_frame, anchor="n")

inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)

# Canvas - Scrollbar
canvas.configure(yscrollcommand=scrollbar.set)
scrollbar.configure(command=canvas.yview)


def _on_mousewheel(event):
    canvas.yview_scroll(-1 * (event.delta // 120), "units")


def open_powerpoint():
    powerpoint_file = r'P:\Partnership_Python_Projects\Creation BOND File Program\Creation_BOND_Python_Program.pptx'

    try:
        os.startfile(powerpoint_file)
    except Exception as e:
        print(f"Error: {e}")


# Bind the function to the MouseWheel event, to make our scrolling function more applicable
canvas.bind_all("<MouseWheel>", _on_mousewheel)

# Widgets
style = ttk.Style()

style.configure("TButton", font=("Rupee", 16, "bold"), width=30, height=2, background="white")
style.map("TButton", foreground=[('active', 'red')], background=[('active', 'blue')])

# Add a title label
title_label = ttk.Label(inner_frame, text="Welcome Partnership Member!",
                        font=("Rupee", 26, "underline"), background="white", foreground="#103d81")
title_label.pack(pady=(20, 10), padx=20)  # Adjust the padding values as necessary

open_powerpoint_button = ttk.Button(inner_frame, text='Open PowerPoint Instructions',
                                    command=open_powerpoint)
open_powerpoint_button.pack(pady=10)

# Add instructions label
instructions_label = tk.Label(inner_frame,
                              text="Instructions:\n"
                                   "To identify missing IPNs between the last and current RAW BOND files and add the \n"
                                   "'Item_Type_Changed_To' and 'Sourced_Type_Changed_To' columns, follow these steps:\n"
                                   "1. Select the Last RAW BOND Creation File.\n"
                                   "2. Select the Current RAW BOND Creation File.",
                              font=("Rupee", 16), background="white")
instructions_label.pack(pady=(10, 20), padx=20)  # Adjust the padding values as necessary

# Add a button to trigger file selection and comparison
run_queries_button = ttk.Button(inner_frame, text="Compare Files", command=compare_files, style="TButton")
run_queries_button.pack(pady=(10, 20), padx=20)

instructions_label = tk.Label(inner_frame,
                              text="Instructions:\n"
                                   "To cross-reference missing IPNs from the 'Removed From Last File' sheet with our\n "
                                   "Active Supplier Contracts, and to update pricing "
                                   "for IPNs in the 'Detail' sheet:\n"
                                   "1. Select the new file containing the 'Removed From Prev File' sheet.\n"
                                   "2. Select the weekly 'Active Supplier Contracts' file.",
                              font=("Rupee", 16), background="white")
instructions_label.pack(pady=(10, 20), padx=20)  # Adjust the padding values as necessary

compare_contract_file_button = ttk.Button(inner_frame, text='Compare w/ Creation Contract',
                                          command=compare_contract_file)
compare_contract_file_button.pack(pady=10)

# Run the GUI window
window.mainloop()
