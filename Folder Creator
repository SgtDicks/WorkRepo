import os
import tkinter as tk
import json
from tkinter import filedialog, ttk

global directory_entry  # Add this line at the beginning of your code


def save_states():
    states = {item: tree.item(item, "values")[1] for item in tree.get_children("")}

    for item in tree.get_children(""):
        for subitem in tree.get_children(item):
            states[subitem] = tree.item(subitem, "values")[1]

    with open('states.json', 'w') as f:
        json.dump(states, f)

    status_label.config(text="States saved successfully!")
    
    
def save_checked_options():
    checked_options = {}
    for item in tree.get_children():
        folder_name, is_checked = tree.item(item, "values")
        is_checked = is_checked == 'True' if is_checked in ['True', 'False'] else bool(int(is_checked))
        print(f"Folder: {folder_name}, Checked: {is_checked}")  # Print each folder's name and checked state
        if is_checked:
            subfolders = [tree.item(child, "values")[0] for child in tree.get_children(item) 
                          if tree.item(child, "values")[1] == 'True' or tree.item(child, "values")[1] == '1']
            print(f"Subfolders: {subfolders}")  # Print the checked subfolders
            checked_options[folder_name] = subfolders
    
    print(checked_options)  # Print the final checked_options dictionary

    file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if file_path:
        with open(file_path, 'w') as f:
            json.dump(checked_options, f)
        status_label.config(text="Checked options saved successfully!")
    else:
        status_label.config(text="Save cancelled.")
        
def load_states():
    try:
        with open('states.json', 'r') as f:
            states = json.load(f)

        for item, state in states.items():
            folder_name = tree.item(item, "text")[3:]
            tree.item(item, values=(folder_name, state), text=('☑ ' if state == '1' else '☐ ') + folder_name)
        
        status_label.config(text="States loaded successfully!")

    except FileNotFoundError:
        status_label.config(text="States file not found!")
        
def load_checked_options():
    file_path = filedialog.askopenfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if file_path:
        with open(file_path, 'r') as f:
            checked_options = json.load(f)
        
        for item in tree.get_children():
            folder_name, _ = tree.item(item, "values")
            if folder_name in checked_options:
                tree.item(item, values=(folder_name, '1'), text='☑ ' + folder_name)
                for child in tree.get_children(item):
                    child_folder_name, _ = tree.item(child, "values")
                    if child_folder_name in checked_options[folder_name]:
                        tree.item(child, values=(child_folder_name, '1'), text='☑ ' + child_folder_name)
                    else:
                        tree.item(child, values=(child_folder_name, '0'), text='☐ ' + child_folder_name)
            else:
                tree.item(item, values=(folder_name, '0'), text='☐ ' + folder_name)
        
        status_label.config(text="Checked options loaded successfully!")
    else:
        status_label.config(text="Load cancelled.")        
 
def toggle_check(item):
    item_values = tree.item(item, "values")
    folder_name = item_values[0]
    is_checked = item_values[1] == '1'  # Corrected this line to convert the string to boolean
    new_state = not is_checked
    tree.item(item, values=(folder_name, str(int(new_state))), text=('☑ ' if new_state else '☐ ') + folder_name)  # Corrected this line to store the boolean as string
    for child in tree.get_children(item):
        toggle_check(child)



def on_item_clicked(event):
    item = tree.selection()[0]
    if item:
        toggle_check(item)


def create_folders(root_folder, directory):
    root_path = os.path.join(directory, root_folder)
    if not os.path.exists(root_path):
        os.makedirs(root_path)
    
    for item in tree.get_children():
        folder_name, is_checked = tree.item(item, "values")
        is_checked = is_checked == 'True' if is_checked in ['True', 'False'] else bool(int(is_checked))
        if is_checked:
            folder_path = os.path.join(root_path, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            for child in tree.get_children(item):
                subfolder_name, is_subfolder_checked = tree.item(child, "values")
                is_subfolder_checked = is_subfolder_checked == 'True' if is_subfolder_checked in ['True', 'False'] else bool(int(is_subfolder_checked))
                if is_subfolder_checked:
                    subfolder_path = os.path.join(folder_path, subfolder_name)
                    os.makedirs(subfolder_path, exist_ok=True)
    
    status_label.config(text=f"Folders created at {root_path}")


def browse_directory():
    folder_selected = filedialog.askdirectory()
    path_entry.delete(0, tk.END)
    path_entry.insert(0, folder_selected)


def on_create_folders():
    root_folder = root_folder_entry.get() or "Root Folder"
    directory = path_entry.get()  # Retrieve directory from path_entry
    create_folders(root_folder, directory)




root = tk.Tk()
root.title("Folder Creator")

# Treeview for folders
tree = ttk.Treeview(root, columns=("Folder Name", "Check"))
tree.heading("#1", text="Folder Name")
tree.column("#1", width=0, stretch=tk.NO)
tree.heading("#2", text="Check")
tree.column("#2", width=0, stretch=tk.NO)

# Add items (folders) to the treeview with your provided folder structure
folders = [
    ("00 RP Contract Files", [
        "0.1 Client Brief",
        "0.2 EOI Submission",
        ("0.3 RFP Submission", [
            "1 Fee Proposal",
            "2 RP Consultancy Agreement",
            "3 SubConsultant Agreements",
            "4 Fee Variations Proposals"
        ]),
        ("0.4 RP Contract", [
            "1 Contract",
            "2 EOTs",
            "3 Variations",
            "4 Insurances",
            "5 Security",
            "6 Payment",
            "7 Consultant",
            "8 Other"
        ]),
        ("0.5 QMS Reports", [
            "1 Project Setup",
            "2 PMP",
            "3 Case Study",
            "4 Progress Report",
            "5 WHS & E Report",
            "6 Appoint Subconsultant",
            "7 Project Close Out",
            "8 Project Suspension Restart",
            "9 Other"
        ]),
        "0.6 Correspondence"
    ]),
    ("01 Project Setup", [
        "1.1 Contact List",
        "1.2 Comms Plan",
        "1.3 Change Management",
        "1.4 Other"
    ]),
    ("02 Project Initiation", [
        "2.1 Background",
        "2.2 Client",
        "2.3 Other"
    ]),
    ("03 Stakeholder & Comms", [
        "3.1 Comms Notices, publications",
        ("3.2 Consultation", [
            "1 Client",
            "2 Stakeholder",
            "3 Meetings",
            "4 Other"
        ]),
        "3.3 Change Requests",
        "3.4 Other"
    ]),
    ("04 Advisory", [
        "4.1 Studies",
        "4.2 Business Case",
        "4.3 Site Analysis & Due Diligence",
        "4.4 Legal",
        "4.5 [insert topic]"
    ]),
    ("05 Design", [
        "5.0 Design Brief",
        "5.1 Site Conditions",
        ("5.2 Master Plan", [
            "1 Drawings",
            "2 Reports",
            "3 Project History"
        ]),
        ("5.3 Concept", [
            "1 Drawings",
            "2 Reports",
            "3 Final Set"
        ]),
        ("5.4 Schematic", [
            "1 Drawings",
            "2 Report",
            "3 Planning",
            "4 Final Set"
        ]),
        ("5.5 Design Development", [
            "1 Drawings",
            "2 Specification",
            "3 Report",
            "4 Final Set"
        ]),
        ("5.6 Tender", [
            "1 Drawings",
            "2 Specification",
            "3 Reports",
            "4 For Information",
            "5 Final Set"
        ]),
        ("5.7 IFC", [
            "1 Drawings",
            "2 Specification",
            "3 Reports",
            "4 Shopdrawings"
        ]),
        "5.8 Value Management",
        "5.9 Meetings",
        "5A Authorities"
    ]),
    ("06 Procurement", [
        "6.1 General",
        "6.2 Consultants",
        ("Template Folder", [
            "1 RFP",
            "2 Proposal",
            "3 Evaluation"
        ]),
        ("6.3 Contractor", [
            "1 EOI",
            "2 RFT",
            "3 Evaluation",
            "4 Contract Negotiation"
        ]),
        "6.4 Novation"
    ]),
    ("07 Consultant", [
        ("7.1 Consultants", [
            "1 Contract",
            "2 EOTs",
            "3 Insurances",
            "4 Security",
            "5 Payment",
            "6 Service Delivery Plan",
            "7 Other"
        ]),
        "7.2 Other"
    ]),
    ("08 Construction", [
        ("8.1 Contract Administration", [
            "1 Contract",
            "2 EOTs",
            "3 Variations",
            "4 Insurances",
            "5 Security",
            "6 Payment",
            "7 Notices",
            "8 Other"
        ]),
        "8.2 Management Plans",
        "8.3 Site Access",
        "8.4 Approvals",
        "8.5 Meetings",
        "8.6 Reports",
        "8.7 RFIs",
        "8.8 Samples Prototypes",
        "8.9 Site Inspections"
    ]),
    ("09 Program", [
        "9.1 Master Program",
        "9.2 Design Program",
        "9.3 Construction Program",
        "9.4 Other"
    ]),
    ("10 Financial", [
        "10.1 Budget",
        "10.2 Cost Estimates",
        "10.3 Cost Reports",
        "10.4 BoQ",
        "10.5 Meetings",
        "10.6 Other"
    ]),
    ("11 Reporting & Governance", [
        "11.1 Client",
        "11.2 Consultant",
        "11.3 Contractor",
        "11.4 Other"
    ]),
    ("12 WHSEQ & IR", [
        "12.1 OHS",
        "12.2 Environmental",
        "12.3 QMS",
        "12.4 Industrial Relations",
        "12.5 Other"
    ]),
    ("13 Images", [
        "13.1 Professional Photos",
        "13.2 Marketing",
        "13.3 Progress"
    ]),
    ("14 Risk Management", [
        "14.1 Plan",
        "14.2 Register",
        "14.3 Other"
    ]),
    ("15 Project Completion", [
        "15.1 Occupancy",
        "15.2 Defects",
        "15.3 ASBUILTs",
        "15.4 O&Ms",
        "15.5 Commissioning",
        "15.6 Checklists",
        "15.7 PC Certificate",
        "15.8 POE"
    ])
    
    # ... (continue adding other folders similarly)
]

for folder, subfolders in folders:
    if isinstance(folder, tuple):
        parent_folder, child_folders = folder
    else:
        parent_folder = folder
        child_folders = []

    parent = tree.insert("", tk.END, text='☑ ' + parent_folder, values=(parent_folder, True))
    
    for subfolder in subfolders:
        if isinstance(subfolder, tuple):
            subfolder_name, child_subfolders = subfolder
            sub_parent = tree.insert(parent, tk.END, text='☑ ' + subfolder_name, values=(subfolder_name, True))
            for child_subfolder in child_subfolders:
                tree.insert(sub_parent, tk.END, text='☑ ' + child_subfolder, values=(child_subfolder, True))
        else:
            tree.insert(parent, tk.END, text='☑ ' + subfolder, values=(subfolder, True))


tree.pack(padx=20, pady=5, fill="both", expand=True)
tree.bind("<Button-1>", on_item_clicked)

# Root Folder Name
root_folder_frame = tk.Frame(root)  # Create a new frame for the root folder widgets
root_folder_frame.pack(fill="x", padx=20, pady=5)  # Pack the frame

root_folder_label = tk.Label(root_folder_frame, text="Root Folder Name:")
root_folder_label.pack(side=tk.LEFT)

root_folder_entry = tk.Entry(root_folder_frame)  # Define the root_folder_entry widget
root_folder_entry.pack(fill="x", expand=True, side=tk.LEFT)  # Make the entry expand with the window

# Directory Path
directory_frame = tk.Frame(root)  # Create a new frame for the directory widgets
directory_frame.pack(fill="x", padx=20, pady=5)  # Pack the frame

directory_label = tk.Label(directory_frame, text="Directory Path:")
directory_label.pack(side=tk.LEFT)

directory_entry = tk.Entry(directory_frame)  # Define the directory_entry widget
directory_entry.pack(fill="x", expand=True, side=tk.LEFT)  # Make the entry expand with the window

browse_button = tk.Button(directory_frame, text="Browse", command=browse_directory)
browse_button.pack(side=tk.LEFT)

# Buttons Frame
buttons_frame = tk.Frame(root)  # Create a new frame for the buttons
buttons_frame.pack(fill="x", padx=20, pady=5)  # Pack the frame

create_folders_button = tk.Button(buttons_frame, text="Create Folders", command=on_create_folders)
create_folders_button.pack(side=tk.LEFT)

save_checked_button = tk.Button(buttons_frame, text="Save Checked Options", command=save_checked_options)
save_checked_button.pack(side=tk.LEFT)

load_checked_button = tk.Button(buttons_frame, text="Load Checked Options", command=load_checked_options)
load_checked_button.pack(side=tk.LEFT)

# Status Label
status_label = tk.Label(root, text="", wraplength=500)
status_label.pack(pady=5, fill="y")

root.mainloop()

