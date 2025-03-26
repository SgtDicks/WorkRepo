import os
import sys
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
from docxtpl import DocxTemplate
import webbrowser
from PIL import Image, ImageTk  # Requires Pillow

def resource_path(relative_path):
    """
    Get absolute path to resource, works for development and for PyInstaller.
    Make sure to include the resource files (template, branding image) with your PyInstaller build.
    """
    try:
        base_path = sys._MEIPASS  # PyInstaller's temporary folder.
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_bundled_template_path():
    """
    Check if an updated bundled template exists in the user folder.
    If so, return its path; otherwise, return the path to the originally bundled template.
    """
    user_template_dir = os.path.join(os.path.expanduser("~"), ".docx_template_generator")
    if not os.path.exists(user_template_dir):
        os.makedirs(user_template_dir)
    updated_template_path = os.path.join(user_template_dir, "default_template.docx")
    if os.path.exists(updated_template_path):
        return updated_template_path
    else:
        return resource_path("default_template.docx")

def update_bundled_template():
    """
    Let the user choose a new DOCX file and copy it to the user folder as the updated template.
    """
    new_template = filedialog.askopenfilename(
        title="Select New Bundled Template",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )
    if not new_template:
        return

    try:
        user_template_dir = os.path.join(os.path.expanduser("~"), ".docx_template_generator")
        if not os.path.exists(user_template_dir):
            os.makedirs(user_template_dir)
        updated_template_path = os.path.join(user_template_dir, "default_template.docx")
        shutil.copyfile(new_template, updated_template_path)
        messagebox.showinfo("Success", "Bundled template updated successfully.")
        status_var.set("Bundled template updated.")
        # If using the bundled template, update the display text.
        if use_bundled_template_var.get():
            template_entry.config(state="disabled")
            template_entry.delete(0, tk.END)
            template_entry.insert(0, "Using updated bundled template")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred updating the bundled template:\n{e}")
        status_var.set("Error updating bundled template.")

def show_instructions():
    """Display a Help window with instructions and a clickable hyperlink."""
    instr_window = tk.Toplevel(root)
    instr_window.title("Instructions")
    instr_window.geometry("600x350")
    instr_window.resizable(False, False)
    
    instructions = (
        "Instructions:\n\n"
        "1. Upload your CV to:\n"
    )
    instr_label = ttk.Label(instr_window, text=instructions, justify="left")
    instr_label.pack(anchor="w", padx=10, pady=5)
    
    # Clickable hyperlink label
    url = "https://chatgpt.com/g/g-67be66c3072c8191b2cccf8570448801-cv-to-json-extractor"
    link_label = tk.Label(instr_window, text=url, fg="blue", cursor="hand2")
    link_label.pack(anchor="w", padx=20)
    link_label.bind("<Button-1>", lambda e: webbrowser.open(url))
    
    additional_instructions = (
        "\n\n2. Download the generated JSON file.\n\n"
        "3. In this application:\n"
        "   - Click 'Browse...' next to 'JSON File' to select your downloaded JSON file.\n"
        "   - If you have your own DOCX template, uncheck 'Use Bundled Template' and click 'Browse...'\n"
        "     to select your template file. Otherwise, leave it checked to use the bundled template.\n"
        "   - To update the bundled template for future use, click the 'Update Bundled Template' button.\n"
        "   - Click 'Browse...' next to 'Output Folder' to select the folder where you want the generated document.\n\n"
        "4. The app will generate a document named 'RPI_<name>_CV.docx', where <name> is extracted from your JSON file.\n\n"
        "5. Click 'Generate Document' to create your personalized CV document."
    )
    add_instr_label = ttk.Label(instr_window, text=additional_instructions, justify="left")
    add_instr_label.pack(anchor="w", padx=10, pady=5)

def select_json():
    json_file = filedialog.askopenfilename(
        title="Select JSON File",
        filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
    )
    json_entry.delete(0, tk.END)
    json_entry.insert(0, json_file)
    status_var.set("JSON file selected.")

def select_template():
    template_file = filedialog.askopenfilename(
        title="Select DOCX Template File",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
    )
    template_entry.config(state="normal")
    template_entry.delete(0, tk.END)
    template_entry.insert(0, template_file)
    status_var.set("Template file selected.")

def select_output():
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_folder)
    status_var.set("Output folder selected.")

def update_template_state():
    """Enable or disable the custom template selection based on the checkbox."""
    if use_bundled_template_var.get():
        template_entry.config(state="disabled")
        template_browse_btn.config(state="disabled")
        template_entry.delete(0, tk.END)
        template_entry.insert(0, "Using bundled template")
    else:
        template_entry.config(state="normal")
        template_browse_btn.config(state="normal")
        template_entry.delete(0, tk.END)
        status_var.set("Please select a template file.")

def generate_doc():
    json_path = json_entry.get()
    template_path = template_entry.get()
    output_folder = output_entry.get()
    
    if not (json_path and output_folder):
        messagebox.showerror("Missing Input", "Please select a JSON file and output folder.")
        status_var.set("Error: Missing JSON file or output folder.")
        return
    
    # Load the JSON data
    try:
        with open(json_path, 'r') as f:
            context = json.load(f)
        status_var.set("JSON loaded successfully.")
    except Exception as e:
        messagebox.showerror("Error Reading JSON", f"An error occurred reading the JSON file:\n{e}")
        status_var.set("Error reading JSON.")
        return
    
    # Extract the user's name from the JSON for the output filename
    user_name = context.get("name", "").strip().replace(" ", "_")
    if not user_name:
        messagebox.showerror("Missing Data", "The JSON file does not contain a valid 'name' field.")
        status_var.set("Error: 'name' field missing in JSON.")
        return
    
    output_filename = f"RPI_{user_name}_CV.docx"
    output_path = os.path.join(output_folder, output_filename)
    
    # Determine which template file to use
    if use_bundled_template_var.get():
        template_path = get_bundled_template_path()
    else:
        if not template_path:
            messagebox.showerror("Missing Input", "Please select a DOCX template file.")
            status_var.set("Error: Template file missing.")
            return
    
    # Render the document using the template and JSON context
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_path)
        messagebox.showinfo("Success", f"Document generated successfully:\n{output_path}")
        status_var.set("Document generated successfully.")
    except Exception as e:
        messagebox.showerror("Error Generating Document", f"An error occurred generating the document:\n{e}")
        status_var.set("Error generating document.")

# Initialize the main application window
root = tk.Tk()
root.title("RPI CV Generator")
root.geometry("700x500")
root.resizable(False, False)

# Create a menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# File Menu (with an Exit option)
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Exit", command=root.quit)
menu_bar.add_cascade(label="File", menu=file_menu)

# Help Menu with clickable instructions
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="Instructions", command=show_instructions)
menu_bar.add_cascade(label="Help", menu=help_menu)

# Use ttk style for a modern look
style = ttk.Style(root)
style.theme_use("clam")

# Main frame for the widgets
main_frame = ttk.Frame(root, padding="20")
main_frame.pack(fill=tk.BOTH, expand=True)

# Branding image at the top
try:
    branding_path = resource_path("RPI_Logo_HOR_Solid_-WITH-WHITE-CORNER_Tagline.png")
    branding_img = Image.open(branding_path)
    # Determine the correct resampling filter for resizing
    try:
        resample = Image.Resampling.LANCZOS
    except AttributeError:
        resample = Image.ANTIALIAS
    # Resize image as needed (here, resizing to 300x100 pixels)
    branding_img = branding_img.resize((300, 50), resample)
    branding_photo = ImageTk.PhotoImage(branding_img)
    branding_label = ttk.Label(main_frame, image=branding_photo)
    branding_label.image = branding_photo  # keep a reference!
    branding_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
except Exception as e:
    print("Error loading branding image:", e)
# Header label (shifted down by one row)
header = ttk.Label(main_frame, text="Docx Template Generator", font=("Helvetica", 16, "bold"))
header.grid(row=1, column=0, columnspan=3, pady=(0, 20))

# JSON File selection
ttk.Label(main_frame, text="JSON File:").grid(row=2, column=0, sticky="e", padx=(0, 10))
json_entry = ttk.Entry(main_frame, width=50)
json_entry.grid(row=2, column=1, padx=(0, 10))
ttk.Button(main_frame, text="Browse...", command=select_json).grid(row=2, column=2)

# Template File selection
ttk.Label(main_frame, text="Template File:").grid(row=3, column=0, sticky="e", padx=(0, 10), pady=10)
template_entry = ttk.Entry(main_frame, width=50)
template_entry.grid(row=3, column=1, padx=(0, 10), pady=10)
template_browse_btn = ttk.Button(main_frame, text="Browse...", command=select_template)
template_browse_btn.grid(row=3, column=2, pady=10)

# Bundled Template Checkbutton and Update button
use_bundled_template_var = tk.BooleanVar(value=True)
bundled_check = ttk.Checkbutton(
    main_frame,
    text="Use Bundled Template",
    variable=use_bundled_template_var,
    command=update_template_state
)
bundled_check.grid(row=4, column=1, sticky="w", pady=(0, 10))

update_template_btn = ttk.Button(main_frame, text="Update Bundled Template", command=update_bundled_template)
update_template_btn.grid(row=4, column=2, pady=(0, 10))
update_template_state()  # Set the initial state

# Output Folder selection
ttk.Label(main_frame, text="Output Folder:").grid(row=5, column=0, sticky="e", padx=(0, 10))
output_entry = ttk.Entry(main_frame, width=50)
output_entry.grid(row=5, column=1, padx=(0, 10))
ttk.Button(main_frame, text="Browse...", command=select_output).grid(row=5, column=2)

# Generate Document button
generate_button = ttk.Button(main_frame, text="Generate Document", command=generate_doc)
generate_button.grid(row=6, column=0, columnspan=3, pady=20)

# Status bar for feedback
status_var = tk.StringVar()
status_var.set("Ready")
status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor="w", padding=5)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()

