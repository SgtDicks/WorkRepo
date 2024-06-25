import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
import subprocess
import os
import threading
import webbrowser
import winreg
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed

magick_path = None
conversion_thread = None
stop_conversion = False
output_dir = None
debug_mode = False
registry_key = r"SOFTWARE\HEICtoJPEGConverter"
registry_value_name = "ImageMagickPath"
max_workers = 10  # Number of threads to use for conversion

def search_imagemagick():
    common_dirs = [
        r"C:\Program Files",
        r"C:\Program Files (x86)"
    ]
    for directory in common_dirs:
        pattern = os.path.join(directory, "ImageMagick*")
        for path in glob.glob(pattern):
            magick_exe = os.path.join(path, "magick.exe")
            if os.path.exists(magick_exe):
                if debug_mode:
                    print(f"Found ImageMagick at {magick_exe}")
                return magick_exe
    return None

def load_magick_path_from_registry():
    global magick_path
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_key) as key:
            magick_path = winreg.QueryValueEx(key, registry_value_name)[0]
            if debug_mode:
                print(f"Loaded ImageMagick path from registry: {magick_path}")
    except FileNotFoundError:
        magick_path = search_imagemagick()
        if magick_path:
            save_magick_path_to_registry(magick_path)
        else:
            prompt_to_download_imagemagick()

def save_magick_path_to_registry(path):
    global magick_path
    magick_path = path
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, registry_key) as key:
        winreg.SetValueEx(key, registry_value_name, 0, winreg.REG_SZ, magick_path)
    if debug_mode:
        print(f"Saved ImageMagick path to registry: {magick_path}")

def set_magick_path():
    global magick_path
    path = filedialog.askopenfilename(title="Select ImageMagick Executable", filetypes=[("Executables", "*.exe")])
    if path:
        save_magick_path_to_registry(path)
        messagebox.showinfo("Info", f"ImageMagick set to: {path}")
    else:
        messagebox.showwarning("Warning", "No executable selected. Please select the ImageMagick executable.")

def prompt_to_download_imagemagick():
    response = messagebox.askyesno("ImageMagick Not Found", "ImageMagick was not found. Would you like to download and install the ImageMagick Windows binaries?")
    if response:
        webbrowser.open("https://imagemagick.org/script/download.php")
    else:
        messagebox.showerror("Error", "ImageMagick is required to use this application. Please install it and try again.")
        root.quit()

def set_output_directory():
    global output_dir
    output_dir = filedialog.askdirectory(title="Select Output Folder")
    if output_dir:
        output_dir_entry.delete(0, tk.END)
        output_dir_entry.insert(0, output_dir)
    else:
        messagebox.showwarning("Warning", "No output directory selected. Please select the output directory.")

def convert_file(heic_path, output_dir, iid):
    global magick_path
    if heic_path.lower().endswith(".heic"):
        jpeg_path = os.path.join(output_dir, os.path.basename(heic_path).rsplit('.', 1)[0] + '.jpg')
        try:
            result = subprocess.run(
                [magick_path, "convert", heic_path, jpeg_path],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            if debug_mode:
                print(f"Converted {heic_path} to {jpeg_path}")
                print(f"stdout: {result.stdout.decode()}")
                print(f"stderr: {result.stderr.decode()}")
            return iid, "✔"
        except subprocess.CalledProcessError as e:
            if debug_mode:
                print(f"Error converting {heic_path}: {e}")
            return iid, "✖"
        except FileNotFoundError as e:
            if debug_mode:
                print(f"ImageMagick not found: {e}")
            return iid, "✖"
    elif transfer_other_files.get():
        try:
            destination_path = os.path.join(output_dir, os.path.basename(heic_path))
            os.makedirs(os.path.dirname(destination_path), exist_ok=True)
            with open(heic_path, 'rb') as fsrc, open(destination_path, 'wb') as fdst:
                fdst.write(fsrc.read())
            if debug_mode:
                print(f"Transferred {heic_path} to {destination_path}")
            return iid, "✔ (Transferred)"
        except Exception as e:
            if debug_mode:
                print(f"Error transferring {heic_path}: {e}")
            return iid, "✖ (Transfer Failed)"
    return iid, "✖"

def convert_heic_to_jpeg(files, output_dir):
    global stop_conversion
    success = True
    stop_conversion = False
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(convert_file, heic_path, output_dir, iid): (heic_path, iid) for heic_path, iid in files}
        
        for future in as_completed(futures):
            heic_path, iid = futures[future]
            if stop_conversion:
                break
            try:
                iid, status = future.result()
                filetree.set(iid, column="Status", value=status)
                if status.startswith("✖"):
                    success = False
            except Exception as e:
                if debug_mode:
                    print(f"Error processing {heic_path}: {e}")
                success = False
    
    return success

def start_conversion():
    global conversion_thread, output_dir
    if not output_dir:
        messagebox.showwarning("Warning", "Output directory not set. Please set it first.")
        return
    
    files = [(filetree.item(iid)['values'][2], iid) for iid in filetree.get_children()]
    if not files:
        messagebox.showwarning("Warning", "No files to convert.")
        return

    # Run the conversion in a separate thread
    conversion_thread = threading.Thread(target=convert_and_notify, args=(files, output_dir))
    conversion_thread.start()

def stop_conversion():
    global stop_conversion
    stop_conversion = True
    if conversion_thread and conversion_thread.is_alive():
        messagebox.showinfo("Info", "Conversion process stopped.")

def convert_and_notify(files, output_dir):
    if convert_heic_to_jpeg(files, output_dir):
        messagebox.showinfo("Success", "All files have been converted successfully!")
    else:
        messagebox.showerror("Error", "Some files were not converted. Check the console for details.")

def add_files():
    files = filedialog.askopenfilenames(title="Select Image Files", filetypes=[("Image files", "*.heic;*.jpg;*.jpeg;*.png;*.bmp;*.gif")])
    for file in files:
        size = os.path.getsize(file) / (1024 * 1024)  # size in MB
        filetree.insert('', 'end', values=(os.path.basename(file), f"{size:.2f} MB", file, ""))

def delete_selected_files():
    selected_items = filetree.selection()
    for item in selected_items:
        filetree.delete(item)

def clear_all_files():
    for item in filetree.get_children():
        filetree.delete(item)

def sort_files_by_name():
    files = [(filetree.item(iid)['values'][0], iid) for iid in filetree.get_children()]
    files.sort()
    for name, iid in files:
        filetree.move(iid, '', 'end')

def sort_files_by_size():
    files = [(filetree.item(iid)['values'][1], iid) for iid in filetree.get_children()]
    files.sort(key=lambda x: float(x[0].split()[0]))
    for size, iid in files:
        filetree.move(iid, '', 'end')

def toggle_debug_mode():
    global debug_mode
    debug_mode = not debug_mode
    debug_status = "enabled" if debug_mode else "disabled"
    messagebox.showinfo("Debug Mode", f"Debug mode {debug_status}.")

def show_about():
    messagebox.showinfo("About", "Batch HEIC to JPEG Converter\nVersion 1.0\nCreated by Aaron")

# Create the GUI
root = tk.Tk()
root.title("Batch HEIC to JPEG Converter")

# Initialize the BooleanVar after creating the root window
transfer_other_files = tk.BooleanVar(value=False)

# Create the menu
menu = Menu(root)
root.config(menu=menu)

file_menu = Menu(menu, tearoff=0)
menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Set ImageMagick Path", command=set_magick_path)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

help_menu = Menu(menu, tearoff=0)
menu.add_cascade(label="Help", menu=help_menu)
help_menu.add_command(label="About", command=show_about)
help_menu.add_command(label="Toggle Debug Mode", command=toggle_debug_mode)

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

control_frame = tk.Frame(frame)
control_frame.pack(pady=5)

add_button = tk.Button(control_frame, text="Add Image Files", command=add_files)
add_button.pack(side=tk.LEFT, padx=5)

delete_button = tk.Button(control_frame, text="Delete Selected Files", command=delete_selected_files)
delete_button.pack(side=tk.LEFT, padx=5)

clear_button = tk.Button(control_frame, text="Clear All Files", command=clear_all_files)
clear_button.pack(side=tk.LEFT, padx=5)

start_button = tk.Button(control_frame, text="Start Conversion", command=start_conversion)
start_button.pack(side=tk.LEFT, padx=5)

stop_button = tk.Button(control_frame, text="Stop Conversion", command=stop_conversion)
stop_button.pack(side=tk.LEFT, padx=5)

sort_frame = tk.Frame(frame)
sort_frame.pack(pady=5)

sort_name_button = tk.Button(sort_frame, text="Sort by Name", command=sort_files_by_name)
sort_name_button.pack(side=tk.LEFT, padx=5)

sort_size_button = tk.Button(sort_frame, text="Sort by Size", command=sort_files_by_size)
sort_size_button.pack(side=tk.LEFT, padx=5)

output_dir_frame = tk.Frame(frame)
output_dir_frame.pack(pady=5)

output_dir_label = tk.Label(output_dir_frame, text="Output Directory:")
output_dir_label.pack(side=tk.LEFT)

output_dir_entry = tk.Entry(output_dir_frame, width=50)
output_dir_entry.pack(side=tk.LEFT, padx=5)

browse_output_button = tk.Button(output_dir_frame, text="Browse", command=set_output_directory)
browse_output_button.pack(side=tk.LEFT, padx=5)

filetree = ttk.Treeview(frame, columns=("Name", "Size", "Path", "Status"), show="headings")
filetree.heading("Name", text="Name")
filetree.heading("Size", text="Size")
filetree.heading("Path", text="File Directory")
filetree.heading("Status", text="Status")
filetree.pack(pady=5)

other_files_checkbox = tk.Checkbutton(frame, text="Transfer other image files", variable=transfer_other_files)
other_files_checkbox.pack(pady=5)

# Load ImageMagick path from registry if available
load_magick_path_from_registry()

root.mainloop()

