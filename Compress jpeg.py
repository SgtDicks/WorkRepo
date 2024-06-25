import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image

def compress_images(input_folder, output_folder, quality=20):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.jpg') or filename.lower().endswith('.jpeg'):
            img_path = os.path.join(input_folder, filename)
            img = Image.open(img_path)
            
            output_path = os.path.join(output_folder, filename)
            img.save(output_path, "JPEG", quality=quality)
            print(f"Compressed {filename} and saved to {output_path}")

def select_input_folder():
    folder = filedialog.askdirectory()
    if folder:
        input_folder_var.set(folder)

def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_var.set(folder)

def start_compression():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    quality = quality_var.get()
    
    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both input and output folders.")
        return
    
    compress_images(input_folder, output_folder, quality)
    messagebox.showinfo("Success", "Images have been compressed successfully!")

# Set up the GUI
root = tk.Tk()
root.title("JPEG Compressor")

input_folder_var = tk.StringVar()
output_folder_var = tk.StringVar()
quality_var = tk.IntVar(value=20)

tk.Label(root, text="Input Folder:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_input_folder).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=output_folder_var, width=50).grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_output_folder).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Quality (1-100):").grid(row=2, column=0, padx=10, pady=10)
tk.Scale(root, from_=1, to=100, orient=tk.HORIZONTAL, variable=quality_var).grid(row=2, column=1, padx=10, pady=10)

tk.Button(root, text="Start Compression", command=start_compression).grid(row=3, column=0, columnspan=3, pady=20)

root.mainloop()

