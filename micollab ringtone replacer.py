import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Button, Frame

def create_folder_if_not_exists(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

def copy_and_rename_file():
    # Define the source path for MiCollab
    app_data_path = os.getenv('APPDATA')
    micollab_path = os.path.join(app_data_path, 'Mitel Networks Corporation', 'MiCollab', 'custom')
    
    # Create the 'custom' folder if it does not exist
    create_folder_if_not_exists(micollab_path)
    
    # Open file dialog to choose an audio file
    file_path = filedialog.askopenfilename(filetypes=[("Audio Files", "*.wav *.mp3")])
    if not file_path:
        return
    
    # Get the file extension (.wav or .mp3)
    _, ext = os.path.splitext(file_path)
    if ext.lower() not in ['.wav', '.mp3']:
        messagebox.showerror("Error", "Please select a valid audio file (.wav or .mp3).")
        return
    
    # Set the new file name based on the extension
    new_file_name = 'ringtone' + ext
    new_file_path = os.path.join(micollab_path, new_file_name)
    
    # Copy and rename the selected file
    shutil.copy(file_path, new_file_path)
    
    # Confirmation message
    messagebox.showinfo("Success", f"File has been copied and renamed to {new_file_path}")

# Create the main window
root = tk.Tk()
root.title("MiCollab Ringtone Replacer")
root.configure(bg='#282828')

# Setup a frame for the buttons
frame = Frame(root, bg='#383838')
frame.pack(padx=10, pady=20)

# Info label
info_label = Label(frame, text="Please select an audio file to use as your ringtone.",
                   bg='#383838', fg='white', font=("Helvetica", 12))
info_label.pack(pady=(0, 10))

# Add a button to the frame
copy_button = Button(frame, text=" browse for a replacement ringtone", command=copy_and_rename_file,
                     bg='#505050', fg='white', font=("Helvetica", 14), padx=10, pady=5)
copy_button.pack()

# Style configuration
root.geometry("400x100")

# Run the application
root.mainloop()

