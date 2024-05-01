import tkinter as tk
import msoffcrypto
import threading
from tkinter import filedialog, scrolledtext

def load_passwords():
    # Opens a file dialog to select the text file containing passwords
    file_path = filedialog.askopenfilename(title="Select Password File", filetypes=[("Text files", "*.txt")])
    if file_path:
        with open(file_path, "r") as file:
            # Loads passwords from the file, stripping any extra whitespace
            return [line.strip() for line in file if line.strip()]
    else:
        return []

def select_document():
    # Opens a file dialog to select the Word document
    file_path = filedialog.askopenfilename(title="Select Word Document", filetypes=[("Word files", "*.doc;*.docx")])
    return file_path

def try_passwords(doc_path, password_list):
    # Opens the Word document
    document = msoffcrypto.OfficeFile(open(doc_path, "rb"))
    for password in password_list:
        if stop_thread:
            return
        password_var.set(password)
        window.update_idletasks()  # Ensures the GUI updates the current password and logs
        log.insert(tk.END, f"Trying password: {password}\n")  # Logs the password attempt
        log.see(tk.END)  # Scrolls to the end of the log
        try:
            # Tries to unlock the document with each password
            document.load_key(password=password)
            result_var.set(f"Success! The password is: {password}")
            return
        except Exception as e:
            result_var.set("Testing...")
    result_var.set("Password not found in the provided list.")

def start_thread():
    global stop_thread
    stop_thread = False
    log.delete('1.0', tk.END)  # Clears the log at the start
    doc_path = select_document()
    if doc_path:
        passwords = load_passwords()
        if passwords:
            result_var.set("Starting test...")
            threading.Thread(target=try_passwords, args=(doc_path, passwords)).start()
        else:
            result_var.set("No passwords loaded or invalid file.")
    else:
        result_var.set("No document selected.")

def stop_process():
    global stop_thread
    stop_thread = True
    result_var.set("Process stopped by user.")

# GUI setup
window = tk.Tk()
window.title("Password Cracker")

password_var = tk.StringVar()
result_var = tk.StringVar()
stop_thread = False

tk.Label(window, text="Testing Password:").pack()
tk.Label(window, textvariable=password_var).pack()
tk.Label(window, textvariable=result_var).pack()

tk.Button(window, text="Load Files & Start", command=start_thread).pack(side=tk.LEFT)
tk.Button(window, text="Stop", command=stop_process).pack(side=tk.RIGHT)

# Scrolled text area for logging password attempts
log = scrolledtext.ScrolledText(window, height=10, width=50)
log.pack(pady=10)

window.mainloop()

