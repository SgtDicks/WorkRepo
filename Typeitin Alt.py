import tkinter as tk
from tkinter import messagebox, Menu, Toplevel, filedialog
from ttkbootstrap import Style
import ttkbootstrap as ttk
import pyautogui
import sqlite3
from pynput import keyboard, mouse
import threading
import time
import json
import webbrowser

# -----------------------------
# Database Setup
# -----------------------------

class MacroDatabase:
    def __init__(self, db_path='macros.db'):
        self.db_path = db_path
        self.initialize_database()

    def initialize_database(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS macros (
                    id INTEGER PRIMARY KEY,
                    description TEXT,
                    commands TEXT
                )
            ''')
            conn.commit()

    def fetch_macros(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, description, commands FROM macros')
            return cursor.fetchall()

    def add_macro(self, description, commands):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT INTO macros (description, commands) VALUES (?, ?)', (description, commands))
            conn.commit()

    def update_macro(self, macro_id, description, commands):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('UPDATE macros SET description = ?, commands = ? WHERE id = ?', (description, commands, macro_id))
            conn.commit()

    def delete_macro(self, macro_id):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM macros WHERE id = ?', (macro_id,))
            conn.commit()
        messagebox.showinfo("Deleted", "Macro deleted successfully!")

# -----------------------------
# Main Application Class
# -----------------------------

class TypeItInAlternative:
    def __init__(self, root):
        self.root = root
        self.style = Style(theme='litera')  # Using ttkbootstrap theme
        self.root.title("TypeItIn Alternative")
        self.root.geometry("400x500")

        # Initialize database
        self.db = MacroDatabase()

        # Initialize Undo and Redo stacks
        self.undo_stack = []
        self.redo_stack = []

        # Initialize recording attributes
        self.recorded_events = []
        self.recording = False
        self.last_event_time = None
        self.keyboard_listener = None
        self.mouse_listener = None

        # Initialize GUI Components
        self.create_menu()
        self.create_search_bar()
        self.create_delay_settings()
        self.create_buttons()
        self.create_macro_list()
        self.create_progress_bar()

        # Load macros
        self.load_macros()

    # -----------------------------
    # Menu Bar (Added File Menu with Export and Import)
    # -----------------------------

    def create_menu(self):
        menu_bar = Menu(self.root)
        self.root.config(menu=menu_bar)

        # File Menu
        file_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Export Macros", command=self.export_macros)
        file_menu.add_command(label="Import Macros", command=self.import_macros)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)

    # -----------------------------
    # Search Bar
    # -----------------------------

    def create_search_bar(self):
        search_frame = ttk.Frame(self.root)
        search_frame.pack(pady=5, fill="x", padx=10)

        ttk.Label(search_frame, text="Search Macros:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side="left", padx=5, fill="x", expand=True)
        search_entry.bind("<KeyRelease>", self.perform_search)

    def perform_search(self, event):
        query = self.search_var.get().lower()
        self.load_macros(query)

    # -----------------------------
    # Delay Settings
    # -----------------------------

    def create_delay_settings(self):
        self.delay_frame = ttk.Frame(self.root)
        self.delay_frame.pack(pady=10, fill="x", padx=10)

        # Initial Delay
        ttk.Label(self.delay_frame, text="Initial Delay (milliseconds):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.initial_delay_entry = ttk.Entry(self.delay_frame, width=15)
        self.initial_delay_entry.grid(row=0, column=1, padx=5, pady=5)
        self.initial_delay_entry.insert(0, "1000")  # Default 1 second

        # Delay Between Commands
        ttk.Label(self.delay_frame, text="Delay Between Commands (milliseconds):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.command_delay_entry = ttk.Entry(self.delay_frame, width=15)
        self.command_delay_entry.grid(row=1, column=1, padx=5, pady=5)
        self.command_delay_entry.insert(0, "500")  # Default 0.5 seconds

        # Post Macro Delay
        ttk.Label(self.delay_frame, text="Post Macro Delay (milliseconds):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.post_delay_entry = ttk.Entry(self.delay_frame, width=15)
        self.post_delay_entry.grid(row=2, column=1, padx=5, pady=5)
        self.post_delay_entry.insert(0, "1000")  # Default 1 second

        # Option to record delays
        self.record_delay_var = tk.BooleanVar(value=True)
        self.record_delay_checkbox = ttk.Checkbutton(
            self.delay_frame,
            text="Record Delays",
            variable=self.record_delay_var,
            bootstyle="info"
        )
        self.record_delay_checkbox.grid(row=3, column=0, columnspan=2, padx=5, pady=10, sticky="w")

    # -----------------------------
    # Control Buttons
    # -----------------------------

    def create_buttons(self):
        self.buttons_frame = ttk.Frame(self.root)
        self.buttons_frame.pack(pady=10)

        # Add Macro Button
        self.add_macro_button = ttk.Button(self.buttons_frame, text="Add New Macro", command=self.add_new_macro, bootstyle="success")
        self.add_macro_button.pack(side="left", padx=10)

        # Record Macro Buttons
        self.record_macro_button = ttk.Button(self.buttons_frame, text="Start Recording", command=self.toggle_recording, bootstyle="warning")
        self.record_macro_button.pack(side="left", padx=10)

        self.stop_recording_button = ttk.Button(self.buttons_frame, text="Stop Recording", command=self.toggle_recording, bootstyle="danger")
        self.stop_recording_button.pack(side="left", padx=10)
        self.stop_recording_button.pack_forget()

    # -----------------------------
    # Macro List with Search, Play, Edit, Delete
    # -----------------------------

    def create_macro_list(self):
        self.macro_frame = ttk.Frame(self.root)
        self.macro_frame.pack(pady=10, fill="both", expand=True, padx=10)

        # Canvas and Scrollbar for macros
        canvas = tk.Canvas(self.macro_frame)
        scrollbar = ttk.Scrollbar(self.macro_frame, orient="vertical", command=canvas.yview)
        self.macro_list_inner = ttk.Frame(canvas)

        self.macro_list_inner.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.macro_list_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def load_macros(self, query=None):
        # Clear existing macros
        for widget in self.macro_list_inner.winfo_children():
            widget.destroy()

        # Fetch macros from database
        macros = self.db.fetch_macros()

        for macro in macros:
            macro_id, description, commands = macro

            # Filter based on search query
            if query:
                if query not in description.lower() and query not in commands.lower():
                    continue

            macro_frame = ttk.Frame(self.macro_list_inner, borderwidth=1, relief="solid")
            macro_frame.pack(fill="x", padx=5, pady=5)

            # Play Button
            play_button = ttk.Button(macro_frame, text=description, bootstyle="outline-primary")
            play_button.config(command=lambda c=commands, b=play_button: self.run_macro(c, b))
            play_button.pack(side="left", padx=5, pady=5, fill="x", expand=True)

            # Edit Button
            edit_button = ttk.Button(macro_frame, text="Edit", command=lambda m_id=macro_id, desc=description, cmd=commands: self.open_editor(m_id, desc, cmd), bootstyle="secondary")
            edit_button.pack(side="left", padx=5)

            # Delete Button
            delete_button = ttk.Button(macro_frame, text="Delete", command=lambda m_id=macro_id: self.delete_macro(m_id), bootstyle="danger")
            delete_button.pack(side="left", padx=5)

    # -----------------------------
    # Progress Bar
    # -----------------------------

    def create_progress_bar(self):
        self.progress = ttk.Progressbar(self.root, orient='horizontal', mode='determinate')
        self.progress.pack(fill="x", padx=10, pady=5)

    def update_progress(self, executed, total):
        if total > 0:
            progress_percent = (executed / total) * 100
            self.progress['value'] = progress_percent
            self.root.update_idletasks()

    # -----------------------------
    # Add New Macro
    # -----------------------------

    def add_new_macro(self):
        self.open_editor()

    # -----------------------------
    # Macro Editor Window
    # -----------------------------

    def open_editor(self, macro_id=None, description="", commands="", script=""):
        editor_window = Toplevel(self.root)
        editor_window.title("Macro Editor")
        editor_window.geometry("800x600")

        # Description
        desc_label = ttk.Label(editor_window, text="Macro Description:")
        desc_label.pack(pady=5)

        desc_entry = ttk.Entry(editor_window, width=80)
        desc_entry.pack(pady=5)
        desc_entry.insert(0, description)

        # Commands
        commands_label = ttk.Label(editor_window, text="Macro Commands:")
        commands_label.pack(pady=5)

        commands_editor = tk.Text(editor_window, width=90, height=15)
        commands_editor.pack(pady=5)
        commands_editor.insert("1.0", commands)
        commands_editor.bind("<<Modified>>", lambda event: self.track_undo(commands_editor))
        self.current_editor = commands_editor

        # Scripting
        script_label = ttk.Label(editor_window, text="Macro Script:")
        script_label.pack(pady=5)

        script_editor = tk.Text(editor_window, width=90, height=10)
        script_editor.pack(pady=5)
        script_editor.insert("1.0", script)
        script_editor.bind("<<Modified>>", lambda event: self.track_undo(script_editor))
        self.current_script_editor = script_editor

        # Special Keys Menu
        menu_bar = Menu(editor_window)
        functions_menu = Menu(menu_bar, tearoff=0)
        special_keys = [
            "[Ctrl v]", "[Ctrl c]", "[Ctrl x]", "[TAB]", "[ENTER]", "[CAPSLOCK]", "[ESC]", "[SPACE]", "[BACKSPACE]", "[DELETE]",
            "[HOME]", "[END]", "[PAGEUP]", "[PAGEDOWN]", "[UP]", "[DOWN]", "[LEFT]", "[RIGHT]", "[F1]", "[F2]", "[F3]", "[F4]",
            "[F5]", "[F6]", "[F7]", "[F8]", "[F9]", "[F10]", "[F11]", "[F12]", "[CTRL]", "[ALT]", "[SHIFT]", "[SCROLLLOCK]",
            "[PRINTSCREEN]", "[INSERT]", "[PAUSE]", "[NUMLOCK]", "[WIN]",
            "[LOOP]", "[ENDLOOP]", "[SET]", "[IF]", "[ENDIF]", "[KEY_COMBO]"
        ]
        for key in special_keys:
            functions_menu.add_command(label=key, command=lambda k=key: self.insert_function(commands_editor, k))
        menu_bar.add_cascade(label="Functions", menu=functions_menu)
        editor_window.config(menu=menu_bar)

        # Buttons Frame
        buttons_frame = ttk.Frame(editor_window)
        buttons_frame.pack(pady=10)

        # Save Button
        save_button = ttk.Button(
            buttons_frame,
            text="Save",
            command=lambda: self.save_macro(
                macro_id,
                desc_entry.get(),
                commands_editor.get("1.0", tk.END).strip(),
                script_editor.get("1.0", tk.END).strip(),
                editor_window
            ),
            bootstyle="primary"
        )
        save_button.pack(side="left", padx=5)

        # Validate Button
        validate_button = ttk.Button(
            buttons_frame,
            text="Validate",
            command=lambda: self.perform_validation(commands_editor.get("1.0", tk.END).strip()),
            bootstyle="info"
        )
        validate_button.pack(side="left", padx=5)

        # Preview Button
        preview_button = ttk.Button(
            buttons_frame,
            text="Preview",
            command=lambda: self.preview_macro(commands_editor.get("1.0", tk.END).strip()),
            bootstyle="warning"
        )
        preview_button.pack(side="left", padx=5)

        # Undo and Redo Buttons
        undo_button = ttk.Button(
            buttons_frame,
            text="Undo",
            command=lambda: self.undo_action(commands_editor),
            bootstyle="secondary"
        )
        undo_button.pack(side="left", padx=5)

        redo_button = ttk.Button(
            buttons_frame,
            text="Redo",
            command=lambda: self.redo_action(commands_editor),
            bootstyle="secondary"
        )
        redo_button.pack(side="left", padx=5)

        # Help Button Removed

    def insert_function(self, editor, function_key):
        if function_key == "[LOOP]":
            editor.insert(tk.INSERT, "LOOP:")  # User inputs the count
        elif function_key == "[ENDLOOP]":
            editor.insert(tk.INSERT, "ENDLOOP ")
        elif function_key == "[SET]":
            editor.insert(tk.INSERT, "SET:")
        elif function_key == "[IF]":
            editor.insert(tk.INSERT, "IF:")
        elif function_key == "[ENDIF]":
            editor.insert(tk.INSERT, "ENDIF ")
        elif function_key == "[KEY_COMBO]":
            editor.insert(tk.INSERT, "KEY_COMBO:")
        else:
            editor.insert(tk.INSERT, f"{function_key} ")

    # -----------------------------
    # Save Macro
    # -----------------------------

    def save_macro(self, macro_id, description, commands, script, editor_window):
        if description and (commands or script):
            # Combine commands and script
            combined_commands = commands
            if script:
                combined_commands += f"\nSCRIPT:{script}"

            if macro_id:
                self.db.update_macro(macro_id, description, combined_commands)
            else:
                self.db.add_macro(description, combined_commands)
            messagebox.showinfo("Success", "Macro saved successfully!")
            editor_window.destroy()
            self.load_macros()
        else:
            messagebox.showwarning("Input Error", "Please enter a description and at least one command or script.")

    # -----------------------------
    # Validate Macro
    # -----------------------------

    def perform_validation(self, commands):
        is_valid, message = self.validate_macro(commands)
        if is_valid:
            messagebox.showinfo("Validation Success", message)
        else:
            messagebox.showerror("Validation Failed", message)

    def validate_macro(self, commands):
        tokens = commands.split()
        loop_stack = []
        condition_stack = []
        variables = set()

        for token in tokens:
            if token.startswith("LOOP:"):
                loop_stack.append("LOOP")
            elif token == "ENDLOOP":
                if not loop_stack:
                    return False, "Unmatched ENDLOOP found."
                loop_stack.pop()
            elif token.startswith("IF:"):
                condition_stack.append("IF")
            elif token == "ENDIF":
                if not condition_stack:
                    return False, "Unmatched ENDIF found."
                condition_stack.pop()
            elif token.startswith("SET:"):
                _, var_assignment = token.split(":")
                if '=' not in var_assignment:
                    return False, f"Invalid SET command syntax: {token}"
                var, _ = var_assignment.split("=")
                variables.add(var)
            # Additional validations can be added here

        if loop_stack:
            return False, "Unmatched LOOP found."
        if condition_stack:
            return False, "Unmatched IF found."
        return True, "Macro is valid."

    # -----------------------------
    # Preview Macro
    # -----------------------------

    def preview_macro(self, commands):
        tokens = commands.split("\n")
        preview_log = []
        variables = {}
        loop_stack = []
        condition_stack = []

        for token in tokens:
            if token.startswith("SET:"):
                _, var_assignment = token.split(":")
                var, value = var_assignment.split("=")
                variables[var] = value
                preview_log.append(f"Set variable '{var}' to '{value}'.")

            elif token.startswith("LOOP:"):
                count = int(token.split(":")[1])
                loop_stack.append({'count': count, 'current': 1, 'start': len(preview_log)})
                preview_log.append(f"Start loop {loop_stack[-1]['current']}/{count}.")

            elif token == "ENDLOOP":
                if loop_stack:
                    loop = loop_stack[-1]
                    if loop['current'] < loop['count']:
                        loop['current'] += 1
                        preview_log.append(f"Loop iteration {loop['current']}/{loop['count']}.")
                    else:
                        loop_stack.pop()
                        preview_log.append("End loop.")
                else:
                    preview_log.append("End loop (no matching loop).")

            elif token.startswith("IF:"):
                condition = token.split(":", 1)[1]
                var, operator, value = condition.split()
                var_value = variables.get(var)
                condition_result = False
                try:
                    var_value = float(var_value)
                    value = float(value)
                except:
                    pass  # Keep as string if not numeric

                if operator == ">":
                    condition_result = var_value > value
                elif operator == "<":
                    condition_result = var_value < value
                elif operator == "==":
                    condition_result = var_value == value
                elif operator == "!=":
                    condition_result = var_value != value
                elif operator == ">=":
                    condition_result = var_value >= value
                elif operator == "<=":
                    condition_result = var_value <= value
                condition_stack.append(condition_result)
                preview_log.append(f"IF condition '{condition}' is {'True' if condition_result else 'False'}.")

            elif token == "ENDIF":
                if condition_stack:
                    condition_stack.pop()
                    preview_log.append("ENDIF.")
                else:
                    preview_log.append("ENDIF (no matching IF).")

            elif token.startswith("KEY_COMBO:"):
                combo = token.split(":",1)[1]
                preview_log.append(f"Press key combination '{combo}'.")

            elif token.startswith("KEY:"):
                key = token.split(":",1)[1]
                preview_log.append(f"Press key '{key}'.")

            elif token.startswith("KEYUP:"):
                key = token.split(":",1)[1]
                preview_log.append(f"Release key '{key}'.")

            elif token.startswith("MOUSE_CLICK:"):
                coords = token.split(":",1)[1]
                preview_log.append(f"Click mouse at coordinates ({coords}).")

            elif token.startswith("DELAY:"):
                delay_ms = token.split(":",1)[1]
                preview_log.append(f"Delay for {delay_ms} milliseconds.")

            elif token.startswith("SCRIPT:"):
                script_content = token.split(":", 1)[1]
                preview_log.append(f"Execute script:\n{script_content}")

            elif token.startswith("VAR:"):
                var = token.split(":",1)[1]
                value = variables.get(var, "")
                preview_log.append(f"Write variable '{var}' with value '{value}'.")

            elif token.strip() == "":
                continue  # Skip empty lines

            else:
                preview_log.append(f"Write '{token}'.")

        # Display the preview in a new window
        preview_window = Toplevel(self.root)
        preview_window.title("Macro Preview")
        preview_window.geometry("600x400")

        text_widget = tk.Text(preview_window, wrap="word")
        text_widget.pack(expand=True, fill="both")
        for line in preview_log:
            text_widget.insert(tk.END, line + "\n")
        text_widget.config(state="disabled")

    # -----------------------------
    # Undo and Redo Functionality
    # -----------------------------

    def track_undo(self, editor):
        if editor.edit_modified():
            current_text = editor.get("1.0", tk.END)
            self.undo_stack.append(current_text)
            # Clear redo stack on new action
            self.redo_stack.clear()
            editor.edit_modified(False)

    def undo_action(self, editor):
        try:
            if self.undo_stack:
                last_action = self.undo_stack.pop()
                self.redo_stack.append(editor.get("1.0", tk.END))
                editor.delete("1.0", tk.END)
                editor.insert("1.0", last_action)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during undo:\n{e}")

    def redo_action(self, editor):
        try:
            if self.redo_stack:
                next_action = self.redo_stack.pop()
                self.undo_stack.append(editor.get("1.0", tk.END))
                editor.delete("1.0", tk.END)
                editor.insert("1.0", next_action)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during redo:\n{e}")

    # -----------------------------
    # Run Macro with All Features
    # -----------------------------

    def run_macro(self, commands, button):
        try:
            initial_delay_ms = float(self.initial_delay_entry.get())
            command_delay_ms = float(self.command_delay_entry.get())
            post_delay_ms = float(self.post_delay_entry.get())
        except ValueError:
            messagebox.showwarning("Input Error", "Please enter valid numbers for all delay settings.")
            return

        initial_delay = initial_delay_ms / 1000
        command_delay = command_delay_ms / 1000
        post_delay = post_delay_ms / 1000

        button.config(state="disabled")
        self.progress['value'] = 0
        total_commands = len(commands.split('\n'))
        executed_commands = 0

        def flash_button():
            if getattr(self, 'is_flashing', False):
                new_style = "danger" if not getattr(self, 'is_flashing_red', False) else "outline-primary"
                button.config(bootstyle=new_style)
                self.is_flashing_red = not getattr(self, 'is_flashing_red', False)
                button.after(500, flash_button)

        def execute_macro():
            variables = {}
            condition_stack = []
            loop_stack = []
            try:
                self.is_flashing = True
                self.is_flashing_red = False
                flash_button()

                time.sleep(initial_delay)

                tokens = commands.split("\n")
                executed_commands = 0

                i = 0
                while i < len(tokens):
                    token = tokens[i].strip()
                    if not token:
                        i +=1
                        continue

                    # Handle SET command
                    if token.startswith("SET:"):
                        _, var_assignment = token.split(":")
                        if '=' not in var_assignment:
                            messagebox.showerror("Error", f"Invalid SET command syntax: {token}")
                            break
                        var, value = var_assignment.split("=")
                        variables[var.strip()] = value.strip()
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        i +=1
                        continue

                    # Handle IF condition
                    if token.startswith("IF:"):
                        condition = token.split(":",1)[1].strip()
                        try:
                            var, operator, value = condition.split()
                            var_value = variables.get(var)
                            if var_value is None:
                                condition_result = False
                            else:
                                # Attempt to convert to float
                                try:
                                    var_value = float(var_value)
                                    value = float(value)
                                except:
                                    pass  # Keep as string

                                if operator == ">":
                                    condition_result = var_value > value
                                elif operator == "<":
                                    condition_result = var_value < value
                                elif operator == "==":
                                    condition_result = var_value == value
                                elif operator == "!=":
                                    condition_result = var_value != value
                                elif operator == ">=":
                                    condition_result = var_value >= value
                                elif operator == "<=":
                                    condition_result = var_value <= value
                                else:
                                    condition_result = False
                        except:
                            condition_result = False
                        condition_stack.append(condition_result)
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        i +=1
                        continue

                    if token == "ENDIF":
                        if condition_stack:
                            condition_stack.pop()
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        i +=1
                        continue

                    # If inside a false condition, skip commands until ENDIF
                    if False in condition_stack:
                        i +=1
                        continue

                    # Handle LOOP command
                    if token.startswith("LOOP:"):
                        try:
                            count = int(token.split(":")[1].strip())
                            loop_stack.append({'count': count, 'start': i +1})
                        except:
                            messagebox.showerror("Error", f"Invalid LOOP command syntax: {token}")
                            break
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        i +=1
                        continue

                    if token == "ENDLOOP":
                        if loop_stack:
                            loop = loop_stack[-1]
                            loop['count'] -=1
                            if loop['count'] >0:
                                i = loop['start']
                                continue
                            else:
                                loop_stack.pop()
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        i +=1
                        continue

                    # Handle KEY_COMBO
                    if token.startswith("KEY_COMBO:"):
                        combo = token.split(":",1)[1].strip()
                        keys = combo.split("+")
                        for key in keys:
                            pyautogui.keyDown(key.lower())
                        for key in reversed(keys):
                            pyautogui.keyUp(key.lower())
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        time.sleep(command_delay)
                        i +=1
                        continue

                    # Handle VAR substitution
                    if token.startswith("VAR:"):
                        var = token.split(":",1)[1].strip()
                        value = variables.get(var, "")
                        pyautogui.write(str(value))
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        time.sleep(command_delay)
                        i +=1
                        continue

                    # Handle SCRIPT
                    if token.startswith("SCRIPT:"):
                        script_content = token.split(":",1)[1].strip()
                        self.execute_macro_script(script_content)
                        executed_commands +=1
                        self.update_progress(executed_commands, total_commands)
                        i +=1
                        continue

                    # Handle other commands
                    self.execute_token(token, variables)
                    executed_commands +=1
                    self.update_progress(executed_commands, total_commands)
                    time.sleep(command_delay)
                    i +=1

                time.sleep(post_delay)
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while running the macro:\n{e}")
            finally:
                self.is_flashing = False
                button.config(bootstyle="outline-primary", state="normal")
                self.progress['value'] = 0

        def execute_token(self, token, variables):
            if token.startswith("MOUSE_CLICK:"):
                _, coords = token.split(":")
                x, y = map(int, coords.split(","))
                pyautogui.click(x, y)
            elif token.startswith("MOUSE_MOVE:"):
                _, coords = token.split(":")
                x, y = map(int, coords.split(","))
                pyautogui.moveTo(x, y)
            elif token.startswith("MOUSE_SCROLL:"):
                _, scroll = token.split(":")
                dx, dy = map(int, scroll.split(","))
                pyautogui.scroll(dy, x=dx, y=dy)
            elif token.startswith("KEY_COMBO:"):
                # Already handled in run_macro
                pass
            elif token.startswith("KEY:"):
                key = token.split(":",1)[1].strip()
                if key in ["CTRL", "ALT", "SHIFT"]:
                    pyautogui.keyDown(key.lower())
                elif key == "ENTER":
                    pyautogui.press("enter")
                elif key == "SPACE":
                    pyautogui.press("space")
                elif key == "BACKSPACE":
                    pyautogui.press("backspace")
                elif key == "TAB":
                    pyautogui.press("tab")
                else:
                    pyautogui.write(key)
            elif token.startswith("KEYUP:"):
                key = token.split(":",1)[1].strip()
                if key in ["CTRL", "ALT", "SHIFT"]:
                    pyautogui.keyUp(key.lower())
            elif token.startswith("DELAY:"):
                delay_time_ms = float(token.split(":",1)[1].strip())
                delay_time = delay_time_ms / 1000
                time.sleep(delay_time)
            elif token.startswith("SCRIPT:"):
                # Already handled in run_macro
                pass
            elif token.startswith("VAR:"):
                # Already handled in run_macro
                pass
            else:
                # Default to writing the key
                pyautogui.write(token)

        def execute_macro_script(self, script):
            # Define a safe execution environment
            safe_globals = {
                'pyautogui': pyautogui,
                'time': time,
                'self': self,
                # Add other allowed modules or functions here
            }
            safe_locals = {}
            try:
                exec(script, safe_globals, safe_locals)
            except Exception as e:
                messagebox.showerror("Script Error", f"An error occurred while executing the script:\n{e}")

        threading.Thread(target=self.run_macro, args=(None, None), daemon=True).start()

    # -----------------------------
    # Macro Recording
    # -----------------------------

    def toggle_recording(self):
        if not self.recording:
            self.start_recording()
        else:
            self.stop_recording()

    def start_recording(self):
        self.recorded_events = []
        self.recording = True
        self.last_event_time = time.time()

        # Start keyboard and mouse listeners
        self.keyboard_listener = keyboard.Listener(on_press=self.on_key_press, on_release=self.on_key_release)
        self.mouse_listener = mouse.Listener(on_click=self.on_click)

        self.keyboard_listener.start()
        self.mouse_listener.start()

        # Update UI buttons
        self.record_macro_button.pack_forget()
        self.stop_recording_button.pack(side="left", padx=10)

    def stop_recording(self):
        self.recording = False

        # Stop listeners
        if self.keyboard_listener:
            self.keyboard_listener.stop()
        if self.mouse_listener:
            self.mouse_listener.stop()

        # Update UI buttons
        self.stop_recording_button.pack_forget()
        self.record_macro_button.pack(side="left", padx=10)

        # Save recorded events as a new macro
        self.save_recorded_macro()

    def on_key_press(self, key):
        if self.recording:
            current_time = time.time()
            if self.record_delay_var.get() and self.last_event_time is not None:
                delay = current_time - self.last_event_time
                delay_ms = delay * 1000  # Convert to milliseconds
                self.recorded_events.append(f"DELAY:{delay_ms:.3f}")
            self.last_event_time = current_time

            try:
                if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                    self.recorded_events.append("KEY:CTRL")
                elif key in [keyboard.Key.alt_l, keyboard.Key.alt_r]:
                    self.recorded_events.append("KEY:ALT")
                elif key in [keyboard.Key.shift, keyboard.Key.shift_r]:
                    self.recorded_events.append("KEY:SHIFT")
                elif key == keyboard.Key.enter:
                    self.recorded_events.append("KEY:ENTER")
                elif key == keyboard.Key.space:
                    self.recorded_events.append("KEY:SPACE")
                elif key == keyboard.Key.backspace:
                    self.recorded_events.append("KEY:BACKSPACE")
                elif key == keyboard.Key.tab:
                    self.recorded_events.append("KEY:TAB")
                else:
                    if hasattr(key, 'char') and key.char:
                        self.recorded_events.append(f"KEY:{key.char}")
                    else:
                        self.recorded_events.append(f"KEY:{key.name.upper()}")
            except AttributeError:
                self.recorded_events.append(f"KEY:{key.name.upper()}")

    def on_key_release(self, key):
        if self.recording:
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                self.recorded_events.append("KEYUP:CTRL")
            elif key in [keyboard.Key.alt_l, keyboard.Key.alt_r]:
                self.recorded_events.append("KEYUP:ALT")
            elif key in [keyboard.Key.shift, keyboard.Key.shift_r]:
                self.recorded_events.append("KEYUP:SHIFT")

    def on_click(self, x, y, button, pressed):
        if self.recording and pressed:
            current_time = time.time()
            if self.record_delay_var.get() and self.last_event_time is not None:
                delay = current_time - self.last_event_time
                delay_ms = delay * 1000  # Convert to milliseconds
                self.recorded_events.append(f"DELAY:{delay_ms:.3f}")
            self.last_event_time = current_time

            self.recorded_events.append(f"MOUSE_CLICK:{x},{y}")

    def save_recorded_macro(self):
        if self.recorded_events:
            commands = "\n".join(self.recorded_events)
            self.open_editor(commands=commands)
        else:
            messagebox.showwarning("No Events", "No events were recorded.")

    # -----------------------------
    # Export and Import Macros
    # -----------------------------

    def export_macros(self):
        macros = self.db.fetch_macros()
        macros_data = []
        for macro in macros:
            macro_id, description, commands = macro
            macros_data.append({
                'id': macro_id,
                'description': description,
                'commands': commands
            })
        
        export_file = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if export_file:
            try:
                with open(export_file, 'w') as f:
                    json.dump(macros_data, f, indent=4)
                messagebox.showinfo("Export Successful", f"Macros exported to {export_file}")
            except Exception as e:
                messagebox.showerror("Export Failed", f"An error occurred while exporting macros:\n{e}")

    def import_macros(self):
        import_file = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if import_file:
            try:
                with open(import_file, 'r') as f:
                    macros_data = json.load(f)
                for macro in macros_data:
                    description = macro.get('description')
                    commands = macro.get('commands')
                    if description and commands:
                        self.db.add_macro(description, commands)
                messagebox.showinfo("Import Successful", f"Macros imported from {import_file}")
                self.load_macros()
            except Exception as e:
                messagebox.showerror("Import Failed", f"An error occurred while importing macros:\n{e}")

    # -----------------------------
    # Delete Macro
    # -----------------------------

    def delete_macro(self, macro_id):
        self.db.delete_macro(macro_id)
        self.load_macros()

    # -----------------------------
    # Main Loop
    # -----------------------------

    def on_closing(self):
        # Ensure all listeners are stopped before closing
        if self.recording:
            self.stop_recording()
        self.root.destroy()

if __name__ == "__main__":
    root = ttk.Window(themename="litera")
    app = TypeItInAlternative(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    # Removed the redundant call to create_macro_list()
    root.mainloop()

