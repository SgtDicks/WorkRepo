import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from openai import OpenAI, OpenAIError
import logging
import os
import winreg as reg

# Initialize logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Function to get the API key from the Windows registry
def get_api_key():
    try:
        reg_key = reg.OpenKey(reg.HKEY_CURRENT_USER, "Software\\OpenAITranscriber", 0, reg.KEY_READ)
        api_key, _ = reg.QueryValueEx(reg_key, "APIKey")
        reg.CloseKey(reg_key)
        return api_key
    except FileNotFoundError:
        logger.error("Registry key not found.")
        return None

# Function to generate speech and save to file
def generate_speech():
    input_text = text_input.get("1.0", tk.END).strip()
    if not input_text:
        logger.error("Input text is empty.")
        status_label.config(text="Input text is empty.")
        return
    
    output_path = filedialog.asksaveasfilename(
        defaultextension=".mp3",
        filetypes=[("MP3 files", "*.mp3"), ("All files", "*.*")]
    )
    if not output_path:
        logger.error("Output path not selected.")
        status_label.config(text="Output path not selected.")
        return
    
    speech_file_path = Path(output_path)

    try:
        # Ensure the directory exists
        os.makedirs(speech_file_path.parent, exist_ok=True)
        
        # Retrieve the API key
        api_key = get_api_key()
        if not api_key:
            status_label.config(text="API key not found. Please check the registry.")
            return

        # Initialize OpenAI client with the API key
        client = OpenAI(api_key=api_key)
        
        # Create speech using OpenAI API
        response = client.audio.speech.create(
            model="tts-1",
            voice="alloy",
            input=input_text
        )

        # Stream response to file
        response.stream_to_file(speech_file_path)
        
        logger.info(f"Speech file successfully saved to {speech_file_path}")
        status_label.config(text=f"Success! File saved to {speech_file_path}")
    except OpenAIError as e:
        logger.error(f"An error occurred with the OpenAI API: {e}")
        status_label.config(text="An error occurred with the OpenAI API. Check logs for details.")
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        status_label.config(text="An error occurred. Check logs for details.")

# Set up the GUI
root = tk.Tk()
root.title("Text to Speech Generator")

# Input text field
tk.Label(root, text="Enter text:").pack(pady=5)
text_input = tk.Text(root, height=10, width=50)
text_input.pack(pady=5)

# Generate button
generate_button = tk.Button(root, text="Generate Speech", command=generate_speech)
generate_button.pack(pady=10)

# Status label
status_label = tk.Label(root, text="")
status_label.pack(pady=5)

# Start the GUI event loop
root.mainloop()

