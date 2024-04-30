import os
import logging
import tkinter as tk
from tkinter import filedialog
import MSWT_Edit
import ExtractSFX

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        file_var.set(file_path)
        update_filename_label()

def format_script():
    file_path = file_var.get()
    if file_path:
        try:
            MSWT_Edit.main(file_path)
            tk.messagebox.showinfo("Success", "Script formatted successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")
            # Log the error
            logging.error(f"An error occurred: {str(e)}")
            # Terminate the active process
            sys.exit(1)
    else:
        tk.messagebox.showerror("Error", "No file selected. Please select a file first.")

def extract_sfx():
    file_path = file_var.get()
    if file_path:
        try:
            ExtractSFX.main(file_path)
            tk.messagebox.showinfo("Success", "Sound Effects extracted successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")
            # Log the error
            logging.error(f"An error occurred: {str(e)}")
            # Terminate the active process
            sys.exit(1)
    else:
        tk.messagebox.showerror("Error", "No file selected. Please select a file first.")

def cancel():
    root.destroy()

def update_filename_label():
    file_path = file_var.get()
    if file_path:
        max_length = 50  # Maximum length of displayed file path
        if len(file_path) > max_length:
            truncated_path = "..." + file_path[-(max_length-3):]
        else:
            truncated_path = file_path
        filename_label.config(text="Selected File: " + truncated_path)

def open_readme():
    try:
        os.startfile("readme.txt")
    except FileNotFoundError:
        tk.messagebox.showerror("Error", "Readme file not found")
    except Exception as e:
        tk.messagebox.showerror("Error", str(e))
        # Log the error
        logging.error(f"An error occurred: {str(e)}")
        # Terminate the active process
        sys.exit(1)

root = tk.Tk()
root.title("Seven Seas Script Formatter v0.1")
root.resizable(False, False)

# Get the screen dimensions
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Set the window dimensions
window_width = 400
window_height = 600

# Calculate the position to center the window
x_position = int((screen_width - window_width) / 2)
y_position = int((screen_height - window_height) / 2)

# Set the window position
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

file_var = tk.StringVar()

# Create a label for user instructions
instructions_label = tk.Label(root, text="Select a Seven Seas (.DOCX format) script.\n\nClick \"Format Script for Lettering\" to create a new TXT file that has SFX and Seven Seas Script Markup removed, and text tags for Bold ([b][/b]), Italic ([i][/i]), and Bold Italic([k][/k]) applied to the script.\n\nClick \"Extract SFX\" to generate a TXT file with only Sound Effects and page labels.\n\nAll TXT files will be generated in the original .DOCX file's directory.", font=("Arial", 10), justify=tk.LEFT, wraplength=350)
instructions_label.pack(expand=True, anchor=tk.CENTER)

# Create a frame to contain the filename label
filename_frame = tk.Frame(root)
filename_frame.pack(expand=True, fill=tk.BOTH)

# Create the filename label and center it within the frame
filename_label = tk.Label(filename_frame, text="Selected File: None", justify=tk.LEFT, wraplength=350)
filename_label.pack(expand=True, anchor=tk.CENTER)

# Create a frame to contain the buttons
button_frame = tk.Frame(root)
button_frame.pack(expand=True, fill=tk.BOTH, anchor=tk.CENTER)

# Create the buttons and center them horizontally within the frame
open_button = tk.Button(button_frame, text="Open File", command=open_file, width=20, height=2, wraplength=80)
open_button.pack(side=tk.TOP, padx=5, pady=10, anchor=tk.CENTER)

format_button = tk.Button(button_frame, text="Format Script\nfor Lettering", command=format_script, width=20, height=2, wraplength=80)
format_button.pack(side=tk.TOP, padx=5, pady=10, anchor=tk.CENTER)

extract_button = tk.Button(button_frame, text="Extract\nSound Effects", command=extract_sfx, width=20, height=2, wraplength=80)
extract_button.pack(side=tk.TOP, padx=5, pady=10, anchor=tk.CENTER)

cancel_button = tk.Button(button_frame, text="Close", command=cancel, width=20, height=2)
cancel_button.pack(side=tk.TOP, padx=5, pady=10, anchor=tk.CENTER)

# Create the "Open Readme" button
readme_button = tk.Button(button_frame, text="Open Readme", command=open_readme)
readme_button.pack(side=tk.BOTTOM, padx=5, pady=10, anchor=tk.CENTER)

# Configure logging
logging.basicConfig(filename='error.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

root.mainloop()
