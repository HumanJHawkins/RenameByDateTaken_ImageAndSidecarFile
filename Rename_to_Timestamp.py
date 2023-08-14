# File renamer to prefix existing file name with date/time stamp
#   Format is yyyymmdd_xMHHMM_ExistingFilename.ExistingExt
#   I.E. 20230626_PM0842_MyFile.png
#
# Format chosen to ensure sort by name is also by date. And,
#   previous file name distinguishes files within each minute. 

# Much of this is pieced together from snippits of code written 
#   by chatGPT. 

# Requires:
#   pip install tkinter
#   pip install exifread

# WARNING: Be very careful using this. It operates on every file 
#   within the folder(s) selected. And has undergone minimal testing 
#   at this point.
import os
import tkinter as tk
from datetime import datetime
import exifread
from tkinter import filedialog

def process_date_string(date_string):
    try:
        datetime_obj = datetime.strptime(date_string, '%Y:%m:%d %H:%M:%S')
        formatted_date = datetime_obj.strftime('%Y%m%d')
        formatted_time = datetime_obj.strftime('%I%M')
        am_pm = datetime_obj.strftime('%p')
        return f"{formatted_date}_{am_pm}{formatted_time}"
    except ValueError:
        print(f"Skipping file due to Value Error: date_string {date_string} not in expected format.")
        return None

def rename_files():
    folder_path = folder_path_entry.get()
    if not os.path.isdir(folder_path):
        result_text.set("Invalid folder path")
        return

    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if not os.path.isfile(file_path):
            print(f"Skipping '{file_name}': Not a file.")
            continue

        try:
            with open(file_path, 'rb') as f:
                tags = exifread.process_file(f)
        except:
            print(f"Skipping file '{file_name}': Error processing file.")
            continue

        date_taken = tags.get('EXIF DateTimeOriginal')
        if date_taken is None:
            print(f"Skipping file '{file_name}': 'Date taken' not found.")
            continue

        date_taken = process_date_string(str(date_taken))
        if date_taken is None:
            continue

        print(f"File: {file_name} - Date taken: {date_taken}")

        if file_name[:15] == date_taken:
            print(f"Skipping file '{file_name}': 'Date taken' already prefixed.")
            continue

        try:
            new_file_name = date_taken + "_" + file_name
            new_file_path = os.path.join(folder_path, new_file_name)
            os.rename(file_path, new_file_path)
            print(f"Renamed file '{file_name}' to '{new_file_name}'.")
        except Exception as e:
            print(f"Error in renaming '{file_name}': {e}")

    result_text.set("Renaming complete.")

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path = folder_path.replace("/", "\\")
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, folder_path)

window = tk.Tk()
window.title("Rename: Prefix with Date Taken")
window.geometry("575x200")

folder_path_label = tk.Label(window, text="Folder Path:")
folder_path_label.grid(row=0, column=0, padx=10, pady=10)

folder_path_entry = tk.Entry(window, width=50)
folder_path_entry.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(window, text="Browse...", command=browse_folder)
browse_button.grid(row=0, column=2, padx=10, pady=10)

result_text = tk.StringVar()
result_label = tk.Label(window, textvariable=result_text)
result_label.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

rename_button = tk.Button(window, text="Rename Files", command=rename_files, font='bold')
rename_button.grid(row=1, column=2, padx=10, pady=10, sticky='E')

window.bind('<Return>', lambda event=None: rename_button.invoke())

window.mainloop()