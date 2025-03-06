import os
import subprocess
import shutil
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import json
import re
import ctypes
import threading
import sys


def get_unc_path(local_path):
    print(f"Original path: {local_path}")

    drive = os.path.splitdrive(local_path)[0].strip().upper()

    try:     
        output = subprocess.check_output(
                    ['powershell', '-Command', 
                        'Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 4 } | Select-Object DeviceID, ProviderName'], 
                    text=True
                )
        lines = output.splitlines()
        for i, line in enumerate(lines):
            if line.strip().startswith(drive):
                match = re.search(r"(\\\\[^\s]+.*)", line)
                if not match and i + 1 < len(lines):
                    match = re.search(r"(\\\\[^\s]+)", lines[i + 1])
                if match:
                    unc_path = match.group(1).strip(" /").replace("\\", "/")  
                    converted_path = local_path.replace(drive, unc_path)
                    print(f"UNC path ditemukan. Menggunakan path {converted_path}")
                    return converted_path

        print(f"Drive {drive} tidak ditemukan di output `net use`.")
    except Exception as e:
        print(f"Error checking mapped drives: {e}")

    print("Path tidak dikenali sebagai mapped drive, menggunakan path asli.")
    return local_path

def browse_directory(entry_widget):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder_selected)

# Fungsi untuk menyimpan ke file
def save_to_file():
    project_name = project_name_entry.get().strip()
    area_name = area_name_entry.get().strip()
    directories = [get_unc_path(entry.get().strip()) for entry in dir_entries]

    if not project_name or not area_name or any(not d for d in directories):
        messagebox.showerror("Error", "You have to fill all entries")
        return

    data = {
        "project_name": project_name,
        "area_name": area_name,
        "directories": directories
    }
    output_file = "conf.json"
    try:
        with open(output_file, "w") as f:
            json.dump(data, f, indent=4)
        messagebox.showinfo("Success", f"Configuration has been saved on {output_file}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed while saving {e}")

def update_progress(value, message):
    progress_bar['value'] = value
    status_label.config(text=message)
    root.update_idletasks() 

def run_tasks():
    try:
        update_progress(25, "Saving configuration...")
        save_to_file()

        update_progress(50, "Initializing directories...")
        start_init()

        update_progress(75, "Creating executable, It might take a while, Do not close this program...")
        pyinstall()

        update_progress(100, "Completed!")
        messagebox.showinfo("Success", "All processes completed successfully!")
        messagebox.showinfo("Ownership", "TIMAS SUPLINDO - Developed by Auvi A")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        print(e)

    finally:
        save_button.config(state=tk.NORMAL)
        update_progress(0, "Idle")

def start_init():
    with open('conf.json', 'r') as file:
        data = json.load(file)

    project = data["project_name"]
    area = data["area_name"]
    engpath = f"{data['directories'][0]}/Monitoring Piping {area} ENG"
    qcpath = f"{data['directories'][1]}/Monitoring Piping {area} QC"
    ppcpath = f"{data['directories'][2]}/Monitoring Piping {area} PPC"
    mhpath = f"{data['directories'][3]}/Monitoring Piping {area} MATERIAL HANDLING"
    dashboard_path = f"{data['directories'][4]}/{project}/Monitoring Piping {area}"

    masterdata_path = f"{engpath}/Master Data"

    mto_data = f"{engpath}/MTO {area}.xlsx"
    mir_data = f"{mhpath}/MIR {area}.xlsx"
    engdata = f"{engpath}/Monitoring Piping {area} ENG.xlsx"
    qcdata = f"{qcpath}/Monitoring Piping {area} QC.xlsx"
    ppcdata = f"{ppcpath}/Monitoring Piping {area} PPC.xlsx"
    mhdata = f"{mhpath}/Monitoring Piping {area} PPC.xlsx"

    os.makedirs(f"{masterdata_path}/BACKUP")
    os.makedirs(f"{qcpath}")
    os.makedirs(f"{ppcpath}")
    os.makedirs(f"{mhpath}")
    os.makedirs(f"{dashboard_path}")

    FILE_ATTRIBUTE_HIDDEN = 0x02
    ctypes.windll.kernel32.SetFileAttributesW(masterdata_path, FILE_ATTRIBUTE_HIDDEN)

    shutil.copy("Template/Piping Template for ENG.xlsx", f"{engpath}/Monitoring Piping {area} ENG.xlsx")
    shutil.copy("Template/Piping Template for QC.xlsx", f"{qcpath}/Monitoring Piping {area} QC.xlsx")
    shutil.copy("Template/Piping Template for PPC.xlsx", f"{ppcpath}/Monitoring Piping {area} PPC.xlsx")

    shutil.copy("Template/Piping Template for MTO.xlsx", f"{mto_data}")
    shutil.copy("Template/Piping Template for MIR.xlsx", f"{mir_data}")

    shutil.copy("Template/masterdata template.xlsx", f"{masterdata_path}/masterdata.xlsx")

    shutil.copyfile("sync.py", f"{project}_{area}_Syncronization.py")
    shutil.copyfile("puller.py", f"{project}_{area}_Puller.py")


    placeholder_eng = "ENGINEER_PLACEHOLDER"
    placeholder_qc = "QAQC_PLACEHOLDER"
    placeholder_ppc = "PPC_PLACEHOLDER"
    placeholder_dashboard = "DASHBOARD_PLACEHOLDER"
    
    with open(f"{project}_{area}_Syncronization.py") as file :
        content = file.read()
    
    with open(f"{project}_{area}_Syncronization.py", 'w') as file:
        new_content = content.replace(placeholder_eng, f"{engdata}") \
                         .replace(placeholder_qc, f"{qcdata}") \
                         .replace(placeholder_ppc, f"{ppcdata}") \
                         .replace(placeholder_dashboard, f"{dashboard_path}") \
                         .replace("MASTERDATA_PLACEHOLDER", f"{masterdata_path}") \
                         .replace("MTO_PLACEHOLDER", f"{mto_data}") \
                         .replace("MIR_PLACEHOLDER", f"{mir_data}") \
        
        file.write(new_content)

    with open(f"{project}_{area}_Puller.py") as file :
        content = file.read()
    
    with open(f"{project}_{area}_Puller.py", 'w') as file:
        new_content = content.replace(placeholder_eng, f"{engdata}") \
                         .replace(placeholder_qc, f"{qcdata}") \
                         .replace(placeholder_ppc, f"{ppcdata}") \
                         .replace(placeholder_dashboard, f"{dashboard_path}") \
                         .replace("MTO_PLACEHOLDER", f"{mto_data}") \
                         .replace("MIR_PLACEHOLDER", f"{mir_data}") \
        
        file.write(new_content)

    messagebox.showinfo("Success", f"Directory has been initiated successfully")

def pyinstall():
    with open('conf.json', 'r') as file:
        data = json.load(file)

    project = data["project_name"]
    area = data["area_name"]

    command = f"pyinstaller --onefile '{project}_{area}_Syncronization.py'"

    subprocess.run(['powershell', "-Command", command],
                   capture_output=True, text=True, check=True)
    
    command2 = f"pyinstaller --onefile '{project}_{area}_Puller.py'"

    subprocess.run(['powershell', "-Command", command2],
                   capture_output=True, text=True, check=True)
    
    shutil.rmtree("./build", ignore_errors=True)
    os.remove(f"{project}_{area}_Syncronization.spec")
    os.remove(f"{project}_{area}_Puller.spec")

def save_and_start():
    save_button.config(state=tk.DISABLED)
    thread = threading.Thread(target=run_tasks)
    thread.start()

root = tk.Tk()
root.title("Piping Monitoring Initiator")

# root.iconbitmap("assets/timaslogo.ico")

progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=8, column=0, columnspan=3, pady=10)

status_label = tk.Label(root, text="Idle", anchor="w")
status_label.grid(row=10, column=0, columnspan=3, sticky="w", padx=10)

save_button = tk.Button(root, text="Save", command=save_and_start, bg="green", fg="white")
save_button.grid(row=9, column=0, columnspan=3, pady=10)

tk.Label(root, text="Project Name:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
project_name_entry = tk.Entry(root, width=50)
project_name_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Area Name:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
area_name_entry = tk.Entry(root, width=50)
area_name_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="C A U T I O N !! \nIt highly reccomended using UNC PATH (//192.168.0.0/path/to/dir) \nrather than using mounted path (Z://path/to/dir). \nYou can copy paste the path to entries below!").grid(row=2, column=1, sticky="w", padx=10, pady=5)

dirlist = ['Engineer', 'QAQC', 'PPC', 'MATERIAL HANDLING', 'DASHBOARD']
dir_entries = []
for i, label in enumerate(dirlist):
    tk.Label(root, text=f"{label} Directory:").grid(row=3+i, column=0, sticky="w", padx=10, pady=5)
    dir_entry = tk.Entry(root, width=50)
    dir_entry.grid(row=3+i, column=1, padx=10, pady=5)
    tk.Button(root, text="Browse", command=lambda e=dir_entry: browse_directory(e)).grid(row=3+i, column=2, padx=10, pady=5)
    dir_entries.append(dir_entry)

tk.Button(root, text="Save", command=save_and_start, bg="green", fg="white").grid(row=9, column=0, columnspan=3, pady=10)

root.mainloop()

#EoL
# pyinstaller --onefile --name "Piping Monitoring Initiator" --noconfirm --distpath .\ --icon=".\assets\timaslogo.ico" --add-data ".\assets\timaslogo.ico;assets" --noconsole .\installment.py