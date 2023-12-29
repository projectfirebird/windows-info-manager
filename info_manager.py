#!env/var/bin python
import cpuinfo
from datetime import datetime
import GPUtil
import platform
import psutil
import re
import subprocess
from tabulate import tabulate
import textwrap
import tkinter as tk
from tkinter import scrolledtext
import webbrowser
import winreg
import wmi


def get_size(bytes, suffix="B"):
    factor = 1024
    for unit in ["", "K", "M", "G", "T", "P"]:
        if bytes < factor:
            return f"{bytes:.2f}{unit}{suffix}"
        bytes /= factor


def parse_installation_date(date_str):
    try:
        match = re.search(r'(\d{2}-[a-zA-Z]{3}-\d{2} \d{2}:\d{2}:\d{2} [APMapm]{2})', date_str)
        if match:
            return datetime.strptime(match.group(0), '%d-%b-%y %I:%M:%S %p')
        else:
            return None
    except ValueError:
        return None


def foo(hive, flag):
    aReg = winreg.ConnectRegistry(None, hive)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                          0, winreg.KEY_READ | flag)
    count_subkey = winreg.QueryInfoKey(aKey)[0]
    software_list = []
    for i in range(count_subkey):
        software = {}
        try:
            asubkey_name = winreg.EnumKey(aKey, i)
            asubkey = winreg.OpenKey(aKey, asubkey_name)
            software['name'] = winreg.QueryValueEx(asubkey, "DisplayName")[0]
            try:
                software['version'] = winreg.QueryValueEx(asubkey, "DisplayVersion")[0]
            except EnvironmentError:
                software['version'] = 'undefined'
            try:
                software['publisher'] = winreg.QueryValueEx(asubkey, "Publisher")[0]
            except EnvironmentError:
                software['publisher'] = 'undefined'
            software_list.append(software)
        except EnvironmentError:
            continue
    return software_list


def get_installed_updates():
    try:
        if platform.system() != "Windows":
            print("This script is designed to run on a Windows system.")
            return []
        result = subprocess.run(['powershell', 'Get-HotFix | Format-Table -AutoSize'], capture_output=True,
                                text=True)
        update_lines = result.stdout.splitlines()
        if len(update_lines) < 4:
            print("No installed updates found.")
            return []
        updates = []
        for line in update_lines[3:]:
            data = line.split()
            if len(data) >= 6:
                description = ' '.join(data[1:-4])
                installed_on_str = ' '.join(data[-4:])
                description = description.replace('NT', '')
                installed_on_str = installed_on_str.replace('NT', '')
                if installed_on_str.startswith('AUTHORITY\\SYSTEM'):
                    date_str = ' '.join(data[-5:])
                    installed_on = parse_installation_date(date_str)
                else:
                    installed_on = parse_installation_date(installed_on_str)
                if installed_on is not None:
                    updates.append({'Description': description, 'InstalledOn': installed_on})
                else:
                    print(f"Failed to parse installation date: {installed_on_str}")
        return updates
    except Exception as e:
        print(f"Failed to check for installed updates: {e}")
        return []


def display_installed_updates(updates):
    sorted_updates = sorted(updates, key=lambda x: x.get('InstalledOn', datetime.min))

    table_headers = ["Description", "Installation Date"]
    table_data = [(update['Description'], update['InstalledOn'].strftime('%m/%d/%Y %I:%M:%S %p')) for update in
                  sorted_updates]

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)
    text.insert(tk.END, "Installed Windows Updates\n (ordered chronologically by Installation Date)\n\n")
    text.insert(tk.END, tabulate(table_data, headers=table_headers, tablefmt="grid"))
    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)
    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)
    text.config(state=tk.DISABLED)


def get_os_information():
    os_info = {}
    try:
        w = wmi.WMI()
        os_info['System Type'] = w.Win32_ComputerSystem()[0].SystemType
        install_date_str = w.Win32_OperatingSystem()[0].InstallDate
        install_date_str = install_date_str.split(".")[0]
        os_info['Installation Date'] = (datetime.strptime(install_date_str, '%Y%m%d%H%M%S').
                                        strftime('%Y-%m-%d %H:%M:%S'))
        os_info['Motherboard Manufacturer'] = w.Win32_ComputerSystem()[0].Manufacturer
        os_info['Motherboard Model'] = w.Win32_ComputerSystem()[0].Model
    except Exception as e:
        print(f"Error getting additional OS information: {e}")
    return os_info


def get_gpu_information():
    try:
        gpus = GPUtil.getGPUs()
        gpu_info = {}
        for i, gpu in enumerate(gpus):
            gpu_info[f'GPU {i + 1}'] = {
                'ID': gpu.id,
                'Name': gpu.name,
                'Driver Version': gpu.driver,
                'Memory Total': f"{round(gpu.memoryTotal, 2)} MB",
                'Memory Free': f"{round(gpu.memoryFree, 2)} MB",
                'Memory Used': f"{round(gpu.memoryUsed, 2)} MB",
                'GPU Usage': f"{gpu.load * 100}%",
                'Temperature': f"{gpu.temperature}°C",
            }
    except Exception as e:
        print(f"Error getting GPU information: {e}")
        gpu_info = {}
    return gpu_info


def display_os_info():
    os_edition = platform.win32_edition()
    uname = platform.uname()
    os_info_str = "=" * 16 + " System & OS Information " + "=" * 16 + "\n\n"
    os_info_str += f"Operating System: {uname.system} {uname.release} {os_edition} - {uname.version}\n"
    os_information = get_os_information()
    for key, value in os_information.items():
        os_info_str += f"{key}: {value}\n"
    os_info_str += "=" * 56 + "\n"

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)
    text.insert(tk.END, os_info_str)
    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)
    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)
    text.config(state=tk.DISABLED)


def display_cpu_info():
    cpufreq = psutil.cpu_freq()
    cpu_info_str = "=" * 23 + " CPU Info " + "=" * 23 + "\n\n"
    cpu_info_str += f"Processor: {cpuinfo.get_cpu_info()['brand_raw']}\n"
    cpu_info_str += f"Physical Cores: {psutil.cpu_count(logical=False)}\n"
    cpu_info_str += f"Total Cores: {psutil.cpu_count(logical=True)}\n"
    cpu_info_str += f"Max. Frequency: {cpufreq.max:.2f}Mhz\n"
    cpu_info_str += f"CPU Usage: {psutil.cpu_percent()}%\n"
    cpu_info_str += "=" * 56 + "\n"

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)
    text.insert(tk.END, cpu_info_str)
    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)
    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)
    text.config(state=tk.DISABLED)


def display_gpu_info():
    gpu_information = get_gpu_information()
    gpu_info_str = "=" * 23 + " GPU Info " + "=" * 23 + "\n\n"
    for gpu, info in gpu_information.items():
        gpu_info_str += f"{gpu}:\n"
        for key, value in info.items():
            gpu_info_str += f"  {key}: {value}\n"
    gpu_info_str += "=" * 56 + "\n"

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)
    text.insert(tk.END, gpu_info_str)
    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)
    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)
    text.config(state=tk.DISABLED)


def display_memory_info():
    memory_info_str = "=" * 18 + " Memory Information " + "=" * 18 + "\n\n"
    memory = psutil.virtual_memory()
    memory_info_str += f"Total Memory: {get_size(memory.total)}\n"
    memory_info_str += f"Available Memory: {get_size(memory.available)}\n"
    memory_info_str += f"Used Memory: {get_size(memory.used)}\n"
    memory_info_str += f"Percentage Used: {memory.percent}%\n\n"
    memory_info_str += "=" * 10 + " SWAP Memory " + "=" * 10 + "\n\n"
    swap = psutil.swap_memory()
    memory_info_str += f"Total Memory: {get_size(swap.total)}\n"
    memory_info_str += f"Available Memory: {get_size(swap.free)}\n"
    memory_info_str += f"Used Memory: {get_size(swap.used)}\n"
    memory_info_str += f"Percentage Used: {swap.percent}%\n"
    memory_info_str += "=" * 56 + "\n"

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)
    text.insert(tk.END, memory_info_str)
    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)
    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)
    text.config(state=tk.DISABLED)


def display_disk_info():
    disk_info_str = "=" * 18 + " Disk Information " + "=" * 18 + "\n\n"
    partitions = psutil.disk_partitions()
    for partition in partitions:
        disk_info_str += f"=== Disk: {partition.device} ===\n"
        disk_info_str += f"  Partition: {partition.mountpoint}\n"
        disk_info_str += f"  File System Type: {partition.fstype}\n"
        try:
            partition_usage = psutil.disk_usage(partition.mountpoint)
        except PermissionError:
            continue
        disk_info_str += f"  Total Disk Size: {get_size(partition_usage.total)}\n"
        disk_info_str += f"  Used Disk: {get_size(partition_usage.used)}\n"
        disk_info_str += f"  Available Disk: {get_size(partition_usage.free)}\n"
        disk_info_str += f"  Percentage Used: {partition_usage.percent}%\n\n"
    disk_info_str += "=" * 56 + "\n"

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)
    # Insert new content and center align
    text.insert(tk.END, disk_info_str)
    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)
    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)
    text.config(state=tk.DISABLED)


def list_installed_applications():
    software_list = (foo(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) +
                     foo(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY) +
                     foo(winreg.HKEY_CURRENT_USER, 0))

    software_list_sorted = sorted(software_list, key=lambda x: x['name'].lower())

    table_headers = ["Name", "Version", "Publisher"]
    table_data = [(software['name'], software['version'], software['publisher']) for software in software_list_sorted]

    # Define the width of the first column
    first_column_width = 30

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)

    text.insert(tk.END, "Installed Windows Applications\n (ordered alphabetically)\n\n ")

    # Wrap text in the first column
    wrapped_data = [(textwrap.fill(name, width=first_column_width), version, publisher) for name, version, publisher in
                    table_data]

    text.insert(tk.END, tabulate(wrapped_data, headers=table_headers, tablefmt="grid",
                                 colalign=("left", "center", "center"), numalign="center",
                                 showindex=False))

    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)
    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)
    text.config(state=tk.DISABLED)


def list_installed_windows_updates():
    installed_updates = get_installed_updates()
    display_installed_updates(installed_updates)


def open_owner_link(event):
    url = "https://github.com/projectfirebird"
    webbrowser.open_new(url)


def open_github_link(event):
    url = "https://github.com/projectfirebird/windows-info-manager"
    webbrowser.open_new(url)


def about_info():
    about_str = "=" * 18 + " About " + "=" * 18 + "\n\n"
    hyperlink_app = "Info Manager GUI - v2.0"
    about_str += hyperlink_app + "\n"
    about_str += "https://github.com/projectfirebird/windows-info-manager\n\n"

    about_str += "by"
    hyperlink_author = " projectfirebird"
    about_str += hyperlink_author + "\n"
    about_str += "https://github.com/projectfirebird\n\n"
    about_str += ("Info Manager GUI is a Python coded Graphical User Interface, designed to provide helpful information"
                  " about a computer running Windows.\n\n")
    about_str += "=" * 18 + " License " + "=" * 18 + "\n\n"
    about_str += "MIT License\n"
    about_str += "Copyright (c) 2023 projectfirebird\n\n"
    about_str += ('Permission is hereby granted, free of charge, to any person obtaining a copy of this software and'
                  ' associated documentation files (the "Software"), to deal in the Software without restriction,'
                  ' including without limitation the rights to use, copy, modify, merge, publish, distribute,'
                  ' sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is'
                  ' furnished to do so, subject to the following conditions:\n\n')
    about_str += ("The above copyright notice and this permission notice shall be included in all copies or substantial"
                  " portions of the Software.\n\n")
    about_str += ('THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT'
                  ' NOT'
                  ' LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.'
                  ' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER'
                  ' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN'
                  ' CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.\n')

    text.config(state=tk.NORMAL)
    text.delete(1.0, tk.END)
    text.insert(tk.END, about_str)

    text.tag_configure("hyperlink_app", foreground="blue", underline=True)
    start_index = text.search(hyperlink_app, 1.0, tk.END)
    end_index = f"{start_index}+{len(hyperlink_app)}c"
    text.tag_add("hyperlink_app", start_index, end_index)
    text.tag_bind("hyperlink_app", "<Button-1>", open_github_link)

    text.tag_configure("hyperlink_author", foreground="blue", underline=True)
    start_index = text.search(hyperlink_author, 1.0, tk.END)
    end_index = f"{start_index}+{len(hyperlink_author)}c"
    text.tag_add("hyperlink_author", start_index, end_index)
    text.tag_bind("hyperlink_author", "<Button-1>", open_owner_link)

    text.tag_configure("center", justify="center")
    text.tag_add("center", 1.0, tk.END)

    text.tag_configure("wrap", wrap=tk.WORD)
    text.tag_add("wrap", 1.0, tk.END)

    text.config(state=tk.DISABLED)


def exit_program():
    root.destroy()


root = tk.Tk()
root.iconbitmap("icon_v2.ico")
root.title("Info Manager GUI - v2.0")

root.rowconfigure(0, minsize=350, weight=1)
root.columnconfigure(1, minsize=750, weight=1)

text_frame = tk.Frame(root)
text = scrolledtext.ScrolledText(text_frame, height=20, width=80)
text.pack(expand=True, fill='both')
frm_buttons = tk.Frame(root, relief=tk.RAISED, bd=2)

text.config(state=tk.DISABLED)

button1 = tk.Button(frm_buttons, text="System & OS Info", command=display_os_info)
button2 = tk.Button(frm_buttons, text="CPU Info", command=display_cpu_info)
button3 = tk.Button(frm_buttons, text="GPU Info", command=display_gpu_info)
button4 = tk.Button(frm_buttons, text="Memory Info", command=display_memory_info)
button5 = tk.Button(frm_buttons, text="Disk Info", command=display_disk_info)
button6 = tk.Button(frm_buttons, text="List Installed Applications", command=list_installed_applications)
button7 = tk.Button(frm_buttons, text="List Installed Windows Updates", command=list_installed_windows_updates)
button8 = tk.Button(frm_buttons, text="Exit Program", command=exit_program)
button9 = tk.Button(frm_buttons, text="About", command=about_info)

button1.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
button2.grid(row=1, column=0, sticky="ew", padx=5)
button3.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
button4.grid(row=3, column=0, sticky="ew", padx=5)
button5.grid(row=4, column=0, sticky="ew", padx=5, pady=5)
button6.grid(row=5, column=0, sticky="ew", padx=5)
button7.grid(row=6, column=0, sticky="ew", padx=5, pady=5)
button8.grid(row=7, column=0, sticky="ew", padx=5)
button9.grid(row=8, column=0, sticky="ew", padx=5)

frm_buttons.grid(row=0, column=0, sticky="ns")
text_frame.grid(row=0, column=1, sticky="nsew")

root.mainloop()
