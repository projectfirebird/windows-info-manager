#!env/var/bin python
import platform
import psutil
from datetime import datetime
import cpuinfo
import winreg
import GPUtil
import wmi
import subprocess
import re
from tabulate import tabulate


def main_menu():
    print("Main Menu:\n")
    print("1. System & Operating System Info")
    print("2. CPU Info")
    print("3. GPU Info")
    print("4. Memory Info")
    print("5. Disk Info")
    print("6. List Installed Applications")
    print("7. List Installed Windows Updates")
    print("8. Exit Program\n")


def get_size(bytes, suffix="B"):
    """
    Scale bytes to its proper format
    e.g:
        1253656 => '1.20MB'
        1253656678 => '1.17GB'
    """
    factor = 1024
    for unit in ["", "K", "M", "G", "T", "P"]:
        if bytes < factor:
            return f"{bytes:.2f}{unit}{suffix}"
        bytes /= factor


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


def parse_installation_date(date_str):
    try:
        match = re.search(r'(\d{2}-[a-zA-Z]{3}-\d{2} \d{2}:\d{2}:\d{2} [APMapm]{2})', date_str)
        if match:
            return datetime.strptime(match.group(0), '%d-%b-%y %I:%M:%S %p')
        else:
            return None
    except ValueError:
        return None


def get_installed_updates():
    print("Checking for previously installed Windows updates...")

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

    print("Installed Windows Updates (ordered chronologically by Installation Date):")
    print(tabulate(table_data, headers=table_headers, tablefmt="grid"))


print("\t\tWelcome to Info Manager for Windows - v1.0")
print("\t\t\t\t by projectfirebird\n")
program_end = False

while not program_end:
    main_menu()
    menu_input = input("Select option: ")

    if menu_input == "1":
        os_edition = platform.win32_edition()
        uname = platform.uname()
        print("=" * 15, "System & OS Information", "=" * 15 + "\n")
        print(f"Operating System: {uname.system} {uname.release} {os_edition} - {uname.version}")
        os_information = get_os_information()
        for key, value in os_information.items():
            print(f"{key}: {value}")
        print("\n")
        print("=" * 55 + "\n")
    elif menu_input == "2":
        cpufreq = psutil.cpu_freq()
        uname = platform.uname()
        print("=" * 25, "CPU Info", "=" * 25 + "\n")
        print(f"Processor: {cpuinfo.get_cpu_info()['brand_raw']}")
        print(f"Processor: {uname.processor}")
        print("Physical Cores:", psutil.cpu_count(logical=False))
        print("Total Cores:", psutil.cpu_count(logical=True))
        print(f"Max. Frequency: {cpufreq.max:.2f}Mhz")
        print(f"CPU Usage: {psutil.cpu_percent()}%\n")
        print("=" * 60 + "\n")
    elif menu_input == "3":
        gpu_information = get_gpu_information()
        print("=" * 15, "GPU Info", "=" * 15 + "\n")
        for gpu, info in gpu_information.items():
            print(f"{gpu}:")
            for key, value in info.items():
                print(f"  {key}: {value}")
        print("\n")
        print("=" * 40 + "\n")
    elif menu_input == "4":
        print("=" * 10, "Memory Information", "=" * 10 + "\n")
        memory = psutil.virtual_memory()
        print(f"Total Memory: {get_size(memory.total)}")
        print(f"Available Memory: {get_size(memory.available)}")
        print(f"Used Memory: {get_size(memory.used)}")
        print(f"Percentage Used: {memory.percent}%\n")
        print("=" * 10, "SWAP Memory", "=" * 10 + "\n")
        swap = psutil.swap_memory()
        print(f"Total Memory: {get_size(swap.total)}")
        print(f"Available Memory: {get_size(swap.free)}")
        print(f"Used Memory: {get_size(swap.used)}")
        print(f"Percentage Used: {swap.percent}%\n")
        print("=" * 40 + "\n")
    elif menu_input == "5":
        print("=" * 10, "Disk Information", "=" * 10 + "\n")
        print("Partitions and Usage:")
        partitions = psutil.disk_partitions()
        for partition in partitions:
            print(f"=== Disk: {partition.device} ===")
            print(f"  Partition: {partition.mountpoint}")
            print(f"  File System Type: {partition.fstype}")
            try:
                partition_usage = psutil.disk_usage(partition.mountpoint)
            except PermissionError:
                continue
            print(f"  Total Disk Size: {get_size(partition_usage.total)}")
            print(f"  Used Disk: {get_size(partition_usage.used)}")
            print(f"  Available Disk: {get_size(partition_usage.free)}")
            print(f"  Percentage Used: {partition_usage.percent}%\n")
            print("=" * 38 + "\n")
    elif menu_input == "6":
        print("=" * 15, "List of Installed Applications", "=" * 15 + "\n")
        software_list = (foo(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) + foo(winreg.HKEY_LOCAL_MACHINE,
                                                                                      winreg.KEY_WOW64_64KEY) +
                         foo(winreg.HKEY_CURRENT_USER, 0))

        for software in software_list:
            print('Name: %s, Version: %s, Publisher: %s' % (software['name'], software['version'],
                                                            software['publisher']))
        print('Number of installed applications: %s' % len(software_list))
        print("=" * 50 + "\n")
    elif menu_input == "7":
        installed_updates = get_installed_updates()
        display_installed_updates(installed_updates)
    elif menu_input == "8":
        print("Thank you for using Info Manager. :)")
        program_end = True
    elif type(menu_input) == int and int(menu_input) > 8:
        print("Please select a valid Menu option, from 1 to 8.\n")
    elif type(menu_input) != int:
        print("Please select a valid Menu option, from 1 to 8.\n")
