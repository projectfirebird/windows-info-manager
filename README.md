# Info Manager GUI - v2.0
## Description 

Python coded GUI that provides helpful information on a Windows running computer.

![test](https://github.com/projectfirebird/windows-info-manager/blob/v2.0/gui_main.png?raw=true)

### Output information

1. System & OS Info
  - Operating System (Windows version, edition and release number)
  - System Type (x64 or x86)
  - Installation Date and Time
  - Motherboard Manufacturer
  - Motherboard Model

2. CPU Info
  - Processor Name and Description
  - Number of Physical Cores
  - Total Number of Cores
  - Max. Frequency
  - Current CPU Usage (%)

3. GPU Info
  - GPU Name and Description
  - GPU Driver Version
  - Total Memory
  - Free Memory
  - Used Memory
  - Current GPU Usage (%)
  - Current GPU Temperature

4. Memory Info
  - Total Memory
  - Available Memory
  - Used Memory
  - Current Memory Usage (%)

5. Disk Info
  - Disk/Partition Name
  - File System Type
  - Total Disk Size
  - Used Disk
  - Available Disk
  - Current Disk Usage (%)

6. List Installed Applications
  - Lists all installed applications in a table, ordered alphabetically, providing their name, version and publisher. (Screenshot bellow)

7. List Installed Windows Updates
  - Lists all installed Windows Updates in a table, ordered chronologically, prividing the update name/id and the installation date.

8. Exit Program
  - Closes the GUI

9. About
  - Lists the current version and license information. (Screenshot bellow) 

## Requirements

1. Obviously, this code will only work/run on a Windows computer.
1. Python 3.x

Pythod modules needed:
- cpuinfo
- datetime
- GPUtil
- platform
- psutil
- re
- subprocess
- tabulate
- textwrap
- tkinter
- webbrowser
- winreg
- wmi

## Screenshots

1. Screenshot of "List Installed Applications"

![test](https://github.com/projectfirebird/windows-info-manager/blob/v2.0/installed_apps.png?raw=true)

2. Screenshot of "About"

![test](https://github.com/projectfirebird/windows-info-manager/blob/v2.0/gui_about.png?raw=true)

## I hope this script is useful and will work on updating it further. Cheers :)
