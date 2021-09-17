import os
import time
import requests
import shutil
import inquirer

import win32com.client
from pathlib import Path

from download import download
from rich.console import Console


# -------------------------------------------------------------------
# -------------------------------------------------------------------

start = time.time() # Start Timer

console = Console(color_system="windows", force_terminal=True) # Console Class

console.print("""[cyan1]
    █████╗ ██╗   ██╗████████╗ ██████╗     ███████╗████████╗ █████╗ ██████╗ ████████╗
    ██╔══██╗██║   ██║╚══██╔══╝██╔═══██╗    ██╔════╝╚══██╔══╝██╔══██╗██╔══██╗╚══██╔══╝
    ███████║██║   ██║   ██║   ██║   ██║    ███████╗   ██║   ███████║██████╔╝   ██║   
    ██╔══██║██║   ██║   ██║   ██║   ██║    ╚════██║   ██║   ██╔══██║██╔══██╗   ██║   
    ██║  ██║╚██████╔╝   ██║   ╚██████╔╝    ███████║   ██║   ██║  ██║██║  ██║   ██║   
    ╚═╝  ╚═╝ ╚═════╝    ╚═╝    ╚═════╝     ╚══════╝   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝   
                            Developer: Ghoul#6066                                                        
[/cyan1]""")

console.print("[cyan1]I will now ask you a few questions so we can get started.[/cyan1]\n")

console.print("[cyan1]Before We Get Started I am Going To Check For Updates.[/cyan1]")

# -------------------------------------------------------------------
# -------------------------------------------------------------------

THIS_VERSION = "1.0.6"

request = requests.get("https://api.github.com/repos/xFGhoul/AutoMenuStart/releases/latest")
data = request.json()
latest = data['tag_name'][1:]
download_url = data['assets'][0]

if latest < THIS_VERSION:
    time.sleep(2)
    os.system('cls')
    console.print("""[cyan1]
    ███╗   ██╗███████╗██╗    ██╗    ██╗   ██╗██████╗ ██████╗  █████╗ ████████╗███████╗
    ████╗  ██║██╔════╝██║    ██║    ██║   ██║██╔══██╗██╔══██╗██╔══██╗╚══██╔══╝██╔════╝
    ██╔██╗ ██║█████╗  ██║ █╗ ██║    ██║   ██║██████╔╝██║  ██║███████║   ██║   █████╗  
    ██║╚██╗██║██╔══╝  ██║███╗██║    ██║   ██║██╔═══╝ ██║  ██║██╔══██║   ██║   ██╔══╝  
    ██║ ╚████║███████╗╚███╔███╔╝    ╚██████╔╝██║     ██████╔╝██║  ██║   ██║   ███████╗
    ╚═╝  ╚═══╝╚══════╝ ╚══╝╚══╝      ╚═════╝ ╚═╝     ╚═════╝ ╚═╝  ╚═╝   ╚═╝   ╚══════╝                                                                       
    [/cyan1]""")
    console.print(f'[cyan1]Latest Version: v{latest}[/cyan1]') 
    console.print(f"[cyan1]DOWNLOADING NEW UPDATE...[/cyan1]")
    dl_path = Path.home() / 'Downloads/AutoStartMenuDL' / f'v{latest}' / f'v{latest}'
    download(download_url['browser_download_url'], dl_path, progressbar=True, kind="zip", verbose=False)
    console.print(f"[cyan1]Updated Downloaded. SEE => {dl_path}[/cyan1]")
    console.print(f"[cyan1]Exiting Because New Version Is Available, Extract And Use The Newwer Version")
    time.sleep(5)
    raise SystemExit



# -------------------------------------------------------------------
# -------------------------------------------------------------------

console.print("[cyan1]FOR THE CHOICES, AN EXAMPLE WOULD BE:\n\nDrive: [red]C[/red]\nProgram Name: [red]discord(MUST MATCH EXACTLY)[/red][/cyan1]\n\n")

drive = console.input("[cyan1]Drive: [/cyan1]", markup=True)

program_name = console.input("[cyan1]Program Name: [/cyan1]", markup=True)

questions = [
  inquirer.List('folder',
                    message="Where Should I Search?",
                    choices=['Program Files', 'Program Files (x86)', 'AppData'],
                    ),
]
answer = inquirer.prompt(questions)

console.print(f"\n\n[cyan1]Good Job Answering!, I will now start to search for [red]{program_name}.exe[/red] in Drive [red]{drive.upper()}:/[/red]\n\n")

console.print(f"[red]NOTE: THIS PROCESS WILL SEARCH YOUR ENTIRE DRIVE(It Will Exlude Some Directories And File Types.), IT WILL NOT FINISH 'FAST', YOU WILL NEED TO BE PATIENT[/red]")

# -------------------------------------------------------------------
# -------------------------------------------------------------------


username = os.getlogin()

exclude = set(['Windows', 'Microsoft', 'Intel', 'Google', 'dotnet', 'ImageGlass', 'Inno Setup 6', 'Microsoft Visual Studio', 'Internet Explorer', 'Microsoft Visual Studio Tools for Unity', 'Microsoft SQL Server', 'Windows Portable Devices', 'Windows Media Player', 'Windows Multimedia Platform', 'Common Files', 'Microsoft Analysis Services', 'Microsoft Office', 'Microsoft SDKs', 'MSBuild', 'Windows NT', 'Windows Kits', 'Windows Photo Viewer', 'Windows Defender', 'WinRar', 'Windows Defender Advanced Threat Protection', 'Windows Security', 'WindowsApps', 'Microsoft.NET', 'K-Lite Codec Pack', 'Reference Assemblies'])

for root, dirs, files in os.walk(f"{drive.upper()}:/Users/{username}/{answer['folder']}" if answer["folder"] == "AppData" else f"{drive.upper()}:/{answer['folder']}" if answer["folder"] == 'Program Files' or 'Program Files (x86)' else None, onerror=None, topdown=True):

    [dirs.remove(d) for d in list(dirs) if d in exclude]

    if any(f.endswith(('.dll', '.txt', '.api', '.htm', '.html', '.py', '.js', '.ts', '.cmake', 'LICENSE', '.rst', '.ttf', '.xml', '.css', '.po', '.mo', '.pak', '.json', '.vdf', '.layout', '.pdf', '.mpp', '.jar', '.otf', '.md', '.pmp', '.png', '.gif', '.webp', '.jpg', '.maifest', '.TXT', '.tex', '.h', '.cpp', '.vbs', '.msg', '.lib', '.c', '.lua', '.mlua', '.wlua', '.dlua', '.mp4', '.mp3', '.so', '.dat', '.policy', '.properties', '.svg', '.conf', '.bin', '.frag', '.vert', '.bat', '.in', '.enc', '.tcl', '.sample', '.crt', '.pm', '.pl', '.code-snippets', '.env', '.rsh', '.bc', '.b', '.out', '.java')) for f in files):
        continue
    
    for file in files:
        s = file
        print(len(s) * " ", '\r', s, end='')

        if file.startswith(f"{program_name}"):

            found_file = (os.path.join(root, file))

            console.print(f"\n\n[cyan1]After searching For [red]{program_name}.exe[/red], I found [red]{found_file}[/red]![/cyan1]\n\n")

            console.print("[cyan1]Alright!, I'm going to go ahead and get the wheels spinning...[/cyan1]\n\n")
  
            console.print("[cyan1]GETTING PROGRAMS MENU[/cyan1]")

            objShell = win32com.client.Dispatch("WScript.Shell")

            UserProgramsMenu = objShell.SpecialFolders("AllUsersPrograms")

            console.print("\n[cyan1]FETCH PROGRAMS MENU: COMPLETE[/cyan1]")

            console.print("\n[cyan1] Copying File...[/cyan1]")

            shutil.copy(found_file, UserProgramsMenu)

            console.print("\n[cyan1]File Has Been Copied, Restart Your Computer And Watch The Magic Happen.")
            
            end = time.time()

            elasped_time = end - start

            console.print(f"\n[cyan1]This Process Took [red]{elasped_time}[/red] Seconds.")

            time.sleep(2)

            raise SystemExit

# -------------------------------------------------------------------
# -------------------------------------------------------------------
       

