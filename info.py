import os
import getpass
import platform
import subprocess
import re
import ctypes

LOG_FILE = "laporan_sistem.txt"

# ------------------ Utility ------------------
def save_to_file(text):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")

def print_and_save(text):
    print(text)
    save_to_file(text)

# ------------------ SISTEM ------------------
def show_system_info():
    try:
        os_name = platform.system() + ' ' + platform.release()
        os_version = platform.version()
        current_user = getpass.getuser()

        users_raw = subprocess.run('net user', capture_output=True, text=True, shell=True)
        collect = False
        temp_line = ''
        for line in users_raw.stdout.splitlines():
            if '-----' in line:
                collect = not collect
                continue
            if collect:
                temp_line += ' ' + line.strip()

        # Filter hanya nama user valid, hapus kata-kata tidak relevan
        ignore_words = {'the','command','completed','successfully','nama','user','users'}
        users_list = [u for u in temp_line.split() if re.fullmatch(r'[A-Za-z0-9_-]+', u) and u.lower() not in ignore_words]

        info = f"\n=== INFORMASI SISTEM ===\nOS Name: {os_name}\nOS Version: {os_version}\nUser saat ini: {current_user}\nDaftar user yang ada: {', '.join(users_list)}\n"
        print_and_save(info)
    except Exception as e:
        print_and_save(f"Gagal mendapatkan informasi sistem: {e}\n")

# ------------------ PROCESSOR ------------------
def show_cpu_info():
    try:
        cpu_raw = subprocess.run('wmic cpu get Name,NumberOfCores,NumberOfLogicalProcessors,LoadPercentage /format:list', capture_output=True, text=True, shell=True)
        name = cores = logical = load = ''
        for line in cpu_raw.stdout.splitlines():
            if line.startswith('Name'):
                name = line.split('=')[1].strip()
            elif line.startswith('NumberOfCores'):
                cores = line.split('=')[1].strip()
            elif line.startswith('NumberOfLogicalProcessors'):
                logical = line.split('=')[1].strip()
            elif line.startswith('LoadPercentage'):
                load = line.split('=')[1].strip()

        info = f"\n=== INFORMASI PROCESSOR ===\nName: {name}\nCores: {cores}\nLogical Processors: {logical}\nLoad Percentage: {load}%\n"
        print_and_save(info)
    except Exception as e:
        print_and_save(f"Gagal mendapatkan informasi processor: {e}\n")

# ------------------ RAM ------------------
def show_ram_info():
    try:
        ram_raw = subprocess.run('wmic OS get TotalVisibleMemorySize,FreePhysicalMemory /format:list', capture_output=True, text=True, shell=True)
        total = free = 0
        for line in ram_raw.stdout.splitlines():
            if 'TotalVisibleMemorySize' in line:
                total = int(line.split('=')[1]) // 1024
            elif 'FreePhysicalMemory' in line:
                free = int(line.split('=')[1]) // 1024
        used = total - free
        percent = round(used / total * 100, 2) if total else 0

        info = f"\n=== INFORMASI RAM ===\nTotal RAM: {total} MB\nTerpakai: {used} MB ({percent}%)"

        ps_cmd = 'Get-CimInstance Win32_PhysicalMemory | Select-Object Manufacturer, Capacity, BankLabel | Format-List'
        output = subprocess.run(f'powershell -Command "{ps_cmd}"', capture_output=True, text=True, shell=True)
        info += "\n\n=== SLOT RAM TERPASANG ==="
        for block in output.stdout.strip().split('\n\n'):
            manufacturer = capacity = bank = ''
            for line in block.splitlines():
                if line.startswith('Manufacturer'):
                    manufacturer = line.split(':')[1].strip()
                elif line.startswith('Capacity'):
                    capacity = int(line.split(':')[1].strip()) // (1024**3)
                elif line.startswith('BankLabel'):
                    bank = line.split(':')[1].strip()
            if manufacturer:
                info += f"\nSlot: {bank} - Merk: {manufacturer} - Kapasitas: {capacity} GB"
        info += "\n"
        print_and_save(info)
    except Exception as e:
        print_and_save(f"Gagal mendapatkan informasi RAM: {e}\n")

# ------------------ HARD DISK ------------------
# Fungsi untuk mendapatkan free/total space sesuai Windows Explorer
def get_disk_space(path):
    free_bytes = ctypes.c_ulonglong(0)
    total_bytes = ctypes.c_ulonglong(0)
    avail_bytes = ctypes.c_ulonglong(0)
    ctypes.windll.kernel32.GetDiskFreeSpaceExW(
        ctypes.c_wchar_p(path),
        ctypes.byref(avail_bytes),
        ctypes.byref(total_bytes),
        ctypes.byref(free_bytes)
    )
    total_gb = total_bytes.value / (1024**3)
    free_gb = avail_bytes.value / (1024**3)
    used_gb = total_gb - free_gb
    return round(total_gb, 2), round(free_gb, 2), round(used_gb, 2)

def show_full_disk_info():
    try:
        # 1️⃣ Ambil info disk fisik
        ps_disk = 'Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size, SerialNumber | Format-List'
        disk_output = subprocess.run(f'powershell -Command "{ps_disk}"', capture_output=True, text=True, shell=True)

        disks = []
        for block in disk_output.stdout.strip().split('\n\n'):
            name = type_ = size = serial = ''
            for line in block.splitlines():
                if line.startswith('FriendlyName'):
                    name = line.split(':', 1)[1].strip()
                elif line.startswith('MediaType'):
                    type_ = line.split(':', 1)[1].strip()
                elif line.startswith('Size'):
                    size = round(int(line.split(':', 1)[1].strip()) / (1024**3), 2)
                elif line.startswith('SerialNumber'):
                    serial = line.split(':', 1)[1].strip()

            if type_ == 'Unspecified':
                if 'NVMe' in name:
                    type_ = 'NVMe SSD'
                    name = re.sub(r'NVMe SSD', '', name).strip()
                else:
                    type_ = 'Unknown'

            if name:
                disks.append({'name': name, 'type': type_, 'size': size, 'serial': serial})

        # 2️⃣ Ambil semua drive letter
        vol_raw = subprocess.run('wmic logicaldisk get Name /format:list', capture_output=True, text=True, shell=True)
        drives = []
        for line in vol_raw.stdout.splitlines():
            if 'Name' in line and '=' in line:
                drive_letter = line.split('=')[1].strip() + "\\"
                total, free, used = get_disk_space(drive_letter)
                # Map ke disk pertama (sederhana)
                disk_info = disks[0] if disks else {'name':'Unknown','type':'Unknown','serial':'Unknown'}
                drives.append({
                    'drive': drive_letter[:-1],
                    'total': total,
                    'free': free,
                    'used': used,
                    'disk_name': disk_info['name'],
                    'disk_type': disk_info['type'],
                    'disk_serial': disk_info['serial']
                })

        # 3️⃣ Bangun info tabel drive
        info = "\n=== INFO DRIVE & DISK ===\n"
        info += f"{'Drive':6} {'Total(GB)':10} {'Free(GB)':10} {'Used(GB)':10} {'Disk / Merk':20} {'Tipe':12} {'Serial'}\n"
        info += "-"*100 + "\n"

        total_user_capacity = 0
        for d in drives:
            total_user_capacity += d['total']
            info += f"{d['drive']:6} {d['total']:10.2f} {d['free']:10.2f} {d['used']:10.2f} {d['disk_name']:20} {d['disk_type']:12} {d['disk_serial']}\n"

        info += "-"*100 + "\n"
        info += f"Total Kapasitas User (semua drive letter): {total_user_capacity:.2f} GB\n"

        # 4️⃣ Total kapasitas fisik
        total_physical = sum(d['size'] for d in disks)
        info += f"Total Kapasitas Disk Fisik: {total_physical:.2f} GB\n"

        print_and_save(info)

    except Exception as e:
        print_and_save(f"Gagal mendapatkan informasi drive & disk: {e}\n")

# ------------------ VGA / GPU ------------------
def show_vga_info():
    try:
        vga_raw = subprocess.run('wmic path win32_videocontroller get Name /format:list', capture_output=True, text=True, shell=True)
        names = [line.split('=')[1].strip() for line in vga_raw.stdout.splitlines() if line.startswith('Name')]
        info = "\n=== INFORMASI VGA ==="
        info += '\n' + '\n'.join(names) if names else info + '\nTidak ditemukan'
        info += "\n"
        print_and_save(info)
    except Exception as e:
        print_and_save(f"Gagal mendapatkan informasi VGA: {e}\n")

# ------------------ MOTHERBOARD ------------------
def show_motherboard_info():
    try:
        ps_cmd = 'Get-WmiObject Win32_BaseBoard | Select-Object Manufacturer, Product, SerialNumber | Format-List'
        output = subprocess.run(f'powershell -Command "{ps_cmd}"', capture_output=True, text=True, shell=True)
        info = "\n=== INFORMASI MOTHERBOARD ==="
        for block in output.stdout.strip().split('\n\n'):
            manufacturer = product = serial = ''
            for line in block.splitlines():
                if line.startswith('Manufacturer'):
                    manufacturer = line.split(':')[1].strip()
                elif line.startswith('Product'):
                    product = line.split(':')[1].strip()
                elif line.startswith('SerialNumber'):
                    serial = line.split(':')[1].strip()
            if manufacturer:
                info += f"\nMerk: {manufacturer} - Tipe: {product} - Serial: {serial}"
        info += "\n"
        print_and_save(info)
    except Exception as e:
        print_and_save(f"Gagal mendapatkan informasi motherboard: {e}\n")

# ------------------ PRINTER ------------------
def show_printers():
    try:
        printer_raw = subprocess.run(
            'wmic printer get Name,Default,PrinterStatus,WorkOffline /format:list',
            capture_output=True, text=True, shell=True
        )

        info = "\n=== PRINTERS INSTALLED ==="
        blocks = [b for b in printer_raw.stdout.strip().split('\n\n') if b.strip()]

        for block in blocks:
            name = default = printer_status = work_offline = ''
            for line in block.splitlines():
                if line.startswith('Name='):
                    name = line.split('=', 1)[1].strip()
                elif line.startswith('Default='):
                    default = 'Yes' if line.split('=', 1)[1].strip() == 'TRUE' else 'No'
                elif line.startswith('PrinterStatus='):
                    printer_status = line.split('=', 1)[1].strip()
                elif line.startswith('WorkOffline='):
                    work_offline = line.split('=', 1)[1].strip()

            # Tentukan status yang lebih manusiawi
            if printer_status == '0':
                status = 'Ready'
            elif printer_status == '1':
                status = 'Paused'
            elif printer_status == '2':
                status = 'Error'
            elif printer_status == '3':
                status = 'Offline'
            elif printer_status == '4':
                status = 'Busy'
            else:
                status = 'Unknown'

            # Tambahkan info offline jika ada
            if work_offline == 'TRUE':
                status += ' (Offline)'

            if name:
                info += f"\nPrinter: {name} - Default: {default} - Status: {status}"

        info += "\n"
        print_and_save(info)

    except Exception as e:
        print_and_save(f"Failed to get printer information: {e}\n")

# ------------------ APLIKASI POPULER ------------------
def get_file_version(exe_path):
    """Ambil versi dari file executable Windows"""
    try:
        info_size = ctypes.windll.version.GetFileVersionInfoSizeW(exe_path, None)
        if not info_size:
            return None
        res = ctypes.create_string_buffer(info_size)
        ctypes.windll.version.GetFileVersionInfoW(exe_path, 0, info_size, res)
        lpdw = ctypes.c_uint()
        puLen = ctypes.c_uint()
        ctypes.windll.version.VerQueryValueW(res, '\\', ctypes.byref(ctypes.pointer(lpdw)), ctypes.byref(puLen))
        # fallback: tidak semua exe punya versi readable
        return None
    except Exception:
        return None

def find_executable_in_paths(exe_dict, extra_paths=None):
    """Cek folder Program Files, Program Files (x86), dan folder tambahan untuk executable tertentu"""
    found = {}
    paths = [os.environ.get("ProgramFiles", "C:\\Program Files"),
             os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")]
    if extra_paths:
        paths.extend(extra_paths)
    
    for base in paths:
        if not os.path.exists(base):
            continue
        for app_name, exe_name in exe_dict.items():
            for root, dirs, files in os.walk(base):
                for f in files:
                    if f.lower() == exe_name.lower():
                        exe_path = os.path.join(root, f)
                        ver = get_file_version(exe_path)
                        found[app_name] = ver if ver else None
                        break
    return found

def find_shortcut_in_start_menu(app_names):
    """Cek shortcut .lnk di Start Menu untuk aplikasi tertentu"""
    found = {}
    start_menu_paths = [
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
        os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs")
    ]
    for base in start_menu_paths:
        if not os.path.exists(base):
            continue
        for root, dirs, files in os.walk(base):
            for app_name in app_names:
                for f in files:
                    if f.lower().endswith(".lnk") and app_name.lower() in f.lower():
                        found[app_name] = None  # versi shortcut tidak tersedia
                        break
    return found

def show_apps_info():
    try:
        # 1️⃣ Ambil apps dari registry (untuk versi)
        ps_cmd = (
            'Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*,'
            'HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*,'
            'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* '
            '| Select-Object DisplayName, DisplayVersion'
        )
        apps_raw = subprocess.run(
            f'powershell -Command "{ps_cmd}"',
            capture_output=True, text=True, shell=True
        )

        apps_lines = [line.strip() for line in apps_raw.stdout.splitlines() if line.strip()]
        registry_apps = {}
        for line in apps_lines:
            if 'DisplayName' in line or line == '----':
                continue
            parts = line.split(None, 1)
            if len(parts) == 2:
                registry_apps[parts[0]] = parts[1]
            else:
                registry_apps[parts[0]] = None

        # 2️⃣ Cek Office Store apps (UWP)
        try:
            ps_office_store = 'Get-AppxPackage -Name "*Microsoft.Office*" | Select Name, Version'
            office_store_raw = subprocess.run(
                f'powershell -Command "{ps_office_store}"',
                capture_output=True, text=True, shell=True
            )
            office_store_apps = {}
            for line in office_store_raw.stdout.splitlines():
                if line.strip() and 'Name' not in line:
                    parts = line.split()
                    if len(parts) >= 2:
                        name, ver = parts[0], parts[-1]
                        office_store_apps[name] = ver
        except Exception:
            office_store_apps = {}

        # 3️⃣ Definisikan aplikasi populer
        zip_names = ['WinRAR', '7-Zip']
        office_names = ['Word', 'Excel', 'PowerPoint', 'OneNote', 'Outlook']
        pdf_executables = {
            "Adobe Acrobat": "Acrobat.exe",
            "Foxit Reader": "FoxitReader.exe",
            "SumatraPDF": "SumatraPDF.exe",
            "PDFsam Basic": "PDFsam.exe"
        }
        user_programs = [os.path.expandvars(r"%LOCALAPPDATA%\Programs")]

        # 4️⃣ ZIP / RAR
        zip_apps = {}
        for name in zip_names:
            ver = registry_apps.get(name)
            zip_apps[name] = ver
        if not any(zip_apps.values()):
            zip_apps.update(find_executable_in_paths({name: f"{name}.exe" for name in zip_names}))

        # 5️⃣ PDF Reader
        pdf_apps = {}
        for name in pdf_executables.keys():
            ver = registry_apps.get(name)
            if ver:
                pdf_apps[name] = ver
        pdf_apps.update(find_executable_in_paths(pdf_executables, extra_paths=user_programs))
        pdf_apps.update(find_shortcut_in_start_menu(list(pdf_executables.keys())))

        # 6️⃣ Microsoft Office
        office_apps = {}
        for name in office_names:
            ver = registry_apps.get(name)
            office_apps[name] = ver
        office_apps.update(office_store_apps)

        # 7️⃣ Format output dengan versi jika ada
        def format_app_versions(app_dict):
            result = []
            for app, ver in app_dict.items():
                if ver:
                    result.append(f"{app} {ver}")
                else:
                    result.append(app)
            return result

        info = "\n=== PENGECEKAN APLIKASI POPULER ==="
        info += f"\nZIP / RAR: {', '.join(format_app_versions(zip_apps)) if zip_apps else 'Tidak ditemukan'}"
        info += f"\nPDF Reader: {', '.join(format_app_versions(pdf_apps)) if pdf_apps else 'Tidak ditemukan'}"
        info += f"\nMicrosoft Office: {', '.join(format_app_versions(office_apps)) if office_apps else 'Tidak ditemukan'}"

        print_and_save(info)

    except Exception as e:
        print_and_save(f"Gagal mendapatkan informasi aplikasi: {e}\n")

# ------------------ MENU ------------------
def main():
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)

    menu = {
        '1': ('Informasi Sistem', show_system_info),
        '2': ('Informasi Processor', show_cpu_info),
        '3': ('Informasi RAM', show_ram_info),
        '4': ('Informasi Hard Disk', show_full_disk_info),
        '5': ('Informasi VGA', show_vga_info),
        '6': ('Aplikasi Populer', show_apps_info),
        '7': ('Motherboard', show_motherboard_info),
        '8': ('Printer', show_printers),
        '0': ('Keluar', None)
    }

    while True:
        print('\n=== MENU ===')
        for k, v in menu.items():
            print(f'{k}. {v[0]}')
        choice = input('Pilih menu: ').strip()
        if choice == '0':
            break
        elif choice in menu:
            menu[choice][1]()
        else:
            print('\nPilihan tidak valid!\n')

if __name__ == '__main__':
    main()
