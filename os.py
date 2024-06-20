import winreg
import platform
import subprocess


def enum_installed_soft(sub_key, sub_key_name):
    try:
        hkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key, 0, winreg.KEY_READ)
        index = 0
        while True:
            try:
                key_name = winreg.EnumKey(hkey, index)
                sub_key_path = sub_key + "\\" + key_name
                hsubkey = winreg.OpenKey(hkey, key_name)
                sub_key_value = winreg.QueryValueEx(hsubkey, sub_key_name)[0]
                print(f"{key_name} : {sub_key_value}")
                winreg.CloseKey(hsubkey)
                index += 1
            except OSError:
                break
    except FileNotFoundError:
        return False
    finally:
        if hkey:
            winreg.CloseKey(hkey)
    return True


def get_operating_system_info():
    os_info = {
        "System": platform.system(),
        "Node Name": platform.node(),
        "Release": platform.release(),
        "Version": platform.version(),
        "Machine": platform.machine(),
        "Processor": platform.processor()
    }
    return os_info


def get_firmware_info():
    firmware_info = {}
    try:
        result = subprocess.run(["wmic", "bios", "get", "Manufacturer,Name,ReleaseDate,Version"], capture_output=True,
                                text=True)
        output = result.stdout.split('\n')
        firmware_info = {
            "Manufacturer": output[1].strip(),
            "Name": output[2].strip(),
            "ReleaseDate": output[3].strip(),
            "Version": output[4].strip()
        }
    except Exception as e:
        print(f"Error retrieving firmware information: {e}")
    return firmware_info


def get_installed_software():
    def foo(hive, flag):
        aReg = winreg.ConnectRegistry(None, hive)
        aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0, winreg.KEY_READ | flag)
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
                try:
                    software['InstallDate'] = winreg.QueryValueEx(asubkey, "InstallDate")[0]
                except EnvironmentError:
                    software['InstallDate'] = 'undefined'
                software_list.append(software)
            except EnvironmentError:
                continue
        return software_list

    software_list = foo(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) + foo(winreg.HKEY_LOCAL_MACHINE,
                                                                                 winreg.KEY_WOW64_64KEY) + foo(
        winreg.HKEY_CURRENT_USER, 0)
    return software_list


def main():
    os_info = get_operating_system_info()
    print("Operating System Information:")
    for key, value in os_info.items():
        print(f"{key}: {value}")
    print()

    firmware_info = get_firmware_info()
    print("Firmware Information:")
    for key, value in firmware_info.items():
        print(f"{key}: {value}")
    print()

    software_list = get_installed_software()
    print("Installed Software Information:")
    for software in software_list:
        print(
            f"Name: {software['name']}, Version: {software['version']}, Publisher: {software['publisher']}, InstallDate: {software['InstallDate']}")
    print(f"Number of installed apps: {len(software_list)}")


if __name__ == "__main__":
    main()
