import winreg
import wmi


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


def get_os_info():
    sub_key = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"
    os_info_keys = [
        "ProductName",
        "ReleaseId",
        "CurrentBuild",
        "CurrentVersion",
        "RegisteredOwner",
        "RegisteredOrganization"
    ]

    try:
        hkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key, 0, winreg.KEY_READ)
        os_info = {}
        for key in os_info_keys:
            try:
                value, _ = winreg.QueryValueEx(hkey, key)
                os_info[key] = value
            except FileNotFoundError:
                os_info[key] = "Not Found"
        winreg.CloseKey(hkey)
        return os_info
    except FileNotFoundError:
        return None


def get_firmware_info():
    try:
        c = wmi.WMI()
        bios_info = c.Win32_BIOS()[0]
        firmware_info = {
            "BIOS Version": bios_info.Version,
            "Manufacturer": bios_info.Manufacturer,
            "Release Date": bios_info.ReleaseDate,
            "SMBIOS Version": bios_info.SMBIOSBIOSVersion
        }
        return firmware_info
    except Exception as e:
        return {"Error": str(e)}


def main():
    os_info = get_os_info()
    if os_info:
        print("Operating System Information:")
        for key, value in os_info.items():
            print(f"{key}: {value}")
    else:
        print("Failed to retrieve OS information.")

    firmware_info = get_firmware_info()
    if firmware_info:
        print("\nFirmware Information:-")
        for key, value in firmware_info.items():
            print(f"{key}: {value}")
    else:
        print("Failed to retrieve firmware information.")

    enum_installed_soft("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "DisplayName")
    enum_installed_soft("Software\\Classes\\Installer\\Products", "ProductName")


if __name__ == "__main__":
    main()
