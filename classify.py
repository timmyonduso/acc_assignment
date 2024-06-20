import winreg
import wmi
import json
import requests
from bs4 import BeautifulSoup
import datetime
import os

# Load predefined lifecycle data
json_file_path = os.path.join(os.path.dirname(__file__), 'product_lifecycle.json')
try:
    with open(json_file_path, 'r') as f:
        lifecycle_data = json.load(f)
except FileNotFoundError:
    print(f"Error: The file {json_file_path} was not found.")
    lifecycle_data = {}


def enum_installed_soft(sub_key, sub_key_name):
    try:
        hkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key, 0, winreg.KEY_READ)
        index = 0
        software_list = []
        while True:
            try:
                key_name = winreg.EnumKey(hkey, index)
                sub_key_path = sub_key + "\\" + key_name
                hsubkey = winreg.OpenKey(hkey, key_name)
                sub_key_value = winreg.QueryValueEx(hsubkey, sub_key_name)[0]
                software_list.append(sub_key_value)
                winreg.CloseKey(hsubkey)
                index += 1
            except OSError:
                break
    except FileNotFoundError:
        return []
    finally:
        if hkey:
            winreg.CloseKey(hkey)
    return software_list

def classify_software(vendor):
    if vendor:
        vendor = vendor.lower()
        if "microsoft" in vendor:
            return "Operating System"
        elif "firmware" in vendor or "bios" in vendor:
            return "Firmware"
        else:
            return "Software"
    return "Unknown"


def determine_license_status(name, vendor):
    # This is a simplified determination based on common knowledge and assumptions
    if "open-source" in name.lower() or "free" in name.lower():
        return "Open-Source/Free"
    if vendor and "microsoft" in vendor.lower():
        return "Licensed"
    return "Unknown"


def get_installed_software():
    conn = wmi.WMI()
    software_list = []

    for software in conn.Win32_Product():
        item = {
            "Name": software.Name,
            "Vendor": software.Vendor,
            "Version": software.Version,
            "Description": software.Description,
            "InstallLocation": software.InstallLocation,
            "Classification": classify_software(software.Vendor),
            "License": determine_license_status(software.Name, software.Vendor),
            "InstallDate": software.InstallDate,
        }

        # Get lifecycle info from predefined data or dynamic sources
        lifecycle_info = lifecycle_data.get(software.Name, get_dynamic_product_info(software.Name))
        item.update(lifecycle_info)

        software_list.append(item)

    return software_list



def convert_to_datetime(wmi_date):
    try:
        if wmi_date:
            return datetime.datetime.strptime(wmi_date.split('.')[0], '%Y%m%d%H%M%S')
    except Exception as e:
        return None
    return None


def get_dynamic_product_info(product_name):
    url = f"https://support.microsoft.com/en-us/lifecycle/search?alpha=Office%202016{product_name.replace(' ', '-').lower()}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Try to find elements and get their text, use default values if not found
    end_of_support = soup.find('span', {'id': 'end-of-support'})
    end_of_life = soup.find('span', {'id': 'end-of-life'})
    update_frequency = soup.find('span', {'id': 'update-frequency'})

    return {
        "end_of_support": end_of_support.text if end_of_support else "Unknown",
        "end_of_life": end_of_life.text if end_of_life else "Unknown",
        "update_frequency": update_frequency.text if update_frequency else "Unknown"
    }


def main():

    software_list = get_installed_software()
    print("\nInstalled Software Information:")
    for software in software_list:
        print(json.dumps(software, indent=2))


if __name__ == "__main__":
    main()
