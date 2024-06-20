import os

import wmi
import datetime
import pandas as pd


def convert_to_datetime(wmi_date):
    """Convert WMI datetime string to a Python datetime object, handling cases where the date might be missing."""
    try:
        if wmi_date:
            return datetime.datetime.strptime(wmi_date.split('.')[0], '%Y%m%d%H%M%S')
    except Exception as e:
        return None
    return None


def get_pnp_signed_drivers(conn):
    drivers_data = []
    for driver in conn.Win32_PnPSignedDriver():
        install_date = convert_to_datetime(driver.InstallDate)
        install_date_str = install_date.strftime('%Y-%m-%d %H:%M:%S') if install_date else "Not Available"
        driver_info = {
            'ClassGuid': driver.ClassGuid,
            'CompatID': driver.CompatID,
            'Description': driver.Description,
            'DeviceClass': driver.DeviceClass,
            'DeviceID': driver.DeviceID,
            'DeviceName': driver.DeviceName,
            'DevLoader': driver.DevLoader,
            'DriverDate': driver.DriverDate,
            'DriverName': driver.DriverName,
            'DriverVersion': driver.DriverVersion,
            'FriendlyName': driver.FriendlyName,
            'HardWareID': driver.HardWareID,
            'InfName': driver.InfName,
            'InstallDate': install_date_str,
            'IsSigned': driver.IsSigned,
            'Location': driver.Location,
            'Manufacturer': driver.Manufacturer,
            'Name': driver.Name,
            'PDO': driver.PDO,
            'DriverProviderName': driver.DriverProviderName,
            'Signer': driver.Signer,
            'Started': driver.Started,
            'StartMode': driver.StartMode,
            'Status': driver.Status,
            'SystemCreationClassName': driver.SystemCreationClassName,
            'SystemName': driver.SystemName,
        }
        drivers_data.append(driver_info)
    return drivers_data


def save_to_excel(drivers_data):
    df = pd.DataFrame(drivers_data)
    # Change the output path to an existing directory
    output_file = os.path.join(os.path.expanduser('~'), 'Desktop', 'type1_drivers.xlsx')
    df.to_excel(output_file, index=False)


def main():
    conn = wmi.WMI()
    drivers_data = get_pnp_signed_drivers(conn)
    save_to_excel(drivers_data)


if __name__ == "__main__":
    main()
