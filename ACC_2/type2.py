import wmi
import datetime
import pandas as pd
import os


def convert_to_datetime(wmi_date):
    """Convert WMI datetime string to a Python datetime object, handling cases where the date might be missing."""
    try:
        if wmi_date:
            return datetime.datetime.strptime(wmi_date.split('.')[0], '%Y%m%d%H%M%S')
    except Exception as e:
        return None
    return None


def get_system_drivers(conn):
    drivers_data = []
    for driver in conn.Win32_SystemDriver():
        install_date = convert_to_datetime(driver.InstallDate)
        install_date_str = install_date.strftime('%Y-%m-%d %H:%M:%S') if install_date else "Not Available"
        driver_info = {
            'Name': driver.Name,
            'Description': driver.Description,
            'DisplayName': driver.DisplayName,
            'ErrorControl': driver.ErrorControl,
            'InstallDate': install_date_str,
            'PathName': driver.PathName,
            'ServiceType': driver.ServiceType,
            'Started': driver.Started,
            'StartMode': driver.StartMode,
            'StartName': driver.StartName,
            'State': driver.State,
            'Status': driver.Status,
            'SystemCreationClassName': driver.SystemCreationClassName,
            'SystemName': driver.SystemName,
            'TagId': driver.TagId
        }
        drivers_data.append(driver_info)
    return drivers_data


def save_to_excel(drivers_data):
    df = pd.DataFrame(drivers_data)
    # Ensure the output path exists
    output_file = os.path.join(os.path.expanduser('~'), 'Desktop', 'type2_drivers.xlsx')
    df.to_excel(output_file, index=False)


def main():
    conn = wmi.WMI()
    drivers_data = get_system_drivers(conn)
    save_to_excel(drivers_data)


if __name__ == "__main__":
    main()
