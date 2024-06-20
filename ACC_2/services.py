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


def get_services(conn):
    services_data = []
    for service in conn.Win32_Service():
        install_date = convert_to_datetime(service.InstallDate)
        install_date_str = install_date.strftime('%Y-%m-%d %H:%M:%S') if install_date else "Not Available"
        service_info = {
            'Name': service.Name,
            'DisplayName': service.DisplayName,
            'Description': service.Description,
            'PathName': service.PathName,
            'ServiceType': service.ServiceType,
            'Started': service.Started,
            'StartMode': service.StartMode,
            'State': service.State,
            'Status': service.Status,
            'SystemCreationClassName': service.SystemCreationClassName,
            'SystemName': service.SystemName,
            'ProcessId': service.ProcessId,
            'InstallDate': install_date_str,
        }
        services_data.append(service_info)
    return services_data


def save_to_excel(services_data):
    df = pd.DataFrame(services_data)
    # Ensure the output path exists
    output_file = os.path.join(os.path.expanduser('~'), 'Desktop', 'services.xlsx')
    df.to_excel(output_file, index=False)


def main():
    conn = wmi.WMI()
    services_data = get_services(conn)
    save_to_excel(services_data)


if __name__ == "__main__":
    main()
