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


def get_processes(conn):
    processes_data = []
    for process in conn.Win32_Process():
        creation_date = convert_to_datetime(process.CreationDate)
        creation_date_str = creation_date.strftime('%Y-%m-%d %H:%M:%S') if creation_date else "Not Available"

        process_info = {
            'CreationClassName': process.CreationClassName,
            'Caption': process.Caption,
            'CommandLine': process.CommandLine,
            'CreationDate': creation_date_str,
            'CSCreationClassName': process.CSCreationClassName,
            'CSName': process.CSName,
            'Description': process.Description,
            'ExecutablePath': process.ExecutablePath,
            'ExecutionState': process.ExecutionState,
            'Handle': process.Handle,
            'HandleCount': process.HandleCount,
            'InstallDate': "Not Available",  # Not available in Win32_Process
            'KernelModeTime': process.KernelModeTime,
            'MaximumWorkingSetSize': process.MaximumWorkingSetSize,
            'MinimumWorkingSetSize': process.MinimumWorkingSetSize,
            'Name': process.Name,
            'OSCreationClassName': process.OSCreationClassName,
            'OSName': process.OSName,
            'OtherOperationCount': process.OtherOperationCount,
            'OtherTransferCount': process.OtherTransferCount,
            'PageFaults': process.PageFaults,
            'PageFileUsage': process.PageFileUsage,
            'ParentProcessId': process.ParentProcessId,
            'PeakPageFileUsage': process.PeakPageFileUsage,
            'PeakVirtualSize': process.PeakVirtualSize,
            'PeakWorkingSetSize': process.PeakWorkingSetSize,
            'Priority': process.Priority,
            'PrivatePageCount': process.PrivatePageCount,
            'ProcessId': process.ProcessId,
            'QuotaNonPagedPoolUsage': process.QuotaNonPagedPoolUsage,
            'QuotaPagedPoolUsage': process.QuotaPagedPoolUsage,
            'QuotaPeakNonPagedPoolUsage': process.QuotaPeakNonPagedPoolUsage,
            'QuotaPeakPagedPoolUsage': process.QuotaPeakPagedPoolUsage,
            'ReadOperationCount': process.ReadOperationCount,
            'ReadTransferCount': process.ReadTransferCount,
            'SessionId': process.SessionId,
            'Status': process.Status,
            'TerminationDate': "Not Available",  # Not available in Win32_Process
            'ThreadCount': process.ThreadCount,
            'UserModeTime': process.UserModeTime,
            'VirtualSize': process.VirtualSize,
            'WindowsVersion': process.WindowsVersion,
            'WorkingSetSize': process.WorkingSetSize,
            'WriteOperationCount': process.WriteOperationCount,
            'WriteTransferCount': process.WriteTransferCount
        }
        processes_data.append(process_info)
    return processes_data


def save_to_excel(processes_data):
    df = pd.DataFrame(processes_data)
    # Ensure the output path exists
    output_file = os.path.join(os.path.expanduser('~'), 'Desktop', 'processes.xlsx')
    df.to_excel(output_file, index=False)


def main():
    conn = wmi.WMI()
    processes_data = get_processes(conn)
    save_to_excel(processes_data)


if __name__ == "__main__":
    main()
