import wmi
import time
import logging

# Configure logging
logging.basicConfig(filename='process_monitor.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def is_risky_process(process):
    """
    Check if the process is risky.
    Criteria for risk can be:
    - High handle count
    - Suspicious executable path or name
    - High resource usage (e.g., memory)
    """
    risky_names = ["suspicious.exe", "malware.exe", "unknown.exe"]
    risky_paths = ["C:\\SuspiciousPath", "C:\\MalwarePath"]
    handle_threshold = 1000  # Example threshold for handle count
    memory_threshold = 100 * 1024 * 1024  # Example threshold for memory usage (100 MB)

    try:
        if process.Name in risky_names:
            return True
        if process.ExecutablePath and any(risk_path in process.ExecutablePath for risk_path in risky_paths):
            return True
        if process.HandleCount > handle_threshold:
            return True
        if process.WorkingSetSize > memory_threshold:
            return True
    except AttributeError as e:
        logging.error(f"AttributeError when checking process {process.ProcessId}: {e}")
    except Exception as e:
        logging.error(f"Unexpected error when checking process {process.ProcessId}: {e}")

    return False


def monitor_processes():
    try:
        conn = wmi.WMI()
    except Exception as e:
        logging.critical(f"Failed to connect to WMI: {e}")
        return

    while True:
        logging.info("Scanning processes...")
        try:
            processes = conn.Win32_Process()
        except wmi.x_wmi as e:
            logging.error(f"Failed to query processes: {e}")
            time.sleep(10)
            continue
        except Exception as e:
            logging.error(f"Unexpected error during process query: {e}")
            time.sleep(10)
            continue

        for process in processes:
            try:
                if is_risky_process(process):
                    logging.warning(f"Risky Process Detected: {process.Name}")
                    logging.warning(f"Process ID: {process.ProcessId}")
                    logging.warning(f"Handle Count: {process.HandleCount}")
                    logging.warning(f"Executable Path: {process.ExecutablePath}")
                    logging.warning(f"Working Set Size: {process.WorkingSetSize}")
                    logging.warning(f"Parent Process ID: {process.ParentProcessId}")
                    logging.warning(f"Command Line: {process.CommandLine}")
                    logging.warning("----------")
            except wmi.x_wmi as e:
                logging.error(f"WMI error processing process {process.ProcessId}: {e}")
            except AttributeError as e:
                logging.error(f"AttributeError processing process {process.ProcessId}: {e}")
            except Exception as e:
                logging.error(f"Unexpected error processing process {process.ProcessId}: {e}")

        time.sleep(10)


if __name__ == "__main__":
    monitor_processes()
