import wmi
import time
import logging

# Configure logging
logging.basicConfig(filename='service_monitor.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def is_risky_service(service):
    """
    Check if the service is risky.
    Criteria for risk can be:
    - Suspicious service name or path
    - Service in a non-running state when it should be running
    - Services with unexpected startup types
    """
    risky_names = ["SuspiciousService", "MalwareService", "UnknownService"]
    risky_paths = ["C:\\SuspiciousPath", "C:\\MalwarePath"]
    non_running_states = ["Stopped", "Paused"]
    unexpected_start_modes = ["Auto", "Manual", "Disabled"]

    try:
        if service.Name in risky_names:
            return True
        if service.PathName and any(risk_path in service.PathName for risk_path in risky_paths):
            return True
        if service.State in non_running_states:
            return True
        if service.StartMode not in unexpected_start_modes:
            return True
    except AttributeError as e:
        logging.error(f"AttributeError when checking service {service.Name}: {e}")
    except Exception as e:
        logging.error(f"Unexpected error when checking service {service.Name}: {e}")

    return False


def monitor_services():
    try:
        conn = wmi.WMI()
    except Exception as e:
        logging.critical(f"Failed to connect to WMI: {e}")
        return

    while True:
        logging.info("Scanning services...")
        try:
            services = conn.Win32_Service()
        except wmi.x_wmi as e:
            logging.error(f"Failed to query services: {e}")
            time.sleep(10)
            continue
        except Exception as e:
            logging.error(f"Unexpected error during service query: {e}")
            time.sleep(10)
            continue

        for service in services:
            try:
                if is_risky_service(service):
                    logging.warning(f"Risky Service Detected: {service.Name}")
                    logging.warning(f"Service Display Name: {service.DisplayName}")
                    logging.warning(f"Process ID: {service.ProcessId}")
                    logging.warning(f"Description: {service.Description}")
                    logging.warning(f"Executable Path: {service.PathName}")
                    logging.warning(f"State: {service.State}")
                    logging.warning(f"Start Mode: {service.StartMode}")
                    logging.warning(f"Install Date: {service.InstallDate}")
                    logging.warning(f"Service Type: {service.ServiceType}")
                    logging.warning("----------")
            except wmi.x_wmi as e:
                logging.error(f"WMI error processing service {service.Name}: {e}")
            except AttributeError as e:
                logging.error(f"AttributeError processing service {service.Name}: {e}")
            except Exception as e:
                logging.error(f"Unexpected error processing service {service.Name}: {e}")

        time.sleep(10)


if __name__ == "__main__":
    monitor_services()
