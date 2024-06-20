import winreg
import pandas as pd
import os


def get_installed_software(sub_key):
    software_list = []
    try:
        hkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key, 0, winreg.KEY_READ)
        index = 0
        while True:
            try:
                key_name = winreg.EnumKey(hkey, index)
                sub_key_path = sub_key + "\\" + key_name
                hsubkey = winreg.OpenKey(hkey, key_name)
                software = {col: None for col in columns}
                for name, col in columns.items():
                    try:
                        software[name] = winreg.QueryValueEx(hsubkey, col)[0]
                    except OSError:
                        software[name] = None
                software_list.append(software)
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


def main():
    global columns
    columns = {
        'AuthorizedCDFPrefix': 'AuthorizedCDFPrefix',
        'BundleAddonCode': 'BundleAddonCode',
        'BundleCachePath': 'BundleCachePath',
        'BundleDetectCode': 'BundleDetectCode',
        'BundlePatchCode': 'BundlePatchCode',
        'BundleProviderKey': 'BundleProviderKey',
        'BundleTag': 'BundleTag',
        'BundleUpgradeCode': 'BundleUpgradeCode',
        'BundleVersion': 'BundleVersion',
        'Comments': 'Comments',
        'Contact': 'Contact',
        'DisplayIcon': 'DisplayIcon',
        'DisplayName': 'DisplayName',
        'DisplayVersion': 'DisplayVersion',
        'EstimatedSize': 'EstimatedSize',
        'HelpLink': 'HelpLink',
        'HelpTelephone': 'HelpTelephone',
        'Inno Setup': 'Inno Setup',
        'Inno Setup CodeFile': 'Inno Setup CodeFile',
        'InstallDate': 'InstallDate',
        'Installed': 'Installed',
        'InstallLanguage': 'InstallLanguage',
        'InstallLocation': 'InstallLocation',
        'InstallSource': 'InstallSource',
        'InstallType': 'InstallType',
        'Language': 'Language',
        'MajorVersion': 'MajorVersion',
        'MinorVersion': 'MinorVersion',
        'ModifyPath': 'ModifyPath',
        'NoElevateOnModify': 'NoElevateOnModify',
        'NoModify': 'NoModify',
        'NoRemove': 'NoRemove',
        'NoRepair': 'NoRepair',
        'PSChildName': 'PSChildName',
        'PSDrive': 'PSDrive',
        'PSParentPath': 'PSParentPath',
        'PSPath': 'PSPath',
        'PSProvider': 'PSProvider',
        'Publisher': 'Publisher',
        'QuietUninstallString': 'QuietUninstallString',
        'Readme': 'Readme',
        'Size': 'Size',
        'SystemComponent': 'SystemComponent',
        'UninstallString': 'UninstallString',
        'URLInfoAbout': 'URLInfoAbout',
        'URLUpdateInfo': 'URLUpdateInfo',
        'Version': 'Version',
        'VersionBuild': 'VersionBuild',
        'VersionMajor': 'VersionMajor',
        'VersionTimestamp': 'VersionTimestamp',
        'WindowsInstaller': 'WindowsInstaller'
    }

    software_data = []
    software_data.extend(get_installed_software("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"))
    software_data.extend(get_installed_software("SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"))

    df = pd.DataFrame(software_data, columns=columns.keys())
    # Change the output path to an existing directory
    output_file = os.path.join(os.path.expanduser('~'), 'Desktop', 'installed_software.xlsx')
    df.to_excel(output_file, index=False)


if __name__ == "__main__":
    main()
