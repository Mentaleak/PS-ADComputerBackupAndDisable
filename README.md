# PS-ADComputerBackupAndDisable
This Script scans AD for Computer Accounts that haven't been used in (365) days. Then it allows the user to choose which ones you would like to backup the laps and bitlocker data for. Afterward you are given the option to disable them and move them to the disabled OU
![image](https://github.com/Mentaleak/PS-ADComputerBackupAndDisable/assets/22431171/49d3ab80-3340-4b28-aeeb-cf926c1beb34)

## Configurable Variables

1. **$daysSinceLogon**: The number of days since the last logon to filter computers. Default is set to 365 days.
    ```powershell
    $daysSinceLogon = 365
    ```

2. **$DisabledComputersOU**: The distinguished name of the Organizational Unit (OU) for disabled computers. Modify this to match your domain and OU structure.
    ```powershell
    $DisabledComputersOU = 'OU=Disabled Computers,DC=Domain,DC=COM'
    ```

## Script Overview

This script automates the process of identifying and disabling inactive computer accounts in Active Directory. It performs the following steps:

1. Loads the necessary .NET assembly for using Windows Forms.
2. Sets a filter date based on the number of days since last logon.
3. Retrieves a list of Windows computers from Active Directory with specific properties.
4. Filters out computers already in the Disabled Computers OU.
5. Further filters computers that haven't been changed, had their password set, or LAPS expired within the specified timeframe.
6. Presents a GUI to select computers to back up data.
7. Collects BitLocker recovery information for the selected computers.
8. Exports the selected computers with their data to a file.
9. Presents a GUI to select computers to disable.
10. Creates a new OU named with the current timestamp.
11. Moves the selected computers to the new OU and disables their accounts.
