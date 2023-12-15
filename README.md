# O365 Monitor
Office365 Hygiene and Account Monitoring Scripts

### Installation Prerequisites
  - An Office 365 account with an administrative role that has access to create and edit Enterprise Apps
  - Administrative access on a Windows computer
  - Note: The setup routine creates an Enterprise App with the absolute minimum (and read-only) rights to your O365 environment

### Setup
  - Download the latest release ZIP
  - Right-click on the ZIP and select Properties
  - Click "Unblock," then "OK"
  - Extract to the location of your choice
  - Open an Administrator Powershell window
  - Navigate to the location of the O365 Monitor scripts
  - Run Install-O365Monitor.ps1
    - During the installation, the tool will prompt for Office 365 administrative credentials.  These are only used in the current session to set up and configure the "O365 Monitor" enterprise app.
  - With the installation is complete, you can close the PowerShell window

### Configuring the O365 Monitor account
#### Note: I recommend running the tool as a non-privileged account.  Administrative rights are not needed.
  - Click Start and type "mmc.exe."  Right click on the result and select "Run As Administrator"
  - With MMC open, click File->Add/Remove Snap-In...
  - Select "Certificates" and click "Add"
  - Select "Computer Account"
  - Select "Local Computer"
  - Click "Ok"
  - Expand Certificates (Local Computer) and click on "Personal"
  - Right-click the "O365Monitor" certificate and select "Properties"->"Manage Private Keys" as shown in the image below
![alt text](https://github.com/Xorlent/O365-Monitor/blob/a3d76a7496205632041604d97adbfa896b07d338/PrivateCertPermissions.png "MMC Certificate Properties")
  - Click "Add" and select the user account that will be used to run the O365 Monitor scripts
  - Select only the "Read" right as shown in the image below with example user "OTHER"
![alt text](https://github.com/Xorlent/O365-Monitor/blob/1d84c5880f7efc114cf16b1ffc0b5d5100c84dd7/PrivateCertPermissions2.png "Certificate Permissions")

### Running the scripts
  - The scripts are designed to be run interactively, but I may enhance and further develop more functionality that would facilitate automated execution and notifications
  - Open a PowerShell command prompt
  - Navigate to the location of the O365 Monitor scripts
  - Execute the desired script
    - Get-ExpiringO365AppRegistrations.ps1 : Outputs a file with any secrets or certificates that are expiring in the next 45 days
    - Get-DormantO365Accounts.ps1 : Outputs a file with any enabled accounts that have been dormant for 45 days (great for identifying and purging old external accounts!)

### The script stopped running after a year
  - Follow the Setup and Configuration steps, but instead of running Install-O365Monitor.ps1, run Renew-O365MonitorCert.ps1
  - The newly generated certificate will be valid for 385 days

Many thanks to [Erik de Bont](https://github.com/erik-de-bont) and AdminDroid ([Twitter](https://twitter.com/admiindroid)|[Facebook](https://www.facebook.com/admindroid)|[LinkedIn](https://www.linkedin.com/company/admindroid/)) for good portions of the scripts.
