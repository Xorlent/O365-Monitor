# O365-Monitor
Office365 Hygiene and Account Monitoring Tool


### Installation Prerequisites
  - An Office 365 account with an administrative role that has access to create and edit Enterprise Apps
  - Administrative access on a Windows computer

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


### Reading the report
  - After a completed run, find the CSV file and open with your favorite editor
  - Column 1 indicates whether the tool remediated the vulnerability or not









Many thanks to [Erik de Bont](https://github.com/erik-de-bont) and AdminDroid ([Twitter](https://twitter.com/admiindroid)|[Facebook](https://www.facebook.com/admindroid)|[LinkedIn](https://www.linkedin.com/company/admindroid/)) for good portions of the scripts.
