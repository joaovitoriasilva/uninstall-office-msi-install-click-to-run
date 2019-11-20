# uninstall-office-msi-install-click-to-run
Uninstall Office MSI and Install Click to run (Powershell)

DISCLAIMER: The Office 365 ProPlus XML installation files provided may not be up to date. Please adapt this files to your needs and based on manufacturer best pratices.

This script will detect Office Versions (2003 and up) installed in the machine where is being runned and will remove them. After remove them, it will install the Office 365 ProPlus click-to-run version.

Script changes:
 - 1.0 - Initial version
 - 1.1 - Added reg keys to remove first run setup
 - 1.2 - Added support for french and portuguese OS LP
 - 1.3 - Added support for uninstalling previously installed Office 365 ProPlus versions; Fixed error when calling Office 365 ProPlus installer with spaces in the path
 - 1.4 - Added detection for solo installations of Office 365 Project and/or Visio clients

If the script detects a version of Visio and/or Project installed, it will also install the click-to-run version.

The script has two parameteres:
 - installVisio365: Variable to define if Visio is found by the script, will it install the Visio 365 Version. By default $false.
 - installProject365: Variable to define if Project is found by the script, will it install the Project 365 Version. By default $false.

Do "get-help .\Office365ProPlusDeploy.ps1 -full" to see the script help.

Requirements:
 - Download the Office 2016 Deployment Tool from Microsoft site: https://www.microsoft.com/en-us/download/details.aspx?id=49117:
 - Copy the file "setup.exe" to the folder where the files are stored;
 - Check Microsoft article to review how to build XMLs files: https://docs.microsoft.com/en-us/deployoffice/configuration-options-for-the-office-2016-deployment-tool;
 - Use Deployment settings page to build your XML files: https://config.office.com/deploymentsettings;
 - Open a CMD/PowerShell window and cd to the folder where the files are stored. Do "setup.exe /download download.xml" to dowload all the necessary Office files;
 - Change the provided XML files to your needs.

The sample scripts are provided AS IS without warranty of any kind.
