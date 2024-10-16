# GetFileHashInformation
Drag and Drop Application to view File Hash Information<br>
 ![FirstLoad](/Images/Application_GFHI_FirstLoad.png)

## Example after loading a file
 ![ExampleLoad](/Images/Application_GFHI_Example00.png)

 ## Quick Access
Quickly access the app by invoking the URL [filehashapp.michaeltheadmin.com](https://filehashapp.michaeltheadmin.com)

### Invoke-Expression
```powershell
iex (irm filehashapp.michaeltheadmin.com)
```
 ![Invoke MSI Application](/Images/Application_GFHI_Example_Run.png)

## Right-Click Context Menu Option
Right-Click Context Menu Option to Lauch the app and automatically load the information<br>
![Right-Click Context Menu](/Images/Application_GFHI_ContextMenu_Example.png)

### Install the Context Menu Option
Install the Context Menu Option from the menu bar<br>
![Install Context Menu](/Images/Application_GFHI_ContextMenu_Install.png)

### Uninstall the Context Menu Option
Uninstall the Context Menu Option from the menu bar<br>
![Uninstall Context Menu](/Images/Application_GFHI_ContextMenu_Uninstall.png)

## File Lock Check
If the file is currently being used by another process, you will receive the below error.
![File Locked](/Images/Application_GFHI_Example_FileLocked.png)

## Do Not Run as Admin
The Drag and Drop events will not work when ran as administrator. Run the script without elevated Priveledges.<br>
 ![Invoke MSI Application](/Images/Application_GFHI_Example_RunAsAdmin.png)