# PowerShell-Windows-Software-Inventory
This PowerShell Script is crafted to ultimately gather the following information from remote servers/nodes in any enterprise environment with proper authorization. 

Must run PowerShell window as administrator prior to execution. 

# Description
This PowerShell script uses the Get-Ciminstance module instead of the Get-WMIObject to obtain a Remote Software Query for Windows Operating System Servers and Clients. 

This overall automates the Software Inventory Collection for Active Directory Enterprise environments.
Devices will be captured from Active Directory via computer name listing to confirm Active Directory authorization.
Results will go to C: Drive by default 

Information Obtained is the following:
Computer Name, Software Name, Software Version, and Software Vendor


# Procedure Example
Run Powershell Winfow as Admin
Run the following for one computer:

PS C:\YOUR-SCRIPT-PATH> .\SW-INV.ps1 -ComputerName PC1 

Run the following for a list of computers from your path 

PS C:\YOUR-SCRIPT-PATH> .\SW-INV.ps1 -ComputerList C:\PATH-TO-LIST\Computer_List.txt
      
All results will be exported to CSV file in the C: drive for administrators to take and move as needed

# Credits
Written by: Javier Walters

# Social Network
LinkedIn: https://www.linkedin.com/in/javier-walters/
