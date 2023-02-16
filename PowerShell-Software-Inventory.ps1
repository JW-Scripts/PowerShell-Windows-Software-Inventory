<#
################################################
Title: Windows OS Remote Software Inventory
Description: 
    - Remote Software Query for Windows OS Servers and Nodes 
Updates:
    - Folder Name with Date with Time 
    - variable Box
    - AD Computer Count
    - Progress Bar for Unit Capture
    - File Movement and Verification 
    - Replaced "cd" with Set-Location
    - Software Inventory Notes 
    - Create Directory with date / time 

################################################
#> 

<#
.SYNOPSIS
        Remote Software Inventory Collection for Enterprise Active Directory Environments.
    
.DESCRIPTION
        Automates the Software Inventory Collection for Enterprise environments
        Capture devices from Active Directory when searching for Computer Name, Software Name, or List of Computers
        Results will go to C: Drive by default 

.EXAMPLE
        PS C:\YOUR-SCRIPT-PATH> .\SW-INV.ps1 -UnitName HRUNIT21
        Executes scan on specified unit name (eg. HRUNIT21)
    
.EXAMPLE
        PS C:\YOUR-SCRIPT-PATH> .\SW-INV.ps1 -ComputerName PC1 
        Executes scan on the specific Computer Name (eg. PC1)

.EXAMPLE
        PS C:\YOUR-SCRIPT-PATH> .\SW-INV.ps1 -ComputerList C:\PATH-TO-LIST\Computer_List.txt
        Executes scan on list of computers from your path 

.EXAMPLE
        PS C:\YOUR-SCRIPT-PATH> .\SW-INV.ps1 -SoftwareName "Trellix" -List C:\Fake_Folder\My_CPU_Targets.txt
        - Executes Software scan for specific software name (eg. Trellix) on remote computer list (eg. Targets.txt) 
        - This will inform you if Software is installed on system or not 
.LINK
        https://docs.microsoft.com/en-us/powershell/module/cimcmdlets/get-ciminstance?view=powershell-7.2
        
        https://ss64.com/ps/get-ciminstance.html
        
        Both links above provide More insight on the Get-CimInstance cmdlet  
#> 

##Parameters for Script 
param (
[string]$ComputerName,$ComputerList,$UnitName,$SoftwareName,$List
) 

## SET Variables and SET Directories
##############################################################################################
## Variables and Directories for Script 
## 
## ---------------- Organizational Units ----------------
##List of Organizational Units within the Active Directory for the Script to be Ran against Specific Units or Sections as well
$OUs = 
## YOUR COMPANY OU; UNIT 1
"OU=UNIT1,OU=COPMANY,OU=DOMAIN,OU=NAME,OU=HERE",
## YOUR COMPANY OU; UNIT 2
"OU=UNIT2,OU=COPMANY,OU=DOMAIN,OU=NAME,OU=HERE"
##
## ----------------------- Date and Time for New Folder Name Creation --------------------------
$DateAndTime = (Get-Date).tostring("MM-dd-yyyy@hh-mm-tt") 
## ----------------------- Date and Time for New Folder Name Creation --------------------------
##
## Date for Report Name
$Date = (Get-Date).tostring("MM-dd-yyyy")
##
## Variables and Directories for Script
##############################################################################################

## ---------- BREAK ---------- ##
## Software Script Begins
## ---------- BREAK ---------- ##

## SET Title Box  
#################################################
## Title Box for Script 
Clear-Host 
Write-Host "
############################
##                            
## Windows Active Directory           
## Remote Software Inventory     
##
############################
" -ForegroundColor Yellow

## Title Box for Script 
################################################


## ---------- BREAK ---------- ##


## COMPUTERNAME PARAM SECTION
############################################################################################################################
## START OF IF SECTION FOR $COMPUTERNAME SELECTION
IF ($ComputerName) {
## == Test Connection of Target Computer
## == IF Computer is Online, make directories & gather information || Else, State Computer is Offline 
IF (Test-Connection -ComputerName $ComputerName -Count 1) {
                        ## Make New Directory When Starting Script 
                        $CN_DIR = mkdir "C:\$($ComputerName)_$($DateAndTime)"
                        IF ($CN_DIR) {Write-Host "Created Directory for $ComputerName at:" -ForegroundColor Green;Write-Host "$CN_DIR" -ForegroundColor Cyan }

                        ## Change Working Directory to New Directory 
                        Set-Location "C:\$($ComputerName)_$($DateAndTime)"

                        ## Create Export Location for Software Results 
                        $Ex_Location = "C:\$($ComputerName)_$($DateAndTime)\$($ComputerName)_Software_List.csv"
                        
                        ## Grab IP Address of Computer Name to verify DNS Registration 
                        $CN_IP = (Resolve-DnsName $ComputerName -Type A).IPAddress

                        ## State IP Address and Computer Name
                        Write-Host "$ComputerName ($CN_IP) Online" -ForegroundColor Green

## Get Software from Remote Machine 
Get-CimInstance -ComputerName $ComputerName -ClassName win32_product -ErrorAction SilentlyContinue | Select-Object PSComputerName, Name, PackageName, Version, Vendor  | Sort-Object Name | Export-Csv -Path $Ex_Location -Append -NoTypeInformation
##
## Confirm that Software File Export Exists in Location --------------------
IF (Get-ChildItem -Path $Ex_Location ) {
                        ## Success
                        Write-Host "Results Exported to:" -ForegroundColor Magenta;Write-Host "$Ex_Location" -ForegroundColor Cyan } ELSE {
                        ## Fail
                        Write-Host "Failed to Export or Obtain Results" -ForegroundColor Red } } 
##
## ELSE Section states if Computer is Offline
} 
##ELSE {Write-Host "$ComputerName OFFLINE" -ForegroundColor Red } 
##
## END OF IF SECTION FOR $COMPUTERNAME SELECTION
############################################################################################################################


## ---------- BREAK ---------- ##


## COMPUTER LIST SECTION
############################################################################################################################
## START OF IF SECTION FOR $COMPUTERLIST PARAM
##
IF ($ComputerList) {
                    ## Make New Directory When Starting Script
                    $CPUListDIR = mkdir "C:\CPU_LIST_$($DateAndTime)";IF ($CPUListDIR) {Write-Host "Created Directory for Computer List at:" -ForegroundColor Green;Write-Host "$CPUListDIR" -ForegroundColor Cyan }
                    
                    ## Make New Directory When Starting Script
                    Set-Location "C:\CPU_LIST_$($DateAndTime)"
                    
                    ## Export Location for Software Results
                    $Ex_Location = "C:\CPU_LIST_$($DateAndTime)\Computer_List_Results.csv"

## Gather the Computer Names from the List of Computers
$ComputerListing = Get-content $ComputerList

## For Each Computer Identified in the List of Computers
Foreach ($Computer in $ComputerListing) {

        ## == Test Connection of Each Computer
        ## == IF Computers are Online, make directory & gather information for each device || Else, State Computer is Offline 
        IF (Test-Connection -ComputerName $Computer -Count 1) {
                    
                    ## Grab IP Address of Target Device to verify DNS Registration || State Computer/IP Address is Online 
                    $CPU_IP = (Resolve-DnsName $Computer -Type A).IPAddress;Write-Host "$Computer ($CPU_IP) Online" -ForegroundColor Cyan
                    
                    ## Send content to Online text file                     ## Send Computer to Offline text file 
                    Add-Content -Value $Computer -Path .\Online.txt } ELSE { Add-Content -Value $Computer -Path .\Offline.txt }

## Get Online Computers
$OnlineCPUs = Get-content -Path .\Online.txt

## Get Software from Online Remote Machines and Export to CSV
Get-CimInstance -ComputerName $OnlineCPUs -ClassName win32_product -ErrorAction SilentlyContinue | Select-Object PSComputerName, Name, PackageName, Version, Vendor  | Sort-Object Name | Export-Csv -Path $Ex_Location -Append -NoTypeInformation }

        ## Confirm that Software File Export Exists in Location 
        IF (Get-ChildItem -Path $Ex_Location ) { 
                
                Write-Host "Results Exported to:" -ForegroundColor Magenta;Write-Host "$Ex_Location" -ForegroundColor Cyan } else { 
                
                Write-Host "Failed to Export or Obtain Results" -ForegroundColor Red } }
##
## END OF IF SECTION FOR $COMPUTERLIST PARAM
############################################################################################################################


## ---------- BREAK ---------- ##


## UNIT NAME OR SECTION NAME IN ACTIVE DIRECTORY OU
############################################################################################################################
## START OF IF SECTION FOR $UNITNAME PARAM
IF ($UnitName) {
                ## Make New Directory with Unit Name as Title 
                $UnitNameDIR = mkdir "C:\$($UnitName)_$($DateAndTime)";IF ($UnitNameDIR) {Write-Host "Created Directory for $UnitName at:" -ForegroundColor Green;Write-Host "$UnitNameDIR" -ForegroundColor Cyan }
                
                ## Change Location to the New Directory 
                Set-Location "C:\$($UnitName)_$($DateAndTime)"
                
                ## Final SW Report Name
                $Name_Of_Report = "$($UnitName)_Sw_Per_CPU.csv"
                
                ## Date for Final SW Report Name
                $Date = (Get-Date).tostring("MM-dd-yyyy")
                
                ## Final SW File Export
                $File_Export = "C:\$($UnitName)_$($DateAndTime)\" + $Date + $Name_Of_Report

## Action 1: Enumeration START -------------------------

## Gather/Export Computer Names from the Active Directory OU based on the Targeted UnitName Parameter 
##

Write-Host "Capturing $UnitName Computers from Active Directory" -ForegroundColor Green
##
foreach ($OU in $OUs) {

                ## Enumerate the Organizational Unit Specified in the Software Inventory Scipt
                $AD_OU = (Get-ADOrganizationalUnit -Filter "Name -Like '*$UnitName*'" -SearchBase $OU).DistinguishedName

                ## Extracts the computer names that are needed from the specified OU
                IF ($AD_OU) {(Get-ADComputer -Filter * -SearchBase $AD_OU).Name | Out-File .\$($UnitName)_List.txt } }
##
## Gather/Export Computer Names from the Active Directory OU based on the Targeted UnitName Parameter 

## Action 1: Enumeration END -------------------------


## Action 2: Counting of Computer Names - START -----------------  
##
## Get the list of AD Computers
$AD_List = Get-Content -Path .\$($UnitName)_List.txt

## Counts Computer Names Found From AD OU
$CPUCount = $AD_List.Count

## States Counted Computers 
Write-Host "$CPUCount Computers Found for $UnitName" -ForegroundColor Cyan 
##
## Action 2: Counting of Computer Names - END -----------------  


## Action 3: Connection Tests - START -----------------------------
##
## TESTS COMPUTER CONNECTION
##

## Comment Box Statement: States Testing Connection to AD Computers
Write-Host ""
Write-Host "---------------------------------------" -ForegroundColor Magenta
Write-host "TESTING CONNECTION TO $UnitName COMPUTERS" -ForegroundColor Yellow
Write-Host "---------------------------------------" -ForegroundColor Magenta

##Gathers List from AD Export
$AD_Computers = Get-Content -Path .\$($UnitName)_List.txt

## Tests the connection to target computers before gathering inventory information || Creates an Online and Offine List of Computers
foreach ($Computer in $AD_Computers) {

IF (Test-Connection -ComputerName $Computer -Quiet -Count 1) {
        
        $UN_IP = (Resolve-DnsName $Computer -Type A).IPAddress;Write-Host "$Computer ($UN_IP) Online" -ForegroundColor Cyan;Write-Host ""
        Add-Content -Value $Computer -Path .\$($UnitName)_Computers_Online.txt
        Add-Content -Value $Computer -Path "C:\Total_Gathered&Missing\Total_Gathered.txt"
            } ELSE {
                    Write-host "$Computer Offline" -ForegroundColor Red;Write-Host ""
                    Add-Content -Value $Computer -Path .\$($UnitName)_Computers_Offline.txt
                    Add-Content -Value $Computer -Path "C:\Total_Gathered&Missing\Total_Missing.txt" 
   }
} 
## 
## Action 3: Connection Tests - END -----------------------------


## Action 4: Get Software - START -----------------------------
##

## Get List of Online Remote AD systems from the Online text file  
$CPU_INV = Get-Content -Path .\$($UnitName)_Computers_Online.txt 

## Counts the amount of Online computers
$NUM = $CPU_INV.Count

## Comment Box Statement: States the Counted Amount of Online AD Computers
Write-host "$NUM $UnitName Computers are Online. Please wait while Software is being gathered for each Computer..." -ForegroundColor Yellow

##Run Get-Ciminstance against the online AD Computers to Gather software from each computer 
$Get_SW = Get-CimInstance -ComputerName $CPU_INV -ClassName win32_product -ErrorAction SilentlyContinue | Select-Object PSComputerName, Name, PackageName, Version, Vendor | Sort-Object -Property Name

##Counts the amount of Computers Inventoried for Software 
$Get_SW_Count = $Get_SW.Count

##
## Action 4: Get Software - END -----------------------------



## Action 5: Show Inventory Progress - START --------------------------------
##

##Progress Bar for Status of Software being gathered for remote systems
$i = 0 
foreach ($SW in $Get_SW) {$SW
$i = $i + 1
$Perc = ($i/$Get_SW_Count)*100 ## Qualitative Calculation for Completion Status
Write-Progress -Activity "SW Status: $Perc" -PercentComplete $Perc ## Progress Bar for Completion Status
}
## 
##Notify If percentage is 100
if ($Perc -eq 100) {Write-host "SOFTWARE GATHERED FOR COMPUTERS IN $UnitName." -ForegroundColor Cyan
                                    } ELSE {
                    Write-host "FAILED TO COMPLETE SOFTWARE INVENTORY FOR COMPUTERS IN $UnitName." -ForegroundColor Red}
##
## Action 5: Show Inventory Progress - END --------------------------------



## Action 6: Export Inventory to CSV - START ----------------------------------------- 
##

## The Software Inventory will be exported to CSV  
$Get_SW | Export-CSV -Path $File_Export

## Verify the Location of the Exported CSV file to determine completed export of Software to CSV 
IF (Get-ChildItem -Path $File_Export) {
Write-host "-----------------------------------------------------------" -ForegroundColor Green
Write-Host "Exported $UnitName Computer Software to CSV at:" -ForegroundColor Magenta
Write-Host "$File_Export" -ForegroundColor Cyan
Write-Host "Software Inventory Complete for $UnitName" -ForegroundColor Green
Write-host "-----------------------------------------------------------" -ForegroundColor Green
} ELSE {
Write-host "FAILED TO EXPORT $UnitName Computer Software to CSV at:" -ForegroundColor Magenta
} 

##
## Action 6: Export Inventory to CSV - END ----------------------------------------- 


## Action 7: Create Software Summary Notes - START ----------------------------------------- 
##

##Sends Summary of Notes that include the Date, Number of computers gathered, and unit inventoried for software || Uses the CMD "echo" method to write to new text file
##
$Note_Export = ".\$($UnitName)_SW_Notes.txt" 
Write-Output "On $Date $NUM AD Computers were Online and Inventoried for SW for $UnitName" >> $Note_Export
## 
## Action 7: Create Software Summary Notes - END ----------------------------------------- 
}
##
## END OF IF SECTION FOR $UNITNAME PARAM
############################################################################################################################


## ---------- BREAK ---------- ##


############################################################################################################################
## FIND SPECIFIC SOFTWARE BY NAME  
##
IF ($SoftwareName) {
## Statement of Searching for Software
Write-Host "Searching $List for $SoftwareName" -ForegroundColor Yellow
##
## Read Content of Software List 
$Listing = Get-Content $List
##
## Action on Each Computer
foreach ($Computer in $Listing) {
## Gathers Specific Software Name through Get-Ciminstance Win32_Product Search 
$FindSW = Get-CimInstance -ComputerName $Computer -ClassName win32_product -ErrorAction SilentlyContinue | Select-Object PSComputerName, Name, PackageName, Version, Vendor | Sort-object Name | Where-Object {$_ -match "$SoftwareName"}
##
## IF Statement for SW Found or Not
IF ($FindSW) {Write-Host "$SoftwareName Found on $Computer" -ForegroundColor Cyan} ELSE {Write-Host "$SoftwareName NOT installed on $Computer"} 
## Export Results to CSV
$FindSW | Export-Csv -Path "C:\$($SoftwareName)_OutPut.csv" -Append -NoTypeInformation }
##
## Verify Exported CSV File Exists 
$Location = "C:\$($SoftwareName)_OutPut.csv"
IF (Get-ChildItem -Path $Location) {Write-Host "Results Exported to:" -ForegroundColor Magenta;Write-Host "$Location"-ForegroundColor Cyan} ELSE {Write-Host "Failed to Export Contents"}
} 
##
## FIND SPECIFIC SOFTWARE BY NAME 
############################################################################################################################


#########################
##### End of Script #####
#########################


