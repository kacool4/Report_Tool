# Reporting Tool -- Work in Progress, The script is not yet finished...

Alternative to RVTools a script with 1 liners and checks in order to report vmware environment. 

## Description : 
   Provides an Excel with all necessary information about VMWare infra. The script sents the output file via mail.
   
## Outputs :
   - VM Info (CPU,Memory,Space,IP,OS)
   - VM Disks Info (Disk ID,Capacity,Type,Location)
   - VM Network Info (VM name,Adapter,Type,IP,Mac,Network)
   - ESXi and Cluster Info (Cluster,Socket,CPU Cores,Threads,Memory,IP,License,Vendor,CPU Type)
   - Datastores Info (Name, Cluster, UUID, Total and Free space %, Hosts attached, Type)
   - Snapshots Info (VM, Name, Created date, Description,Size)
   - VMTools Info (VM, Version, Status)
   - NIC Info (Host,Model,Driver,Firmware,Description)
   - HBA Info (Host,Model,Driver,Firmware,Description)
   - WWN/WWP Info (Host,Model,Status,WWN,WWP)
   - VMKernel Adapter Info (Host,Device,Mac,IP,Subnet,Port group)
   - vDS Info (Hosts,vDS Name,No of Ports,Vlan,Security configuration)
   - sDS Info (Hosts,vDS Name,No of Ports,Vlan,Security configuration)
   - License Info (Provides list of licenses, total and used.)
   - ...More to be added in the near future.

## Issues :
   - 

## Updates :
  
 (20/04/2022)
   
    - Added Template tab in the report
    - Reordering some columns
  
 (24/02/2022)
   
    - Minor changes/Fixes
    - Nics assign to vSS
    - Nics assign to vDS
    - List of vDS and portgroups
    - List of vSS and portgroups
 
(23/02/2022)
   
    - Minor changes/Fixes
    - Re-arranege of tabs
    - Align code for better reading
 
(22/02/2022)
   
    - Fixed list of ESXis on Datastore Info tab 
    - Fixed Bios date on Esxi and Cluster Info
    - Added Powerstate and Connection status in Esxi and Cluster Info 
    - VM datastores in VM Info (Shows all datastores that the vm has disks)
    - Network IP-per-Mac, Status and Connected at start up was added  (VM network)
    - Zombie disk size in GB
    - Disk SCSI (VM DIsk)
 
 (14/02/2022)
 
    - Use param -noconnection in order to run the script when you are already logged in a vCenter. Result files will not be sent via mail. Suitable for people that want only results for the vCenter that is already logged 
    - Use param -mailreport in order to sent the zip file via mail to the recipients that are specified in the script
    - Use param -version to check version, created and last modified date of the script
    - Use param -help to get info about available parameters



## Requirements:
* Windows Server 2012 and above or Windows 10
* Powershell 5.1 and above
* ImportExcel Module (Install-Module -Name ImportExcel  or check below on how to install it offline) 
* PowerCLI 12.0 + either standalone or import the module in Powershell (Preferred)
* A text file "List.txt" in order to specify the vCenters to take report from

## User Role Creation:
* Create a new Role in vCenter with the following permissions
   * Datastore --> Browse Datastore (We need this permission in order to search for Zombie files)
   * Global --> Licenses (We need this permission in order to view the licesnses)

## Offline ImportExcel Module installation:
* Download the zip file from github page [Here](https://github.kyndryl.net/Dimitrios-Kakoulidis/Report_Tool/blob/main/ImportExcel.zip)
* Unzip and copy the folder to Powershell folder (C:\Program Files\WindowsPowerShell\Modules)
* Open powershell command line and type "Import-Module -Name ImportExcel" 

## Configuration

Specify the location of the List.txt or leave it as default (It will search for the file in the same folder as the script is stored)
```powershell
foreach($line in Get-Content .\list.txt) 
```
Write the vCenters in the list.txt as a list and not as a single line seperated by comma.
```
Valid list of vms 
vCenter1
vCenter2
vCenter3
vCenter4

Not Valid
vCenter1,vCenter2,vCenter3,vCenter4
```

## Run the script

  The script reads the list.txt file and then tries to connect to the first vCenter. Export all the info. Creates a excel file with the vCenter name, disconnects from the current session and read the other line of the txt file. Repeats the steps of Connect - output excel file - disconnect from vCenter for all the vCenters on the txt list.
  
  ```
    Report_Tool.ps1               // By default it will run the script asks for credentials but no mail will be sent.
    Report_Tool.ps1 -noconnection // Will not ask for login. No mail will be sent.
    Report_Tool.ps1 -Sent_mail    // Will ask for credentials and sent mail
    Report_Tool.ps1 -help         // Displays param options to run the script
    Report_Tool.ps1 -version      // Gives latest version of script
  ```

In order to sent the results automatically after all reports are created you need to remove # from the following line and setup the correct details:

```
  send-mailmessage -from 'ReportTools@Something.COM' -to 'Test@test.com' -subject 'ReportTools: Here is your report $(get-date -f 'dd-MM-yyyy')' -body 'Below you can find the rvtools report. Please see attachment `n `n `n' -Attachments $destination -smtpServer IP
```

## Credentials in order to run the script
 Due to security reasons there is no username and password written in the script. The encryption is not strong so you can run the script with 2 possible ways :
 - Use a domain account that has access to the vCenters. In this case it will connect and disconnect from all without asking you a password
 - Type each time it tries to connecto to vCenter the username and password


# Report Tools Output:
## General Tabs
![Alt text](/screenshots/output4.jpg?raw=true "General Tabs")
## VM Details Tab
![Alt text](/screenshots/output1.jpg?raw=true "VM Details Tab")
## Details for Network Adaptets Tab
![Alt text](/screenshots/output2.jpg?raw=true "Details for Network Adaptets Tab")
## WWN Info Tab
![Alt text](/screenshots/output3.jpg?raw=true "WWN Info Tab")
## Datastore Info Tab
![Alt text](/screenshots/output5.jpg?raw=true "Datastore Info Tab")

....and many more detailed tabs about your vmware infra.
