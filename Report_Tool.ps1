
<#

 Author: Dimitrios Kakoulidis
 Date Create : 11-02-2022
 Last Update : 06-05-2022
 Version: 1.19

 .Description 
   Alternative to RVTools. The script provides an Excel with all necessary information about VMWare infra. The script sents the output file via mail 
   
 .Outputs
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
   - License Info (Provides list of licenses, total and used)
   - Zombie Disk Info (VMDK info, Datastore, Size)
   - ATS Heart Beat Info (0=Disabled // 1=Enabled)
#> 
    ###########################################################################################################################################
    ######################## Start of Script ##################################################################################################
    ########################################################################################################################################### 

    #Param values #####################################################################################################
 

    param([switch] $help,
          [switch] $mailreport,
          [switch] $noconnection,
          [switch] $version
    )


    # Bypass  policy #############
        $Bypass = Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
        $Bypass

    # General Variables #####################################################################################################
 
    # Date and version of Script
        $DMTLversion = 'v1.19' 
        $modDate = '06-05-2022'

    # Steps and Total steps variables

    $CurrentStep = 1
    $TotalSteps = 20

    # Get Date and create folder
        $Date = (Get-Date -f 'ddMMyyyy')
        New-Item -Path "C:\Scripts\DimiTools\$Date" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        $folder = 'C:\Scripts\DimiTools\'+$Date+'\'
        $source = $folder
        $destination = 'C:\Scripts\DimiTools\Archives\'+$Date+'.zip'
        $Connect=$false
        $sent_mail=$false

       

    # Check param values #####################################################################################################
    if ($help)
    {
        Write-Host
        Write-Host ' No param used, By default it will run the script asks for credentials but no mail will be sent.'
        Write-Host
        Write-Host ' Use param -noconnection in order to run the script when you are already logged in a vCenter. No mail will be sent. Suitable only if you want results for the current vCenter. List.txt MUST be filled with the hostname for the vCenter'
        Write-Host
        Write-Host ' Use param -mailreport in order to sent the zip file via mail to the recipients that are specified in the script'
        Write-Host
        Write-Host ' Use param -version to check version, created and last modified date of the script'
        Write-Host ' '
        Exit
    }
    elseif ($mailreport)
    {
        $sent_mail=$true
    }
    elseif ($noconnection)
    {
        $Connect=$true

    }elseif ($version)
    {
        Write-Host ' Author      : Dimitrios Kakoulidis'
        Write-Host ' Date Create : 11-02-2022'
        Write-Host ' Last Update :',$modDate
        Write-Host ' Version     :',$DMTLversion
        Exit
    }

    # Start Collecting Info  #####################################################################################################

    foreach($line in Get-Content C:\Scripts\DimiTools\list.txt) 
    {


    ######## Get vCenter name and create a save name/path #####################################################################
    $vCenter = $line
    $outputpath=$folder+$vCenter+'.xlsx'
    

    ############  Connect to vCenter ######################################################################################## 

    if (!($Connect) -or ($sent_mail)) {
       Write-Host 'Connecting to vCenter...'
       Connect-VIServer $vCenter
       Write-Host 'Starting script.......'
    }  

    ######## VM Info output ##################################################################################################
    Write-Host "($CurrentStep/$TotalSteps) Gathering VM Info..."

    Get-VM | Select Name,
                    PowerState,
                    @{N='CPU';E={$_.NumCpu}},          
                    @{N='Memory';E={$_.MemoryGB}}, 
                    @{N='HDD Total Space (GB)'; E={[math]::round($_.ProvisionedSpaceGB)}},
                    @{N='IP Address';E={@($_.guest.IPAddress -join '|')}},
                    @{N='ESXi Host';E={Get-VMHost -VM $_}}, 
                    @{N='Guest OS';E={$_.Guest.OSFullName}},
                    @{N='Hardware Version'; E={$_.version}},
                    @{N='Notes '; E={$_.Notes}},
                    @{N=’Datastore’;E={[string]::join(“;”, (Get-Datastore -VM $_))}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'VM Info'
    
    Write-Host "($CurrentStep/$TotalSteps) VM Info Completed"     
    $CurrentStep ++

    ######## Template Info output ##################################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering Template Info..."

    Get-Template | Select Name,
                          @{N='CPU';E={$_.ExtensionData.Config.Hardware.NumCPU}},
                          @{N='Memory';E={$_.ExtensionData.Config.Hardware.MemoryMB}},
                          @{N='Guest OS'; E = { $_.ExtensionData.Config.GuestId}},
                          @{N='Storage (GB)';E={[Math]::Round(($_.ExtensionData.Summary.Storage.Committed/1GB),1)}}, 
                          @{N='Host';E={(Get-VMhost -id $_.HostID).Name}}, 
                          @{N='Datastore'; E={(Get-Datastore -id $_.DatastoreIDlist).Name -join ','}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Template Info'
   
    Write-Host "($CurrentStep/$TotalSteps) Template Info Completed"        
    $CurrentStep ++
    
    ######## Disk Info output ##################################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering Disk Info..."

    Get-VM | Get-HardDisk |Select @{N='VM';E={$_.Parent}}, 
                                  @{N='Disk ID';E={$_.Name}}, 
                                  @{N='HDD Capacity (GB)'; E={[math]::round($_.CapacityGB)}},
                                  @{N='Disk Type'; E={$_.DiskType}}, 
                                  @{N='Format';E={$_.StorageFormat}}, 
                                  @{N='SCSI  id';E={$hd = $_
                                                $ctrl = $hd.Parent.Extensiondata.Config.Hardware.Device | where{$_.Key -eq $hd.ExtensionData.ControllerKey} 
                                               "$($ctrl.BusNumber):$($_.ExtensionData.UnitNumber)"}},
                                  @{N='VMDK Location';E={$_.Filename}}| Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'VM Disk'
    
    Write-Host "($CurrentStep/$TotalSteps) Disk Info Completed"   
    $CurrentStep ++
    
    ######## Disk Partition Info output ############################################################################################      
    
    Write-Host "($CurrentStep/$TotalSteps) Gathering Disk Partition Info..."
    
    $TVMs = Get-VM
    ForEach ($VM in $TVMs){
       $VM.Extensiondata.Guest.Disk | Select @{N="Name";E={$VM.Name}},
                                             DiskPath, 
                                             @{N="Capacity(GB)";E={[math]::Round($_.Capacity/ 1GB)}}, 
                                             @{N="Free Space(GB)";E={[math]::Round($_.FreeSpace / 1GB)}},
                                             @{N="Free Space %";E={[math]::Round(((100* ($_.FreeSpace))/ ($_.Capacity)),0)}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Partition'
    }
                     
    Write-Host "($CurrentStep/$TotalSteps) Disk Partition Info Completed"   
    $CurrentStep ++                 
                              
    
    ######## VM Network Info output ############################################################################################                            
    
    Write-Host "($CurrentStep/$TotalSteps) Gathering Network Info..."

    Get-VM -PipelineVariable vm | Get-NetworkAdapter | Select @{N='VM Name';E={$vm.Name}},
                                                               @{N='Power State';E={$vm.Powerstate}},
                                                               @{N='Adapter ';E={$_.Name}},
                                                               Type,
                                                               @{N='IP Address';E={$nic = $_; ($vm.Guest.Nics | where{$_.Device.Name -eq $nic.Name}).IPAddress -join '|'}},
                                                               MacAddress,
                                                               @{N='Network';E={$_.NetworkName}},
                                                               @{N='Connected';E={$_.ConnectionState.Connected}},
                                                               @{N='Connected at Power on';E={$_.ConnectionState.StartConnected}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'VM Network'
                                                               
    Write-Host "($CurrentStep/$TotalSteps) Network Info Completed." 
    $CurrentStep ++

                                                               
    ######## ESXi and Cluster output ########################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering ESXi and Cluster Info..."

    Get-VMHost | select @{N='Datacenter';E={@(Get-Datacenter -vmhost $_.name)}}, 
                        @{N='Cluster';E={@($_.Parent)}},   
                        Name,
                        @{N='Power Status ';E={$_.PowerState}},
                        @{N='Connection State';E={$_.ConnectionState}},
                        @{N='Sockets';E={$_.ExtensionData.Hardware.cpuinfo.NumCPUPackages}},
                        @{N='CPU Cores';E={$_.ExtensionData.Hardware.cpuinfo.NumCPUCores}},
                        @{N='CPU Threads';E={$_.ExtensionData.Hardware.cpuinfo.NumCPUThreads}},
                        @{N='Memory (GB)'; E={[math]::round($_.MemoryTotalGB)}},
                        @{N='IP Address';E={Get-VMHostNetworkAdapter -VMHost $_ -VMKernel | ?{$_.ManagementTrafficEnabled} | %{$_.Ip}}},
                        @{N='ESXi Version';E={@($_.version)}},
                        @{N='ESXi Build';E={@($_.build)}},
                        @{N='Uptime in Days'; E={New-Timespan -Start $_.ExtensionData.Summary.Runtime.BootTime -End (Get-Date) | Select -ExpandProperty Days}},
                        @{N='License';E={@($_.LicenseKey)}},
                        @{N='Vendor';E={$_.ExtensionData.Hardware.SystemInfo.Vendor}},
                        @{N='Model';E={$_.ExtensionData.Hardware.SystemInfo.Model}},
                        @{N='BIOS Version';E={$_.ExtensionData.Hardware.BiosInfo.BiosVersion}},
                        @{N='BIOS Release Date';E={$_.ExtensionData.Hardware.BiosInfo.releaseDate}},
                        @{N='CPU Type';E={$_.ProcessorType}} | Sort-Object Datacenter,Cluster| Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'ESXi And Cluster'
    
    Write-Host "($CurrentStep/$TotalSteps) ESXi and Cluster Info Completed."
    $CurrentStep ++

                                               
    ######## Datastore output #################################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering Datastore Info..."

    Get-Datastore | Select Name,
                        @{N='NAA Address ';E={$_.ExtensionData.Info.Vmfs.Extent[0].DiskName}},
                        @{N='Total (GB)'; E={[math]::round($_.CapacityGB)}},
                        @{N='Free (GB)'; E={[math]::round($_.FreeSpaceGB)}},
                        @{N='Free (%)'; E={[math]::Round(($_.freespacegb / $_.capacitygb * 100),2)}},
                        @{N='DS Folder'; E={$_.ParentFolder}},
                        @{N='DS Cluster'; E={Get-Datastorecluster -Datastore $_}},
                        @{N=’Hosts’;E={[string]::join(“;”,(Get-Datastore $_ | Get-VMHost))}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Datastores'
    
    Write-Host "($CurrentStep/$TotalSteps) Datastore Info Completed"
    $CurrentStep ++
                        
    ######## VM Tools output ###################################################################################################
    
    Write-Host "($CurrentStep/$TotalSteps) Gathering VMTools Info..."
    
    Get-VM | % { get-view $_.id } | select name, 
                                   @{N='Tools Version'; E={$_.config.tools.toolsversion}}, 
                                   @{N='Tool Status'; E={$_.Guest.ToolsStatus}},
                                   @{N='Version Status'; E={$_.Guest.ToolsVersionStatus}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'VMTools'
     
    Write-Host "($CurrentStep/$TotalSteps) VMTools Info Completed" 
    $CurrentStep ++
                                   
    ########### NIC output ########################################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering NIC Info..."

    $vmhosts = Get-VMHost | Where{$_.ConnectionState -eq 'Connected'}
    $OutputNIC = @()

    foreach ($ESXHost in $vmhosts) {
       $esxcli = Get-EsxCli -vmhost $ESXHost
       $nicfirmware = $esxcli.network.nic.list()
       $driversoft = $esxcli.software.vib.list()

      foreach($nicfirmwareselect in $nicfirmware)
      {
        $NetworDescription = $nicfirmwareselect.Description
        $NetworDriver = $driversoft | where {$_.name -eq ($nicfirmwareselect.Driver)}
        $NetworkName = $nicfirmwareselect.Name
        $NetworkFirmware = ($esxcli.network.nic.get($nicfirmwareselect.Name)).DriverInfo.FirmwareVersion
        $OutputNIC += '' | select @{N='Hostname';E={$ESXHost.Name}},
                                  @{N='ESXi Model';E={$ESXHost.Model}},
                                  @{N='Driver Ver.';E={$NetworDriver.Version}},
                                  @{N='Firmware Ver.';E={$NetworkFirmware}},
                                  @{N='NIC Descr.';E={$NetworDescription}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'NIC Info'
       }
 
     }
   
    Write-Host "($CurrentStep/$TotalSteps) NIC Info Completed."
    $CurrentStep ++
    
     
    ######## HBA output ###########################################################################################################
  
    Write-Host "($CurrentStep/$TotalSteps) Gathering HBA Info..."

    $HBAHosts = Get-VMHost | Where{$_.ConnectionState -eq 'Connected'} 
    $OutputHBA = @()
    
    foreach ($HBAHost in $HBAHosts) {
          $esxcli = Get-EsxCli -vmhost $HBAHost
          $fcfirmware = $esxcli.storage.san.fc.list()
          $driverhbasoft = $esxcli.software.vib.list()

         foreach($fcfirmwareselect in $fcfirmware){
            $fcDescription = $fcfirmwareselect.ModelDescription
            $fcDriver = $driversoft | where {$_.name -eq ($fcfirmwareselect.DriverName)}
            $fcName = $fcfirmwareselect.Adapter
            $fcFirmware = $fcfirmwareselect.FirmwareVersion
            $OutputHBA += '' | select @{N='Hostname';E={$HBAHost.Name}},
                                      @{N='ESXi Model';E={$HBAHost.Model}},
                                      @{N='Driver Ver.';E={$fcDriver.Version}},
                                      @{N='Firmware Ver.';E={$fcFirmware}},
                                      @{N='HBA Descr.';E={$fcDescription}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'HBA Info'
     }
    }

    Write-Host "($CurrentStep/$TotalSteps) HBA Info Completed."
    $CurrentStep ++


    ######## HBA WWN/WWNP output ###########################################################################################################
  
    Write-Host "($CurrentStep/$TotalSteps) Gathering WWN/WWPN Info..."
   
    Get-VMhost | Get-VMHostHBA -Type FibreChannel | Select VMHost,
                                                           Device,
                                                           @{N='Status';E={$_.Status}},
                                                           @{N='Driver';E={$_.Driver}},
                                                           @{N='Model';E={$_.Model}},
                                                           @{N='WWN';E={'{0:X}'-f$_.NodeWorldWideName}},
                                                           @{N='WWP';E={'{0:X}'-f$_.PortWorldWideName}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'WWN_WWPN Info'

    Write-Host "($CurrentStep/$TotalSteps) WWN/WWPN Info Completed."
    $CurrentStep ++


    ######## ATS Heart Beat  ###########################################################################################################
  
    Write-Host "($CurrentStep/$TotalSteps) Gathering ATS Info..."

    Get-VMHost | Get-AdvancedSetting -Name VMFS3.UseATSForHBOnVMFS5 | select @{N='Host';E={$_.Entity}},
                                                                             @{N='ATS Heart Beat';E={$_.Name}}, 
                                                                             @{N='0=Disabled/1=Enabled';E={$_.Value}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'ATS HB Info'

    Write-Host "($CurrentStep/$TotalSteps) ATS Info Completed."
    $CurrentStep ++


    ######## VMKernel Adapterstput #####################################################################################################
 
    Write-Host "($CurrentStep/$TotalSteps) Gathering VMKernel Info..."

    Get-VMHostNetworkAdapter -VMKernel | select @{N='ESXi Host';E={$_.VMHost}},
                                                 @{N='Device Name';E={$_.DeviceName}},
                                                 @{N='Mac Address';E={$_.Mac}},
                                                 @{N='IP Address';E={$_.IP}},
                                                 @{N='SubNet Mask';E={$_.SubnetMask}},
                                                 @{N='Port Group';E={$_.PortGroupName}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'VMKernel Adapters'
     
    Write-Host "($CurrentStep/$TotalSteps) VMKernel Info Completed." 
    $CurrentStep ++                                             


    ######## Distributed Switches output ###########################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering vDS Info..."

    Get-VDPortgroup  | select @{N='Datacenter';E={$_.Datacenter}},
                              @{N='vLan Name';E={$_.Name}},
                              @{N='vSwitch Name';E={$_.VDSwitch}},
                              @{N='Vlan ID';E={$_.VlanConfiguration}},
                              @{N='Number of ports';E={$_.NumPorts}},
                              @{N='Promiscuous Mode';E={$_.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.AllowPromiscuous.Value}},
                              @{N='Forged Transmits';E={$_.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.ForgedTransmits.Value}}, 
                              @{N='Mac Changes';E={$_.Extensiondata.Config.DefaultPortConfig.SecurityPolicy.MacChanges.Value}} | Sort-Object Datacenter,'vSwitch Name' |  Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'vDS'
    
    Write-Host "($CurrentStep/$TotalSteps) vDS Info Completed."
    $CurrentStep ++                             


    ######## Nics Assign to Distributed Switches output ################################################################################
   
    Write-Host "($CurrentStep/$TotalSteps) Gathering vDS Nics Info..."
    
    $Nic_vDS = @()
    foreach($sw in (Get-VirtualSwitch -Distributed)){
         $uuid = $sw.ExtensionData.Summary.Uuid
         $sw.ExtensionData.Config.Host | %{
            $esx = Get-View $_.Config.Host
            $netSys = Get-View $esx.ConfigManager.NetworkSystem
            $netSys.NetworkConfig.ProxySwitch | where {$_.Uuid -eq $uuid} | %{
                $_.Spec.Backing.PnicSpec | %{
                    $row = '' | Select Host,dvSwitch,PNic
                    $row.Host = $esx.Name
                    $row.dvSwitch = $sw.Name
                    $row.PNic = $_.PnicDevice
                    $Nic_vDS += $row      
                 }
           }
        }
    }
    $Nic_vDS | Sort-Object Host | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Nics vDS'
   
    Write-Host "($CurrentStep/$TotalSteps) vDS Nics Info Completed."
    $CurrentStep ++



    ######## Standard Switches output ########################################################################################
  
    Write-Host "($CurrentStep/$TotalSteps) Gathering vSS Info..."

    $vmhosts_vss = Get-VMHost
    foreach ($ESXHost in $vmhosts_vss) {
          Get-VirtualPortGroup -Standard -VMHost $ESXHost | select @{N='Cluster';E={$ESXHost.Parent}},
                                                                   @{N='Host';E={$ESXHost.Name}},
                                                                   VirtualSwitch,
                                                                   @{N='Port Name';E={$_.Name}},
                                                                   @{N='Vlan';E={$_.vlanid}} | Sort-Object Host,VirstualSwitch | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'vSS'                                                                                                           
    }

    Write-Host "($CurrentStep/$TotalSteps) vSS Info Completed."
    $CurrentStep ++


    ######## Nics Assign to Standard Switches output ################################################################################
    
    Write-Host "($CurrentStep/$TotalSteps) vSS Nics Info..."

    foreach($esx in Get-VMHost){
        $vNicTab = @{}
        $esx.ExtensionData.Config.Network.Vnic | %{
        $vNicTab.Add($_.Portgroup,$_)
        }
        foreach($vsw in (Get-VirtualSwitch -Standard -VMHost $esx)){
            foreach($pg in (Get-VirtualPortGroup -VirtualSwitch $vsw)){
                 Select -InputObject $pg -Property @{N='ESX';E={$esx.name}},
                                                   @{N='vSwitch';E={$vsw.Name}},
                                                   @{N='Active NIC';E={[string]::Join(',',$vsw.ExtensionData.Spec.Policy.NicTeaming.NicOrder.ActiveNic)}},
                                                   @{N='Portgroup';E={$pg.Name}},
                                                   @{N='VLAN';E={$pg.VLanId}} | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Nics vSS'
            }
        }
    }

    Write-Host "($CurrentStep/$TotalSteps) vSS Nics Info Completed."
    $CurrentStep ++

    
    ######## Licenses output ####################################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering Licenses Info..."

        $LINfo = @()
        foreach ($licenseManager in (Get-View LicenseManager)){
        $vCenter = ([System.uri]$licenseManager.Client.ServiceUrl).Host
        foreach ($license in $licenseManager.Licenses)
        {
            $licenseProp = $license.Properties
            $licenseExpiryInfo = $licenseProp | Where-Object {$_.Key -eq 'expirationDate'} | Select-Object -ExpandProperty Value
            if ($license.Name -eq 'Product Evaluation')
            {
                $expirationDate = 'Evaluation'
            }
            elseif ($null -eq $licenseExpiryInfo)
            {
                $expirationDate = 'Never'
            }
            else
            {
                $expirationDate = $licenseExpiryInfo
            } 
    
            if ($license.Total -eq 0)
            {
                $totalLicenses = 'Unlimited'
            } 
            else 
            {
                $totalLicenses = $license.Total
            }     
            $licenseObj = New-Object psobject
            $licenseObj | Add-Member -MemberType NoteProperty -Name 'Name'  -Value $license.Name
            $licenseObj | Add-Member -MemberType NoteProperty -Name 'License Key' -Value $license.LicenseKey
            $licenseObj | Add-Member -MemberType NoteProperty -Name 'Product Name' -Value ($licenseProp | Where-Object {$_.Key -eq 'ProductName'} | Select-Object -ExpandProperty Value)
            $licenseObj | Add-Member -MemberType NoteProperty -Name 'Version' -Value ($licenseProp | Where-Object {$_.Key -eq 'ProductVersion'} | Select-Object -ExpandProperty Value)
            $licenseObj | Add-Member -MemberType NoteProperty -Name 'Total' -Value $totalLicenses
            $licenseObj | Add-Member -MemberType NoteProperty -Name 'Used' -Value $license.Used
            $licenseObj | Add-Member -MemberType NoteProperty -Name 'Expiration   Date' -Value $expirationDate
            $LINfo += $licenseObj 
            }
         } 
         $LINfo | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Licenses'
         
     Write-Host "($CurrentStep/$TotalSteps) Licenses Info Completed."          
     $CurrentStep ++

    
     ######## Snapshots output ###################################################################################################

     Write-Host "($CurrentStep/$TotalSteps) Gathering Snapshot Info..."
     
     Get-VM | Sort Name | Get-Snapshot | Where { $_.Name.Length -gt 0 } | Select VM,
                                                                                Name,
                                                                                @{N='SnapShot  Created  on';E={@($_.created)}},
                                                                                Description,
                                                                                @{N='SizeGB';E={[math]::Round(($_.SizeMB/1024),2)}}  | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Snapshots'
     
     Write-Host "($CurrentStep/$TotalSteps) Snapshot Info Completed." 
     $CurrentStep ++

                                                                                
    ######## Zombie VMDKs output ####################################################################################################

    Write-Host "($CurrentStep/$TotalSteps) Gathering Zombie Disk Info..."

    $report = @()
    $arrUsedDisks = Get-View -ViewType VirtualMachine | % {$_.Layout} | % {$_.Disk} | % {$_.DiskFile}
    $arrDS = Get-Datastore | Sort-Object -property Name
    foreach ($strDatastore in $arrDS) {
          $ds = Get-Datastore -Name $strDatastore.Name | % {Get-View $_.Id}
          $fileQueryFlags = New-Object VMware.Vim.FileQueryFlags
          $fileQueryFlags.FileSize = $true
          $fileQueryFlags.FileType = $true
          $fileQueryFlags.Modification = $true
          $searchSpec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec
          $searchSpec.details = $fileQueryFlags
          $searchSpec.matchPattern = '*.vmdk'
          $searchSpec.sortFoldersFirst = $true
          $dsBrowser = Get-View $ds.browser
          $rootPath = '[' + $ds.Name + ']'
          $searchResult = $dsBrowser.SearchDatastoreSubFolders($rootPath, $searchSpec)

          foreach ($folders in $searchResult){
            foreach ($fileResult in $folders.File){
            if ($fileResult.Path){
             if ($fileResult.Path -notmatch '-flat.vmdk|-ctk.vmdk|-delta.vmdk|-rdmp.vmdk' -and
                    (-not ($arrUsedDisks -contains ($folders.FolderPath + $fileResult.Path)))){
                    $row = '' | Select @{N='Cluster';E={$strDatastore.Name}},
                                       @{N='Path';E={$folders.FolderPath}},
                                       @{N='File';E={$fileResult.Path}},
                                       @{N='Size in GB';E={[math]::round($fileResult.FileSize/(1024 * 1024 * 1024),3)}},
                                       @{N='Date Last modifided';E={$fileResult.Modification}}
                                                                         
                   $report += $row
               }
            }
         }
        }
       } $report | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Zombie Disks'
   
    Write-Host "($CurrentStep/$TotalSteps) Zombi Disk Info Completed."
    $CurrentStep ++  

    ######## MetaData ####################################################################################################
    
    Write-Host "Writing Metadata..."

    $Infov =' Report Tools version : '+$DMTLversion
    $Note =' Scan performed on vCenter '+$vCenter+' on '+$Date
    $Note,$Infov | Export-Excel -Append -AutoSize –Path $outputpath -WorksheetName 'Meta Data'
   
    Write-Host "Report Completed."


    ############ Disconnect from vCenter ######################################################################################## 
    
     if (!($Connect)) {
           Disconnect-VIServer -confirm:$false
           Write-Host 'Disconnecting from vCenter.'
     }   
   
     ## Reset steps to 1
     $CurrentStep = 1
   
   } ######## End of all exports ###################################################################################################

     
    ###### Check if Archive folder exists. If not it will create it ################################################################
    
    if (-Not (Test-Path -Path 'C:\Scripts\DimiTools\Archives')) {
        New-Item -Path 'C:\Scripts\DimiTools\Archives' -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    }
    
    ## Compress the folder and place it to Archive folder. Delete the original folder ###############################################

      Start-Sleep -s 10
      Compress-Archive -Path $source -DestinationPath $destination
      Start-Sleep -s 10
      Remove-Item -path $source -recurse


    ######## Sent Report via Mail ####################################################################################################
    if ($sent_mail){
       Write-Host 'Sending mail...'
       send-mailmessage -from 'NEW_TOOLS@KYNDRYL.COM' -to 'dimitrios.kakoulidis@kyndryl.com' -subject "Customer Name: Weekly Tools report $(get-date -f 'dd-MM-yyyy')" -body 'Below you can find the Reporting Tools excel file. Please see attachment ' -Attachments $destination -smtpServer  192.168.21.16
    } 

    ###########################################################################################################################################
    ######################## End of Script ####################################################################################################
    ###########################################################################################################################################
