<#-----------------------------------------------------------
VMware vSphere Data Dump
Created by Jason Chandler
10/15/2018

Edited
11/1/2018

Directions:
Create text file "C:\Scripts\VCenter Script Data\VMwareMasterDataDump_vCenterArr.txt" which contains the vCenters you want to connect to.  This script will generate various content from VM Guests, and ESXi Hosts and send it to C:\Reports\
Add the following Folders to this server "C:\Windows\System32\config\systemprofile\Desktop" & "C:\Windows\SysWOW64\config\systemprofile\Desktop"
------------------------------------------------------------#>

Start-Transcript -Path "C:\Scripts\VMwareDataDump\transcript_VMwareMasterDataDump.log"
Write-Output $((Get-Date).ToString('MM-dd-yyyy_hh:mm:ss'))"    : *Started* download of VMware and Host Data and Excel Upload" | Out-File "\\hq-nas-projects\departments\IT\Private\VMware\VMwareDataDump\log.txt" -Append
Import-Module VMware.VimAutomation.Core
#NULL VARIABLES
$vms = $null
$ESXiHosts = $null
$Object = $null

#closes all open Excel processes
Stop-Process -n excel -Force

sleep -Seconds 5

#Move VMwareMasterDataDump to local computer
Move-Item -Path "\\hq-nas-projects\departments\IT\Private\VMware\VMwareDataDump\VMwareMasterDataDump.xlsx" -Destination "C:\Scripts\VMwareDataDump\VMwareMasterDataDump.xlsx"

#Remove CSV if present
if ((Test-Path -path "C:\Scripts\VMwareDataDump\HostDataDump.csv") -eq $true) { Remove-Item -Path "C:\Scripts\VMwareDataDump\HostDataDumpTemp.csv"} else { Write-Host "HostDataDumpTemp.csv is already gone"}
if ((Test-Path -path "C:\Scripts\VMwareDataDump\VMDataDumpTemp.csv") -eq $true) { Remove-Item -Path "C:\Scripts\VMwareDataDump\VMDataDumpTemp.csv"} else { Write-Host "VMDataDumpTemp.csv is already gone"}

# Auth variables
$vcvCreds = "C:\Vcv\vcitals.xml"
#$vcvCreds = "C:\Scripts\VMwareDataDump\credtemp.xml"
$creds = Get-VICredentialStoreItem -file $vcvCreds

#List vCenters within txt file
$vCenterArr = @(get-Content "C:\Scripts\VMwareDataDump\VMwareMasterDataDump_vCenterArr.txt")

#Connect to vCenters within VMwareMasterDataDump_vCenterArr.txt and create variables
Connect-VIServer -Server $vCenterArr -User $creds.User -Password $creds.Password 
$vms = Get-VM | Sort-Object Name
$ESXiHosts = Get-VMHost | Sort-Object Name

#New Objects & VI Properties
New-VIProperty -Name "VMTag" -ObjectType VirtualMachine -Value {$(Get-TagAssignment -Entity $args[0] | Select-Object -ExpandProperty Tag).Name}
New-VIProperty -Name "HostTag" -ObjectType VMHost -Value {$(Get-TagAssignment -Entity $args[0] | Select-Object -ExpandProperty Tag).Name}
New-VIProperty -Name "ClusterName" -ObjectType VirtualMachine -Value {Get-Cluster -VM $($args[0].Name) | Select-Object -ExpandProperty Name}

#Collecting Host data
$counter1 = 0
Foreach ($ESXiHost in $ESXiHosts) {
    $Object = Get-VMHost $ESXiHost | Select-Object Name,HostTag,
    @{N="vCenter Server";E={$ESXiHost.ExtensionData.Client.ServiceUrl.Split('/')[2].trimend(":443")}},
    @{N="Parent Cluster";E={$ESXiHost.Parent}},
    @{N="VMware Version";E={$ESXiHost.APIVersion + " " + $ESXiHost.build}},
    @{N="Hardware Model";E={$ESXiHost.Model}},
    @{N="CPU Socket(s)";E={$ESXiHost.ExtensionData.Hardware.CpuInfo.NumCpuPackages}},
    @{N="CPU Cores";E={$ESXiHost.ExtensionData.Hardware.CpuInfo.NumCpuCores}},
    @{N="CPU Threads";E={$ESXiHost.ExtensionData.Hardware.CpuInfo.NumCpuThreads}},
    @{N="Memory (GB)";E={"" + [math]::round($ESXiHost.ExtensionData.Hardware.MemorySize / 1GB, 0)}}
    $counter1 +=1
    $Object | Export-Csv -path "C:\Scripts\VMwareDataDump\HostDataDumpTemp.csv" -NoTypeInformation -Append
    Write-Host "Gathered $counter1  " " $ESXiHost Information"
}

#Collecting VM data
$counter1 = 0
Foreach ($vm in $vms) {
    $Object = Get-VM $vm| Select-Object Name,VMTag,
    @{N="vCenter Server"; E={$vm.extensiondata.Client.ServiceUrl.Split('/')[2].trimend(":443")}},
    @{N="Parent Cluster"; E={$_.ClusterName}},
    @{N="Parent Host"; E={$vm.VMHost.Name}},
    @{N="OS";E={$vm.ExtensionData.Guest.GuestFullName}},
    @{N="PowerState";E={$vm.PowerState}},
    @{N="Department";E={($vm.CustomFields | Where-Object {$_.Key -eq "Department:"}).Value}},
    @{N="Supporting Application";E={($vm.CustomFields | Where-Object {$_.Key -eq "Supporting Application:"}).Value}},
    @{N="Technical Lead";E={($vm.CustomFields | Where-Object {$_.Key -eq "Technical Lead:"}).Value}},
    @{N="Number of CPU(s)";E={$vm.NumCpu}},
    @{N="Memory (GB)";E={$vm.MemoryGB}},
    @{N="Provisioned Space (GB)";E={$vm.ProvisionedSpaceGB}},
    @{N="Used Space (GB)";E={$vm.UsedSpaceGB}}
    $counter1 +=1
    Write-Host "Gathered $counter1 " " $vm Information"
    $Object | Export-Csv -path "C:\Scripts\VMwareDataDump\VMDataDumpTemp.csv" -NoTypeInformation -Append
    }

#Remove VIProperty's
Remove-VIProperty -Name *VMTag* -ObjectType *
Remove-VIProperty -Name *HostTag* -ObjectType *
Remove-VIProperty -Name *ClusterName* -ObjectType *

#Disconnect all vCenters
Disconnect-VIServer * -Force -Confirm:$false

#Excel Work

#Define Variables
$VMDataDump = 'C:\Scripts\VMwareDataDump\VMDataDumpTemp.csv' # source's fullpath
$HostDataDump = 'C:\Scripts\VMwareDataDump\HostDataDumpTemp.csv' # source's fullpath
$VMwareMasterDataDump = 'C:\Scripts\VMwareDataDump\VMwareMasterDataDump.xlsx' # destination's fullpath
$CurrentDate = "Latest Upload " + $((Get-Date).ToString('MM-dd-yyyy hh:mm:ss'))

#####################    VM Data     #####################

#Clear contents of VMDataDump Worksheet
$xl = new-object -c excel.application
$xl.displayAlerts = $false # don't prompt the user
$wb2 = $xl.workbooks.open($VMwareMasterDataDump) # open target
$sh2_wb2 = $wb2.sheets | Where-Object {$_.name -eq "VMDataDump"}
$sh2_wb2.Range("A1:M900").Clear() # clear contents in range
$wb2.close($true) # close and save destination workbook 
$xl.Quit()
Stop-Process -n excel -Force

#Copy VMDataDumpTemp Worksheet to VMwareMasterDataDump Workbook
$xl = new-object -c excel.application
$xl.displayAlerts = $false # don't prompt the user
$wb2 = $xl.workbooks.open($VMDataDump, $null, $true) # open source, readonly
$wb1 = $xl.workbooks.open($VMwareMasterDataDump) # open target
$sh1_wb1 = $wb1.sheets.item(1) # second sheet in destination workbook
$sheetToCopy = $wb2.sheets.item('VMDataDumpTemp') # source sheet to copy
$sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook
$wb2.close($false) # close source workbook w/o saving
$wb1.close($true) # close and save destination workbook
$xl.quit()
Stop-Process -n excel -Force

#Copy VMDataDumpTemp Contents into VMDataDump Worksheet
$xl = New-Object -ComObject excel.application 
$xl.displayAlerts = $false # don't prompt the user
$Workbook = $xl.Workbooks.open($VMwareMasterDataDump) 
$Worksheet = $Workbook.WorkSheets.item("VMDataDumpTemp") 
$worksheet.activate()
$range = $WorkSheet.Range("A1:M1").EntireColumn
$range.Copy() | out-null
$Worksheet = $Workbook.Worksheets.item("VMDataDump") 
$Range = $Worksheet.Range("A1")
$Worksheet.Paste($range)  
$workbook.UsedRange.Columns.Autofit() | Out-Null #Autosize Columns
$workbook.close($true) # close and save destination workbook
$xl.Quit()

#Delete VMDataDumpTemp Worksheet
$xl = new-object -c excel.application
$xl.displayAlerts = $false # don't prompt the user
$wb2 = $xl.workbooks.open($VMwareMasterDataDump) # open target
$sh2_wb2 = $wb2.sheets | Where-Object {$_.name -eq "VMDataDumpTemp"}
$sh2_wb2.delete() #Delete original sheet in template
$wb2.close($true) # close and save destination workbook
$xl.quit()
Stop-Process -n excel -Force


#####################    Host Data     #####################

#Clear contents of HostDataDump Worksheet
$xl = new-object -c excel.application
$xl.displayAlerts = $false # don't prompt the user
$wb2 = $xl.workbooks.open($VMwareMasterDataDump) # open target
$sh2_wb2 = $wb2.sheets | Where-Object {$_.name -eq "HostDataDump"}
$sh2_wb2.Range("A1:J200").Clear()
$wb2.close($true) # close and save destination workbook 
$xl.Quit()
Stop-Process -n excel -Force

#Copy HostDataDumpTemp Worksheet to VMwareMasterDataDump Workbook
$xl = new-object -c excel.application
$xl.displayAlerts = $false # don't prompt the user
$wb2 = $xl.workbooks.open($HostDataDump, $null, $true) # open source, readonly
$wb1 = $xl.workbooks.open($VMwareMasterDataDump) # open target
$sh1_wb1 = $wb1.sheets.item(1) # second sheet in destination workbook
$sheetToCopy = $wb2.sheets.item('HostDataDumpTemp') # source sheet to copy
$sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook
$wb2.close($false) # close source workbook w/o saving
$wb1.close($true) # close and save destination workbook
$xl.quit()
Stop-Process -n excel -Force

#Copy HostDataDumpTemp Contents into VMDataDump Worksheet
$xl = New-Object -ComObject excel.application 
$xl.displayAlerts = $false # don't prompt the user
$Workbook = $xl.Workbooks.open($VMwareMasterDataDump) 
$Worksheet = $Workbook.WorkSheets.item("HostDataDumpTemp") 
$worksheet.activate()
$range = $WorkSheet.Range("A1:M1").EntireColumn
$range.Copy() | out-null
$Worksheet = $Workbook.Worksheets.item("HostDataDump") 
$Range = $Worksheet.Range("A1")
$Worksheet.Paste($range) 
$workbook.UsedRange.Columns.Autofit() | Out-Null #Autosize Columns
$workbook.close($true) # close and save destination workbook
$xl.Quit()

#Delete HostDataDumpTemp Worksheet
$xl = new-object -c excel.application
$xl.displayAlerts = $false # don't prompt the user
$wb2 = $xl.workbooks.open($VMwareMasterDataDump) # open target
$sh2_wb2 = $wb2.sheets | Where-Object {$_.name -eq "HostDataDumpTemp"}
$sh2_wb2.delete() #Delete original sheet in template
$wb2.close($true) # close and save destination workbook
$xl.quit()
Stop-Process -n excel -Force

#####################


#Delete CSV's
Remove-Item -Path "C:\Scripts\VMwareDataDump\HostDataDumpTemp.csv"
Remove-Item -Path "C:\Scripts\VMwareDataDump\VMDataDumpTemp.csv"
Write-Output $((Get-Date).ToString('MM-dd-yyyy_hh:mm:ss'))"    : *Completed* download of VMware and Host Data and Excel Upload" | Out-File "\\hq-nas-projects\departments\IT\Private\VMware\VMwareDataDump\log.txt" -Append

#Move VMwareMasterDataDump to projects share
Move-Item -Path "C:\Scripts\VMwareDataDump\VMwareMasterDataDump.xlsx" -Destination "\\hq-nas-projects\departments\IT\Private\VMware\VMwareDataDump\VMwareMasterDataDump.xlsx"
sleep -Seconds 5

#Insert Date into .xlsx
$xldate = new-object -c excel.application
$xldate.displayAlerts = $false # don't prompt the user
$wbdate = $xldate.workbooks.open("\\hq-nas-projects\departments\IT\Private\VMware\VMwareDataDump\VMwareMasterDataDump.xlsx")
$wsdate = $wbdate.WorkSheets.item(3)
$wbdate.UsedRange.Columns.Autofit() | Out-Null #Autosize Columns
$xldate.Visible=$true
$wsdate.Cells.Item(1,6) = $currentDate
$wbdate.close($true)
$xldate.quit()
Stop-Process -n excel -Force


Stop-Transcript 
