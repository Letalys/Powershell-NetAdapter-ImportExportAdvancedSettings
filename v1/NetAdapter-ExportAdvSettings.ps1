#Requires -Version 4.0
<#
.SYNOPSIS
  This script can be use for export NetAdapter AdvancedSetting to XML format
.DESCRIPTION
  Using NetAdapter Module, create a exportable object to XML. Generated XML can be import with the script NetAdapter-ImportAdvSettings.ps1.
  Must be used in console UI.
.OUTPUTS
  .\Export\<InterfaceDescription>.NetAdapterExport.xml
.NOTES
  Version:        1.0
  Author:         Letalys
  Creation Date:  19/02/2023
  Purpose/Change: Initial script development
#>

cls
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host -ForegroundColor Cyan "       NetAdapter Advanced Settings Exporting           "
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host ""

Write-Host -Foreground Yellow "Physical Network Adapter List :"
Get-NetAdapter -Physical | Select @{label='Index';expression={$_.IfIndex}},Status,InterfaceDescription,Name, DriverFileName, DriverVersion, DriverDate  | Format-Table -Autosize

$SelectedNetAdapterIndex = Read-Host "Enter the index of the network adapter to export"

Write-host ""
Write-Host -NoNewLine  -ForegroundColor Yellow "Export..."

$SelectedNetAdapter = Get-NetAdapter -InterfaceIndex $SelectedNetAdapterIndex -ErrorAction SilentlyContinue

if ($SelectedNetAdapter -ne $null){
    Try {
        $WMIComputerInfo = Get-WmiObject -Class  Win32_ComputerSystemProduct
        $WMIOSInfo = Get-WmiObject -Class  Win32_OperatingSystem
        $SelectedNetAdapterAdvSettings = $SelectedNetAdapter | Get-NetAdapterAdvancedProperty  | Select * | Sort DisplayName

        $ObjExport = New-Object PSObject

        $ObjComputerInfo = New-Object PSObject
        $ObjComputerInfo | Add-Member -MemberType NoteProperty -Name "ExportDateTime" -Value $(Get-Date -Format "yyyy-MM-dd HH:mm")
        $ObjComputerInfo | Add-Member -MemberType NoteProperty -Name "ComputerSource" -Value $env:COMPUTERNAME
        $ObjComputerInfo | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value $WMIComputerInfo.Vendor
        $ObjComputerInfo | Add-Member -MemberType NoteProperty -Name "ModelName" -Value $WMIComputerInfo.Name
        $ObjComputerInfo | Add-Member -MemberType NoteProperty -Name "ModelVersion" -Value $WMIComputerInfo.Version
        $ObjComputerInfo | Add-Member -MemberType NoteProperty -Name "SerialNumber" -Value $WMIComputerInfo.IdentifyingNumber
        $ObjComputerInfo | Add-Member -MemberType NoteProperty -Name "OperatingSystemVersion" -Value $WMIOSInfo.Version

        $ObjExport | Add-Member -MemberType NoteProperty -Name "ExportedFrom" -Value $ObjComputerInfo

        $ObjNetAdapterInfo = New-Object PSObject
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "InterfaceDescription" -Value $SelectedNetAdapter.InterfaceDescription
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "DriverFileName" -Value $SelectedNetAdapter.DriverFileName
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "DriverVersion" -Value $SelectedNetAdapter.DriverVersion
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "DriverDate" -Value $SelectedNetAdapter.DriverDate

        [System.Collections.Arraylist]$SettingsList =@()
        foreach($Setting in $SelectedNetAdapterAdvSettings){
            $ObjNetAdapterAdvSettingValueInfo = New-Object PSObject
            $ObjNetAdapterAdvSettingValueInfo | Add-Member -MemberType NoteProperty -Name "Value" -Value "$($setting.DisplayValue)"

            $SelectNetAdapterValidDisplayAdvSettings = $Setting | Select -ExpandProperty ValidDisplayValues

            [System.Collections.Arraylist]$ValidDisplayValues =@()
            foreach($ValidDisplayValue in $SelectNetAdapterValidDisplayAdvSettings){
                $ValidDisplayValues.Add($ValidDisplayValue) | Out-Null
            }

            $ObjNetAdapterAdvSettingValueInfo | Add-Member -MemberType NoteProperty -Name "ValidDisplayValues" -Value $ValidDisplayValues

            $ObjNetAdapterAdvSetting = New-Object PSObject
            $ObjNetAdapterAdvSetting | Add-Member -MemberType NoteProperty -Name "$($setting.RegistryKeyWord)" -Value $ObjNetAdapterAdvSettingValueInfo

            $SettingsList.Add($ObjNetAdapterAdvSetting) | Out-Null
        }

        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "AdvancedSettings" -Value $SettingsList
        $ObjExport | Add-Member -MemberType NoteProperty -Name "NetAdapterInformations" -Value $ObjNetAdapterInfo

        $PathToSave = "$PSScriptRoot\Exports\$($SelectedNetAdapter.InterfaceDescription).NetAdapteExportCfg.xml"
        $NewXML = ConvertTo-Xml -As "Document" -InputObject $ObjExport -Depth 5 -NoTypeInformation
        $NewXML.Save($PathToSave)

        Write-Host -ForegroundColor Green "Completed"
        Write-Host -NoNewLine -ForegroundColor Yellow "Save To : "
        Write-Host $PathToSave 
    }Catch{
        Write-Error $_
    }
}else{
    Write-Error "No NetworkAdapter exists for the specified index." -Category ObjectNotFound
}
Write-Host ""
