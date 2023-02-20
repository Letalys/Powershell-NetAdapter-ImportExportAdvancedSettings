<#
.SYNOPSIS
  This script can be use for export NetAdapter AdvancedSetting to XML format
.DESCRIPTION
  Using WMI and registry, create a exportable object to XML. Generated XML can be import with the script NetAdapter-ImportAdvSettings.ps1.
  Must be used in console UI.
.OUTPUTS
  .\Export\<InterfaceDescription>.NetAdapterExport.xml
.NOTES
  Version:        2.0
  Author:         Letalys
  Creation Date:  19/02/2023
  Purpose/Change: Initial script development
#>

Clear-Host
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host -ForegroundColor Cyan "        NetAdapter Advanced Settings Exporting    v2        "
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host ""

#region Variables
    $NetClassRegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"
#endregion Variables

Write-Host -Foreground Yellow "Physical Network Adapter List :"
Get-CimInstance Win32_NetworkAdapter -Filter "PhysicalAdapter='True'" | Select-Object @{label='Index';expression={$_.DeviceID}},Name, ServiceName | Format-Table -AutoSize

$SelectedNetAdapterIndex = Read-Host "Enter the index of the network adapter to export"

Write-host ""
Write-Host -NoNewLine  -ForegroundColor Yellow "Export..."

$SelectedNetAdapter = Get-CIMInstance -Class Win32_NetworkAdapter -Filter "DeviceID=$SelectedNetAdapterIndex AND PhysicalAdapter='True'"

if($null -ne $SelectedNetAdapter){
    Try{
        $WMIComputerInfo = Get-CimInstance -Class  Win32_ComputerSystemProduct
        $WMIOSInfo = Get-CimInstance -Class  Win32_OperatingSystem
        $SelectNedAdpterDriver = Get-CimInstance -Class Win32_PnpSignedDriver -Filter "DeviceName='$($SelectedNetAdapter.Name)'"

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
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "InterfaceDescription" -Value $SelectedNetAdapter.Name
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "DriverFileName" -Value $SelectNedAdpterDriver.InfName
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "DriverVersion" -Value $SelectNedAdpterDriver.DriverVersion
        #$ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "DriverDate" -Value ($SelectNedAdpterDriver.DriverDate.ToString().Substring(0,4) + "-" + $SelectNedAdpterDriver.DriverDate.Substring(4,2) + "-" + $SelectNedAdpterDriver.DriverDate.Substring(6,2))
        $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "DriverDate" -Value (Get-Date -Date $SelectNedAdpterDriver.DriverDate -Format "yyyy-MM-dd")

        #Set the real Registry Index (4 digit)
        While($SelectedNetAdapterIndex.Length -ne 4){
            $SelectedNetAdapterIndex = "0$SelectedNetAdapterIndex"
        }

        $NetAdapterRegistryPath = Join-Path -Path $NetClassRegistryPath -ChildPath $SelectedNetAdapterIndex
        $NetAdapterAdvSettingsRegistryPath = Join-Path -Path $NetAdapterRegistryPath -ChildPath "Ndi\Params"
        $NetAdapterAdvSettingsItems = Get-ChildItem -Path $NetAdapterAdvSettingsRegistryPath

        [System.Collections.Arraylist]$SettingsList =@()
        Foreach($Setting in $NetAdapterAdvSettingsItems){
            $SettingRegistryPath = $Setting.PSPath

            $ObjNetAdapterAdvSettingValueInfo = New-Object PSObject
            $ObjNetAdapterAdvSettingValueInfo | Add-Member -Name "Description" -membertype Noteproperty -Value $((Get-ItemProperty -Path $SettingRegistryPath).ParamDesc)
            $ObjNetAdapterAdvSettingValueInfo | Add-Member -MemberType NoteProperty -Name "Value" -Value "$($(Get-ItemProperty -Path $NetAdapterRegistryPath).$($setting.PSChildName))"

            if(Test-Path (Join-Path -Path $SettingRegistryPath -ChildPath "Enum")){
                $ValidRegistryValues = Get-Item -Path (Join-Path -Path $SettingRegistryPath -ChildPath "Enum") | Select-object -ExpandProperty Property
                $ObjValidValues = New-Object PSObject
                Foreach($vrv in $ValidRegistryValues){
                    $ValidDisplayValue = Get-ItemProperty -Path (Join-Path -Path $SettingRegistryPath -ChildPath "Enum")
                    $ObjValidValues | Add-Member -MemberType NoteProperty -Name "$vrv" -Value $ValidDisplayValue.$vrv
                }
            }
            
            $ObjNetAdapterAdvSettingValueInfo | Add-Member -MemberType NoteProperty -Name "ValidDisplayValues" -Value $ObjValidValues

            $ObjNetAdapterAdvSetting = New-Object PSObject
            $ObjNetAdapterAdvSetting | Add-Member -MemberType NoteProperty -Name "$($setting.PSChildName)" -Value $ObjNetAdapterAdvSettingValueInfo

            $SettingsList.Add($ObjNetAdapterAdvSetting) | Out-Null
         }

         $ObjNetAdapterInfo | Add-Member -MemberType NoteProperty -Name "AdvancedSettings" -Value $SettingsList
         $ObjExport | Add-Member -MemberType NoteProperty -Name "NetAdapterInformations" -Value $ObjNetAdapterInfo

         $PathToSave = "$PSScriptRoot\Exports\$($SelectedNetAdapter.Name).NetAdapteExportCfg.xml"
         $NewXML = ConvertTo-Xml -As "Document" -InputObject $ObjExport -Depth 5 -NoTypeInformation
         $NewXML.Save($PathToSave)
 
         Write-Host -ForegroundColor Green "Completed"
         Write-Host -NoNewLine -ForegroundColor Yellow "Save To : "
         Write-Host $PathToSave 
    }Catch{
        Write-Error  "$($_.InvocationInfo.ScriptLineNumber) : $($_)"
    }
}else{
     Write-Error  "NetworkAdapter Index not found"
}