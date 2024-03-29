﻿#Requires -Version 4.0
<#
.SYNOPSIS
  This script can be use for Import NetAdapter AdvancedSetting from XML format
.DESCRIPTION
  Using NetAdapter Module, create a import object from XML. XML is generated by the script NetAdapter-ExportAdvSettings.ps1.
  Must be used in console UI. Need administrator rights.
.PARAMETER ImportXMLFile
    Specifies the XML file path to import. If this parameter is not configured, an open file window will be displayed.
    If it's set, you can use this script in silent mode.
               Required?                    false
               Default value                null
               Accept pipeline input?       false
               Accept wildcard characters?  false
.NOTES
  Version:        1.0
  Author:         Letalys
  Creation Date:  19/02/2023
  Purpose/Change: Initial script development
#>
Param([String]$ImportXMLFile)

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $false){
    Write-Error "This script requires elevation of privilege to run" -Category AuthenticationError
    pause
    exit
}

Clear-Host
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host -ForegroundColor Cyan "       NetAdapter Advanced Settings Importing   v1          "
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host ""

if($ImportXMLFile -eq ""){
    Add-Type -AssemblyName System.Windows.Forms

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        Filter = 'XML (*.xml)|*.xml'
        RestoreDirectory = $true
        Title ="Choose the configuration file to import"
        Multiselect =$false
    }
    $null = $FileBrowser.ShowDialog()
    $ImportXMLFile = $FileBrowser.FileName
}

If((Test-Path $ImportXMLFile) -eq $false){Write-Error "Specified file not exist" -Category NotSpecified;exit}

Try {
    Write-Host -NoNewLine -ForegroundColor Yellow "Loading File... "
    $CurrentFile = Get-Item $ImportXMLFile -ErrorAction SilentlyContinue
    
    [xml]$XML = Get-Content $CurrentFile.FullName
    Write-Host -ForegroundColor Green "Completed"

    $NetAdapterIfName = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'InterfaceDescription')]" | Select-Object -ExpandProperty Node
    $NetAdapterDriverVers  = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'DriverVersion')]" | Select-Object -ExpandProperty Node
    $NetAdapterAdvSettings = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'AdvancedSettings')]/Property/Property" | Select-Object -ExpandProperty Node

    Write-Host -NoNewline -ForegroundColor Yellow "Checking the presence of the network adapter specified in XML File... "
    $NetAdapterList = Get-NetAdapter -Physical -InterfaceDescription "$($NetAdapterIfName.("#text"))" -ErrorAction SilentlyContinue

    Switch($true){
        ($($NetAdapterList | Measure-Object).count -eq 0){Write-Error "No matches found" -Category InvalidResult;exit}
        ($($NetAdapterList | Measure-Object).count -gt 1){Write-Error "Multiple matches found, unable to continue." -Category InvalidResult;exit}
        ($($NetAdapterList | Measure-Object).count -eq 1){
            Write-Host -ForegroundColor Green "Match found"
            Write-host -ForegroundColor Yellow -NoNewline "Checking the driver version specified in XML File... "

            if($NetAdapterList.DriverVersion -eq $NetAdapterDriverVers.("#text")){
                Write-Host -ForegroundColor Green "Match found"
                Write-host -ForegroundColor Yellow  "Import settings... "

                Foreach($setting in $NetAdapterAdvSettings){
                    $GetValue = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'$($setting.Name)')]//Property[@Name='Value']" | Select-Object -ExpandProperty Node
        
                    Write-Host -ForegroundColor Yellow -NoNewline "`t $($setting.Name) : "
                    if($null -ne $GetValue.("#text")){
                        Set-NetAdapterAdvancedProperty -InterfaceDescription $NetAdapterList.InterfaceDescription -RegistryKeyword "$($($setting.Name))" -DisplayValue $($GetValue.("#text"))
                        Write-Host -ForegroundColor green "Applied"
                        
                    }else{
                        Write-Host -ForegroundColor Red "No Value to set"
                    }
                }
                Write-Host -ForegroundColor green "Process Completed"
            }else{
                Write-Error "Drivers Versions do not match" -Category InvalidResult
                exit
            }
        }
    }
}catch{
    Write-Error  "$($_.InvocationInfo.ScriptLineNumber) : $($_)"
}





