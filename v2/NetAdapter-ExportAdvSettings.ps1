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
Function Get-RegistryClassPath{
    param(
        [Parameter(Mandatory=$true)][String]$Class
	)

    write-host ""
    Write-Host -NoNewLine -ForegroundColor Yellow "Searching $Class GUID Class :"

    $ParentClassKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Class"
    $NetClassGUID = Get-ChildItem -Path $ParentClassKey | Get-ItemProperty -Name "Class" -ErrorAction SilentlyContinue | Where-Object {$_.class -eq "$Class"}

    if($null -ne $NetClassGUID){
        Write-Host -ForegroundColor Green $NetClassGUID.PSChildName

        $ParentNetClassKey =  Join-Path -Path $ParentClassKey -ChildPath $NetClassGUID.PSChildName

        return $ParentNetClassKey
    }
}

Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host -ForegroundColor Cyan "       Exportation des paramètres des cartes réseau v2      "
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host ""


Write-Host -Foreground Yellow "Physical Network Adapter Available :"
Get-WMIObject Win32_NetworkAdapter -Filter "PhysicalAdapter='True'" | Select-Object DeviceID,Name, ServiceName  |Format-Table -AutoSize

$SelectedIndex= Read-Host "Enter the index of the network adapter to export : "
$SelectedNetAdapter = Get-WMIObject Win32_NetworkAdapter -Filter "DeviceID=$SelectedIndex AND PhysicalAdapter='True'"

if($null -ne $SelectedNetAdapter){
    Write-Host ""
    Write-Host -Foreground Yellow "Votre sélection : "

    $SelectedNetAdapter | Select-Object DeviceId, Name, ServiceName | Format-Table

    Write-Host ""
    Write-Host -Foreground Yellow "Information about linked Drivers (Wait) :"

    $SelectNedAdpterDriver = Get-WMIObject Win32_PnpSignedDriver -Filter "DeviceName='$($SelectedNetAdapter.Name)'" | Select-Object DriverVersion,DriverDate,infName 
    $SelectNedAdpterDriver | Format-Table -AutoSize

    $GUIDClass = Get-RegistryClassPath -Class NET


    #Récupération de l'index au format XXXX
    $SelectedIndexFormated = $SelectedIndex
    While($SelectedIndexFormated.Length -ne 4){
        $SelectedIndexFormated = "0$SelectedIndexFormated"
    }

     Write-Host -NoNewline -Foreground Yellow "Index réél : "
     Write-Host $SelectedIndexFormated

     $RegistryAdapterPath = Join-Path -Path $GUIDClass -ChildPath $SelectedIndexFormated

     Write-Host -NoNewline -Foreground Yellow "Registry : "
     Write-Host $RegistryAdapterPath

     Write-Host ""
     Write-Host -NoNewline -Foreground Yellow "Get registry Information : "

     Try {
         $objNetAdapter = New-Object PSobject
         $ObjNetAdapter | Add-Member -Name "Name" -membertype Noteproperty -Value $SelectedNetAdapter.Name
         $ObjNetAdapter | Add-Member -Name "FileName" -membertype Noteproperty -Value $SelectedNetAdapter.ServiceName
         $ObjNetAdapter | Add-Member -Name "INF" -membertype Noteproperty -Value $SelectNedAdpterDriver.InfName
         $ObjNetAdapter | Add-Member -Name "Version" -membertype Noteproperty -Value $SelectNedAdpterDriver.DriverVersion
         $ObjNetAdapter | Add-Member -Name "DateVersion" -membertype Noteproperty -Value ($SelectNedAdpterDriver.DriverDate.Substring(0,4) + "-" + $SelectNedAdpterDriver.DriverDate.Substring(4,2) + "-" + $SelectNedAdpterDriver.DriverDate.Substring(6,2))
         $ObjNetAdapter | Add-Member -Name "RegistryPath" -MemberType NoteProperty -Value $RegistryAdapterPath

         $ParamsList = Get-ChildItem -Path "$($ObjNetAdapter.RegistryPath)\Ndi\Params"
     
         [System.Collections.ArrayList]$AdvancedParamsTab=@()
         Foreach($param in $ParamsList){
            $ObjAdvancedParam = New-Object PSobject

            #On récupère le RegistryPath du paramètre
            $ObjAdvancedParam | Add-Member -Name "RegistryKeyPath" -membertype Noteproperty -Value $param.PSPath

            #On récupère le nom de la sous clé de Ndi\Params qui est également le nom de la propriété avancée RegistryKeyWord
            $ObjAdvancedParam | Add-Member -Name "RegistryKeyword" -membertype Noteproperty -Value $param.PSChildName

            #Récupération de la description du paramètre.
            $CurrentParam = Get-ItemProperty -Path $ObjAdvancedParam.RegistryKeyPath
            $ObjAdvancedParam | Add-Member -Name "ParamDesc" -membertype Noteproperty -Value $($CurrentParam.ParamDesc)

            #On remonte dans la clé principale (de l'ObjAdapter) pour récupérer la valeur utilisée
            $CurrentParamRegistryValue = Get-ItemProperty -Path $ObjNetAdapter.RegistryPath 
            $ObjAdvancedParam | Add-Member -Name "RegistryValue" -membertype Noteproperty -Value $CurrentParamRegistryValue.$($ObjAdvancedParam.RegistryKeyword)

            #On récupère la displayValue associée à la registryKey qui est le nom de la propriété dans les énumérations de la propriété
            if(Test-Path "$($ObjAdvancedParam.RegistryKeyPath)\enum"){
                $CurrentDisplayValue = Get-ItemProperty -Path "$($ObjAdvancedParam.RegistryKeyPath)\enum"
                $ObjAdvancedParam | Add-Member -Name "DisplayValue" -membertype Noteproperty -Value $CurrentDisplayValue.$($ObjAdvancedParam.RegistryValue)

                #On récupère la liste complète des RegistryValue et des DisplayValue dans les Enum pour obtenir la liste des valeurs autorisées.
                [System.Collections.ArrayList]$ValidParamValuesTab=@()

                #Récupération des registryValue disponibles
                $ValidRegistryValues = Get-Item -Path "$($ObjAdvancedParam.RegistryKeyPath)\enum" | Select-object -ExpandProperty Property

                Foreach($vrv in $ValidRegistryValues){
                    #Write-Host -ForegroundColor Yellow "Valeur Possible pour $($ObjAdvancedParam.RegistryKeyWord)"
                    #write-host -ForegroundColor Green $vrv

                    $ObjAdvancedParamValidValue = New-Object PSobject
                    $ObjAdvancedParamValidValue | Add-Member -Name "ValidRegistryValue" -MemberType NoteProperty -Value $vrv
            
                    #On récupère les DisplayValue associée
                    $ValidDisplayValue = Get-ItemProperty -Path "$($ObjAdvancedParam.RegistryKeyPath)\enum"
                    #Write-Host $ValidDisplayValue.$vrv

                    $ObjAdvancedParamValidValue | Add-Member -Name "ValidDisplayValue" -MemberType NoteProperty -Value $ValidDisplayValue.$vrv

                    #$ValidDisplayValue | Ft

                    $ValidParamValuesTab.Add($ObjAdvancedParamValidValue) | Out-Null
                }
        
                $ObjAdvancedParam | Add-Member -Name "ValidValues" -membertype Noteproperty -Value $ValidParamValuesTab
            }
            $AdvancedParamsTab.Add($ObjAdvancedParam) | Out-Null
         }

         $ObjNetAdapter | Add-Member -Name "AdvancedParam" -MemberType NoteProperty -Value $AdvancedParamsTab

         Write-Host -ForegroundColor Green "Completed"
     }Catch{
        Write-Error -Message $_.Message
        $objNetAdapter = $null
     }
     #L'objet NetAdapter est construit


    If($null -ne $objNetAdapter){
        Write-host ""
        Write-Host -ForegroundColor Yellow "Export to XML : "
    
    Try {
        [xml]$XMLFile = New-Object System.Xml.XmlDocument
        $XMLDeclaration = $XMLFile.CreateXmlDeclaration("1.0","UTF-8",$null)
        $XMLFile.AppendChild($XMLDeclaration) | out-null

        $XMLRoot = $XMLFile.CreateNode("element","NetAdapterExport",$null)

        #region XML : ExportInformation
        $ExportInformationsNode = $XMLFile.CreateNode("element","ExportInformations","$null")

        $Element = $XMLFile.CreateElement("Origin")
        $Element.InnerText = $env:COMPUTERNAME
        $ExportInformationsNode.AppendChild($Element) | Out-Null

        $WMIProductInfo = Get-WmiObject -Class  Win32_ComputerSystemProduct

        $Element = $XMLFile.CreateElement("Manufacturer")
        $Element.InnerText = $WMIProductInfo.Vendor
        $ExportInformationsNode.AppendChild($Element) | Out-Null

        $Element = $XMLFile.CreateElement("Model")
        $Element.InnerText = $WMIProductInfo.Name
        $ExportInformationsNode.AppendChild($Element) | Out-Null

        $Element = $XMLFile.CreateElement("Version")
        $Element.InnerText = $WMIProductInfo.Version
        $ExportInformationsNode.AppendChild($Element) | Out-Null

        $Element = $XMLFile.CreateElement("SN")
        $Element.InnerText = $WMIProductInfo.IdentifyingNumber
        $ExportInformationsNode.AppendChild($Element) | Out-Null

        $WMIOSInfo = Get-WmiObject -Class  Win32_OperatingSystem

        $Element = $XMLFile.CreateElement("OperatingSystemVersion")
        $Element.InnerText = $WMIOSInfo.Version
        $ExportInformationsNode.AppendChild($Element) | Out-Null

        $Element = $XMLFile.CreateElement("ExportDateTime")
        $Element.InnerText = $(Get-Date -Format "yyyy-MM-dd HH:mm")
        $ExportInformationsNode.AppendChild($Element) | Out-Null

        $XMLRoot.AppendChild($ExportInformationsNode) | out-null
        #endregion XML : ExportInformation

        #region XML : NetAdapterInformations
        $ExportNetAdapterInformations = $XMLFile.CreateNode("element","NetAdapterInformations","$null")

        $Element = $XMLFile.CreateElement("Name")
        $Element.InnerText = $objNetAdapter.Name
        $ExportNetAdapterInformations.AppendChild($Element) | Out-Null

        $ExportDriverInfos = $XMLFile.CreateNode("element","DriverInfos","$null")
        $ExportNetAdapterInformations.AppendChild($ExportDriverInfos) | Out-Null

        $Element = $XMLFile.CreateElement("FileName")
        $Element.InnerText = $objNetAdapter.FileName
        $ExportDriverInfos.AppendChild($Element) | Out-Null

        $Element = $XMLFile.CreateElement("INF")
        $Element.InnerText = $objNetAdapter.INF
        $ExportDriverInfos.AppendChild($Element) | Out-Null

        $Element = $XMLFile.CreateElement("Version")
        $Element.InnerText = $objNetAdapter.Version
        $ExportDriverInfos.AppendChild($Element) | Out-Null

        $Element = $XMLFile.CreateElement("DateVersion")
        $Element.InnerText = $objNetAdapter.DateVersion
        $ExportDriverInfos.AppendChild($Element) | Out-Null

        $ExportAdapterConfiguration = $XMLFile.CreateNode("element","AdapterConfiguration","$null")
        $ExportNetAdapterInformations.AppendChild($ExportAdapterConfiguration) | Out-Null

        foreach($p in $objNetAdapter.AdvancedParam){
            $ExportAdapterConfigurationParam = $XMLFile.CreateNode("element","Param","$null")
            $ExportAdapterConfigurationParam.SetAttribute("RegistryKeyWord",$p.RegistryKeyWord)
            $ExportAdapterConfigurationParam.SetAttribute("RegistryValue",$p.RegistryValue)

            $Element = $XMLFile.CreateElement("DisplayName")
            $Element.InnerText = $p.ParamDesc
            $ExportAdapterConfigurationParam.AppendChild($Element) | Out-Null
    
            $ExportValidValue = $XMLFile.CreateNode("element","ValidDisplayValues","$null")
            $ExportAdapterConfigurationParam.AppendChild($ExportValidValue) | Out-Null
    
            foreach($v in $p.ValidValues){
                $Element = $XMLFile.CreateElement("Value")
                $Element.InnerText = $v.ValidDisplayValue
                $Element.SetAttribute("ValidRegistryValue",$v.ValidRegistryValue)
                $ExportValidValue.AppendChild($Element) | Out-Null
            }

            $ExportAdapterConfiguration.AppendChild($ExportAdapterConfigurationParam) | Out-Null
        }

        $XMLRoot.AppendChild($ExportNetAdapterInformations) | out-null
        #endregion NetAdapterInformations
	
        $XMLFile.AppendChild($XMLRoot) | out-null

        #Windows 7 : définition PSSCriptRoot
        if($null -eq $PSScriptRootCustom){
            $PSScriptRootCustom = split-path -parent $MyInvocation.MyCommand.Definition
        }

        $SaveFile = "$PSScriptRootCustom\Exports\$($SelectedNetAdapter.Name).NetAdapteExportCfg.xml"

        Write-Host -ForegroundColor Yellow -NoNewline "`t Save Path : "
        Write-host "$SaveFile"

		$XMLFile.Save($SaveFile)

        }Catch{
            Write-Host -ForegroundColor Red "Failed : $_"
        }


            }
}else{
     Write-host -ForegroundColor Red "NetworkAdapter Index not found"
}
 pause