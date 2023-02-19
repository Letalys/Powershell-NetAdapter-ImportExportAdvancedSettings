#Requires -Version 4.0
Param([String]$ImportXMLFile)

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $false){
    Write-Error "This script requires elevation of privilege to run" -Category AuthenticationError
    pause
    exit
}

cls
Write-Host -ForegroundColor Cyan "------------------------------------------------------------"
Write-Host -ForegroundColor Cyan "       NetAdapter Advanced Settings Importing               "
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

    $NetAdapterIfName = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'InterfaceDescription')]" | Select -ExpandProperty Node
    $NetAdapterDriverVers  = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'DriverVersion')]" | Select -ExpandProperty Node
    $NetAdapterAdvSettings = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'AdvancedSettings')]/Property/Property" | Select -ExpandProperty Node

    Write-Host -NoNewline -ForegroundColor Yellow "Checking the presence of the network adapter specified in XML File... "
    $NetAdapterList = Get-NetAdapter -Physical -InterfaceDescription "$($NetAdapterIfName.("#text"))" -ErrorAction SilentlyContinue

    Switch($true){
        ($($NetAdapterList | Measure).count -eq 0){Write-Error "`nNo matches found" -Category InvalidResult;exit}
        ($($NetAdapterList | Measure).count -gt 1){Write-Error "`n Multiple matches found, unable to continue." -Category InvalidResult;exit}
        ($($NetAdapterList | Measure).count -eq 1){
            Write-Host -ForegroundColor Green "Match found"
            Write-host -ForegroundColor Yellow -NoNewline "Checking the driver version specified in XML File... "

            if($NetAdapterList.DriverVersion -eq $NetAdapterDriverVers.("#text")){
                Write-Host -ForegroundColor Green "Match found"
                Write-host -ForegroundColor Yellow  "Import settings... "

                Foreach($setting in $NetAdapterAdvSettings){
                    $GetValue = Select-Xml -Xml $XML -XPath "//Property[contains(@Name,'$($setting.Name)')]//Property[@Name='Value']" | Select -ExpandProperty Node
        
                    Write-Host -ForegroundColor Yellow -NoNewline "`t $($setting.Name) : "
                    if($GetValue.("#text") -ne $null){
                        Set-NetAdapterAdvancedProperty -InterfaceDescription $NetAdapterList.InterfaceDescription -RegistryKeyword "$($($setting.Name))" -DisplayValue $($GetValue.("#text"))
                        Write-Host -ForegroundColor green "Applied"
                        
                    }else{
                        Write-Host -ForegroundColor Red "No Value to set"
                    }
                }
                Write-Host -ForegroundColor green "Process Completed"
            }else{
                Write-Error "`nDrivers Versions do not match" -Category InvalidResult
                exit
            }
        }
    }
}catch{
    Write-Error "`n $_"
}





