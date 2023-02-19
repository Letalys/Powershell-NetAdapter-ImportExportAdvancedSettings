# PowerShell : NetAdapter-ImportExport-AdvancedSettings

The objective of this project is to be able to easily Export and Import all the advanced parameters of the physical network cards in order to re-import them or integrate them on another machine requiring the same parameters.

This script was born from a request to standardize the parameters of my company's network cards, in particular for Wi-Fi cards.

I'm using this scripts with Microsoft EndPoint Configuration Manager for deploying networkCard configuration.

## Scripts Versions
### [v1](./v1) : Using NetAdapter Module
This version of Import/Export NetworkAdapter AdvancedSettings use the NetAdapter module of Powershell.This Module was introduced by __Powershell 4.0.__ Works only with version 4.0 or higher of Powershell on __Windows 8 and higher__.

### [v2](./v2) : Using WMI and Registry
This Version of Import/Export NetworkAdapter AdvancedSettings use WMI and Registry to work. It can use on all Windows Version. But i recommand to use v1 for compatible Operating System.

**This version is being remake to optimize the code and reduce its weight, but current works**

## NetworkAdapter Models and Drivers Versions
Imports of settings can only be carried out if the network adapter and its driver version correspond to the elements present on the target system.

Some advanced parameters and authorized values ​​are only available depending on the model of the network card and especially its driver version.


## 🔗 Liens
https://github.com/Letalys/Powershell-NetAdapter-ImportExportAdvancedSettings


## Auteur
- [@Letalys (GitHUb)](https://www.github.com/Letalys)
