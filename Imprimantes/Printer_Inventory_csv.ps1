#@=============================================
#@ FileName: Printer_Inventory.ps1
#@=============================================
#@ Script Name: Inventaire des imprimantes
#@ Created: [DATE_DMY]- 18/04/2016 
#@ Author: Maxime Cholon
#@ Email: maxime.cholon@maif.fr
#@ OS: Windows 7 / 2012
#@ Version History: 1.1
#@=============================================
#@ Objectif: Récolter l'intégralité des imprimantes sur les serveurs.
#@ Prérequis : Liste des serveurs dans un fichier Serveurs.txt situé à la racine du script.
#@ Prérequis : Executer le script avec un compte ADS.
#@=============================================

#@================Code Start===================
clear
$SCRIPT_PARENT   = Split-Path -Parent $MyInvocation.MyCommand.Definition 

#=========================================================================
Function HostList {$Printservers =  Get-Content ($SCRIPT_PARENT + "\Servers.txt")}
# ========================================================================
Function Host { $Printservers = Read-Host "Entrer le nom d'un poste ou son IP"}
# ========================================================================
Function LocalHost { $Printservers = $env:computername }
#@========================================================================
#Récupération des Informations fournies par l'utilisateur.

Write-Host "**************************"               -ForegroundColor Green
Write-Host "Inventaire des Imprimantes"               -ForegroundColor Green
Write-Host "**************************"               -ForegroundColor Green
Write-Host " "

$strResponse = Read-Host "
[1] Depuis une fichier (Servers.txt).
[2] Entrer le nom d'une imprimante manuellement.
[3] Poste local.
"

If($strResponse -eq "1"){. HostList}
    elseif($strResponse -eq "2"){. Host}
    elseif($strResponse -eq "3"){. LocalHost}
    else
    {
        Write-Host "La réponse fournie n'est pas correcte, Veuillez relancer le script." -foregroundColor Red
    }                                                                

Write-Progress -Activity "Recherche dans l'inventaire" -status "En Cours..." -id 1
###########################################################################################################################
ForEach ($Printserver in $Printservers)
{   $Printers = Get-WmiObject Win32_Printer -ComputerName $Printserver
    ForEach ($Printer in $Printers)
    {
        if ($Printer.Name -like "*")
        {
            $PrinterServer = $Printserver
            $PrinterName = $Printer.Name
        if ($Printer.PortName -notlike "*\*")
            {   $Ports = Get-WmiObject Win32_TcpIpPrinterPort -Filter "name = '$($Printer.Portname)'" -ComputerName $Printserver
                ForEach ($Port in $Ports)
                {
                    $PrinterPort = $Port.HostAddress
                }
            }
            $PrinterShare = $Printer.ShareName   
            $PrinterDriver = $Printer.DriverName
            ####################       
            $Drivers = Get-WmiObject Win32_PrinterDriver -Filter "__path like '%$($Printer.DriverName)%'" -ComputerName $Printserver
            ForEach ($Driver in $Drivers)
            {
                $PrinterDrivers = $Driver.DriverPath.Substring(0,1)
                $PrinterVersion = (Get-ItemProperty ($Driver.DriverPath.Replace("$Drive`:","\\$PrintServer\$Drive`$"))).VersionInfo.ProductVersion
                $PrinterPath = Split-Path $Driver.DriverPath -Leaf
            }
            ####################
            $PrinterLocation = $Printer.Location
            $PrinterComment = $Printer.Comment
            $PrinterShared = $Printer.Shared
        }
    }
}
##############################################################################################################################