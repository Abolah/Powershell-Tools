#@ FileName: Printer_Inventory.ps1
#========================================================================
#@ Script Name: Inventaire des imprimantes
#@ Created: [DATE_DMY]- 18/04/2016 
#@ Author: Maxime Cholon
#@ Email: maxime.cholon@maif.fr
#@ OS: Windows 7 / 2012
#@ Version History: 1.2
#========================================================================
#@ Objectif: Récolter l'intégralité des imprimantes qui ping sur les serveurs.
#@ Prérequis : Liste des serveurs dans un fichier Serveurs.txt situé à la racine du script.
#@ Prérequis : Executer le script avec un compte ADS.
#========================================================================

#===========================Code Start===================================
clear
$SCRIPT_PARENT   = Split-Path -Parent $MyInvocation.MyCommand.Definition 

#=========================================================================
Function HostList {$Printservers =  Get-Content ($SCRIPT_PARENT + "\Input\Servers.txt")}
#=========================================================================
Function Host { $Printservers = Read-Host}
#=========================================================================
Function LocalHost { $Printservers = $env:computername }
#=========================================================================

#=========================================================================
#Récupération des Informations fournies par l'utilisateur.
Write-Host "**************************" -ForegroundColor Green
Write-Host "Inventaire des Imprimantes" -ForegroundColor Green
Write-Host "**************************" -ForegroundColor Green
Write-Host " "

$strResponse = Read-Host "
[1] Depuis un fichier (Servers.txt).
[2] Entrer le nom d'un serveur manuellement.
[3] Poste local.
[4] Quitter.
"

If($strResponse -eq "1"){. HostList}
    elseif($strResponse -eq "2"){. Host}
    elseif($strResponse -eq "3"){. LocalHost}
    elseif($strResponse -eq "4"){Break}
    else
    {
        Write-Host "La réponse fournie n'est pas correcte, Veuillez relancer le script." -foregroundColor Red
        Break
    }                                                                

Write-Progress -Activity "Recherche dans l'inventaire" -status "Le script prend beaucoup de temps pour faire l'inventaire. Allez prendre un café" -id 1
#========================================================================
#Création d'une feuille Excel.
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True
$Excel = $Excel.Workbooks.Add()
$Sheet = $Excel.Worksheets.Item(1)
$Sheet.Name = "Inventaire des imprimantes"
#========================================================================
$Sheet.Cells.Item(1,1) = "Serveur d'impression"
$Sheet.Cells.Item(1,2) = "Nom de l'imprimante"
$Sheet.Cells.Item(1,3) = "Nom du port"
$Sheet.Cells.Item(1,4) = "Nom du partage"
$Sheet.Cells.Item(1,5) = "Nom du pilote"
$Sheet.Cells.Item(1,6) = "Version du pilote"
$Sheet.Cells.Item(1,7) = "Pilote"
$Sheet.Cells.Item(1,8) = "Location"
$Sheet.Cells.Item(1,9) = "Commentaires"
$Sheet.Cells.Item(1,10) = "Etat du partage"
#========================================================================
$intRow = 2
$WorkBook = $Sheet.UsedRange
$WorkBook.Interior.ColorIndex = 40
$WorkBook.Font.ColorIndex = 11
$WorkBook.Font.Bold = $True
#========================================================================
ForEach ($Printserver in $Printservers)
{   $Printers = Get-WmiObject Win32_Printer -ComputerName $Printserver
    ForEach ($Printer in $Printers)
    {
        if ($Printer.Name -ne "Microsoft*")
        {
            $Sheet.Cells.Item($intRow, 1) = $Printserver
            $Sheet.Cells.Item($intRow, 2) = $Printer.Name
            $Testping = test-Connection -ComputerName $Printer.Name -Count 1 -Quiet
            if ($Testping -Or $Printer.Name -like"IMPRIC*")
            {
                if ($Printer.Shared -eq "VRAI")
                {
                    $Sheet.Cells.Item($intRow, 1) = $Printserver
                    $Sheet.Cells.Item($intRow, 2) = $Printer.Name

                    if ($Printer.PortName -notlike "*\*")
                    {
                        $Ports = Get-WmiObject Win32_TcpIpPrinterPort -Filter "name = '$($Printer.Portname)'" -ComputerName $Printserver
                        ForEach ($Port in $Ports)
                        {
                            $Sheet.Cells.Item($intRow, 3) = $Printer.Name
                        }
                    }
                    $PrintShare = $Printer.ShareName
                    $Sheet.Cells.Item($intRow, 4) = "\\"+ $Printserver + "\"+ $Printer.Name
                    $Sheet.Cells.Item($intRow, 5) = $Printer.DriverName
                    $Drivers = Get-WmiObject Win32_PrinterDriver -Filter "__path like '%$($Printer.DriverName)%'" -ComputerName $Printserver
                    ForEach ($Driver in $Drivers)
                    {
                        $Drive = $Driver.DriverPath.Substring(0,1)
                        $Sheet.Cells.Item($intRow,6) = (Get-ItemProperty ($Driver.DriverPath.Replace("$Drive`:","\\$PrintServer\$Drive`$"))).VersionInfo.ProductVersion
                        $Sheet.Cells.Item($intRow,7) = Split-Path $Driver.DriverPath -Leaf
                    }
                    $Sheet.Cells.Item($intRow, 8) = $Printer.Location
                    $Sheet.Cells.Item($intRow, 9) = $Printer.Comment
                    $Sheet.Cells.Item($intRow, 10) = $Printer.Shared
                    $intRow ++
                }
            }else
            {
                #Tous les postes qui ne ping pas et qui ne sont pas partagés + les IMPRIC
            }
        }
    }
    $WorkBook.EntireColumn.AutoFit() | Out-Null
}
##############################################
#Fin de l'inventaire
$intRow ++ 
$Sheet.Cells.Item($intRow,1) = "Inventaire terminé"
$Sheet.Cells.Item($intRow,1).Font.Bold = $True
$Sheet.Cells.Item($intRow,1).Interior.ColorIndex = 40
$Sheet.Cells.Item($intRow,2).Interior.ColorIndex = 40

##############################################
#Sauvegarde de la feuille Excel
$filename = Read-Host " Donnez un nom au fichier sans préciser l'extension"
$output = ($SCRIPT_PARENT + "\Output\" +$filename+".xlsx")
$Excel.SaveAs($output)
##############################################
clear