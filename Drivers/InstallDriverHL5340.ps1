#----------------------------------------------------------------
# InstallDriverHL5340.ps1
# v1.0 - 27/10/2015
# Cholon Maxime
#----------------------------------------------------------------
# Description :
# Ce script sert à remplacer les drivers d'impressions des postes de travail
# Usage:
# Ce script est positionné au Logon du Poste
#----------------------------------------------------------------
# Modification :
# Maxime Cholon- Optimisation du script (suppression de l'ajout du driver intermediaire et affectation du bon script après l'ajout)
#                 - Supression de la variable du driver intermediaire
#                 - Ajout du fichier CSV
#                 - Ajout du Psexec.exe pour l'éxecution du "prndrv.vbs" depuis l'ordinateur distant
#                 - Test à distance OK
#                 - Boucle d'éxecution en fonctions du nombre de postes renseignés dans le fichier CSV
#-----------------------------------------------------------------------------------------------------------------------------------------
#########
# var   #
#########
#Variable Auto Incrémentale
$traitement = 1;
#Variable du CSV
$ComputerName = ""
$CSV_Path="\List_Computers.csv"
#Variables de l'imprimante
    $architecture = "Windows NT x86"
    $printerName = "Brother Mono Universal Printer (PCL)"

#Variable du script prndrvr.vbs
    $scriptPath ="C:\Windows\System32\Printing_Admin_Scripts\fr-FR\prndrvr.vbs"

#Variables du Driver
    $driverName = "Brother HL-5440D BR-Script3"
    $driverModel = "Brother HL-5440D BR-Script3"
    $oldDriverModel = "Brother Mono Universal Printer (PCL)"
    $driverPath = "\DrvHl5340D"
    $infPath = "\DrvHl5340D\\BRPHL11A.INF"

###########
#Fonctions#
###########

function restart_spooler()
{
    net stop spooler
    net start spooler
}

#########
#Actions#
#########
clear
$CSV_File = Import-Csv -Path $CSV_Path -Delimiter ";" 

write-host "`n#####################################################" -ForegroundColor White
write-host "#                Début du Script                    #" -ForegroundColor White
write-host "#####################################################" -ForegroundColor White

       
if (-not $CSV_File)
{
       Write-Host "[ERREUR]: Le fichier CSV n'est pas renseigné" -ForegroundColor red
	   Write-Host " Syntaxe :  .\create_printer_v4 -ListPrinters Photocop.csv"
       exit 1
}
else {write-host "`n[CSV] : le fichier CSV est bien renseigné" -ForegroundColor green}


Foreach ($Line in $CSV_File) 
{ 
	$ComputerName = $Line.Postes
        write-host "`n-----------------------------------------------------" -ForegroundColor Yellow
        write-host "-----------------Traitement $traitement------------------------" -ForegroundColor Yellow
        write-host "-----------------------------------------------------" -ForegroundColor Yellow
        write-host "`n`n[Poste : $ComputerName connecté]" -ForegroundColor Green

    #Obtention des informations de l'imprimante
    $printers = get-WmiObject Win32_Printer -ComputerName $ComputerName  #Appel de la class WIN32_Printer de l'ordinateur distant
    $drivers = Get-WmiObject Win32_PrinterDriver -ComputerName $ComputerName

    #Recherche de l'imprimante par son nom
    $printer = $printers | where {$_.name -eq "Brother Mono Universal Printer (PCL)" -or $_.name -like "Brother HL*"}

    write-host "`n[Ci-dessous, l'imprimante à configurer ]" -ForegroundColor Cyan
    $printer
    #Récupère la version du pilote actuel
    write-host "[Pilote Actuel : "$printer.DriverName"]" -ForegroundColor gray


    if($printer.DriverName -eq $driverName)
    {
        Write-Warning "Vous avez le mauvais pilote, patientez... Nous installons le bon pilote"
        # Vide la file d'impression
        $hide = $printer.CancelAllJobs()

        # RedÃ©marrage du spooler
        restart_spooler

        # ajoute le pilote d'impression
        \\PWF70610\TestNewDriver5440\.\psexec.exe -s \\$ComputerName Cscript $scriptPath -a -m $driverModel -v 3 -e $architecture  -i  $infPath -h $driverPath 2>null

        # Affectation du bon pilote  
        $printer.DriverName = $driverModel
        $hide = $printer.Put()

        # supprime le pilote d'impression
        \\PWF70610\TestNewDriver5440\.\psexec.exe -s \\$ComputerName Cscript $scriptPath -d -m $oldDriverModel -v 3 -e $architecture 2>null

   
        # Print a test page
        #$hide = $printer.PrintTestPage()
        if($printer.DriverName -eq $driverName)
        {
            Write-host "`nVous avez maintenant le bon pilote !" -ForegroundColor green
        }
        else{write-host "`n[Execution] : False" -ForegroundColor Red
        }   
    }
    else
    {
        Write-host "`nVous avez le bon pilote !" -ForegroundColor green 
    
    }
    write-host "`n[Pilote Actuel : "$printer.DriverName"]" -ForegroundColor green
    $traitement++
}


write-host "`n#####################################################" -ForegroundColor White
write-host "#                Fin du Script                      #" -ForegroundColor White
write-host "#####################################################" -ForegroundColor White
