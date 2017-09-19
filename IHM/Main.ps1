#@ FileName: Main.ps1
#========================================================================
#@ Script Name: WinCheck
#@ Created: [DATE_DMY]- 18/05/2016 
#@ Author: Maxime Cholon
#@ OS: Windows 7 / 2012
#@ Version History: 1.1
#========================================================================
#@ Objectif: Inventaire de scripts couramment utilisés
#@ Prérequis : Vérifier ques les chemins d'accès aux scripts sont les bons.
#@ INSTRUCTIONS: NE PAS MODIFIER LE FICHIER XAML SOUS PEINE DE TOUT FAIRE PLANTER.
#========================================================================
#ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window x:Class="WinCheck.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WinCheck"
        mc:Ignorable="d"
        Title="WinCheck" Height="570" Width="640">
    <Grid ScrollViewer.VerticalScrollBarVisibility="Disabled">
        <Image x:Name="sevenup" Margin="10,10,461,416" Source="D:\Scripts\IHM\img\7up.png" Stretch="Fill"/>
        <Label x:Name="label" Content="Outil de Diagnostic des postes de travail" HorizontalAlignment="Left" Height="33" Margin="229,26,0,0" VerticalAlignment="Top" Width="317" FontSize="14"/>
        <Button x:Name="diag" Content="Diag desPostes" HorizontalAlignment="Left" Height="79" Margin="48,146,0,0" VerticalAlignment="Top" Width="112" Grid.RowSpan="10" Grid.Row="1"/>
        <Button x:Name="sccm" Content="Administration SCCM" HorizontalAlignment="Left" Height="79" Margin="229,146,0,0" VerticalAlignment="Top" Width="121"/>
        <Button x:Name="inventory" Content="Inventaire" HorizontalAlignment="Left" Height="79" Margin="406,146,0,0" VerticalAlignment="Top" Width="126"/>
        <Button x:Name="driver" Content="Installation de pilotes" HorizontalAlignment="Left" Height="72" Margin="48,274,0,0" VerticalAlignment="Top" Width="112"/>
        <Button x:Name="wol" Content="WakeOnLan" HorizontalAlignment="Left" Height="72" Margin="229,274,0,0" VerticalAlignment="Top" Width="121"/>
        <Button x:Name="provisio" Content="Provisionning" HorizontalAlignment="Left" Height="72" Margin="406,274,0,0" VerticalAlignment="Top" Width="126"/>
        <Button x:Name="quit" Content="Quitter" HorizontalAlignment="Left" Height="41" Margin="522,475,0,0" VerticalAlignment="Top" Width="89"/>
    </Grid>
</Window>
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF_$($_.Name)" -Value $Form.FindName($_.Name)}
#===========================================================================
# Les boutons d'appel de scripts.
#===========================================================================
$WPF_diag.Add_Click({
    Write-Host "Chemin du script de diag des postes maif"
    $Form.close() | Out-Null
})
$WPF_sccm.Add_Click({
    Write-Host "Chemin vers SCCM"
    $Form.close() | Out-Null
})
$WPF_inventory.Add_Click({
    Write-Host "Chemin du script d'inventaire"
    $Form.close() |Out-Null
})

$WPF_driver.Add_Click({
    Write-Host "Chemin du script d'install pilotes"
    $Form.close() |Out-Null
})
$WPF_wol.Add_Click({
    Write-Host "Chemin du script de WoL"
    $Form.close() | Out-Null
})
$WPF_provisio.Add_Click({
    Write-Host "Chemin vers script de provisionning"
    $Form.Close() | Out-Null
})
#===========================================================================
#Bouton Quitter
#===========================================================================
$WPF_quit.Add_Click({$form.Close()})
#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | Out-Null