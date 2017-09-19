<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.123
	 Created on:   	01/07/2016 13:55
	 Created by:   	maxime Cholon
	 Organization: 	Maif
	 Filename:     	Chrome_Launcher.ps1
	===========================================================================
	.DESCRIPTION
		lance Chrome automatiquement en plein écran.
#>
start chrome.exe -ArgumentList "--start-fullscreen http://www.maif.fr"