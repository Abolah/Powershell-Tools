<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.123
	 Created on:   	06/07/2016 15:00
	 Created by:   	Maxime Cholon
	 Filename:message.ps1
	===========================================================================
	.DESCRIPTION
		Envoie un message sous forme de popup sur le PC distant.
#>
>
cls
Invoke-Command -ComputerName "Name of the Computer" -Scriptblock {
	$CmdMessage = { C:\windows\system32\msg.exe * "message to send" }
	$CmdMessage | Invoke-Expression
} 
