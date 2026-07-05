#requires -version 5.1
<#
.SYNOPSIS
    Startpunkt fuer das Tool "DHCP MAC-Reservierung".

.DESCRIPTION
    Laedt Settings, Services und die WinForms-Oberflaeche des Projekts.

.NOTES
    Projekt: DHCP MAC-Reservierung
    Umgebung: Windows Server 2022 / Windows PowerShell 5.1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$settingsPath = Join-Path $scriptRoot 'Settings.psd1'

if (-not (Test-Path $settingsPath)) {
    [System.Windows.Forms.MessageBox]::Show("Settings.psd1 wurde nicht gefunden:`r`n$settingsPath", 'Fehler', 'OK', 'Error') | Out-Null
    return
}

$settings = Import-PowerShellDataFile -Path $settingsPath

. (Join-Path $scriptRoot 'services\CsvMacListService.ps1')
. (Join-Path $scriptRoot 'services\DhcpReservationService.ps1')
. (Join-Path $scriptRoot 'views\DhcpMacReservierungForm.ps1')

Show-DhcpMacReservierungForm -Settings $settings
