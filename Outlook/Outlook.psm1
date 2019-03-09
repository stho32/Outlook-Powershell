<#
    OUTLOOK POWERSHELL MODULE

    This module uses the COM Interop functionality to directly communicate 
    with the outlook application providing cmdlets to enable
    integration into workflows.
#>
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

Push-Location $PSScriptRoot

Get-ChildItem -Filter "*.ps1" -Recurse | 
	ForEach-Object {
		#Write-Host "Loading $($_.Name) ..."
		. ($_.Fullname)
}

Pop-Location
