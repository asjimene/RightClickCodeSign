<#
.SYNOPSIS
    Script that adds a right click menu to sign files using a specified code signing certificate.
.DESCRIPTION
     RightClickToCodeSign is a simple PowerShell script that adds a right click context menu to Code Sign Files using the right click menu.
.EXAMPLE
    Install the Script, select the Code Signing Cert from a menu
    RightClickToCodeSign.ps1 -Install -InstallSelectCert -timestamp "http://timestamp.digicert.com" -Algorithm SHA256

    Install the Script, select the Code Signing Cert using the Subject and Issuer (use * as a wildcard)
    RightClickToCodeSign.ps1 -Install -Subject "Andrews Code Si*" -Issuer "Andrew*" -timestamp "http://timestamp.digicert.com" -Algorithm SHA256
.EXAMPLE
    Uninstall the Script
    RightClickToCodeSign.ps1 -Uninstall
.NOTES
    This script is installed in the User Context
    Created by Andrew Jimenez (@asjimene) 2020-04-12
#>

Param (
    # Specifies a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is
    # used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters,
    # enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any
    # characters as escape sequences.
    [Parameter(Mandatory = $false,
        Position = 0,
        ParameterSetName = "LiteralPath",
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Literal path to one or more locations.")]
    [Alias("PSPath")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $LiteralPath,

    # Uninstall Switch Run this Script with the uninstall Switch to uninstall the Script and remove the Registry changes to your account
    [Parameter(Mandatory = $false, ParameterSetName = "Uninstall")]
    [switch]
    $Uninstall = $false,

    # Install Switch Run this Script with the Install Switch to Install the Script and add the Registry changes to your account
    [Parameter(Mandatory = $false, ParameterSetName = "Install")]
    [Parameter(Mandatory = $false, ParameterSetName = "InstallChooseCert")]
    [switch]
    $Install = $false,

    [Parameter(Mandatory = $true, ParameterSetName = "Install")]
    [Parameter(Mandatory = $true, ParameterSetName = "InstallChooseCert")]
    [String[]]
    $ExtensionList = ".ps1",

    [Parameter(Mandatory = $true, ParameterSetName = "InstallChooseCert")]
    [switch]
    $InstallChooseCert = $false,

    [Parameter(Mandatory = $true, ParameterSetName = "Install")]
    [String]
    $Issuer,

    [Parameter(Mandatory = $true, ParameterSetName = "Install")]
    [String]
    $Subject,

    [Parameter(Mandatory = $true, ParameterSetName = "Install")]
    [Parameter(Mandatory = $true, ParameterSetName = "InstallChooseCert")]
    [String]
    $Timestamp,

    [Parameter(Mandatory = $true, ParameterSetName = "Install")]
    [Parameter(Mandatory = $true, ParameterSetName = "InstallChooseCert")]
    [String]
    $Algorithm
)

$Global:NameOfThisScript = "RightClickToCodeSign"
$Global:ScriptFileName = "CodeSignFile"
$Global:RightClickName = "Code Sign File"

if ($Install) {
    Write-Output "Creating $Global:ScriptFileName folder in LOCALAPPDATA folder"
    New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Name $Global:ScriptFileName -ErrorAction SilentlyContinue

    if ($InstallChooseCert) {
        $SelectedCert = Get-ChildItem 'Cert:\CurrentUser\My' -CodeSigningCert | Out-GridView -Title "Choose the Signing Certificate" -OutputMode Single
        $Subject = $SelectedCert.Subject
        $Issuer = $SelectedCert.Issuer
    }

    Write-Output "Copying Script to $Global:ScriptFileName Folder"
    #Copy-Item "$PSScriptRoot\$($Global:NameOfThisScript).ps1" -Destination "$env:LOCALAPPDATA\$($Global:ScriptFileName)\$($Global:ScriptFileName).ps1" -ErrorAction SilentlyContinue
    (Get-Content "$PSScriptRoot\$($Global:NameOfThisScript).ps1").Replace('_ISSUER_', "$Issuer").Replace('_SUBJECT_', "$Subject").Replace('_TIMESTAMP_', $Timestamp).Replace('_ALGORITHM_', $Algorithm) | Out-File -FilePath "$env:LOCALAPPDATA\$($Global:ScriptFileName)\$($Global:ScriptFileName).ps1" -Encoding utf8 -Force

    foreach ($extension in $ExtensionList) {
        # Reg2CI (c) 2020 by Roger Zander
        if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\$($extension)") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\$($extension)" -force -ea SilentlyContinue };
        if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell" -force -ea SilentlyContinue };
        if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell\$($Global:RightClickName)") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell\$($Global:RightClickName)" -force -ea SilentlyContinue };
        if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell\$($Global:RightClickName)\command") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell\$($Global:RightClickName)\command" -force -ea SilentlyContinue };
        New-ItemProperty -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell\$($Global:RightClickName)" -Name '(default)' -Value "$($Global:RightClickName)" -PropertyType String -Force -ea SilentlyContinue;
        New-ItemProperty -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\$($extension)\shell\$($Global:RightClickName)\command" -Name '(default)' -Value "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$env:LOCALAPPDATA\$($Global:ScriptFileName)\$($Global:ScriptFileName).ps1`" -LiteralPath '%1'" -PropertyType String -Force -ea SilentlyContinue;
    }

    Write-Output "Installation Complete, You should test CodeSigning on the file: `"$env:LOCALAPPDATA\$($Global:ScriptFileName)\$($Global:ScriptFileName).ps1`""
    Pause
}

if ($Uninstall) {
    Write-Output "Removing Script from LOCALAPPDATA"
    Remove-Item "$env:LOCALAPPDATA\$($Global:ScriptFileName)" -Force -Recurse -ErrorAction SilentlyContinue
    $InstalledLocations = (Get-ChildItem "HKCU:\Software\Classes\SystemFileAssociations\*\shell\$($Global:RightClickName)").PSPath
    Write-Output "Cleaning Up Registry"
    foreach ($Location in $InstalledLocations) {
        if ((Test-Path -LiteralPath $Location) -eq $true) { 
            $ShellPath = (Get-Item $Location).PSParentPath
            $ExtensionPath = (Get-Item (Get-Item $Location).PSParentPath).PSParentPath
            Write-Output "Removing $Location"
            Remove-Item $Location -force -Recurse -ea SilentlyContinue 
            if ([System.String]::IsNullOrEmpty((Get-ChildItem $ShellPath))) {
                Write-Output "Removing" $ShellPath
                Remove-Item $ShellPath -force -Recurse -ea SilentlyContinue 
            }

            if ([System.String]::IsNullOrEmpty((Get-ChildItem $ExtensionPath))) {
                Write-Output "Removing" $ExtensionPath
                Remove-Item $ExtensionPath -force -Recurse -ea SilentlyContinue 
            }
        }
    }

    Write-Output "Uninstallation Complete!"
    Pause
}

if ((-not $Install) -and (-not $Uninstall)) {
    $cert = Get-ChildItem Cert:\CurrentUser\My\ -CodeSigningCert | Where-Object Issuer -like '_ISSUER_' | Where-Object Subject -like '_SUBJECT_'
    Set-AuthenticodeSignature -Certificate $cert -TimestampServer '_TIMESTAMP_' -HashAlgorithm _ALGORITHM_ -FilePath "$LiteralPath" 
}
