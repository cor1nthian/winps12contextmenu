<#

    .SYNOPSIS
    Adds PowerShell script to explorer context menu

    .DESCRIPTION
    Adds PowerShell script to explorer context menu

    .INPUTS
    PAth to the script (mandatory) and context menu title (optional)

    .OUTPUTS
    None

    .PARAMETER ScriptPath
    Mandatory. 0 position. Path to the script file

    .PARAMETER ContextMenuOption
    Optional. 1 position. Default - "Find files by content". Context menu option name

#>

Param ( [Parameter(Position = 0, Mandatory = $true)]  [System.String] $ScriptPath,
        [Parameter(Position = 1, Mandatory = $false)] [System.String] $ContextMenuOption = "Find files by content" )

############### SCRIPT ###############
[System.String] $psExe = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
[System.String[]] $subKeys = @( 'Directory', 'Directory\Background', 'Drive' )

foreach ($keyName in $subKeys) {
    [System.String] $path = "Registry::HKEY_CLASSES_ROOT\$keyName\shell\PowerShellCustomScript"
    [System.Void] (New-Item -Path $path -Name 'command' -Force -ErrorAction Stop)
    Set-ItemProperty -Path $path -Name '(default)' -Value $ContextMenuOption
    Set-ItemProperty -Path $path -Name 'Icon' -Value "$psExe,0"
    if($PSVersionTable.PSVersion.Major -eq 2) {
        [System.String] $runStr = "$psExe -NoProfile -File `"$ScriptPath`" `"%V`""
    } else {
        [System.String] $runStr = "$psExe -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" `"%V`""
    }
    Set-ItemProperty -Path "$path\command" -Name '(default)' -Value $runStr
}
