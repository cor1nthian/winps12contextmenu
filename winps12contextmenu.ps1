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

    .PARAMETER UsePWSH
    Optional. 2 position. Default - false. If true, always use pwsh as script environment placing the script to context menu
    
#>

Param ( [Parameter(Position = 0, Mandatory = $true)]  [System.String]  $ScriptPath,
        [Parameter(Position = 1, Mandatory = $false)] [System.String]  $ContextMenuOption = "Find files by content",
        [Parameter(Position = 2, Mandatory = $false)] [System.Boolean] $UsePWSH = $false )

############### SCRIPT ###############
[System.String] $psExe
if($UsePWSH) {
    $psExe = "$env:PROGRAMFILES\PowerShell\7\pwsh.exe"
} else {
    if(Test-Path "$env:PROGRAMFILES\PowerShell\7\pwsh.exe") {
        $psExe = "$env:PROGRAMFILES\PowerShell\7\pwsh.exe"
    } else {
        if(Test-Path "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe") {
            $psExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
        } else {
            Write-Host -ForegroundColor Red "NO POWERSHELL EXECUTABLE FOUND"
            Write-Host "`nPRESS ANY KEY TO CONTINUE"
            [System.Void][System.Console]::ReadKey($true)
            Exit 1
        }
    }
}
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
