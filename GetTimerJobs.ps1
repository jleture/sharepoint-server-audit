[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [string]$OutputFilePath = "ExportTimerJobs.csv"
)

$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
try {
    . ("$ScriptDirectory\_Helpers.ps1")
}
catch {
    Write-Error "Error while loading PowerShell scripts" 
    Write-Error $_.Exception.Message
}

Invoke-Start $MyInvocation.MyCommand.Name $ScriptDirectory

function ExportTimerJobs {
    param(
        [string]$outputFilePath,
        [string]$sep
    )

    $outputObj = @() 

    Get-SPTimerJob | Where-Object { $_.TypeName -notlike "Microsoft*" } | ForEach-Object {
        Write-Verbose "$($_.Name)"
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty Name $_.Name
        $obj | Add-Member NoteProperty TypeName $_.TypeName
        $obj | Add-Member NoteProperty DisplayName $_.DisplayName
        $obj | Add-Member NoteProperty Status $_.Status
        $obj | Add-Member NoteProperty LastRunTime $_.LastRunTime
        $outputObj += $obj
    }

    Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
    $outputObj | Export-CSV -Path $outputFilePath -NoTypeInformation -Append -Delimiter $sep -Encoding UTF8
    Write-Host " [OK]" -ForegroundColor Green
}

try {
    $config = Get-Config $Env
    $config

    $OutputFilePath = "$($config.OutputDir)/$OutputFilePath"
    ExportTimerJobs -outputFilePath $OutputFilePath -sep $config.Sep
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}