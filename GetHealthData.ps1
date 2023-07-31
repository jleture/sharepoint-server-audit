[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [string]$OutputFilePath = "ExportHealthData.csv"
)

$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$OutputFilePath = "$ScriptDirectory/$OutputFilePath"
try {
    . ("$ScriptDirectory\_Helpers.ps1")
}
catch {
    Write-Error "Error while loading PowerShell scripts" 
    Write-Error $_.Exception.Message
}

Invoke-Start $MyInvocation.MyCommand.Name $ScriptDirectory

try {
    $config = Get-Config $Env
    $config

    ExportHealthData -outputFilePath $OutputFilePath -sep $config.Sep
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}


function ExportTimerJobs {
    param(
        [string]$outputFilePath,
        [string]$sep
    )

    $outputObj = @() 

    $caWebApp = (Get-SPWebApplication -IncludeCentralAdministration) | Where-Object { $_.IsAdministrationWebApplication -eq $true }
    $caWeb = Get-SPWeb -Identity $caWebApp.Url
    Write-Verbose "Central administration is [$($caWebApp.Url)]"
    $healthList = $caWeb.GetList("\Lists\HealthReports")

    $Query = New-Object Microsoft.SharePoint.SPQuery
    $Query.ViewAttributes = "Scope='Recursive'"
    $Query.Query = "<Where><Or><BeginsWith><FieldRef Name='HealthReportSeverityIcon'/><Value Type='Text'>1</Value></BeginsWith><BeginsWith><FieldRef Name='HealthReportSeverityIcon'/><Value Type='Text'>2</Value></BeginsWith></Or></Where>"

    $ListItems = $healthList.GetItems($Query)
    foreach ($spListItem in $ListItems) {
        Write-Verbose "$($spListItem['Title'])"
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty Severity $spListItem["Severity"]
        $obj | Add-Member NoteProperty HealthReportCategory $spListItem["HealthReportCategory"]
        $obj | Add-Member NoteProperty HealthReportServices $spListItem["HealthReportServices"]
        $obj | Add-Member NoteProperty Title $spListItem["Title"]
        $obj | Add-Member NoteProperty Created $spListItem["Created"]
        $obj | Add-Member NoteProperty Modified $spListItem["Modified"]
        $obj | Add-Member NoteProperty HealthReportServers $spListItem["HealthReportServers"]
        $outputObj += $obj
    }

    Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
    $outputObj | Export-CSV -Path $outputFilePath -NoTypeInformation -Append -Delimiter $sep
    Write-Host " [OK]" -ForegroundColor Green
}