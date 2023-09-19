[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [string]$OutputFilePath = "ExportSitesSize.csv"
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

function ExportSitesSize {
    param(
        [string]$webappUrl,
        [string]$outputFilePath,
        [string]$sep
    ) 

    Write-Verbose "$webappUrl"
    $webApp = Get-SPWebApplication $webappUrl
    $Sites = $webApp | Get-SPSite -Limit ALL -ErrorAction SilentlyContinue
 
    $outputObj = @() 

    foreach ($Site in $Sites) {
        Write-Verbose " - $($Site.URL)"
        $obj = New-Object PSObject
        $obj | Add-Member -type NoteProperty -name "Site URL" -value $Site.Url
        $obj | Add-Member -type NoteProperty -name "Database" -value $Site.ContentDatabase.Name
        $obj | Add-Member -type NoteProperty -name "SizeInMB" -value ($Site.Usage.Storage/1024)
        $outputObj += $obj
    }

    Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
    $outputObj | Export-CSV -Path $outputFilePath -NoTypeInformation -Append -Delimiter $sep
    Write-Host " [OK]" -ForegroundColor Green
}

try {
    $config = Get-Config $Env
    $config

    $OutputFilePath = "$($config.OutputDir)/$OutputFilePath"
    foreach ($webApp in $config.WebApplications) {
        ExportSitesSize -webappUrl $webApp -outputFilePath $OutputFilePath -sep $config.Sep -Encoding UTF8
    }
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}