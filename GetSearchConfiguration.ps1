[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [string]$OutputFilePathCrawl = "ExportSearchCrawl-{0}.csv",
    [string]$OutputFilePathManagedProperties = "ExportSearchManagedProperties.csv"
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

    ExportSearchCrawlDuration -outputFilePath $OutputFilePathCrawl -sep $config.Sep

    ExportSearchManagedProperties -outputFilePath $OutputFilePathManagedProperties -sep $config.Sep
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}


function ExportSearchCrawlDuration {
    param(
        [string]$outputFilePath,
        [string]$sep
    )

    $outputObj = @() 

    $ssa = Get-SPEnterpriseSearchServiceApplication
    $sources = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $ssa

    foreach ($source in $sources) {
        Write-Verbose "$($source.Name)"
        $log = New-Object Microsoft.Office.Server.Search.Administration.CrawlLog $ssa
        $last = $log.GetCrawlHistory(10, $source.Id)[0]

        $outputFilePath = $outputFilePath -f $source.Name
        Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
        $last | Export-CSV -Path $outputFilePath -NoTypeInformation -Append -Delimiter $sep
        Write-Host " [OK]" -ForegroundColor Green
    }
}

function ExportSearchManagedProperties {
    param(
        [string]$outputFilePath,
        [string]$sep
    )

    $ssa = Get-SPEnterpriseSearchServiceApplication

    $outputObj = Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $ssa | Where-Object { $_.SystemDefined -eq $false } | ForEach-Object {
        New-Object -TypeName PSObject -Property @{
            PID         = $_.PID
            Name        = $_.Name
            ManagedType = $_.ManagedType 
        }
    }

    Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
    $outputObj | Export-CSV -Path $outputFilePath -NoTypeInformation -Append -Delimiter $sep
    Write-Host " [OK]" -ForegroundColor Green
}