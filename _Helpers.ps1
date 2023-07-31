Function Invoke-Start($scriptName, $currentDir) {
    Write-Host ""
    Write-Host -ForegroundColor Green "-------------------------------------------------"
    Write-Host -ForegroundColor Green " Start script: $scriptName"
    Write-Host -ForegroundColor Green "-------------------------------------------------"

    $Global:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

    $currentDate = (Get-Date).ToString("yyyyMMdd-HHmmss")

    try {
        $logsPath = "$currentDir\Logs"
        $global:LogFilePath = [string]::Format("$logsPath\{0}_{1}.log", $currentDate, $scriptName)
        if (!(Test-Path -Path $logsPath)) {
            New-Item -ItemType directory -Path $logsPath | Out-Null
        }
    }
    catch {
        Write-Host "ERROR New-Item $logsPath" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    try {
        Start-Transcript $global:LogFilePath
    }
    catch {
    }
}

Function Invoke-Stop() {   
    $Global:Stopwatch.Stop()
    $totalSecs = [Math]::Round($Global:Stopwatch.Elapsed.TotalSeconds, 0)

    try {
        Stop-Transcript
    }
    catch {
    }

    Write-Host ""
    Write-Host -ForegroundColor Green "-------------------------------------------------"
    Write-Host -ForegroundColor Green " Executing time: $totalSecs s"
    Write-Host -ForegroundColor Green "-------------------------------------------------"
}

Function Get-Config($env) {
    $configFile = "Config.$env.json"
    if (!(Test-Path -Path $configFile)) {
        throw "Configuration file [$configFile] does not exist!"
    }

    Write-Host "Get-Content $configFile" -NoNewline:$True
    $json = Get-Content $configFile | ConvertFrom-Json
    Write-Host " [OK]" -ForegroundColor Green
    return $json
}