[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [string]$OutputFilePath = "ExportTestContentDatabase.csv"
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

    foreach ($webApp in $config.WebApplications) {
        ExportTestContentDatabase -webappUrl $webApp -outputFilePath $OutputFilePath -sep $config.Sep
    }
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}


function ExportTestContentDatabase {
    param(
        [string]$webappUrl,
        [string]$outputFilePath,
        [string]$sep
    )

    $outputObj = @() 

    Get-SPContentDatabase -WebApplication $webappUrl | ForEach-Object {
        Write-Verbose "Process $webappUrl / $($_.Name)..."
        $problems = Test-SPContentDatabase -Name $_.Name -WebApplication $webappUrl
        $problems | Where-Object { $_.Category -eq "MissingFeature" } | ForEach-Object {
            $obj = New-Object PSObject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Category -Value $_.Category
            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentDb -Value ($_.Message | Select-String -Pattern "Database \[(?<db>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["db"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name RefCount -Value ""
            Add-Member -InputObject $obj -MemberType NoteProperty -Name File -Value ($_.Message | Select-String -Pattern "missing feature: Id = \[(?<feature>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["feature"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Id -Value ""
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Class -Value ""
            return $obj
        } | ForEach-Object {
            $outputObj += $_
        }

        $problems | Where-Object { $_.Category -eq "MissingSetupFile" } | ForEach-Object {
            $obj = New-Object PSObject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Category -Value $_.Category
            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentDb -Value ($_.Message | Select-String -Pattern "in the database \[(?<db>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["db"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name RefCount -Value ($_.Message | Select-String -Pattern "is referenced \[(?<refcount>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["refcount"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name File -Value ($_.Message | Select-String -Pattern "File \[(?<filename>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["filename"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Id -Value ""
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Class -Value ""
            return $obj
        } | ForEach-Object {
            $outputObj += $_
        }


        $problems | Where-Object { $_.Category -eq "MissingAssembly" } | ForEach-Object {
            $obj = New-Object PSObject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Category -Value $_.Category
            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentDb -Value ($_.Message | Select-String -Pattern "in the database \[(?<db>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["db"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name RefCount -Value ""
            Add-Member -InputObject $obj -MemberType NoteProperty -Name File -Value ($_.Message | Select-String -Pattern "Assembly \[(?<assembly>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["assembly"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Id -Value ""
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Class -Value ""
            return $obj
        } | ForEach-Object {
            $outputObj += $_
        }


        $problems | Where-Object { $_.Category -eq "MissingWebPart" } | ForEach-Object {
            $obj = New-Object PSObject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Category -Value $_.Category
            Add-Member -InputObject $obj -MemberType NoteProperty -Name ContentDb -Value ($_.Message | Select-String -Pattern "in the database \[(?<db>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["db"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name RefCount -Value ($_.Message | Select-String -Pattern "is referenced \[(?<refcount>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["refcount"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name File -Value ($_.Message | Select-String -Pattern "from assembly \[(?<assembly>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["assembly"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Id -Value ($_.Message | Select-String -Pattern "WebPart class \[(?<id>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["id"].Value })
            Add-Member -InputObject $obj -MemberType NoteProperty -Name Class -Value ($_.Message | Select-String -Pattern " \(class \[(?<class>[^\[]*)\]" | Select-Object -Expand Matches | % { $_.Groups["class"].Value })
            return $obj
        } | ForEach-Object {
            $outputObj += $_
        }
    }

    Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
    $outputObj | Export-CSV -Path $outputFilePath -NoTypeInformation -Append -Delimiter $sep
    Write-Host " [OK]" -ForegroundColor Green
}