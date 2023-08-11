[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [string]$OutputFilePath = "SharePointWebParts.csv"
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

function ProcessWebApp {
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
        foreach ($Web in $Site.AllWebs) {
            Write-Verbose "   - $($Web.URL)"
            if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($Web)) {
                $PubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($Web)
                $PagesLib = $PubWeb.PagesList
            }
            else {
                $PagesLib = $Web.Lists["Site Pages"]
            }             
            Write-Verbose "     - $($PagesLib.Title)"

            foreach ($Page in $PagesLib.Items | Where-Object { $_.Name -match ".aspx" }) {
                $PageURL = $web.site.Url + "/" + $Page.File.URL
                $WebPartManager = $Page.File.GetLimitedWebPartManager([System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
                 
                foreach ($WebPart in $WebPartManager.WebParts) {
                    $obj = New-Object PSObject
                    $obj | Add-Member -type NoteProperty -name "Site URL" -value $web.Url
                    $obj | Add-Member -type NoteProperty -name "Page URL" -value $PageURL
                    $obj | Add-Member -type NoteProperty -name "Web Part Title" -value $WebPart.Title
                    $obj | Add-Member -type NoteProperty -name "Web Part Type" -value $WebPart.GetType().ToString()
                    $outputObj += $obj
                }
            }
        }

        Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
        $outputObj | Export-csv $outputFilePath -NoTypeInformation -Append -Delimiter $sep -Encoding UTF8
        Write-Host " [OK]" -ForegroundColor Green
    }
}

try {
    $config = Get-Config $Env
    $config

    $OutputFilePath = "$($config.OutputDir)/$OutputFilePath"
    foreach ($webApp in $config.WebApplications) {
        ProcessWebApp -webappUrl $webApp -outputFilePath $OutputFilePath -sep $config.Sep
    }
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}