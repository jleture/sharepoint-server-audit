[CmdletBinding()]
param(
    [string]$Env = "LAB",
    [string]$OutputFilePath = "ExportFormsAndWorkflows.csv"
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

function ExportFormsAndWorkflows {
    param(
        [string]$webappUrl,
        [string]$outputFilePath,
        [string]$sep
    )

    Write-Verbose "$webappUrl"
    $webApp = Get-SPWebApplication $webappUrl
    $Sites = $webApp | Get-SPSite -Limit ALL -ErrorAction SilentlyContinue

    $outputObj = @() 

    $Sites | Get-SPWeb -Limit ALL -ErrorAction SilentlyContinue | ForEach-Object {
        $Web = $_

        $wfManager = New-object Microsoft.SharePoint.WorkflowServices.WorkflowServicesManager($Web)
        if ($null -ne $wfManager) {
            $wfsService = $wfManager.GetWorkflowSubscriptionService()
        }

        for ($i = 0; $i -ne $_.Lists.count; $i++) {
            $list = $_.Lists[$i]
            Write-Verbose " - $($list.Title)"

            Write-Verbose "  - InfoPath"
            try {
                if ($list.ContentTypes[0].ResourceFolder.Properties["_ipfs_infopathenabled"]) {
                    $obj = New-Object PSObject
                    $obj | Add-Member NoteProperty Type "InfoPath" 
                    $obj | Add-Member NoteProperty Site $($List.ParentWeb.Title)
                    $obj | Add-Member NoteProperty URL $($List.ParentWeb.URL)
                    $obj | Add-Member NoteProperty Title $($List.Title)
                    $obj | Add-Member NoteProperty Template $list.ContentTypes[0].ResourceFolder.Properties["_ipfs_solutionName"]
                    $outputObj += $obj
                }  
            }
            catch {}

            Write-Verbose "  - InfoPathXML"
            try {
                if ($list.BaseTemplate -eq "XMLForm" -and $list.BaseType -eq "DocumentLibrary") {
                    $obj = New-Object PSObject
                    $obj | Add-Member NoteProperty Type "InfoPathXML" 
                    $obj | Add-Member NoteProperty Site $($List.ParentWeb.Title)
                    $obj | Add-Member NoteProperty URL $($List.ParentWeb.URL)
                    $obj | Add-Member NoteProperty Title $($List.Title)
                    $obj | Add-Member NoteProperty Template ""
                    $outputObj += $obj
                }
            }
            catch {}

            Write-Verbose "  - Workflow"
            try {
                foreach ($wf in $list.WorkflowAssociations) {
                    $obj = New-Object PSObject 
                    $obj | Add-Member NoteProperty Type "Workflow" 
                    $obj | Add-Member NoteProperty Site $($List.ParentWeb.Title)
                    $obj | Add-Member NoteProperty URL $($List.ParentWeb.URL)
                    $obj | Add-Member NoteProperty Title $($List.Title)
                    $obj | Add-Member NoteProperty Template $($wf.BaseTemplate)
                    $obj | Add-Member NoteProperty WorkflowName $($wf.Name)
                    $obj | Add-Member NoteProperty AssociationData $($wf.AssociationData)
                    $obj | Add-Member NoteProperty Enabled $($wf.Enabled)
                    $outputObj += $obj
                }
            }
            catch {}

            try {
                if($null -ne $wfsService) {
                    $subscriptions = $wfsService.EnumerateSubscriptionsByList($list.ID)
                    if ($null -ne $subscriptions) {
                        Write-Verbose "  - Workflow2013"
                        foreach ($subscription in $subscriptions) {
                            if (($Web.Url + $list.Title + $subscriptions.Name) -ne $currentItem) {
                                $currentItem = $Web.Url + $list.Title + $subscription.Name   
                                $wfID = $subscription.PropertyDefinitions["SharePointWorkflowContext.ActivationProperties.WebId"]       

                                $obj = New-Object PSObject
                                $obj | Add-Member NoteProperty Type "Workflow2013" 
                                $obj | Add-Member NoteProperty Site $($List.ParentWeb.Title)
                                $obj | Add-Member NoteProperty URL $($List.ParentWeb.URL)
                                $obj | Add-Member NoteProperty Title $($List.Title)
                                $obj | Add-Member NoteProperty Template ""
                                $obj | Add-Member NoteProperty WorkflowName $($subscription.Name)
                                $obj | Add-Member NoteProperty WorkflowId $wfID
                                $outputObj += $obj
                            }
                        }
                    }
                }
            }
            catch {}
            }

        try {
            $nintexList = $_.Lists["NintexForms"]
            if ($null -ne $nintexList) {
                Write-Verbose "  - NintexForms"
                $obj = New-Object PSObject 
                $obj | Add-Member NoteProperty Type "NintexForms" 
                $obj | Add-Member NoteProperty Site $($List.ParentWeb.Title)
                $obj | Add-Member NoteProperty URL $($List.ParentWeb.URL)
                $obj | Add-Member NoteProperty Title $($List.Title)
                $obj | Add-Member NoteProperty Template ""
                $obj | Add-Member NoteProperty Count $($($nintexList.Items | Where-Object { $_.ContentType.Name -eq "Document" } | Where-Object { $_.level -notcontains "Draft" }).Count)
                $outputObj += $obj
            }
        }
        catch {}
    }

    Write-Host "Export-CSV to $outputFilePath" -NoNewline:$True
    $outputObj | Export-CSV -Path $outputFilePath -NoTypeInformation -Append -Delimiter $sep -Encoding UTF8
    Write-Host " [OK]" -ForegroundColor Green
}

try {
    $config = Get-Config $Env
    $config

    $OutputFilePath = "$($config.OutputDir)/$OutputFilePath"
    foreach ($webApp in $config.WebApplications) {
        ExportFormsAndWorkflows -webappUrl $webApp -outputFilePath $OutputFilePath -sep $config.Sep
    }
}
catch {
    Write-Error $_
}
finally {
    Invoke-Stop
}