# sharepoint-server-audit
PowerShell scripts to audit SharePoint Server sites

## Files
| File | Role |
| - | - |
| **_Helpers.ps1** | Useful methods |
| **Config.Lab.json** | Configuration file with tenant and SharePoint URL, app registration and CSV separator |
| **GetSharePointWebParts.ps1** | Script to get webparts used in SharePoint sites |
| **GetTestContentDatabases.ps1** | Script to get errors in content databases |
| **GetFormsAndWorkflows.ps1** | Script to get worflows and custom forms |
| **GetHealthData.ps1** | Script to get errors from the central administration health reports  |
| **GetTimerJobs.ps1** | Script to get custom timer jobs |
| **GetSearchConfiguration.ps1** | Script to get search managed properties and the last crawl duration |

## Prerequisities

## Configuration

Create a new configuration file based on `Config.LAB.json` or edit this one.

When executing the scripts, the code-name of the configuration should be passed as an argument:

~~~powershell
.\GetSharePointWebParts.ps1 -Env LAB
.\GetSharePointWebParts.ps1 -Env PROD
~~~

To connect to SharePoint, it's better to XXX.

~~~json
{
    "OutputDir": "C:/Temp",
    "Sep" : ";",
    "WebApplications": ["https://your-webapp.local", "https://your-mysites.local"]
}
~~~


## Get webparts [GetSharePointWebParts.ps1]

The script has three steps:
1. Get SharePoint sites
2. For each site, open every pages and get webparts
3. Generate a CSV file with webparts, pages and sites

The CSV files are generated in the folder specified in the configuration file (`OutputDir`).

~~~powershell
.\GetSharePointWebParts.ps1 -Env LAB -OutputFilePath "SharePointWebParts.csv"
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetSharePointWebParts.ps1 -Env LAB -OutputFilePath "SharePointWebParts.csv" -Verbose
~~~

## Get custom forms and workflows [GetFormsAndWorkflows.ps1]

The script detects 3 objects:
1. InfoPath forms
2. Nintex forms
3. Workflows (custom, buit-in, Nintex, with the 2010 or 2013 engine)

The CSV files are generated in the folder specified in the configuration file (`OutputDir`).

~~~powershell
.\GetFormsAndWorkflows.ps1 -Env LAB -OutputFilePath "ExportFormsAndWorkflows.csv"
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetFormsAndWorkflows.ps1 -Env LAB -OutputFilePath "ExportFormsAndWorkflows.csv" -Verbose
~~~

## Get custom timer jobs [GetTimerJobs.ps1]

Returns the custom timer jobs (not `Microsoft`) deployed on the farm.

~~~powershell
.\GetTimerJobs.ps1 -Env LAB -OutputFilePath "ExportTimerJobs.csv"
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetTimerJobs.ps1 -Env LAB -OutputFilePath "ExportTimerJobs.csv" -Verbose
~~~

## Get errors on content databases [GetTestContentDatabases.ps1]

~~~powershell
.\GetTestContentDatabases.ps1 -Env LAB -OutputFilePath "ExportTestContentDatabase.csv"
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetTestContentDatabases.ps1 -Env LAB -OutputFilePath "ExportTestContentDatabase.csv" -Verbose
~~~

## Get search managed properties and crawl duration [GetSearchConfiguration.ps1]

The script returns the search managed properties and the last crawl duration for each sources.

For the crawl duration, there are one CSV file generated per sources with the last 10 executions. The CSV filename must include a pattern to inject the source name (example: `ExportSearchCrawl-{0}.csv`)

~~~powershell
.\GetSearchConfiguration.ps1 -Env LAB -OutputFilePathCrawl "ExportSearchCrawl-{0}.csv" -OutputFilePathManagedProperties "ExportSearchManagedProperties.csv"
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetSearchConfiguration.ps1 -Env LAB -OutputFilePathCrawl "ExportSearchCrawl-{0}.csv" -OutputFilePathManagedProperties "ExportSearchManagedProperties.csv" -Verbose
~~~


## Get errors from the central administration health reports [GetHealthData.ps1]

~~~powershell
.\GetHealthData.ps1 -Env LAB -OutputFilePath "ExportHealthData.csv"
~~~

You can add `-Verbose` to display more information in the terminal.

~~~powershell
.\GetHealthData.ps1 -Env LAB -OutputFilePath "ExportHealthData.csv" -Verbose
~~~