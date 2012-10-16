#requires -version 2.0
param (
    [parameter(Mandatory=$true)] [Uri]$CollectionUri,
    [parameter(Mandatory=$true)] [string]$ProjectName,
    [parameter(Mandatory=$true)] [string]$QCExportPath,
    [string] $QCPrefix,
    [switch] $Fix = $false
)

$ErrorActionPreference = "Stop"

function Import-Excel($path, $sheetName)
{
    Write-Verbose "Reading $sheetName from Excel sheet $path"
    $connection = New-Object System.Data.OleDb.OleDbConnection "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=`"$path`";Extended Properties=`"Excel 12.0 Xml;HDR=YES;IMEX=1`""
    $connection.Open()
    $query = "SELECT * FROM [$sheetName`$]"
    $command = New-Object System.Data.OleDb.OleDbCommand @($query, $connection)
    $adapter = New-Object System.Data.OleDb.OleDbDataAdapter
    $adapter.SelectCommand = $command
    $table = New-Object System.Data.DataTable $sheetName
    $rowCount = $adapter.Fill($table)
    Write-Verbose "Read $rowCount rows from Excel sheet"
    $connection.Close()
    $table
}

Add-Type -AssemblyName 'Microsoft.TeamFoundation.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
Add-Type -AssemblyName 'Microsoft.TeamFoundation.WorkItemTracking.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'

$Collection = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($CollectionUri)
Write-Debug "About to attempt a connection to TFS"
$Collection.EnsureAuthenticated()
Write-Verbose "Successfully connected and authenticated to TFS"

$WorkItemStore = $Collection.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])

$ProjectFound = @($WorkItemStore.Projects | Where-Object { $_.Name -eq $ProjectName }).Length -eq 1
if ($ProjectFound) {
    Write-Verbose "Found a team project named $ProjectName"
} else {
    throw "We connected to the TFS collection successfully, but couldn't find a project named $ProjectName. Did you use the right collection URI ($CollectionUri), and spell the project name correctly?"
}

$TfsWorkItemsQueryText = "
    SELECT Id
    FROM WorkItems
    WHERE [Team Project] = '$ProjectName'
    AND [Work Item Type] = 'Bug'
    AND Title contains 'QC'"
$TfsWorkItemsCount = $WorkItemStore.QueryCount($TfsWorkItemsQueryText)
Write-Verbose "$TfsWorkItemsCount QC-related work items found in TFS"

$CountProcessed = 0
$TfsWorkItems = $WorkItemStore.Query($TfsWorkItemsQueryText) |
    %{
        $CountProcessed++
        Write-Progress -Activity "Retrieving work items from TFS" -PercentComplete ($CountProcessed / $TfsWorkItemsCount * 100)

        if (-not ($_.Title -match '^(?<QCPrefix>.*?)\s*QC\s*(?<QCId>\d+)')) {
            Write-Warning "TFS Work Item $($_.Id) does not have a readable QC Id. The current title is `"$($_.Title)`". Update the work item title to start with something like `"QC 123 - `", or remove the text `"QC`" from the title entirely. This work item will be ignored during processing."
            return;
        }
        if ($QCPrefix -ne $matches["QCPrefix"]) {
            Write-Debug "Ignoring TFS Work Item $($_.Id) because prefix doesn't match"
            return;
        }
        $QCId = [int]$matches["QCId"]

        New-Object PSObject -Property @{
            "QCId" = $QCId;
            "TfsId" = $_.Id;
            "TfsState" = $_.State;
            "TfsAssignedTo" = $_["Assigned To"];
        }
    } |
    Sort-Object -Property QCId
Write-Progress -Activity "Retrieving work items from TFS" -Complete

if (-not (Test-Path $QCExportPath)) {
    throw "We couldn't find or access the QC export file that is meant to be at $QCExportPath"
}
$QCExportPath = (Resolve-Path $QCExportPath).Path

$DefectsInQC = Import-Excel $QCExportPath "Sheet1"

Write-Verbose "$($DefectsInQC.Length) defects found in QC export"

$CountProcessed = 0
$SyncIssuesFound = 0
$DefectsInQC | `
    %{
        $CountProcessed++
        Write-Progress -Activity "Processing QC defects" -PercentComplete ($CountProcessed / $DefectsInQC.Length * 100)

        $QCDefect = $_
        $QCId = [int]$QCDefect["Defect ID"]
        $TfsWorkItemsForThisQC = @($TfsWorkItems | ` Where-Object { $_.QCId -eq $QCId })

        $OpenTfsWorkItemsForThisQC = @($TfsWorkItemsForThisQC | ` Where-Object {
            $_.TfsState -ne 'Removed' -and
            $_.TfsState -ne 'Done'
        })

        if ($TfsWorkItemsForThisQC.Length -eq 0) {
            $SyncIssuesFound++
            "QC $QCId is not tracked in TFS at all"
            $Title = "QC $QCId - $($QCDefect.Summary)"
            if ($Fix -eq $true) {
                Write-Verbose "$Title"
            }
        }
        elseif ($OpenTfsWorkItemsForThisQC.Length -gt 1) {
            $SyncIssuesFound++
            $DuplicateTfsIds = $OpenTfsWorkItemsForThisQC | Select-Object -ExpandProperty TfsId
            "QC $QCId is tracked by multiple open TFS work items: $DuplicateTfsIds"
        }
    }
Write-Progress -Activity "Processing QC defects" -Complete

"Found $SyncIssuesFound sync issues across $($DefectsInQC.Length) supplied QC defects and $TfsWorkItemsCount QC-related TFS work items"
