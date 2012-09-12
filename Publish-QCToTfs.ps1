#requires -version 2.0
param (
    [parameter(Mandatory=$true)] [Uri]$CollectionUri,
    [parameter(Mandatory=$true)] [string]$ProjectName,
    [parameter(Mandatory=$true)] [string]$QCExportPath
)

$ErrorActionPreference = "Stop"

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

$QCWorkItemsInTfsQueryText = "
    SELECT Id
    FROM WorkItems
    WHERE [Team Project] = '$ProjectName'
    AND [Work Item Type] = 'Bug'
    AND Title contains 'QC'"
$QCWorkItemsInTfsCount = $WorkItemStore.QueryCount($QCWorkItemsInTfsQueryText)
Write-Verbose "$QCWorkItemsInTfsCount QC-related work items found in TFS"

$CountProcessed = 0
$QCWorkItemsInTfs = $WorkItemStore.Query($QCWorkItemsInTfsQueryText) |
    %{
        $CountProcessed++
        Write-Progress -Activity "Retrieving work items from TFS" -PercentComplete ($CountProcessed / $QCWorkItemsInTfsCount * 100)

        if (-not ($_.Title -match '^QC (?<QCId>\d+)')) {
            Write-Error "TFS Work Item $($_.Id) does not have a readable QC Id. Update the work item title to start with something like `"QC 123 - `", or remove the text `"QC`" from the title entirely. This work item will be ignored during processing."
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

$QCWorkItemsInTfs

if (-not (Test-Path $QCExportPath)) {
    throw "We couldn't find or access the QC export file that is meant to be at $QCExportPath"
}
$QCExportPath = (Resolve-Path $QCExportPath).Path

function Get-ExcelHeaderRow($worksheet) {
    Write-Verbose "Reading the Excel header row"
    $columnIndex = 1
    $lastColumnValue = $null
    $columnValues = @()
    do
    {
        $lastColumnValue = $worksheet.Cells.Item(1, $columnIndex).Value()
        if ($lastColumnValue -ne $null) {
            $columnValues += $lastColumnValue
        }
        $columnIndex++
    }
    until ($lastColumnValue -eq $null)
    $columnValues
}

function Get-ExcelDataRow($worksheet, $headers, $rowIndex) {
    Write-Verbose "Reading Excel data row $rowIndex"
    $data = @{}
    for ($columnIndex = 0; $columnIndex -lt $headers.Length; $columnIndex++) {
        $columnName = $headers[$columnIndex]
        Write-Verbose "Reading ($rowIndex,$($columnIndex + 1)): $columnName"
        $data[$columnName] = $worksheet.Cells.Item($rowIndex, $columnIndex + 1).Value()
    }
    New-Object PSObject -Property $data
}

$Excel = New-Object -ComObject excel.application
$QCWorkbook = $Excel.Workbooks.Open($QCExportPath)
Write-Verbose "Opened the QC workbook in Excel"
$QCWorksheet = $QCWorkbook.Worksheets.Item(1)

$QCHeaderRow = Get-ExcelHeaderRow $QCWorksheet
(Get-ExcelDataRow $QCWorksheet $QCHeaderRow 2).Status
