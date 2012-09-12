#requires -version 2.0
param (
    [parameter(Mandatory=$true)] [Uri]$CollectionUri,
    [parameter(Mandatory=$true)] [string]$ProjectName
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
            Write-Error "TFS Work Item $($_.Id) does not have a readable QC Id; ignoring"
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
