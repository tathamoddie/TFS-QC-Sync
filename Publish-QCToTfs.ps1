#requires -version 2.0
param (
    [parameter(Mandatory=$true)] [Uri]$CollectionUri,
    [parameter(Mandatory=$true)] [string]$ProjectName,
    [parameter(Mandatory=$true)] [string]$QCExportPath,
    [parameter(Mandatory=$true)] $IterationMapping,
    [string] $QCPrefix,
    [switch] $Fix = $false
)

$ErrorActionPreference = "Stop"

$DescriptionField = "Microsoft.VSTS.TCM.ReproSteps"
$CommentsField = "Microsoft.VSTS.Common.AcceptanceCriteria"

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

function New-BugInTfs($WorkItemType, $QCDefect)
{
    $WorkItem = New-Object Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem $WorkItemType
    $WorkItem["Title"] = Format-TfsWorkItemTitle $QCDefect
    $WorkItem["Severity"] = $DefectToTfsSeverity[$QCDefect.Severity]
    if ($QCDefect.Priority -ne [System.DBNull]::Value) {
        $WorkItem["Microsoft.VSTS.Common.BusinessValue"] = [int]::Parse($QCDefect.Priority[0])
    }
    $WorkItem["$DescriptionField"] = Format-TfsWorkItemTextAsHtml $QCDefect.Description
    $WorkItem["$CommentsField"] = Format-TfsWorkItemTextAsHtml $QCDefect.Comments
    if (-not $WorkItem.IsValid()) {
        $InvalidFieldNames = $WorkItem.Fields | Where-Object { -not $_.IsValid } | %{ "$($_.Name) is $($_.Status) and value is `"$($_.Value)`"" }
        Write-Error "The newly created TFS work item was not valid for saving. Invalid fields were: $InvalidFieldNames" -ErrorAction Continue
    }
    $WorkItem
}

function Format-TfsWorkItemTitle($QCDefect)
{
    $Prefix = ''
    if (-not [string]::IsNullOrWhiteSpace($QCPrefix)) {
        $Prefix = "$($QCPrefix) "
    }
    $QCTitle = $QCDefect.Summary
    if ($QCTitle.Length -gt 150) {
        $QCTitle = "$($QCTitle.Substring(0, 149))…"
    }
    "$($Prefix)QC $($QCDefect["Defect ID"]) - $($QCTitle)"
}

function Format-TfsWorkItemTextAsHtml($Text)
{
    $Text -replace "<", "&lt;" `
          -replace ">", "&gt;" `
          -replace "`n", "<br>"
}

Add-Type -AssemblyName 'Microsoft.TeamFoundation.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'
Add-Type -AssemblyName 'Microsoft.TeamFoundation.WorkItemTracking.Client, Version=11.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'

$Collection = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($CollectionUri)
Write-Debug "About to attempt a connection to TFS"
$Collection.EnsureAuthenticated()
Write-Verbose "Successfully connected and authenticated to TFS"

$WorkItemStore = $Collection.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])

$TfsProject = @($WorkItemStore.Projects | Where-Object { $_.Name -eq $ProjectName })[0]
if ($TfsProject -ne $null) {
    Write-Verbose "Found a team project named $ProjectName"
} else {
    throw "We connected to the TFS collection successfully, but couldn't find a project named $ProjectName. Did you use the right collection URI ($CollectionUri), and spell the project name correctly?"
}

$BugWorkItemType = $TfsProject.WorkItemTypes['Bug']
if ($BugWorkItemType -eq $null) {
    throw "Couldn't find the work item type definition for bugs"
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
            Write-Verbose "Ignoring TFS Work Item $($_.Id) because prefix doesn't match"
            return;
        }
        $QCId = [int]$matches["QCId"]

        New-Object PSObject -Property @{
            "QCId" = $QCId;
            "TfsId" = $_.Id;
            "TfsState" = $_.State;
            "TfsAssignedTo" = $_["Assigned To"];
            "TfsWorkItem" = $_;
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

$DefectToTfsSeverity = @{
    "1-Critical" = "1 - Critical";
    "2-High" = "2 - High";
    "3-Medium" = "3 - Medium";
    "4-Low" = "4 - Low";
}

$QCStatusToTfsStateToNewTfsStateMapping = @{
    "Assigned" = @{
        "Done" = "Committed";
        "Removed" = "New";
    };
    "New" = @{
        "Done" = "Committed";
        "Removed" = "New";
    };
    "Open" = @{
        "Done" = "Committed";
        "Removed" = "New";
    };
    "Fix" = @{
        "Done" = "Committed";
        "Removed" = "New";
    };
    "Analyse" = @{
        "Done" = "Committed";
        "Removed" = "New";
    };
    "Reopened" = @{
        "Done" = "Committed";
        "Removed" = "New";
    };
    "Closed" = @{
        "New" = "Removed";
        "Approved" = "Removed";
        "Committed" = "Done";
    };
    "Fixed" = @{
        "New" = "Removed";
        "Approved" = "Removed";
        "Committed" = "Done";
    };
    "Retest" = @{
        "New" = "Removed";
        "Approved" = "Removed";
        "Committed" = "Done";
    };
    "Deploy" = @{
        "New" = "Removed";
        "Approved" = "Removed";
        "Committed" = "Done";
    };
    "Deferred" = @{
    };
}

$CountProcessed = 0
$SyncIssuesFound = 0
$TfsChanges = @()
$DefectsInQC | `
    %{
        $CountProcessed++
        Write-Progress -Activity "Processing QC defects" -PercentComplete ($CountProcessed / $DefectsInQC.Length * 100)

        $QCDefect = $_
        $QCId = [int]$QCDefect["Defect ID"]
        $TfsWorkItemsForThisQC = @($TfsWorkItems | ` Where-Object { $_.QCId -eq $QCId })

        Write-Verbose "Processing QC $QCId which has $($TfsWorkItemsForThisQC.Length) TFS work items: $($TfsWorkItemsForThisQC | Select-Object -ExpandProperty TfsId)"

        $OpenTfsWorkItemsForThisQC = @($TfsWorkItemsForThisQC | ` Where-Object {
            $_.TfsState -ne 'Removed' -and
            $_.TfsState -ne 'Done'
        })

        Write-Debug 'Assessing item counts'
        if ($TfsWorkItemsForThisQC.Length -eq 0) {
            if (@('Assigned', 'New', 'Open', 'Fix', 'Analyse') -contains $QCDefect.Status) {
                $SyncIssuesFound++
                "QC $QCId is $($QCDefect.Status) but not tracked in TFS at all (should be created)"
                $TfsChanges += New-BugInTfs $BugWorkItemType $QCDefect
            } else {
                Write-Verbose "QC $QCId is not in TFS, but doesn't need to be because state is $($QCDefect.Status)"
            }
            return
        }
        elseif ($OpenTfsWorkItemsForThisQC.Length -gt 1) {
            $SyncIssuesFound++
            $DuplicateTfsIds = $OpenTfsWorkItemsForThisQC | Select-Object -ExpandProperty TfsId
            "QC $QCId is tracked by multiple open TFS work items: $DuplicateTfsIds"
            return
        }

        if ($OpenTfsWorkItemsForThisQC.Length -eq 1) {
            $TfsWorkItem = $OpenTfsWorkItemsForThisQC[0].TfsWorkItem
        }
        else {
            $TfsWorkItem = ($TfsWorkItemsForThisQC | Select-Object -Last 1).TfsWorkItem
        }

        Write-Debug 'Checking title'
        $ExpectedTitle = Format-TfsWorkItemTitle $QCDefect
        Write-Debug "Expected title is $ExpectedTitle"
        Write-Debug "Current title is $($TfsWorkItem["Title"])"
        if ($TfsWorkItem["Title"] -ne $ExpectedTitle) {
            $SyncIssuesFound++
            "TFS $($TfsWorkItem.Id) has title '$($TfsWorkItem["Title"])' (should be '$ExpectedTitle')"
            $TfsWorkItem.Open()
            $TfsWorkItem["Title"] = $ExpectedTitle
            $TfsChanges += $TfsWorkItem
        }

        Write-Debug 'Checking severity'
        $ExpectedSeverity = $DefectToTfsSeverity[$QCDefect.Severity]
        if ($TfsWorkItem["Severity"] -ne $ExpectedSeverity) {
            $SyncIssuesFound++
            "QC $QCId has severity '$($QCDefect.Severity)', but TFS $($TfsWorkItem.Id) has '$($TfsWorkItem["Severity"])' (should be '$ExpectedSeverity')"
            $TfsWorkItem.Open()
            $TfsWorkItem["Severity"] = $ExpectedSeverity
            $TfsChanges += $TfsWorkItem
        }

        Write-Debug 'Checking priority'
        if ($QCDefect.Priority -ne [System.DBNull]::Value) {
            $ExpectedBusinessValue = [int]::Parse($QCDefect.Priority[0])
            if ($TfsWorkItem["Microsoft.VSTS.Common.BusinessValue"] -ne $ExpectedBusinessValue) {
                $SyncIssuesFound++
                "QC $QCId has priority '$($QCDefect.Priority)', but TFS $($TfsWorkItem.Id) has '$($TfsWorkItem["Microsoft.VSTS.Common.BusinessValue"])' (should be '$ExpectedBusinessValue')"
                $TfsWorkItem.Open()
                $TfsWorkItem["Microsoft.VSTS.Common.BusinessValue"] = $ExpectedBusinessValue
                $TfsChanges += $TfsWorkItem
            }
        }

        Write-Debug 'Checking iteration path'
        $ExpectedIterationPath = $IterationMapping[$QCDefect["Detected in Release"]]
        if ((-not $TfsWorkItem["Iteration Path"].Contains("\")) -and
            ($TfsWorkItem["Iteration Path"] -ne $ExpectedIterationPath)) {
            $SyncIssuesFound++
            "QC $QCId was detected in release '$($QCDefect["Detected in Release"])', but TFS $($TfsWorkItem.Id) is in iteration path '$($TfsWorkItem["Iteration Path"])' (should be '$ExpectedIterationPath')"
            $TfsWorkItem.Open()
            $TfsWorkItem["Iteration Path"] = $ExpectedIterationPath
            $TfsChanges += $TfsWorkItem
        }

        Write-Debug 'Checking state'
        if ($QCStatusToTfsStateToNewTfsStateMapping.Contains($QCDefect.Status)) {
            $ExpectedState = $QCStatusToTfsStateToNewTfsStateMapping[$QCDefect.Status][$TfsWorkItem["State"]]
        } else {
            $ExpectedState = $null
            Write-Warning "QC $QCId has status has status '$($QCDefect.Status)' that we didn't know what to do with (relates to TFS $($TfsWorkItem.Id) which is $($TfsWorkItem["State"]))"
        }
        if (($ExpectedState -ne $null) -and
            ($TfsWorkItem["State"] -ne $ExpectedState)) {
            $SyncIssuesFound++
            "QC $QCId has status '$($QCDefect.Status)', but TFS $($TfsWorkItem.Id) has '$($TfsWorkItem["State"])' (should be '$ExpectedState')"
            $TfsWorkItem.Open()
            $TfsWorkItem["State"] = $ExpectedState
            $TfsChanges += $TfsWorkItem
        }

        Write-Debug 'Checking description'
        $ExpectedDescription = Format-TfsWorkItemTextAsHtml $QCDefect.Description
        $ActualDescription = $TfsWorkItem["$DescriptionField"]
        Write-Debug "Expected description is $ExpectedDescription"
        Write-Debug "Current description is $ActualDescription"
        if ($ActualDescription -ne $ExpectedDescription) {
            $SyncIssuesFound++
            "TFS $($TfsWorkItem.Id) has out of date description"
            $TfsWorkItem.Open()
            $TfsWorkItem["$DescriptionField"] = $ExpectedDescription
            $TfsChanges += $TfsWorkItem
        }

        Write-Debug 'Checking comments'
        $ExpectedComments = Format-TfsWorkItemTextAsHtml $QCDefect.Comments
        $ActualComments = $TfsWorkItem["$CommentsField"]
        Write-Debug "Expected comments is $ExpectedComments"
        Write-Debug "Current comments is $ActualComments"
        if ($ActualComments -ne $ExpectedComments) {
            $SyncIssuesFound++
            "TFS $($TfsWorkItem.Id) has out of date comments"
            $TfsWorkItem.Open()
            $TfsWorkItem["$CommentsField"] = $ExpectedComments
            $TfsChanges += $TfsWorkItem
        }
    }
Write-Progress -Activity "Processing QC defects" -Complete

"Found $SyncIssuesFound sync issues across $($DefectsInQC.Length) supplied QC defects and $TfsWorkItemsCount QC-related TFS work items"

Write-Progress -Activity "Publishing $($TfsChanges.Length) changes to TFS" -PercentComplete 0

if ($Fix -eq $true) {
    $TfsChanges = @($TfsChanges | `
        Group-Object -Property Title | `
        %{ $_.Group[0] })

    $SaveErrors = $WorkItemStore.BatchSave($TfsChanges)
    $PublishedTfsIds = $TfsChanges | Select-Object -ExpandProperty Id | Sort-Object
    "Published $($TfsChanges.Length - $SaveErrors.Length) work items to TFS $PublishedTfsIds"
    if ($SaveErrors.Length -ne 0) {
        Write-Error "$($SaveErrors.Length) work items failed to publish to TFS" -ErrorAction Continue
    }
}
else {
    "Skipping $($TfsChanges.Length) changes to TFS because -Fix switch was not supplied"
}

Write-Progress -Activity "Publishing $($TfsChanges.Length) changes to TFS" -Complete
