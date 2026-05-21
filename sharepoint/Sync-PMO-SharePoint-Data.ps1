<#
.SYNOPSIS
Exports PMO SharePoint lists to the Vercel portal data file.

.DESCRIPTION
Reads SharePoint lists:
- PMO Projects
- PMO Tasks
- PMO Risks
- PMO Status Reports
- PMO Resources
- PMO Project Files document library

Then writes:
- public/data/sharepoint-data.json

After running this script, commit/push to GitHub. Vercel will redeploy and the portal will show SharePoint data automatically.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl = "https://leaderig.sharepoint.com/sites/Solution",

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "c431f27e-b41c-4488-90b0-5492f8622279",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\public\data\sharepoint-data.json"
)

$ErrorActionPreference = "Stop"

function Write-Step { param([string]$Message) Write-Host "[PMO Sync] $Message" -ForegroundColor Cyan }

function Ensure-PnPModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
}

function Get-ListItemsSafe {
    param([string]$ListName)
    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "List not found: $ListName" -ForegroundColor Yellow
        return @()
    }
    return Get-PnPListItem -List $ListName -PageSize 500
}

function FieldValue {
    param($Item, [string[]]$Names)
    foreach ($name in $Names) {
        if ($Item.FieldValues.ContainsKey($name) -and $null -ne $Item.FieldValues[$name]) {
            $value = $Item.FieldValues[$name]
            if ($value -is [Microsoft.SharePoint.Client.FieldUserValue]) { return $value.LookupValue }
            if ($value -is [Microsoft.SharePoint.Client.FieldUrlValue]) { return $value.Url }
            return "$value"
        }
    }
    return ""
}

function NumberValue {
    param($Value, [double]$Default = 0)
    if ([string]::IsNullOrWhiteSpace("$Value")) { return $Default }
    $clean = ("$Value" -replace "%", "")
    $parsed = 0.0
    if ([double]::TryParse($clean, [ref]$parsed)) { return $parsed }
    return $Default
}

Ensure-PnPModule
Write-Step "Connecting to $SiteUrl"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

Write-Step "Reading SharePoint PMO lists"
$projectItems = Get-ListItemsSafe "PMO Projects"
$taskItems = Get-ListItemsSafe "PMO Tasks"
$riskItems = Get-ListItemsSafe "PMO Risks"
$reportItems = Get-ListItemsSafe "PMO Status Reports"
$resourceItems = Get-ListItemsSafe "PMO Resources"
$fileItems = Get-ListItemsSafe "PMO Project Files"

$projects = @()
foreach ($item in $projectItems) {
    $progress = NumberValue (FieldValue $item @("PMOProgress", "Progress")) 0
    $projects += [ordered]@{
        id = "P-$($item.Id)"
        ProjectName = FieldValue $item @("PMOProjectName", "Title", "ProjectName")
        ProjectManager = FieldValue $item @("PMOProjectManager", "ProjectManager")
        Status = FieldValue $item @("PMOStatus", "Status")
        Health = FieldValue $item @("PMOHealth", "Health")
        Progress = $progress
        DueDate = FieldValue $item @("PMODueDate", "DueDate")
        FilesLink = FieldValue $item @("PMOFilesLink", "FilesLink")
        Tasks = 0
        Risks = 0
    }
}

$tasks = @()
foreach ($item in $taskItems) {
    $tasks += [ordered]@{
        id = "T-$($item.Id)"
        title = FieldValue $item @("Title", "TaskName")
        Project = FieldValue $item @("PMOTaskProject", "Project")
        AssignedTo = FieldValue $item @("PMOAssignedTo", "AssignedTo")
        Status = FieldValue $item @("PMOTaskStatus", "Status")
        Priority = FieldValue $item @("PMOPriority", "Priority")
        DueDate = FieldValue $item @("PMOTaskDueDate", "DueDate")
        Progress = NumberValue (FieldValue $item @("PMOTaskProgress", "Progress")) 0
    }
}

$risks = @()
foreach ($item in $riskItems) {
    $risks += [ordered]@{
        id = "R-$($item.Id)"
        title = FieldValue $item @("Title", "RiskTitle")
        Project = FieldValue $item @("PMORiskProject", "Project")
        Owner = FieldValue $item @("PMORiskOwner", "Owner")
        Severity = FieldValue $item @("PMOSeverity", "Severity")
        Status = FieldValue $item @("PMORiskStatus", "Status")
        MitigationPlan = FieldValue $item @("PMOMitigationPlan", "MitigationPlan")
    }
}

$reports = @()
foreach ($item in $reportItems) {
    $reports += [ordered]@{
        id = "WR-$($item.Id)"
        title = FieldValue $item @("Title")
        Project = FieldValue $item @("PMOReportProject", "Project")
        ReportDate = FieldValue $item @("PMOReportDate", "ReportDate")
        OverallHealth = FieldValue $item @("PMOOverallHealth", "OverallHealth")
        Achievements = FieldValue $item @("PMOAchievements", "Achievements")
        Challenges = FieldValue $item @("PMOChallenges", "Challenges")
        NextSteps = FieldValue $item @("PMONextSteps", "NextSteps")
    }
}

$resources = @()
foreach ($item in $resourceItems) {
    $util = NumberValue (FieldValue $item @("PMOUtilization", "Utilization")) 0
    $resources += [ordered]@{
        id = "RES-$($item.Id)"
        name = FieldValue $item @("PMOEmployee", "Employee", "Title")
        role = FieldValue $item @("PMORole", "Role")
        CapacityHours = NumberValue (FieldValue $item @("PMOCapacityHours", "CapacityHours")) 0
        PlannedHours = NumberValue (FieldValue $item @("PMOPlannedHours", "PlannedHours")) 0
        ActualHours = NumberValue (FieldValue $item @("PMOActualHours", "ActualHours")) 0
        Utilization = $util
        load = $util
        status = if ($util -gt 90) { "Overloaded" } elseif ($util -gt 75) { "Normal" } else { "Available" }
    }
}

$files = @()
foreach ($item in $fileItems) {
    $files += [ordered]@{
        id = "F-$($item.Id)"
        title = FieldValue $item @("FileLeafRef", "Title")
        link = FieldValue $item @("FileRef")
        count = 1
    }
}

# Build sheet-like tables for Databases page
$sheets = @(
    [ordered]@{ name = "PMO Projects"; headers = @("ID", "Project", "Manager", "Status", "Progress", "Due Date", "Files Link"); rows = @($projects | ForEach-Object { @($_.id, $_.ProjectName, $_.ProjectManager, $_.Status, "$($_.Progress)%", $_.DueDate, $_.FilesLink) }) },
    [ordered]@{ name = "PMO Tasks"; headers = @("ID", "Task", "Project", "Assigned To", "Status", "Priority", "Due Date", "Progress"); rows = @($tasks | ForEach-Object { @($_.id, $_.title, $_.Project, $_.AssignedTo, $_.Status, $_.Priority, $_.DueDate, "$($_.Progress)%") }) },
    [ordered]@{ name = "PMO Risks"; headers = @("ID", "Risk", "Project", "Owner", "Severity", "Status", "Mitigation Plan"); rows = @($risks | ForEach-Object { @($_.id, $_.title, $_.Project, $_.Owner, $_.Severity, $_.Status, $_.MitigationPlan) }) },
    [ordered]@{ name = "PMO Status Reports"; headers = @("ID", "Project", "Report Date", "Health", "Achievements", "Challenges", "Next Steps"); rows = @($reports | ForEach-Object { @($_.id, $_.Project, $_.ReportDate, $_.OverallHealth, $_.Achievements, $_.Challenges, $_.NextSteps) }) },
    [ordered]@{ name = "PMO Resources"; headers = @("ID", "Employee", "Role", "Capacity", "Planned", "Actual", "Utilization"); rows = @($resources | ForEach-Object { @($_.id, $_.name, $_.role, $_.CapacityHours, $_.PlannedHours, $_.ActualHours, "$($_.Utilization)%") }) }
)

$data = [ordered]@{
    source = [ordered]@{
        type = "sharepoint"
        siteUrl = $SiteUrl
        updatedAt = (Get-Date).ToString("s")
    }
    projects = $projects
    tasks = $tasks
    risks = $risks
    reports = $reports
    resources = $resources
    files = $files
    sheets = $sheets
}

$folder = Split-Path $OutputPath -Parent
if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Force -Path $folder | Out-Null }
$data | ConvertTo-Json -Depth 20 | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "SharePoint data exported successfully." -ForegroundColor Green
Write-Host "Output: $OutputPath" -ForegroundColor Green
Write-Host ""
Write-Host "Next commands:" -ForegroundColor Yellow
Write-Host "git add public/data/sharepoint-data.json"
Write-Host "git commit -m 'Sync PMO SharePoint data'"
Write-Host "git push"
