<#
.SYNOPSIS
Creates a professional PMO Project Control Hub page in SharePoint using PnP.PowerShell.

.DESCRIPTION
This script creates:
- PMO SharePoint lists if they do not exist
- PMO document library if it does not exist
- A modern SharePoint page with bilingual PMO layout
- Executive dashboard cards
- Quick action buttons
- Portfolio / tasks / risks / reports / documents sections
- Links to Vercel Project Control Hub and PMO Starter Kit

NOTES
- Run in PowerShell 7.4+
- Requires PnP.PowerShell
- If your tenant requires an Entra App Registration, pass -ClientId from IT Admin
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl = "https://leaderig.sharepoint.com/sites/Solution",

    [Parameter(Mandatory = $false)]
    [string]$PageName = "PMO-Project-Control-Hub.aspx",

    [Parameter(Mandatory = $false)]
    [string]$PageTitle = "Project Control Hub | PMO Portal",

    [Parameter(Mandatory = $false)]
    [string]$PortalUrl = "https://pmo-live-portal.vercel.app/",

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "",

    [Parameter(Mandatory = $false)]
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[PMO] $Message" -ForegroundColor Cyan
}

function Ensure-PnPModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Step "Installing PnP.PowerShell for current user..."
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
}

function Connect-PMOSite {
    Write-Step "Connecting to $SiteUrl"
    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    }
    else {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    }
}

function Ensure-List {
    param(
        [string]$Title,
        [hashtable[]]$Fields
    )

    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Step "Creating list: $Title"
        New-PnPList -Title $Title -Template GenericList -OnQuickLaunch | Out-Null
    }
    else {
        Write-Step "List exists: $Title"
    }

    foreach ($field in $Fields) {
        $internalName = $field.InternalName
        $existingField = Get-PnPField -List $Title -Identity $internalName -ErrorAction SilentlyContinue
        if (-not $existingField) {
            Add-PnPField -List $Title -DisplayName $field.DisplayName -InternalName $internalName -Type $field.Type -AddToDefaultView | Out-Null
        }
    }
}

function Ensure-DocumentLibrary {
    param([string]$Title)
    $lib = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $lib) {
        Write-Step "Creating document library: $Title"
        New-PnPList -Title $Title -Template DocumentLibrary -OnQuickLaunch | Out-Null
    }
    else {
        Write-Step "Document library exists: $Title"
    }
}

function New-CardHtml {
    param(
        [string]$Icon,
        [string]$Title,
        [string]$Value,
        [string]$Sub,
        [string]$Color = "#2563eb"
    )

    return @"
<div style='background:#ffffff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;box-shadow:0 8px 24px rgba(15,23,42,.06);min-height:118px;'>
  <div style='font-size:24px;margin-bottom:8px;color:$Color;'>$Icon</div>
  <div style='font-size:13px;color:#697386;font-weight:600;'>$Title</div>
  <div style='font-size:32px;font-weight:800;color:#17212f;line-height:1.1;margin-top:8px;'>$Value</div>
  <div style='font-size:12px;color:#697386;margin-top:6px;'>$Sub</div>
</div>
"@
}

function Add-PMOSection {
    param(
        [string]$Page,
        [int]$Order,
        [string]$Html
    )

    Add-PnPPageSection -Page $Page -SectionTemplate OneColumn -Order $Order | Out-Null
    Add-PnPPageTextPart -Page $Page -Section $Order -Column 1 -Text $Html | Out-Null
}

function Build-HeroHtml {
    return @"
<div style='direction:rtl;background:linear-gradient(135deg,#111827 0%,#1f3b8a 58%,#2563eb 100%);color:white;border-radius:28px;padding:34px;box-shadow:0 22px 60px rgba(15,23,42,.22);font-family:Segoe UI,Tahoma,Arial,sans-serif;'>
  <div style='display:flex;gap:18px;align-items:center;justify-content:space-between;flex-wrap:wrap;'>
    <div style='max-width:860px;'>
      <div style='display:inline-block;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.18);padding:7px 12px;border-radius:999px;font-size:13px;font-weight:700;margin-bottom:14px;'>PMO Project Control Hub</div>
      <h1 style='font-size:42px;line-height:1.1;margin:0 0 12px;font-weight:800;'>بوابة التحكم وإدارة المشاريع</h1>
      <p style='font-size:16px;line-height:1.8;margin:0;color:#dbeafe;'>Professional bilingual workspace for portfolio, projects, tasks, risks, resources, files, weekly reports and executive follow-up.</p>
    </div>
    <div style='display:flex;gap:10px;flex-wrap:wrap;'>
      <a href='$PortalUrl' target='_blank' style='background:white;color:#111827;text-decoration:none;padding:13px 18px;border-radius:14px;font-weight:800;'>Open PMO Portal</a>
      <a href='#PMOActions' style='background:rgba(255,255,255,.12);color:white;text-decoration:none;padding:13px 18px;border-radius:14px;font-weight:800;border:1px solid rgba(255,255,255,.22);'>Quick Actions</a>
    </div>
  </div>
</div>
"@
}

function Build-KpiHtml {
    $cards = @(
        New-CardHtml -Icon "📁" -Title "Total Projects" -Value "24" -Sub "Portfolio items" -Color "#2563eb"
        New-CardHtml -Icon "✅" -Title "Open Tasks" -Value "120" -Sub "Team activities" -Color "#16a34a"
        New-CardHtml -Icon "⚠️" -Title "High Risks" -Value "8" -Sub "Needs attention" -Color "#dc2626"
        New-CardHtml -Icon "📊" -Title "Portfolio Progress" -Value "68%" -Sub "Average completion" -Color "#7c3aed"
        New-CardHtml -Icon "👥" -Title "Utilization" -Value "74%" -Sub "Team capacity" -Color "#d97706"
    ) -join ""

    return @"
<div style='direction:rtl;font-family:Segoe UI,Tahoma,Arial,sans-serif;'>
  <h2 style='margin:0 0 14px;font-size:26px;color:#17212f;'>Executive Dashboard | لوحة المؤشرات</h2>
  <div style='display:grid;grid-template-columns:repeat(5,minmax(160px,1fr));gap:14px;'>$cards</div>
</div>
"@
}

function Build-QuickActionsHtml {
    $actions = @(
        @{T="Portfolio Dashboard"; A="لوحة المحفظة"; I="📊"; U=$PortalUrl},
        @{T="Projects List"; A="قائمة المشاريع"; I="📁"; U="../Lists/PMO%20Projects/AllItems.aspx"},
        @{T="Tasks Board"; A="لوحة المهام"; I="✅"; U="../Lists/PMO%20Tasks/AllItems.aspx"},
        @{T="Risks & Issues"; A="المخاطر والمشكلات"; I="⚠️"; U="../Lists/PMO%20Risks/AllItems.aspx"},
        @{T="Weekly Reports"; A="التقارير الأسبوعية"; I="📝"; U="../Lists/PMO%20Status%20Reports/AllItems.aspx"},
        @{T="Team Resources"; A="الموارد والفريق"; I="👥"; U="../Lists/PMO%20Resources/AllItems.aspx"},
        @{T="Project Files"; A="ملفات المشاريع"; I="📚"; U="../PMO%20Project%20Files/Forms/AllItems.aspx"},
        @{T="Automation Rules"; A="الأتمتة والتنبيهات"; I="🔔"; U="#PMOAutomation"}
    )

    $html = ""
    foreach ($a in $actions) {
        $html += @"
<a href='$($a.U)' target='_blank' style='text-decoration:none;color:#17212f;background:#fff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;display:block;box-shadow:0 8px 24px rgba(15,23,42,.05);'>
  <div style='font-size:26px;margin-bottom:10px;'>$($a.I)</div>
  <div style='font-size:15px;font-weight:800;'>$($a.A)</div>
  <div style='font-size:12px;color:#697386;margin-top:4px;'>$($a.T)</div>
</a>
"@
    }

    return @"
<div id='PMOActions' style='direction:rtl;font-family:Segoe UI,Tahoma,Arial,sans-serif;'>
  <h2 style='margin:0 0 14px;font-size:26px;color:#17212f;'>Quick Actions | إجراءات سريعة</h2>
  <div style='display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:14px;'>$html</div>
</div>
"@
}

function Build-PortalEmbedFallbackHtml {
    return @"
<div style='direction:rtl;font-family:Segoe UI,Tahoma,Arial,sans-serif;background:#ffffff;border:1px solid #e6eaf0;border-radius:22px;padding:24px;box-shadow:0 8px 24px rgba(15,23,42,.05);'>
  <h2 style='margin:0 0 8px;font-size:26px;'>Project Control Hub</h2>
  <p style='margin:0 0 16px;color:#697386;line-height:1.8;'>افتح البوابة التفاعلية الكاملة لعرض Dashboard، Kanban Board، Databases، Tasks، Risks، Team، Files، Reports باللغتين العربية والإنجليزية.</p>
  <a href='$PortalUrl' target='_blank' style='display:inline-block;background:#111827;color:white;text-decoration:none;padding:13px 18px;border-radius:14px;font-weight:800;'>Open Interactive Portal</a>
  <div style='margin-top:16px;background:#f8fafc;border:1px dashed #cbd5e1;border-radius:16px;padding:14px;color:#475569;font-size:13px;line-height:1.7;'>If Embed is allowed by IT, add an Embed web part manually with:<br/><code>&lt;iframe src="$PortalUrl" width="100%" height="950" style="border:0;border-radius:18px;"&gt;&lt;/iframe&gt;</code></div>
</div>
"@
}

function Build-ListsHtml {
    return @"
<div style='direction:rtl;font-family:Segoe UI,Tahoma,Arial,sans-serif;'>
  <h2 style='margin:0 0 14px;font-size:26px;color:#17212f;'>PMO Databases | قواعد البيانات</h2>
  <div style='display:grid;grid-template-columns:repeat(3,minmax(220px,1fr));gap:14px;'>
    <div style='background:#fff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;'><b>PMO Projects</b><p style='color:#697386;line-height:1.7;'>Project portfolio, ownership, status, progress and links.</p><a href='../Lists/PMO%20Projects/AllItems.aspx'>Open List</a></div>
    <div style='background:#fff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;'><b>PMO Tasks</b><p style='color:#697386;line-height:1.7;'>Team tasks, due dates, priorities and completion.</p><a href='../Lists/PMO%20Tasks/AllItems.aspx'>Open List</a></div>
    <div style='background:#fff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;'><b>PMO Risks</b><p style='color:#697386;line-height:1.7;'>Risks, issues, severity, owners and mitigation plans.</p><a href='../Lists/PMO%20Risks/AllItems.aspx'>Open List</a></div>
    <div style='background:#fff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;'><b>Status Reports</b><p style='color:#697386;line-height:1.7;'>Weekly and monthly reports for each project.</p><a href='../Lists/PMO%20Status%20Reports/AllItems.aspx'>Open List</a></div>
    <div style='background:#fff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;'><b>Resources</b><p style='color:#697386;line-height:1.7;'>Team capacity, planned hours and utilization.</p><a href='../Lists/PMO%20Resources/AllItems.aspx'>Open List</a></div>
    <div style='background:#fff;border:1px solid #e6eaf0;border-radius:18px;padding:18px;'><b>Project Files</b><p style='color:#697386;line-height:1.7;'>Schedules, reports, contracts, documents and deliverables.</p><a href='../PMO%20Project%20Files/Forms/AllItems.aspx'>Open Library</a></div>
  </div>
</div>
"@
}

function Build-AutomationHtml {
    return @"
<div id='PMOAutomation' style='direction:rtl;font-family:Segoe UI,Tahoma,Arial,sans-serif;'>
  <h2 style='margin:0 0 14px;font-size:26px;color:#17212f;'>Automation & Governance | الأتمتة والحوكمة</h2>
  <div style='display:grid;grid-template-columns:repeat(2,minmax(260px,1fr));gap:14px;'>
    <div style='background:#fff7ed;border:1px solid #fed7aa;border-radius:18px;padding:18px;'><b>Delayed Task Alert</b><p>إذا تأخرت المهمة يتم تنبيه مدير المشروع.</p></div>
    <div style='background:#fef2f2;border:1px solid #fecaca;border-radius:18px;padding:18px;'><b>High Risk Escalation</b><p>إذا زادت المخاطر يتم إنشاء تصعيد ومتابعة أسبوعية.</p></div>
    <div style='background:#eff6ff;border:1px solid #bfdbfe;border-radius:18px;padding:18px;'><b>Weekly Report Reminder</b><p>تذكير أسبوعي لتحديث تقارير حالة المشاريع.</p></div>
    <div style='background:#f0fdf4;border:1px solid #bbf7d0;border-radius:18px;padding:18px;'><b>Resource Utilization Warning</b><p>تنبيه عند تجاوز استخدام الموارد 90%.</p></div>
  </div>
</div>
"@
}

Ensure-PnPModule
Connect-PMOSite

Write-Step "Creating PMO lists and document library..."
Ensure-List -Title "PMO Projects" -Fields @(
    @{DisplayName="Project Name"; InternalName="PMOProjectName"; Type="Text"},
    @{DisplayName="Project Manager"; InternalName="PMOProjectManager"; Type="Text"},
    @{DisplayName="Status"; InternalName="PMOStatus"; Type="Choice"},
    @{DisplayName="Progress"; InternalName="PMOProgress"; Type="Number"},
    @{DisplayName="Health"; InternalName="PMOHealth"; Type="Choice"},
    @{DisplayName="Due Date"; InternalName="PMODueDate"; Type="DateTime"},
    @{DisplayName="Files Link"; InternalName="PMOFilesLink"; Type="URL"}
)

Ensure-List -Title "PMO Tasks" -Fields @(
    @{DisplayName="Project"; InternalName="PMOTaskProject"; Type="Text"},
    @{DisplayName="Assigned To"; InternalName="PMOAssignedTo"; Type="Text"},
    @{DisplayName="Status"; InternalName="PMOTaskStatus"; Type="Choice"},
    @{DisplayName="Priority"; InternalName="PMOPriority"; Type="Choice"},
    @{DisplayName="Due Date"; InternalName="PMOTaskDueDate"; Type="DateTime"},
    @{DisplayName="Progress"; InternalName="PMOTaskProgress"; Type="Number"}
)

Ensure-List -Title "PMO Risks" -Fields @(
    @{DisplayName="Project"; InternalName="PMORiskProject"; Type="Text"},
    @{DisplayName="Owner"; InternalName="PMORiskOwner"; Type="Text"},
    @{DisplayName="Severity"; InternalName="PMOSeverity"; Type="Choice"},
    @{DisplayName="Status"; InternalName="PMORiskStatus"; Type="Choice"},
    @{DisplayName="Mitigation Plan"; InternalName="PMOMitigationPlan"; Type="Note"}
)

Ensure-List -Title "PMO Status Reports" -Fields @(
    @{DisplayName="Project"; InternalName="PMOReportProject"; Type="Text"},
    @{DisplayName="Report Date"; InternalName="PMOReportDate"; Type="DateTime"},
    @{DisplayName="Overall Health"; InternalName="PMOOverallHealth"; Type="Choice"},
    @{DisplayName="Achievements"; InternalName="PMOAchievements"; Type="Note"},
    @{DisplayName="Challenges"; InternalName="PMOChallenges"; Type="Note"},
    @{DisplayName="Next Steps"; InternalName="PMONextSteps"; Type="Note"}
)

Ensure-List -Title "PMO Resources" -Fields @(
    @{DisplayName="Employee"; InternalName="PMOEmployee"; Type="Text"},
    @{DisplayName="Role"; InternalName="PMORole"; Type="Text"},
    @{DisplayName="Capacity Hours"; InternalName="PMOCapacityHours"; Type="Number"},
    @{DisplayName="Planned Hours"; InternalName="PMOPlannedHours"; Type="Number"},
    @{DisplayName="Actual Hours"; InternalName="PMOActualHours"; Type="Number"},
    @{DisplayName="Utilization"; InternalName="PMOUtilization"; Type="Number"}
)

Ensure-DocumentLibrary -Title "PMO Project Files"

$pageExists = Get-PnPPage -Identity $PageName -ErrorAction SilentlyContinue
if ($pageExists -and $Overwrite) {
    Write-Step "Removing existing page: $PageName"
    Remove-PnPPage -Identity $PageName -Force
    $pageExists = $null
}

if (-not $pageExists) {
    Write-Step "Creating page: $PageName"
    Add-PnPPage -Name $PageName -LayoutType Article -Title $PageTitle | Out-Null
}
else {
    Write-Step "Page already exists: $PageName. Use -Overwrite to recreate it."
    Write-Host "Page URL: $SiteUrl/SitePages/$PageName" -ForegroundColor Yellow
    return
}

Add-PMOSection -Page $PageName -Order 1 -Html (Build-HeroHtml)
Add-PMOSection -Page $PageName -Order 2 -Html (Build-KpiHtml)
Add-PMOSection -Page $PageName -Order 3 -Html (Build-QuickActionsHtml)
Add-PMOSection -Page $PageName -Order 4 -Html (Build-PortalEmbedFallbackHtml)
Add-PMOSection -Page $PageName -Order 5 -Html (Build-ListsHtml)
Add-PMOSection -Page $PageName -Order 6 -Html (Build-AutomationHtml)

try {
    Set-PnPPage -Identity $PageName -Title $PageTitle -Publish | Out-Null
}
catch {
    Write-Step "Page created, but automatic publish failed. Open the page and click Publish manually."
}

Write-Host ""
Write-Host "PMO SharePoint page created successfully." -ForegroundColor Green
Write-Host "Page URL: $SiteUrl/SitePages/$PageName" -ForegroundColor Green
Write-Host ""
Write-Host "To overwrite later, run the script with -Overwrite" -ForegroundColor Yellow
