# PMO Portal Auto Update to SharePoint

This guide enables automatic updates from the PMO web app forms to SharePoint Lists.

## Goal

When a user submits a form in the PMO Portal, the record should be created automatically in the matching SharePoint list:

- projects -> PMO Projects
- tasks -> PMO Tasks
- risks -> PMO Risks
- reports -> PMO Status Reports
- resources -> PMO Resources
- schedules -> PMO Schedules
- files -> PMO Files Register

## Recommended option

Use Power Automate as a secure bridge between the Vercel web app and SharePoint.

PMO Portal Form -> Power Automate HTTP Trigger -> Switch by formType -> Create item in SharePoint List

## Step 1: Create the SharePoint lists

Use the existing script:

```powershell
.\sharepoint\Create-PMO-SharePoint-Portal.ps1 `
  -SiteUrl "https://leaderig.sharepoint.com/sites/Solution" `
  -ClientId "c431f27e-b41c-4488-90b0-5492f8622279" `
  -Overwrite
```

Recommended additional lists:

- PMO Schedules
- PMO Files Register

## Step 2: Create a Power Automate cloud flow

1. Go to https://make.powerautomate.com
2. Create > Automated cloud flow or Instant cloud flow
3. Trigger: When an HTTP request is received
4. Paste this JSON schema:

```json
{
  "type": "object",
  "properties": {
    "formType": { "type": "string" },
    "submittedAt": { "type": "string" },
    "submittedBy": { "type": "string" },
    "record": { "type": "object" }
  },
  "required": ["formType", "record"]
}
```

## Step 3: Add a Switch action

Switch on:

```text
formType
```

Cases:

```text
projects
 tasks
 risks
 reports
 resources
 schedules
 files
```

## Step 4: Create SharePoint item for each case

For every case, use:

```text
SharePoint > Create item
```

Site Address:

```text
https://leaderig.sharepoint.com/sites/Solution
```

List mapping examples:

### projects -> PMO Projects

- Title = record['en'] or record['ar']
- PMOProjectName = record['en'] or record['ar']
- PMOProjectManager = record['owner']
- PMOStatus = record['status']
- PMOProgress = record['progress']
- PMODueDate = record['due']
- PMOFilesLink = record['link']

### tasks -> PMO Tasks

- Title = record['title']
- PMOTaskProject = record['project']
- PMOAssignedTo = record['owner']
- PMOTaskStatus = record['status']
- PMOPriority = record['priority']
- PMOTaskDueDate = record['due']
- PMOTaskProgress = record['progress']

### risks -> PMO Risks

- Title = record['title']
- PMORiskProject = record['project']
- PMORiskOwner = record['owner']
- PMOSeverity = record['severity']
- PMORiskStatus = record['status']
- PMOMitigationPlan = record['mitigation']

### reports -> PMO Status Reports

- Title = record['project'] + ' - ' + record['date']
- PMOReportProject = record['project']
- PMOReportDate = record['date']
- PMOOverallHealth = record['health']
- PMOAchievements = record['achievements']
- PMOChallenges = record['challenges']
- PMONextSteps = record['nextSteps']

### resources -> PMO Resources

- Title = record['name']
- PMOEmployee = record['name']
- PMORole = record['role']
- PMOCapacityHours = record['capacity']
- PMOPlannedHours = record['planned']
- PMOActualHours = record['actual']
- PMOUtilization = record['util']

## Step 5: Copy the HTTP POST URL

After saving the flow, Power Automate will generate an HTTP POST URL.

Copy it. It will look like:

```text
https://prod-xx.region.logic.azure.com/workflows/...
```

## Step 6: Add the URL to Vercel

In Vercel:

1. Project > Settings > Environment Variables
2. Add:

```text
VITE_POWER_AUTOMATE_URL = <Power Automate HTTP POST URL>
```

3. Redeploy the project.

## Step 7: Test

Submit a record from the PMO Portal Forms page.

Expected result:

- The record is saved locally in the browser.
- The same record is posted to Power Automate.
- Power Automate creates an item in the matching SharePoint list.

## Important notes

- The HTTP trigger may require Power Automate premium depending on the tenant license.
- If HTTP trigger is not allowed, use Microsoft Forms + SharePoint List as a non-premium fallback.
- For enterprise-grade security, use Azure Function or Microsoft Graph with Entra ID authentication.
