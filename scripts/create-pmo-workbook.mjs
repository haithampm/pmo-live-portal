import XLSX from 'xlsx';
import fs from 'node:fs';
import path from 'node:path';

const outDir = path.resolve('output');
fs.mkdirSync(outDir, { recursive: true });

const now = new Date().toISOString().slice(0, 10);

const sheets = {
  Projects: [
    {
      ProjectID: 'P001',
      ProjectName: 'EPM-Smart Lightning',
      ArabicName: 'مشروع تجريبي',
      ProjectManager: 'Haitham Elmohamady',
      Region: 'EPM',
      BusinessLine: 'Le Source',
      Status: 'In Progress',
      Health: 'Green',
      Progress: 65,
      StartDate: now,
      EndDate: '',
      FilesLink: 'PASTE_ONEDRIVE_OR_TEAMS_FOLDER_LINK',
      ScheduleLink: 'PASTE_SCHEDULE_LINK',
      ReportsLink: 'PASTE_REPORTS_LINK',
      DocsLink: 'PASTE_DOCS_LINK',
      Notes: 'Replace sample data with your real project data',
    },
  ],
  Tasks: [
    {
      TaskID: 'T001',
      ProjectID: 'P001',
      ProjectName: 'EPM-Smart Lightning',
      TaskName: 'Prepare project plan',
      AssignedTo: 'Haitham Elmohamady',
      Status: 'In Progress',
      Priority: 'High',
      Phase: 'Planning',
      DueDate: '',
      Progress: 50,
      Notes: 'Initial sample task',
      FileLink: 'PASTE_FILE_LINK',
    },
  ],
  Risks: [
    {
      RiskID: 'R001',
      ProjectID: 'P001',
      ProjectName: 'EPM-Smart Lightning',
      RiskTitle: 'Pending approval may delay schedule',
      Owner: 'Haitham Elmohamady',
      Severity: 'Medium',
      Status: 'Open',
      MitigationPlan: 'Follow up weekly',
      DueDate: '',
      RelatedFileLink: 'PASTE_FILE_LINK',
    },
  ],
  StatusReports: [
    {
      ReportID: 'WR001',
      ProjectID: 'P001',
      ProjectName: 'EPM-Smart Lightning',
      ReportDate: now,
      OverallHealth: 'Green',
      Achievements: 'Initial setup completed',
      Challenges: 'Pending links and real data',
      NextSteps: 'Upload real project data',
      SupportNeeded: '',
      ReportFileLink: 'PASTE_REPORT_LINK',
    },
  ],
  Resources: [
    {
      ResourceID: 'RES001',
      Employee: 'Haitham Elmohamady',
      Role: 'Project Manager',
      ProjectID: 'P001',
      CapacityHours: 40,
      PlannedHours: 30,
      ActualHours: 20,
      Utilization: 50,
      Status: 'Normal',
    },
  ],
  Documents: [
    {
      DocumentID: 'D001',
      ProjectID: 'P001',
      ProjectName: 'EPM-Smart Lightning',
      DocumentType: 'Schedule',
      Title: 'Project Schedule',
      FileLink: 'PASTE_FILE_LINK',
      Owner: 'Haitham Elmohamady',
      UpdatedDate: now,
    },
  ],
  Choices: [
    { Field: 'Project Status', Values: 'Not Started, In Progress, On Hold, Delayed, Completed, Cancelled' },
    { Field: 'Health', Values: 'Green, Amber, Red' },
    { Field: 'Task Status', Values: 'New, In Progress, Waiting, Blocked, Completed, Cancelled' },
    { Field: 'Priority', Values: 'Low, Medium, High, Critical' },
    { Field: 'Severity', Values: 'Low, Medium, High, Critical' },
    { Field: 'Resource Status', Values: 'Available, Normal, Overloaded, Unavailable' },
  ],
  ReadMe: [
    { Step: 1, Instruction: 'Open this workbook in Excel Online using your Microsoft personal account haitham.pm@hotmail.com.' },
    { Step: 2, Instruction: 'Upload it to OneDrive under a folder named PMO Portal.' },
    { Step: 3, Instruction: 'Replace sample rows with your real project data.' },
    { Step: 4, Instruction: 'Use the Vercel portal Upload Workbook button to render the workbook as a dashboard.' },
    { Step: 5, Instruction: 'Paste the Excel Online edit/view links in the Live Workbook Connection section of the portal.' },
  ],
};

const wb = XLSX.utils.book_new();
Object.entries(sheets).forEach(([name, rows]) => {
  const ws = XLSX.utils.json_to_sheet(rows);
  const range = XLSX.utils.decode_range(ws['!ref']);
  ws['!cols'] = Array.from({ length: range.e.c + 1 }, () => ({ wch: 24 }));
  XLSX.utils.book_append_sheet(wb, ws, name);
});

const filePath = path.join(outDir, 'PMO_Master_Data.xlsx');
XLSX.writeFile(wb, filePath);
console.log(`Created ${filePath}`);
