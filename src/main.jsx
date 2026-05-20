import React from 'react';
import { createRoot } from 'react-dom/client';
import { motion } from 'framer-motion';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  BarChart3,
  CalendarDays,
  ClipboardList,
  Download,
  ExternalLink,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Gauge,
  Home,
  KanbanSquare,
  Link as LinkIcon,
  Mail,
  PlusCircle,
  RefreshCcw,
  Save,
  Search,
  Settings,
  ShieldCheck,
  Table2,
  Upload,
  Users,
} from 'lucide-react';
import './styles.css';

const CONFIG = {
  companyName: 'Leader Investment Group',
  portalName: 'PMO Live Portal',
  microsoftLists: {
    projects: 'https://lists.microsoft.com/',
    tasks: 'https://lists.microsoft.com/',
    risks: 'https://lists.microsoft.com/',
    reports: 'https://lists.microsoft.com/',
    resources: 'https://lists.microsoft.com/',
  },
  forms: {
    addTask: 'https://forms.office.com/',
    updateTask: 'https://forms.office.com/',
    weeklyReport: 'https://forms.office.com/',
    registerRisk: 'https://forms.office.com/',
  },
};

const DEFAULT_PROJECTS = [
  { id: 'P001', name: 'EPM-IDT Phase 3', arabicName: 'مشروع التحول الرقمي في عمليات الحوكمة والرقابة على المشاريع الاستثمارية بامانة المنطقة الشرقية', region: 'EPM', businessLine: 'Le Source', owner: 'Mohamed Ahmed / Mohamed AL Khatim', status: 'Development', health: 'Amber', progress: 64, risk: 'High', openTasks: 28, overdue: 0, criticalRisks: 0, taskCount: 28, files: 'https://teams.microsoft.com/', schedule: 'https://teams.microsoft.com/', reports: 'https://teams.microsoft.com/', docs: 'https://teams.microsoft.com/', sheet: 'EPM-IDT Phase 3' },
  { id: 'P002', name: 'EPM- EPM III - HR', arabicName: 'مشروع زين الشرقية', region: 'EPM', businessLine: 'Le Sys', owner: 'Rami Mamoon / Ahmed Al Said / Abdelrahman Abdellatif', status: 'Development', health: 'Amber', progress: 59, risk: 'High', openTasks: 28, overdue: 0, criticalRisks: 0, taskCount: 28, files: 'https://teams.microsoft.com/', schedule: 'https://teams.microsoft.com/', reports: 'https://teams.microsoft.com/', docs: 'https://teams.microsoft.com/', sheet: 'EPM- EPM III - HR' },
];

const DEFAULT_TASKS = [
  { id: 'T001', project: 'EPM- EPM III - Finance', title: 'أوامر التغير', assignee: 'Ahmed Al Said', status: 'انشاء ملف متطلبات', phase: '2- Analysis', due: 'No due date', priority: 'High', progress: 0, note: 'فور الانتهاء من مراجعة العقود سوف يتم عقد اجتماع لمناقشة ما تم انجازة لاوامر التغير' },
  { id: 'T002', project: 'EPM- EPM III - Finance', title: 'التكامل مع فواتير بلدي', assignee: 'Ahmed Al Said', status: 'انشاء ملف متطلبات', phase: '2- Analysis', due: 'No due date', priority: 'High', progress: 0, note: 'طلب تغيري - الربط مع منصة بلدي' },
  { id: 'T003', project: 'EPM-IDT Phase 3', title: 'الرسائل التفاعلية (تشات بوت)', assignee: 'Mohamed Ahmed', status: 'قيد التطوير', phase: '5- Development', due: 'No due date', priority: 'Medium', progress: 55, note: 'تم طلب الرقم الخاص بالشات بوت ومن ثم سيتم البدء التهيئة' },
];

const DEFAULT_RISKS = [
  { id: 'R001', project: 'EPM-IDT Phase 3', title: 'الرسائل التفاعلية (تشات بوت)', severity: 'Medium', owner: 'Mohamed Ahmed', status: 'قيد التطوير', note: 'تم طلب الرقم الخاص بالشات بوت ومن ثم سيتم البدء التهيئة' },
  { id: 'R002', project: 'EPM- EPM III - Finance', title: 'أوامر التغير', severity: 'High', owner: 'Ahmed Al Said', status: 'انشاء ملف متطلبات', note: 'فور الانتهاء من مراجعة العقود سوف يتم عقد اجتماع لمناقشة ما تم انجازة لاوامر التغير' },
];

function norm(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).replace(/\s+/g, ' ').trim();
}

function normalizeHeader(header, index) {
  return norm(header) || `Column ${index + 1}`;
}

function isCompleted(text) {
  const v = norm(text).toLowerCase();
  return ['تم التطوير', 'تم التسليم', 'تم الانتهاء', 'تم الاطلاق', 'closed', 'support', 'الدعم', 'قبول'].some(x => v.includes(x));
}

function phaseProgress(phase, status) {
  const txt = `${phase || ''} ${status || ''}`.toLowerCase();
  const match = norm(phase).match(/^(\d{1,2})\s*[-–]/);
  if (match) return Math.max(0, Math.min(100, Math.round((Number(match[1]) / 11) * 100)));
  if (isCompleted(txt)) return 92;
  if (txt.includes('go-live') || txt.includes('go live') || txt.includes('اطلاق') || txt.includes('الاطلاق')) return 82;
  if (txt.includes('training') || txt.includes('تدريب')) return 73;
  if (txt.includes('deployment') || txt.includes('production') || txt.includes('خوادم') || txt.includes('انتاج')) return 64;
  if (txt.includes('development') || txt.includes('تطوير')) return 55;
  if (txt.includes('analysis') || txt.includes('تحليل')) return 25;
  if (txt.includes('planning') || txt.includes('تخطيط')) return 10;
  return 0;
}

function priorityFrom(progress, status, note) {
  const txt = `${status || ''} ${note || ''}`.toLowerCase();
  if (progress < 30 || ['متوقف', 'معلق', 'بانتظار', 'blocked', 'critical'].some(x => txt.includes(x))) return 'Critical';
  if (progress < 60 || ['انتظار', 'لم يتم', 'تأخير'].some(x => txt.includes(x))) return 'High';
  if (progress < 85) return 'Medium';
  return 'Low';
}

function statusFromProgress(progress) {
  if (progress >= 96) return 'Completed';
  if (progress >= 85) return 'Support';
  if (progress >= 70) return 'Deployment';
  if (progress >= 50) return 'Development';
  if (progress >= 25) return 'Analysis';
  return 'Planning';
}

function healthFrom(progress, openTasks, criticalRisks) {
  if (criticalRisks >= 3 || progress < 45) return 'Red';
  if (criticalRisks >= 1 || progress < 70 || openTasks > 20) return 'Amber';
  return 'Green';
}

function sheetToRows(ws) {
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  const headerIndex = matrix.findIndex(row => row.filter(cell => norm(cell)).length >= 3);
  const safeHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
  const headers = (matrix[safeHeaderIndex] || []).map(normalizeHeader);
  const lastHeader = headers.reduce((last, h, i) => (norm(h) ? i : last), 0);
  const finalHeaders = headers.slice(0, lastHeader + 1).map((h, i) => h || `Column ${i + 1}`);
  const rows = matrix.slice(safeHeaderIndex + 1).map(row => {
    const obj = {};
    finalHeaders.forEach((h, i) => { obj[h] = norm(row[i]); });
    return obj;
  }).filter(row => Object.values(row).some(v => norm(v)));
  return { headers: finalHeaders, rows };
}

function parseWorkbook(workbook, fileName = 'Uploaded workbook') {
  const sheets = workbook.SheetNames.map(name => ({ name, ...sheetToRows(workbook.Sheets[name]) }));
  const erp = sheets.find(s => s.name === 'ERP Project');
  const serviceSheets = sheets.filter(s => !['ERP Project', 'Radar', 'LOV'].includes(s.name));
  const tasks = [];

  serviceSheets.forEach(sheet => {
    sheet.rows.forEach(row => {
      const title = row['الخدمة'] || row['Implementation Activates'] || row['Project Name'] || '';
      const phase = row['المرحلة'] || row['Implementation Life Cycle'] || '';
      const status = row['الحالة'] || '';
      const action = row['نوع الاجراء'] || row['الإجراء التالي'] || row['الاجراء التالي'] || '';
      const note = row['ملاحظة'] || '';
      const due = row['تاريخ الاستحقاق'] || '';
      const assignee = row['المعنيين'] || row['Project Team'] || '';
      if (!title) return;
      if (![phase, status, action, note, due, assignee].some(Boolean)) return;
      const progress = phaseProgress(phase, status);
      const priority = priorityFrom(progress, status, note);
      tasks.push({
        id: `T${String(tasks.length + 1).padStart(4, '0')}`,
        project: sheet.name,
        title,
        assignee: assignee || 'Unassigned',
        status: status || action || 'Open',
        phase,
        due: due || 'No due date',
        priority,
        progress,
        note,
        action,
        sheet: sheet.name,
      });
    });
  });

  const grouped = tasks.reduce((acc, task) => {
    acc[task.project] = acc[task.project] || [];
    acc[task.project].push(task);
    return acc;
  }, {});

  const projects = (erp?.rows || serviceSheets.map((s, i) => ({ 'Project Name': s.name, '#': i + 1 }))).filter(row => row['Project Name'] || row['الخدمة']).map((row, i) => {
    const name = row['Project Name'] || row['الخدمة'] || `Project ${i + 1}`;
    const projectTasks = grouped[name] || [];
    const progress = projectTasks.length ? Math.round(projectTasks.reduce((sum, t) => sum + t.progress, 0) / projectTasks.length) : 0;
    const openTasks = projectTasks.filter(t => t.progress < 90).length;
    const criticalRisks = projectTasks.filter(t => t.priority === 'Critical').length;
    const health = healthFrom(progress, openTasks, criticalRisks);
    return {
      id: `P${String(i + 1).padStart(3, '0')}`,
      name,
      arabicName: row['Arabic Name'] || '',
      businessLine: row['Business Line'] || '',
      region: row['Region'] || '',
      owner: row['Project Manager'] || row['Project Team'] || 'Project Team',
      start: row['Start Date'] || '',
      end: row['End Date'] || '',
      applications: row['Application/s'] || '',
      status: statusFromProgress(progress),
      health,
      progress,
      risk: health === 'Red' ? 'Critical' : health === 'Amber' ? 'High' : 'Medium',
      openTasks,
      overdue: 0,
      criticalRisks,
      taskCount: projectTasks.length,
      files: 'https://teams.microsoft.com/',
      schedule: 'https://teams.microsoft.com/',
      reports: 'https://teams.microsoft.com/',
      docs: 'https://teams.microsoft.com/',
      sheet: name,
    };
  });

  const risks = tasks.filter(t => ['Critical', 'High'].includes(t.priority)).map((t, i) => ({
    id: `R${String(i + 1).padStart(4, '0')}`,
    project: t.project,
    title: t.title,
    severity: t.priority,
    owner: t.assignee,
    status: t.status,
    phase: t.phase,
    note: t.note,
  }));

  return {
    sourceInfo: {
      workbook: fileName,
      projects: projects.length,
      sourceTasks: tasks.length,
      displayRisks: risks.length,
      sheets: sheets.length,
      updatedAt: new Date().toISOString().slice(0, 10),
    },
    projects,
    tasks,
    risks,
    sheets,
  };
}

function pillClass(value) {
  const map = {
    Green: 'pill green', Amber: 'pill amber', Red: 'pill red', Completed: 'pill green', Deployment: 'pill green', Support: 'pill green', Development: 'pill blue', Analysis: 'pill amber', Planning: 'pill amber', UAT: 'pill blue', Critical: 'pill red', High: 'pill orange', Medium: 'pill blue', Low: 'pill neutral', 'تم التطوير': 'pill green', 'تم التسليم': 'pill green', 'قبول': 'pill green', 'تحت التطوير': 'pill amber', 'في مرحلة التطوير': 'pill blue', 'قيد التطوير': 'pill blue', 'جاري التحليل': 'pill amber', 'انشاء ملف متطلبات': 'pill amber', 'الدعم': 'pill green', 'الدغم': 'pill amber', Blocked: 'pill red', Overdue: 'pill red', New: 'pill neutral'
  };
  return map[value] || 'pill neutral';
}

function Pill({ children }) { return <span className={pillClass(children)}>{children || '—'}</span>; }
function ButtonLink({ href, children, icon: Icon = ExternalLink }) { return <a className="btn" href={href} target="_blank" rel="noreferrer"><Icon size={16} />{children}</a>; }
function Metric({ icon: Icon, label, value, sub }) { return <div className="metric"><div><p>{label}</p><strong>{value}</strong><span>{sub}</span></div><div className="metricIcon"><Icon size={23} /></div></div>; }
function shortText(text, length = 120) { const value = norm(text); return value.length > length ? `${value.slice(0, length)}...` : value; }

function App() {
  const [workbookData, setWorkbookData] = React.useState(() => ({
    sourceInfo: { workbook: 'Starter sample. Upload Excel to load full workbook.', projects: DEFAULT_PROJECTS.length, sourceTasks: DEFAULT_TASKS.length, displayRisks: DEFAULT_RISKS.length, sheets: 0, updatedAt: new Date().toISOString().slice(0, 10) },
    projects: DEFAULT_PROJECTS,
    tasks: DEFAULT_TASKS,
    risks: DEFAULT_RISKS,
    sheets: [],
  }));
  const [query, setQuery] = React.useState('');
  const [taskQuery, setTaskQuery] = React.useState('');
  const [activeSheetName, setActiveSheetName] = React.useState('');
  const [sheetQuery, setSheetQuery] = React.useState('');
  const [selectedProjectId, setSelectedProjectId] = React.useState('P001');
  const [liveSheetUrl, setLiveSheetUrl] = React.useState(() => localStorage.getItem('liveSheetUrl') || '');
  const [editSheetUrl, setEditSheetUrl] = React.useState(() => localStorage.getItem('editSheetUrl') || '');

  const projects = workbookData.projects.length ? workbookData.projects : DEFAULT_PROJECTS;
  const tasks = workbookData.tasks.length ? workbookData.tasks : DEFAULT_TASKS;
  const risks = workbookData.risks.length ? workbookData.risks : DEFAULT_RISKS;
  const sheets = workbookData.sheets || [];
  const selected = projects.find(p => p.id === selectedProjectId) || projects[0];
  const activeSheet = sheets.find(s => s.name === activeSheetName) || sheets[0];

  React.useEffect(() => {
    if (!projects.find(p => p.id === selectedProjectId) && projects[0]) setSelectedProjectId(projects[0].id);
  }, [projects, selectedProjectId]);

  React.useEffect(() => {
    if (!activeSheetName && sheets[0]) setActiveSheetName(sheets[0].name);
  }, [sheets, activeSheetName]);

  const summary = React.useMemo(() => ({
    total: projects.length,
    atRisk: projects.filter(p => p.health === 'Amber' || p.health === 'Red').length,
    openTasks: projects.reduce((a,p)=>a + Number(p.openTasks || 0),0),
    overdue: projects.reduce((a,p)=>a + Number(p.overdue || 0),0),
    criticalRisks: projects.reduce((a,p)=>a + Number(p.criticalRisks || 0),0),
    avgProgress: projects.length ? Math.round(projects.reduce((a,p)=>a + Number(p.progress || 0),0) / projects.length) : 0,
  }), [projects]);

  const filteredProjects = projects.filter(p => `${p.name} ${p.arabicName} ${p.owner} ${p.status} ${p.region}`.toLowerCase().includes(query.toLowerCase()));
  const filteredTasks = tasks.filter(t => `${t.project} ${t.title} ${t.assignee} ${t.status} ${t.phase} ${t.note}`.toLowerCase().includes(taskQuery.toLowerCase())).slice(0, 100);
  const selectedTasks = tasks.filter(t => t.project === selected?.name).slice(0, 8);
  const selectedRisks = risks.filter(r => r.project === selected?.name).slice(0, 6);
  const filteredSheetRows = activeSheet ? activeSheet.rows.filter(row => row.some(cell => norm(cell).toLowerCase().includes(sheetQuery.toLowerCase()))).slice(0, 250) : [];

  async function handleUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const parsed = parseWorkbook(workbook, file.name);
    setWorkbookData(parsed);
    setActiveSheetName(parsed.sheets[0]?.name || '');
    setSelectedProjectId(parsed.projects[0]?.id || 'P001');
  }

  function saveLiveLinks() {
    localStorage.setItem('liveSheetUrl', liveSheetUrl);
    localStorage.setItem('editSheetUrl', editSheetUrl);
  }

  function exportCurrentData() {
    const payload = JSON.stringify(workbookData, null, 2);
    const blob = new Blob([payload], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'pmo-workbook-data.json';
    a.click();
    URL.revokeObjectURL(url);
  }

  return <div className="app">
    <aside className="sidebar">
      <div className="brand"><div className="logo"><Gauge size={23}/></div><div><h1>{CONFIG.portalName}</h1><p>{CONFIG.companyName}</p></div></div>
      <nav>{[[Home,'Dashboard'],[FileSpreadsheet,'Live Workbook'],[Table2,'Sheet Viewer'],[FolderOpen,'Projects'],[KanbanSquare,'Tasks Board'],[AlertTriangle,'Risks & Issues'],[Users,'Resources'],[FileText,'Weekly Reports']].map(([Icon,label])=><a key={label} href={`#${label.replaceAll(' ','-')}`}><Icon size={18}/><span>{label}</span></a>)}</nav>
      <div className="sideCard"><h3>Excel Live Control</h3><p>Upload the workbook to instantly render all sheets as professional web tables. Add the Excel Online edit link to keep editing in Microsoft.</p><label className="btn uploadBtn"><Upload size={16}/>Upload Excel<input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleUpload}/></label><ButtonLink href={editSheetUrl || '#'} icon={ExternalLink}>Edit Live Sheet</ButtonLink></div>
    </aside>

    <main>
      <motion.section className="hero" initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} transition={{duration:.45}}>
        <div><p className="eyebrow">Workbook web viewer • {workbookData.sourceInfo.updatedAt}</p><h2>Solution Department PMO Dashboard</h2><p className="intro">The portal now displays your Excel workbook as a web application with projects, tasks, risks, sheet tabs, search, filtering, and direct edit/open links for the source Excel file.</p></div>
        <div className="actions"><label className="primary fileAction"><Upload size={16}/> Upload Workbook<input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleUpload}/></label><ButtonLink href={CONFIG.forms.weeklyReport} icon={Mail}>Submit Weekly Report</ButtonLink><button className="darkBtn lightBtn" onClick={exportCurrentData}><Download size={16}/> Export JSON</button></div>
      </motion.section>

      <section className="metrics"><Metric icon={FolderOpen} label="Total Projects" value={summary.total} sub="From ERP Project / sheets"/><Metric icon={ClipboardList} label="Tracked Activities" value={workbookData.sourceInfo.sourceTasks || tasks.length} sub={`${summary.openTasks} open activities`}/><Metric icon={AlertTriangle} label="Attention Items" value={workbookData.sourceInfo.displayRisks || risks.length} sub={`${summary.atRisk} projects amber/red`}/><Metric icon={BarChart3} label="Portfolio Progress" value={`${summary.avgProgress}%`} sub="Average completion"/></section>

      <section className="card livePanel" id="Live-Workbook"><div className="cardHead"><div><h3>Live Workbook Connection</h3><p>Use Excel Online / OneDrive / Teams link for direct editing. The web dashboard remains the viewer and PMO layer.</p></div><Pill>{workbookData.sourceInfo.workbook}</Pill></div><div className="liveGrid"><div><label>Excel Online view/embed link</label><input value={liveSheetUrl} onChange={e=>setLiveSheetUrl(e.target.value)} placeholder="Paste OneDrive / SharePoint Excel view link"/></div><div><label>Excel Online edit link</label><input value={editSheetUrl} onChange={e=>setEditSheetUrl(e.target.value)} placeholder="Paste Excel edit link"/></div><button className="btn" onClick={saveLiveLinks}><Save size={16}/> Save links</button><a className="btn" href={liveSheetUrl || '#'} target="_blank" rel="noreferrer"><ExternalLink size={16}/> Open Live Workbook</a></div>{liveSheetUrl ? <iframe className="sheetFrame" src={liveSheetUrl} title="Live Excel Workbook"/> : <div className="emptyFrame"><FileSpreadsheet size={32}/><b>No live Excel link yet</b><span>Paste an Excel Online / OneDrive / SharePoint link above, or upload the workbook to render it in the portal.</span></div>}</section>

      <section className="gridMain">
        <div className="card wide" id="Projects"><div className="cardHead"><div><h3>Projects Portfolio</h3><p>Professional portfolio table generated from the uploaded workbook.</p></div><div className="search"><Search size={16}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search projects, owner, region..."/></div></div><div className="tableWrap"><table><thead><tr><th>Project</th><th>Region</th><th>Owner</th><th>Status</th><th>Progress</th><th>Files</th></tr></thead><tbody>{filteredProjects.map(p=><tr key={p.id} onClick={()=>setSelectedProjectId(p.id)} className={selected?.id===p.id?'selected':''}><td><b>{p.name}</b><small>{p.arabicName || p.id}</small></td><td>{p.region || '-'}</td><td>{shortText(p.owner, 70)}</td><td><Pill>{p.status}</Pill></td><td><div className="progress"><span><i style={{width:`${p.progress}%`}}/></span>{p.progress}%</div></td><td><a href={p.files} target="_blank" rel="noreferrer"><LinkIcon size={14}/> Open</a></td></tr>)}</tbody></table></div></div>
        <div className="card"><div className="cardHead"><div><h3>Project Workspace</h3><p>{selected?.name}</p></div><Pill>{selected?.health}</Pill></div><div className="workspace"><div className="workTop"><span>Progress</span><b>{selected?.progress || 0}%</b></div><div className="bar"><i style={{width:`${selected?.progress || 0}%`}}/></div><div className="mini"><div><b>{selected?.openTasks || 0}</b><span>Open</span></div><div><b>{selected?.taskCount || 0}</b><span>Total</span></div><div><b>{selected?.criticalRisks || 0}</b><span>Attention</span></div></div></div><div className="projectMeta"><p><b>Arabic name:</b> {selected?.arabicName || '-'}</p><p><b>Business line:</b> {selected?.businessLine || '-'}</p><p><b>Sheet:</b> {selected?.sheet || '-'}</p></div><div className="linkGrid"><ButtonLink href={selected?.files || '#'} icon={FolderOpen}>Main Files</ButtonLink><ButtonLink href={selected?.schedule || '#'} icon={CalendarDays}>Project Schedule</ButtonLink><ButtonLink href={selected?.reports || '#'} icon={FileText}>Status Reports</ButtonLink><ButtonLink href={selected?.docs || '#'} icon={FileText}>Documentations</ButtonLink></div></div>
      </section>

      <section className="card" id="Sheet-Viewer"><div className="cardHead"><div><h3>Workbook Sheet Viewer</h3><p>Browse all uploaded Excel sheets as searchable web tables.</p></div><div className="search"><Search size={16}/><input value={sheetQuery} onChange={e=>setSheetQuery(e.target.value)} placeholder="Search inside sheet..."/></div></div><div className="tabs">{sheets.length ? sheets.map(s=><button key={s.name} className={activeSheet?.name===s.name?'active':''} onClick={()=>setActiveSheetName(s.name)}>{s.name}<span>{s.rows.length}</span></button>) : <button className="active">Upload Excel to show sheets</button>}</div>{activeSheet ? <div className="tableWrap sheetTable"><table><thead><tr>{activeSheet.headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{filteredSheetRows.map((row, i)=><tr key={i}>{activeSheet.headers.map((h, j)=><td key={`${i}-${j}`}>{row[j]}</td>)}</tr>)}</tbody></table></div> : <div className="emptyFrame"><Upload size={32}/><b>Upload your workbook</b><span>The sheet tabs and tables will appear here instantly.</span></div>}</section>

      <section className="twoCols"><div className="card" id="Tasks-Board"><div className="cardHead"><div><h3>Task Control</h3><p>Search and review activities extracted from all project sheets.</p></div><div className="search smallSearch"><Search size={16}/><input value={taskQuery} onChange={e=>setTaskQuery(e.target.value)} placeholder="Search tasks..."/></div></div><div className="list">{filteredTasks.map(t=><div className="item" key={t.id}><div><b>{t.title}</b><p>{t.project} • {shortText(t.assignee, 80)} • {t.due}</p><p>{t.phase}</p></div><div><Pill>{t.status}</Pill><Pill>{t.priority}</Pill></div></div>)}</div></div><div className="card" id="Risks-&-Issues"><div className="cardHead"><div><h3>Risks & Issues</h3><p>Generated attention list from low-progress or pending activities.</p></div><ButtonLink href={CONFIG.forms.registerRisk} icon={PlusCircle}>Register</ButtonLink></div><div className="list">{risks.slice(0, 40).map(r=><div className="item" key={r.id}><div><b>{r.title}</b><p>{r.project} • Owner: {shortText(r.owner, 80)}</p><p>{shortText(r.note, 110)}</p></div><Pill>{r.severity}</Pill></div>)}</div></div></section>

      <section className="twoCols"><div className="card"><div className="cardHead"><div><h3>Selected Project Tasks</h3><p>Activities for {selected?.name}</p></div><Pill>{selectedTasks.length}</Pill></div><div className="list">{selectedTasks.length ? selectedTasks.map(t=><div className="item" key={t.id}><div><b>{t.title}</b><p>{shortText(t.assignee, 80)} • {t.phase}</p></div><div><Pill>{t.status}</Pill></div></div>) : <div className="item"><div><b>No tasks displayed</b><p>Open the sheet viewer for full workbook details.</p></div></div>}</div></div><div className="card"><div className="cardHead"><div><h3>Selected Project Risks</h3><p>Attention items for {selected?.name}</p></div><Pill>{selectedRisks.length}</Pill></div><div className="list">{selectedRisks.length ? selectedRisks.map(r=><div className="item" key={r.id}><div><b>{r.title}</b><p>{shortText(r.owner, 80)} • {r.status}</p></div><Pill>{r.severity}</Pill></div>) : <div className="item"><div><b>No generated risks</b><p>This project has no generated high-priority risks.</p></div></div>}</div></div></section>

      <section className="report" id="Weekly-Reports"><div><h3>Weekly Reporting Center</h3><p>The dashboard can render the workbook instantly. For true live editing, paste your Excel Online edit link above. Microsoft permissions remain controlled by Teams/OneDrive/SharePoint.</p></div><div className="actions"><a className="primary" href={CONFIG.forms.weeklyReport} target="_blank" rel="noreferrer"><Mail size={16}/> Submit Report</a><a className="darkBtn" href={CONFIG.microsoftLists.reports} target="_blank" rel="noreferrer"><FileText size={16}/> Open Reports List</a></div></section>
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App />);
