 باimport React from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  AlertTriangle, BarChart3, Bell, BriefcaseBusiness, CalendarDays, ClipboardCheck,
  Database, Download, Edit3, Eye, FileText, Filter, Folder, Gauge, Home,
  KanbanSquare, Languages, Link as LinkIcon, Mail, MoreHorizontal, Plus,
  RefreshCcw, Search, Settings, Share2, ShieldCheck, Sparkles, Trash2, Upload,
  Users, Workflow, Save, FormInput
} from 'lucide-react';
import './styles.css';
import './sharepoint-sync.css';
import './forms.css';

const LS_KEY = 'pmo-control-hub-records-v4';
const TYPES = ['projects', 'tasks', 'risks', 'reports', 'resources', 'schedules', 'files'];
const EMPTY = Object.fromEntries(TYPES.map((type) => [type, []]));

const COPY = {
  en: {
    app: 'Project Control Hub',
    sub: 'Professional PMO workspace for portfolio, delivery, schedules, resources and reporting',
    search: 'Search projects, tasks, risks...', upload: 'Upload Excel', save: 'Save', export: 'Export', share: 'Share', refresh: 'Refresh', new: 'New', filters: 'Filters',
    backend: 'Backend', connected: 'Connected', local: 'Local fallback',
    nav: { dashboard: 'Dashboard', portfolio: 'Portfolio', board: 'Board', tasks: 'Tasks', risks: 'Risks', schedule: 'Schedule', teams: 'Resources', reports: 'Reports', files: 'Files', databases: 'Databases', forms: 'Forms' },
    status: { all: 'All', on: 'On Track', delayed: 'Delayed', blocked: 'At Risk', done: 'Completed', open: 'Open' },
    labels: { project: 'Project', owner: 'Owner', status: 'Status', priority: 'Priority', progress: 'Progress', due: 'Due', actions: 'Actions', health: 'Project Health Score', budget: 'Budget Utilization', active: 'Active Projects', total: 'Total Projects', risks: 'Open Risks', tasks: 'Delayed Tasks', team: 'Team Utilization', milestones: 'Milestones Completed' }
  },
  ar: {
    app: 'مركز التحكم بالمشاريع',
    sub: 'مساحة PMO احترافية لإدارة المحفظة والتنفيذ والجداول والموارد والتقارير',
    search: 'بحث في المشاريع والمهام والمخاطر...', upload: 'رفع Excel', save: 'حفظ', export: 'تصدير', share: 'مشاركة', refresh: 'تحديث', new: 'جديد', filters: 'الفلاتر',
    backend: 'قاعدة البيانات', connected: 'متصل', local: 'محلي',
    nav: { dashboard: 'الرئيسية', portfolio: 'المحفظة', board: 'البورد', tasks: 'المهام', risks: 'المخاطر', schedule: 'الجدول الزمني', teams: 'الموارد', reports: 'التقارير', files: 'الملفات', databases: 'القواعد', forms: 'الإدخال' },
    status: { all: 'الكل', on: 'على المسار', delayed: 'متأخر', blocked: 'خطر', done: 'مكتمل', open: 'مفتوح' },
    labels: { project: 'المشروع', owner: 'المالك', status: 'الحالة', priority: 'الأولوية', progress: 'الإنجاز', due: 'الموعد', actions: 'إجراءات', health: 'مؤشر صحة المشروع', budget: 'استخدام الميزانية', active: 'المشاريع النشطة', total: 'إجمالي المشاريع', risks: 'المخاطر المفتوحة', tasks: 'المهام المتأخرة', team: 'استخدام الفريق', milestones: 'المعالم المكتملة' }
  }
};

const FORM_CONFIG = {
  projects: { icon: BriefcaseBusiness, en: 'Project', ar: 'مشروع', fields: [['en', 'Project name', 'text'], ['ar', 'Arabic name', 'text'], ['owner', 'Project manager', 'text'], ['department', 'Department', 'text'], ['client', 'Client', 'text'], ['region', 'Region', 'text'], ['status', 'Status', 'select:on,delayed,blocked,done'], ['priority', 'Priority', 'select:Low,Medium,High,Critical'], ['progress', 'Progress %', 'number'], ['start', 'Start date', 'date'], ['due', 'Due date', 'date'], ['budget', 'Budget utilization %', 'number'], ['util', 'Team utilization %', 'number'], ['link', 'Files link', 'url'], ['notes', 'Notes', 'textarea']] },
  tasks: { icon: ClipboardCheck, en: 'Task', ar: 'مهمة', fields: [['project', 'Project', 'text'], ['title', 'Task title', 'text'], ['owner', 'Assigned to', 'text'], ['status', 'Status', 'select:open,on,delayed,blocked,done'], ['priority', 'Priority', 'select:Low,Medium,High,Critical'], ['progress', 'Progress %', 'number'], ['due', 'Due date', 'date'], ['notes', 'Notes', 'textarea']] },
  risks: { icon: AlertTriangle, en: 'Risk / Issue', ar: 'مخاطرة', fields: [['project', 'Project', 'text'], ['title', 'Risk / issue', 'text'], ['owner', 'Owner', 'text'], ['severity', 'Severity', 'select:Low,Medium,High,Critical'], ['status', 'Status', 'select:open,on,delayed,blocked,done'], ['impact', 'Impact', 'select:Low,Medium,High,Critical'], ['probability', 'Probability', 'select:Low,Medium,High,Critical'], ['due', 'Target date', 'date'], ['mitigation', 'Mitigation plan', 'textarea']] },
  schedules: { icon: CalendarDays, en: 'Schedule Item', ar: 'عنصر جدول', fields: [['project', 'Project', 'text'], ['milestone', 'Milestone / phase', 'text'], ['owner', 'Owner', 'text'], ['start', 'Start date', 'date'], ['due', 'End date', 'date'], ['status', 'Status', 'select:on,delayed,blocked,done'], ['progress', 'Progress %', 'number'], ['notes', 'Notes', 'textarea']] },
  resources: { icon: Users, en: 'Resource', ar: 'مورد', fields: [['name', 'Employee', 'text'], ['role', 'Role', 'text'], ['department', 'Department', 'text'], ['project', 'Project', 'text'], ['capacity', 'Capacity hours', 'number'], ['planned', 'Planned hours', 'number'], ['actual', 'Actual hours', 'number'], ['util', 'Utilization %', 'number'], ['status', 'Status', 'select:Available,Normal,Overloaded,Unavailable']] },
  reports: { icon: FileText, en: 'Status Report', ar: 'تقرير حالة', fields: [['project', 'Project', 'text'], ['date', 'Report date', 'date'], ['health', 'Overall health', 'select:on,delayed,blocked,done'], ['achievements', 'Achievements', 'textarea'], ['challenges', 'Challenges', 'textarea'], ['nextSteps', 'Next steps', 'textarea'], ['support', 'Support needed', 'textarea']] },
  files: { icon: Folder, en: 'File / Document', ar: 'ملف', fields: [['project', 'Project', 'text'], ['title', 'File title', 'text'], ['type', 'Document type', 'select:Schedule,Report,Contract,MOM,Deliverable,Other'], ['owner', 'Owner', 'text'], ['date', 'Updated date', 'date'], ['link', 'OneDrive / Teams / SharePoint link', 'url'], ['notes', 'Notes', 'textarea']] }
};

const SAMPLE = {
  source: { type: 'sample', updatedAt: new Date().toISOString() },
  projects: [
    { id: 'P-001', en: 'EPM-IDT Phase 3', ar: 'التحول الرقمي للحوكمة', owner: 'Mohamed Ahmed', department: 'Digital Transformation', client: 'EPM', region: 'EPM', status: 'on', priority: 'Medium', progress: 70, start: '2024-01-01', due: '2026-12-30', budget: 73, util: 70, tasks: 8, risks: 1, files: 12 },
    { id: 'P-002', en: 'EPM-Cleaning', ar: 'الإشراف على النظافة', owner: 'Mohamed Idris', department: 'Operations', client: 'EPM', region: 'EPM', status: 'delayed', priority: 'High', progress: 35, start: '2025-11-03', due: '2028-11-02', budget: 82, util: 79, tasks: 16, risks: 4, files: 22 },
    { id: 'P-003', en: 'EPM-940 Phase 5', ar: 'إدارة بلاغات 940', owner: 'Mohamed Al Khatem', department: 'Call Center', client: 'EPM', region: 'EPM', status: 'on', priority: 'Medium', progress: 92, start: '2025-12-10', due: '2028-12-09', budget: 66, util: 82, tasks: 7, risks: 1, files: 16 },
    { id: 'P-004', en: 'Hail Gardening', ar: 'حدائق حائل', owner: 'Khaled Sultan', department: 'Operations', client: 'Hail', region: 'Hail', status: 'on', priority: 'Medium', progress: 92, start: '2023-02-16', due: '2026-06-04', budget: 75, util: 87, tasks: 11, risks: 1, files: 14 },
    { id: 'P-005', en: 'Nazaha - 980', ar: 'نزاهة 980', owner: 'Mohamed Al Khatem', department: 'Contact Center', client: 'Nazaha', region: 'Nazaha', status: 'blocked', priority: 'Critical', progress: 20, start: '2026-05-06', due: '2028-05-05', budget: 55, util: 89, tasks: 12, risks: 5, files: 7 }
  ],
  tasks: [], risks: [], resources: [], reports: [], schedules: [], files: [], sheets: []
};

function norm(value) { return value == null ? '' : String(value).trim(); }
function num(value, fallback = 0) { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : fallback; }
function titleOf(row, lang) { return lang === 'ar' ? (row.ar || row.title || row.name || row.en || row.project || '—') : (row.en || row.title || row.name || row.ar || row.project || '—'); }
function statusOf(status, lang) { return COPY[lang].status[status] || status || '—'; }
function loadLocal() { try { return JSON.parse(localStorage.getItem(LS_KEY)) || EMPTY; } catch { return EMPTY; } }
function saveLocal(value) { localStorage.setItem(LS_KEY, JSON.stringify(value)); }
function uid(type) { return `${type.slice(0, 1).toUpperCase()}-${Date.now()}`; }
function downloadJson(name, data) { const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url); }

async function apiList(type) { const res = await fetch(`/api/records${type ? `?type=${type}` : ''}`, { cache: 'no-store' }); const json = await res.json(); if (!res.ok) throw new Error(json.error || 'API failed'); return json.rows || []; }
async function apiCreate(type, record) { const res = await fetch('/api/records', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type, record }) }); const json = await res.json(); if (!res.ok) throw new Error(json.error || 'API failed'); return json.row; }
async function apiDelete(id, type) { const res = await fetch('/api/records', { method: 'DELETE', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id, type }) }); const json = await res.json(); if (!res.ok) throw new Error(json.error || 'API failed'); return json; }

function Chip({ type, children }) { return <span className={`chip ${type}`}>{children}</span>; }
function Progress({ value }) { const v = Math.max(0, Math.min(100, num(value))); return <div className="progress"><span><i style={{ width: `${v}%` }} /></span><b>{v}%</b></div>; }
function Cmd({ icon: Icon, label, tone = '', onClick }) { return <button className={`commandButton ${tone}`} onClick={onClick}><Icon size={16} /><span>{label}</span></button>; }
function Stat({ icon: Icon, title, value, sub, tone }) { return <div className="statCard"><div className={`statIcon ${tone}`}><Icon size={22} /></div><div><p>{title}</p><strong>{value}</strong><span>{sub}</span></div></div>; }

function sheetToRows(ws) {
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  const headerIndex = Math.max(0, matrix.findIndex((row) => row.filter((cell) => norm(cell)).length >= 3));
  const headers = (matrix[headerIndex] || []).map((header, index) => norm(header) || `Column ${index + 1}`);
  const rows = matrix.slice(headerIndex + 1).map((row) => headers.map((_, index) => norm(row[index]))).filter((row) => row.some(Boolean));
  return { headers, rows };
}

function parseWorkbook(workbook, fileName) {
  const sheets = workbook.SheetNames.map((name) => ({ name, ...sheetToRows(workbook.Sheets[name]) }));
  const sourceRows = (sheets.find((sheet) => /project|erp/i.test(sheet.name))?.rows || sheets.flatMap((sheet) => sheet.rows.slice(0, 2)));
  const projects = sourceRows.slice(0, 60).map((row, index) => ({
    id: row[0] || `P-${index + 1}`,
    en: row[4] || row[1] || row[0] || `Project ${index + 1}`,
    ar: row[3] || row[1] || '',
    owner: row[7] || row[2] || 'Project Manager',
    department: row[1] || '', client: row[2] || '', region: row[2] || '',
    status: 'on', priority: 'Medium', progress: 60,
    start: row[5] || '', due: row[6] || '', budget: 70, util: 70, tasks: 0, risks: 0, files: 0
  }));
  return { ...SAMPLE, source: { type: 'excel', workbook: fileName, updatedAt: new Date().toISOString() }, sheets, projects: projects.length ? projects : SAMPLE.projects };
}

function TopBar({ t, lang, setLang, upload, backendState }) {
  return <header className="topBar"><div className="brand"><div className="brandLogo"><Workflow size={22} /></div><div><b>{t.app}</b><span>{t.sub}</span></div></div><div className="globalSearch"><Search size={18} /><input placeholder={t.search} /></div><div className="dataSource"><Database size={15} /><span>{t.backend}: {backendState === 'connected' ? t.connected : t.local}</span></div><div className="topActions"><label className="uploadBtn"><Upload size={16} />{t.upload}<input type="file" accept=".xlsx,.xls,.xlsm" onChange={upload} /></label><button onClick={() => setLang(lang === 'en' ? 'ar' : 'en')}><Languages size={16} />{lang === 'en' ? 'AR' : 'EN'}</button><button><Bell size={16} /></button><button><Settings size={16} /></button></div></header>;
}

function Sidebar({ t, page, setPage }) {
  const nav = [['dashboard', Home], ['portfolio', BriefcaseBusiness], ['board', KanbanSquare], ['tasks', ClipboardCheck], ['risks', AlertTriangle], ['schedule', CalendarDays], ['teams', Users], ['reports', BarChart3], ['files', Folder], ['databases', Database], ['forms', FormInput]];
  return <aside className="sidebar"><div className="workspaceCard"><Sparkles size={18} /><div><b>PMO Workspace</b><span>Projects • Teams • Delivery</span></div></div><nav>{nav.map(([key, Icon]) => <button key={key} className={page === key ? 'active' : ''} onClick={() => setPage(key)}><Icon size={18} /><span>{t.nav[key]}</span></button>)}</nav><div className="sideNote"><ShieldCheck size={18} /><p>Official PMO app for portfolio control, schedules, resources, risks, files and weekly reporting.</p></div></aside>;
}

function Hero({ t, page, setPage, refresh, exportAll }) {
  return <section className="pageHero"><div><div className="breadcrumb"><Home size={14} /><span>PMO</span><span>/</span><span>{t.nav[page]}</span></div><h1>{t.app}</h1><p>{t.sub}</p></div><div className="heroActions"><Cmd icon={Plus} label={t.new} tone="primary" onClick={() => setPage('forms')} /><Cmd icon={RefreshCcw} label={t.refresh} onClick={refresh} /><Cmd icon={Download} label={t.export} onClick={exportAll} /><Cmd icon={Share2} label={t.share} /></div></section>;
}

function Filters({ t, filters, setFilters, owners, departments, regions }) {
  return <section className="filterDock"><div><Filter size={16} /><b>{t.filters}</b></div><select value={filters.owner} onChange={(event) => setFilters({ ...filters, owner: event.target.value })}><option value="all">Project Manager: {t.status.all}</option>{owners.map((owner) => <option key={owner}>{owner}</option>)}</select><select value={filters.department} onChange={(event) => setFilters({ ...filters, department: event.target.value })}><option value="all">Department: {t.status.all}</option>{departments.map((department) => <option key={department}>{department}</option>)}</select><select value={filters.status} onChange={(event) => setFilters({ ...filters, status: event.target.value })}><option value="all">Project Status: {t.status.all}</option><option value="on">{t.status.on}</option><option value="delayed">{t.status.delayed}</option><option value="blocked">{t.status.blocked}</option><option value="done">{t.status.done}</option></select><select value={filters.region} onChange={(event) => setFilters({ ...filters, region: event.target.value })}><option value="all">Region: {t.status.all}</option>{regions.map((region) => <option key={region}>{region}</option>)}</select><button className="commandButton" onClick={() => setFilters({ owner: 'all', department: 'all', status: 'all', region: 'all' })}>Reset Filters</button></section>;
}

function groupAverage(rows, key, metric) { const groups = {}; rows.forEach((row) => { const name = row[key] || 'Other'; groups[name] ||= []; groups[name].push(num(row[metric], 60)); }); return Object.entries(groups).map(([name, values]) => ({ name, value: Math.round(values.reduce((a, b) => a + b, 0) / values.length) })).slice(0, 6); }
function MiniBars({ title, rows }) { return <div className="panelPro pm-card"><div className="panelHeader"><h3>{title}</h3><MoreHorizontal size={16} /></div><div className="miniBars">{rows.map((row) => <p key={row.name}><span>{row.name}</span><i><b style={{ width: `${row.value}%` }} /></i><em>{row.value}%</em></p>)}</div></div>; }
function HealthChart({ projects }) { const total = projects.length || 1; const on = projects.filter((p) => p.status === 'on').length; const delayed = projects.filter((p) => p.status === 'delayed').length; const blocked = projects.filter((p) => p.status === 'blocked').length; const done = projects.filter((p) => p.status === 'done').length; return <div className="panelPro pm-card"><div className="panelHeader"><h3>Projects by Progress</h3><MoreHorizontal size={16} /></div><div className="pulseBody"><div className="donutPro" style={{ background: `conic-gradient(#2563eb 0 ${(on / total) * 100}%,#f59e0b ${(on / total) * 100}% ${((on + delayed) / total) * 100}%,#ef4444 ${((on + delayed) / total) * 100}% ${((on + delayed + blocked) / total) * 100}%,#16a34a ${((on + delayed + blocked) / total) * 100}% 100%)` }}><span>{total}<small>Total</small></span></div><div className="legendPro"><p><i className="green" />On Track<b>{on}</b></p><p><i className="amber" />Delayed<b>{delayed}</b></p><p><i className="red" />At Risk<b>{blocked}</b></p><p><i className="gray" />Completed<b>{done}</b></p></div></div></div>; }
function Gantt({ projects }) { return <div className="panelPro pm-card pm-gantt"><div className="panelHeader"><h3>Project Gantt Chart</h3><MoreHorizontal size={16} /></div><div className="ganttHead"><span>Project Name</span><span>% Complete</span><span>Jan</span><span>Mar</span><span>May</span><span>Jul</span><span>Sep</span><span>Dec</span></div>{projects.slice(0, 7).map((project, index) => <div className="ganttRow" key={project.id || project.en}><b>{project.en || project.title}</b><em>{project.progress || 0}%</em><div className="ganttTrack"><i className={project.status} style={{ marginLeft: `${(index * 9) % 35}%`, width: `${Math.max(18, Math.min(70, num(project.progress)))}%` }} /></div></div>)}</div>; }
function HeatMap() { const values = ['Very high', 'High', 'Medium', 'Low', 'Very low']; return <div className="panelPro pm-card"><div className="panelHeader"><h3>Risk Heat Map</h3></div><div className="heatMap">{values.map((y, yIndex) => <React.Fragment key={y}>{values.map((x, xIndex) => <i key={`${x}-${y}`} className={`h${Math.min(4, Math.max(0, yIndex + xIndex - 2))}`} />)}</React.Fragment>)}</div></div>; }
function Trend() { return <div className="panelPro pm-card"><div className="panelHeader"><h3>Task Completion Trend</h3></div><div className="trendLine"><b>73%</b></div></div>; }
function Workload({ resources }) { const rows = resources.length ? resources : [{ name: 'Mohamed Idris', role: 'PM', util: 70 }, { name: 'Mohamed Ahmed', role: 'PM', util: 72 }, { name: 'Rami Mamoon', role: 'Lead', util: 78 }, { name: 'Khaled Sultan', role: 'Lead', util: 82 }]; return <div className="panelPro pm-card"><div className="panelHeader"><h3>Team Workload</h3></div><div className="workload">{rows.slice(0, 6).map((row) => <p key={row.name}><b>{row.name}</b><span>{row.role || row.department}</span><Progress value={row.util || row.load} /></p>)}</div></div>; }
function Milestones({ projects }) { return <div className="panelPro pm-card"><div className="panelHeader"><h3>Milestone Tracking</h3></div><div className="milestones"><div className="donutPro"><span>{Math.max(1, projects.filter((p) => p.progress > 80).length * 18)}</span></div><p><b>Completed</b><span>{projects.filter((p) => p.progress > 80).length}</span></p><p><b>In Progress</b><span>{projects.filter((p) => p.progress <= 80 && p.progress > 40).length}</span></p><p><b>Delayed</b><span>{projects.filter((p) => p.status !== 'on').length}</span></p></div></div>; }

function Dashboard({ t, projects, tasks, risks, resources }) {
  const total = projects.length;
  const active = projects.filter((project) => project.status !== 'done').length;
  const avg = total ? Math.round(projects.reduce((sum, project) => sum + num(project.progress), 0) / total) : 0;
  const util = resources.length ? Math.round(resources.reduce((sum, resource) => sum + num(resource.util), 0) / resources.length) : 68;
  const openRisks = risks.length || projects.reduce((sum, project) => sum + num(project.risks), 0);
  const delayed = tasks.filter((task) => task.status === 'delayed' || task.status === 'blocked').length || projects.filter((project) => project.status === 'delayed' || project.status === 'blocked').length * 3;
  const milestones = Math.max(1, projects.filter((project) => num(project.progress) > 80).length * 18);
  return <><section className="statsGrid pm-kpis"><Stat icon={BriefcaseBusiness} title={t.labels.total} value={total} sub="Portfolio scope" tone="blue" /><Stat icon={Folder} title={t.labels.active} value={active} sub="In delivery" tone="purple" /><Stat icon={Gauge} title={t.labels.budget} value={`${Math.round(projects.reduce((sum, project) => sum + num(project.budget), 0) / (total || 1))}%`} sub="Average" tone="green" /><Stat icon={BarChart3} title={t.labels.health} value={`${avg}/100`} sub="Portfolio score" tone="amber" /><Stat icon={AlertTriangle} title={t.labels.risks} value={openRisks} sub="Risk items" tone="red" /><Stat icon={ClipboardCheck} title={t.labels.tasks} value={delayed} sub="Need attention" tone="red" /><Stat icon={Users} title={t.labels.team} value={`${util}%`} sub="Capacity" tone="purple" /><Stat icon={CalendarDays} title={t.labels.milestones} value={milestones} sub="Tracked" tone="blue" /></section><section className="pm-grid"><Gantt projects={projects} /><HealthChart projects={projects} /><MiniBars title="Resource Allocation by Department" rows={groupAverage(projects, 'department', 'util')} /><HeatMap /><Milestones projects={projects} /><Trend /><MiniBars title="Project Status by Department" rows={groupAverage(projects, 'department', 'progress')} /><Workload resources={resources} /></section></>;
}

function Table({ t, lang, rows, title, onDelete, type }) { return <section className="panelPro tablePanel"><div className="panelHeader"><h3>{title}</h3><div className="viewTabs"><button className="active">Table</button><button>Detail</button></div></div><div className="tableWrap"><table><thead><tr><th>{t.labels.project}</th><th>{t.labels.owner}</th><th>{t.labels.status}</th><th>{t.labels.priority}</th><th>{t.labels.progress}</th><th>{t.labels.due}</th><th>{t.labels.actions}</th></tr></thead><tbody>{rows.map((row, index) => <tr key={row._backendId || row.id || index}><td><div className="projectTitle"><span>{row.id || row._backendId || `#${index + 1}`}</span><b>{titleOf(row, lang)}</b></div></td><td>{row.owner || row.name || '—'}</td><td><Chip type={row.status || row.health}>{statusOf(row.status || row.health, lang)}</Chip></td><td><Chip type={(row.priority || row.severity || 'Medium').toLowerCase()}>{row.priority || row.severity || 'Medium'}</Chip></td><td><Progress value={row.progress || row.util} /></td><td>{row.due || row.date || '—'}</td><td><div className="rowActions"><Eye size={16} /><Edit3 size={16} />{row.link && <a href={row.link} target="_blank" rel="noreferrer"><LinkIcon size={16} /></a>}{row._backendId && <button className="iconButton" onClick={() => onDelete(row._backendId, type)}><Trash2 size={16} /></button>}<MoreHorizontal size={16} /></div></td></tr>)}</tbody></table></div></section>; }
function Board({ t, lang, projects }) { const columns = [['blocked', t.status.blocked], ['delayed', t.status.delayed], ['on', t.status.on], ['done', t.status.done]]; return <section className="boardPage">{columns.map(([status, label]) => <div className={`boardColumn ${status}`} key={status}><div className="boardHead"><b>{label}</b><span>{projects.filter((project) => project.status === status).length}</span></div>{projects.filter((project) => project.status === status).map((project) => <div className="taskCard" key={project.id || project._backendId}><div className="taskTop"><b>{titleOf(project, lang)}</b><MoreHorizontal size={16} /></div><p>{project.owner}</p><Progress value={project.progress} /><div className="taskMeta"><Chip type={(project.priority || 'Medium').toLowerCase()}>{project.priority}</Chip><span>{project.due}</span></div></div>)}</div>)}</section>; }

function FormCenter({ t, lang, onSave, backendState }) {
  const [type, setType] = React.useState('projects');
  const [form, setForm] = React.useState({});
  const [saving, setSaving] = React.useState(false);
  const config = FORM_CONFIG[type];
  async function save() { setSaving(true); await onSave(type, { ...form, id: form.id || uid(type), createdAt: new Date().toISOString() }); setForm({}); setSaving(false); }
  return <section className="formsLayout"><div className="formNav panelPro"><h3>Input Center</h3>{Object.entries(FORM_CONFIG).map(([key, item]) => { const Icon = item.icon; return <button key={key} className={type === key ? 'active' : ''} onClick={() => setType(key)}><Icon size={17} /><span>{lang === 'ar' ? item.ar : item.en}</span><b>{key}</b></button>; })}</div><div className="formPanel panelPro"><div className="panelHeader"><h3>{lang === 'ar' ? config.ar : config.en}</h3><Chip type={backendState === 'connected' ? 'on' : 'delayed'}>{backendState === 'connected' ? 'Backend' : 'Local fallback'}</Chip></div><div className="smartForm">{config.fields.map(([key, label, definition]) => { const [kind, options] = definition.split(':'); return <label key={key} className={kind === 'textarea' ? 'wide' : ''}><span>{label}</span>{kind === 'textarea' ? <textarea value={form[key] || ''} onChange={(event) => setForm({ ...form, [key]: event.target.value })} /> : kind === 'select' ? <select value={form[key] || ''} onChange={(event) => setForm({ ...form, [key]: event.target.value })}><option value="">Select</option>{options.split(',').map((option) => <option key={option} value={option}>{option}</option>)}</select> : <input type={kind} value={form[key] || ''} onChange={(event) => setForm({ ...form, [key]: event.target.value })} />}</label>; })}</div><div className="formActions"><Cmd icon={Save} label={saving ? 'Saving...' : t.save} tone="primary" onClick={save} /></div></div><div className="reviewPanel panelPro"><div className="panelHeader"><h3>Actions</h3></div><div className="recordList"><div className="recordCard"><b>Backend API</b><p>/api/records</p><small>{backendState === 'connected' ? 'Supabase active' : 'Configure Supabase env vars in Vercel'}</small></div><div className="recordCard"><b>Working modules</b><p>Create, List, Delete, Export, Upload Excel, Dashboard</p></div></div></div></section>;
}

function Team({ resources }) { return <section className="teamGrid">{resources.map((resource, index) => <div className="teamCard" key={resource._backendId || resource.id || index}><div className="avatar">{norm(resource.name || 'T').slice(0, 1)}</div><h3>{resource.name}</h3><p>{resource.role || resource.department}</p><Progress value={resource.util || resource.load} /><Chip type={num(resource.util || resource.load) > 90 ? 'blocked' : num(resource.util || resource.load) > 75 ? 'delayed' : 'on'}>{resource.status || 'Available'}</Chip></div>)}</section>; }
function Reports({ t, projects, reports }) { const total = projects.length; const avg = total ? Math.round(projects.reduce((sum, project) => sum + num(project.progress), 0) / total) : 0; const attention = projects.filter((project) => ['blocked', 'delayed'].includes(project.status)).length; return <section className="reportGrid"><div className="panelPro reportHero"><h3>{t.labels.health}</h3><p>Executive summary generated from portfolio, schedules, resources, risks and reports.</p><div className="reportNumbers"><span><b>{total}</b>Projects</span><span><b>{avg}%</b>Progress</span><span><b>{attention}</b>Attention</span><span><b>{reports.length}</b>Reports</span></div><Cmd icon={Download} label={t.export} tone="primary" onClick={() => downloadJson('pmo-report.json', { projects, reports })} /></div><HealthChart projects={projects} /><Trend /></section>; }
function Files({ files, projects, lang }) { const rows = files.length ? files : projects.map((project) => ({ title: titleOf(project, lang), count: project.files, link: project.link })); return <section className="filesGrid">{rows.map((file, index) => <div className="fileCard" key={file._backendId || file.id || index}><Folder size={26} /><b>{file.title || file.name}</b><p>{file.type || file.count || 0} files</p><a href={file.link || '#'} target="_blank" rel="noreferrer"><LinkIcon size={14} /> OneDrive / Teams</a></div>)}</section>; }
function Databases({ t, data, team, backend, local, activeSheet, setActiveSheet, query, setQuery }) { const make = (name, items) => ({ name, headers: Object.keys(items[0] || {}), rows: items.map((item) => Object.keys(items[0] || {}).map((header) => item[header])) }); const sheets = [...(data.sheets || []), ...(team.sheets || []), ...TYPES.map((type) => make(`Backend ${type}`, backend[type] || [])).filter((sheet) => sheet.rows.length), ...TYPES.map((type) => make(`Local ${type}`, local[type] || [])).filter((sheet) => sheet.rows.length)]; const current = sheets.find((sheet) => sheet.name === activeSheet) || sheets[0]; const rows = current ? current.rows.filter((row) => row.some((cell) => norm(cell).toLowerCase().includes(query.toLowerCase()))).slice(0, 200) : []; return <section className="panelPro databasePage"><div className="databaseHero"><div><h3>Databases & Control Lists</h3><p>Backend, SharePoint, Excel and local records in one clean review area.</p></div><div className="dbStats"><span><b>{sheets.length}</b>Tables</span><span><b>{current?.rows?.length || 0}</b>Rows</span><span><b>{current?.headers?.length || 0}</b>Fields</span><span><b>{rows.length}</b>Results</span></div></div><div className="databaseToolbar"><div className="field"><label>Select database</label><select value={current?.name || ''} onChange={(event) => setActiveSheet(event.target.value)}>{sheets.map((sheet) => <option key={sheet.name} value={sheet.name}>{sheet.name} - {sheet.rows.length}</option>)}</select></div><div className="databaseSearch"><Search size={16} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder={t.search} /></div></div>{current ? <div className="tableWrap databaseTable"><table><thead><tr>{current.headers.map((header, index) => <th key={index}>{header}</th>)}</tr></thead><tbody>{rows.map((row, rowIndex) => <tr key={rowIndex}>{current.headers.map((header, columnIndex) => <td key={columnIndex}>{String(row[columnIndex] ?? '—')}</td>)}</tr>)}</tbody></table></div> : <div className="emptyState"><Database size={36} /><b>No data yet</b></div>}</section>; }

function App() {
  const [lang, setLang] = React.useState('en');
  const [page, setPage] = React.useState('dashboard');
  const [data, setData] = React.useState(SAMPLE);
  const [team, setTeam] = React.useState({ resources: [], sheets: [] });
  const [local, setLocal] = React.useState(loadLocal);
  const [backend, setBackend] = React.useState(EMPTY);
  const [backendState, setBackendState] = React.useState('fallback');
  const [activeSheet, setActiveSheet] = React.useState('');
  const [query, setQuery] = React.useState('');
  const [filters, setFilters] = React.useState({ owner: 'all', department: 'all', status: 'all', region: 'all' });
  const t = COPY[lang];

  async function refresh() { try { const rows = await apiList(); const next = { ...EMPTY }; rows.forEach((row) => { if (next[row.type]) next[row.type].push({ ...row.record, _backendId: row.id, _createdAt: row.created_at }); }); setBackend(next); setBackendState('connected'); } catch { setBackendState('fallback'); } }
  React.useEffect(() => { fetch('/data/sharepoint-data.json', { cache: 'no-store' }).then((res) => res.ok ? res.json() : null).then((json) => { if (json) { setData({ ...SAMPLE, ...json }); setActiveSheet(json.sheets?.[0]?.name || ''); } }).catch(() => {}); fetch('/data/team-data.json', { cache: 'no-store' }).then((res) => res.ok ? res.json() : null).then((json) => { if (json) setTeam(json); }).catch(() => {}); refresh(); }, []);
  async function onSave(type, record) { try { const row = await apiCreate(type, record); setBackend((previous) => ({ ...previous, [type]: [{ ...row.record, _backendId: row.id }, ...(previous[type] || [])] })); setBackendState('connected'); } catch { const next = { ...local, [type]: [record, ...(local[type] || [])] }; setLocal(next); saveLocal(next); setBackendState('fallback'); } }
  async function onDelete(id, type) { try { await apiDelete(id, type); setBackend((previous) => ({ ...previous, [type]: previous[type].filter((item) => item._backendId !== id) })); } catch (error) { alert(error.message); } }
  async function upload(event) { const file = event.target.files?.[0]; if (!file) return; const buffer = await file.arrayBuffer(); const workbook = XLSX.read(buffer, { type: 'array', cellDates: true }); const parsed = parseWorkbook(workbook, file.name); setData(parsed); setActiveSheet(parsed.sheets[0]?.name || ''); setPage('databases'); }

  const projects = [...data.projects, ...backend.projects, ...local.projects];
  const tasks = [...backend.tasks, ...local.tasks];
  const risks = [...backend.risks, ...local.risks];
  const resources = [...(team.resources || []), ...backend.resources, ...local.resources];
  const reports = [...backend.reports, ...local.reports];
  const files = [...backend.files, ...local.files];
  const schedules = [...backend.schedules, ...local.schedules];
  const unique = (key) => [...new Set(projects.map((project) => project[key]).filter(Boolean))];
  const filtered = projects.filter((project) => (filters.owner === 'all' || project.owner === filters.owner) && (filters.department === 'all' || project.department === filters.department) && (filters.status === 'all' || project.status === filters.status) && (filters.region === 'all' || project.region === filters.region));
  const exportAll = () => downloadJson('pmo-full-export.json', { data, team, backend, local, exportedAt: new Date().toISOString() });

  return <div className={`appShell ${lang === 'ar' ? 'rtl' : 'ltr'}`} dir={lang === 'ar' ? 'rtl' : 'ltr'}><TopBar t={t} lang={lang} setLang={setLang} upload={upload} backendState={backendState} /><Sidebar t={t} page={page} setPage={setPage} /><main className="mainArea"><Hero t={t} page={page} setPage={setPage} refresh={refresh} exportAll={exportAll} />{page === 'dashboard' && <><Filters t={t} filters={filters} setFilters={setFilters} owners={unique('owner')} departments={unique('department')} regions={unique('region')} /><Dashboard t={t} projects={filtered} tasks={tasks} risks={risks} resources={resources} /></>}{page === 'portfolio' && <Table t={t} lang={lang} rows={filtered} title={t.nav.portfolio} onDelete={onDelete} type="projects" />}{page === 'board' && <Board t={t} lang={lang} projects={filtered} />}{page === 'tasks' && <Table t={t} lang={lang} rows={tasks.map((task) => ({ ...task, en: task.title, ar: task.title }))} title={t.nav.tasks} onDelete={onDelete} type="tasks" />}{page === 'risks' && <Table t={t} lang={lang} rows={risks.map((risk) => ({ ...risk, en: risk.title, ar: risk.title, priority: risk.severity, progress: 0 }))} title={t.nav.risks} onDelete={onDelete} type="risks" />}{page === 'schedule' && <Table t={t} lang={lang} rows={schedules.map((schedule) => ({ ...schedule, en: schedule.milestone, ar: schedule.milestone, priority: 'Medium' }))} title={t.nav.schedule} onDelete={onDelete} type="schedules" />}{page === 'teams' && <Team resources={resources} />}{page === 'reports' && <Reports t={t} projects={filtered} reports={reports} />}{page === 'files' && <Files files={files} projects={filtered} lang={lang} />}{page === 'databases' && <Databases t={t} data={data} team={team} backend={backend} local={local} activeSheet={activeSheet} setActiveSheet={setActiveSheet} query={query} setQuery={setQuery} />}{page === 'forms' && <FormCenter t={t} lang={lang} onSave={onSave} backendState={backendState} />}</main></div>;
}

createRoot(document.getElementById('root')).render(<App />);
