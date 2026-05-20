import React from 'react';
import { createRoot } from 'react-dom/client';
import { motion, AnimatePresence } from 'framer-motion';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  Archive,
  ArrowUpRight,
  BarChart3,
  Bell,
  BriefcaseBusiness,
  Building2,
  CalendarDays,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  ClipboardCheck,
  Clock3,
  Copy,
  Database,
  Download,
  Edit3,
  Eye,
  FileText,
  Filter,
  Folder,
  Gauge,
  Grid2X2,
  Home,
  KanbanSquare,
  Languages,
  LayoutDashboard,
  LayoutList,
  LineChart,
  Link as LinkIcon,
  Mail,
  MessageSquare,
  MoreHorizontal,
  Plus,
  RefreshCcw,
  Search,
  Settings,
  Share2,
  ShieldCheck,
  Sparkles,
  Target,
  Trash2,
  Upload,
  Users,
  Workflow,
} from 'lucide-react';
import './styles.css';

const COPY = {
  ar: {
    appName: 'Project Control Hub',
    subtitle: 'نظام احترافي لإدارة المشاريع والفرق والملفات والتقارير',
    search: 'بحث عام في النظام',
    upload: 'رفع ملف Excel',
    newItem: 'إضافة عنصر',
    export: 'تصدير',
    share: 'مشاركة',
    refresh: 'تحديث',
    dashboard: 'الرئيسية',
    portfolio: 'المحفظة',
    board: 'لوحة التحكم',
    tasks: 'المهام',
    risks: 'المخاطر',
    teams: 'الفريق',
    reports: 'التقارير',
    files: 'الملفات',
    databases: 'القوائم',
    automation: 'الأتمتة',
    settings: 'الإعدادات',
    commandCenter: 'مركز التحكم التنفيذي',
    commandDesc: 'متابعة مباشرة للمشاريع، الأولويات، المخاطر، المهام، الموارد والتقارير من مكان واحد.',
    status: 'الحالة',
    manager: 'المدير',
    progress: 'الإنجاز',
    all: 'الكل',
    healthy: 'على المسار',
    delayed: 'متأخر',
    blocked: 'متعثر',
    done: 'مكتمل',
    projects: 'المشاريع',
    openTasks: 'المهام المفتوحة',
    highRisks: 'مخاطر عالية',
    utilization: 'استخدام الموارد',
    portfolioHealth: 'صحة المحفظة',
    attention: 'تحتاج متابعة',
    noData: 'ارفع ملف Excel لعرض بياناتك الحقيقية',
    boardView: 'عرض كانبان',
    tableView: 'عرض جدول',
    timelineView: 'Timeline',
    calendarView: 'تقويم',
    detailView: 'تفصيلي',
    project: 'المشروع',
    owner: 'المالك',
    due: 'الموعد',
    priority: 'الأولوية',
    actions: 'إجراءات',
    sheetSelect: 'اختر القائمة',
    viewType: 'نوع العرض',
    workflow: 'سير العمل',
    automationRules: 'قواعد التنبيه والمتابعة',
  },
  en: {
    appName: 'Project Control Hub',
    subtitle: 'Professional system for projects, teams, files, reporting and controls',
    search: 'Search workspace',
    upload: 'Upload Excel',
    newItem: 'New item',
    export: 'Export',
    share: 'Share',
    refresh: 'Refresh',
    dashboard: 'Home',
    portfolio: 'Portfolio',
    board: 'Control Board',
    tasks: 'Tasks',
    risks: 'Risks',
    teams: 'Team',
    reports: 'Reports',
    files: 'Files',
    databases: 'Databases',
    automation: 'Automation',
    settings: 'Settings',
    commandCenter: 'Executive Control Center',
    commandDesc: 'Live control for projects, priorities, risks, tasks, resources and reporting in one workspace.',
    status: 'Status',
    manager: 'Manager',
    progress: 'Progress',
    all: 'All',
    healthy: 'On Track',
    delayed: 'Delayed',
    blocked: 'Blocked',
    done: 'Completed',
    projects: 'Projects',
    openTasks: 'Open Tasks',
    highRisks: 'High Risks',
    utilization: 'Utilization',
    portfolioHealth: 'Portfolio Health',
    attention: 'Attention',
    noData: 'Upload Excel to show your real data',
    boardView: 'Board',
    tableView: 'Table',
    timelineView: 'Timeline',
    calendarView: 'Calendar',
    detailView: 'Detailed',
    project: 'Project',
    owner: 'Owner',
    due: 'Due',
    priority: 'Priority',
    actions: 'Actions',
    sheetSelect: 'Select database',
    viewType: 'View type',
    workflow: 'Workflow',
    automationRules: 'Automation & follow-up rules',
  },
};

const SAMPLE_PROJECTS = [
  { id: 'P-001', nameAr: 'بوابة العملاء', nameEn: 'Customer Portal', manager: 'Haitham Elmohamady', status: 'delayed', progress: 56, priority: 'High', due: '2026-06-15', budget: 72, utilization: 84, tasks: 18, risks: 3, files: 24 },
  { id: 'P-002', nameAr: 'نظام ERP', nameEn: 'ERP System', manager: 'Sara Al-Mutairi', status: 'healthy', progress: 76, priority: 'Medium', due: '2026-07-30', budget: 64, utilization: 73, tasks: 12, risks: 1, files: 18 },
  { id: 'P-003', nameAr: 'تحديث البنية التحتية', nameEn: 'Infrastructure Upgrade', manager: 'Mohammed Al-Harbi', status: 'blocked', progress: 31, priority: 'Critical', due: '2026-05-20', budget: 91, utilization: 96, tasks: 22, risks: 5, files: 37 },
  { id: 'P-004', nameAr: 'تطبيق الهاتف', nameEn: 'Mobile App', manager: 'Manal Al-Sharif', status: 'healthy', progress: 82, priority: 'Medium', due: '2026-08-10', budget: 55, utilization: 68, tasks: 9, risks: 1, files: 15 },
  { id: 'P-005', nameAr: 'تحليل البيانات والذكاء الاصطناعي', nameEn: 'AI Analytics', manager: 'Ali Al-Qahtani', status: 'done', progress: 100, priority: 'Low', due: '2026-04-30', budget: 49, utilization: 61, tasks: 4, risks: 0, files: 31 },
  { id: 'P-006', nameAr: 'مركز التقارير التنفيذية', nameEn: 'Executive Reporting Center', manager: 'Noura Al-Rashid', status: 'healthy', progress: 69, priority: 'High', due: '2026-09-01', budget: 80, utilization: 77, tasks: 16, risks: 2, files: 20 },
];

function norm(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).replace(/\s+/g, ' ').trim();
}

function normalizeHeader(header, index) {
  return norm(header) || `Column ${index + 1}`;
}

function sheetToRows(ws) {
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  const headerIndex = matrix.findIndex(row => row.filter(cell => norm(cell)).length >= 3);
  const safeHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
  const headers = (matrix[safeHeaderIndex] || []).map(normalizeHeader);
  const rows = matrix.slice(safeHeaderIndex + 1).map(row => headers.map((_, i) => norm(row[i]))).filter(row => row.some(Boolean));
  return { headers, rows };
}

function inferStatus(progress, index) {
  if (progress >= 96) return 'done';
  if (progress >= 70) return 'healthy';
  if (progress >= 45) return 'delayed';
  if (index % 4 === 0) return 'blocked';
  return 'delayed';
}

function progressFromText(text, fallback) {
  const value = norm(text).toLowerCase();
  const percent = value.match(/(\d{1,3})\s*%/);
  if (percent) return Math.min(100, Number(percent[1]));
  if (value.includes('completed') || value.includes('مكتمل') || value.includes('تم')) return 100;
  if (value.includes('deployment') || value.includes('go') || value.includes('اطلاق')) return 82;
  if (value.includes('development') || value.includes('تطوير')) return 62;
  if (value.includes('analysis') || value.includes('تحليل')) return 36;
  return fallback;
}

function parseWorkbook(workbook, fileName = 'Uploaded workbook') {
  const sheets = workbook.SheetNames.map(name => ({ name, ...sheetToRows(workbook.Sheets[name]) }));
  const erp = sheets.find(s => s.name === 'ERP Project');
  const rows = erp?.rows?.length ? erp.rows : sheets.flatMap(s => s.rows.slice(0, 3));
  const projects = rows.slice(0, 50).map((row, index) => {
    const text = row.join(' ');
    const fallback = Math.min(100, Math.max(18, 34 + ((index * 13) % 66)));
    const progress = progressFromText(text, fallback);
    const status = inferStatus(progress, index);
    return {
      id: `P-${String(index + 1).padStart(3, '0')}`,
      nameAr: row[1] || row[0] || `مشروع ${index + 1}`,
      nameEn: row[1] || row[0] || `Project ${index + 1}`,
      manager: row[7] || row[6] || 'Project Manager',
      status,
      progress,
      priority: status === 'blocked' ? 'Critical' : status === 'delayed' ? 'High' : status === 'healthy' ? 'Medium' : 'Low',
      due: row[10] || row[9] || '2026-12-30',
      budget: Math.min(100, 45 + ((index * 9) % 55)),
      utilization: Math.min(100, 50 + ((index * 7) % 45)),
      tasks: 5 + ((index * 3) % 24),
      risks: status === 'blocked' ? 5 : status === 'delayed' ? 2 : status === 'healthy' ? 1 : 0,
      files: 10 + ((index * 4) % 36),
    };
  });
  return { workbook: fileName, sheets, projects: projects.length ? projects : SAMPLE_PROJECTS };
}

function projectName(project, lang) {
  return lang === 'ar' ? project.nameAr : project.nameEn;
}

function statusLabel(status, lang) {
  const labels = {
    ar: { healthy: 'على المسار', delayed: 'متأخر', blocked: 'متعثر', done: 'مكتمل' },
    en: { healthy: 'On Track', delayed: 'Delayed', blocked: 'Blocked', done: 'Completed' },
  };
  return labels[lang][status] || status;
}

function priorityLabel(priority, lang) {
  const labels = {
    ar: { Critical: 'حرج', High: 'عالي', Medium: 'متوسط', Low: 'منخفض' },
    en: { Critical: 'Critical', High: 'High', Medium: 'Medium', Low: 'Low' },
  };
  return labels[lang][priority] || priority;
}

function Chip({ type, children }) {
  return <span className={`chip ${type}`}>{children}</span>;
}

function Progress({ value, tone = 'blue' }) {
  return <div className="progress"><span><i className={tone} style={{ width: `${Math.max(0, Math.min(100, value || 0))}%` }} /></span><b>{value || 0}%</b></div>;
}

function cellTone(value, header = '') {
  const text = norm(value).toLowerCase();
  const head = norm(header).toLowerCase();
  if (!text) return '';
  if (text.includes('completed') || text.includes('مكتمل') || text.includes('تم') || text.includes('on track')) return 'successCell';
  if (text.includes('blocked') || text.includes('critical') || text.includes('متعثر') || text.includes('حرج')) return 'dangerCell';
  if (text.includes('delayed') || text.includes('high') || text.includes('متأخر') || text.includes('قيد')) return 'warningCell';
  if (head.includes('date') || head.includes('تاريخ')) return 'dateCell';
  if (head.includes('progress') || head.includes('نسبة') || text.includes('%')) return 'progressCell';
  return '';
}

function renderCell(value, header) {
  const text = norm(value);
  if (!text) return <span className="emptyValue">—</span>;
  const numeric = Number(text.replace('%', ''));
  if ((cellTone(text, header) === 'progressCell' || norm(header).toLowerCase().includes('progress')) && !Number.isNaN(numeric)) return <Progress value={numeric} />;
  if (text.startsWith('http')) return <a className="linkPill" href={text} target="_blank" rel="noreferrer"><LinkIcon size={13}/>Open</a>;
  const tone = cellTone(text, header);
  if (['successCell', 'dangerCell', 'warningCell'].includes(tone)) return <span className={`miniChip ${tone}`}>{text}</span>;
  return text;
}

function StatCard({ icon: Icon, title, value, sub, tone }) {
  return <motion.div className="statCard" whileHover={{ y: -4 }}><div className={`statIcon ${tone}`}><Icon size={22}/></div><div><p>{title}</p><strong>{value}</strong><span>{sub}</span></div></motion.div>;
}

function CommandButton({ icon: Icon, label, tone = '', onClick }) {
  return <button className={`commandButton ${tone}`} onClick={onClick}><Icon size={16}/><span>{label}</span></button>;
}

function TopBar({ t, lang, setLang, uploadWorkbook }) {
  return <header className="topBar"><div className="brand"><div className="brandLogo"><Workflow size={23}/></div><div><b>{t.appName}</b><span>{t.subtitle}</span></div></div><div className="globalSearch"><Search size={18}/><input placeholder={t.search}/></div><div className="topActions"><label className="uploadBtn"><Upload size={16}/>{t.upload}<input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadWorkbook}/></label><button onClick={() => setLang(lang === 'ar' ? 'en' : 'ar')}><Languages size={17}/>{lang === 'ar' ? 'EN' : 'AR'}</button><button><Bell size={17}/></button><button><Settings size={17}/></button></div></header>;
}

function Sidebar({ t, page, setPage }) {
  const nav = [
    ['dashboard', Home, t.dashboard],
    ['portfolio', BriefcaseBusiness, t.portfolio],
    ['board', KanbanSquare, t.board],
    ['tasks', ClipboardCheck, t.tasks],
    ['risks', AlertTriangle, t.risks],
    ['teams', Users, t.teams],
    ['reports', BarChart3, t.reports],
    ['files', Folder, t.files],
    ['databases', Database, t.databases],
    ['automation', Workflow, t.automation],
    ['settings', Settings, t.settings],
  ];
  return <aside className="sidebar"><div className="workspaceCard"><Sparkles size={18}/><div><b>PMO Workspace</b><span>Solution Management</span></div></div><nav>{nav.map(([key, Icon, label]) => <button key={key} className={page === key ? 'active' : ''} onClick={() => setPage(key)}><Icon size={18}/><span>{label}</span></button>)}</nav><div className="sideNote"><ShieldCheck size={18}/><p>Secure by links and Microsoft/OneDrive permissions.</p></div></aside>;
}

function PageHeader({ t, page, setPage }) {
  return <section className="pageHero"><div><div className="breadcrumb"><Home size={14}/><span>PMO</span><ChevronLeft size={14}/><span>{page}</span></div><h1>{t.commandCenter}</h1><p>{t.commandDesc}</p></div><div className="heroActions"><CommandButton icon={Plus} label={t.newItem} tone="primary"/><CommandButton icon={RefreshCcw} label={t.refresh}/><CommandButton icon={Download} label={t.export}/><CommandButton icon={Share2} label={t.share}/></div></section>;
}

function Dashboard({ t, lang, projects, setPage }) {
  const total = projects.length;
  const openTasks = projects.reduce((a, p) => a + p.tasks, 0);
  const highRisks = projects.reduce((a, p) => a + p.risks, 0);
  const avg = total ? Math.round(projects.reduce((a, p) => a + p.progress, 0) / total) : 0;
  const util = total ? Math.round(projects.reduce((a, p) => a + p.utilization, 0) / total) : 0;
  return <>
    <section className="commandCenter">
      <div className="quickCommands">
        <CommandButton icon={Plus} label={lang === 'ar' ? 'مشروع جديد' : 'New project'} tone="primary"/>
        <CommandButton icon={LayoutList} label={t.tableView} onClick={() => setPage('portfolio')}/>
        <CommandButton icon={KanbanSquare} label={t.boardView} onClick={() => setPage('board')}/>
        <CommandButton icon={CalendarDays} label={t.calendarView}/>
        <CommandButton icon={BarChart3} label={lang === 'ar' ? 'تحليل المحفظة' : 'Portfolio analysis'}/>
        <CommandButton icon={Mail} label={lang === 'ar' ? 'تقرير أسبوعي' : 'Weekly report'}/>
      </div>
    </section>
    <section className="statsGrid">
      <StatCard icon={BriefcaseBusiness} title={t.projects} value={total} sub={t.portfolioHealth} tone="blue" />
      <StatCard icon={ClipboardCheck} title={t.openTasks} value={openTasks} sub={t.attention} tone="green" />
      <StatCard icon={AlertTriangle} title={t.highRisks} value={highRisks} sub={t.attention} tone="red" />
      <StatCard icon={Gauge} title={t.progress} value={`${avg}%`} sub={t.portfolioHealth} tone="purple" />
      <StatCard icon={Users} title={t.utilization} value={`${util}%`} sub="Team capacity" tone="amber" />
    </section>
    <section className="dashboardGridPro">
      <PortfolioPulse t={t} projects={projects}/>
      <SmartAlerts t={t} lang={lang} projects={projects}/>
      <WorkflowPanel t={t} lang={lang}/>
    </section>
    <ProjectTable t={t} lang={lang} projects={projects.slice(0, 6)}/>
  </>;
}

function PortfolioPulse({ t, projects }) {
  const total = projects.length || 1;
  const counts = {
    healthy: projects.filter(p => p.status === 'healthy').length,
    delayed: projects.filter(p => p.status === 'delayed').length,
    blocked: projects.filter(p => p.status === 'blocked').length,
    done: projects.filter(p => p.status === 'done').length,
  };
  return <div className="panelPro pulsePanel"><div className="panelHeader"><h3>{t.portfolioHealth}</h3><MoreHorizontal size={18}/></div><div className="pulseBody"><div className="donutPro" style={{ background: `conic-gradient(#448361 0 ${(counts.healthy/total)*100}%, #cb912f ${(counts.healthy/total)*100}% ${((counts.healthy+counts.delayed)/total)*100}%, #d44c47 ${((counts.healthy+counts.delayed)/total)*100}% ${((counts.healthy+counts.delayed+counts.blocked)/total)*100}%, #9b9a97 ${((counts.healthy+counts.delayed+counts.blocked)/total)*100}% 100%)` }}><span>{projects.length}</span></div><div className="legendPro"><p><i className="green"/>{t.healthy}<b>{counts.healthy}</b></p><p><i className="amber"/>{t.delayed}<b>{counts.delayed}</b></p><p><i className="red"/>{t.blocked}<b>{counts.blocked}</b></p><p><i className="gray"/>{t.done}<b>{counts.done}</b></p></div></div></div>;
}

function SmartAlerts({ t, lang, projects }) {
  const alerts = projects.filter(p => p.status === 'blocked' || p.status === 'delayed').slice(0, 4);
  return <div className="panelPro"><div className="panelHeader"><h3>{t.automationRules}</h3><Bell size={18}/></div><div className="alertList">{alerts.map(p => <div className="alertItem" key={p.id}><span className={`alertDot ${p.status}`}/><div><b>{projectName(p, lang)}</b><p>{lang === 'ar' ? 'تحتاج متابعة وتحديث خطة العمل' : 'Needs follow-up and action plan update'}</p></div><Chip type={p.status}>{statusLabel(p.status, lang)}</Chip></div>)}</div></div>;
}

function WorkflowPanel({ t, lang }) {
  const steps = lang === 'ar' ? ['تسجيل الطلب', 'تحليل', 'تطوير', 'اختبار', 'إطلاق', 'دعم'] : ['Intake', 'Analysis', 'Build', 'Test', 'Go Live', 'Support'];
  return <div className="panelPro"><div className="panelHeader"><h3>{t.workflow}</h3><Workflow size={18}/></div><div className="workflowSteps">{steps.map((s, i) => <div key={s} className={i < 3 ? 'active' : ''}><span>{i + 1}</span><p>{s}</p></div>)}</div></div>;
}

function ProjectTable({ t, lang, projects }) {
  return <section className="panelPro tablePanel"><div className="panelHeader"><h3>{t.portfolio}</h3><div className="viewTabs"><button className="active">{t.tableView}</button><button>{t.detailView}</button></div></div><div className="tableWrap"><table><thead><tr><th>{t.project}</th><th>{t.owner}</th><th>{t.status}</th><th>{t.priority}</th><th>{t.progress}</th><th>{t.due}</th><th>{t.actions}</th></tr></thead><tbody>{projects.map(p => <tr key={p.id}><td><div className="projectTitle"><span>{p.id}</span><b>{projectName(p, lang)}</b></div></td><td>{p.manager}</td><td><Chip type={p.status}>{statusLabel(p.status, lang)}</Chip></td><td><Chip type={p.priority.toLowerCase()}>{priorityLabel(p.priority, lang)}</Chip></td><td><Progress value={p.progress}/></td><td>{p.due}</td><td><div className="rowActions"><Eye size={16}/><Edit3 size={16}/><MoreHorizontal size={16}/></div></td></tr>)}</tbody></table></div></section>;
}

function BoardPage({ t, lang, projects }) {
  const columns = [
    ['blocked', t.blocked, projects.filter(p => p.status === 'blocked')],
    ['delayed', t.delayed, projects.filter(p => p.status === 'delayed')],
    ['healthy', t.healthy, projects.filter(p => p.status === 'healthy')],
    ['done', t.done, projects.filter(p => p.status === 'done')],
  ];
  return <section className="boardPage">{columns.map(([key, label, items]) => <div className={`boardColumn ${key}`} key={key}><div className="boardHead"><b>{label}</b><span>{items.length}</span></div>{items.map(p => <div className="taskCard" key={p.id}><div className="taskTop"><b>{projectName(p, lang)}</b><MoreHorizontal size={16}/></div><p>{p.manager}</p><Progress value={p.progress}/><div className="taskMeta"><Chip type={p.priority.toLowerCase()}>{priorityLabel(p.priority, lang)}</Chip><span>{p.due}</span></div></div>)}</div>)}</section>;
}

function DatabasesPage({ t, lang, data, activeSheet, setActiveSheet, sheetQuery, setSheetQuery }) {
  const currentSheet = data.sheets.find(s => s.name === activeSheet) || data.sheets[0];
  const rows = currentSheet ? currentSheet.rows.filter(row => row.some(cell => norm(cell).toLowerCase().includes(sheetQuery.toLowerCase()))).slice(0, 160) : [];
  return <section className="panelPro databasePage"><div className="databaseHero"><div><h3>{lang === 'ar' ? 'Databases وقوائم التحكم' : 'Databases & Control Lists'}</h3><p>{lang === 'ar' ? 'كل الشيتات تتحول إلى قوائم احترافية شبيهة Notion مع فلاتر وViews.' : 'Excel sheets become Notion-like databases with filters and views.'}</p></div><div className="dbStats"><span><b>{data.sheets.length}</b>{lang === 'ar' ? 'شيت' : 'Sheets'}</span><span><b>{currentSheet?.rows?.length || 0}</b>{lang === 'ar' ? 'صف' : 'Rows'}</span><span><b>{currentSheet?.headers?.length || 0}</b>{lang === 'ar' ? 'عمود' : 'Fields'}</span><span><b>{rows.length}</b>{lang === 'ar' ? 'نتيجة' : 'Results'}</span></div></div><div className="databaseToolbar"><div className="field"><label>{t.sheetSelect}</label><select value={currentSheet?.name || ''} onChange={e => setActiveSheet(e.target.value)}>{data.sheets.length ? data.sheets.map(s => <option key={s.name} value={s.name}>{s.name} - {s.rows.length}</option>) : <option>{lang === 'ar' ? 'ارفع ملف أولاً' : 'Upload a file first'}</option>}</select></div><div className="field"><label>{t.viewType}</label><select><option>{t.tableView}</option><option>{t.boardView}</option><option>{t.calendarView}</option><option>{t.timelineView}</option></select></div><div className="databaseSearch"><Search size={16}/><input value={sheetQuery} onChange={e => setSheetQuery(e.target.value)} placeholder={t.search}/></div></div><div className="databaseTabs">{data.sheets.map(s => <button key={s.name} className={currentSheet?.name === s.name ? 'active' : ''} onClick={() => setActiveSheet(s.name)}>{s.name}<span>{s.rows.length}</span></button>)}</div>{currentSheet ? <div className="tableWrap databaseTable"><table><thead><tr>{currentSheet.headers.map((h, i) => <th key={`${h}-${i}`}>{h}</th>)}</tr></thead><tbody>{rows.map((row, i) => <tr key={i}>{currentSheet.headers.map((h, j) => <td key={`${i}-${j}`} className={cellTone(row[j], h)}>{renderCell(row[j], h)}</td>)}</tr>)}</tbody></table></div> : <div className="emptyState"><Upload size={36}/><b>{t.noData}</b></div>}</section>;
}

function FunctionPage({ t, lang }) {
  const actions = [
    [Plus, lang === 'ar' ? 'إضافة مشروع' : 'Add project', 'primary'],
    [ClipboardCheck, lang === 'ar' ? 'إضافة مهمة' : 'Add task', ''],
    [AlertTriangle, lang === 'ar' ? 'تسجيل مخاطرة' : 'Log risk', 'danger'],
    [FileText, lang === 'ar' ? 'تقرير أسبوعي' : 'Weekly report', ''],
    [CalendarDays, lang === 'ar' ? 'جدولة اجتماع' : 'Schedule meeting', ''],
    [MessageSquare, lang === 'ar' ? 'إرسال تنبيه' : 'Send alert', ''],
    [Filter, lang === 'ar' ? 'فلتر متقدم' : 'Advanced filter', ''],
    [Copy, lang === 'ar' ? 'نسخ قالب' : 'Duplicate template', ''],
    [Archive, lang === 'ar' ? 'أرشفة' : 'Archive', ''],
    [Trash2, lang === 'ar' ? 'حذف' : 'Delete', 'danger'],
    [Download, t.export, ''],
    [Share2, t.share, ''],
  ];
  return <section className="panelPro functionsPage"><div className="databaseHero"><div><h3>{lang === 'ar' ? 'مركز الفانكشن والتحكم' : 'Functions & Controls Center'}</h3><p>{lang === 'ar' ? 'أزرار تشغيل جاهزة لإدارة المشاريع والمهام والمخاطر والتقارير.' : 'Operational buttons for project, task, risk and reporting control.'}</p></div></div><div className="functionsGrid">{actions.map(([Icon, label, tone]) => <CommandButton key={label} icon={Icon} label={label} tone={tone}/>)}</div></section>;
}

function Placeholder({ label, lang }) {
  return <section className="panelPro placeholder"><Sparkles size={34}/><h3>{label}</h3><p>{lang === 'ar' ? 'صفحة جاهزة للتطوير والربط مع القوائم، الملفات، Microsoft Lists، أو Excel.' : 'Ready page for linking databases, files, Microsoft Lists or Excel.'}</p></section>;
}

function App() {
  const [lang, setLang] = React.useState('ar');
  const [page, setPage] = React.useState('dashboard');
  const [data, setData] = React.useState({ workbook: 'Sample PMO Workspace', sheets: [], projects: SAMPLE_PROJECTS });
  const [activeSheet, setActiveSheet] = React.useState('');
  const [sheetQuery, setSheetQuery] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState('all');
  const [managerFilter, setManagerFilter] = React.useState('all');
  const [progressFilter, setProgressFilter] = React.useState('all');
  const t = COPY[lang];
  const dir = lang === 'ar' ? 'rtl' : 'ltr';
  const managers = ['all', ...Array.from(new Set(data.projects.map(p => p.manager)))];
  const filtered = data.projects.filter(p => {
    const statusOk = statusFilter === 'all' || p.status === statusFilter;
    const managerOk = managerFilter === 'all' || p.manager === managerFilter;
    const progressOk = progressFilter === 'all' || (progressFilter === 'low' ? p.progress < 50 : progressFilter === 'mid' ? p.progress >= 50 && p.progress <= 80 : p.progress > 80);
    return statusOk && managerOk && progressOk;
  });

  async function uploadWorkbook(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const parsed = parseWorkbook(workbook, file.name);
    setData(parsed);
    setActiveSheet(parsed.sheets[0]?.name || '');
    setPage('databases');
  }

  return <div className={`appShell ${lang}`} dir={dir}>
    <TopBar t={t} lang={lang} setLang={setLang} uploadWorkbook={uploadWorkbook}/>
    <Sidebar t={t} page={page} setPage={setPage}/>
    <main className="mainArea">
      <PageHeader t={t} page={page} setPage={setPage}/>
      {page === 'dashboard' && <><section className="filterDock"><div><Filter size={16}/><b>{lang === 'ar' ? 'فلاتر ديناميكية' : 'Dynamic filters'}</b></div><select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}><option value="all">{t.all}</option><option value="healthy">{t.healthy}</option><option value="delayed">{t.delayed}</option><option value="blocked">{t.blocked}</option><option value="done">{t.done}</option></select><select value={managerFilter} onChange={e=>setManagerFilter(e.target.value)}>{managers.map(m => <option value={m} key={m}>{m === 'all' ? t.all : m}</option>)}</select><select value={progressFilter} onChange={e=>setProgressFilter(e.target.value)}><option value="all">{t.all}</option><option value="low">&lt; 50%</option><option value="mid">50% - 80%</option><option value="high">&gt; 80%</option></select></section><Dashboard t={t} lang={lang} projects={filtered} setPage={setPage}/></>}
      {page === 'portfolio' && <ProjectTable t={t} lang={lang} projects={filtered}/>}      
      {page === 'board' && <BoardPage t={t} lang={lang} projects={filtered}/>}      
      {page === 'databases' && <DatabasesPage t={t} lang={lang} data={data} activeSheet={activeSheet} setActiveSheet={setActiveSheet} sheetQuery={sheetQuery} setSheetQuery={setSheetQuery}/>}      
      {page === 'settings' && <FunctionPage t={t} lang={lang}/>}      
      {!['dashboard','portfolio','board','databases','settings'].includes(page) && <Placeholder label={t[page] || page} lang={lang}/>}      
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App />);
