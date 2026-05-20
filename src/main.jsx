import React from 'react';
import { createRoot } from 'react-dom/client';
import { motion } from 'framer-motion';
import {
  AlertTriangle,
  BarChart3,
  CalendarDays,
  ClipboardList,
  ExternalLink,
  FileText,
  FolderOpen,
  Gauge,
  Home,
  KanbanSquare,
  Link as LinkIcon,
  Mail,
  PlusCircle,
  Search,
  ShieldCheck,
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

const sourceInfo = {
  "workbook": "Solution Department - Projects(2).xlsx",
  "projects": 24,
  "sourceTasks": 383,
  "displayTasks": 120,
  "displayRisks": 40,
  "updatedAt": "2026-05-20"
};

const projects = [
  {
    "id": "P001",
    "name": "EPM-IDT Phase 3",
    "arabicName": "مشروع التحول الرقمي في عمليات الحوكمة والرقابة على المشاريع الاستثمارية بامانة المنطقة الشرقية",
    "region": "EPM",
    "businessLine": "Le Source",
    "owner": "Mohamed Ahmed",
    "status": "Deployment",
    "health": "Green",
    "progress": 96,
    "risk": "Medium",
    "openTasks": 3,
    "overdue": 0,
    "criticalRisks": 1,
    "taskCount": 28,
    "start": "2023-12-31",
    "end": "2026-12-30",
    "applications": "Investment Management System \"Fortunail\"\n• FOX ERP\n• Opward \n• Infoore\n• Axionic",
    "files": "https://teams.microsoft.com/",
    "schedule": "https://teams.microsoft.com/",
    "reports": "https://teams.microsoft.com/",
    "docs": "https://teams.microsoft.com/",
    "sheet": "EPM-IDT Phase 3"
  },
  {
    "id": "P002",
    "name": "EPM- EPM III - HR",
    "arabicName": "مشروع زين الشرقية",
    "region": "EPM",
    "businessLine": "Le Sys",
    "owner": "Rami Mamoon",
    "status": "UAT",
    "health": "Green",
    "progress": 86,
    "risk": "High",
    "openTasks": 18,
    "overdue": 2,
    "criticalRisks": 7,
    "taskCount": 28,
    "start": "2021-11-01",
    "end": "2024-10-30",
    "applications": "•  FOX ERP",
    "files": "https://teams.microsoft.com/",
    "schedule": "https://teams.microsoft.com/",
    "reports": "https://teams.microsoft.com/",
    "docs": "https://teams.microsoft.com/",
    "sheet": "EPM- EPM III - HR"
  }
];

const tasks = [
  {
    "id": "P003-T016",
    "project": "EPM- EPM III - Finance",
    "title": "أوامر التغير",
    "assignee": "Ahmed Al Said",
    "status": "انشاء ملف متطلبات",
    "phase": "2- Analysis",
    "due": "No due date",
    "priority": "High",
    "progress": 0,
    "note": "فور الانتهاء من مراجعة العقود سوف يتم عقد اجتماع لمناقشة ما تم انجازة لاوامر التغير",
    "sheet": "EPM- EPM III - Finance"
  },
  {
    "id": "P003-T037",
    "project": "EPM- EPM III - Finance",
    "title": "التكامل مع فواتير بلدي",
    "assignee": "Ahmed Al Said",
    "status": "انشاء ملف متطلبات",
    "phase": "2- Analysis",
    "due": "No due date",
    "priority": "High",
    "progress": 0,
    "note": "طلب تغيري - الربط مع منصة بلدي",
    "sheet": "EPM- EPM III - Finance"
  }
];

const risks = [
  {
    "id": "R001",
    "project": "EPM-IDT Phase 3",
    "title": "الرسائل التفاعلية (تشات بوت)",
    "severity": "Critical",
    "owner": "Mohamed Ahmed",
    "status": "قيد التطوير",
    "note": "تم طلب الرقم الخاص بالشات بوت ومن ثم سيتم البدء التهيئة"
  },
  {
    "id": "R002",
    "project": "EPM- EPM III - HR",
    "title": "اعارة الموظفين",
    "severity": "High",
    "owner": "أ.عبد الله الدخيل",
    "status": "في مرحلة التطوير",
    "note": "في مرحلة التطوير"
  }
];

function pillClass(value) {
  const map = {
    Green: 'pill green', Amber: 'pill amber', Red: 'pill red',
    'On Track': 'pill green', 'At Risk': 'pill amber', Delayed: 'pill red',
    Deployment: 'pill green', UAT: 'pill blue', 'Go live': 'pill green', 'Go-Live': 'pill green',
    Support: 'pill green', Analysis: 'pill amber', Training: 'pill blue', Infra: 'pill neutral',
    planning: 'pill amber', Critical: 'pill red', High: 'pill orange', Medium: 'pill blue', Low: 'pill neutral',
    'تم التطوير': 'pill green', 'تم التطوير ': 'pill green', 'تم التسليم': 'pill green', 'قبول': 'pill green',
    'تحت التطوير': 'pill amber', 'في مرحلة التطوير': 'pill blue', 'قيد التطوير': 'pill blue',
    'جاري التحليل': 'pill amber', 'انشاء ملف متطلبات': 'pill amber', 'الدعم': 'pill green', 'الدغم': 'pill amber',
    Completed: 'pill green', Blocked: 'pill red', Overdue: 'pill red', 'In Progress': 'pill blue', New: 'pill neutral'
  };
  return map[value] || 'pill neutral';
}

function Pill({ children }) { return <span className={pillClass(children)}>{children}</span>; }
function ButtonLink({ href, children, icon: Icon = ExternalLink }) { return <a className="btn" href={href} target="_blank" rel="noreferrer"><Icon size={16} />{children}</a>; }
function Metric({ icon: Icon, label, value, sub }) { return <div className="metric"><div><p>{label}</p><strong>{value}</strong><span>{sub}</span></div><div className="metricIcon"><Icon size={23} /></div></div>; }
function shortText(text, length = 120) {
  if (!text) return '';
  return text.length > length ? `${text.slice(0, length)}...` : text;
}

function App() {
  const [query, setQuery] = React.useState('');
  const [selected, setSelected] = React.useState(projects[0]);
  const summary = React.useMemo(() => ({
    total: projects.length,
    atRisk: projects.filter(p => p.health === 'Amber' || p.health === 'Red').length,
    openTasks: projects.reduce((a,p)=>a+p.openTasks,0),
    overdue: projects.reduce((a,p)=>a+p.overdue,0),
    criticalRisks: projects.reduce((a,p)=>a+p.criticalRisks,0),
    avgProgress: Math.round(projects.reduce((a,p)=>a+p.progress,0)/projects.length),
  }), []);
  const filtered = projects.filter(p => `${p.name} ${p.arabicName} ${p.owner} ${p.status} ${p.region}`.toLowerCase().includes(query.toLowerCase()));
  const selectedTasks = tasks.filter(t => t.project === selected.name).slice(0, 6);
  const selectedRisks = risks.filter(r => r.project === selected.name).slice(0, 5);
  const visibleTasks = tasks.slice(0, 30);

  return <div className="app">
    <aside className="sidebar">
      <div className="brand"><div className="logo"><Gauge size={23}/></div><div><h1>{CONFIG.portalName}</h1><p>{CONFIG.companyName}</p></div></div>
      <nav>{[[Home,'Dashboard'],[FolderOpen,'Projects'],[KanbanSquare,'Tasks Board'],[AlertTriangle,'Risks & Issues'],[CalendarDays,'Schedule'],[Users,'Resources'],[FileText,'Weekly Reports'],[ShieldCheck,'Governance']].map(([Icon,label])=><a key={label} href={`#${label.replaceAll(' ','-')}`}><Icon size={18}/><span>{label}</span></a>)}</nav>
      <div className="sideCard"><h3>Microsoft Live Links</h3><p>Secure data remains inside Microsoft. This portal opens Lists, Forms, and Teams files using existing permissions.</p><ButtonLink href={CONFIG.microsoftLists.tasks} icon={ClipboardList}>Open Tasks List</ButtonLink><ButtonLink href={CONFIG.forms.addTask} icon={PlusCircle}>Add Task Form</ButtonLink></div>
    </aside>
    <main>
      <motion.section className="hero" initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} transition={{duration:.45}}>
        <div><p className="eyebrow">Updated from Excel workbook • {sourceInfo.updatedAt}</p><h2>Solution Department PMO Dashboard</h2><p className="intro">Portfolio dashboard loaded from <b>{sourceInfo.workbook}</b>. It includes {sourceInfo.projects} projects and {sourceInfo.sourceTasks} tracked activities from your project sheets.</p></div>
        <div className="actions"><ButtonLink href={CONFIG.forms.addTask} icon={PlusCircle}>Add Task</ButtonLink><ButtonLink href={CONFIG.forms.weeklyReport} icon={Mail}>Submit Weekly Report</ButtonLink><ButtonLink href={CONFIG.forms.registerRisk} icon={AlertTriangle}>Register Risk</ButtonLink></div>
      </motion.section>
      <section className="metrics"><Metric icon={FolderOpen} label="Total Projects" value={summary.total} sub="From ERP Project sheet"/><Metric icon={ClipboardList} label="Open Activities" value={summary.openTasks} sub={`${summary.overdue} overdue activities`}/><Metric icon={AlertTriangle} label="Attention Items" value={summary.criticalRisks} sub={`${summary.atRisk} projects amber/red`}/><Metric icon={BarChart3} label="Portfolio Progress" value={`${summary.avgProgress}%`} sub="Average completion from sheets"/></section>
      <section className="gridMain">
        <div className="card wide" id="Projects"><div className="cardHead"><div><h3>Projects Portfolio</h3><p>Data imported from ERP Project and project activity sheets.</p></div><div className="search"><Search size={16}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search projects, owner, region..."/></div></div><div className="tableWrap"><table><thead><tr><th>Project</th><th>Region</th><th>Owner</th><th>Status</th><th>Progress</th><th>Files</th></tr></thead><tbody>{filtered.map(p=><tr key={p.id} onClick={()=>setSelected(p)} className={selected.id===p.id?'selected':''}><td><b>{p.name}</b><small>{p.arabicName || p.id}</small></td><td>{p.region || '-'}</td><td>{p.owner}</td><td><Pill>{p.status}</Pill></td><td><div className="progress"><span><i style={{width:`${p.progress}%`}}/></span>{p.progress}%</div></td><td><a href={p.files} target="_blank" rel="noreferrer"><LinkIcon size={14}/> Open</a></td></tr>)}</tbody></table></div></div>
        <div className="card"><div className="cardHead"><div><h3>Project Workspace</h3><p>{selected.name}</p></div><Pill>{selected.health}</Pill></div><div className="workspace"><div className="workTop"><span>Progress</span><b>{selected.progress}%</b></div><div className="bar"><i style={{width:`${selected.progress}%`}}/></div><div className="mini"><div><b>{selected.openTasks}</b><span>Open</span></div><div><b>{selected.overdue}</b><span>Overdue</span></div><div><b>{selected.criticalRisks}</b><span>Attention</span></div></div></div><div className="projectMeta"><p><b>Arabic name:</b> {selected.arabicName || '-'}</p><p><b>Applications:</b> {shortText(selected.applications, 160) || '-'}</p><p><b>Sheet:</b> {selected.sheet || '-'}</p></div><div className="linkGrid"><ButtonLink href={selected.files} icon={FolderOpen}>Main Files</ButtonLink><ButtonLink href={selected.schedule} icon={CalendarDays}>Project Schedule</ButtonLink><ButtonLink href={selected.reports} icon={FileText}>Status Reports</ButtonLink><ButtonLink href={selected.docs} icon={FileText}>Documentations</ButtonLink></div></div>
      </section>
      <section className="twoCols"><div className="card" id="Tasks-Board"><div className="cardHead"><div><h3>Task Control</h3><p>Top attention activities from the uploaded Excel workbook.</p></div><ButtonLink href={CONFIG.microsoftLists.tasks}>Open List</ButtonLink></div><div className="list">{visibleTasks.map(t=><div className="item" key={t.id}><div><b>{t.title}</b><p>{t.project} • {t.assignee} • {t.due}</p><p>{t.phase}</p></div><div><Pill>{t.status}</Pill><Pill>{t.priority}</Pill></div></div>)}</div></div><div className="card" id="Risks-&-Issues"><div className="cardHead"><div><h3>Risks & Issues</h3><p>Generated attention list from low-progress or pending activities.</p></div><ButtonLink href={CONFIG.forms.registerRisk} icon={PlusCircle}>Register</ButtonLink></div><div className="list">{risks.slice(0, 20).map(r=><div className="item" key={r.id}><div><b>{r.title}</b><p>{r.project} • Owner: {r.owner}</p><p>{shortText(r.note, 120)}</p></div><Pill>{r.severity}</Pill></div>)}</div></div></section>
      <section className="twoCols"><div className="card"><div className="cardHead"><div><h3>Selected Project Tasks</h3><p>Activities for {selected.name}</p></div><Pill>{selectedTasks.length}</Pill></div><div className="list">{selectedTasks.length ? selectedTasks.map(t=><div className="item" key={t.id}><div><b>{t.title}</b><p>{t.assignee} • {t.phase}</p></div><div><Pill>{t.status}</Pill></div></div>) : <div className="item"><div><b>No priority tasks displayed</b><p>Open the project sheet or Microsoft List for full details.</p></div></div>}</div></div><div className="card"><div className="cardHead"><div><h3>Selected Project Risks</h3><p>Attention items for {selected.name}</p></div><Pill>{selectedRisks.length}</Pill></div><div className="list">{selectedRisks.length ? selectedRisks.map(r=><div className="item" key={r.id}><div><b>{r.title}</b><p>{r.owner} • {r.status}</p></div><Pill>{r.severity}</Pill></div>) : <div className="item"><div><b>No generated risks</b><p>This project has no low-progress risks in the displayed set.</p></div></div>}</div></div></section>
      <section className="report" id="Weekly-Reports"><div><h3>Weekly Reporting Center</h3><p>Current dashboard is loaded from the uploaded workbook. Next step is replacing placeholder Microsoft Lists, Forms, and Teams links with your live links for real-time operations.</p></div><div className="actions"><a className="primary" href={CONFIG.forms.weeklyReport} target="_blank" rel="noreferrer"><Mail size={16}/> Submit Report</a><a className="darkBtn" href={CONFIG.microsoftLists.reports} target="_blank" rel="noreferrer"><FileText size={16}/> Open Reports List</a></div></section>
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App />);
