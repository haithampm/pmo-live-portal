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

const projects = [
  { id: 'P001', name: 'EPM-Smart Lightning', owner: 'Haitham Elmohamady', status: 'On Track', health: 'Green', progress: 68, risk: 'Medium', openTasks: 14, overdue: 2, criticalRisks: 1, files: 'https://teams.microsoft.com/', schedule: 'https://teams.microsoft.com/', reports: 'https://teams.microsoft.com/', docs: 'https://teams.microsoft.com/' },
  { id: 'P002', name: 'Hail Projects', owner: 'Project Manager', status: 'At Risk', health: 'Amber', progress: 46, risk: 'High', openTasks: 22, overdue: 5, criticalRisks: 3, files: 'https://teams.microsoft.com/', schedule: 'https://teams.microsoft.com/', reports: 'https://teams.microsoft.com/', docs: 'https://teams.microsoft.com/' },
  { id: 'P003', name: 'EPM Zain ERP', owner: 'Project Manager', status: 'Delayed', health: 'Red', progress: 38, risk: 'Critical', openTasks: 19, overdue: 7, criticalRisks: 4, files: 'https://teams.microsoft.com/', schedule: 'https://teams.microsoft.com/', reports: 'https://teams.microsoft.com/', docs: 'https://teams.microsoft.com/' },
  { id: 'P004', name: 'HBM Cleaning Project', owner: 'Project Manager', status: 'On Track', health: 'Green', progress: 81, risk: 'Low', openTasks: 9, overdue: 0, criticalRisks: 0, files: 'https://teams.microsoft.com/', schedule: 'https://teams.microsoft.com/', reports: 'https://teams.microsoft.com/', docs: 'https://teams.microsoft.com/' },
];

const tasks = [
  { id: 'T001', project: 'EPM-Smart Lightning', title: 'Review project schedule baseline', assignee: 'Ahmed', status: 'In Progress', due: '2026-05-24', priority: 'High' },
  { id: 'T002', project: 'Hail Projects', title: 'Submit weekly project status', assignee: 'Sara', status: 'Blocked', due: '2026-05-21', priority: 'Critical' },
  { id: 'T003', project: 'EPM Zain ERP', title: 'Upload delivery notes', assignee: 'Mohamed', status: 'Overdue', due: '2026-05-18', priority: 'High' },
  { id: 'T004', project: 'HBM Cleaning Project', title: 'Confirm documentation folder', assignee: 'Noura', status: 'Completed', due: '2026-05-20', priority: 'Medium' },
  { id: 'T005', project: 'EPM-Smart Lightning', title: 'Prepare next week action plan', assignee: 'Haitham', status: 'New', due: '2026-05-27', priority: 'Medium' },
];

const risks = [
  { id: 'R001', project: 'EPM Zain ERP', title: 'Schedule slippage due to pending approvals', severity: 'Critical', owner: 'PMO', status: 'Open' },
  { id: 'R002', project: 'Hail Projects', title: 'Resource availability conflict', severity: 'High', owner: 'Delivery Lead', status: 'In Progress' },
  { id: 'R003', project: 'EPM-Smart Lightning', title: 'Missing project status document update', severity: 'Medium', owner: 'Project Owner', status: 'Open' },
];

function pillClass(value) {
  const map = {
    Green: 'pill green', Amber: 'pill amber', Red: 'pill red',
    'On Track': 'pill green', 'At Risk': 'pill amber', Delayed: 'pill red',
    Completed: 'pill neutral', Blocked: 'pill red', Overdue: 'pill red', Critical: 'pill red',
    High: 'pill orange', Medium: 'pill blue', Low: 'pill neutral', 'In Progress': 'pill blue', New: 'pill neutral'
  };
  return map[value] || 'pill neutral';
}

function Pill({ children }) { return <span className={pillClass(children)}>{children}</span>; }
function ButtonLink({ href, children, icon: Icon = ExternalLink }) { return <a className="btn" href={href} target="_blank" rel="noreferrer"><Icon size={16} />{children}</a>; }
function Metric({ icon: Icon, label, value, sub }) { return <div className="metric"><div><p>{label}</p><strong>{value}</strong><span>{sub}</span></div><div className="metricIcon"><Icon size={23} /></div></div>; }

function App() {
  const [query, setQuery] = React.useState('');
  const [selected, setSelected] = React.useState(projects[0]);
  const summary = React.useMemo(() => ({
    total: projects.length,
    atRisk: projects.filter(p => p.status === 'At Risk' || p.health === 'Amber').length,
    openTasks: projects.reduce((a,p)=>a+p.openTasks,0),
    overdue: projects.reduce((a,p)=>a+p.overdue,0),
    criticalRisks: projects.reduce((a,p)=>a+p.criticalRisks,0),
    avgProgress: Math.round(projects.reduce((a,p)=>a+p.progress,0)/projects.length),
  }), []);
  const filtered = projects.filter(p => `${p.name} ${p.owner} ${p.status}`.toLowerCase().includes(query.toLowerCase()));

  return <div className="app">
    <aside className="sidebar">
      <div className="brand"><div className="logo"><Gauge size={23}/></div><div><h1>{CONFIG.portalName}</h1><p>{CONFIG.companyName}</p></div></div>
      <nav>{[[Home,'Dashboard'],[FolderOpen,'Projects'],[KanbanSquare,'Tasks Board'],[AlertTriangle,'Risks & Issues'],[CalendarDays,'Schedule'],[Users,'Resources'],[FileText,'Weekly Reports'],[ShieldCheck,'Governance']].map(([Icon,label])=><a key={label} href={`#${label.replaceAll(' ','-')}`}><Icon size={18}/><span>{label}</span></a>)}</nav>
      <div className="sideCard"><h3>Microsoft Live Links</h3><p>Keep secure data inside Microsoft. This portal opens Lists, Forms, and Teams files using existing permissions.</p><ButtonLink href={CONFIG.microsoftLists.tasks} icon={ClipboardList}>Open Tasks List</ButtonLink><ButtonLink href={CONFIG.forms.addTask} icon={PlusCircle}>Add Task Form</ButtonLink></div>
    </aside>
    <main>
      <motion.section className="hero" initial={{opacity:0,y:12}} animate={{opacity:1,y:0}} transition={{duration:.45}}>
        <div><p className="eyebrow">Live project control center</p><h2>PMO Portfolio Dashboard</h2><p className="intro">External web portal connected by secure links to Microsoft Lists, Forms, Teams files, OneDrive/SharePoint libraries, and weekly reporting workflows.</p></div>
        <div className="actions"><ButtonLink href={CONFIG.forms.addTask} icon={PlusCircle}>Add Task</ButtonLink><ButtonLink href={CONFIG.forms.weeklyReport} icon={Mail}>Submit Weekly Report</ButtonLink><ButtonLink href={CONFIG.forms.registerRisk} icon={AlertTriangle}>Register Risk</ButtonLink></div>
      </motion.section>
      <section className="metrics"><Metric icon={FolderOpen} label="Total Projects" value={summary.total} sub="Microsoft Teams project groups"/><Metric icon={ClipboardList} label="Open Tasks" value={summary.openTasks} sub={`${summary.overdue} overdue tasks`}/><Metric icon={AlertTriangle} label="Critical Risks" value={summary.criticalRisks} sub={`${summary.atRisk} projects at risk`}/><Metric icon={BarChart3} label="Portfolio Progress" value={`${summary.avgProgress}%`} sub="Average project completion"/></section>
      <section className="gridMain">
        <div className="card wide" id="Projects"><div className="cardHead"><div><h3>Projects Portfolio</h3><p>Select a project to open its live Microsoft links.</p></div><div className="search"><Search size={16}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search projects..."/></div></div><div className="tableWrap"><table><thead><tr><th>Project</th><th>Owner</th><th>Status</th><th>Progress</th><th>Files</th></tr></thead><tbody>{filtered.map(p=><tr key={p.id} onClick={()=>setSelected(p)} className={selected.id===p.id?'selected':''}><td><b>{p.name}</b><small>{p.id}</small></td><td>{p.owner}</td><td><Pill>{p.status}</Pill></td><td><div className="progress"><span><i style={{width:`${p.progress}%`}}/></span>{p.progress}%</div></td><td><a href={p.files} target="_blank" rel="noreferrer"><LinkIcon size={14}/> Open</a></td></tr>)}</tbody></table></div></div>
        <div className="card"><div className="cardHead"><div><h3>Project Workspace</h3><p>{selected.name}</p></div><Pill>{selected.health}</Pill></div><div className="workspace"><div className="workTop"><span>Progress</span><b>{selected.progress}%</b></div><div className="bar"><i style={{width:`${selected.progress}%`}}/></div><div className="mini"><div><b>{selected.openTasks}</b><span>Open</span></div><div><b>{selected.overdue}</b><span>Overdue</span></div><div><b>{selected.criticalRisks}</b><span>Risks</span></div></div></div><div className="linkGrid"><ButtonLink href={selected.files} icon={FolderOpen}>Main Files</ButtonLink><ButtonLink href={selected.schedule} icon={CalendarDays}>Project Schedule</ButtonLink><ButtonLink href={selected.reports} icon={FileText}>Status Reports</ButtonLink><ButtonLink href={selected.docs} icon={FileText}>Documentations</ButtonLink></div></div>
      </section>
      <section className="twoCols"><div className="card" id="Tasks-Board"><div className="cardHead"><div><h3>Task Control</h3><p>Live actions open Microsoft Lists or Forms.</p></div><ButtonLink href={CONFIG.microsoftLists.tasks}>Open List</ButtonLink></div><div className="list">{tasks.map(t=><div className="item" key={t.id}><div><b>{t.title}</b><p>{t.project} • {t.assignee} • Due {t.due}</p></div><div><Pill>{t.status}</Pill><Pill>{t.priority}</Pill></div></div>)}</div></div><div className="card" id="Risks-&-Issues"><div className="cardHead"><div><h3>Risks & Issues</h3><p>Critical items requiring PMO follow-up.</p></div><ButtonLink href={CONFIG.forms.registerRisk} icon={PlusCircle}>Register</ButtonLink></div><div className="list">{risks.map(r=><div className="item" key={r.id}><div><b>{r.title}</b><p>{r.project} • Owner: {r.owner}</p></div><Pill>{r.severity}</Pill></div>)}</div></div></section>
      <section className="report" id="Weekly-Reports"><div><h3>Weekly Reporting Center</h3><p>Use Microsoft Forms/List Forms for live input. Export weekly views from Microsoft Lists or send automatic reminders once Power Automate is enabled.</p></div><div className="actions"><a className="primary" href={CONFIG.forms.weeklyReport} target="_blank" rel="noreferrer"><Mail size={16}/> Submit Report</a><a className="darkBtn" href={CONFIG.microsoftLists.reports} target="_blank" rel="noreferrer"><FileText size={16}/> Open Reports List</a></div></section>
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App />);
