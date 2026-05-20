import React from 'react';
import { createRoot } from 'react-dom/client';
import { motion } from 'framer-motion';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  BarChart3,
  Bell,
  BookOpen,
  BriefcaseBusiness,
  CalendarDays,
  CheckCircle2,
  ChevronLeft,
  ClipboardCheck,
  Clock3,
  Copy,
  Download,
  Edit3,
  Eye,
  FileText,
  Filter,
  Folder,
  Grid2X2,
  Home,
  KanbanSquare,
  LayoutList,
  LineChart,
  MoreHorizontal,
  Plus,
  Search,
  Settings,
  Share2,
  Table2,
  Trash2,
  Upload,
  Users,
} from 'lucide-react';
import './styles.css';

const DEFAULT_PROJECTS = [
  { name: 'تطوير بوابة العملاء', manager: 'أحمد العتيبي', avatar: 'https://i.pravatar.cc/60?img=12', status: 'متأخر', progress: 55, end: '2025/06/15', health: 'late', priority: 'High' },
  { name: 'نظام ERP', manager: 'سارة المري', avatar: 'https://i.pravatar.cc/60?img=47', status: 'على المسار', progress: 72, end: '2025/07/30', health: 'on', priority: 'Medium' },
  { name: 'تحديث البنية التحتية', manager: 'محمد الحربي', avatar: 'https://i.pravatar.cc/60?img=33', status: 'متعثّر', progress: 30, end: '2025/05/20', health: 'risk', priority: 'Critical' },
  { name: 'تطبيق الهاتف المحمول', manager: 'منال الشريف', avatar: 'https://i.pravatar.cc/60?img=49', status: 'على المسار', progress: 80, end: '2025/08/10', health: 'on', priority: 'Medium' },
  { name: 'تحليل البيانات والذكاء الاصطناعي', manager: 'علي القحطاني', avatar: 'https://i.pravatar.cc/60?img=11', status: 'مكتمل', progress: 100, end: '2025/04/30', health: 'done', priority: 'Low' },
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

function phaseProgress(text) {
  const value = norm(text).toLowerCase();
  if (value.includes('تم') || value.includes('مكتمل') || value.includes('completed') || value.includes('support')) return 100;
  if (value.includes('go') || value.includes('اطلاق')) return 85;
  if (value.includes('deploy') || value.includes('انتاج')) return 75;
  if (value.includes('develop') || value.includes('تطوير')) return 60;
  if (value.includes('analysis') || value.includes('تحليل')) return 35;
  if (value.includes('prepare') || value.includes('planning') || value.includes('تخطيط')) return 15;
  return 50;
}

function projectHealth(progress) {
  if (progress >= 90) return 'done';
  if (progress >= 70) return 'on';
  if (progress >= 45) return 'late';
  return 'risk';
}

function healthStatus(health) {
  return health === 'done' ? 'مكتمل' : health === 'on' ? 'على المسار' : health === 'late' ? 'متأخر' : 'متعثّر';
}

function priorityFromHealth(health) {
  return health === 'risk' ? 'Critical' : health === 'late' ? 'High' : health === 'on' ? 'Medium' : 'Low';
}

function parseWorkbook(workbook, fileName = 'Uploaded workbook') {
  const sheets = workbook.SheetNames.map(name => ({ name, ...sheetToRows(workbook.Sheets[name]) }));
  const erp = sheets.find(s => s.name === 'ERP Project');
  const serviceSheets = sheets.filter(s => !['ERP Project', 'Radar', 'LOV'].includes(s.name));
  const projectRows = erp?.rows || [];
  const tasks = [];

  serviceSheets.forEach(sheet => {
    sheet.rows.forEach((row, idx) => {
      const title = row[1] || row[0] || '';
      const phase = row[2] || '';
      const status = row[3] || row[4] || '';
      const assignee = row[5] || row[6] || 'غير محدد';
      const note = row[7] || row[8] || '';
      if (!title || title === sheet.name) return;
      const progress = phaseProgress(`${phase} ${status}`);
      const health = projectHealth(progress);
      tasks.push({ id: `${sheet.name}-${idx}`, project: sheet.name, title, phase, status: status || 'مفتوح', assignee, note, progress, health, priority: priorityFromHealth(health) });
    });
  });

  const grouped = tasks.reduce((acc, t) => {
    acc[t.project] = acc[t.project] || [];
    acc[t.project].push(t);
    return acc;
  }, {});

  const projects = projectRows.length ? projectRows.slice(0, 40).map((row, i) => {
    const name = row[1] || row[0] || `مشروع ${i + 1}`;
    const manager = row[7] || row[6] || 'مدير المشروع';
    const projectTasks = grouped[name] || [];
    const progress = projectTasks.length ? Math.round(projectTasks.reduce((a, t) => a + t.progress, 0) / projectTasks.length) : Math.min(100, Math.max(20, 35 + ((i * 11) % 65)));
    const health = projectHealth(progress);
    return { name, manager, avatar: `https://i.pravatar.cc/60?img=${(i % 50) + 1}`, status: healthStatus(health), progress, end: row[10] || row[9] || '2025/07/30', health, tasks: projectTasks.length, priority: priorityFromHealth(health) };
  }) : DEFAULT_PROJECTS;

  return { workbook: fileName, sheets, projects, tasks };
}

function StatusPill({ type, children }) {
  return <span className={`statusPill ${type}`}>{children}</span>;
}

function PriorityPill({ value }) {
  return <span className={`priorityPill ${String(value).toLowerCase()}`}>{value}</span>;
}

function TopMetric({ title, value, delta, icon: Icon, tone }) {
  return <motion.div className="topMetric" whileHover={{ y: -4 }} transition={{ duration: 0.2 }}><div className={`metricBadge ${tone}`}><Icon size={22} /></div><p>{title}</p><strong>{value}</strong><span className={delta?.includes('▼') ? 'down' : 'up'}>{delta}</span></motion.div>;
}

function ProgressBar({ value, tone = 'blue' }) {
  return <div className="progressLine"><i className={tone} style={{ width: `${Math.max(0, Math.min(100, value || 0))}%` }} /></div>;
}

function cellTone(value, header = '') {
  const text = norm(value).toLowerCase();
  const head = norm(header).toLowerCase();
  if (!text) return '';
  if (text.includes('مكتمل') || text.includes('تم') || text.includes('done') || text.includes('completed') || text.includes('على المسار')) return 'cellSuccess';
  if (text.includes('متأخر') || text.includes('تعثر') || text.includes('متعثر') || text.includes('critical') || text.includes('high') || text.includes('blocked')) return 'cellDanger';
  if (text.includes('انتظار') || text.includes('تحت') || text.includes('قيد') || text.includes('تحليل') || text.includes('medium')) return 'cellWarning';
  if (head.includes('date') || head.includes('تاريخ') || text.match(/^\d{4}[/-]\d{1,2}[/-]\d{1,2}/)) return 'cellDate';
  if (text.includes('%') || head.includes('progress') || head.includes('نسبة')) return 'cellProgress';
  return '';
}

function renderSheetCell(value, header) {
  const text = norm(value);
  if (!text) return <span className="mutedCell">—</span>;
  const tone = cellTone(text, header);
  const numeric = Number(text.replace('%', ''));
  if ((tone === 'cellProgress' || norm(header).toLowerCase().includes('progress') || norm(header).includes('نسبة')) && !Number.isNaN(numeric)) {
    const valueNum = Math.max(0, Math.min(100, numeric));
    return <div className="sheetProgress"><b>{valueNum}%</b><span><i style={{ width: `${valueNum}%` }} /></span></div>;
  }
  if (tone === 'cellSuccess' || tone === 'cellDanger' || tone === 'cellWarning') return <span className={`sheetChip ${tone}`}>{text}</span>;
  if (text.startsWith('http')) return <a className="sheetLink" href={text} target="_blank" rel="noreferrer">فتح الرابط</a>;
  return text;
}

function Donut({ projects }) {
  const total = projects.length || 1;
  const on = projects.filter(p => p.health === 'on').length;
  const late = projects.filter(p => p.health === 'late').length;
  const risk = projects.filter(p => p.health === 'risk').length;
  const done = projects.filter(p => p.health === 'done').length;
  return <div className="donutWrap"><div className="donut dynamic" style={{ background: `conic-gradient(#448361 0 ${(on / total) * 100}%, #cb912f ${(on / total) * 100}% ${((on + late) / total) * 100}%, #d44c47 ${((on + late) / total) * 100}% ${((on + late + risk) / total) * 100}%, #9b9a97 ${((on + late + risk) / total) * 100}% 100%)` }}><span /></div><div className="legend"><p><i className="greenDot" />على المسار: {on}</p><p><i className="amberDot" />متأخر: {late}</p><p><i className="redDot" />متعثر: {risk}</p><p><i className="grayDot" />مكتمل: {done}</p></div><b>إجمالي المشاريع: {projects.length}</b></div>;
}

function ActionButton({ icon: Icon, label, tone = '' }) {
  return <button className={`actionButton ${tone}`}><Icon size={16}/>{label}</button>;
}

function BoardCard({ item }) {
  return <div className="boardCard"><div className="boardCardTop"><b>{item.name || item.title}</b><MoreHorizontal size={16}/></div><p>{item.manager || item.assignee || 'غير محدد'}</p><ProgressBar value={item.progress || 0}/><div className="boardMeta"><PriorityPill value={item.priority || 'Medium'} /><span>{item.progress || 0}%</span></div></div>;
}

function MondayBoard({ projects, tasks }) {
  const columns = [
    ['risk', 'متعثر / Critical', projects.filter(p => p.health === 'risk')],
    ['late', 'متأخر / High', projects.filter(p => p.health === 'late')],
    ['on', 'على المسار', projects.filter(p => p.health === 'on')],
    ['done', 'مكتمل', projects.filter(p => p.health === 'done')],
  ];
  return <div className="mondayBoard">{columns.map(([key, title, items]) => <div className={`boardColumn ${key}`} key={key}><div className="boardColumnHead"><span>{title}</span><b>{items.length}</b></div>{items.slice(0, 8).map(item => <BoardCard key={item.name} item={item}/>)}{items.length === 0 && <div className="emptyColumn">لا توجد عناصر</div>}</div>)}</div>;
}

function FunctionPanel({ setPage }) {
  const actions = [
    [Plus, 'إضافة عنصر', 'primary'], [Filter, 'فلترة متقدمة', ''], [LayoutList, 'عرض جدول', ''], [KanbanSquare, 'عرض كانبان', ''], [CalendarDays, 'تقويم', ''], [BarChart3, 'تحليل', ''], [Copy, 'نسخ', ''], [Eye, 'معاينة', ''], [Edit3, 'تعديل', ''], [Trash2, 'حذف', 'danger'], [Download, 'تصدير', ''], [Share2, 'مشاركة', '']
  ];
  return <section className="functionPanel panel"><div className="functionHead"><div><h3>أدوات التشغيل</h3><p>Buttons & functions بشكل قريب من Notion و Monday.</p></div><button onClick={()=>setPage('sheets')}>فتح القوائم</button></div><div className="functionGrid">{actions.map(([Icon, label, tone]) => <ActionButton key={label} icon={Icon} label={label} tone={tone}/>)}</div></section>;
}

function App() {
  const [data, setData] = React.useState({ workbook: 'Sample dashboard', sheets: [], projects: DEFAULT_PROJECTS, tasks: [] });
  const [page, setPage] = React.useState('dashboard');
  const [sheetQuery, setSheetQuery] = React.useState('');
  const [activeSheet, setActiveSheet] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState('الكل');
  const [managerFilter, setManagerFilter] = React.useState('الكل');
  const [progressFilter, setProgressFilter] = React.useState('الكل');
  const [viewMode, setViewMode] = React.useState('table');
  const currentSheet = data.sheets.find(s => s.name === activeSheet) || data.sheets[0];
  const projects = data.projects;

  const managers = ['الكل', ...Array.from(new Set(projects.map(p => p.manager).filter(Boolean))).slice(0, 30)];
  const filteredProjects = projects.filter(p => {
    const statusOk = statusFilter === 'الكل' || p.status === statusFilter;
    const managerOk = managerFilter === 'الكل' || p.manager === managerFilter;
    const progressOk = progressFilter === 'الكل' || (progressFilter === 'أقل من 50%' ? p.progress < 50 : progressFilter === '50% - 80%' ? p.progress >= 50 && p.progress <= 80 : p.progress > 80);
    return statusOk && managerOk && progressOk;
  });
  const sheetRows = currentSheet ? currentSheet.rows.filter(row => row.some(cell => norm(cell).toLowerCase().includes(sheetQuery.toLowerCase()))).slice(0, 120) : [];
  const totalRows = currentSheet?.rows?.length || 0;
  const totalCols = currentSheet?.headers?.length || 0;

  const avgProgress = filteredProjects.length ? Math.round(filteredProjects.reduce((a, p) => a + p.progress, 0) / filteredProjects.length) : 0;
  const onCount = filteredProjects.filter(p => p.health === 'on').length;
  const riskCount = filteredProjects.filter(p => p.health === 'risk').length;
  const lateCount = filteredProjects.filter(p => p.health === 'late').length;
  const doneCount = filteredProjects.filter(p => p.health === 'done').length;

  async function uploadWorkbook(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const parsed = parseWorkbook(workbook, file.name);
    setData(parsed);
    setActiveSheet(parsed.sheets[0]?.name || '');
    setPage('sheets');
  }

  const nav = [
    ['dashboard', Home, 'الصفحة الرئيسية'],
    ['portfolio', Folder, 'المحفظة والمشاريع'],
    ['board', KanbanSquare, 'لوحة Monday'],
    ['functions', Settings, 'الأزرار والفانكشن'],
    ['sheets', Table2, 'قوائم الشيتات'],
    ['tasks', ClipboardCheck, 'المهام'],
    ['reports', BarChart3, 'التقارير ولوحات المعلومات'],
    ['documents', FileText, 'المستندات'],
    ['approvals', CheckCircle2, 'الموافقات'],
    ['resources', Users, 'الموارد والقدرات'],
    ['risks', AlertTriangle, 'المخاطر والمشكلات'],
    ['lessons', BookOpen, 'الدروس المستفادة'],
  ];

  return <div className="spPage" dir="rtl">
    <header className="spTopBar"><div className="topIcons"><img src="https://i.pravatar.cc/48?img=13" /><Settings size={21} /><span>؟</span><Bell size={20} /></div><div className="spSearch"><Search size={19} /><input placeholder="البحث في هذا الموقع" /></div><div className="spBrand"><b>PMO Workspace</b><Grid2X2 size={24} /></div></header>
    <aside className="rightNav">{nav.map(([key, Icon, label]) => <a key={key} className={page === key ? 'active' : ''} onClick={() => setPage(key)}><Icon size={22} />{label}</a>)}<a className="edit"><Edit3 size={20} />تحرير</a><a className="back"><ChevronLeft size={20} />عودة إلى SharePoint</a></aside>

    <main className="spContent">
      <section className="pageHeader"><div className="headerIcon"><BriefcaseBusiness size={44} /></div><div><h1>بوابة إدارة المشاريع</h1><p>{page === 'dashboard' ? 'Dashboard ديناميكي بأسلوب Notion / Monday' : page === 'sheets' ? 'قوائم وDatabases للشيتات' : page === 'board' ? 'لوحة كانبان احترافية مثل Monday' : page === 'functions' ? 'مركز الأزرار والفانكشن' : 'صفحة مستقلة لإدارة بيانات المحفظة'}</p></div></section>
      <section className="commandBar"><label className="primaryCmd"><Plus size={20} />رفع شيت جديد<input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadWorkbook} /></label><button onClick={()=>setPage('functions')}><Settings size={18} />الفانكشن</button><button onClick={()=>setPage('board')}><KanbanSquare size={18} />Monday Board</button><button><Download size={18} />تصدير</button><button><Share2 size={18} />مشاركة</button><span className="workbookName">{data.workbook}</span></section>

      {page === 'dashboard' && <>
        <FunctionPanel setPage={setPage}/>
        <section className="filterBar panel"><div><Filter size={18}/><b>فلاتر الداشبورد</b></div><select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}><option>الكل</option><option>على المسار</option><option>متأخر</option><option>متعثّر</option><option>مكتمل</option></select><select value={managerFilter} onChange={e=>setManagerFilter(e.target.value)}>{managers.map(m => <option key={m}>{m}</option>)}</select><select value={progressFilter} onChange={e=>setProgressFilter(e.target.value)}><option>الكل</option><option>أقل من 50%</option><option>50% - 80%</option><option>أكثر من 80%</option></select><div className="viewToggle"><button className={viewMode==='table'?'active':''} onClick={()=>setViewMode('table')}><LayoutList size={16}/>جدول</button><button className={viewMode==='board'?'active':''} onClick={()=>setViewMode('board')}><KanbanSquare size={16}/>كانبان</button></div></section>
        <section className="kpiGrid"><TopMetric title="عدد المشاريع" value={filteredProjects.length} delta="حسب الفلاتر الحالية" icon={Folder} tone="blue" /><TopMetric title="المشاريع المتأخرة" value={lateCount} delta="تحتاج متابعة" icon={Clock3} tone="red" /><TopMetric title="المشاريع عالية المخاطر" value={riskCount} delta="تحتاج تصعيد" icon={AlertTriangle} tone="orange" /><TopMetric title="نسبة الإنجاز" value={`${avgProgress}%`} delta="متوسط المحفظة" icon={LineChart} tone="green" /><TopMetric title="على المسار" value={onCount} delta="حالة مستقرة" icon={Users} tone="purple" /><TopMetric title="مكتملة" value={doneCount} delta="تم الانتهاء" icon={FileText} tone="cyan" /></section>
        {viewMode === 'board' ? <MondayBoard projects={filteredProjects} tasks={data.tasks}/> : <section className="dashboardGrid"><div className="panel chartPanel"><div className="panelHead"><MoreHorizontal size={22} /><h3>توزيع حالة المشاريع</h3></div><Donut projects={filteredProjects} /></div><div className="panel summaryPanel"><div className="panelHead"><BriefcaseBusiness size={22} /><h3>ملخص المحفظة الديناميكي</h3></div><div className="portfolioStats"><div><CheckCircle2 className="greenIcon" /><b>{onCount}</b><span>على المسار</span></div><div><AlertTriangle className="redIcon" /><b>{riskCount}</b><span>متعثر</span></div><div><Clock3 className="amberIcon" /><b>{lateCount}</b><span>متأخر</span></div><div><CheckCircle2 className="grayIcon" /><b>{doneCount}</b><span>مكتمل</span></div></div><div className="dualProgress"><div><p>نسبة الإنجاز الإجمالية <b>{avgProgress}%</b></p><ProgressBar value={avgProgress} /></div><div><p>عدد المهام المستخرجة <b>{data.tasks?.length || 0}</b></p><ProgressBar value={Math.min(100, (data.tasks?.length || 0) / 5)} tone="purple" /></div></div></div><div className="panel alertPanel"><div className="panelHead"><Bell size={22} /><h3>تنبيهات ومتابعات</h3></div><div className="alerts"><div><span className="alertIcon red"><AlertTriangle /></span><p><b>{riskCount} مشروع متعثر</b><small>راجع صفحة المحفظة والمشاريع</small></p><em>الآن</em></div><div><span className="alertIcon amber"><AlertTriangle /></span><p><b>{lateCount} مشروع متأخر</b><small>تحتاج خطة معالجة</small></p><em>الآن</em></div><div><span className="alertIcon orange"><CalendarDays /></span><p><b>{data.sheets.length} شيت متاح</b><small>افتح صفحة قوائم الشيتات</small></p><em>الآن</em></div></div><a className="allAlerts" onClick={()=>setPage('sheets')}>عرض الشيتات <ChevronLeft size={17} /></a></div></section>}
      </>}

      {page === 'portfolio' && <section className="projectsTable panel"><div className="tableTop"><a>عرض الكل <ChevronLeft size={17} /></a><h3>المشاريع</h3></div><table><thead><tr><th>المشروع</th><th>المدير</th><th>الحالة</th><th>الأولوية</th><th>نسبة الإنجاز</th><th>الموعد النهائي</th><th></th></tr></thead><tbody>{filteredProjects.map(p => <tr key={p.name}><td><span className="projectIcon"><Grid2X2 size={16} /></span>{p.name}</td><td><img src={p.avatar} />{p.manager}</td><td><StatusPill type={p.health}>{p.status}</StatusPill></td><td><PriorityPill value={p.priority}/></td><td><b>{p.progress}%</b><ProgressBar value={p.progress} /></td><td className={p.health === 'late' || p.health === 'risk' ? 'dateRed' : p.health === 'done' ? 'dateGreen' : ''}>{p.end}</td><td><MoreHorizontal size={20} /></td></tr>)}</tbody></table></section>}
      {page === 'board' && <MondayBoard projects={filteredProjects} tasks={data.tasks}/>}      
      {page === 'functions' && <FunctionPanel setPage={setPage}/>}      
      {page === 'sheets' && <section className="sheetViewer panel"><div className="sheetHero"><div><h3>Databases & Lists</h3><p>قوائم منظمة بأسلوب Notion مع فلاتر وTags وViews.</p></div><div className="sheetMiniStats"><span><b>{data.sheets.length}</b> شيت</span><span><b>{totalRows}</b> صف</span><span><b>{totalCols}</b> عمود</span><span><b>{sheetRows.length}</b> نتيجة</span></div></div><div className="sheetControls enhanced"><div className="selectBlock"><label>اختر الشيت</label><select value={currentSheet?.name || ''} onChange={e => setActiveSheet(e.target.value)}>{data.sheets.length ? data.sheets.map(s => <option key={s.name} value={s.name}>{s.name} - {s.rows.length} صف</option>) : <option>ارفع ملف Excel أولاً</option>}</select></div><div className="selectBlock"><label>نوع العرض</label><select><option>جدول تفصيلي</option><option>Board / Kanban</option><option>Calendar</option><option>Timeline</option><option>عناصر تحتاج متابعة</option></select></div><div className="selectBlock"><label>حالة العناصر</label><select><option>كل الحالات</option><option>Open</option><option>In Progress</option><option>Blocked</option><option>Completed</option></select></div><div className="sheetSearch"><Search size={17} /><input value={sheetQuery} onChange={e => setSheetQuery(e.target.value)} placeholder="بحث داخل الشيت" /></div></div><div className="tabs compactTabs">{data.sheets.map(s => <button key={s.name} className={(currentSheet?.name === s.name) ? 'active' : ''} onClick={() => setActiveSheet(s.name)}>{s.name}<span>{s.rows.length}</span></button>)}</div>{currentSheet ? <div className="sheetTable enhancedTable"><table><thead><tr>{currentSheet.headers.map((h, index) => <th key={`${h}-${index}`}><span>{h}</span></th>)}</tr></thead><tbody>{sheetRows.map((row, i) => <tr key={i}>{currentSheet.headers.map((h, j) => <td className={cellTone(row[j], h)} key={`${i}-${j}`}>{renderSheetCell(row[j], h)}</td>)}</tr>)}</tbody></table></div> : <div className="emptySheet"><Upload size={36}/><b>ارفع ملف Excel لعرض القوائم</b><p>استخدم زر رفع شيت جديد بالأعلى.</p></div>}</section>}
      {!['dashboard','portfolio','sheets','board','functions'].includes(page) && <section className="panel placeholderPage"><h3>صفحة {nav.find(n => n[0] === page)?.[2]}</h3><p>سيتم ربط هذه الصفحة لاحقًا بقوائم Microsoft Lists أو شيتات Excel حسب اختيارك.</p></section>}
    </main>
  </div>;
}

createRoot(document.getElementById('root')).render(<App />);