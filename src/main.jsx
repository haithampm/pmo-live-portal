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
  Download,
  Edit3,
  FileText,
  Folder,
  Grid2X2,
  Home,
  LineChart,
  MoreHorizontal,
  Plus,
  Search,
  Settings,
  Share2,
  Upload,
  Users,
} from 'lucide-react';
import './styles.css';

const DEFAULT_PROJECTS = [
  { name: 'تطوير بوابة العملاء', manager: 'أحمد العتيبي', avatar: 'https://i.pravatar.cc/60?img=12', status: 'متأخر', progress: 55, end: '2025/06/15', health: 'late' },
  { name: 'نظام ERP', manager: 'سارة المري', avatar: 'https://i.pravatar.cc/60?img=47', status: 'على المسار', progress: 72, end: '2025/07/30', health: 'on' },
  { name: 'تحديث البنية التحتية', manager: 'محمد الحربي', avatar: 'https://i.pravatar.cc/60?img=33', status: 'متعثّر', progress: 30, end: '2025/05/20', health: 'risk' },
  { name: 'تطبيق الهاتف المحمول', manager: 'منال الشريف', avatar: 'https://i.pravatar.cc/60?img=49', status: 'على المسار', progress: 80, end: '2025/08/10', health: 'on' },
  { name: 'تحليل البيانات والذكاء الاصطناعي', manager: 'علي القحطاني', avatar: 'https://i.pravatar.cc/60?img=11', status: 'مكتمل', progress: 100, end: '2025/04/30', health: 'done' },
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

function parseWorkbook(workbook, fileName = 'Uploaded workbook') {
  const sheets = workbook.SheetNames.map(name => ({ name, ...sheetToRows(workbook.Sheets[name]) }));
  const erp = sheets.find(s => s.name === 'ERP Project');
  const projectRows = erp?.rows || [];
  const projects = projectRows.length ? projectRows.slice(0, 24).map((row, i) => ({
    name: row[1] || row[0] || `مشروع ${i + 1}`,
    manager: row[7] || row[6] || 'مدير المشروع',
    avatar: `https://i.pravatar.cc/60?img=${(i % 50) + 1}`,
    status: i % 5 === 0 ? 'متأخر' : i % 4 === 0 ? 'متعثّر' : i % 6 === 0 ? 'مكتمل' : 'على المسار',
    progress: Math.min(100, Math.max(20, 35 + ((i * 11) % 65))),
    end: row[10] || row[9] || '2025/07/30',
    health: i % 5 === 0 ? 'late' : i % 4 === 0 ? 'risk' : i % 6 === 0 ? 'done' : 'on',
  })) : DEFAULT_PROJECTS;
  return { workbook: fileName, sheets, projects };
}

function StatusPill({ type, children }) {
  return <span className={`statusPill ${type}`}>{children}</span>;
}

function TopMetric({ title, value, delta, icon: Icon, tone }) {
  return (
    <motion.div className="topMetric" whileHover={{ y: -4 }} transition={{ duration: 0.2 }}>
      <div className={`metricBadge ${tone}`}><Icon size={22} /></div>
      <p>{title}</p>
      <strong>{value}</strong>
      <span className={delta?.includes('-') ? 'down' : 'up'}>{delta}</span>
    </motion.div>
  );
}

function ProgressBar({ value, tone = 'blue' }) {
  return <div className="progressLine"><i className={tone} style={{ width: `${value}%` }} /></div>;
}

function Donut() {
  return (
    <div className="donutWrap">
      <div className="donut"><span /></div>
      <div className="legend">
        <p><i className="greenDot" />على المسار</p>
        <p><i className="amberDot" />متأخر</p>
        <p><i className="redDot" />متعثر</p>
        <p><i className="grayDot" />مكتمل</p>
      </div>
      <b>إجمالي المشاريع: 24</b>
    </div>
  );
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

function App() {
  const [data, setData] = React.useState({ workbook: 'Sample dashboard', sheets: [], projects: DEFAULT_PROJECTS });
  const [sheetQuery, setSheetQuery] = React.useState('');
  const [activeSheet, setActiveSheet] = React.useState('');
  const currentSheet = data.sheets.find(s => s.name === activeSheet) || data.sheets[0];
  const sheetRows = currentSheet ? currentSheet.rows.filter(row => row.some(cell => norm(cell).toLowerCase().includes(sheetQuery.toLowerCase()))).slice(0, 120) : [];
  const totalRows = currentSheet?.rows?.length || 0;
  const totalCols = currentSheet?.headers?.length || 0;
  const matchedRows = sheetRows.length;

  async function uploadWorkbook(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const parsed = parseWorkbook(workbook, file.name);
    setData(parsed);
    setActiveSheet(parsed.sheets[0]?.name || '');
  }

  const projects = data.projects;

  return (
    <div className="spPage" dir="rtl">
      <header className="spTopBar">
        <div className="topIcons"><img src="https://i.pravatar.cc/48?img=13" /><Settings size={21} /><span>؟</span><Bell size={20} /></div>
        <div className="spSearch"><Search size={19} /><input placeholder="البحث في هذا الموقع" /></div>
        <div className="spBrand"><b>SharePoint</b><Grid2X2 size={24} /></div>
      </header>

      <aside className="rightNav">
        <a className="active"><Home size={22} />الصفحة الرئيسية</a>
        <a><Folder size={21} />المحفظة والمشاريع</a>
        <a><ClipboardCheck size={21} />المهام</a>
        <a><BarChart3 size={21} />التقارير ولوحات المعلومات</a>
        <a><FileText size={21} />المستندات</a>
        <a><CheckCircle2 size={21} />الموافقات</a>
        <a><Users size={21} />الموارد والقدرات</a>
        <a><AlertTriangle size={21} />المخاطر والمشكلات</a>
        <a><BookOpen size={21} />الدروس المستفادة</a>
        <a><Settings size={21} />الإعدادات</a>
        <a className="edit"><Edit3 size={20} />تحرير</a>
        <a className="back"><ChevronLeft size={20} />عودة إلى SharePoint</a>
      </aside>

      <main className="spContent">
        <section className="pageHeader">
          <div className="headerIcon"><BriefcaseBusiness size={44} /></div>
          <div><h1>بوابة إدارة المشاريع</h1><p>لوحة تحكم المحفظة والمشاريع</p></div>
        </section>

        <section className="commandBar">
          <label className="primaryCmd"><Plus size={20} />مشروع جديد<input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadWorkbook} /></label>
          <button><FileText size={18} />إنشاء تقرير</button>
          <button><Download size={18} />تصدير</button>
          <button><Share2 size={18} />مشاركة</button>
          <span className="workbookName">{data.workbook}</span>
        </section>

        <section className="kpiGrid">
          <TopMetric title="عدد المشاريع" value="24" delta="▲ 3 منذ الشهر الماضي" icon={Folder} tone="blue" />
          <TopMetric title="المهام المتأخرة" value="18" delta="▲ 5 منذ الشهر الماضي" icon={Clock3} tone="red" />
          <TopMetric title="المشاريع عالية المخاطر" value="6" delta="▲ 1 منذ الشهر الماضي" icon={AlertTriangle} tone="orange" />
          <TopMetric title="نسبة الإنجاز" value="68%" delta="▲ 6% منذ الشهر الماضي" icon={LineChart} tone="green" />
          <TopMetric title="الاستخدام" value="74%" delta="▼ 2% منذ الشهر الماضي" icon={Users} tone="purple" />
          <TopMetric title="التقارير المفتوحة" value="11" delta="▼ 3 منذ الشهر الماضي" icon={FileText} tone="cyan" />
        </section>

        <section className="dashboardGrid">
          <div className="panel chartPanel"><div className="panelHead"><MoreHorizontal size={22} /><h3>توزيع حالة المشاريع</h3></div><Donut /></div>
          <div className="panel summaryPanel"><div className="panelHead"><BriefcaseBusiness size={22} /><h3>ملخص المحفظة</h3></div><div className="portfolioStats"><div><CheckCircle2 className="greenIcon" /><b>11</b><span>على المسار<br />(45%)</span></div><div><AlertTriangle className="redIcon" /><b>4</b><span>متعثر<br />(15%)</span></div><div><Clock3 className="amberIcon" /><b>6</b><span>متأخر<br />(25%)</span></div><div><CheckCircle2 className="grayIcon" /><b>3</b><span>مكتمل<br />(15%)</span></div></div><div className="dualProgress"><div><p>نسبة الإنجاز الإجمالية <b>68%</b></p><ProgressBar value={68} /></div><div><p>الاستخدام الإجمالي للموارد <b>74%</b></p><ProgressBar value={74} tone="purple" /></div></div></div>
          <div className="panel alertPanel"><div className="panelHead"><Bell size={22} /><h3>تنبيهات ومتابعات</h3></div><div className="alerts"><div><span className="alertIcon red"><AlertTriangle /></span><p><b>تأخر في المشروع "تطوير بوابة العملاء"</b><small>تم تأخير الموعد النهائي 5 أيام</small></p><em>منذ 1 يوم</em></div><div><span className="alertIcon amber"><AlertTriangle /></span><p><b>مخاطر عالية في مشروع "نظام ERP"</b><small>تحديث خطة المخاطر مطلوب</small></p><em>منذ 2 يوم</em></div><div><span className="alertIcon orange"><CalendarDays /></span><p><b>اجتماع لجنة التوجيه</b><small>غداً 10:00 ص</small></p><em>منذ 2 يوم</em></div></div><a className="allAlerts">عرض جميع التنبيهات <ChevronLeft size={17} /></a></div>
        </section>

        <section className="projectsTable panel">
          <div className="tableTop"><a>عرض الكل <ChevronLeft size={17} /></a><h3>المشاريع</h3></div>
          <table><thead><tr><th>المشروع</th><th>المدير</th><th>الحالة</th><th>نسبة الإنجاز</th><th>الموعد النهائي</th><th></th></tr></thead><tbody>{projects.slice(0, 8).map((p) => <tr key={p.name}><td><span className="projectIcon"><Grid2X2 size={16} /></span>{p.name}</td><td><img src={p.avatar} />{p.manager}</td><td><StatusPill type={p.health}>{p.status}</StatusPill></td><td><b>{p.progress}%</b><ProgressBar value={p.progress} /></td><td className={p.health === 'late' || p.health === 'risk' ? 'dateRed' : p.health === 'done' ? 'dateGreen' : ''}>{p.end}</td><td><MoreHorizontal size={20} /></td></tr>)}</tbody></table>
        </section>

        {data.sheets.length > 0 && <section className="sheetViewer panel"><div className="sheetHero"><div><h3>عرض الشيتات مباشرة</h3><p>جداول احترافية مع ألوان للحالة، الأولوية، التواريخ، ونسب الإنجاز.</p></div><div className="sheetMiniStats"><span><b>{data.sheets.length}</b> شيت</span><span><b>{totalRows}</b> صف</span><span><b>{totalCols}</b> عمود</span><span><b>{matchedRows}</b> نتيجة</span></div></div><div className="sheetControls enhanced"><div className="tabs">{data.sheets.map(s => <button key={s.name} className={(currentSheet?.name === s.name) ? 'active' : ''} onClick={() => setActiveSheet(s.name)}>{s.name}<span>{s.rows.length}</span></button>)}</div><div className="sheetSearch"><Search size={17} /><input value={sheetQuery} onChange={e => setSheetQuery(e.target.value)} placeholder="بحث داخل الشيت" /></div></div>{currentSheet && <div className="sheetTable enhancedTable"><table><thead><tr>{currentSheet.headers.map((h, index) => <th key={`${h}-${index}`}><span>{h}</span></th>)}</tr></thead><tbody>{sheetRows.map((row, i) => <tr key={i}>{currentSheet.headers.map((h, j) => <td className={cellTone(row[j], h)} key={`${i}-${j}`}>{renderSheetCell(row[j], h)}</td>)}</tr>)}</tbody></table></div>}</section>}
      </main>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);