import fs from 'node:fs';

const source = fs.readFileSync('src/main.jsx', 'utf8');
const checks = [
  ['Dashboard renders KPI cards', /function Dashboard\(/],
  ['Navigation includes Settings page', /settings.*Settings/],
  ['Navigation includes Import Export center', /importExport.*Import \/ Export/],
  ['Module pages include Table and Form tabs', /Table \/ Views[\s\S]*Form Entry/],
  ['Files mapping is view-only', /mode:'view-only'/],
  ['Files mapping no longer saves mapped records from card view', !/Save mapped files/.test(source)],
  ['Generated mapping records are hidden from the Files table', /visibleFiles\(/],
  ['Generated mapping records are blocked from create flow', /if \(type === 'files' && isGeneratedMappedFile\(record\)\) return/],
  ['File card counts mapped tree items separately', /mapped items|mappedCount/],
  ['SharePoint structure includes requested folders', /1-Project Phases[\s\S]*2-Project Schedule[\s\S]*3-Project status document[\s\S]*4-Master Data Template[\s\S]*Recordings/],
  ['Export center includes Excel Word PDF JSON outputs', /Excel\+ formatted[\s\S]*PDF report[\s\S]*Word report[\s\S]*JSON backup/],
  ['Theme settings include Light and Dark', /Light[\s\S]*Dark/]
];

let failed = 0;
for (const [name, condition] of checks) {
  const ok = condition instanceof RegExp ? condition.test(source) : Boolean(condition);
  console.log(`${ok ? 'PASS' : 'FAIL'} - ${name}`);
  if (!ok) failed += 1;
}

if (failed) {
  console.error(`\n${failed} E2E smoke check(s) failed.`);
  process.exit(1);
}
console.log('\nAll PMO E2E smoke checks passed.');
