// Excel parser (runs in the browser). Finds columns by header text, so column order doesn't matter.
const GRADES = [{min:90,label:'ممتاز'},{min:80,label:'جيد جدا'},{min:60,label:'جيد'},{min:50,label:'مقبول'},{min:0,label:'راسب'}];
const gradeFor = p => GRADES.find(g => p >= g.min).label;
const norm = v => String(v ?? '').replace(/\s+/g, '').trim();
const num = v => { const n = parseFloat(String(v ?? '').replace('%', '')); return isNaN(n) ? null : n; };

function parseExcel(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  const h = rows.findIndex(r => r.some(c => norm(c) === 'الكود' || norm(c) === 'رقمالجلوس'));
  if (h < 0) throw new Error('لم أجد صف العناوين (عمود "الكود").');
  const head = rows[h].map(norm);
  const find = (...k) => head.findIndex(c => c && k.some(x => c.includes(x)));
  const cName = find('اسمالطالب', 'الاسم'), cCode = find('الكود', 'رقمالجلوس');
  const cTotal = find('المجموع'), cPct = find('النسبة'), cGrade = find('التقدير');
  const skip = new Set([cName, cCode, cTotal, cPct, cGrade]);
  const maxRow = rows[h + 1] || [];
  const subjects = head.map((c, i) => (c && c !== 'م' && !skip.has(i))
    ? { col: i, name: String(rows[h][i]).trim(), max: num(maxRow[i]) || 100 } : null).filter(Boolean);
  const maxTotal = subjects.reduce((a, s) => a + s.max, 0);
  const students = [], warnings = [], seen = new Set();
  for (const r of rows.slice(h + 1)) {
    const seat = String(r[cCode] ?? '').trim();
    if (!/^\d{4,}$/.test(seat)) continue;
    const name = String(r[cName]).trim();
    const scores = subjects.map(s => num(r[s.col]));
    const absent = scores.every(x => x === null);
    const total = absent ? 0 : scores.reduce((a, b) => a + (b || 0), 0);
    const percent = absent ? 0 : Math.round(total / maxTotal * 10000) / 100;
    if (seen.has(seat)) warnings.push(`رقم جلوس مكرر: ${seat}`);
    seen.add(seat);
    if (!absent) {
      const ft = num(r[cTotal]), fp = num(r[cPct]);
      if (ft !== null && ft !== total) warnings.push(`${name} (${seat}): المجموع في الملف ${ft} والمحسوب ${total}`);
      if (fp !== null && Math.abs(fp - percent) > 0.1) warnings.push(`${name} (${seat}): النسبة في الملف ${fp} والمحسوبة ${percent}`);
    }
    students.push({
      seat_no: seat, name, absent,
      subjects: subjects.map((s, i) => ({ name: s.name, score: scores[i], max: s.max })),
      total, max_total: maxTotal, percent, grade: absent ? 'غائب' : gradeFor(percent),
    });
  }
  if (!students.length) throw new Error('لم أجد أي طالب في الملف.');
  return { subjects, students, warnings };
}
