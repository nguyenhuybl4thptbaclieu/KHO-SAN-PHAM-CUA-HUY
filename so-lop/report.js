// Dựng nội dung cho Excel và bản in từ dữ liệu.
import { feeTable, monthSummary, paymentsIn, currentSegment, WEEKDAY, WEEKDAY_SHORT, fmt, fmtShort, money, today } from './core.js';

const STATUS = { done: 'Đã dạy', off: 'Nghỉ', planned: 'Dự kiến', today: 'Hôm nay' };
export const groupLine = (g, short = false) => {
  const seg = currentSegment(g);
  if (short) return seg.days.map(d => WEEKDAY_SHORT[d]).join(' · ') + (g.time.start ? ` · ${g.time.start}${g.time.end ? '–' + g.time.end : ''}` : '');
  const t = g.time.start ? ` · ${g.time.start}${g.time.end ? '–' + g.time.end : ''}` : '';
  return seg.days.map(d => WEEKDAY[d]).join(', ') + t;
};
const cellText = c => c.state === 'paid' ? `${fmt(c.payment.date)} · ${c.payment.method === 'CK' ? 'CK' : 'Tiền mặt'}` : c.state === 'overdue' ? 'Chưa đóng' : c.state === 'na' ? 'Chưa vào học' : '';
const cellStyle = c => ({ paid: 'paid', overdue: 'due', na: 'open', open: 'normal' }[c.state]);

export function feeSheet(g, holidays, name) {
  const t = feeTable(g, holidays), n = t.months.length;
  const info = `${groupLine(g)}${g.fee ? ' · Học phí ' + money(g.fee) + '/tháng' : ''} · Xuất ngày ${fmt(today())}`;
  const rows = [
    [{ v: `Học phí · ${g.name}`, s: 'title' }],
    [{ v: info, s: 'muted' }],
    [],
    ['TT', 'Họ và tên', 'Ngày vào học', ...t.months.map(m => `Tháng ${m.number}\n${fmtShort(m.first)}–${fmtShort(m.last)}`)].map(v => ({ v, s: 'head' })),
    ...t.rows.map(r => [r.no, { v: r.student.name, s: 'bold' }, fmt(r.student.joined), ...r.cells.map(c => ({ v: cellText(c), s: cellStyle(c) }))]),
    [],
  ];
  const sums = t.months.map(m => monthSummary(g, m, t.cal.current));
  rows.push([null, { v: 'Đã đóng', s: 'bold' }, null, ...sums.map(s => `${s.paid}/${s.students}`)]);
  rows.push([null, 'Chuyển khoản', null, ...sums.map(s => s.CK)]);
  rows.push([null, 'Tiền mặt', null, ...sums.map(s => s.TM)]);
  if (g.fee) {
    rows.push([null, 'Tiền CK (đ)', null, ...sums.map(s => ({ v: s.amountCK, s: 'money' }))]);
    rows.push([null, 'Tiền mặt (đ)', null, ...sums.map(s => ({ v: s.amountTM, s: 'money' }))]);
  }
  return { name: name || g.name, cols: [5, 26, 13, ...Array(n).fill(19)], rows, freeze: { col: 2, row: 4 }, merges: [`A1:H1`, `A2:H2`], titleRow: true };
}

export function scheduleSheet(g, holidays, cal) {
  const rows = [[{ v: `Lịch dạy · ${g.name}`, s: 'title' }], [{ v: groupLine(g), s: 'muted' }], [],
    ['Tháng', 'Buổi', 'Ngày', 'Thứ', 'Trạng thái', 'Ghi chú'].map(v => ({ v, s: 'head' }))];
  for (const r of cal.rows.filter(r => r.month <= cal.current + 1))
    rows.push([r.month, r.status === 'off' ? '' : r.slot, fmt(r.date), WEEKDAY[new Date(r.date + 'T12:00:00Z').getUTCDay()], STATUS[r.status], [r.extra ? 'Học bù' : '', r.holiday ? 'Nghỉ lễ: ' + r.holiday : ''].filter(Boolean).join(' · ')]);
  return { name: 'Lịch ' + g.name, cols: [8, 7, 13, 11, 12, 26], rows, freeze: { row: 4 }, merges: ['A1:F1', 'A2:F2'], titleRow: true };
}

export function paymentsSheet(data, ym, label) {
  const list = paymentsIn(data, ym);
  const ck = list.filter(p => p.method === 'CK'), tm = list.filter(p => p.method === 'TM');
  const sum = a => a.reduce((s, p) => s + p.amount, 0);
  const rows = [[{ v: `Thống kê học phí · ${label}`, s: 'title' }], [],
    [{ v: 'Chuyển khoản', s: 'bold' }, `${ck.length} lượt`, { v: sum(ck), s: 'money' }],
    [{ v: 'Tiền mặt', s: 'bold' }, `${tm.length} lượt`, { v: sum(tm), s: 'money' }],
    [{ v: 'Tổng', s: 'bold' }, `${list.length} lượt`, { v: sum(list), s: 'money' }], [],
    ['Ngày đóng', 'Học sinh', 'Nhóm', 'Tháng học', 'Hình thức', 'Số tiền (đ)'].map(v => ({ v, s: 'head' })),
    ...list.map(p => [fmt(p.date), p.student.name, p.group.name, `Tháng ${p.month}`, p.method === 'CK' ? 'Chuyển khoản' : 'Tiền mặt', p.amount ? { v: p.amount, s: 'money' } : ''])];
  return { name: 'Thống kê', cols: [14, 26, 22, 12, 15, 14], rows, freeze: { row: 7 }, merges: ['A1:F1'], titleRow: true };
}

const h = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export function printHtml(groups, holidays) {
  return groups.map(g => {
    const t = feeTable(g, holidays);
    const sums = t.months.map(m => monthSummary(g, m, t.cal.current));
    return `<section class="print-page"><h1>Học phí · ${h(g.name)}</h1><p>${h(groupLine(g))}${g.fee ? ' · ' + money(g.fee) + '/tháng' : ''} · In ngày ${fmt(today())}</p>
<table><thead><tr><th>TT</th><th>Họ và tên</th><th>Vào học</th>${t.months.map(m => `<th>T${m.number}<br><small>${fmtShort(m.first)}–${fmtShort(m.last)}</small></th>`).join('')}</tr></thead>
<tbody>${t.rows.map(r => `<tr><td>${r.no}</td><td class="n">${h(r.student.name)}</td><td>${fmtShort(r.student.joined)}</td>${r.cells.map(c => `<td class="${c.state}">${c.state === 'paid' ? `${fmtShort(c.payment.date)} ${c.payment.method === 'CK' ? 'CK' : 'TM'}` : c.state === 'overdue' ? 'Chưa đóng' : c.state === 'na' ? '–' : ''}</td>`).join('')}</tr>`).join('')}
<tr class="sum"><td></td><td>Đã đóng</td><td></td>${sums.map(s => `<td>${s.paid}/${s.students}<br><small>CK ${s.CK} · TM ${s.TM}</small></td>`).join('')}</tr></tbody></table></section>`;
  }).join('');
}
