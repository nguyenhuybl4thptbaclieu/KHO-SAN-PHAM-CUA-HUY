// Sổ lớp · logic thuần (không phụ thuộc giao diện). Mọi ngày dạng 'YYYY-MM-DD'.
export const PER_MONTH = 8;          // số buổi = 1 tháng học phí
export const DAYS_PER_WEEK = 2;      // số buổi mỗi tuần
export const DATA_VERSION = 2;

export const WEEKDAY = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];
export const WEEKDAY_SHORT = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];

export function today(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
export function validDate(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T12:00:00Z');
  return !isNaN(d) && d.toISOString().slice(0, 10) === s && s >= '1900-01-01' && s <= '2200-12-31';
}
export function validTime(s) { return s === '' || (typeof s === 'string' && /^([01]\d|2[0-3]):[0-5]\d$/.test(s)); }
export function plus(s, n = 1) { const d = new Date(s + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
export const weekday = s => new Date(s + 'T12:00:00Z').getUTCDay();
export const fmt = s => s ? `${s.slice(8, 10)}/${s.slice(5, 7)}/${s.slice(0, 4)}` : '';
export const fmtShort = s => s ? `${s.slice(8, 10)}/${s.slice(5, 7)}` : '';
export const money = n => (Number(n) || 0).toLocaleString('vi-VN') + ' đ';

/** Bỏ dấu tiếng Việt để tìm kiếm: "Nguyễn Đức" → "nguyen duc". */
export function fold(s) {
  return String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().trim();
}

export function holidayOn(holidays, date) {
  for (const h of holidays || []) if (h.from <= date && date <= h.to) return h;
  return null;
}

/**
 * Tính lịch dạy của một nhóm.
 * overrides[date]: 'off' = nghỉ · 'done' = xác nhận đã dạy (hôm nay) · 'on' = vẫn dạy dù trùng ngày nghỉ lễ
 * extras: các ngày học bù (ngoài lịch cố định), tính như một buổi bình thường.
 */
export function schedule(group, holidays = [], asOf = today(), minMonths = 12) {
  const segs = [...group.schedules].sort((a, b) => a.from.localeCompare(b.from));
  const extras = new Set(group.extras || []);
  const lastExtra = [...extras].sort().at(-1) || '';
  const rows = []; let count = 0, done = 0, idx = 0;
  const horizon = plus(asOf, 400) > lastExtra ? plus(asOf, 400) : lastExtra;
  const target = Math.max(12, Math.min(240, minMonths)) * PER_MONTH;
  for (let date = group.start, steps = 0; steps < 90000; date = plus(date), steps++) {
    while (idx + 1 < segs.length && segs[idx + 1].from <= date) idx++;
    const scheduled = segs[idx].from <= date && segs[idx].days.includes(weekday(date));
    const extra = !scheduled && extras.has(date);
    if (!scheduled && !extra) continue;
    const ov = group.overrides[date];
    const hol = extra ? null : holidayOn(holidays, date);
    const off = ov === 'off' || (hol && ov !== 'on' && ov !== 'done');
    const status = off ? 'off' : date > asOf ? 'planned' : (ov === 'done' || date < asOf) ? 'done' : 'today';
    const month = Math.floor(count / PER_MONTH) + 1, slot = count % PER_MONTH + 1;
    rows.push({ date, status, month, slot, extra, holiday: hol && off && ov !== 'off' ? hol.name : '' });
    if (status !== 'off') count++;
    if (status === 'done') done++;
    if (count >= target && date >= horizon && count % PER_MONTH === 0) break;
  }
  const months = [];
  for (const row of rows) {
    const m = months[row.month - 1] ?? (months[row.month - 1] = { number: row.month, rows: [], done: 0 });
    m.rows.push(row); if (row.status === 'done') m.done++;
  }
  for (const m of months) {
    const live = m.rows.filter(r => r.status !== 'off');
    m.first = live[0]?.date || m.rows[0].date; m.last = live.at(-1)?.date || m.rows.at(-1).date;
  }
  const current = months.find(m => m.done < PER_MONTH)?.number ?? months.length + 1;
  return { rows, months, current, done };
}

export function currentSegment(group, asOf = today()) {
  return group.schedules.filter(s => s.from <= asOf).at(-1) || group.schedules[0];
}

/** Học sinh có học trong tháng đó không (có buổi không nghỉ từ ngày vào học). */
export function eligible(student, month) { return !!month && month.rows.some(r => r.status !== 'off' && r.date >= student.joined); }

/**
 * Trạng thái học phí của 1 học sinh ở 1 tháng học:
 * 'paid' đã đóng · 'overdue' chưa đóng (tháng đã kết thúc) · 'open' đang học, chưa đóng (không cảnh báo) · 'na' chưa vào học
 */
export function payState(student, month, current) {
  if (student.payments[month.number]) return 'paid';
  if (!eligible(student, month)) return 'na';
  return month.number < current ? 'overdue' : 'open';
}

export function overdueList(group, cal) {
  const out = [];
  for (const s of group.students) for (const m of cal.months) {
    if (m.number >= cal.current) break;
    if (payState(s, m, cal.current) === 'overdue') out.push({ student: s, month: m.number });
  }
  return out;
}

/** Số lượt nghỉ/xác nhận sẽ bị bỏ khi đổi lịch (chỉ những ngày không còn trong lịch mới). */
export function changePreview(g, from, days) {
  return Object.keys(g.overrides).filter(d => d >= from && !(g.extras || []).includes(d) && !days.includes(weekday(d))).length;
}
export function changeSchedule(g, from, days) {
  if (!validDate(from) || from < g.start || days.length !== DAYS_PER_WEEK || new Set(days).size !== DAYS_PER_WEEK || days.some(d => !Number.isInteger(d) || d < 0 || d > 6)) throw Error('Lịch học không hợp lệ.');
  g.schedules = g.schedules.filter(s => s.from < from);
  g.schedules.push({ from, days: [...days].sort() });
  for (const d of Object.keys(g.overrides)) if (d >= from && !(g.extras || []).includes(d) && !days.includes(weekday(d))) delete g.overrides[d];
}
export function isScheduledDay(g, date) {
  const seg = g.schedules.filter(s => s.from <= date).at(-1);
  return !!seg && seg.days.includes(weekday(date));
}
export function addExtra(g, date) {
  if (!validDate(date) || date < g.start) throw Error('Ngày học bù phải từ ngày bắt đầu nhóm trở đi.');
  if (isScheduledDay(g, date)) throw Error('Ngày này đã có trong lịch học. Hãy sửa trạng thái buổi đó.');
  if ((g.extras || []).includes(date)) throw Error('Ngày này đã có buổi học bù.');
  g.extras = [...(g.extras || []), date].sort();
}
export function removeExtra(g, date) { g.extras = (g.extras || []).filter(d => d !== date); delete g.overrides[date]; }

/* ---------- Thống kê ---------- */
export function monthSummary(group, month, current) {
  const r = { students: 0, paid: 0, CK: 0, TM: 0, overdue: 0, open: 0, amountCK: 0, amountTM: 0 };
  for (const s of group.students) {
    const st = payState(s, month, current);
    if (st === 'na') continue;
    r.students++;
    if (st === 'paid') { const p = s.payments[month.number]; r.paid++; r[p.method]++; r['amount' + p.method] += payAmount(group, p); }
    else r[st]++;
  }
  return r;
}
/** Số tiền của một lượt đóng: số đã lưu, nếu không có thì lấy học phí nhóm. */
export function payAmount(group, p) { return Number.isInteger(p.amount) ? p.amount : (group.fee || 0); }
/** Đổi học phí nhóm; các lần đóng cũ giữ nguyên số tiền theo học phí trước đó. */
export function setFee(group, fee) {
  if (group.fee && group.fee !== fee) for (const s of group.students) for (const p of Object.values(s.payments)) if (!Number.isInteger(p.amount)) p.amount = group.fee;
  group.fee = fee;
}
/** Học sinh còn nợ: tháng đã kết thúc chưa đóng + buổi đã dạy của tháng đang học (học phí ÷ 8 mỗi buổi). */
export function debtOf(group, student, cal) {
  const fee = group.fee || 0, months = [];
  let amount = 0, partial = null;
  for (const m of cal.months) {
    if (m.number >= cal.current) break;
    if (payState(student, m, cal.current) === 'overdue') { months.push(m.number); amount += fee; }
  }
  const cm = cal.months[cal.current - 1];
  if (cm && !student.payments[cm.number]) {
    const sessions = cm.rows.filter(r => r.status === 'done' && r.date >= student.joined).length;
    if (sessions) { partial = { month: cm.number, sessions }; amount += sessions * fee / PER_MONTH; }
  }
  return { months, partial, amount, hasFee: !!group.fee };
}
/** Các lượt đóng tiền có ngày đóng trong tháng dương lịch 'YYYY-MM' (hoặc tất cả nếu ym rỗng). */
export function paymentsIn(data, ym) {
  const out = [];
  for (const g of data.groups) for (const s of g.students) for (const [m, p] of Object.entries(s.payments))
    if (!ym || p.date.startsWith(ym)) out.push({ group: g, student: s, month: Number(m), date: p.date, method: p.method, amount: payAmount(g, p) });
  return out.sort((a, b) => b.date.localeCompare(a.date) || a.student.name.localeCompare(b.student.name, 'vi'));
}

/** Bảng học phí dùng chung cho màn hình bảng, Excel và bản in. */
export function feeTable(group, holidays, asOf = today()) {
  const cal = schedule(group, holidays, asOf);
  const maxPaid = Math.max(0, ...group.students.flatMap(s => Object.keys(s.payments).map(Number)));
  const n = Math.max(cal.current, maxPaid, 1);
  const months = cal.months.slice(0, n);
  const rows = group.students.map((s, i) => ({
    no: i + 1, student: s,
    cells: months.map(m => { const st = payState(s, m, cal.current), p = s.payments[m.number]; return { month: m.number, state: st, payment: p || null }; }),
  }));
  return { cal, months, rows };
}

/* ---------- Kiểm tra & chuyển đổi dữ liệu ---------- */
export function emptyData() { return { version: DATA_VERSION, updatedAt: '', holidays: [], groups: [] }; }

/** Chuyển dữ liệu phiên bản 1 (app cũ) sang phiên bản 2. */
export function migrate(data) {
  if (data && data.version === 1 && Array.isArray(data.groups)) {
    return {
      version: 2, updatedAt: '', holidays: [],
      groups: data.groups.map(g => ({ ...g, time: { start: '', end: '' }, fee: null, extras: [] })),
    };
  }
  return data;
}

export function validate(data) {
  try { return validateStrict(data); }
  catch (e) { throw Error(e instanceof TypeError ? 'Tệp không đúng định dạng Sổ lớp.' : e.message); }
}
function validateStrict(data) {
  if (!data || data.version !== DATA_VERSION || !Array.isArray(data.groups) || data.groups.length > 500 || !Array.isArray(data.holidays)) throw Error('Tệp không đúng định dạng Sổ lớp.');
  if (typeof data.updatedAt !== 'string') throw Error('Tệp không đúng định dạng Sổ lớp.');
  if (data.defaultFee !== undefined && data.defaultFee !== null && !(Number.isInteger(data.defaultFee) && data.defaultFee >= 0 && data.defaultFee <= 1e9)) throw Error('Học phí chung không hợp lệ.');
  const ids = new Set();
  const id = x => { if (typeof x !== 'string' || !x.length || x.length > 100 || ids.has(x)) throw Error('Mã dữ liệu không hợp lệ.'); ids.add(x); };
  const name = x => { if (typeof x !== 'string' || !x.trim() || x.length > 120) throw Error('Tên không hợp lệ.'); };
  const obj = x => x && typeof x === 'object' && !Array.isArray(x);
  for (const h of data.holidays) { id(h.id); name(h.name); if (!validDate(h.from) || !validDate(h.to) || h.from > h.to) throw Error('Ngày nghỉ lễ không hợp lệ.'); }
  for (const g of data.groups) {
    id(g.id); name(g.name);
    if (!validDate(g.start) || !Array.isArray(g.schedules) || !g.schedules.length || !Array.isArray(g.students) || !obj(g.overrides) || !Array.isArray(g.extras) || !obj(g.time)) throw Error('Dữ liệu nhóm không hợp lệ.');
    if (!validTime(g.time.start) || !validTime(g.time.end)) throw Error('Giờ học không hợp lệ.');
    if (g.fee !== null && !(Number.isInteger(g.fee) && g.fee >= 0 && g.fee <= 1e9)) throw Error('Học phí không hợp lệ.');
    let previous = '';
    for (const s of g.schedules) {
      if (!validDate(s.from) || s.from < g.start || s.from <= previous || !Array.isArray(s.days) || s.days.length !== DAYS_PER_WEEK || new Set(s.days).size !== DAYS_PER_WEEK || s.days.some(d => !Number.isInteger(d) || d < 0 || d > 6)) throw Error('Lịch học không hợp lệ.');
      previous = s.from;
    }
    if (g.schedules[0].from !== g.start) throw Error('Thiếu lịch ban đầu.');
    for (const d of g.extras) if (!validDate(d) || d < g.start) throw Error('Ngày học bù không hợp lệ.');
    if (new Set(g.extras).size !== g.extras.length) throw Error('Ngày học bù bị trùng.');
    for (const [d, v] of Object.entries(g.overrides)) if (!validDate(d) || !['off', 'done', 'on'].includes(v)) throw Error('Trạng thái buổi học không hợp lệ.');
    for (const s of g.students) {
      id(s.id); name(s.name);
      if (!validDate(s.joined) || !obj(s.payments)) throw Error('Dữ liệu học sinh không hợp lệ.');
      for (const [m, p] of Object.entries(s.payments)) if (!/^[1-9]\d{0,2}$/.test(m) || !obj(p) || !validDate(p.date) || !['CK', 'TM'].includes(p.method) || (p.amount !== undefined && !(Number.isInteger(p.amount) && p.amount >= 0 && p.amount <= 1e9))) throw Error('Học phí không hợp lệ.');
    }
  }
  return data;
}

/** Đọc chuỗi JSON (bản lưu, tệp sao lưu cũ hoặc mới) → dữ liệu hợp lệ. */
export function parseData(raw) {
  let obj;
  try { obj = JSON.parse(raw); } catch { throw Error('Tệp không đúng định dạng Sổ lớp.'); }
  const d = validate(migrate(obj));
  delete d.exportedAt;
  return d;
}
