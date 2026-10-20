import * as C from './core.js';
import { buildXlsx, XLSX_MIME } from './xlsx.js';
import { feeSheet, scheduleSheet, paymentsSheet, printHtml, groupLine } from './report.js';
import { createSync } from './sync.js';

const KEY = 'so-lop-v2', OLD_KEY = 'so-lop-v1', APP_VERSION = '2.1';
const $ = s => document.querySelector(s);
const app = $('#app'), sheet = $('#sheet');
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const { fmt, fmtShort, money, today, WEEKDAY, WEEKDAY_SHORT, PER_MONTH } = C;

/* ================= Dữ liệu ================= */
let data = C.emptyData(), storageError = '';
try {
  const raw = localStorage.getItem(KEY);
  const old = raw ? null : localStorage.getItem(OLD_KEY);
  if (raw || old) data = C.parseData(raw || old);
  if (old) localStorage.setItem(KEY, JSON.stringify(data));
} catch { storageError = 'Không đọc được dữ liệu trên máy. Hãy khôi phục từ Google Drive hoặc tệp sao lưu.'; }
navigator.storage?.persist?.().catch(() => {});

function save(d) {
  try { localStorage.setItem(KEY, JSON.stringify(d)); }
  catch { throw Error('Không lưu được trên máy. Bộ nhớ có thể đã đầy.'); }
}
function commit(change, { restore = false } = {}) {
  if (storageError && !restore) throw Error(storageError);
  const copy = structuredClone(data);
  change(copy);
  copy.updatedAt = new Date().toISOString();
  C.validate(copy);
  save(copy);
  data = copy; storageError = '';
  render();
  sync.markDirty();
}
function replaceData(d) { C.validate(d); save(d); data = d; storageError = ''; render(); }
/** Ghi thay đổi, kèm nút Hoàn tác trên thông báo. */
function commitUndoable(message, change) {
  const before = structuredClone(data);
  commit(change);
  toast(message, { action: 'Hoàn tác', fn: () => commit(d => { d.groups = before.groups; d.holidays = before.holidays; }) });
}

const sync = createSync({
  getData: () => data,
  applyRemote: remote => { replaceData(C.parseData(JSON.stringify(remote))); toast('Đã tải dữ liệu mới nhất từ Google Drive.'); },
  onConflict: remote => conflictSheet(remote),
  onStatus: () => paintSync(),
});

/* ================= Trạng thái giao diện ================= */
const ui = { screen: 'groups', groupId: null, tab: 'fees', feeMonth: null, filter: 'all', q: '', view: 'list', calMonth: null, homeQ: '', statsYm: today().slice(0, 7), statsMethod: 'all', debtGroup: 'all' };
const group = (id = ui.groupId) => data.groups.find(g => g.id === id);
const calCache = new WeakMap();
function calOf(g, min = 12) {
  const k = today() + '|' + min + '|' + data.holidays.length + JSON.stringify(data.holidays);
  const hit = calCache.get(g);
  if (hit && hit.k === k) return hit.cal;
  const cal = C.schedule(g, data.holidays, today(), min);
  calCache.set(g, { k, cal });
  return cal;
}
const nextLesson = cal => cal.rows.find(r => r.status === 'today' || r.status === 'planned');

/* ================= Biểu tượng ================= */
const ICON = {
  users: '<circle cx="9" cy="8" r="3.6"/><path d="M2.5 20v-.8A5.2 5.2 0 0 1 7.7 14h2.6a5.2 5.2 0 0 1 5.2 5.2v.8M16 4.6a3.6 3.6 0 0 1 0 6.8M18.5 14.2a5.2 5.2 0 0 1 3 4.8v1"/>',
  chart: '<path d="M4 20V11M10 20V5M16 20v-6M21 20H3"/>',
  holiday: '<rect x="3" y="4.5" width="18" height="16.5" rx="3.5"/><path d="M8 2.5v4M16 2.5v4M3 9.5h18M9.8 13.3l4.4 4.4M14.2 13.3l-4.4 4.4"/>',
  gear: '<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/><circle cx="15" cy="7" r="2.2"/><circle cx="9" cy="17" r="2.2"/>',
  search: '<circle cx="11" cy="11" r="6.5"/><path d="M20 20l-4.2-4.2"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  back: '<path d="M15 4.5 7.5 12l7.5 7.5"/>',
  next: '<path d="M9 4.5 16.5 12 9 19.5"/>',
  more: '<circle cx="5.5" cy="12" r="1.4" fill="currentColor"/><circle cx="12" cy="12" r="1.4" fill="currentColor"/><circle cx="18.5" cy="12" r="1.4" fill="currentColor"/>',
  cloud: '<path d="M7 18.5a4.5 4.5 0 0 1-.4-9A6 6 0 0 1 18.2 9 4.8 4.8 0 0 1 17.5 18.5z"/>',
  table: '<rect x="3" y="4.5" width="18" height="15" rx="2.5"/><path d="M3 10h18M3 14.8h18M9 4.5v15"/>',
  list: '<path d="M9 6.5h11M9 12h11M9 17.5h11M4.5 6.5h.01M4.5 12h.01M4.5 17.5h.01"/>',
  check: '<path d="M5 12.5 9.5 17 19 7.5"/>',
  share: '<path d="M12 15V3.5M7.5 8 12 3.5 16.5 8M5 12.5V19a1.5 1.5 0 0 0 1.5 1.5h11A1.5 1.5 0 0 0 19 19v-6.5"/>',
  print: '<path d="M7 8.5V3.5h10v5M7 17H5a1.5 1.5 0 0 1-1.5-1.5v-5A2 2 0 0 1 5.5 8.5h13a2 2 0 0 1 2 2v5A1.5 1.5 0 0 1 19 17h-2M7 13.5h10v7H7z"/>',
  edit: '<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/>',
  trash: '<path d="M4 7h16M9.5 7V4.5h5V7M6.5 7l1 13h9l1-13"/>',
  calendar: '<rect x="3" y="4.5" width="18" height="16.5" rx="3.5"/><path d="M8 2.5v4M16 2.5v4M3 9.5h18"/>',
  swap: '<path d="M4 8h13l-3.5-3.5M20 16H7l3.5 3.5"/>',
  close: '<path d="M6 6l12 12M18 6 6 18"/>',
};
const ic = (n, cls = '') => `<svg class="ic ${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[n]}</svg>`;

/* ================= Vẽ màn hình ================= */
let lastScreenKey = '';
function render() {
  const screenKey = ui.screen + '|' + ui.groupId + '|' + ui.tab;
  const fresh = screenKey !== lastScreenKey; lastScreenKey = screenKey;
  const keep = {};
  if (!fresh) app.querySelectorAll('[data-keep]').forEach(el => { keep[el.dataset.keep] = [el.scrollLeft, el.scrollTop]; });
  const y = window.scrollY;
  const focusKey = document.activeElement?.dataset?.focus;
  app.innerHTML = view();
  app.querySelectorAll('[data-keep]').forEach(el => { const k = keep[el.dataset.keep]; if (k) { el.scrollLeft = k[0]; el.scrollTop = k[1]; } });
  if (fresh) { window.scrollTo(0, 0); app.querySelector('.chip.on')?.scrollIntoView({ block: 'nearest', inline: 'center' }); }
  else window.scrollTo(0, y);
  if (focusKey && !sheet.open) focusBy(focusKey);
  paintSync();
}
function focusBy(key) { const el = app.querySelector(`[data-focus="${CSS.escape(key)}"]`); if (el) el.focus({ preventScroll: true }); }
function go(screen, extra = {}) { Object.assign(ui, { screen }, extra); render(); }

function view() {
  const s = ui.screen;
  const body = s === 'group' && group() ? groupView() : s === 'stats' ? statsView() : s === 'holidays' ? holidaysView() : s === 'settings' ? settingsView() : groupsView();
  return `${storageError ? `<div class="alert" role="alert">${esc(storageError)} <button class="link" data-action="settings">Mở Cài đặt</button></div>` : ''}${body}${tabbar()}`;
}
function tabbar() {
  const t = ui.screen === 'group' ? 'groups' : ui.screen;
  const item = (id, icon, label) => `<button class="tab ${t === id ? 'on' : ''}" data-action="nav" data-to="${id}" ${t === id ? 'aria-current="page"' : ''}>${ic(icon)}<span>${label}</span></button>`;
  return `<nav class="tabbar" aria-label="Điều hướng">${item('groups', 'users', 'Nhóm học')}${item('stats', 'chart', 'Thống kê')}${item('holidays', 'holiday', 'Nghỉ lễ')}${item('settings', 'gear', 'Cài đặt')}</nav>`;
}
const syncChip = () => `<button class="sync" data-action="settings" data-sync></button>`;
function paintSync() {
  const st = sync.status;
  const t = st.at ? new Date(st.at) : null;
  const when = t ? `${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')} ${fmtShort(today(t))}` : '';
  const text = { off: 'Chưa lưu lên Drive', idle: 'Google Drive', busy: 'Đang lưu lên Drive…', ok: `Đã lưu Drive ${when}`, offline: 'Chờ mạng để lưu Drive', error: 'Lỗi lưu Drive', conflict: 'Cần chọn bản dữ liệu' }[st.state];
  document.querySelectorAll('[data-sync]').forEach(el => { el.className = `sync s-${st.state}`; el.innerHTML = `${ic('cloud')}<span>${text}</span>`; });
  const box = document.querySelector('[data-sync-detail]');
  if (box) box.textContent = st.message || text;
}

/* ---------- Nhóm học ---------- */
function groupsView() {
  return `<header class="bar"><div class="bar-row"><div><h1>Nhóm học</h1><div class="sub">${WEEKDAY[new Date().getDay()]}, ${fmt(today())}</div></div>
<button class="round primary" data-action="new-group" aria-label="Thêm nhóm học">${ic('plus')}</button></div>
<label class="search">${ic('search')}<input id="q-home" type="search" placeholder="Tìm học sinh trong mọi nhóm" autocomplete="off" autocorrect="off" spellcheck="false" enterkeyhint="search" value="${esc(ui.homeQ)}" data-focus="q-home"></label>${syncChip()}</header>
<main class="content" id="home-body">${homeBody()}</main>`;
}
function homeBody() {
  if (ui.homeQ.trim()) return searchResults();
  if (!data.groups.length) return `<section class="empty">${ic('users', 'big')}<h2>Chưa có nhóm học</h2><p>Tạo nhóm, chọn hai ngày học trong tuần và ngày bắt đầu.</p><button class="btn primary" data-action="new-group">${ic('plus')} Tạo nhóm học</button></section>`;
  return `<div class="cards">${data.groups.map(g => {
    const cal = calOf(g), m = cal.months[cal.current - 1], due = C.overdueList(g, cal).length, nx = nextLesson(cal);
    return `<button class="gcard" data-action="open-group" data-id="${esc(g.id)}" data-focus="g-${esc(g.id)}">
<span class="gtop"><span class="avatar">${esc(g.name.trim().slice(0, 1).toUpperCase())}</span><span class="gname"><strong>${esc(g.name)}</strong><small>${esc(groupLine(g, true))}</small></span>${due ? `<span class="badge due">${due} chưa đóng</span>` : ''}</span>
<span class="gprog"><span>Tháng ${cal.current}</span><span class="meter"><i style="width:${(m?.done || 0) / PER_MONTH * 100}%"></i></span><span>${m?.done || 0}/${PER_MONTH} buổi</span></span>
<span class="gfoot"><span>${g.students.length} học sinh</span><span>${nx ? (nx.status === 'today' ? 'Có buổi hôm nay' : `Buổi tới: ${WEEKDAY_SHORT[C.weekday(nx.date)]} ${fmtShort(nx.date)}`) : ''}</span></span></button>`;
  }).join('')}</div>`;
}
function searchResults() {
  const q = C.fold(ui.homeQ), hits = [];
  for (const g of data.groups) {
    const cal = calOf(g);
    for (const s of g.students) if (C.fold(s.name).includes(q)) {
      const due = cal.months.filter(m => m.number < cal.current && C.payState(s, m, cal.current) === 'overdue').map(m => 'T' + m.number);
      const cur = cal.months[cal.current - 1], paidNow = cur && s.payments[cur.number];
      hits.push({ g, s, tag: due.length ? `<span class="tag due">Chưa đóng ${due.join(', ')}</span>` : paidNow ? `<span class="tag paid">Đã đóng T${cur.number}</span>` : '' });
    }
  }
  if (!hits.length) return `<p class="none">Không tìm thấy học sinh “${esc(ui.homeQ.trim())}”.</p>`;
  return `<div class="list-head">${hits.length} kết quả</div><section class="list">${hits.map(h => `<button class="row" data-action="goto-student" data-g="${esc(h.g.id)}" data-s="${esc(h.s.id)}"><span class="row-main"><strong>${esc(h.s.name)}</strong><small>${esc(h.g.name)} · Vào học ${fmt(h.s.joined)}</small></span>${h.tag}${ic('next', 'chev')}</button>`).join('')}</section>`;
}

/* ---------- Chi tiết nhóm ---------- */
function groupView() {
  const g = group(), cal = calOf(g, Math.max(12, ui.calMonth || 0));
  const cur = cal.months[cal.current - 1], nx = nextLesson(cal), todayRow = cal.rows.find(r => r.date === today() && r.status === 'today');
  return `<header class="bar compact"><div class="bar-row"><button class="back" data-action="home" aria-label="Về danh sách nhóm">${ic('back')}</button>
<h1 class="g-title">${esc(g.name)}</h1><button class="round" data-action="group-menu" aria-label="Quản lý nhóm">${ic('more')}</button></div></header>
<main class="content"><div class="sub g-sub">${esc(groupLine(g))}${g.fee ? ` · ${money(g.fee)}/tháng` : ''}</div>
<section class="hero"><div class="hero-top"><div><div class="eyebrow">Tháng học hiện tại</div><div class="hero-month">Tháng ${cal.current}</div></div><div class="hero-num"><b>${cur?.done || 0}</b>/${PER_MONTH}<small>buổi đã dạy</small></div></div>
<div class="meter light"><i style="width:${(cur?.done || 0) / PER_MONTH * 100}%"></i></div>
<div class="hero-foot"><span>${cur ? `${fmtShort(cur.first)} – ${fmtShort(cur.last)}${cur.done < PER_MONTH ? ' (dự kiến)' : ''}` : ''}</span><span>${nx && !todayRow ? `Buổi tới: ${WEEKDAY[C.weekday(nx.date)]} ${fmtShort(nx.date)}` : ''}</span></div>
${todayRow ? `<button class="btn light" data-action="confirm-today" data-date="${todayRow.date}">${ic('check')} Xác nhận đã dạy hôm nay</button>` : ''}</section>
<div class="seg" role="tablist"><button role="tab" data-action="tab" data-tab="fees" aria-selected="${ui.tab === 'fees'}" class="${ui.tab === 'fees' ? 'on' : ''}">Học phí</button><button role="tab" data-action="tab" data-tab="calendar" aria-selected="${ui.tab === 'calendar'}" class="${ui.tab === 'calendar' ? 'on' : ''}">Lịch dạy</button></div>
${ui.tab === 'fees' ? feesView(g, cal) : calendarView(g, cal)}
</main>`;
}

function feeMonths(g, cal) {
  const maxPaid = Math.max(0, ...g.students.flatMap(s => Object.keys(s.payments).map(Number)));
  return Math.min(cal.months.length, Math.max(cal.current + 1, maxPaid));
}
function feesView(g, cal) {
  const n = feeMonths(g, cal);
  if (!ui.feeMonth || ui.feeMonth > n) ui.feeMonth = cal.current;
  const overdue = C.overdueList(g, cal);
  const per = {}; for (const o of overdue) per[o.month] = (per[o.month] || 0) + 1;
  const m = cal.months[ui.feeMonth - 1], sum = C.monthSummary(g, m, cal.current);
  const chips = Array.from({ length: n }, (_, i) => {
    const k = i + 1, label = k === cal.current ? 'hiện tại' : k > cal.current ? 'sắp tới' : 'đã xong';
    return `<button class="chip ${k === ui.feeMonth ? 'on' : ''}" data-action="fee-month" data-m="${k}" data-focus="chip-${k}" aria-pressed="${k === ui.feeMonth}">T${k}<small>${label}</small>${per[k] ? `<b class="dot" aria-label="${per[k]} chưa đóng">${per[k]}</b>` : ''}</button>`;
  }).join('');
  return `${overdue.length ? `<button class="banner due" data-action="show-overdue">${overdue.length} lượt chưa đóng ở tháng đã kết thúc · ${Object.entries(per).map(([k, v]) => `T${k}: ${v}`).join(' · ')}${ic('next', 'chev')}</button>` : ''}
<div class="chips" data-keep="chips">${chips}</div>
<div class="month-info"><span>Tháng ${m.number}: ${fmtShort(m.first)} – ${fmtShort(m.last)}</span></div>
<section class="summary"><div><b>${sum.paid}</b>/${sum.students}<small>đã đóng</small></div><div class="ck"><b>${sum.CK}</b><small>Chuyển khoản${g.fee ? '<br>' + money(sum.amountCK) : ''}</small></div><div class="tm"><b>${sum.TM}</b><small>Tiền mặt${g.fee ? '<br>' + money(sum.amountTM) : ''}</small></div></section>
<div class="toolbar"><label class="search small">${ic('search')}<input id="q-group" type="search" placeholder="Tìm tên trong nhóm" autocomplete="off" autocorrect="off" spellcheck="false" enterkeyhint="search" value="${esc(ui.q)}" data-focus="q-group"></label>
<button class="round" data-action="toggle-view" data-focus="toggle-view" aria-label="${ui.view === 'list' ? 'Xem dạng bảng' : 'Xem dạng danh sách'}">${ic(ui.view === 'list' ? 'table' : 'list')}</button>
<button class="round primary" data-action="new-student" aria-label="Thêm học sinh">${ic('plus')}</button></div>
<div id="fee-body">${feeBody(g, cal)}</div>`;
}
function feeBody(g, cal) {
  const m = cal.months[ui.feeMonth - 1], q = C.fold(ui.q);
  const all = g.students.map((s, i) => ({ s, no: i + 1, st: C.payState(s, m, cal.current) })).filter(x => !q || C.fold(x.s.name).includes(q));
  const cnt = { all: all.length, unpaid: all.filter(x => x.st === 'overdue' || x.st === 'open').length, paid: all.filter(x => x.st === 'paid').length };
  const list = all.filter(x => ui.filter === 'all' || (ui.filter === 'paid' ? x.st === 'paid' : x.st === 'overdue' || x.st === 'open'));
  const seg = `<div class="seg small">${[['all', 'Tất cả'], ['unpaid', 'Chưa đóng'], ['paid', 'Đã đóng']].map(([k, l]) => `<button data-action="filter" data-f="${k}" data-focus="f-${k}" class="${ui.filter === k ? 'on' : ''}" aria-pressed="${ui.filter === k}">${l} <span class="n">${cnt[k]}</span></button>`).join('')}</div>`;
  if (!g.students.length) return `<section class="empty small"><h2>Chưa có học sinh</h2><p>Thêm tên và ngày vào học của từng học sinh.</p><button class="btn primary" data-action="new-student">${ic('plus')} Thêm học sinh</button></section>`;
  const none = !list.length ? `<p class="none">${q ? `Không có học sinh khớp “${esc(ui.q.trim())}”.` : ui.filter === 'paid' ? 'Chưa có học sinh nào đóng tháng này.' : 'Tất cả học sinh đã đóng tháng này.'}</p>` : '';
  if (ui.view === 'table') return seg + (none || tableView(g, cal, list));
  return seg + (none || `<section class="list">${list.map(({ s, no, st }) => `<div class="srow"><button class="srow-main" data-action="edit-student" data-id="${esc(s.id)}" data-focus="st-${esc(s.id)}"><span class="no">${no}</span><span class="nm"><strong>${esc(s.name)}</strong><small>Vào học ${fmt(s.joined)}</small></span></button>${payButton(s, m.number, st, 'pay')}</div>`).join('')}</section>`);
}
function payLabel(s, k, st, compact) {
  const p = s.payments[k];
  if (st === 'paid') return `<b>${compact ? fmtShort(p.date) : fmtShort(p.date)}</b><small>${p.method === 'CK' ? 'Chuyển khoản' : 'Tiền mặt'}</small>`;
  if (st === 'overdue') return '<b>Chưa đóng</b>';
  if (st === 'na') return '<small>Chưa vào học</small>';
  return compact ? '<small>＋</small>' : '<b>Ghi đóng</b>';
}
function payButton(s, k, st, kind) {
  const label = st === 'paid' ? `đã đóng ${fmt(s.payments[k].date)}` : st === 'overdue' ? 'chưa đóng' : st === 'na' ? 'chưa vào học' : 'chưa ghi';
  return `<button class="${kind} p-${st} ${s.payments[k]?.method === 'TM' ? 'tm' : ''}" data-action="pay" data-id="${esc(s.id)}" data-m="${k}" data-focus="${kind}-${esc(s.id)}-${k}" aria-label="${esc(s.name)}, tháng ${k}, ${label}">${payLabel(s, k, st, kind === 'cell')}</button>`;
}
function tableView(g, cal, list) {
  const n = feeMonths(g, cal), months = cal.months.slice(0, n);
  return `<div class="table-wrap" data-keep="table" tabindex="0" aria-label="Bảng học phí"><table><thead><tr><th class="sticky">Học sinh</th>${months.map(m => `<th class="${m.number === ui.feeMonth ? 'sel' : ''}">T${m.number}<small>${fmtShort(m.first)}–${fmtShort(m.last)}</small></th>`).join('')}</tr></thead>
<tbody>${list.map(({ s, no }) => `<tr><td class="sticky"><button class="tname" data-action="edit-student" data-id="${esc(s.id)}" data-focus="tst-${esc(s.id)}"><span class="no">${no}</span>${esc(s.name)}</button></td>${months.map(m => `<td class="${m.number === ui.feeMonth ? 'sel' : ''}">${payButton(s, m.number, C.payState(s, m, cal.current), 'cell')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}

function calendarView(g, cal) {
  if (!ui.calMonth) ui.calMonth = cal.current;
  ui.calMonth = Math.min(ui.calMonth, cal.months.length);
  const m = cal.months[ui.calMonth - 1], time = g.time.start ? ` · ${g.time.start}${g.time.end ? '–' + g.time.end : ''}` : '';
  const status = { done: 'Đã dạy', off: 'Nghỉ', planned: 'Dự kiến', today: 'Hôm nay' };
  return `<div class="cal-nav"><button class="round" data-action="cal-prev" data-focus="cal-prev" ${ui.calMonth <= 1 ? 'disabled' : ''} aria-label="Tháng trước">${ic('back')}</button>
<div class="cal-title"><strong>Tháng ${m.number}${m.number === cal.current ? ' · hiện tại' : ''}</strong><small>${fmt(m.first)} – ${fmt(m.last)} · ${m.done}/${PER_MONTH} đã dạy</small></div>
<button class="round" data-action="cal-next" data-focus="cal-next" ${ui.calMonth >= 240 ? 'disabled' : ''} aria-label="Tháng sau">${ic('next')}</button></div>
<div class="pair"><button class="btn tonal" data-action="add-extra">${ic('plus')} Buổi học bù</button><button class="btn tonal" data-action="change-schedule">${ic('swap')} Đổi lịch</button></div>
<section class="list">${m.rows.map(r => `<button class="lesson st-${r.status}" data-action="lesson" data-date="${r.date}" data-focus="ls-${r.date}">
<span class="dtile"><b>${r.date.slice(8)}</b><small>Th${Number(r.date.slice(5, 7))}</small></span>
<span class="row-main"><strong>${WEEKDAY[C.weekday(r.date)]}${r.extra ? ' <em class="tag extra">Học bù</em>' : ''}</strong><small>${r.status === 'off' ? (r.holiday ? 'Nghỉ lễ: ' + esc(r.holiday) : 'Không tính buổi') : `Buổi ${r.slot}${time}`}</small></span>
<span class="pill ${r.status}">${status[r.status]}</span></button>`).join('')}</section>`;
}

/* ---------- Thống kê ---------- */
function statsView() {
  const allPays = C.paymentsIn(data, '');
  const yms = [...new Set([today().slice(0, 7), ...allPays.map(p => p.date.slice(0, 7))])].sort().reverse();
  if (ui.statsYm !== 'all' && !yms.includes(ui.statsYm)) ui.statsYm = yms[0];
  const sum = a => a.reduce((s, p) => s + p.amount, 0);
  const split = a => ({ ck: a.filter(p => p.method === 'CK'), tm: a.filter(p => p.method === 'TM') });
  const life = split(allPays);
  const period = C.paymentsIn(data, ui.statsYm === 'all' ? '' : ui.statsYm), per = split(period);
  const list = ui.statsMethod === 'all' ? period : period.filter(p => p.method === ui.statsMethod);
  const label = ym => ym === 'all' ? 'Từ đầu' : `Th${Number(ym.slice(5))}/${ym.slice(0, 4)}`;
  const longLabel = ym => ym === 'all' ? 'từ lúc bắt đầu' : `tháng ${Number(ym.slice(5))}/${ym.slice(0, 4)}`;
  const noFee = data.groups.filter(g => !g.fee);
  const byGroup = data.groups.map(g => { const x = period.filter(p => p.group === g), sp = split(x); return { g, all: x, ...sp }; }).filter(x => x.all.length);
  // Học sinh còn nợ
  const debts = [];
  for (const g of data.groups) { const cal = calOf(g); for (const s of g.students) { const d = C.debtOf(g, s, cal); if (d.months.length || d.partial) debts.push({ g, s, ...d }); } }
  if (ui.debtGroup !== 'all' && !data.groups.some(g => g.id === ui.debtGroup)) ui.debtGroup = 'all';
  const debtList = debts.filter(d => ui.debtGroup === 'all' || d.g.id === ui.debtGroup).sort((a, b) => b.amount - a.amount || a.s.name.localeCompare(b.s.name, 'vi'));
  const debtTotal = debtList.reduce((t, d) => t + d.amount, 0);
  const debtWhat = d => [d.months.length ? 'T' + d.months.join(', T') : '', d.partial ? `${d.partial.sessions} buổi T${d.partial.month}` : ''].filter(Boolean).join(' + ');
  let lastDate = '';
  return `<header class="bar"><div class="bar-row"><div><h1>Thống kê</h1><div class="sub">Học phí đã thu và còn nợ</div></div>
<button class="round" data-action="export-all" aria-label="Xuất Excel tất cả nhóm">${ic('share')}</button></div></header>
<main class="content">
<section class="lifetime"><small>Tổng đã thu từ lúc bắt đầu</small><b>${money(sum(allPays))}</b><span class="lt-split"><span class="ck">CK ${money(sum(life.ck))}</span><span class="tm">Tiền mặt ${money(sum(life.tm))}</span></span><em>${allPays.length} lượt đóng</em></section>
${noFee.length ? `<button class="banner warn" data-action="fix-fee" data-g="${esc(noFee[0].id)}">${noFee.length === 1 ? `Nhóm ${esc(noFee[0].name)}` : `${noFee.length} nhóm`} chưa nhập học phí nên chưa tính được tiền. Chạm để nhập.${ic('next', 'chev')}</button>` : ''}
<h2 class="section">Theo tháng dương lịch</h2>
<div class="chips" data-keep="ym">${[...yms, 'all'].map(ym => `<button class="chip ${ui.statsYm === ym ? 'on' : ''}" data-action="stats-ym" data-ym="${ym}" data-focus="ym-${ym}">${label(ym)}</button>`).join('')}</div>
<div class="period-total">Đã thu ${longLabel(ui.statsYm)}: <b>${money(sum(period))}</b></div>
<div class="seg small" role="group" aria-label="Chọn hình thức">${[['all', 'Tất cả', sum(period)], ['CK', 'Chuyển khoản', sum(per.ck)], ['TM', 'Tiền mặt', sum(per.tm)]].map(([k, l, v]) => `<button data-action="stats-method" data-m="${k}" data-focus="sm-${k}" class="${ui.statsMethod === k ? 'on' : ''}" aria-pressed="${ui.statsMethod === k}"><span class="seg-l">${l}</span><span class="seg-v">${money(v)}</span></button>`).join('')}</div>
${byGroup.length ? `<h2 class="section">Theo nhóm · ${longLabel(ui.statsYm)}</h2><section class="list">${byGroup.map(x => `<div class="row static grow"><span class="row-main"><strong>${esc(x.g.name)}</strong><small>${x.all.length} lượt</small></span><span class="amt">${x.g.fee || x.all.some(p => p.amount) ? money(sum(x.all)) : '<small>Chưa có học phí</small>'}</span><span class="tags"><span class="tag ck">CK ${x.ck.length}</span><span class="tag tm">TM ${x.tm.length}</span></span></div>`).join('')}</section>` : ''}
<h2 class="section due-t">Học sinh chưa đóng${debtList.length ? ` · ${debtList.length} em · ${money(debtTotal)}` : ''}</h2>
${data.groups.length > 1 ? `<div class="chips" data-keep="debt">${[['all', 'Tất cả nhóm'], ...data.groups.map(g => [g.id, g.name])].map(([id, n]) => `<button class="chip small ${ui.debtGroup === id ? 'on' : ''}" data-action="debt-group" data-g="${esc(id)}" data-focus="dg-${esc(id)}">${esc(n)}</button>`).join('')}</div>` : ''}
${debtList.length ? `<section class="list">${debtList.map(d => `<button class="row grow" data-action="stats-pay" data-g="${esc(d.g.id)}" data-s="${esc(d.s.id)}" data-m="${d.months[0] || d.partial.month}"><span class="row-main"><strong>${esc(d.s.name)}</strong><small>${esc(d.g.name)} · ${debtWhat(d)}</small></span><span class="amt due">${d.hasFee ? money(d.amount) : '<small>Chưa có học phí</small>'}</span><span class="tag due">Chưa đóng</span></button>`).join('')}</section>` : `<p class="none">Không có học sinh nào còn nợ.</p>`}
<h2 class="section">${ui.statsMethod === 'CK' ? 'Các lượt chuyển khoản' : ui.statsMethod === 'TM' ? 'Các lượt tiền mặt' : 'Các lượt đóng tiền'} · ${longLabel(ui.statsYm)}</h2>
${list.length ? `<section class="list">${list.map(p => { const head = p.date !== lastDate ? `<div class="list-date">${WEEKDAY[C.weekday(p.date)]}, ${fmt(p.date)}</div>` : ''; lastDate = p.date; return `${head}<button class="row" data-action="stats-pay" data-g="${esc(p.group.id)}" data-s="${esc(p.student.id)}" data-m="${p.month}"><span class="row-main"><strong>${esc(p.student.name)}</strong><small>${esc(p.group.name)} · Tháng ${p.month}${p.amount ? ' · ' + money(p.amount) : ''}</small></span><span class="tag ${p.method === 'CK' ? 'ck' : 'tm'}">${p.method === 'CK' ? 'CK' : 'Tiền mặt'}</span></button>`; }).join('')}</section>` : `<p class="none">Chưa có lượt đóng tiền nào.</p>`}
${data.groups.length ? `<div class="pair"><button class="btn tonal" data-action="export-all">${ic('share')} Xuất Excel</button><button class="btn tonal" data-action="print-all">${ic('print')} In bảng học phí</button></div>` : ''}
</main>`;
}

/* ---------- Nghỉ lễ ---------- */
function holidaysView() {
  const list = [...data.holidays].sort((a, b) => b.from.localeCompare(a.from));
  const affected = h => data.groups.reduce((n, g) => n + calOf(g).rows.filter(r => r.date >= h.from && r.date <= h.to && r.holiday).length, 0);
  return `<header class="bar"><div class="bar-row"><div><h1>Nghỉ lễ</h1><div class="sub">Áp dụng cho tất cả nhóm học</div></div>
<button class="round primary" data-action="new-holiday" aria-label="Thêm ngày nghỉ lễ">${ic('plus')}</button></div></header>
<main class="content">${list.length ? `<section class="list">${list.map(h => `<button class="row" data-action="edit-holiday" data-id="${esc(h.id)}" data-focus="h-${esc(h.id)}"><span class="row-main"><strong>${esc(h.name)}</strong><small>${h.from === h.to ? fmt(h.from) : `${fmt(h.from)} – ${fmt(h.to)}`} · ${affected(h)} buổi được nghỉ</small></span>${ic('next', 'chev')}</button>`).join('')}</section>` :
    `<section class="empty">${ic('holiday', 'big')}<h2>Chưa có ngày nghỉ lễ</h2><p>Thêm một lần, tất cả nhóm học đều nghỉ và lịch tự nối đủ 8 buổi.</p><button class="btn primary" data-action="new-holiday">${ic('plus')} Thêm ngày nghỉ lễ</button></section>`}</main>`;
}

/* ---------- Cài đặt ---------- */
function settingsView() {
  const st = sync.status;
  return `<header class="bar"><div class="bar-row"><div><h1>Cài đặt</h1></div></div></header>
<main class="content">
<h2 class="section">Google Drive</h2>
<section class="card">${sync.configured ? `<div class="drive-ok">${ic('cloud')}<div><strong>Đã kết nối</strong><small data-sync-detail></small></div></div>
${st.state === 'conflict' ? `<button class="btn primary wide" data-action="resolve-conflict">Chọn bản dữ liệu</button>` : ''}
<div class="pair"><button class="btn tonal" data-action="sync-now">Đồng bộ ngay</button><button class="btn tonal danger" data-action="disconnect">Ngắt kết nối</button></div>` :
    `<p class="muted">Mỗi lần sửa, dữ liệu tự lưu lên Google Drive của bạn. Đổi máy chỉ cần kết nối lại là có đủ dữ liệu.</p>
<form id="drive-form"><label class="field"><span>Địa chỉ ứng dụng web</span><input name="url" type="url" inputmode="url" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="https://script.google.com/macros/s/…/exec" required></label>
<label class="field"><span>Mã kết nối</span><input name="key" autocomplete="off" autocapitalize="characters" spellcheck="false" placeholder="20 ký tự" required></label>
<p class="error" id="drive-error" role="alert"></p><button class="btn primary wide" type="submit">Kết nối Google Drive</button></form>`}</section>
<h2 class="section">Tệp sao lưu</h2>
<section class="list">
<button class="row" data-action="export-json">${ic('share')}<span class="row-main"><strong>Tải tệp sao lưu</strong><small>Toàn bộ nhóm, lịch và học phí</small></span></button>
<button class="row" data-action="import-json">${ic('swap')}<span class="row-main"><strong>Khôi phục từ tệp</strong><small>Thay thế toàn bộ dữ liệu hiện tại</small></span></button>
</section><input id="import-file" type="file" accept="application/json,.json" hidden>
<p class="version">Sổ lớp ${APP_VERSION}</p></main>`;
}

/* ================= Bảng thao tác (bottom sheet) ================= */
let returnFocus = null;
function openSheet(title, body, setup) {
  if (!sheet.open) returnFocus = document.activeElement?.dataset?.focus || null;
  sheet.innerHTML = `<div class="grab" aria-hidden="true"></div><header class="sheet-head"><h2 tabindex="-1" autofocus>${title}</h2><button type="button" class="round ghost" data-close aria-label="Đóng">${ic('close')}</button></header><div class="sheet-body">${body}</div>`;
  sheet.querySelector('[data-close]').onclick = closeSheet;
  if (!sheet.open) sheet.showModal();
  sheet.scrollTop = 0;
  setup?.(sheet);
}
function closeSheet() { if (sheet.open) sheet.close(); }
sheet.addEventListener('close', () => {
  if (sheet.open) return; // đã mở bảng khác ngay sau khi đóng
  sheet.innerHTML = '';
  if (returnFocus) focusBy(returnFocus);
});
sheet.addEventListener('click', ev => { if (ev.target === sheet) closeSheet(); });
function formHandler(form, fn) {
  form.addEventListener('submit', ev => {
    ev.preventDefault();
    const err = form.querySelector('.error');
    try { if (err) err.textContent = ''; fn(new FormData(form), ev.submitter); }
    catch (e) { if (err) err.textContent = e.message; else toast(e.message); }
  });
}
const errBox = '<p class="error" role="alert"></p>';

function daysPicker(chosen = []) {
  return `<div class="days" role="group" aria-label="Hai ngày học mỗi tuần">${[1, 2, 3, 4, 5, 6, 0].map(d => `<label class="day"><input type="checkbox" name="days" value="${d}" ${chosen.includes(d) ? 'checked' : ''}><span>${WEEKDAY_SHORT[d]}</span></label>`).join('')}</div>`;
}
function bindDays(root, onChange) {
  const boxes = [...root.querySelectorAll('input[name=days]')];
  const sync2 = () => { const n = boxes.filter(b => b.checked).length; boxes.forEach(b => { b.disabled = !b.checked && n >= C.DAYS_PER_WEEK; }); onChange?.(); };
  boxes.forEach(b => b.addEventListener('change', sync2)); sync2();
}
function readDays(f) { const ds = f.getAll('days').map(Number); if (ds.length !== C.DAYS_PER_WEEK) throw Error('Hãy chọn đúng 2 ngày học trong tuần.'); return ds; }
const parseFee = s => { const d = String(s || '').replace(/\D/g, ''); return d ? Math.min(1e9, Number(d)) : null; };
const feeText = n => n ? Number(n).toLocaleString('vi-VN') : '';
function bindFee(root) { const i = root.querySelector('input[name=fee]'); i?.addEventListener('input', () => { const v = parseFee(i.value); i.value = feeText(v); }); }

function groupSheet(edit) {
  const g = edit ? group() : null;
  openSheet(edit ? 'Sửa nhóm học' : 'Nhóm học mới', `<form>
<label class="field"><span>Tên nhóm</span><input name="name" required maxlength="120" placeholder="Ví dụ: Toán 8 – tối" value="${esc(g?.name || '')}"></label>
${edit ? '' : `<label class="field"><span>Ngày bắt đầu</span><input name="start" type="date" required value="${today()}"></label><div class="field"><span>Hai ngày học mỗi tuần</span>${daysPicker()}</div>`}
<div class="field-row"><label class="field"><span>Giờ bắt đầu</span><input name="tstart" type="time" value="${g?.time.start || ''}"></label><label class="field"><span>Giờ kết thúc</span><input name="tend" type="time" value="${g?.time.end || ''}"></label></div>
<label class="field"><span>Học phí mỗi tháng (không bắt buộc)</span><input name="fee" inputmode="numeric" placeholder="Ví dụ: 400.000" value="${feeText(g?.fee)}"><em class="suffix">đ</em></label>
${errBox}<button class="btn primary wide" type="submit">${edit ? 'Lưu thay đổi' : 'Tạo nhóm'}</button></form>`, root => {
    if (!edit) bindDays(root);
    bindFee(root);
    formHandler(root.querySelector('form'), f => {
      const name = f.get('name').trim(), time = { start: f.get('tstart') || '', end: f.get('tend') || '' }, fee = parseFee(f.get('fee'));
      if (!name) throw Error('Hãy nhập tên nhóm.');
      if (time.end && !time.start) throw Error('Hãy nhập giờ bắt đầu.');
      if (time.start && time.end && time.end <= time.start) throw Error('Giờ kết thúc phải sau giờ bắt đầu.');
      if (edit) { commit(d => Object.assign(d.groups.find(x => x.id === g.id), { name, time, fee })); closeSheet(); toast('Đã lưu nhóm học.'); return; }
      const start = f.get('start'), days = readDays(f), id = uid();
      if (!C.validDate(start)) throw Error('Hãy chọn ngày bắt đầu.');
      commit(d => d.groups.push({ id, name, start, time, fee, schedules: [{ from: start, days: days.sort() }], extras: [], overrides: {}, students: [] }));
      closeSheet(); openGroup(id); toast('Đã tạo nhóm. Thêm học sinh để bắt đầu.');
    });
  });
}
function groupMenu() {
  const g = group();
  openSheet(esc(g.name), `<section class="menu">
<button data-m="edit">${ic('edit')}Sửa tên, giờ học, học phí</button>
<button data-m="schedule">${ic('swap')}Đổi lịch dạy</button>
<button data-m="extra">${ic('plus')}Thêm buổi học bù</button>
<button data-m="excel">${ic('share')}Xuất Excel bảng học phí</button>
<button data-m="print">${ic('print')}In bảng học phí</button>
<button data-m="delete" class="danger">${ic('trash')}Xóa nhóm</button></section>`, root => {
    root.querySelector('.menu').addEventListener('click', ev => {
      const b = ev.target.closest('[data-m]'); if (!b) return;
      const m = b.dataset.m;
      if (m === 'excel') { closeSheet(); exportGroup(g); return; }
      if (m === 'print') { closeSheet(); printGroups([g]); return; }
      ({ edit: () => groupSheet(true), schedule: scheduleSheet2, extra: extraSheet, delete: deleteGroupSheet })[m]();
    });
  });
}
function deleteGroupSheet() {
  const g = group();
  openSheet('Xóa nhóm học?', `<p>Nhóm <b>${esc(g.name)}</b> cùng ${g.students.length} học sinh, lịch học và học phí sẽ bị xóa.</p><div class="pair"><button class="btn tonal" data-close2>Không xóa</button><button class="btn danger-fill" data-yes>Xóa nhóm</button></div>`, root => {
    root.querySelector('[data-close2]').onclick = closeSheet;
    root.querySelector('[data-yes]').onclick = () => { closeSheet(); history.state?.view === 'group' ? history.back() : null; ui.screen = 'groups'; commitUndoable(`Đã xóa nhóm ${g.name}.`, d => { d.groups = d.groups.filter(x => x.id !== g.id); }); };
  });
}
function scheduleSheet2() {
  const g = group(), seg = C.currentSegment(g), from0 = today() < g.start ? g.start : today();
  openSheet('Đổi lịch dạy', `<form><label class="field"><span>Áp dụng từ ngày</span><input name="from" type="date" min="${g.start}" required value="${from0}"></label>
<div class="field"><span>Hai ngày học mới</span>${daysPicker(seg.days)}</div><p class="note" id="sch-note"></p>${errBox}<button class="btn primary wide" type="submit">Áp dụng lịch mới</button></form>`, root => {
    const f = root.querySelector('form'), note = root.querySelector('#sch-note');
    const upd = () => {
      const from = f.from.value, ds = [...f.querySelectorAll('input[name=days]:checked')].map(i => Number(i.value));
      if (!C.validDate(from) || ds.length !== 2) { note.textContent = ''; return; }
      const lost = C.changePreview(g, from, ds);
      note.textContent = `Các buổi trước ${fmt(from)} giữ nguyên.${lost ? ` ${lost} buổi đã báo nghỉ/xác nhận sẽ bỏ vì không còn trong lịch mới.` : ''}${from < today() ? ' Lịch từ ngày này đến hôm nay sẽ được tính lại.' : ''}`;
    };
    bindDays(root, upd); f.from.addEventListener('input', upd); upd();
    formHandler(f, fd => { const from = fd.get('from'), ds = readDays(fd); commit(d => C.changeSchedule(d.groups.find(x => x.id === g.id), from, ds)); closeSheet(); toast('Đã đổi lịch dạy.'); });
  });
}
function extraSheet() {
  const g = group();
  openSheet('Thêm buổi học bù', `<form><label class="field"><span>Ngày học bù</span><input name="date" type="date" min="${g.start}" required value="${today()}"></label><p class="note">Buổi học bù được tính như một buổi trong tháng học.</p>${errBox}<button class="btn primary wide" type="submit">Thêm buổi học bù</button></form>`, root => {
    formHandler(root.querySelector('form'), f => { const date = f.get('date'); commit(d => C.addExtra(d.groups.find(x => x.id === g.id), date)); closeSheet(); ui.tab = 'calendar'; ui.calMonth = calOf(group(), 240).rows.find(r => r.date === date)?.month || null; render(); toast(`Đã thêm buổi học bù ${fmt(date)}.`); });
  });
}
function lessonSheet(date) {
  const g = group(), cal = calOf(g, Math.max(12, ui.calMonth || 0)), r = cal.rows.find(x => x.date === date); if (!r) return;
  const t = today(), ov = g.overrides[date], hol = !r.extra && C.holidayOn(data.holidays, date);
  const auto = date > t ? 'Dự kiến' : date < t ? 'Đã dạy' : 'Chưa xác nhận';
  let opts;
  if (r.extra) opts = [['auto', auto], ...(date === t ? [['done', 'Đã dạy']] : [])];
  else if (hol) opts = [['auto', `Nghỉ lễ · ${hol.name}`], ['on', `Vẫn dạy buổi này${date < t ? ' · Đã dạy' : ''}`], ...(date === t ? [['done', 'Đã dạy']] : [])];
  else opts = [['auto', auto], ...(date === t ? [['done', 'Đã dạy']] : []), ['off', 'Nghỉ · không tính buổi']];
  const curV = !ov || (hol && ov === 'off') ? 'auto' : ov;
  openSheet(`${WEEKDAY[C.weekday(date)]}, ${fmt(date)}`, `<p class="sheet-sub">${r.extra ? 'Buổi học bù · ' : ''}${r.status === 'off' ? 'Không tính buổi' : `Tháng ${r.month} · Buổi ${r.slot}`}</p>
<section class="options">${opts.map(([v, l]) => `<button class="option ${curV === v ? 'on' : ''}" data-v="${v}" aria-pressed="${curV === v}"><span>${esc(l)}</span>${curV === v ? ic('check') : ''}</button>`).join('')}</section>
${r.extra ? `<button class="btn tonal danger wide" data-remove>${ic('trash')} Xóa buổi học bù</button>` : ''}`, root => {
    root.querySelector('.options').addEventListener('click', ev => {
      const b = ev.target.closest('[data-v]'); if (!b) return;
      const v = b.dataset.v; closeSheet();
      if (v === curV) return;
      commit(d => { const x = d.groups.find(y => y.id === g.id); if (v === 'auto') delete x.overrides[date]; else x.overrides[date] = v; });
      toast('Đã cập nhật buổi học. Các tháng phía sau tự sắp xếp lại.');
    });
    root.querySelector('[data-remove]')?.addEventListener('click', () => { closeSheet(); commitUndoable('Đã xóa buổi học bù.', d => C.removeExtra(d.groups.find(x => x.id === g.id), date)); });
  });
}
function studentSheet(sid) {
  const g = group(), s = g.students.find(x => x.id === sid), cal = calOf(g);
  const history = s ? cal.months.slice(0, feeMonths(g, cal)).map(m => ({ m, st: C.payState(s, m, cal.current) })) : [];
  openSheet(s ? 'Học sinh' : 'Thêm học sinh', `<form><label class="field"><span>Họ và tên</span><input name="name" required maxlength="120" autocomplete="off" autocapitalize="words" placeholder="Nhập họ và tên" value="${esc(s?.name || '')}"></label>
<label class="field"><span>Ngày vào học</span><input name="joined" type="date" required value="${s?.joined || today()}"></label>${errBox}
${s ? `<button class="btn primary wide" type="submit">Lưu</button>` : `<div class="pair"><button class="btn tonal" type="submit" value="more">Lưu, thêm tiếp</button><button class="btn primary" type="submit" value="done">Lưu</button></div>`}</form>
${s ? `<h3 class="sheet-h3">Học phí các tháng</h3><section class="hist">${history.map(({ m, st }) => `<button class="hrow p-${st}" data-m="${m.number}"><span>Tháng ${m.number}<small>${fmtShort(m.first)}–${fmtShort(m.last)}</small></span><span>${st === 'paid' ? `${fmt(s.payments[m.number].date)} · ${s.payments[m.number].method === 'CK' ? 'CK' : 'Tiền mặt'}` : st === 'overdue' ? 'Chưa đóng' : st === 'na' ? 'Chưa vào học' : 'Chưa ghi'}</span></button>`).join('')}</section>
<button class="btn tonal danger wide" data-delete>${ic('trash')} Xóa học sinh</button>` : ''}`, root => {
    const form = root.querySelector('form');
    if (!s) setTimeout(() => form.name.focus(), 60);
    formHandler(form, (f, submitter) => {
      const name = f.get('name').trim().replace(/\s+/g, ' '), joined = f.get('joined');
      if (!name) throw Error('Hãy nhập họ và tên.');
      if (!C.validDate(joined)) throw Error('Hãy chọn ngày vào học.');
      if (!s && g.students.some(x => C.fold(x.name) === C.fold(name)) && form.dataset.dup !== name) { form.dataset.dup = name; throw Error('Tên này đã có trong nhóm. Bấm Lưu lần nữa nếu vẫn muốn thêm.'); }
      if (s) commit(d => Object.assign(d.groups.find(x => x.id === g.id).students.find(x => x.id === sid), { name, joined }));
      else commit(d => d.groups.find(x => x.id === g.id).students.push({ id: uid(), name, joined, payments: {} }));
      if (!s && submitter?.value === 'more') { form.name.value = ''; form.dataset.dup = ''; form.name.focus(); toast(`Đã thêm ${name}.`); return; }
      closeSheet(); toast(s ? 'Đã lưu học sinh.' : `Đã thêm ${name}.`);
    });
    root.querySelector('.hist')?.addEventListener('click', ev => { const b = ev.target.closest('[data-m]'); if (b) paySheet(g.id, sid, Number(b.dataset.m)); });
    root.querySelector('[data-delete]')?.addEventListener('click', () => { closeSheet(); commitUndoable(`Đã xóa ${s.name}.`, d => { const x = d.groups.find(y => y.id === g.id); x.students = x.students.filter(y => y.id !== sid); }); });
  });
}
function paySheet(gid, sid, k) {
  const g = group(gid), s = g?.students.find(x => x.id === sid); if (!s) return;
  const cal = calOf(g), m = cal.months[k - 1], p = s.payments[k];
  openSheet(`Học phí tháng ${k}`, `<p class="sheet-sub"><b>${esc(s.name)}</b> · ${esc(g.name)}${m ? ` · ${fmtShort(m.first)}–${fmtShort(m.last)}` : ''}${g.fee ? ` · ${money(g.fee)}` : ''}</p>
<div class="field-row"><label class="field"><span>Ngày đóng</span><input id="pay-date" type="date" required value="${p?.date || today()}"></label>
<label class="field"><span>Số tiền</span><input id="pay-amount" name="fee" inputmode="numeric" placeholder="${g.fee ? '' : 'Chưa có học phí'}" value="${feeText(p ? C.payAmount(g, p) : g.fee)}"><em class="suffix">đ</em></label></div>
<div class="field"><span>Chạm hình thức để lưu</span><div class="paybtns"><button class="paybtn ck ${p?.method === 'CK' ? 'on' : ''}" data-method="CK">${p?.method === 'CK' ? ic('check') : ''}Chuyển khoản</button><button class="paybtn tm ${p?.method === 'TM' ? 'on' : ''}" data-method="TM">${p?.method === 'TM' ? ic('check') : ''}Tiền mặt</button></div></div>
${errBox}${p ? `<button class="btn tonal danger wide" data-clear>${ic('trash')} Xóa ngày đóng</button>` : ''}`, root => {
    bindFee(root);
    root.querySelectorAll('[data-method]').forEach(b => b.addEventListener('click', () => {
      const date = root.querySelector('#pay-date').value, method = b.dataset.method;
      if (!C.validDate(date)) { root.querySelector('.error').textContent = 'Hãy chọn ngày đóng.'; return; }
      const amount = parseFee(root.querySelector('#pay-amount').value);
      const rec = amount === null ? { date, method } : { date, method, amount };
      try { commit(d => { d.groups.find(x => x.id === gid).students.find(x => x.id === sid).payments[k] = rec; }); }
      catch (e) { root.querySelector('.error').textContent = e.message; return; }
      closeSheet(); toast(`${s.name} · Tháng ${k} · ${method === 'CK' ? 'Chuyển khoản' : 'Tiền mặt'}${amount ? ' ' + money(amount) : ''} · ${fmtShort(date)}`);
    }));
    root.querySelector('[data-clear]')?.addEventListener('click', () => { closeSheet(); commitUndoable(`Đã xóa ngày đóng tháng ${k} của ${s.name}.`, d => { delete d.groups.find(x => x.id === gid).students.find(x => x.id === sid).payments[k]; }); });
  });
}
function holidaySheet(hid) {
  const h = data.holidays.find(x => x.id === hid);
  openSheet(h ? 'Sửa ngày nghỉ lễ' : 'Thêm ngày nghỉ lễ', `<form><label class="field"><span>Tên</span><input name="name" required maxlength="80" placeholder="Ví dụ: Nghỉ Tết Nguyên đán" value="${esc(h?.name || '')}"></label>
<div class="field-row"><label class="field"><span>Từ ngày</span><input name="from" type="date" required value="${h?.from || today()}"></label><label class="field"><span>Đến ngày</span><input name="to" type="date" required value="${h?.to || today()}"></label></div>
<p class="note">Tất cả nhóm học đều nghỉ các buổi trong khoảng này; lịch tự nối đủ 8 buổi.</p>${errBox}<button class="btn primary wide" type="submit">Lưu</button></form>
${h ? `<button class="btn tonal danger wide" data-delete>${ic('trash')} Xóa ngày nghỉ lễ</button>` : ''}`, root => {
    const f = root.querySelector('form');
    f.from.addEventListener('change', () => { if (!f.to.value || f.to.value < f.from.value) f.to.value = f.from.value; });
    formHandler(f, fd => {
      const name = fd.get('name').trim(), from = fd.get('from'), to = fd.get('to');
      if (!name) throw Error('Hãy nhập tên ngày nghỉ.');
      if (!C.validDate(from) || !C.validDate(to) || to < from) throw Error('Ngày kết thúc phải từ ngày bắt đầu trở đi.');
      if (C.plus(from, 92) < to) throw Error('Mỗi đợt nghỉ tối đa 3 tháng.');
      commit(d => { if (h) Object.assign(d.holidays.find(x => x.id === hid), { name, from, to }); else d.holidays.push({ id: uid(), name, from, to }); });
      closeSheet(); toast('Đã lưu ngày nghỉ lễ cho tất cả nhóm.');
    });
    root.querySelector('[data-delete]')?.addEventListener('click', () => { closeSheet(); commitUndoable('Đã xóa ngày nghỉ lễ.', d => { d.holidays = d.holidays.filter(x => x.id !== hid); }); });
  });
}
function conflictSheet(remote) {
  const desc = d => { const n = d.groups.reduce((s, g) => s + g.students.length, 0); const t = d.updatedAt ? new Date(d.updatedAt) : null; return `${d.groups.length} nhóm · ${n} học sinh${t ? ` · sửa lúc ${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')} ${fmt(today(t))}` : ''}`; };
  let parsed; try { parsed = C.parseData(JSON.stringify(remote)); } catch { toast('Dữ liệu trên Drive bị lỗi. App giữ bản trên máy.'); sync.useLocal(); return; }
  openSheet('Chọn bản dữ liệu', `<p>Dữ liệu trên Google Drive khác với máy này. Bản không chọn vẫn được giữ trong thư mục Lịch sử trên Drive.</p>
<section class="options"><button class="option" data-c="remote"><span><b>Dùng bản trên Google Drive</b><small>${desc(parsed)}</small></span></button>
<button class="option" data-c="local"><span><b>Giữ bản trên máy này</b><small>${desc(data)}</small></span></button></section>`, root => {
    root.querySelector('.options').addEventListener('click', ev => {
      const b = ev.target.closest('[data-c]'); if (!b) return; closeSheet();
      if (b.dataset.c === 'remote') sync.useRemote(parsed); else sync.useLocal();
    });
  });
}

/* ================= Xuất, in, sao lưu ================= */
const slug = s => C.fold(s).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'nhom';
function deliver(bytes, name, type) {
  const file = new File([bytes], name, { type });
  if (/iPhone|iPad|iPod|Android/i.test(navigator.userAgent) && navigator.canShare?.({ files: [file] })) {
    navigator.share({ files: [file] }).catch(e => { if (e.name !== 'AbortError') download(file); });
    return;
  }
  download(file);
}
function download(file) {
  const url = URL.createObjectURL(file), a = document.createElement('a');
  a.href = url; a.download = file.name; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
function exportGroup(g) {
  const cal = calOf(g);
  deliver(buildXlsx([feeSheet(g, data.holidays, 'Học phí'), scheduleSheet(g, data.holidays, cal)]), `Hoc-phi-${slug(g.name)}-${today()}.xlsx`, XLSX_MIME);
}
function exportAll() {
  if (!data.groups.length) { toast('Chưa có nhóm học để xuất.'); return; }
  deliver(buildXlsx([...data.groups.map(g => feeSheet(g, data.holidays)), paymentsSheet(data, '', 'Tất cả')]), `So-lop-hoc-phi-${today()}.xlsx`, XLSX_MIME);
}
function printGroups(groups) {
  $('#print').innerHTML = printHtml(groups, data.holidays);
  setTimeout(() => window.print(), 50);
}
window.addEventListener('afterprint', () => { $('#print').innerHTML = ''; });
function exportJson() {
  const raw = storageError ? (localStorage.getItem(KEY) || '') : JSON.stringify(data, null, 1);
  deliver(new TextEncoder().encode(raw), `so-lop-sao-luu-${today()}.json`, 'application/json');
}
async function importJson(file) {
  try {
    if (!file) return;
    if (file.size > 10 * 1024 * 1024) throw Error('Tệp quá lớn.');
    const restored = C.parseData(await file.text());
    const n = restored.groups.reduce((s, g) => s + g.students.length, 0);
    openSheet('Khôi phục dữ liệu?', `<p>Tệp có <b>${restored.groups.length} nhóm</b>, <b>${n} học sinh</b>. Toàn bộ dữ liệu hiện tại sẽ được thay thế.</p><div class="pair"><button class="btn tonal" data-no>Hủy</button><button class="btn danger-fill" data-yes>Khôi phục</button></div>`, root => {
      root.querySelector('[data-no]').onclick = closeSheet;
      root.querySelector('[data-yes]').onclick = () => { closeSheet(); commit(d => { d.groups = restored.groups; d.holidays = restored.holidays; }, { restore: true }); ui.screen = 'groups'; render(); toast('Đã khôi phục dữ liệu.'); };
    });
  } catch (e) { toast('Không khôi phục được: ' + e.message); }
}

/* ================= Thông báo ================= */
let toastTimer;
function toast(text, { action, fn, sticky } = {}) {
  const el = $('#toast');
  el.innerHTML = `<span>${esc(text)}</span>${action ? `<button type="button">${esc(action)}</button>` : ''}`;
  if (action) el.querySelector('button').onclick = () => { el.classList.remove('show'); try { fn(); } catch (e) { toast(e.message); } };
  el.classList.add('show');
  clearTimeout(toastTimer);
  if (!sticky) toastTimer = setTimeout(() => el.classList.remove('show'), action ? 6500 : 3200);
}

/* ================= Sự kiện ================= */
function openGroup(id, extra = {}) {
  Object.assign(ui, { screen: 'group', groupId: id, tab: 'fees', feeMonth: null, filter: 'all', q: '', calMonth: null }, extra);
  if (history.state?.view !== 'group') history.pushState({ view: 'group' }, ''); else history.replaceState({ view: 'group' }, '');
  render();
}
app.addEventListener('click', ev => {
  const b = ev.target.closest('[data-action]'); if (!b || b.disabled) return;
  const a = b.dataset.action;
  try {
    switch (a) {
      case 'nav': if (history.state?.view === 'group') history.replaceState(null, ''); go(b.dataset.to); break;
      case 'home': if (history.state?.view === 'group') history.back(); else go('groups'); break;
      case 'settings': go('settings'); break;
      case 'new-group': groupSheet(false); break;
      case 'open-group': openGroup(b.dataset.id); break;
      case 'goto-student': { const s = group(b.dataset.g)?.students.find(x => x.id === b.dataset.s); openGroup(b.dataset.g, { q: s?.name || '' }); break; }
      case 'group-menu': groupMenu(); break;
      case 'tab': ui.tab = b.dataset.tab; render(); break;
      case 'fee-month': ui.feeMonth = Number(b.dataset.m); render(); break;
      case 'filter': ui.filter = b.dataset.f; render(); break;
      case 'toggle-view': ui.view = ui.view === 'list' ? 'table' : 'list'; try { localStorage.setItem('so-lop-view', ui.view); } catch { } render(); break;
      case 'show-overdue': { const g = group(), o = C.overdueList(g, calOf(g)); ui.feeMonth = o[0]?.month || ui.feeMonth; ui.filter = 'unpaid'; ui.q = ''; render(); break; }
      case 'new-student': studentSheet(null); break;
      case 'edit-student': studentSheet(b.dataset.id); break;
      case 'pay': paySheet(ui.groupId, b.dataset.id, Number(b.dataset.m)); break;
      case 'confirm-today': commit(d => { d.groups.find(x => x.id === ui.groupId).overrides[b.dataset.date] = 'done'; }); toast('Đã xác nhận buổi dạy hôm nay.'); break;
      case 'cal-prev': ui.calMonth--; render(); break;
      case 'cal-next': ui.calMonth++; render(); break;
      case 'lesson': lessonSheet(b.dataset.date); break;
      case 'add-extra': extraSheet(); break;
      case 'change-schedule': scheduleSheet2(); break;
      case 'stats-ym': ui.statsYm = b.dataset.ym; render(); break;
      case 'stats-method': ui.statsMethod = b.dataset.m; render(); break;
      case 'debt-group': ui.debtGroup = b.dataset.g; render(); break;
      case 'fix-fee': openGroup(b.dataset.g); groupSheet(true); break;
      case 'stats-pay': paySheet(b.dataset.g, b.dataset.s, Number(b.dataset.m)); break;
      case 'export-all': exportAll(); break;
      case 'print-all': printGroups(data.groups); break;
      case 'new-holiday': holidaySheet(null); break;
      case 'edit-holiday': holidaySheet(b.dataset.id); break;
      case 'sync-now': sync.pull(); break;
      case 'disconnect': sync.disconnect(); render(); toast('Đã ngắt kết nối Google Drive. Dữ liệu trên Drive vẫn còn.'); break;
      case 'resolve-conflict': sync.pull(); break;
      case 'export-json': exportJson(); break;
      case 'import-json': $('#import-file').click(); break;
    }
  } catch (e) { toast(e.message); }
});
app.addEventListener('input', ev => {
  const t = ev.target;
  if (t.id === 'q-home') { ui.homeQ = t.value; $('#home-body').innerHTML = homeBody(); }
  if (t.id === 'q-group') { ui.q = t.value; const g = group(); $('#fee-body').innerHTML = feeBody(g, calOf(g)); }
});
app.addEventListener('keydown', ev => { if (ev.key === 'Enter' && ev.target.type === 'search') ev.target.blur(); });
app.addEventListener('change', ev => { if (ev.target.id === 'import-file') { importJson(ev.target.files[0]); ev.target.value = ''; } });
app.addEventListener('submit', async ev => {
  if (ev.target.id !== 'drive-form') return;
  ev.preventDefault();
  const f = new FormData(ev.target), err = $('#drive-error'), btn = ev.target.querySelector('button[type=submit]');
  err.textContent = ''; btn.disabled = true; btn.textContent = 'Đang kết nối…';
  try { await sync.connect(f.get('url'), f.get('key')); render(); toast('Đã kết nối Google Drive.'); }
  catch (e) { err.textContent = e.name === 'AbortError' || e instanceof TypeError ? 'Không kết nối được. Kiểm tra mạng và địa chỉ ứng dụng web.' : e.message; btn.disabled = false; btn.textContent = 'Kết nối Google Drive'; }
});
window.addEventListener('popstate', () => { if (sheet.open) closeSheet(); if (ui.screen === 'group') go('groups'); });
document.addEventListener('visibilitychange', () => { if (document.hidden) { sync.flush(); return; } if (!sheet.open) render(); sync.pullIfStale(); });
window.addEventListener('online', () => sync.flush());
window.addEventListener('storage', ev => { if (ev.key === KEY && ev.newValue) { try { data = C.parseData(ev.newValue); if (!sheet.open) render(); } catch { } } });
let lastDay = today();
setInterval(() => { if (today() !== lastDay) { lastDay = today(); if (!sheet.open) render(); } }, 30000);

// Bàn phím iPhone: đẩy bảng thao tác lên trên bàn phím.
if (window.visualViewport) {
  const vv = window.visualViewport;
  const kb = () => document.documentElement.style.setProperty('--kb', Math.max(0, window.innerHeight - vv.height - vv.offsetTop) + 'px');
  vv.addEventListener('resize', kb); vv.addEventListener('scroll', kb);
}

// Cập nhật phiên bản app.
if ('serviceWorker' in navigator) {
  let wantReload = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (wantReload) location.reload(); });
  navigator.serviceWorker.register('./sw.js').then(reg => {
    const offer = w => toast('Đã có phiên bản mới của Sổ lớp.', { action: 'Cập nhật', sticky: true, fn: () => { wantReload = true; w.postMessage('skip'); } });
    if (reg.waiting && navigator.serviceWorker.controller) offer(reg.waiting);
    reg.addEventListener('updatefound', () => { const w = reg.installing; w?.addEventListener('statechange', () => { if (w.state === 'installed' && navigator.serviceWorker.controller) offer(w); }); });
    document.addEventListener('visibilitychange', () => { if (!document.hidden) reg.update().catch(() => {}); });
  }).catch(() => {});
}

try { ui.view = localStorage.getItem('so-lop-view') === 'table' ? 'table' : 'list'; } catch { }
render();
sync.pull();
