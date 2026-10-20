// ===== Google Apps Script — ألصقه في Extensions > Apps Script داخل الشيت =====
// كلمات السر هنا على السيرفر (مش ظاهرة في كود التطبيق). غيّرها من هنا.
const SUP_PASS = ['111'];
const OWNER_PASS = ['999', 'honor'];
const WIN_PERCENT = 0.20;
const TAB = 'Students';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(TAB);
  if (!s) {
    s = ss.insertSheet(TAB);
    s.appendRow(['id', 'name', 'phone', 'parent', 'area', 'marks', 'exam', 'created']);
    s.getRange('A:E').setNumberFormat('@');
  }
  return s;
}

function rows_() {
  const v = sheet_().getDataRange().getValues();
  v.shift();
  return v.filter(r => r[0] !== '').map((r, i) => ({
    row: i + 2,
    id: String(r[0]), name: r[1], phone: String(r[2]), parent: String(r[3]), area: r[4],
    marks: r[5] ? JSON.parse(r[5]) : [],
    exam: r[6] || 'none',
    created: r[7]
  }));
}

function pub_(s) {
  return { id: s.id, name: s.name, phone: s.phone, parent: s.parent, area: s.area, marks: s.marks, exam: s.exam };
}

function addStudent_(q) {
  const name = String(q.name || '').trim();
  if (!name) throw new Error('اكتب اسم الطالب');
  const all = rows_();
  const max = all.reduce((m, s) => Math.max(m, Number(s.id) || 0), 1000);
  const id = String(max + 1);
  sheet_().appendRow([id, name, String(q.phone || ''), String(q.parent || ''), String(q.area || '').trim(), '[]', 'none', new Date()]);
  return id;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return out_(handle_(JSON.parse(e.postData.contents)));
  } catch (err) {
    return out_({ error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out_({ ok: true }); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(q) {
  const role = OWNER_PASS.includes(q.pass) ? 'owner' : SUP_PASS.includes(q.pass) ? 'sup' : '';
  const sh = sheet_();
  const all = rows_();
  const find = id => all.find(s => s.id === String(id));

  switch (q.action) {
    case 'register': {
      return { id: addStudent_(q) };
    }
    case 'me': {
      const s = find(q.id);
      if (!s) throw new Error('الكود غير موجود');
      return { s: pub_(s) };
    }
    case 'mark': {
      const s = find(q.id);
      if (!s) throw new Error('الكود غير موجود');
      const n = Number(q.surah);
      let m = s.marks.filter(x => x !== n);
      if (q.on) m.push(n);
      sh.getRange(s.row, 6).setValue(JSON.stringify(m));
      return { ok: true };
    }
    case 'exam': {
      const s = find(q.id);
      if (!s) throw new Error('الكود غير موجود');
      if (s.exam === 'none') sh.getRange(s.row, 7).setValue('requested');
      return { ok: true };
    }
    case 'login': {
      if (!role) throw new Error('كلمة السر غلط');
      return { role };
    }
    case 'list': {
      if (!role) throw new Error('كلمة السر غلط');
      const out = { role, students: all.map(pub_) };
      if (role === 'owner') {
        const ok = all.filter(s => s.exam === 'approved').sort((a, b) => b.marks.length - a.marks.length);
        const n = ok.length ? Math.max(1, Math.ceil(ok.length * WIN_PERCENT)) : 0;
        out.winners = ok.slice(0, n).map(pub_);
        const areas = {};
        all.forEach(s => { const a = s.area || 'غير محدد'; areas[a] = (areas[a] || 0) + 1; });
        out.areas = areas;
      }
      return out;
    }
    case 'add': {
      if (!role) throw new Error('كلمة السر غلط');
      return { id: addStudent_(q) };
    }
    case 'approve': {
      if (!role) throw new Error('كلمة السر غلط');
      const s = find(q.id);
      if (!s) throw new Error('الكود غير موجود');
      sh.getRange(s.row, 7).setValue(q.value === 'approved' ? 'approved' : 'none');
      return { ok: true };
    }
    case 'reset': {
      if (role !== 'owner') throw new Error('للمالك فقط');
      all.forEach(s => { sh.getRange(s.row, 6, 1, 2).setValues([['[]', 'none']]); });
      return { ok: true };
    }
  }
  throw new Error('طلب غير معروف');
}
