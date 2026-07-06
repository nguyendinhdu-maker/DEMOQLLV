/*************************************************************************
 * QUẢN LÝ LUẬN VĂN — KHOA KỸ THUẬT CÔNG TRÌNH (LHU)
 * BACKEND API — Google Apps Script + Google Sheets + Google Drive
 *
 * Frontend (GitHub Pages / tên miền riêng) gọi API này qua fetch().
 * Triển khai: Deploy → Web app → Execute as: Me → Who has access: Anyone
 * Lần đầu: chạy hàm setup() để tạo sheet + tài khoản quản lý mặc định.
 *************************************************************************/

var SPREADSHEET_ID = '';   // để trống nếu script gắn liền Trang tính
var SALT = 'LHU-KTCT-2026';
var TOKEN_TTL = 21600;     // 6 giờ

var GRADER_ROLES = ['gvhd', 'gvpb', 'chutich', 'uyvien1', 'uyvien2'];
var ROLE_LABEL = {
  gvhd: 'Giảng viên hướng dẫn', gvpb: 'Giảng viên phản biện',
  chutich: 'Chủ tịch hội đồng', uyvien1: 'Ủy viên hội đồng 1', uyvien2: 'Ủy viên hội đồng 2'
};

/* ===================== HTTP ===================== */
function doGet() {
  return json_({ ok: true, service: 'QUANLYLUANVAN API', time: new Date().toISOString() });
}
function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    var action = req.action, data = req.data || {};
    var PUBLIC = { login: 1, register: 1 };
    var me = null;
    if (!PUBLIC[action]) {
      me = auth_(req.token);
      if (!me) return json_({ error: 'AUTH', message: 'Phiên đăng nhập hết hạn. Vui lòng đăng nhập lại.' });
    }
    var out;
    switch (action) {
      case 'login':            out = login_(data); break;
      case 'register':         out = register_(data); break;
      case 'bootstrap':        out = bootstrap_(me); break;
      case 'svRegisterThesis': out = svRegisterThesis_(me, data); break;
      case 'svUploadFile':     out = svUploadFile_(me, data); break;
      case 'svAddTienDo':      out = svAddTienDo_(me, data); break;
      case 'qlDuyet':          out = qlDuyet_(me, data); break;
      case 'qlPhanCong':       out = qlPhanCong_(me, data); break;
      case 'qlHoiDongMember':  out = qlHoiDongMember_(me, data); break;
      case 'qlSetSVHoiDong':   out = qlSetSVHoiDong_(me, data); break;
      case 'qlImportGV':       out = qlImportGV_(me, data); break;
      case 'qlUploadRubric':   out = qlUploadRubric_(me, data); break;
      case 'gvDiemDanh':       out = gvDiemDanh_(me, data); break;
      case 'gvGhiChuTienDo':   out = gvGhiChuTienDo_(me, data); break;
      case 'chamDiem':         out = chamDiem_(me, data); break;
      case 'luuNhanXet':       out = luuNhanXet_(me, data); break;
      case 'kySo':             out = kySo_(me, data); break;
      case 'exportPdf':        out = exportPdf_(me, data); break;
      default: out = { error: 'ACTION', message: 'Hành động không hợp lệ: ' + action };
    }
    return json_(out);
  } catch (err) {
    return json_({ error: 'SERVER', message: String(err && err.message || err) });
  }
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(sanitize_(obj)))
    .setMimeType(ContentService.MimeType.JSON);
}
function sanitize_(obj) {
  return JSON.parse(JSON.stringify(obj, function (k, v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'dd/MM/yyyy HH:mm');
    return v === undefined ? '' : v;
  }));
}
function tz_() { return Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh'; }

/* ===================== SHEET HELPERS ===================== */
function ss_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var a = SpreadsheetApp.getActiveSpreadsheet();
  if (a) return a;
  throw new Error('Chưa cấu hình SPREADSHEET_ID.');
}
function sheet_(name) { return ss_().getSheetByName(name) || ss_().insertSheet(name); }
function readTable_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var head = v[0], rows = [];
  for (var i = 1; i < v.length; i++) {
    if (String(v[i].join('')) === '') continue;
    var o = { _row: i + 1 };
    for (var c = 0; c < head.length; c++) o[head[c]] = v[i][c];
    rows.push(o);
  }
  return rows;
}
function updateRow_(name, keyCol, keyVal, patch) {
  var sh = sheet_(name), v = sh.getDataRange().getValues(), head = v[0];
  var ik = head.indexOf(keyCol);
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][ik]) === String(keyVal)) {
      Object.keys(patch).forEach(function (col) {
        var ic = head.indexOf(col);
        if (ic > -1) sh.getRange(i + 1, ic + 1).setValue(patch[col]);
      });
      return true;
    }
  }
  return false;
}
function appendObj_(name, obj) {
  var sh = sheet_(name), head = sh.getDataRange().getValues()[0];
  sh.appendRow(head.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }));
}
function upsert2_(name, k1c, k1v, k2c, k2v, patch) {
  var sh = sheet_(name), v = sh.getDataRange().getValues(), head = v[0];
  var i1 = head.indexOf(k1c), i2 = head.indexOf(k2c);
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][i1]) === String(k1v) && String(v[i][i2]) === String(k2v)) {
      Object.keys(patch).forEach(function (col) {
        var ic = head.indexOf(col); if (ic > -1) sh.getRange(i + 1, ic + 1).setValue(patch[col]);
      });
      return;
    }
  }
  var obj = {}; obj[k1c] = k1v; obj[k2c] = k2v; Object.keys(patch).forEach(function (c) { obj[c] = patch[c]; });
  appendObj_(name, obj);
}
function upsert3_(name, k1c, k1v, k2c, k2v, k3c, k3v, patch) {
  var sh = sheet_(name), v = sh.getDataRange().getValues(), head = v[0];
  var i1 = head.indexOf(k1c), i2 = head.indexOf(k2c), i3 = head.indexOf(k3c);
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][i1]) === String(k1v) && String(v[i][i2]) === String(k2v) && String(v[i][i3]) === String(k3v)) {
      Object.keys(patch).forEach(function (col) {
        var ic = head.indexOf(col); if (ic > -1) sh.getRange(i + 1, ic + 1).setValue(patch[col]);
      });
      return;
    }
  }
  var obj = {}; obj[k1c] = k1v; obj[k2c] = k2v; obj[k3c] = k3v;
  Object.keys(patch).forEach(function (c) { obj[c] = patch[c]; });
  appendObj_(name, obj);
}

/* ===================== AUTH ===================== */
function hash_(pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw + SALT)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function login_(d) {
  var email = String(d.email || '').trim().toLowerCase();
  var users = readTable_('NguoiDung');
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].Email).toLowerCase() === email) {
      if (String(users[i].MatKhauHash) !== hash_(String(d.matkhau || ''))) {
        return { error: 'LOGIN', message: 'Sai mật khẩu.' };
      }
      var token = Utilities.getUuid();
      CacheService.getScriptCache().put('tk_' + token, email, TOKEN_TTL);
      return { token: token, boot: bootstrap_(getUser_(email)) };
    }
  }
  return { error: 'LOGIN', message: 'Email chưa có tài khoản. Sinh viên hãy bấm "Tạo tài khoản".' };
}
function register_(d) {
  var email = String(d.email || '').trim().toLowerCase();
  if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { error: 'REG', message: 'Email không hợp lệ.' };
  if (String(d.matkhau || '').length < 6) return { error: 'REG', message: 'Mật khẩu tối thiểu 6 ký tự.' };
  if (!d.hoten || !d.mssv) return { error: 'REG', message: 'Thiếu họ tên hoặc MSSV.' };
  var users = readTable_('NguoiDung');
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].Email).toLowerCase() === email) return { error: 'REG', message: 'Email đã có tài khoản.' };
  }
  appendObj_('NguoiDung', { Email: email, MatKhauHash: hash_(String(d.matkhau)), HoTen: d.hoten, VaiTro: 'sinhvien', ChucDanh: 'Sinh viên' });
  appendObj_('SinhVien', { MaSV: d.mssv, HoTen: d.hoten, Lop: d.lop || '', Nganh: d.nganh || '', Email: email, TrangThai: 'chua_dang_ky' });
  return login_({ email: email, matkhau: d.matkhau });
}
function auth_(token) {
  if (!token) return null;
  var email = CacheService.getScriptCache().get('tk_' + token);
  if (!email) return null;
  CacheService.getScriptCache().put('tk_' + token, email, TOKEN_TTL); // gia hạn
  return getUser_(email);
}
function getUser_(email) {
  var users = readTable_('NguoiDung');
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].Email).toLowerCase() === String(email).toLowerCase()) {
      return { email: String(users[i].Email).toLowerCase(), name: users[i].HoTen, vaiTro: users[i].VaiTro, chucDanh: users[i].ChucDanh };
    }
  }
  return null;
}

/* ===================== PHÂN QUYỀN THEO PHÂN CÔNG ===================== */
// Trả về mọi "nhiệm vụ" của một giảng viên: hướng dẫn ai, phản biện ai, thuộc hội đồng nào
function rolesOf_(me) {
  var out = { sinhvien: null, quanly: me.vaiTro === 'quanly', gvhd: [], gvpb: [], hoidong: [] };
  var svs = readTable_('SinhVien');
  if (me.vaiTro === 'sinhvien') {
    for (var i = 0; i < svs.length; i++) if (String(svs[i].Email).toLowerCase() === me.email) out.sinhvien = svs[i].MaSV;
    return out;
  }
  svs.forEach(function (s) {
    if (String(s.EmailGVHD).toLowerCase() === me.email) out.gvhd.push(s.MaSV);
    if (String(s.EmailGVPB).toLowerCase() === me.email) out.gvpb.push(s.MaSV);
  });
  readTable_('HoiDong').forEach(function (h) {
    if (String(h.EmailGV).toLowerCase() === me.email) out.hoidong.push({ soHD: h.SoHoiDong, viTri: h.ViTri });
  });
  return out;
}
// Giảng viên này có được chấm SV này với vai trò vaiTroCham không?
function canGrade_(me, maSV, vaiTroCham) {
  var r = rolesOf_(me);
  if (vaiTroCham === 'gvhd') return r.gvhd.indexOf(maSV) > -1;
  if (vaiTroCham === 'gvpb') return r.gvpb.indexOf(maSV) > -1;
  var sv = readTable_('SinhVien').filter(function (s) { return String(s.MaSV) === String(maSV); })[0];
  if (!sv || !sv.SoHoiDong) return false;
  return r.hoidong.some(function (h) { return String(h.soHD) === String(sv.SoHoiDong) && h.viTri === vaiTroCham; });
}

/* ===================== BOOTSTRAP (dữ liệu theo vai trò) ===================== */
function bootstrap_(me) {
  var roles = rolesOf_(me);
  var cfg = {}; readTable_('CauHinh').forEach(function (r) { cfg[r.Khoa] = r.GiaTri; });
  var crits = readTable_('Rubric').sort(function (a, b) { return (a.STT || 0) - (b.STT || 0); })
    .map(function (r) { return { id: r.MaTieuChi, label: r.TieuChi, short: r.VietTat, max: Number(r.DiemToiDa) || 0 }; });
  var lich = readTable_('LichBieu').map(function (r) { return { moc: r.Moc, thoiGian: r.ThoiGian, moTa: r.MoTa }; });
  var users = readTable_('NguoiDung');
  var nameOf = {}; users.forEach(function (u) { nameOf[String(u.Email).toLowerCase()] = u.HoTen; });
  var lecturers = users.filter(function (u) { return u.VaiTro === 'giangvien' || u.VaiTro === 'quanly'; })
    .map(function (u) { return { email: String(u.Email).toLowerCase(), name: u.HoTen, chucDanh: u.ChucDanh }; });

  var svsAll = readTable_('SinhVien');
  function svView(s) {
    return {
      maSV: s.MaSV, hoTen: s.HoTen, lop: s.Lop, nganh: s.Nganh, email: String(s.Email).toLowerCase(),
      tenDeTai: s.TenDeTai, tomTat: s.TomTat, trangThai: s.TrangThai,
      emailGVHD: String(s.EmailGVHD || '').toLowerCase(), tenGVHD: nameOf[String(s.EmailGVHD || '').toLowerCase()] || '',
      emailGVPB: String(s.EmailGVPB || '').toLowerCase(), tenGVPB: nameOf[String(s.EmailGVPB || '').toLowerCase()] || '',
      soHoiDong: s.SoHoiDong, ngayBaoVe: s.NgayBaoVe, diaDiem: s.DiaDiem,
      fileLuanVan: s.FileLuanVan, fileName: s.FileName
    };
  }

  var visibleSV = [];
  if (roles.quanly) visibleSV = svsAll.map(svView);
  else if (me.vaiTro === 'sinhvien') visibleSV = svsAll.filter(function (s) { return String(s.Email).toLowerCase() === me.email; }).map(svView);
  else {
    var mine = {};
    roles.gvhd.concat(roles.gvpb).forEach(function (m) { mine[m] = 1; });
    var myHDs = roles.hoidong.map(function (h) { return String(h.soHD); });
    svsAll.forEach(function (s) {
      if (mine[s.MaSV] || (s.SoHoiDong && myHDs.indexOf(String(s.SoHoiDong)) > -1)) visibleSV.push(svView(s));
    });
  }
  var visibleIds = visibleSV.map(function (s) { return String(s.maSV); });

  function pick(table) {
    return readTable_(table).filter(function (r) { return visibleIds.indexOf(String(r.MaSV)) > -1; });
  }
  var scores = {}, comments = {}, sigs = {};
  pick('Diem').forEach(function (r) {
    scores[r.MaSV] = scores[r.MaSV] || {};
    scores[r.MaSV][r.VaiTroCham] = scores[r.MaSV][r.VaiTroCham] || {};
    scores[r.MaSV][r.VaiTroCham][r.MaTieuChi] = Number(r.Diem) || 0;
  });
  pick('NhanXet').forEach(function (r) {
    comments[r.MaSV + '|' + r.VaiTroCham] = r.NhanXet;
  });
  pick('ChuKy').forEach(function (r) {
    sigs[r.MaSV] = sigs[r.MaSV] || {};
    sigs[r.MaSV][r.VaiTroCham] = { name: r.HoTen, at: r.ThoiGian, serial: r.Serial };
  });
  var tienDo = pick('TienDo').map(function (r) { return { maSV: r.MaSV, ngay: r.Ngay, noiDung: r.NoiDung, ghiChuGV: r.GhiChuGV }; });
  var diemDanh = pick('DiemDanh').map(function (r) { return { maSV: r.MaSV, ngay: r.Ngay, coMat: r.CoMat, ghiChu: r.GhiChu }; });

  var hoiDong = readTable_('HoiDong').map(function (h) {
    return { soHD: h.SoHoiDong, viTri: h.ViTri, email: String(h.EmailGV).toLowerCase(), name: nameOf[String(h.EmailGV).toLowerCase()] || h.EmailGV };
  });

  return {
    me: { email: me.email, name: me.name, chucDanh: me.chucDanh, vaiTro: me.vaiTro, roles: roles },
    config: cfg, crits: crits, lich: lich,
    lecturers: roles.quanly ? lecturers : lecturers.map(function (l) { return { email: l.email, name: l.name }; }),
    students: visibleSV, hoiDong: hoiDong,
    scores: scores, comments: comments, signatures: sigs,
    tienDo: tienDo, diemDanh: diemDanh,
    roleLabels: ROLE_LABEL, graderRoles: GRADER_ROLES
  };
}

/* ===================== SINH VIÊN ===================== */
function mySV_(me) {
  var svs = readTable_('SinhVien');
  for (var i = 0; i < svs.length; i++) if (String(svs[i].Email).toLowerCase() === me.email) return svs[i];
  throw new Error('Không tìm thấy hồ sơ sinh viên.');
}
function svRegisterThesis_(me, d) {
  if (me.vaiTro !== 'sinhvien') throw new Error('Chỉ sinh viên được đăng ký luận văn.');
  var sv = mySV_(me);
  updateRow_('SinhVien', 'MaSV', sv.MaSV, {
    TenDeTai: d.tenDeTai || '', TomTat: d.tomTat || '', Lop: d.lop || sv.Lop, Nganh: d.nganh || sv.Nganh,
    TrangThai: 'cho_duyet'
  });
  return bootstrap_(me);
}
function svUploadFile_(me, d) {
  if (me.vaiTro !== 'sinhvien') throw new Error('Chỉ sinh viên được nộp file.');
  var sv = mySV_(me);
  var bytes = Utilities.base64Decode(d.base64);
  var blob = Utilities.newBlob(bytes, d.mimeType || 'application/pdf', d.fileName || 'luanvan.pdf');
  var folder = folder_('Luan van - File SV nop (KTCT)');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  updateRow_('SinhVien', 'MaSV', sv.MaSV, { FileLuanVan: file.getUrl(), FileName: d.fileName || file.getName() });
  return bootstrap_(me);
}
function svAddTienDo_(me, d) {
  if (me.vaiTro !== 'sinhvien') throw new Error('Chỉ sinh viên được ghi tiến độ.');
  var sv = mySV_(me);
  appendObj_('TienDo', { MaSV: sv.MaSV, Ngay: Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy'), NoiDung: d.noiDung || '', GhiChuGV: '' });
  return bootstrap_(me);
}

/* ===================== QUẢN LÝ / THƯ KÝ ===================== */
function requireQL_(me) { if (me.vaiTro !== 'quanly') throw new Error('Chỉ Người quản lý/Thư ký có quyền này.'); }
function qlDuyet_(me, d) { requireQL_(me); updateRow_('SinhVien', 'MaSV', d.maSV, { TrangThai: d.trangThai }); return bootstrap_(me); }
function qlPhanCong_(me, d) {
  requireQL_(me);
  var patch = {};
  if (d.emailGVHD !== undefined) patch.EmailGVHD = String(d.emailGVHD).toLowerCase();
  if (d.emailGVPB !== undefined) patch.EmailGVPB = String(d.emailGVPB).toLowerCase();
  if (d.ngayBaoVe !== undefined) patch.NgayBaoVe = d.ngayBaoVe;
  if (d.diaDiem !== undefined) patch.DiaDiem = d.diaDiem;
  updateRow_('SinhVien', 'MaSV', d.maSV, patch);
  return bootstrap_(me);
}
function qlHoiDongMember_(me, d) {
  requireQL_(me);
  upsert2_('HoiDong', 'SoHoiDong', d.soHD, 'ViTri', d.viTri, { EmailGV: String(d.emailGV || '').toLowerCase() });
  return bootstrap_(me);
}
function qlSetSVHoiDong_(me, d) { requireQL_(me); updateRow_('SinhVien', 'MaSV', d.maSV, { SoHoiDong: d.soHD }); return bootstrap_(me); }
function qlImportGV_(me, d) {
  requireQL_(me);
  var users = readTable_('NguoiDung');
  var existing = {}; users.forEach(function (u) { existing[String(u.Email).toLowerCase()] = true; });
  var added = 0;
  (d.rows || []).forEach(function (r) {
    var email = String(r.email || '').trim().toLowerCase();
    if (!email || existing[email]) return;
    appendObj_('NguoiDung', {
      Email: email, MatKhauHash: hash_(String(r.matkhau || '123456')),
      HoTen: r.hoten || email, VaiTro: 'giangvien', ChucDanh: r.chucdanh || 'Giảng viên'
    });
    existing[email] = true; added++;
  });
  var b = bootstrap_(me); b._imported = added; return b;
}
function qlUploadRubric_(me, d) {
  requireQL_(me);
  var sh = sheet_('Rubric'); sh.clear();
  sh.appendRow(['STT', 'MaTieuChi', 'TieuChi', 'VietTat', 'DiemToiDa']);
  (d.rows || []).forEach(function (r, i) { sh.appendRow([i + 1, 'c' + (i + 1), r.label, r.short || '', Number(r.max) || 0]); });
  ['Diem', 'ChuKy'].forEach(function (n) {
    var s = sheet_(n); s.clearContents();
    s.appendRow(n === 'Diem' ? ['MaSV', 'VaiTroCham', 'EmailNguoiCham', 'MaTieuChi', 'Diem'] : ['MaSV', 'VaiTroCham', 'EmailNguoiCham', 'HoTen', 'ThoiGian', 'Serial']);
  });
  return bootstrap_(me);
}

/* ===================== GIẢNG VIÊN ===================== */
function gvDiemDanh_(me, d) {
  if (!canGrade_(me, d.maSV, 'gvhd')) throw new Error('Bạn không phải GVHD của sinh viên này.');
  upsert2_('DiemDanh', 'MaSV', d.maSV, 'Ngay', d.ngay, { CoMat: d.coMat ? 'x' : '', GhiChu: d.ghiChu || '' });
  return bootstrap_(me);
}
function gvGhiChuTienDo_(me, d) {
  if (!canGrade_(me, d.maSV, 'gvhd')) throw new Error('Bạn không phải GVHD của sinh viên này.');
  var sh = sheet_('TienDo'), v = sh.getDataRange().getValues(), head = v[0];
  var iMa = head.indexOf('MaSV'), iNg = head.indexOf('Ngay'), iNd = head.indexOf('NoiDung'), iGc = head.indexOf('GhiChuGV');
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][iMa]) === String(d.maSV) && String(v[i][iNg]) === String(d.ngay) && String(v[i][iNd]) === String(d.noiDung)) {
      sh.getRange(i + 1, iGc + 1).setValue(d.ghiChu || ''); return bootstrap_(me);
    }
  }
  throw new Error('Không tìm thấy mục tiến độ.');
}
function chamDiem_(me, d) {
  if (!canGrade_(me, d.maSV, d.vaiTroCham)) throw new Error('Bạn không được phân công chấm sinh viên này với vai trò ' + (ROLE_LABEL[d.vaiTroCham] || d.vaiTroCham) + '.');
  if (signed_(d.maSV, d.vaiTroCham)) throw new Error('Phiếu đã ký số, không thể sửa điểm.');
  var crit = readTable_('Rubric').filter(function (r) { return r.MaTieuChi === d.critId; })[0];
  if (!crit) throw new Error('Tiêu chí không hợp lệ.');
  var v = Math.max(0, Math.min(Number(crit.DiemToiDa), Number(d.diem) || 0));
  upsert3_('Diem', 'MaSV', d.maSV, 'VaiTroCham', d.vaiTroCham, 'MaTieuChi', d.critId, { EmailNguoiCham: me.email, Diem: v });
  return { ok: true, diem: v };
}
function luuNhanXet_(me, d) {
  if (!canGrade_(me, d.maSV, d.vaiTroCham)) throw new Error('Bạn không được phân công nhận xét sinh viên này.');
  if (signed_(d.maSV, d.vaiTroCham)) throw new Error('Phiếu đã ký số, không thể sửa nhận xét.');
  upsert2_('NhanXet', 'MaSV', d.maSV, 'VaiTroCham', d.vaiTroCham, { EmailNguoiCham: me.email, NhanXet: d.text || '' });
  return { ok: true };
}
function signed_(maSV, vaiTroCham) {
  return readTable_('ChuKy').some(function (r) { return String(r.MaSV) === String(maSV) && r.VaiTroCham === vaiTroCham; });
}
function kySo_(me, d) {
  if (!canGrade_(me, d.maSV, d.vaiTroCham)) throw new Error('Bạn không được phân công vai trò này.');
  var at = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm:ss');
  var serial = 'LHU-' + hash_(me.email + d.maSV + d.vaiTroCham).slice(0, 10).toUpperCase();
  upsert2_('ChuKy', 'MaSV', d.maSV, 'VaiTroCham', d.vaiTroCham, { EmailNguoiCham: me.email, HoTen: me.name, ThoiGian: at, Serial: serial });
  return bootstrap_(me);
}

/* ===================== XUẤT PDF ===================== */
function folder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function exportPdf_(me, d) {
  var b = bootstrap_(me);
  var sv = b.students.filter(function (s) { return String(s.maSV) === String(d.maSV); })[0];
  if (!sv) throw new Error('Bạn không có quyền với sinh viên này.');
  var loai = d.loai; // 'gvhd' | 'gvpb' | 'hoidong' | 'tonghop'
  var html = (loai === 'tonghop') ? pdfTongHop_(b, sv) : pdfPhieu_(b, sv, loai);
  var pdf = Utilities.newBlob(html, 'text/html', 'x.html').getAs('application/pdf')
    .setName(('Phieu_' + loai + '_' + sv.maSV + '_' + sv.hoTen + '.pdf').replace(/\s+/g, '_'));
  var file = folder_('Luan van - Phieu & Bien ban (KTCT)').createFile(pdf);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl(), name: file.getName() };
}
function headerHtml_() {
  return '<div style="display:flex;justify-content:space-between;border-bottom:2px solid #000;padding-bottom:10px">'
    + '<div style="text-align:center;flex:1;font-size:12px;line-height:1.5"><b>TRƯỜNG ĐẠI HỌC LẠC HỒNG</b><br><b>KHOA KỸ THUẬT CÔNG TRÌNH</b></div>'
    + '<div style="text-align:center;flex:1;font-size:12px;line-height:1.5"><b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br><b>Độc lập – Tự do – Hạnh phúc</b></div></div>';
}
function svInfoHtml_(sv) {
  return '<div style="font-size:13px;line-height:1.9;margin:16px 0">'
    + '<div><b>Họ và tên sinh viên:</b> ' + sv.hoTen + ' &nbsp;&nbsp; <b>MSSV:</b> ' + sv.maSV + ' &nbsp;&nbsp; <b>Lớp:</b> ' + (sv.lop || '') + '</div>'
    + '<div><b>Ngành:</b> ' + (sv.nganh || '') + '</div>'
    + '<div><b>Tên đề tài:</b> ' + (sv.tenDeTai || '') + '</div>'
    + '<div><b>GVHD:</b> ' + (sv.tenGVHD || '—') + ' &nbsp;&nbsp; <b>GVPB:</b> ' + (sv.tenGVPB || '—') + ' &nbsp;&nbsp; <b>Hội đồng số:</b> ' + (sv.soHoiDong || '—') + '</div>'
    + (sv.ngayBaoVe ? '<div><b>Ngày bảo vệ:</b> ' + sv.ngayBaoVe + ' &nbsp;&nbsp; <b>Địa điểm:</b> ' + (sv.diaDiem || '') + '</div>' : '')
    + '</div>';
}
function sigBoxHtml_(sig, roleLabel, name) {
  var inner = sig
    ? '<div style="border:1px solid #1f8a5b;border-radius:6px;padding:6px 12px;background:#f0faf4;display:inline-block;text-align:left"><div style="font-size:10px;color:#1f8a5b;font-weight:bold">✓ ĐÃ KÝ SỐ</div><div style="font-size:11px;font-weight:bold">' + sig.name + '</div><div style="font-size:9px;color:#555">' + sig.at + '</div><div style="font-size:9px;color:#555">SN: ' + sig.serial + '</div></div>'
    : '<div style="font-style:italic;color:#999;font-size:11px">(chưa ký)</div>';
  return '<div style="text-align:center;font-size:12px"><div style="font-weight:bold">' + roleLabel + '</div>'
    + '<div style="height:70px;display:flex;align-items:center;justify-content:center">' + inner + '</div>'
    + '<div style="font-weight:bold">' + (name || '') + '</div></div>';
}
// Phiếu nhận xét + điểm của 1 nhóm (GVHD / GVPB / cả hội đồng)
function pdfPhieu_(b, sv, loai) {
  var roles = loai === 'gvhd' ? ['gvhd'] : loai === 'gvpb' ? ['gvpb'] : ['chutich', 'uyvien1', 'uyvien2'];
  var title = loai === 'gvhd' ? 'PHIẾU NHẬN XÉT VÀ ĐÁNH GIÁ CỦA GIẢNG VIÊN HƯỚNG DẪN'
    : loai === 'gvpb' ? 'PHIẾU NHẬN XÉT VÀ ĐÁNH GIÁ CỦA GIẢNG VIÊN PHẢN BIỆN'
    : 'PHIẾU NHẬN XÉT VÀ CHẤM ĐIỂM CỦA HỘI ĐỒNG ĐÁNH GIÁ';
  function nameOfRole(rk) {
    if (rk === 'gvhd') return sv.tenGVHD;
    if (rk === 'gvpb') return sv.tenGVPB;
    var m = b.hoiDong.filter(function (h) { return String(h.soHD) === String(sv.soHoiDong) && h.viTri === rk; })[0];
    return m ? m.name : '';
  }
  var body = roles.map(function (rk) {
    var sc = (b.scores[sv.maSV] || {})[rk] || {};
    var rows = b.crits.map(function (c, i) {
      return '<tr><td style="border:1px solid #000;padding:6px;text-align:center">' + (i + 1) + '</td>'
        + '<td style="border:1px solid #000;padding:6px">' + c.label + '</td>'
        + '<td style="border:1px solid #000;padding:6px;text-align:center">' + c.max.toFixed(1) + '</td>'
        + '<td style="border:1px solid #000;padding:6px;text-align:center;font-weight:bold">' + (Number(sc[c.id]) || 0).toFixed(2) + '</td></tr>';
    }).join('');
    var tot = b.crits.reduce(function (s, c) { return s + (Number(sc[c.id]) || 0); }, 0);
    var cm = b.comments[sv.maSV + '|' + rk] || '';
    var sig = (b.signatures[sv.maSV] || {})[rk];
    return '<div style="margin-top:22px"><div style="font-size:14px;font-weight:bold;border-left:4px solid #16345f;padding-left:8px">' + b.roleLabels[rk] + ': ' + (nameOfRole(rk) || '—') + '</div>'
      + '<table style="width:100%;border-collapse:collapse;font-size:12.5px;margin:10px 0"><tr style="background:#eee"><th style="border:1px solid #000;padding:6px;width:32px">TT</th><th style="border:1px solid #000;padding:6px">Tiêu chí đánh giá</th><th style="border:1px solid #000;padding:6px;width:70px">Tối đa</th><th style="border:1px solid #000;padding:6px;width:70px">Điểm</th></tr>'
      + rows
      + '<tr style="background:#f3f3f3"><td colspan="3" style="border:1px solid #000;padding:6px;text-align:right;font-weight:bold">TỔNG ĐIỂM</td><td style="border:1px solid #000;padding:6px;text-align:center;font-weight:bold">' + tot.toFixed(2) + '</td></tr></table>'
      + '<div style="font-size:13px"><b>Nhận xét và đề nghị:</b><div style="border:1px solid #999;border-radius:4px;min-height:56px;padding:8px;margin-top:4px;line-height:1.6">' + (cm || '<i style="color:#999">(chưa có nhận xét)</i>') + '</div></div>'
      + '<div style="display:flex;justify-content:flex-end;margin-top:14px">' + sigBoxHtml_(sig, b.roleLabels[rk], nameOfRole(rk)) + '</div></div>';
  }).join('');
  return '<div style="font-family:\'Times New Roman\',serif;padding:28px;color:#000">' + headerHtml_()
    + '<div style="text-align:center;margin:22px 0 4px"><div style="font-size:17px;font-weight:bold">' + title + '</div>'
    + '<div style="font-style:italic;font-size:12px">Năm học ' + (b.config.NamHoc || '') + '</div></div>'
    + svInfoHtml_(sv) + body + '</div>';
}
// Biên bản tổng hợp 5 điểm
function pdfTongHop_(b, sv) {
  var rows = b.graderRoles.map(function (rk, i) {
    var sc = (b.scores[sv.maSV] || {})[rk] || {};
    var tot = b.crits.reduce(function (s, c) { return s + (Number(sc[c.id]) || 0); }, 0);
    var nm = rk === 'gvhd' ? sv.tenGVHD : rk === 'gvpb' ? sv.tenGVPB
      : (b.hoiDong.filter(function (h) { return String(h.soHD) === String(sv.soHoiDong) && h.viTri === rk; })[0] || {}).name;
    return { rk: rk, nm: nm || '—', tot: tot };
  });
  var avg = rows.reduce(function (s, r) { return s + r.tot; }, 0) / rows.length;
  var xl = avg >= 9 ? 'Xuất sắc' : avg >= 8 ? 'Giỏi' : avg >= 7 ? 'Khá' : avg >= 5.5 ? 'Trung bình' : 'Không đạt';
  var trs = rows.map(function (r, i) {
    return '<tr><td style="border:1px solid #000;padding:7px;text-align:center">' + (i + 1) + '</td><td style="border:1px solid #000;padding:7px">' + r.nm + '</td><td style="border:1px solid #000;padding:7px;text-align:center">' + b.roleLabels[r.rk] + '</td><td style="border:1px solid #000;padding:7px;text-align:center;font-weight:bold">' + r.tot.toFixed(2) + '</td></tr>';
  }).join('');
  var sigs = rows.map(function (r) { return sigBoxHtml_((b.signatures[sv.maSV] || {})[r.rk], b.roleLabels[r.rk], r.nm); }).join('');
  return '<div style="font-family:\'Times New Roman\',serif;padding:28px;color:#000">' + headerHtml_()
    + '<div style="text-align:center;margin:24px 0 4px"><div style="font-size:18px;font-weight:bold">PHIẾU TỔNG HỢP ĐIỂM ĐÁNH GIÁ LUẬN VĂN</div>'
    + '<div style="font-style:italic;font-size:12px">Hội đồng số ' + (sv.soHoiDong || '—') + ' · Năm học ' + (b.config.NamHoc || '') + '</div></div>'
    + svInfoHtml_(sv)
    + '<table style="width:100%;border-collapse:collapse;font-size:13px;margin:8px 0"><tr style="background:#eee"><th style="border:1px solid #000;padding:7px;width:36px">TT</th><th style="border:1px solid #000;padding:7px">Người chấm</th><th style="border:1px solid #000;padding:7px">Vai trò</th><th style="border:1px solid #000;padding:7px;width:70px">Điểm</th></tr>'
    + trs + '<tr style="background:#f3f3f3"><td colspan="3" style="border:1px solid #000;padding:7px;text-align:right;font-weight:bold">ĐIỂM TRUNG BÌNH</td><td style="border:1px solid #000;padding:7px;text-align:center;font-weight:bold">' + avg.toFixed(2) + '</td></tr></table>'
    + '<div style="font-size:13px;margin:8px 0"><b>Xếp loại:</b> ' + xl + '</div>'
    + '<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px 30px;margin-top:26px">' + sigs + '</div></div>';
}

/* ===================== SETUP (chạy 1 lần) ===================== */
function setup() {
  var book = ss_();
  function reset(name, head, rows) {
    var sh = book.getSheetByName(name) || book.insertSheet(name);
    sh.clear(); sh.appendRow(head);
    (rows || []).forEach(function (r) { sh.appendRow(r); });
    sh.setFrozenRows(1);
  }
  reset('CauHinh', ['Khoa', 'GiaTri'], [
    ['TenTruong', 'TRƯỜNG ĐẠI HỌC LẠC HỒNG'], ['TenKhoa', 'KHOA KỸ THUẬT CÔNG TRÌNH'],
    ['NamHoc', '2025 – 2026'], ['HocKy', 'Học kỳ 2'], ['Email', 'ktct@lhu.edu.vn']
  ]);
  reset('NguoiDung', ['Email', 'MatKhauHash', 'HoTen', 'VaiTro', 'ChucDanh'], [
    ['quanly@lhu.edu.vn', hash_('ktct@2026'), 'Nguyễn Thị Hà', 'quanly', 'Thư ký Khoa'],
    ['son.do@lhu.edu.vn', hash_('123456'), 'PGS.TS. Đỗ Văn Sơn', 'giangvien', 'Trưởng Bộ môn Kết cấu'],
    ['hong.le@lhu.edu.vn', hash_('123456'), 'TS. Lê Thị Hồng', 'giangvien', 'Giảng viên'],
    ['bao.pham@lhu.edu.vn', hash_('123456'), 'ThS. Phạm Quốc Bảo', 'giangvien', 'Giảng viên'],
    ['hoang.tran@lhu.edu.vn', hash_('123456'), 'TS. Trần Minh Hoàng', 'giangvien', 'Giảng viên'],
    ['tung.nguyen@lhu.edu.vn', hash_('123456'), 'TS. Nguyễn Thanh Tùng', 'giangvien', 'Giảng viên'],
    ['an.nguyen@lhu.edu.vn', hash_('123456'), 'Nguyễn Văn An', 'sinhvien', 'Sinh viên']
  ]);
  reset('SinhVien', ['MaSV', 'HoTen', 'Lop', 'Nganh', 'Email', 'TenDeTai', 'TomTat', 'TrangThai', 'EmailGVHD', 'EmailGVPB', 'SoHoiDong', 'NgayBaoVe', 'DiaDiem', 'FileLuanVan', 'FileName'], [
    ['2151010023', 'Nguyễn Văn An', '21XD111', 'Kỹ thuật xây dựng', 'an.nguyen@lhu.edu.vn',
      'Phân tích ứng xử kết cấu nhà cao tầng chịu tải trọng động đất',
      'Mô phỏng kết cấu nhà cao tầng chịu động đất bằng ETABS...', 'da_duyet',
      'hoang.tran@lhu.edu.vn', 'hong.le@lhu.edu.vn', '3', '28/06/2026', 'Phòng C2.05', '', '']
  ]);
  reset('HoiDong', ['SoHoiDong', 'ViTri', 'EmailGV'], [
    ['3', 'chutich', 'son.do@lhu.edu.vn'],
    ['3', 'uyvien1', 'bao.pham@lhu.edu.vn'],
    ['3', 'uyvien2', 'tung.nguyen@lhu.edu.vn']
  ]);
  reset('Rubric', ['STT', 'MaTieuChi', 'TieuChi', 'VietTat', 'DiemToiDa'], [
    [1, 'c1', 'Hình thức trình bày Thuyết minh', 'Thuyết minh', 1.5],
    [2, 'c2', 'Hình thức trình bày Bản vẽ', 'Bản vẽ', 1.5],
    [3, 'c3', 'Khối lượng tính toán của đề tài', 'Tính toán', 3.0],
    [4, 'c4', 'Thuyết trình Luận văn', 'Thuyết trình', 2.0],
    [5, 'c5', 'Trả lời câu hỏi của Hội đồng', 'Vấn đáp', 2.0]
  ]);
  reset('Diem', ['MaSV', 'VaiTroCham', 'EmailNguoiCham', 'MaTieuChi', 'Diem']);
  reset('NhanXet', ['MaSV', 'VaiTroCham', 'EmailNguoiCham', 'NhanXet']);
  reset('ChuKy', ['MaSV', 'VaiTroCham', 'EmailNguoiCham', 'HoTen', 'ThoiGian', 'Serial']);
  reset('TienDo', ['MaSV', 'Ngay', 'NoiDung', 'GhiChuGV'], [
    ['2151010023', '10/05/2026', 'Hoàn thành chương 1: Tổng quan', 'Đạt yêu cầu, tiếp tục chương 2'],
    ['2151010023', '05/06/2026', 'Hoàn thành mô hình ETABS + chương 3', '']
  ]);
  reset('DiemDanh', ['MaSV', 'Ngay', 'CoMat', 'GhiChu'], [
    ['2151010023', '10/05/2026', 'x', 'Duyệt bài lần 1'],
    ['2151010023', '05/06/2026', 'x', 'Duyệt bài lần 2']
  ]);
  reset('LichBieu', ['Moc', 'ThoiGian', 'MoTa'], [
    ['Đăng ký đề tài', '01/03 – 15/03/2026', 'Sinh viên đăng ký tên đề tài và thông tin trên hệ thống'],
    ['Phân công GVHD', '16/03 – 22/03/2026', 'Khoa duyệt và phân công giảng viên hướng dẫn'],
    ['Thực hiện luận văn', '23/03 – 05/06/2026', 'SV thực hiện, gặp GVHD duyệt bài định kỳ'],
    ['Nộp file luận văn', '06/06 – 12/06/2026', 'Nộp bản PDF thuyết minh + bản vẽ lên hệ thống'],
    ['Phản biện', '13/06 – 20/06/2026', 'GV phản biện đọc, nhận xét và chấm điểm'],
    ['Bảo vệ trước hội đồng', '25/06 – 30/06/2026', 'Theo lịch từng hội đồng, xem thông báo của Khoa']
  ]);
}
