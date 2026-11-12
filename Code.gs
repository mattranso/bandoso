var SPREADSHEET_ID = '1zNJKKs9bKOoNzcnrQIiHBb02ehU25anGPmlam0cBUPc';

function getSS_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) throw new Error("Lỗi ID Google Sheets.");
  return ss;
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Bản Đồ Số Mặt Trận – Cộng Đồng')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function initializeSystem() {
  try {
    var ss = getSS_();
    setupSheet_(ss, 'USERS', ['UserID', 'Username', 'PasswordHash', 'HoTen', 'VaiTro', 'Cap', 'UnitID', 'XaPhuongID', 'XaPhuong', 'ThonApKhuPhoID', 'ThonApKhuPho', 'SoDienThoai', 'Email', 'TrangThai', 'NgayTao', 'LanDangNhapCuoi']);
    setupSheet_(ss, 'LOCATIONS', ['LocationID', 'MaDiem', 'TenDiem', 'NhomDuLieu', 'LoaiDiem', 'MoTa', 'DiaChi', 'XaPhuongID', 'XaPhuong', 'ThonApKhuPhoID', 'ThonApKhuPho', 'ViDo', 'KinhDo', 'NguoiTao', 'UserID', 'DonVi', 'NgayTao', 'NguoiCapNhat', 'NgayCapNhat', 'TrangThai', 'TrangThaiXacMinh', 'TrangThaiCongKhai', 'GhiChu']);
    setupSheet_(ss, 'SYSTEM_LOG', ['LogID', 'UserID', 'Username', 'HoTen', 'ThaoTac', 'DoiTuong', 'DoiTuongID', 'ThoiGian', 'IP', 'GhiChu']);
    
    var catSheet = setupSheet_(ss, 'CATEGORIES', ['CategoryID', 'NhomDuLieu', 'LoaiDiem', 'MoTa', 'Icon', 'TrangThai', 'ThuTu']);
    if (catSheet.getLastRow() <= 1) seedCategories_(catSheet);

    var unitSheet = setupSheet_(ss, 'UNITS', ['UnitID', 'TenDonVi', 'Cap', 'XaPhuongID', 'XaPhuong', 'ThonApKhuPhoID', 'ThonApKhuPho', 'DiaChi', 'NguoiPhuTrach', 'DienThoai', 'Email', 'TrangThai']);
    // Tự động bơm 95 đơn vị nếu bảng UNITS đang trống hoặc chỉ có 1 dòng
    if (unitSheet.getLastRow() <= 2) seedUnits_(unitSheet);

    return { success: true, message: "Khởi tạo bảng và nạp tự động 95 đơn vị thành công!" };
  } catch (error) { return { success: false, message: error.message }; }
}

function setupSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); sheet.appendRow(headers); sheet.setFrozenRows(1); }
  return sheet;
}

function setupLeNguyen() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName('USERS');
    if (!sheet) throw new Error("Chưa có bảng USERS.");

    var data = sheet.getDataRange().getValues();
    var lenguyenExists = false;
    var passHash = hashPassword_('Mttq@2026');

    for (var i = data.length - 1; i >= 1; i--) {
      var username = data[i][1] ? data[i][1].toString().trim().toLowerCase() : '';
      if (username === 'admin') {
        sheet.deleteRow(i + 1);
      } else if (username === 'lenguyen') {
        lenguyenExists = true;
        sheet.getRange(i + 1, 3).setValue(passHash);
        sheet.getRange(i + 1, 4).setValue('Lê Văn Nguyên');
        sheet.getRange(i + 1, 5).setValue('ADMIN');
        sheet.getRange(i + 1, 14).setValue('HOAT_DONG');
      }
    }

    if (!lenguyenExists) {
      var userId = 'USER-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheet.appendRow([
        userId, 'lenguyen', passHash, 'Lê Văn Nguyên', 'ADMIN', 'THANH_PHO',
        'MTTQ-TP', 'TP_DONG_NAI', 'Thành phố Đồng Nai', '', '',
        '0944119644', 'lenguyen041291@gmail.com', 'HOAT_DONG', new Date(), new Date()
      ]);
    }
    return { success: true, message: "Đã thiết lập tài khoản lenguyen thành công!" };
  } catch (error) { return { success: false, message: error.message }; }
}

function seedCategories_(sheet) {
  var categories = [
    ['CAT-001', 'AN SINH XÃ HỘI', 'Điểm hỗ trợ an sinh', 'Nơi cung cấp lương thực, thực phẩm', 'home', 'HOAT_DONG', 1],
    ['CAT-002', 'AN SINH XÃ HỘI', 'Nhà Đại đoàn kết', 'Nhà đại đoàn kết hỗ trợ hộ nghèo', 'home', 'HOAT_DONG', 2],
    ['CAT-003', 'AN SINH XÃ HỘI', 'Địa chỉ nhân đạo', 'Cơ sở nuôi dưỡng trẻ mồ côi, người già', 'favorite', 'HOAT_DONG', 3],
    ['CAT-004', 'CÔNG TRÌNH – CỘNG ĐỒNG', 'Công trình Mặt trận', 'Công trình chào mừng Đại hội', 'apartment', 'HOAT_DONG', 4],
    ['CAT-005', 'CÔNG TRÌNH – CỘNG ĐỒNG', 'Công trình dân sinh', 'Đường giao thông nông thôn', 'bridge', 'HOAT_DONG', 5],
    ['CAT-010', 'MẶT TRẬN – ĐOÀN THỂ', 'Ban Công tác Mặt trận', 'Văn phòng Ban CTMT', 'groups', 'HOAT_DONG', 10],
    ['CAT-016', 'MÔ HÌNH CỘNG ĐỒNG', 'Mô hình tự quản', 'Mô hình an ninh trật tự', 'shield', 'HOAT_DONG', 16]
  ];
  categories.forEach(function(row) { sheet.appendRow(row); });
}

function seedUnits_(sheet) {
  // 95 Đơn vị của anh đã được tích hợp sẵn tại đây
  var units = [
    ['MTTQ-TP', 'Ủy ban MTTQ Việt Nam thành phố Đồng Nai', 'THANH_PHO', 'TP_DONG_NAI', 'Thành phố Đồng Nai', '', '', 'Trung tâm hành chính tỉnh Đồng Nai', 'Quản trị viên', '02513822111', 'mttq.dongnai@gov.vn', 'HOAT_DONG'],
    ['UNIT-001', 'Ủy ban MTTQ Việt Nam Phường An Lộc', 'XA_PHUONG', 'XA-AN-LOC', 'Phường An Lộc', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-002', 'Ủy ban MTTQ Việt Nam Xã An Phước', 'XA_PHUONG', 'XA-AN-PHUOC', 'Xã An Phước', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-003', 'Ủy ban MTTQ Việt Nam Xã An Viễn', 'XA_PHUONG', 'XA-AN-VIEN', 'Xã An Viễn', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-004', 'Ủy ban MTTQ Việt Nam Phường Bảo Vinh', 'XA_PHUONG', 'XA-BAO-VINH', 'Phường Bảo Vinh', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-005', 'Ủy ban MTTQ Việt Nam Xã Bàu Hàm', 'XA_PHUONG', 'XA-BAU-HAM', 'Xã Bàu Hàm', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-006', 'Ủy ban MTTQ Việt Nam Phường Biên Hoà', 'XA_PHUONG', 'XA-BIEN-HOA', 'Phường Biên Hoà', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-007', 'Ủy ban MTTQ Việt Nam Xã Bình An', 'XA_PHUONG', 'XA-BINH-AN', 'Xã Bình An', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-008', 'Ủy ban MTTQ Việt Nam Phường Bình Lộc', 'XA_PHUONG', 'XA-BINH-LOC', 'Phường Bình Lộc', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-009', 'Ủy ban MTTQ Việt Nam Phường Bình Long', 'XA_PHUONG', 'XA-BINH-LONG', 'Phường Bình Long', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-010', 'Ủy ban MTTQ Việt Nam Xã Bình Minh', 'XA_PHUONG', 'XA-BINH-MINH', 'Xã Bình Minh', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-011', 'Ủy ban MTTQ Việt Nam Phường Bình Phước', 'XA_PHUONG', 'XA-BINH-PHUOC', 'Phường Bình Phước', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-012', 'Ủy ban MTTQ Việt Nam Xã Bình Tân', 'XA_PHUONG', 'XA-BINH-TAN', 'Xã Bình Tân', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-013', 'Ủy ban MTTQ Việt Nam Xã Bom Bo', 'XA_PHUONG', 'XA-BOM-BO', 'Xã Bom Bo', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-014', 'Ủy ban MTTQ Việt Nam Xã Bù Đăng', 'XA_PHUONG', 'XA-BU-DANG', 'Xã Bù Đăng', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-015', 'Ủy ban MTTQ Việt Nam Xã Cẩm Mỹ', 'XA_PHUONG', 'XA-CAM-MY', 'Xã Cẩm Mỹ', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-016', 'Ủy ban MTTQ Việt Nam Xã Nam Cát Tiên', 'XA_PHUONG', 'XA-NAM-CAT-TIEN', 'Xã Nam Cát Tiên', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-017', 'Ủy ban MTTQ Việt Nam Phường Chơn Thành', 'XA_PHUONG', 'XA-CHON-THANH', 'Phường Chơn Thành', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-018', 'Ủy ban MTTQ Việt Nam Xã Đa Kia', 'XA_PHUONG', 'XA-DA-KIA', 'Xã Đa Kia', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-019', 'Ủy ban MTTQ Việt Nam Xã Đại Phước', 'XA_PHUONG', 'XA-DAI-PHUOC', 'Xã Đại Phước', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-020', 'Ủy ban MTTQ Việt Nam Xã Đak Lua', 'XA_PHUONG', 'XA-DAK-LUA', 'Xã Đak Lua', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-021', 'Ủy ban MTTQ Việt Nam Xã Đak Nhau', 'XA_PHUONG', 'XA-DAK-NHAU', 'Xã Đak Nhau', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-022', 'Ủy ban MTTQ Việt Nam Xã Đăk Ơ', 'XA_PHUONG', 'XA-DAK-O', 'Xã Đăk Ơ', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-023', 'Ủy ban MTTQ Việt Nam Phường Dầu Giây', 'XA_PHUONG', 'XA-DAU-GIAY', 'Phường Dầu Giây', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-024', 'Ủy ban MTTQ Việt Nam Xã Định Quán', 'XA_PHUONG', 'XA-DINH-QUAN', 'Xã Định Quán', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-025', 'Ủy ban MTTQ Việt Nam Phường Đồng Phú', 'XA_PHUONG', 'XA-DONG-PHU', 'Phường Đồng Phú', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-026', 'Ủy ban MTTQ Việt Nam Xã Đồng Tâm', 'XA_PHUONG', 'XA-DONG-TAM', 'Xã Đồng Tâm', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-027', 'Ủy ban MTTQ Việt Nam Phường Đồng Xoài', 'XA_PHUONG', 'XA-DONG-XOAI', 'Phường Đồng Xoài', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-028', 'Ủy ban MTTQ Việt Nam Xã Gia Kiệm', 'XA_PHUONG', 'XA-GIA-KIEM', 'Xã Gia Kiệm', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-029', 'Ủy ban MTTQ Việt Nam Xã Bù Gia Mập', 'XA_PHUONG', 'XA-BU-GIA-MAP', 'Xã Bù Gia Mập', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-030', 'Ủy ban MTTQ Việt Nam Phường Hàng Gòn', 'XA_PHUONG', 'XA-HANG-GON', 'Phường Hàng Gòn', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-031', 'Ủy ban MTTQ Việt Nam Phường Hố Nai', 'XA_PHUONG', 'XA-HO-NAI', 'Phường Hố Nai', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-032', 'Ủy ban MTTQ Việt Nam Xã Hưng Phước', 'XA_PHUONG', 'XA-HUNG-PHUOC', 'Xã Hưng Phước', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-033', 'Ủy ban MTTQ Việt Nam Xã Hưng Thịnh', 'XA_PHUONG', 'XA-HUNG-THINH', 'Xã Hưng Thịnh', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-034', 'Ủy ban MTTQ Việt Nam Xã La Ngà', 'XA_PHUONG', 'XA-LA-NGA', 'Xã La Ngà', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-035', 'Ủy ban MTTQ Việt Nam Xã Lộc Hưng', 'XA_PHUONG', 'XA-LOC-HUNG', 'Xã Lộc Hưng', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-036', 'Ủy ban MTTQ Việt Nam Phường Lộc Ninh', 'XA_PHUONG', 'XA-LOC-NINH', 'Phường Lộc Ninh', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-037', 'Ủy ban MTTQ Việt Nam Xã Lộc Quang', 'XA_PHUONG', 'XA-LOC-QUANG', 'Xã Lộc Quang', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-038', 'Ủy ban MTTQ Việt Nam Xã Lộc Tấn', 'XA_PHUONG', 'XA-LOC-TAN', 'Xã Lộc Tấn', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-039', 'Ủy ban MTTQ Việt Nam Xã Lộc Thành', 'XA_PHUONG', 'XA-LOC-THANH', 'Xã Lộc Thành', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-040', 'Ủy ban MTTQ Việt Nam Xã Lộc Thạnh', 'XA_PHUONG', 'XA-LOC-THANH-2', 'Xã Lộc Thạnh', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-041', 'Ủy ban MTTQ Việt Nam Phường Long Bình', 'XA_PHUONG', 'XA-LONG-BINH', 'Phường Long Bình', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-042', 'Ủy ban MTTQ Việt Nam Xã Long Hà', 'XA_PHUONG', 'XA-LONG-HA', 'Xã Long Hà', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-043', 'Ủy ban MTTQ Việt Nam Phường Long Hưng', 'XA_PHUONG', 'XA-LONG-HUNG', 'Phường Long Hưng', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-044', 'Ủy ban MTTQ Việt Nam Phường Long Khánh', 'XA_PHUONG', 'XA-LONG-KHANH', 'Phường Long Khánh', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-045', 'Ủy ban MTTQ Việt Nam Xã Long Phước', 'XA_PHUONG', 'XA-LONG-PHUOC', 'Xã Long Phước', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-046', 'Ủy ban MTTQ Việt Nam Phường Long Thành', 'XA_PHUONG', 'XA-LONG-THANH', 'Phường Long Thành', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-047', 'Ủy ban MTTQ Việt Nam Xã Minh Đức', 'XA_PHUONG', 'XA-MINH-DUC', 'Xã Minh Đức', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-048', 'Ủy ban MTTQ Việt Nam Phường Minh Hưng', 'XA_PHUONG', 'XA-MINH-HUNG', 'Phường Minh Hưng', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-049', 'Ủy ban MTTQ Việt Nam Xã Nghĩa Trung', 'XA_PHUONG', 'XA-NGHIA-TRUNG', 'Xã Nghĩa Trung', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-050', 'Ủy ban MTTQ Việt Nam Xã Nha Bích', 'XA_PHUONG', 'XA-NHA-BICH', 'Xã Nha Bích', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-051', 'Ủy ban MTTQ Việt Nam Phường Nhơn Trạch', 'XA_PHUONG', 'XA-NHON-TRACH', 'Phường Nhơn Trạch', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-052', 'Ủy ban MTTQ Việt Nam Xã Phú Hoà', 'XA_PHUONG', 'XA-PHU-HOA', 'Xã Phú Hoà', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-053', 'Ủy ban MTTQ Việt Nam Xã Phú Lâm', 'XA_PHUONG', 'XA-PHU-LAM', 'Xã Phú Lâm', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-054', 'Ủy ban MTTQ Việt Nam Xã Phú Lý', 'XA_PHUONG', 'XA-PHU-LY', 'Xã Phú Lý', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-055', 'Ủy ban MTTQ Việt Nam Xã Phú Nghĩa', 'XA_PHUONG', 'XA-PHU-NGHIA', 'Xã Phú Nghĩa', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-056', 'Ủy ban MTTQ Việt Nam Xã Phú Riềng', 'XA_PHUONG', 'XA-PHU-RIENG', 'Xã Phú Riềng', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-057', 'Ủy ban MTTQ Việt Nam Xã Phú Trung', 'XA_PHUONG', 'XA-PHU-TRUNG', 'Xã Phú Trung', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-058', 'Ủy ban MTTQ Việt Nam Xã Phú Vinh', 'XA_PHUONG', 'XA-PHU-VINH', 'Xã Phú Vinh', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-059', 'Ủy ban MTTQ Việt Nam Xã Phước An', 'XA_PHUONG', 'XA-PHUOC-AN', 'Xã Phước An', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-060', 'Ủy ban MTTQ Việt Nam Phường Phước Bình', 'XA_PHUONG', 'XA-PHUOC-BINH', 'Phường Phước Bình', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-061', 'Ủy ban MTTQ Việt Nam Phường Phước Long', 'XA_PHUONG', 'XA-PHUOC-LONG', 'Phường Phước Long', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-062', 'Ủy ban MTTQ Việt Nam Xã Phước Sơn', 'XA_PHUONG', 'XA-PHUOC-SON', 'Xã Phước Sơn', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-063', 'Ủy ban MTTQ Việt Nam Phường Phước Tân', 'XA_PHUONG', 'XA-PHUOC-TAN', 'Phường Phước Tân', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-064', 'Ủy ban MTTQ Việt Nam Xã Phước Thái', 'XA_PHUONG', 'XA-PHUOC-THAI', 'Xã Phước Thái', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-065', 'Ủy ban MTTQ Việt Nam Xã Sông Ray', 'XA_PHUONG', 'XA-SONG-RAY', 'Xã Sông Ray', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-066', 'Ủy ban MTTQ Việt Nam Xã Tà Lài', 'XA_PHUONG', 'XA-TA-LAI', 'Xã Tà Lài', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-067', 'Ủy ban MTTQ Việt Nam Phường Tam Hiệp', 'XA_PHUONG', 'XA-TAM-HIEP', 'Phường Tam Hiệp', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-068', 'Ủy ban MTTQ Việt Nam Phường Tam Phước', 'XA_PHUONG', 'XA-TAM-PHUOC', 'Phường Tam Phước', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-069', 'Ủy ban MTTQ Việt Nam Xã Tân An', 'XA_PHUONG', 'XA-TAN-AN', 'Xã Tân An', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-070', 'Ủy ban MTTQ Việt Nam Xã Tân Hưng', 'XA_PHUONG', 'XA-TAN-HUNG', 'Xã Tân Hưng', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-071', 'Ủy ban MTTQ Việt Nam Phường Tân Khai', 'XA_PHUONG', 'XA-TAN-KHAI', 'Phường Tân Khai', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-072', 'Ủy ban MTTQ Việt Nam Xã Tân Lợi', 'XA_PHUONG', 'XA-TAN-LOI', 'Xã Tân Lợi', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-073', 'Ủy ban MTTQ Việt Nam Phường Tân Phú', 'XA_PHUONG', 'XA-TAN-PHU', 'Phường Tân Phú', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-074', 'Ủy ban MTTQ Việt Nam Xã Tân Quan', 'XA_PHUONG', 'XA-TAN-QUAN', 'Xã Tân Quan', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-075', 'Ủy ban MTTQ Việt Nam Xã Tân Tiến', 'XA_PHUONG', 'XA-TAN-TIEN', 'Xã Tân Tiến', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-076', 'Ủy ban MTTQ Việt Nam Phường Tân Triều', 'XA_PHUONG', 'XA-TAN-TRIEU', 'Phường Tân Triều', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-077', 'Ủy ban MTTQ Việt Nam Xã Thanh Sơn', 'XA_PHUONG', 'XA-THANH-SON', 'Xã Thanh Sơn', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-078', 'Ủy ban MTTQ Việt Nam Xã Thiện Hưng', 'XA_PHUONG', 'XA-THIEN-HUNG', 'Xã Thiện Hưng', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-079', 'Ủy ban MTTQ Việt Nam Xã Thọ Sơn', 'XA_PHUONG', 'XA-THO-SON', 'Xã Thọ Sơn', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-080', 'Ủy ban MTTQ Việt Nam Xã Thống Nhất', 'XA_PHUONG', 'XA-THONG-NHAT', 'Xã Thống Nhất', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-081', 'Ủy ban MTTQ Việt Nam Xã Thuận Lợi', 'XA_PHUONG', 'XA-THUAN-LOI', 'Xã Thuận Lợi', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-082', 'Ủy ban MTTQ Việt Nam Phường Trấn Biên', 'XA_PHUONG', 'XA-TRAN-BIEN', 'Phường Trấn Biên', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-083', 'Ủy ban MTTQ Việt Nam Phường Trảng Bom', 'XA_PHUONG', 'XA-TRANG-BOM', 'Phường Trảng Bom', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-084', 'Ủy ban MTTQ Việt Nam Phường Trảng Dài', 'XA_PHUONG', 'XA-TRANG-DAI', 'Phường Trảng Dài', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-085', 'Ủy ban MTTQ Việt Nam Phường Trị An', 'XA_PHUONG', 'XA-TRI-AN', 'Phường Trị An', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-086', 'Ủy ban MTTQ Việt Nam Xã Xuân Bắc', 'XA_PHUONG', 'XA-XUAN-BAC', 'Xã Xuân Bắc', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-087', 'Ủy ban MTTQ Việt Nam Xã Xuân Định', 'XA_PHUONG', 'XA-XUAN-DINH', 'Xã Xuân Định', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-088', 'Ủy ban MTTQ Việt Nam Xã Xuân Đông', 'XA_PHUONG', 'XA-XUAN-DONG', 'Xã Xuân Đông', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-089', 'Ủy ban MTTQ Việt Nam Xã Xuân Đường', 'XA_PHUONG', 'XA-XUAN-DUONG', 'Xã Xuân Đường', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-090', 'Ủy ban MTTQ Việt Nam Xã Xuân Hoà', 'XA_PHUONG', 'XA-XUAN-HOA', 'Xã Xuân Hoà', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-091', 'Ủy ban MTTQ Việt Nam Phường Xuân Lập', 'XA_PHUONG', 'XA-XUAN-LAP', 'Phường Xuân Lập', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-092', 'Ủy ban MTTQ Việt Nam Phường Xuân Lộc', 'XA_PHUONG', 'XA-XUAN-LOC', 'Phường Xuân Lộc', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-093', 'Ủy ban MTTQ Việt Nam Xã Xuân Phú', 'XA_PHUONG', 'XA-XUAN-PHU', 'Xã Xuân Phú', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-094', 'Ủy ban MTTQ Việt Nam Xã Xuân Quế', 'XA_PHUONG', 'XA-XUAN-QUE', 'Xã Xuân Quế', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG'],
    ['UNIT-095', 'Ủy ban MTTQ Việt Nam Xã Xuân Thành', 'XA_PHUONG', 'XA-XUAN-THANH', 'Xã Xuân Thành', '', '', '', 'Cán bộ', '', '', 'HOAT_DONG']
  ];
  sheet.clearContents(); // Dọn dẹp bảng cũ
  sheet.appendRow(['UnitID', 'TenDonVi', 'Cap', 'XaPhuongID', 'XaPhuong', 'ThonApKhuPhoID', 'ThonApKhuPho', 'DiaChi', 'NguoiPhuTrach', 'DienThoai', 'Email', 'TrangThai']);
  sheet.setFrozenRows(1);
  units.forEach(function(row) { sheet.appendRow(row); });
}