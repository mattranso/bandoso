// ==========================================
// HÀM TIỆN ÍCH: DÒ TÌM CỘT THÔNG MINH
// ==========================================
function getSafeCol(headers, colName) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === colName) return i;
  }
  return -1; 
}

// ==========================================
// LẤY DANH SÁCH ĐỊA ĐIỂM
// ==========================================
function getLocations(token) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOCATIONS');
  if (!sheet) return { success: false, locations: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, locations: [] };

  var headers = data[0];
  var locations = [];

  var iId = getSafeCol(headers, 'LocationID');
  var iName = getSafeCol(headers, 'TenDiem');
  var iGroup = getSafeCol(headers, 'NhomDuLieu');
  var iType = getSafeCol(headers, 'LoaiDiem');
  var iWard = getSafeCol(headers, 'XaPhuong');
  var iAddress = getSafeCol(headers, 'DiaChi');
  var iDesc = getSafeCol(headers, 'MoTa');
  var iLat = getSafeCol(headers, 'ViDo');
  var iLng = getSafeCol(headers, 'KinhDo');
  var iStatus = getSafeCol(headers, 'TrangThai');

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (iStatus !== -1 && row[iStatus] !== 'DA_DUYET') continue; 

    locations.push({
      LocID: iId !== -1 ? row[iId] : '', 
      TenDiem: iName !== -1 ? row[iName] : '',
      NhomDuLieu: iGroup !== -1 ? row[iGroup] : '',
      LoaiDiem: iType !== -1 ? row[iType] : '',
      XaPhuong: iWard !== -1 ? row[iWard] : '',
      DiaChi: iAddress !== -1 ? row[iAddress] : '',
      MoTa: iDesc !== -1 ? row[iDesc] : '',
      ViDo: iLat !== -1 ? row[iLat] : '',
      KinhDo: iLng !== -1 ? row[iLng] : '',
      TrangThai: iStatus !== -1 ? row[iStatus] : ''
    });
  }
  return { success: true, locations: locations };
}

// ==========================================
// THÊM ĐỊA ĐIỂM MỚI (CHẶN LỖI ĐỊNH DẠNG SỐ)
// ==========================================
function createLocation(locData, token) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOCATIONS');
  if (!sheet) return { success: false, message: 'Lỗi: Không tìm thấy bảng LOCATIONS' };

  var headers = sheet.getDataRange().getValues()[0];
  var newRow = new Array(headers.length).fill(''); 
  var newId = 'LOC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  for (var i = 0; i < headers.length; i++) {
    var headerName = String(headers[i]).trim();
    if (headerName === 'LocationID') newRow[i] = newId;
    if (headerName === 'TenDiem') newRow[i] = locData.TenDiem;
    if (headerName === 'NhomDuLieu') newRow[i] = locData.NhomDuLieu;
    if (headerName === 'LoaiDiem') newRow[i] = locData.LoaiDiem;
    if (headerName === 'XaPhuongID') newRow[i] = locData.XaPhuongID;
    if (headerName === 'XaPhuong') newRow[i] = locData.XaPhuong;
    if (headerName === 'DiaChi') newRow[i] = locData.DiaChi;
    if (headerName === 'MoTa') newRow[i] = locData.MoTa;
    
    // ĐÃ SỬA: Thêm dấu nháy đơn (') để ép Google Sheets hiểu đây là Chữ, cấm tự đổi thành số tỷ
    if (headerName === 'ViDo') newRow[i] = "'" + String(locData.ViDo).replace(',', '.');
    if (headerName === 'KinhDo') newRow[i] = "'" + String(locData.KinhDo).replace(',', '.');
    
    if (headerName === 'TrangThai') newRow[i] = 'CHO_DUYET'; 
    if (headerName === 'UserID' && token !== 'GUEST') {
      var session = verifySession(token);
      if(session) newRow[i] = session.UserID;
    }
  }

  sheet.appendRow(newRow);
  return { success: true, message: 'Gửi thành công! Vui lòng chờ Admin phê duyệt để hiển thị.' };
}

// ==========================================
// CÁC HÀM CHO ADMIN QUẢN LÝ ĐỊA ĐIỂM
// ==========================================
function getManageLocationsAdmin(token) {
  var session = verifySession(token);
  if (!session || (session.VaiTro !== 'ADMIN' && session.VaiTro !== 'CAN_BO_XA')) return { success: false };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOCATIONS');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var locations = [];

  var iId = getSafeCol(headers, 'LocationID');
  var iName = getSafeCol(headers, 'TenDiem');
  var iGroup = getSafeCol(headers, 'NhomDuLieu');
  var iType = getSafeCol(headers, 'LoaiDiem');
  var iWard = getSafeCol(headers, 'XaPhuong');
  var iStatus = getSafeCol(headers, 'TrangThai');
  var iAddress = getSafeCol(headers, 'DiaChi');
  var iDesc = getSafeCol(headers, 'MoTa');
  var iLat = getSafeCol(headers, 'ViDo');
  var iLng = getSafeCol(headers, 'KinhDo');

  if (iId === -1) return { success: false, message: 'Lỗi: Cột LocationID bị thiếu!' };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    locations.push({
      LocID: row[iId], 
      TenDiem: iName !== -1 ? row[iName] : '',
      NhomDuLieu: iGroup !== -1 ? row[iGroup] : '',
      LoaiDiem: iType !== -1 ? row[iType] : '',
      XaPhuong: iWard !== -1 ? row[iWard] : '',
      TrangThai: iStatus !== -1 ? row[iStatus] : '',
      DiaChi: iAddress !== -1 ? row[iAddress] : '',
      MoTa: iDesc !== -1 ? row[iDesc] : '',
      ViDo: iLat !== -1 ? row[iLat] : '',
      KinhDo: iLng !== -1 ? row[iLng] : ''
    });
  }
  
  locations.sort(function(a, b) {
    if (a.TrangThai === 'CHO_DUYET' && b.TrangThai !== 'CHO_DUYET') return -1;
    if (a.TrangThai !== 'CHO_DUYET' && b.TrangThai === 'CHO_DUYET') return 1;
    return 0;
  });

  return { success: true, data: locations };
}

function adminUpdateLocationStatus(token, locId, newStatus) {
  var session = verifySession(token);
  if (!session || session.VaiTro !== 'ADMIN') return { success: false, message: 'Bạn không có quyền duyệt.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOCATIONS');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = getSafeCol(headers, 'LocationID');
  var statusCol = getSafeCol(headers, 'TrangThai');

  if (idCol === -1 || statusCol === -1) return { success: false };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(locId).trim()) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      return { success: true, message: 'Đã cập nhật trạng thái!' };
    }
  }
  return { success: false, message: 'Không tìm thấy địa điểm.' };
}

function adminUpdateLocation(token, locData) {
  var session = verifySession(token);
  if (!session || session.VaiTro !== 'ADMIN') return { success: false, message: 'Không có quyền.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOCATIONS');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = getSafeCol(headers, 'LocationID');
  if (idCol === -1) return { success: false };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(locData.LocID).trim()) {
      var iName = getSafeCol(headers, 'TenDiem'); if(iName !== -1) sheet.getRange(i + 1, iName + 1).setValue(locData.TenDiem);
      var iGroup = getSafeCol(headers, 'NhomDuLieu'); if(iGroup !== -1) sheet.getRange(i + 1, iGroup + 1).setValue(locData.NhomDuLieu);
      var iType = getSafeCol(headers, 'LoaiDiem'); if(iType !== -1) sheet.getRange(i + 1, iType + 1).setValue(locData.LoaiDiem);
      var iAddress = getSafeCol(headers, 'DiaChi'); if(iAddress !== -1) sheet.getRange(i + 1, iAddress + 1).setValue(locData.DiaChi);
      var iDesc = getSafeCol(headers, 'MoTa'); if(iDesc !== -1) sheet.getRange(i + 1, iDesc + 1).setValue(locData.MoTa);
      return { success: true, message: 'Đã cập nhật thông tin địa điểm!' };
    }
  }
  return { success: false, message: 'Không tìm thấy.' };
}

function adminDeleteLocation(token, locId) {
  var session = verifySession(token);
  if (!session || session.VaiTro !== 'ADMIN') return { success: false };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOCATIONS');
  var data = sheet.getDataRange().getValues();
  var idCol = getSafeCol(data[0], 'LocationID');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(locId).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Đã xóa địa điểm!' };
    }
  }
  return { success: false };
}