// Lấy danh sách 95 Đơn vị Xã/Phường (MỞ CÔNG KHAI CHO KHÁCH)
function getUnits() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UNITS');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('UNITS');
    sheet.appendRow(['UnitID', 'TenDonVi']);
    return { success: true, units: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  var units = [];
  for (var i = 1; i < data.length; i++) {
    units.push({ UnitID: data[i][0], TenDonVi: data[i][1] });
  }
  return { success: true, units: units };
}

// Thêm hoặc Sửa Đơn vị (Chỉ ADMIN)
function adminSaveUnit(token, unitData) {
  var session = verifySession(token);
  if (!session || session.VaiTro !== 'ADMIN') return { success: false, message: 'Chỉ Admin mới có quyền thực hiện.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UNITS');
  var data = sheet.getDataRange().getValues();
  
  if (unitData.UnitID) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === unitData.UnitID) {
        sheet.getRange(i + 1, 2).setValue(unitData.TenDonVi);
        return { success: true, message: 'Đã cập nhật đơn vị thành công!' };
      }
    }
  }

  var newId = 'UNIT-' + Utilities.getUuid().substring(0, 5).toUpperCase();
  sheet.appendRow([newId, unitData.TenDonVi]);
  return { success: true, message: 'Đã thêm đơn vị mới thành công!' };
}

// Xóa Đơn vị (Chỉ ADMIN)
function adminDeleteUnit(token, unitId) {
  var session = verifySession(token);
  if (!session || session.VaiTro !== 'ADMIN') return { success: false, message: 'Không có quyền truy cập.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UNITS');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === unitId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Đã xóa đơn vị!' };
    }
  }
  return { success: false, message: 'Không tìm thấy đơn vị.' };
}