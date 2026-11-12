// Lấy danh sách danh mục (Bỏ qua dòng tiêu đề và chỉ lấy mục HOAT_DONG)
function getCategories() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CATEGORIES');
  if (!sheet) return { success: false, message: 'Không tìm thấy trang tính CATEGORIES.' };
  
  var data = sheet.getDataRange().getValues();
  var categories = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] !== 'HOAT_DONG') continue; 
    
    categories.push({
      CatID: data[i][0],      
      NhomDuLieu: data[i][1], 
      LoaiDiem: data[i][2],   
      MoTa: data[i][3],       
      IconName: data[i][4] || 'place' 
    });
  }
  return { success: true, categories: categories };
}

// Thêm hoặc Cập nhật danh mục (Chỉ ADMIN)
function adminSaveCategory(token, catData) {
  var session = verifySession(token);
  if (!session || session.VaiTro !== 'ADMIN') return { success: false, message: 'Chỉ Admin mới có quyền thực hiện.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CATEGORIES');
  var data = sheet.getDataRange().getValues();
  
  if (catData.CatID) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === catData.CatID) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[catData.NhomDuLieu, catData.LoaiDiem, catData.MoTa, catData.IconName]]);
        return { success: true, message: 'Đã cập nhật danh mục thành công!' };
      }
    }
  }

  var newId = 'CAT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
  var nextThuTu = data.length; 
  
  sheet.appendRow([newId, catData.NhomDuLieu, catData.LoaiDiem, catData.MoTa, catData.IconName, 'HOAT_DONG', nextThuTu]);
  return { success: true, message: 'Đã thêm danh mục mới thành công!' };
}

// Xóa danh mục
function adminDeleteCategory(token, catId) {
  var session = verifySession(token);
  if (!session || session.VaiTro !== 'ADMIN') return { success: false, message: 'Không có quyền truy cập.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CATEGORIES');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === catId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Đã xóa danh mục!' };
    }
  }
  return { success: false, message: 'Không tìm thấy danh mục.' };
}