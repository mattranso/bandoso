function login(username, password) {
  try {
    var sheet = getSS_().getSheetByName('USERS');
    var data = sheet.getDataRange().getValues();
    var passHash = hashPassword_(password);
    
    var scriptProps = PropertiesService.getScriptProperties();
    var usernameKey = username.trim().toLowerCase();
    var attemptKey = 'LOGIN_FAILS_' + usernameKey;
    var lockKey = 'LOGIN_LOCK_' + usernameKey;
    
    // Kiểm tra khóa 10 phút
    var lockTime = scriptProps.getProperty(lockKey);
    if (lockTime) {
      var timeElapsed = new Date().getTime() - parseInt(lockTime);
      if (timeElapsed < 10 * 60 * 1000) {
        return { success: false, message: "Tài khoản đang bị khóa bảo mật. Thử lại sau " + Math.ceil((10 * 60 * 1000 - timeElapsed) / 60000) + " phút." };
      } else {
        scriptProps.deleteProperty(lockKey); 
        scriptProps.deleteProperty(attemptKey);
      }
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === usernameKey) {
        // Sai mật khẩu
        if (data[i][2] !== passHash) {
          var fails = parseInt(scriptProps.getProperty(attemptKey) || '0') + 1;
          scriptProps.setProperty(attemptKey, fails.toString());
          if (fails >= 5) {
            scriptProps.setProperty(lockKey, new Date().getTime().toString());
            return { success: false, message: "Tạm khóa bảo mật 10 phút do nhập sai 5 lần." };
          }
          return { success: false, message: "Sai mật khẩu! Bạn còn " + (5 - fails) + " lần thử." };
        }
        
        // Trạng thái tài khoản
        if (data[i][13] === 'CHO_DUYET') return { success: false, message: "Tài khoản đang chờ Xã/Phường duyệt." };
        if (data[i][13] === 'TU_CHOI') return { success: false, message: "Tài khoản đã bị khóa." };
        
        // Đăng nhập thành công
        scriptProps.deleteProperty(attemptKey); 
        scriptProps.deleteProperty(lockKey);
        var sessionData = { 
          UserID: data[i][0], Username: data[i][1], HoTen: data[i][3], 
          VaiTro: data[i][4], XaPhuongID: data[i][7], XaPhuong: data[i][8], UnitID: data[i][6] 
        };
        var token = Utilities.getUuid();
        CacheService.getScriptCache().put(token, JSON.stringify(sessionData), 21600);
        return { success: true, token: token, user: sessionData };
      }
    }
    return { success: false, message: "Tài khoản không tồn tại." };
  } catch (e) { return { success: false, message: e.message }; }
}

function verifySession(token) { var s = CacheService.getScriptCache().get(token); return s ? JSON.parse(s) : null; }
function getCurrentUser(token) { return { success: true, user: verifySession(token) }; }
function logout(token) { CacheService.getScriptCache().remove(token); return { success: true }; }

// ĐĂNG KÝ VÀ TỰ ĐỘNG PHÂN LOẠI
function registerUser(formData) {
  try {
    var sheet = getSS_().getSheetByName('USERS');
    var data = sheet.getDataRange().getValues();
    var username = formData.Username.trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === username) {
        return { success: false, message: "Tên đăng nhập đã tồn tại." };
      }
    }
    
    // Tự động phân loại: Có nhập Thôn/Ấp -> Cán bộ thôn, Bỏ trống -> Cán bộ Xã
    var vaiTro = formData.ThonAp ? "CAN_BO_THON" : "CAN_BO_XA";
    var cap = formData.ThonAp ? "THON_AP" : "XA_PHUONG";
    
    sheet.appendRow([
      "USER-" + Utilities.getUuid().substring(0, 8).toUpperCase(), // UserID
      formData.Username, 
      hashPassword_(formData.Password), 
      formData.HoTen, 
      vaiTro, 
      cap, 
      "", // UnitID
      formData.XaPhuongID, 
      formData.XaPhuongName, 
      "", // ThonApID
      formData.ThonAp, // Ghi tên Thôn/Ấp vào cột K
      formData.SoDienThoai, 
      "", // Email
      "CHO_DUYET", 
      new Date(), 
      "" // Lan đăng nhập cuối
    ]);
    
    return { success: true, message: "Đăng ký thành công! Đang chờ Cán bộ quản lý phê duyệt." };
  } catch (e) { 
    return { success: false, message: e.message }; 
  }
}

function getPendingUsers(token) {
  var user = verifySession(token);
  if (!user || (user.VaiTro !== 'ADMIN' && user.VaiTro !== 'CAN_BO_XA')) return { success: false };
  var data = getSS_().getSheetByName('USERS').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][13] === 'CHO_DUYET') {
      if (user.VaiTro === 'ADMIN' || data[i][7] === user.XaPhuongID) {
        list.push({ UserID: data[i][0], Username: data[i][1], HoTen: data[i][3], XaPhuong: data[i][8] });
      }
    }
  }
  return { success: true, users: list };
}

function approveUser(userId, action, token) {
  var user = verifySession(token);
  if (!user || (user.VaiTro !== 'ADMIN' && user.VaiTro !== 'CAN_BO_XA')) return { success: false };
  var sheet = getSS_().getSheetByName('USERS');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      if (user.VaiTro === 'CAN_BO_XA' && data[i][7] !== user.XaPhuongID) return { success: false, message: "Vượt quá thẩm quyền đơn vị." };
      sheet.getRange(i + 1, 14).setValue(action === 'APPROVE' ? 'HOAT_DONG' : 'TU_CHOI');
      return { success: true, message: "Đã xử lý thành công." };
    }
  }
}

function getAllUsersForAdmin(token) {
  var user = verifySession(token);
  if (!user || (user.VaiTro !== 'ADMIN' && user.VaiTro !== 'CAN_BO_XA')) return { success: false };
  var sheet = getSS_().getSheetByName('USERS');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if(data[i][0]) {
      if (user.VaiTro === 'ADMIN' || (data[i][7] === user.XaPhuongID && data[i][4] !== 'ADMIN')) {
        list.push({ UserID: data[i][0], Username: data[i][1], HoTen: data[i][3], VaiTro: data[i][4], XaPhuong: data[i][8], TrangThai: data[i][13] });
      }
    }
  }
  return { success: true, data: list };
}

function adminResetPassword(token, targetUserId) {
  var user = verifySession(token);
  if (!user || (user.VaiTro !== 'ADMIN' && user.VaiTro !== 'CAN_BO_XA')) return { success: false };
  var sheet = getSS_().getSheetByName('USERS');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === targetUserId) { 
      if (user.VaiTro === 'CAN_BO_XA' && data[i][4] === 'ADMIN') return { success: false, message: "Không được phép reset tài khoản cấp cao hơn." };
      sheet.getRange(i + 1, 3).setValue(hashPassword_("Mttq@2026")); 
      var sp = PropertiesService.getScriptProperties();
      sp.deleteProperty('LOGIN_FAILS_' + data[i][1].toString().toLowerCase());
      sp.deleteProperty('LOGIN_LOCK_' + data[i][1].toString().toLowerCase());
      return { success: true, message: "Đã reset mật khẩu về Mttq@2026 và mở khóa." }; 
    }
  }
}

function adminDeleteUser(token, targetUserId) {
  var user = verifySession(token);
  if (!user || (user.VaiTro !== 'ADMIN' && user.VaiTro !== 'CAN_BO_XA')) return { success: false };
  var sheet = getSS_().getSheetByName('USERS');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === targetUserId) { 
      if (user.VaiTro === 'CAN_BO_XA' && data[i][4] === 'ADMIN') return { success: false, message: "Từ chối quyền xóa." };
      sheet.deleteRow(i + 1); return { success: true }; 
    }
  }
}

function adminUpdateUser(token, updateData) {
  var user = verifySession(token);
  if (!user || (user.VaiTro !== 'ADMIN' && user.VaiTro !== 'CAN_BO_XA')) return { success: false };
  var sheet = getSS_().getSheetByName('USERS');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === updateData.UserID) {
      var newRole = updateData.VaiTro;
      // Ép quyền: Cán bộ xã không thể set ai đó làm Admin hoặc Cán bộ xã ngang hàng
      if (user.VaiTro === 'CAN_BO_XA' && (newRole === 'ADMIN' || newRole === 'CAN_BO_XA')) newRole = 'CAN_BO_THON'; 
      sheet.getRange(i + 1, 4).setValue(updateData.HoTen);
      sheet.getRange(i + 1, 5).setValue(newRole);
      sheet.getRange(i + 1, 9).setValue(updateData.XaPhuong);
      sheet.getRange(i + 1, 14).setValue(updateData.TrangThai);
      return { success: true, message: "Cập nhật thành công." };
    }
  }
}

function importUsersThon(token, userList) {
  var user = verifySession(token);
  if (!user || (user.VaiTro !== 'ADMIN' && user.VaiTro !== 'CAN_BO_XA')) return { success: false, message: "Không có quyền!" };
  var sheet = getSS_().getSheetByName('USERS');
  var existingData = sheet.getDataRange().getValues();
  var existUsernames = existingData.map(r => r[1] ? r[1].toString().trim().toLowerCase() : '');
  
  var added = 0; var defaultPass = hashPassword_('Mttq@2026');
  userList.forEach(u => {
    var uname = u.Username.trim().toLowerCase();
    if (uname && existUsernames.indexOf(uname) === -1) {
      sheet.appendRow(["USER-" + Utilities.getUuid().substring(0, 8).toUpperCase(), u.Username, defaultPass, u.HoTen, "CAN_BO_THON", "THON_AP", user.UnitID, user.XaPhuongID, user.XaPhuong, "", u.ThonAp, u.SoDienThoai, "", "HOAT_DONG", new Date(), ""]);
      existUsernames.push(uname); added++;
    }
  });
  return { success: true, message: `Đã import thành công ${added} tài khoản! Mật khẩu mặc định: Mttq@2026` };
}

function changePassword(token, oldPass, newPass) {
  var user = verifySession(token);
  var sheet = getSS_().getSheetByName('USERS');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === user.UserID) {
      if (data[i][2] !== hashPassword_(oldPass)) return { success: false, message: "Mật khẩu cũ sai." };
      sheet.getRange(i + 1, 3).setValue(hashPassword_(newPass));
      return { success: true, message: "Đổi mật khẩu thành công!" };
    }
  }
}