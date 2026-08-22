/**
 * DỰ ÁN QUẢN LÝ KHO HỶ SÂM KÝ
 * Developer: Senior Developer & AI Partner
 * Cấu hình khớp chính xác với Spreadsheet ID: 1P7PGYjDZFxIT1kG_EmuFI91omVHKBFHqLMgi-GtAbZU
 */

const SPREADSHEET_ID = '1P7PGYjDZFxIT1kG_EmuFI91omVHKBFHqLMgi-GtAbZU';

// --- HÀM KHỞI TẠO DATABASE (ĐẢM BẢO KHỚP TÊN TAB THỰC TẾ) ---
function getDb() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return {
      ss: ss,
      inventory: ss.getSheetByName('Kho'), // Khớp với tab "Kho"
      loginLogs: ss.getSheetByName('LoginLogs'), // Khớp với tab "LoginLogs"
      users: ss.getSheetByName('Users'), // Khớp với tab "Users"
      reports: ss.getSheetByName('InventoryReports') || ss.insertSheet('InventoryReports'),
      transactionLogs: ss.getSheetByName('Logs') || ss.insertSheet('Logs')
    };
  } catch (e) {
    console.error("Lỗi kết nối Spreadsheet: " + e.message);
    return null;
  }
}

// --- XỬ LÝ GET REQUEST (LẤY DỮ LIỆU) ---
function doGet(e) {
  const db = getDb();
  if (!db) return sendResponse({ status: 'error', message: 'Database Connection Failed' });

  const action = e.parameter.action;

  try {
    switch (action) {
      case 'getLogins':
        return sendResponse(fetchData(db.loginLogs));
      case 'getInventory':
        return sendResponse(fetchData(db.inventory));
      case 'getUsers':
        return sendResponse(fetchData(db.users));
      default:
        return sendResponse({ status: 'error', message: 'Invalid Action' });
    }
  } catch (err) {
    return sendResponse({ status: 'error', message: err.toString() });
  }
}

// --- XỬ LÝ POST REQUEST (GHI DỮ LIỆU) ---
function doPost(e) {
  const db = getDb();
  if (!db) return sendResponse({ status: 'error', message: 'Database Connection Failed' });

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    // 1. Xử lý Đăng nhập (Ghi log)
    if (action === 'login') {
      const sheet = db.loginLogs;
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["Timestamp", "Username", "Role", "Details"]);
      }
      sheet.appendRow([new Date(), data.username, data.role, data.details]);
      return sendResponse({ status: 'success' });
    }

    // 2. Xử lý Đồng bộ kho (Sync từ App lên Sheet)
    if (action === 'sync_inventory') {
      const sheet = db.inventory;
      sheet.clear(); // Xóa dữ liệu cũ để đồng bộ mới hoàn toàn
      sheet.appendRow(["ID", "Tên", "Danh mục", "Đơn vị", "Tồn kho", "Vị trí", "Ngưỡng tối thiểu", "Giá nhập", "Mô tả"]);
      
      if (data.inventory && data.inventory.length > 0) {
        const rows = data.inventory.map(item => [
          item.id, item.name, item.category, item.unit, 
          item.actualStock, item.location, item.minThreshold, 
          item.importPrice, item.description
        ]);
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      }

      // Xử lý phiếu Nhập kho -> sheet "IN"
      if (data.inboundLogs && data.inboundLogs.length > 0) {
        const inSheet = db.ss.getSheetByName("IN") || db.ss.insertSheet("IN");
        inSheet.clear();
        inSheet.appendRow(["Timestamp", "User", "Type", "Item", "Category", "Amount", "Unit", "Price", "Reason", "Stock After"]);
        const rowsIn = data.inboundLogs.map(log => [
          log.timestamp || new Date(), log.user, log.type, log.name, log.category || '', 
          log.amount, log.unit || '', log.importPrice || 0, log.reason || '', log.newStock || 0
        ]);
        inSheet.getRange(2, 1, rowsIn.length, rowsIn[0].length).setValues(rowsIn);
      }

      // Xử lý phiếu Xuất kho -> sheet "OUT"
      if (data.outboundLogs && data.outboundLogs.length > 0) {
        const outSheet = db.ss.getSheetByName("OUT") || db.ss.insertSheet("OUT");
        outSheet.clear();
        outSheet.appendRow(["Timestamp", "User", "Type", "Item", "Category", "Amount", "Unit", "Price", "Reason", "Stock After"]);
        const rowsOut = data.outboundLogs.map(log => [
          log.timestamp || new Date(), log.user, log.type, log.name, log.category || '', 
          log.amount, log.unit || '', log.importPrice || 0, log.reason || '', log.newStock || 0
        ]);
        outSheet.getRange(2, 1, rowsOut.length, rowsOut[0].length).setValues(rowsOut);
      }

      return sendResponse({ status: 'success', message: 'Kho và các nhật ký IN/OUT đã được đồng bộ' });
    }

    // 3. Xử lý Báo cáo Kiểm kê (Hàng ngày/Tuần/Tháng)
    if (action === 'audit_report') {
      const auditType = data.auditType || "daily";
      let sheetName = "everyDAY";
      if (auditType === "weekly") sheetName = "everyWEEK";
      if (auditType === "monthly") sheetName = "everyMONTH";
      
      const sheet = db.ss.getSheetByName(sheetName) || db.ss.insertSheet(sheetName);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["Timestamp", "User", "Item", "Category", "Theoretical", "Actual", "Difference", "Unit", "Reason"]);
      }
      
      if (data.auditData && data.auditData.length > 0) {
        const rows = data.auditData.map(log => [
          new Date(), log.user, log.name, log.category,
          log.theoreticalStock, log.actualStock, 
          (log.actualStock - log.theoreticalStock),
          log.unit, log.reason
        ]);
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
      return sendResponse({ status: 'success' });
    }

    // 4. Xử lý Ghi log (Logs)
    if (action === 'log_inventory_change') {
      const sheet = db.transactionLogs;
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["Timestamp", "User", "Type", "Item", "Category", "Amount", "Unit", "Price", "Reason", "Stock After"]);
      }
      if (data.logs && data.logs.length > 0) {
        const rows = data.logs.map(log => [
          new Date(), log.user, log.type, log.name, log.category || '', 
          log.amount, log.unit || '', log.price || 0, log.reason || '', log.stockAfter || 0
        ]);
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        
        // Cập nhật thêm vào IN và OUT tương ứng
        const inLogs = data.logs.filter(l => l.type === "Nhập kho");
        if (inLogs.length > 0) {
          const inSheet = db.ss.getSheetByName("IN") || db.ss.insertSheet("IN");
          if (inSheet.getLastRow() === 0) {
             inSheet.appendRow(["Timestamp", "User", "Type", "Item", "Category", "Amount", "Unit", "Price", "Reason", "Stock After"]);
          }
          const rowsIn = inLogs.map(log => [
            new Date(), log.user, log.type, log.name, log.category || '', 
            log.amount, log.unit || '', log.price || 0, log.reason || '', log.stockAfter || 0
          ]);
          inSheet.getRange(inSheet.getLastRow() + 1, 1, rowsIn.length, rowsIn[0].length).setValues(rowsIn);
        }

        const outLogs = data.logs.filter(l => l.type === "Xuất kho");
        if (outLogs.length > 0) {
          const outSheet = db.ss.getSheetByName("OUT") || db.ss.insertSheet("OUT");
          if (outSheet.getLastRow() === 0) {
             outSheet.appendRow(["Timestamp", "User", "Type", "Item", "Category", "Amount", "Unit", "Price", "Reason", "Stock After"]);
          }
           const rowsOut = outLogs.map(log => [
            new Date(), log.user, log.type, log.name, log.category || '', 
            log.amount, log.unit || '', log.price || 0, log.reason || '', log.stockAfter || 0
          ]);
          outSheet.getRange(outSheet.getLastRow() + 1, 1, rowsOut.length, rowsOut[0].length).setValues(rowsOut);
        }
      }
      return sendResponse({ status: 'success' });
    }

    return sendResponse({ status: 'error', message: 'Action not supported: ' + action });
  } catch (err) {
    return sendResponse({ status: 'error', message: 'Post Error: ' + err.toString() });
  }
}

// --- HÀM HỖ TRỢ ---
function fetchData(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, i) => obj[header] = row[i]);
    return obj;
  });
}

function sendResponse(content) {
  return ContentService.createTextOutput(JSON.stringify(content))
    .setMimeType(ContentService.MimeType.JSON);
}
