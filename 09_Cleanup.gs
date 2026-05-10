// =========================================================================
// TÍNH NĂNG 3: DỌN DẸP
// =========================================================================
function checkAndCleanData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const apiKey = getApiKey();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIDEO);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const videoIds = []; const rowMap = {}; const rowsToDelete = []; const seenIds = new Set();

    for (let i = 0; i < data.length; i++) {
      const vId = extractVideoId(data[i][2]);
      if (!vId) continue;
      if (seenIds.has(vId)) rowsToDelete.push(i + 2);
      else { seenIds.add(vId); videoIds.push(vId); rowMap[vId] = i + 2; }
    }

    const threeMonthsAgo = new Date(); threeMonthsAgo.setDate(threeMonthsAgo.getDate() - 90);

    for (let i = 0; i < videoIds.length; i += 50) {
      const json = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/videos?part=snippet&id=${videoIds.slice(i, i + 50).join(',')}`, apiKey);
      if (json.items) {
        json.items.forEach(item => {
          if (new Date(item.snippet.publishedAt) < threeMonthsAgo) rowsToDelete.push(rowMap[item.id]);
        });
      }
    }

    if (rowsToDelete.length > 0) {
      const uniqueRows = [...new Set(rowsToDelete)].sort((a, b) => b - a);
      deleteRowsInBlocks_(sheet, uniqueRows);
      recalculateSTT(SHEET_VIDEO);
      ui.alert(`✅ Đã dọn dẹp ${uniqueRows.length} hàng.`);
    } else ui.alert(`✅ Dữ liệu đã sạch.`);
  } catch (error) { ui.alert(`❌ LỖI: ${error.message}`); }
}

function deleteLowViewVideos() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🗑️ Xác nhận dọn dẹp', 
    'Bạn có chắc chắn muốn xóa tất cả các video có lượt view dưới 20.000 trong Sheet Video không?\n\n(Hệ thống sẽ không thể hoàn tác sau khi thực hiện)', 
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.NO) return;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIDEO);
    if (!sheet) return ui.alert("Không tìm thấy Sheet: " + SHEET_VIDEO);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return ui.alert("Sheet hiện tại không có dữ liệu để quét.");

    const viewData = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    const rowsToDelete = [];
    
    for (let i = 0; i < viewData.length; i++) {
      const viewString = viewData[i][0];
      if (viewString !== null && viewString !== undefined && viewString !== "") {
        const views = unformatNumber(viewString);
        if (views < 20000) {
          rowsToDelete.push(i + 2);
        }
      }
    }

    if (rowsToDelete.length > 0) {
      deleteRowsInBlocks_(sheet, rowsToDelete);
      recalculateSTT(SHEET_VIDEO);
      ui.alert(`✅ ĐÃ DỌN DẸP HOÀN TẤT!\n\nĐã tìm thấy và xóa thành công ${rowsToDelete.length} video có lượt view dưới 20.000.`);
    } else {
      ui.alert(`✅ QUÉT HOÀN TẤT!\n\nKhông tìm thấy video nào có lượt view dưới 20.000.`);
    }
  } catch (error) {
    ui.alert(`❌ HỆ THỐNG GẶP LỖI:\n${error.message}`);
  }
}

// =========================================================================
// TỰ ĐỘNG XÓA KÊNH KHÔNG HOẠT ĐỘNG
// =========================================================================

/**
 * 🧹 Xóa các kênh "chết" (0 views trong 4 tuần / 0 video trong 30 ngày)
 * Sử dụng cột G (Views/4 tuần) và F (Vids/Tháng) làm tiêu chí
 */
function cleanInactiveChannels() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:18px;background:#f8fafc;color:#202124;line-height:1.45;">
      ${buildDialogHeader_("Dọn dẹp kênh không hoạt động", "Xóa các kênh có views/tháng thấp hơn hoặc bằng ngưỡng bạn chọn.", "#dc2626")}
      <div style="background:#fff;border:1px solid #dfe3eb;border-radius:10px;padding:14px;">
        <label style="display:block;font-weight:bold;margin-bottom:6px;">Ngưỡng VIEWS/THÁNG tối thiểu để giữ lại kênh</label>
        <input id="threshold" type="number" min="0" value="0" style="width:100%;box-sizing:border-box;padding:11px;border:2px solid #dc2626;border-radius:8px;font-size:14px;">
        <div style="font-size:12px;color:#6b7280;margin-top:8px;">
          Nhập <b>0</b> để xóa kênh có 0 views/tháng. Nhập <b>1000</b> để xóa kênh có views/tháng <= 1000.
        </div>
        <label style="display:block;margin-top:12px;font-size:13px;"><input type="checkbox" id="confirmDelete"> Tôi hiểu thao tác này sẽ xóa dòng và không thể hoàn tác.</label>
        <div id="status" style="display:none;margin-top:12px;border:1px solid #dfe3eb;border-radius:8px;padding:10px;font-size:13px;white-space:pre-wrap;"></div>
      </div>
      <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px;">
        <button onclick="google.script.host.close()" style="padding:10px 18px;background:#9ca3af;color:#fff;border:none;border-radius:8px;cursor:pointer;">Hủy</button>
        <button id="runBtn" onclick="run()" style="padding:10px 22px;background:#dc2626;color:#fff;border:none;border-radius:8px;font-weight:bold;cursor:pointer;">Dọn dẹp</button>
      </div>
      <script>
        function setStatus(text, isErr){
          const el = document.getElementById('status');
          el.style.display = 'block';
          el.style.borderColor = isErr ? '#fecaca' : '#bfdbfe';
          el.style.background = isErr ? '#fff7f7' : '#eff6ff';
          el.textContent = text;
        }
        function run(){
          const threshold = parseInt(document.getElementById('threshold').value, 10);
          if (isNaN(threshold) || threshold < 0){ setStatus('Ngưỡng không hợp lệ.', true); if(window.dlgProgress) window.dlgProgress.fail('Ngưỡng không hợp lệ'); return; }
          if (!document.getElementById('confirmDelete').checked){ setStatus('Vui lòng tick xác nhận trước khi xóa.', true); if(window.dlgProgress) window.dlgProgress.fail('Chưa xác nhận'); return; }
          document.getElementById('runBtn').disabled = true;
          setStatus('Đang dọn dẹp kênh có views/tháng <= ' + threshold + '...', false);
          if(window.dlgProgress) window.dlgProgress.start(4, 'Đang dọn dẹp kênh ≤ ' + threshold + ' views/tháng');
          google.script.run.withSuccessHandler(function(res){
            setStatus((res.success ? 'Hoàn tất: ' : 'Lỗi: ') + res.message, !res.success);
            document.getElementById('runBtn').disabled = !!res.success;
            if(window.dlgProgress){ res.success ? window.dlgProgress.complete(res.message) : window.dlgProgress.fail('Lỗi: ' + res.message); }
          }).withFailureHandler(function(err){
            setStatus('Lỗi: ' + (err.message || err), true);
            document.getElementById('runBtn').disabled = false;
            if(window.dlgProgress) window.dlgProgress.fail('Lỗi: ' + (err.message || err));
          }).executeCleanInactiveChannels(threshold);
        }
        document.getElementById('threshold').addEventListener('keydown', function(e){ if(e.key==='Enter'){ e.preventDefault(); run(); } });
      </script>
    </div>
  `).setWidth(620).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, "Dọn dẹp kênh không hoạt động");
}

function executeCleanInactiveChannels(threshold) {
  try {
    threshold = parseInt(threshold, 10);
    if (isNaN(threshold) || threshold < 0) throw new Error("Ngưỡng không hợp lệ.");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHANNEL);
    if (!sheet) throw new Error("Không tìm thấy Sheet: " + SHEET_CHANNEL);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, message: "Sheet không có dữ liệu." };
    const viewsData = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
    const rowsToDelete = [];
    for (let i = 0; i < viewsData.length; i++) {
      const v = unformatNumber(viewsData[i][0]);
      if (v <= threshold) rowsToDelete.push(i + 2);
    }
    if (rowsToDelete.length > 0) {
      deleteRowsInBlocks_(sheet, rowsToDelete);
      recalculateSTT(SHEET_CHANNEL);
      return { success: true, message: `Đã xóa ${rowsToDelete.length} kênh có views/tháng <= ${threshold}.` };
    }
    return { success: true, message: `Không có kênh nào cần dọn dẹp với ngưỡng ${threshold}.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 🧹 DỌN DẸP TỔNG HỢP: Chạy cả 2 module:
 *   1. Xóa video > 3 tháng (tái sử dụng checkAndCleanData)
 *   2. Xóa kênh chết (cleanInactiveChannels)
 */
function runFullAutoCleanup() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '🧹 DỌN DẸP TỔNG HỢP',
    'Hệ thống sẽ tự động:\n\n' +
    '1. Xóa các video > 3 tháng (Sheet VIDEO)\n' +
    '2. Xóa các video trùng lặp\n' +
    '3. Cập nhật & xóa kênh có 0 views/tháng (Sheet KÊNH)\n\n' +
    'Quá trình này có thể mất vài phút. Tiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  
  try {
    const apiKey = getApiKey();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const videoSheet = ss.getSheetByName(SHEET_VIDEO);
    let videoCleaned = 0;
    
    if (videoSheet) {
      const lastRow = videoSheet.getLastRow();
      if (lastRow >= 2) {
        const data = videoSheet.getRange(2, 1, lastRow - 1, 3).getValues();
        const videoIds = []; const rowMap = {}; const rowsToDelete = []; const seenIds = new Set();
        
        for (let i = 0; i < data.length; i++) {
          const vId = extractVideoId(data[i][2]);
          if (!vId) continue;
          if (seenIds.has(vId)) rowsToDelete.push(i + 2);
          else { seenIds.add(vId); videoIds.push(vId); rowMap[vId] = i + 2; }
        }
        
        const threeMonthsAgo = new Date(); threeMonthsAgo.setDate(threeMonthsAgo.getDate() - 90);
        for (let i = 0; i < videoIds.length; i += 50) {
          const json = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/videos?part=snippet&id=${videoIds.slice(i, i + 50).join(',')}`, apiKey);
          if (json.items) {
            json.items.forEach(item => {
              if (new Date(item.snippet.publishedAt) < threeMonthsAgo) rowsToDelete.push(rowMap[item.id]);
            });
          }
        }
        
        if (rowsToDelete.length > 0) {
          const uniqueRows = [...new Set(rowsToDelete)].sort((a, b) => b - a);
          deleteRowsInBlocks_(videoSheet, uniqueRows);
          recalculateSTT(SHEET_VIDEO);
          videoCleaned = uniqueRows.length;
        }
      }
    }
    
    let channelCleaned = 0;
    const channelSheet = ss.getSheetByName(SHEET_CHANNEL);
    if (channelSheet) {
      const lastRow = channelSheet.getLastRow();
      if (lastRow >= 2) {
        const viewsData = channelSheet.getRange(2, 7, lastRow - 1, 1).getValues();
        const rowsToDelete = [];
        for (let i = 0; i < viewsData.length; i++) {
          const v = unformatNumber(viewsData[i][0]);
          if (v <= 0) rowsToDelete.push(i + 2);
        }
        if (rowsToDelete.length > 0) {
          deleteRowsInBlocks_(channelSheet, rowsToDelete);
          recalculateSTT(SHEET_CHANNEL);
          channelCleaned = rowsToDelete.length;
        }
      }
    }
    
    ui.alert(
      `✅ DỌN DẸP TỔNG HỢP HOÀN TẤT!\n\n` +
      `📹 Video đã xóa: ${videoCleaned} (quá 3 tháng / trùng)\n` +
      `📺 Kênh đã xóa: ${channelCleaned} (0 views/tháng)\n\n` +
      `💡 GỢI Ý: Trước khi chạy, nên cập nhật Sheet KÊNH để có dữ liệu views/tháng mới nhất.`
    );
  } catch (e) {
    ui.alert(`❌ LỖI: ${e.message}`);
  }
}

// =========================================================================
// CÁC HÀM BỔ TRỢ
// =========================================================================

function recalculateSTT(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const links = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  const sttValues = []; let stt = 1;
  for (let i = 0; i < links.length; i++) {
    if (links[i][0]) sttValues.push([stt++]);
    else sttValues.push([""]);
  }
  sheet.getRange(2, 1, lastRow - 1, 1).setValues(sttValues);
}

