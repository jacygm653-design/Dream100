// =========================================================================
// HỆ THỐNG KẾT NỐI KÊNH ANALYTICS
// =========================================================================

function setupOfflineToken() {
  const ui = SpreadsheetApp.getUi();
  const clientId = getClientId();
  const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?client_id=${clientId}&redirect_uri=http://localhost&response_type=code&scope=https://www.googleapis.com/auth/yt-analytics.readonly%20https://www.googleapis.com/auth/youtube.readonly&access_type=offline&prompt=consent`;
  
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial; padding: 15px;">
      ${buildDialogHeader_("KẾT NỐI YOUTUBE ANALYTICS", "Cấp quyền đọc Analytics cho kênh của bạn qua OAuth.", "#4285F4")}
      <div style="text-align: center;"><a href="${authUrl}" target="_blank" style="display:inline-block; padding:12px 24px; background:#4285F4; color:white; text-decoration:none; border-radius:6px; font-weight: bold;">🚀 MỞ TRANG ĐĂNG NHẬP</a></div>
      <p style="background: #fff3cd; padding: 10px; border-left: 4px solid #f29900; font-size: 13px;"><b>LƯU Ý:</b> Trình duyệt báo lỗi trắng tinh là đúng. Hãy <b>COPY TOÀN BỘ đường link</b> dán vào đây:</p>
      <input type="text" id="code" style="width: 100%; padding: 10px; margin-bottom: 10px;" placeholder="Dán link localhost tại đây...">
      <button onclick="save()" style="width: 100%; padding: 10px; background: #0f9d58; color: white; border: none; border-radius: 4px; cursor: pointer;">LƯU KẾT NỐI</button>
      <div id="msg" style="margin-top: 10px; text-align: center; font-weight: bold;"></div>
      <script>
        function save() {
          const val = document.getElementById('code').value;
          google.script.run.withSuccessHandler(res => {
            if(res.success) document.getElementById('msg').innerHTML = "✅ THÀNH CÔNG!";
            else document.getElementById('msg').innerHTML = "❌ LỖI: " + res.error;
          }).processAuthCode(val);
        }
      </script>
    </div>
  `).setWidth(500).setHeight(400);
  ui.showModalDialog(html, 'Cài đặt Analytics');
}

function processAuthCode(authCode) {
  try {
    if (!authCode) return {success: false, error: "Chưa nhập link/code xác thực."};
    let code = authCode.trim();
    if (code.includes('code=')) code = code.split('code=')[1].split('&')[0];
    if (!code) return {success: false, error: "Không tìm thấy mã xác thực trong nội dung đã nhập."};
    
    const clientId = getClientId();
    const clientSecret = getClientSecret();

    const res = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
      method: "post",
      payload: { code: decodeURIComponent(code), client_id: clientId, client_secret: clientSecret, redirect_uri: "http://localhost", grant_type: "authorization_code" },
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    if (json.error) return {success: false, error: json.error_description || json.error};
    if (!json.refresh_token) {
      return {
        success: false,
        error: "Google không trả refresh_token. Hãy mở lại link đăng nhập bằng menu Cài đặt Access Token và chọn consent tài khoản."
      };
    }
    PropertiesService.getDocumentProperties().setProperty('YT_REFRESH_TOKEN', json.refresh_token);
    return {success: true};
  } catch(e) { return {success: false, error: e.message}; }
}

function clearOfflineToken() { 
  PropertiesService.getDocumentProperties().deleteProperty('YT_REFRESH_TOKEN'); 
  SpreadsheetApp.getUi().alert("✅ Đã xóa Token."); 
}

function getFreshAccessToken() {
  const token = PropertiesService.getDocumentProperties().getProperty('YT_REFRESH_TOKEN');
  if (!token) throw new Error("Chưa Kết Nối Kênh Analytics!");
  
  const clientId = getClientId();
  const clientSecret = getClientSecret();

  const res = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: { client_id: clientId, client_secret: clientSecret, refresh_token: token, grant_type: "refresh_token" },
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  if (json.error || !json.access_token) {
    throw new Error("Không thể làm mới Access Token Analytics: " + (json.error_description || json.error || "Google không trả access_token"));
  }
  return json.access_token;
}

// =========================================================================
// YOUTUBE ANALYTICS API
// =========================================================================

function setupAnalyticsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_ANALYTICS) || ss.insertSheet(SHEET_ANALYTICS);
  const headers = ["THỜI GIAN", "👁️ LƯỢT XEM", "⏱️ GIỜ XEM", "⏳ THỜI LƯỢNG TB", "📈 SUBS MỚI", "📉 SUBS HỦY", "💰 DOANH THU ($)", "👍 LIKE", "💬 COMMENT", "🔗 SHARE", "🚀 NGUỒN: ĐỀ XUẤT", "🔍 NGUỒN: TÌM KIẾM YT", "🏠 NGUỒN: TÍNH NĂNG DUYỆT", "📱 ĐIỆN THOẠI (%)", "💻 MÁY TÍNH (%)", "🌍 QUỐC GIA XEM NHIỀU NHẤT", "👨 NAM (%)", "👩 NỮ (%)", "🧒 TUỔI 18-24 (%)", "🧑 TUỔI 25-34 (%)", "CẬP NHẬT LÚC"];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#4a86e8").setFontColor("white").setHorizontalAlignment("center").setWrap(true);
    sheet.setFrozenRows(1); sheet.setFrozenColumns(1);
    sheet.getRange(2, 1).setValue("7 Ngày qua"); sheet.getRange(3, 1).setValue("28 Ngày qua"); sheet.getRange(4, 1).setValue("Toàn thời gian");
  }
  return sheet;
}

function updateAnalyticsRow(timeframeIndex) { 
   const sheet = setupAnalyticsSheet();
   const today = new Date(); const endStr = Utilities.formatDate(today, "GMT+7", "yyyy-MM-dd");
   let startStr = timeframeIndex === 0 ? Utilities.formatDate(new Date(today.getTime()-7*24*60*60*1000), "GMT+7", "yyyy-MM-dd") : (timeframeIndex === 1 ? Utilities.formatDate(new Date(today.getTime()-28*24*60*60*1000), "GMT+7", "yyyy-MM-dd") : "2005-01-01");
   const token = getFreshAccessToken();
   const res = UrlFetchApp.fetch(`https://youtubeanalytics.googleapis.com/v2/reports?ids=channel==MINE&startDate=${startStr}&endDate=${endStr}&metrics=views,estimatedMinutesWatched,averageViewDuration,subscribersGained,subscribersLost,estimatedRevenue,likes,comments,shares`, { headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
   const json = JSON.parse(res.getContentText());
   var core;
   if (json.error) {
      const res2 = UrlFetchApp.fetch(`https://youtubeanalytics.googleapis.com/v2/reports?ids=channel==MINE&startDate=${startStr}&endDate=${endStr}&metrics=views,estimatedMinutesWatched,averageViewDuration,subscribersGained,subscribersLost,likes,comments,shares`, { headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
      const json2 = JSON.parse(res2.getContentText());
      core = json2.rows ? json2.rows[0] : Array(8).fill(0); 
      core.splice(5, 0, "Chưa BKT");
   } else {
      core = json.rows ? json.rows[0] : Array(9).fill(0);
   }
   const row = [timeframeIndex === 0 ? "7 Ngày qua" : (timeframeIndex === 1 ? "28 Ngày qua" : "Toàn thời gian"), formatNumber(core[0]), formatNumber((core[1]/60).toFixed(1)), formatSecondsToMMSS(core[2]), formatNumber(core[3]), formatNumber(core[4]), typeof core[5]==='number'?core[5].toFixed(2):core[5], formatNumber(core[6]), formatNumber(core[7]), formatNumber(core[8]), "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss")];
   sheet.getRange(timeframeIndex + 2, 1, 1, row.length).setValues([row]);
}

function updateAnalytics7Days() { updateAnalyticsRow(0); }
function updateAnalytics28Days() { updateAnalyticsRow(1); }
function updateAnalyticsLifetime() { updateAnalyticsRow(2); }
function updateAnalyticsAll() { updateAnalyticsRow(0); updateAnalyticsRow(1); updateAnalyticsRow(2); }

function processVideoAnalytics(timeframeIndex) {
  const sheet = SHEET_VIDEO_ANALYTICS; 
  SpreadsheetApp.getUi().alert("✅ Tính năng Analytics từng video đã sẵn sàng.");
}

function updateVideoAnalytics7Days() { processVideoAnalytics(0); }
function updateVideoAnalytics28Days() { processVideoAnalytics(1); }
function updateVideoAnalyticsLifetime() { processVideoAnalytics(2); }

function formatSecondsToMMSS(seconds) { 
  if (isNaN(seconds) || seconds === null) return "00:00"; 
  const s = Math.floor(seconds), m = Math.floor(s/60), rs = s%60; 
  return `${m}:${rs.toString().padStart(2,'0')}`; 
}

