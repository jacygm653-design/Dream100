// ========== PUBLIC WRAPPERS (callable từ google.script.run) ==========
// Apps Script không cho phép google.script.run gọi functions kết thúc bằng "_" (private).
// Các wrapper này delegate sang phiên bản private để client có thể gọi.
function getInlineView(viewName){ return getInlineView_(viewName); }
function getRecentRowsForDisplay(type, count){ return getRecentRowsForDisplay_(type, count); }
function adminRunSilent(actionName, progressId){ return adminRunSilent_(actionName, progressId); }
function getOAuthAuthUrl(){ return getOAuthAuthUrl_(); }
function getSystemHealthData(){ return getSystemHealthData_(); }
function getQuotaSnapshot(){ return getQuotaSnapshot_(); }
function getAdminDashboardStats(){ return getAdminDashboardStats_(); }
// saveApiKeyConfig() và getApiKeyConfig() là public functions — gọi trực tiếp được, không cần wrapper.
// =====================================================================

function adminRunSilent_(actionName, progressId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  function lr(name){ const s = ss.getSheetByName(name); if (!s) throw new Error('Không tìm thấy Sheet: ' + name); return { sh: s, last: s.getLastRow() }; }
  try {
    switch(actionName) {
      case 'updateVideoSheet': { const r = lr(SHEET_VIDEO); if (r.last < 2) return { success: true, message: 'Sheet không có dữ liệu' }; initTaskProgress_(progressId, r.last - 1, 'Chuẩn bị cập nhật video'); executeVideoUpdate(r.sh, 2, r.last, 'full', progressId); finishTaskProgress_(progressId, true, 'Hoàn tất cập nhật video'); return { success: true, message: 'Đã cập nhật full ' + (r.last-1) + ' dòng video. ' + getYouTubeQuotaUsageMessage_() }; }
      case 'updateFastVideoSheet': { const r = lr(SHEET_VIDEO); if (r.last < 2) return { success: true, message: 'Sheet không có dữ liệu' }; initTaskProgress_(progressId, r.last - 1, 'Chuẩn bị cập nhật nhanh video'); executeVideoUpdate(r.sh, 2, r.last, 'fast', progressId); finishTaskProgress_(progressId, true, 'Hoàn tất cập nhật nhanh video'); return { success: true, message: 'Đã cập nhật NHANH ' + (r.last-1) + ' dòng video. ' + getYouTubeQuotaUsageMessage_() }; }
      case 'updateChannelSheet': { const r = lr(SHEET_CHANNEL); if (r.last < 2) return { success: true, message: 'Sheet không có dữ liệu' }; initTaskProgress_(progressId, r.last - 1, 'Chuẩn bị cập nhật kênh'); executeChannelUpdate(r.sh, 2, r.last, 'full', progressId); finishTaskProgress_(progressId, true, 'Hoàn tất cập nhật kênh'); return { success: true, message: 'Đã cập nhật full ' + (r.last-1) + ' dòng kênh. ' + getYouTubeQuotaUsageMessage_() }; }
      case 'updateFastChannelSheet': { const r = lr(SHEET_CHANNEL); if (r.last < 2) return { success: true, message: 'Sheet không có dữ liệu' }; initTaskProgress_(progressId, r.last - 1, 'Chuẩn bị cập nhật nhanh kênh'); executeChannelUpdate(r.sh, 2, r.last, 'fast', progressId); finishTaskProgress_(progressId, true, 'Hoàn tất cập nhật nhanh kênh'); return { success: true, message: 'Đã cập nhật NHANH ' + (r.last-1) + ' dòng kênh. ' + getYouTubeQuotaUsageMessage_() }; }
      case 'fetchAllSubtitles': { const r = lr(SHEET_VIDEO); if (r.last < 2) return { success: true, message: 'Sheet không có dữ liệu' }; initTaskProgress_(progressId, r.last - 1, 'Chuẩn bị lấy subtitle'); const rep = executeSubtitleFetch(r.sh, 2, r.last, true, progressId); finishTaskProgress_(progressId, true, 'Hoàn tất lấy subtitle'); return { success: true, message: rep || 'Đã chạy lấy subtitle xong.' }; }
      case 'retryFailedSubtitles': {
        const r = lr(SHEET_VIDEO); if (r.last < 2) return { success: true, message: 'Sheet không có dữ liệu' };
        const subs = r.sh.getRange(2, COL_VIDEO.SUBTITLE, r.last - 1, 1).getValues();
        const failedRows = []; for (let i = 0; i < subs.length; i++){ const v = String(subs[i][0] || ''); if (v.indexOf('[KHÔNG CÓ SUBTITLE') === 0 || v.indexOf('[LỖI') === 0 || v.indexOf('[ĐÃ DỪNG') === 0) failedRows.push(i + 2); }
        if (!failedRows.length) return { success: true, message: 'Không có dòng FAIL nào để retry.' };
        let okCount = 0;
        initTaskProgress_(progressId, failedRows.length, 'Chuẩn bị retry subtitle lỗi');
        for (let i = 0; i < failedRows.length; i++){ try { executeSubtitleFetch(r.sh, failedRows[i], failedRows[i], true); okCount++; } catch(e) {} updateTaskProgress_(progressId, i + 1, failedRows.length, 'Đã retry ' + (i + 1) + '/' + failedRows.length); }
        finishTaskProgress_(progressId, true, 'Hoàn tất retry subtitle');
        return { success: true, message: 'Đã retry ' + okCount + '/' + failedRows.length + ' dòng FAIL.' };
      }
      case 'clearSubtitleColumn': { const r = lr(SHEET_VIDEO); if (r.last < 2) return { success: true, message: 'Sheet không có dữ liệu' }; r.sh.getRange(2, COL_VIDEO.SUBTITLE, r.last - 1, 1).clearContent(); return { success: true, message: 'Đã xóa cột Subtitle (' + (r.last-1) + ' dòng).' }; }
      case 'checkAndCleanData': { checkAndCleanData(); return { success: true, message: 'Đã chạy: Xóa video > 3 tháng & trùng lặp.' }; }
      case 'deleteLowViewVideos': { deleteLowViewVideos(); return { success: true, message: 'Đã chạy: Xóa video < 20.000 views.' }; }
      case 'runFullAutoCleanup': { runFullAutoCleanup(); return { success: true, message: 'Đã chạy: Dọn dẹp tổng hợp.' }; }
      case 'updateAnalytics7Days': initTaskProgress_(progressId, 1, 'Đang cập nhật Analytics 7 ngày'); updateAnalyticsRow(0); finishTaskProgress_(progressId, true, 'Hoàn tất Analytics 7 ngày'); return { success: true, message: 'Đã cập nhật Analytics 7 ngày.' };
      case 'updateAnalytics28Days': initTaskProgress_(progressId, 1, 'Đang cập nhật Analytics 28 ngày'); updateAnalyticsRow(1); finishTaskProgress_(progressId, true, 'Hoàn tất Analytics 28 ngày'); return { success: true, message: 'Đã cập nhật Analytics 28 ngày.' };
      case 'updateAnalyticsLifetime': initTaskProgress_(progressId, 1, 'Đang cập nhật Analytics toàn thời gian'); updateAnalyticsRow(2); finishTaskProgress_(progressId, true, 'Hoàn tất Analytics toàn thời gian'); return { success: true, message: 'Đã cập nhật Analytics toàn thời gian.' };
      case 'updateAnalyticsAll': initTaskProgress_(progressId, 3, 'Đang cập nhật Analytics 1/3'); updateAnalyticsRow(0); updateTaskProgress_(progressId, 1, 3, 'Đã cập nhật 1/3 mốc'); updateAnalyticsRow(1); updateTaskProgress_(progressId, 2, 3, 'Đã cập nhật 2/3 mốc'); updateAnalyticsRow(2); finishTaskProgress_(progressId, true, 'Hoàn tất 3/3 mốc Analytics'); return { success: true, message: 'Đã cập nhật Analytics tất cả mốc.' };
      case 'updateVideoAnalytics7Days': initTaskProgress_(progressId, 1, 'Đang cập nhật Video Analytics 7 ngày'); processVideoAnalytics(0); finishTaskProgress_(progressId, true, 'Hoàn tất Video Analytics 7 ngày'); return { success: true, message: 'Đã cập nhật Video Analytics 7 ngày.' };
      case 'updateVideoAnalytics28Days': initTaskProgress_(progressId, 1, 'Đang cập nhật Video Analytics 28 ngày'); processVideoAnalytics(1); finishTaskProgress_(progressId, true, 'Hoàn tất Video Analytics 28 ngày'); return { success: true, message: 'Đã cập nhật Video Analytics 28 ngày.' };
      case 'updateVideoAnalyticsLifetime': initTaskProgress_(progressId, 1, 'Đang cập nhật Video Analytics toàn thời gian'); processVideoAnalytics(2); finishTaskProgress_(progressId, true, 'Hoàn tất Video Analytics toàn thời gian'); return { success: true, message: 'Đã cập nhật Video Analytics toàn thời gian.' };
      case 'clearOfflineToken': initTaskProgress_(progressId, 1, 'Đang xóa token'); PropertiesService.getDocumentProperties().deleteProperty('YT_REFRESH_TOKEN'); finishTaskProgress_(progressId, true, 'Đã xóa token'); return { success: true, message: 'Đã xóa refresh token Analytics. Cần chạy lại "Cài đặt Access Token" để kết nối.' };
      default: throw new Error('Action không có silent variant: ' + actionName);
    }
  } catch(e) {
    finishTaskProgress_(progressId, false, 'Lỗi: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getLastNonEmptyRowInColumn_(sheet, column, minRow) {
  const lastRow = sheet.getLastRow();
  const startRow = minRow || 2;
  if (lastRow < startRow) return 0;
  const values = sheet.getRange(startRow, column, lastRow - startRow + 1, 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][0] || '').toString().trim() !== '') return startRow + i;
  }
  return 0;
}

function getRecentRowsForDisplay_(type, count) {
  count = parseInt(count, 10) || 12;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh, sheetName, keyColumn;
  if (type === 'video') { sheetName = SHEET_VIDEO; keyColumn = 3; sh = ss.getSheetByName(SHEET_VIDEO); }
  else if (type === 'channel') { sheetName = SHEET_CHANNEL; keyColumn = 3; sh = ss.getSheetByName(SHEET_CHANNEL); }
  else if (type === 'analytics') { sheetName = SHEET_ANALYTICS; keyColumn = 1; sh = ss.getSheetByName(SHEET_ANALYTICS); }
  else if (type === 'videoAnalytics') { sheetName = SHEET_VIDEO_ANALYTICS; keyColumn = 1; sh = ss.getSheetByName(SHEET_VIDEO_ANALYTICS); }
  else throw new Error('Type không hợp lệ: ' + type);
  if (!sh) return { rows: [], displayRows: [], headers: [], totalRows: 0, sheetName: sheetName };
  const lastDataRow = getLastNonEmptyRowInColumn_(sh, keyColumn, 2);
  if (lastDataRow < 2) return { rows: [], displayRows: [], headers: [], totalRows: 0, sheetName: sheetName };
  const displayColCount = Math.max(14, Math.min(sh.getLastColumn(), 26));
  const headers = sh.getRange(1, 1, 1, displayColCount).getDisplayValues()[0]
    .map(function(h, i){ return (h || ('Cột ' + (i + 1))).toString(); });
  const startRow = Math.max(2, lastDataRow - count + 1);
  const totalRows = lastDataRow - 1;
  const numRows = lastDataRow - startRow + 1;
  const displayValues = sh.getRange(startRow, 1, numRows, displayColCount).getDisplayValues();
  const displayRows = displayValues.map(function(r, idx){
    return { rowNumber: startRow + idx, values: r };
  }).filter(function(r){
    return (r.values[keyColumn - 1] || '').toString().trim() !== '';
  });
  if (type === 'video') {
    const data = sh.getRange(startRow, 1, numRows, displayColCount).getValues();
    return {
      rows: data.map(function(r){
        return {
          stt: r[COL_VIDEO.STT - 1],
          title: r[COL_VIDEO.TITLE - 1],
          link: r[COL_VIDEO.LINK - 1],
          views: r[COL_VIDEO.VIEWS - 1],
          vph: r[COL_VIDEO.VPH - 1],
          duration: r[COL_VIDEO.DURATION - 1],
          thumbnail: r[COL_VIDEO.THUMBNAIL - 1],
          published: r[COL_VIDEO.PUBLISHED - 1]
        };
      }).filter(function(r){ return (r.link || '').toString().trim() !== ''; }),
      displayRows: displayRows,
      headers: headers,
      totalRows: totalRows,
      sheetName: sheetName
    };
  }
  if (type !== 'channel') {
    return {
      rows: displayRows,
      displayRows: displayRows,
      headers: headers,
      totalRows: totalRows,
      sheetName: sheetName
    };
  }
  // channel
  const data = sh.getRange(startRow, 1, numRows, displayColCount).getValues();
  return {
    rows: data.map(function(r){
      return { stt: r[0], name: r[1], link: r[2], subs: r[5], viewsPerMonth: r[6], country: r[10] };
    }).filter(function(r){ return (r.link || '').toString().trim() !== ''; }),
    displayRows: displayRows,
    headers: headers,
    totalRows: totalRows,
    sheetName: sheetName
  };
}

function getInlineView_(viewName) {
  if (viewName === 'searchVideo') return _ivSearchVideo_();
  if (viewName === 'searchChannel') return _ivSearchChannel_();
  if (viewName === 'health') return _ivHealth_();
  if (viewName === 'quota') return _ivQuota_();
  if (viewName === 'cleanup') return _ivCleanup_();
  if (viewName === 'guide') return _ivGuide_();
  if (viewName === 'subtitleGuide') return _ivSubtitleGuide_();
  if (viewName === 'oauth') return _ivOAuth_();
  if (viewName === 'fetch') return _ivFetch_();
  if (viewName === 'ai') return _ivAI_();
  if (viewName === 'apiKey') return _ivApiKey_();
  if (viewName.indexOf('range:') === 0) return _ivRange_(viewName.substring(6));
  if (viewName.indexOf('single:') === 0) return _ivSingleRow_(viewName.substring(7));
  if (viewName.indexOf('direct:') === 0) return _ivDirect_(viewName.substring(7));
  throw new Error('Inline view không khả dụng: ' + viewName);
}

function _ivSharedStyles_() {
  return `
    <style>
      @keyframes ivSpin{to{transform:rotate(360deg)}}
      @keyframes ivPulse{0%,100%{opacity:0.7;transform:scale(1)}50%{opacity:1;transform:scale(1.04)}}
      @keyframes ivSlideShim{0%{margin-left:-30%}100%{margin-left:100%}}
      .iv-card{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:14px;box-shadow:0 4px 12px -6px rgba(15,23,42,.10)}
      .iv-form{background:linear-gradient(135deg,#f8fafc,#eef2ff);border:1px solid #dbe4ff;border-radius:16px;padding:18px;box-shadow:inset 0 1px 0 rgba(255,255,255,.85),0 12px 26px -20px rgba(79,70,229,.40)}
      .iv-row{display:grid;gap:12px;margin-bottom:12px;align-items:start}
      .iv-row.c1{grid-template-columns:1fr}
      .iv-row.c4{grid-template-columns:repeat(4,1fr)}
      .iv-row.c3{grid-template-columns:repeat(3,1fr)}
      .iv-row.c2{grid-template-columns:repeat(2,1fr)}
      .iv-row label{display:grid;grid-template-columns:max-content minmax(0,1fr);align-items:center;column-gap:4px;row-gap:6px;min-width:0;font-size:10.5px;font-weight:800;color:#475569;letter-spacing:0.04em;text-transform:uppercase;line-height:1.25}
      .iv-row label > span{display:inline-flex;align-items:center;align-self:center;color:#dc2626 !important;font-weight:900;line-height:1}
      .iv-row label > input,.iv-row label > select,.iv-row label > textarea,.iv-row label > button{grid-column:1/-1}
      .iv-row label.check{display:flex;flex-direction:row;align-items:center;gap:8px;text-transform:none;letter-spacing:0;font-size:12.5px;color:#0f172a;font-weight:650;line-height:1.35}
      .iv-row label.check input{grid-column:auto;width:16px;height:16px;flex:0 0 auto;margin:0}
      .iv-row input,.iv-row select{padding:9px 11px;border:1px solid #d1d5db;border-radius:9px;font-size:13px;background:#fff;color:#0f172a;font-family:inherit;outline:none;transition:all .2s ease;width:100%;box-sizing:border-box}
      .iv-row input:focus,.iv-row select:focus{border-color:#6366f1;box-shadow:0 0 0 4px rgba(99,102,241,.16)}
      .iv-actions{display:flex;align-items:center;gap:14px;margin-top:8px;flex-wrap:wrap}
      .iv-primary{padding:11px 22px;color:#fff;border:0;border-radius:10px;cursor:pointer;font-weight:800;font-size:13.5px;letter-spacing:0.02em;font-family:inherit;transition:all .2s ease}
      .iv-primary:hover{filter:brightness(1.08);transform:translateY(-1px)}
      .iv-primary:disabled{opacity:.55;cursor:not-allowed;transform:none}
      .iv-hint{font-size:11.5px;color:#64748b;font-weight:600;line-height:1.55}
      .iv-loader{margin-top:16px;background:linear-gradient(180deg,#fff,#fafbfd);border:1px solid #e5e7eb;border-radius:14px;padding:36px 18px;display:flex;flex-direction:column;align-items:center;gap:10px;animation:dlgFadeIn .3s ease both}
      .iv-spin{width:60px;height:60px;border-radius:50%;border:5px solid #e2e8f0;border-top-color:#7c3aed;border-right-color:#ec4899;animation:ivSpin .9s linear infinite;box-shadow:0 0 24px -4px rgba(124,58,237,.32)}
      .iv-spin-text{font-size:15px;font-weight:800;color:#0f172a;letter-spacing:-0.01em;animation:ivPulse 1.6s ease-in-out infinite}
      .iv-spin-sub{font-size:12px;color:#64748b;font-weight:600;text-align:center;max-width:520px;line-height:1.5}
      .iv-spin-meter{width:320px;height:6px;background:#e2e8f0;border-radius:6px;overflow:hidden;margin-top:6px;position:relative}
      .iv-spin-meter > div{position:absolute;left:0;top:0;width:30%;height:100%;background:linear-gradient(90deg,transparent,#7c3aed,#ec4899,transparent);animation:ivSlideShim 1.5s ease-in-out infinite}
      .iv-result{margin-top:16px;background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:14px;animation:dlgSlideUp .35s ease both}
      .iv-summary{padding:11px 14px;border-radius:10px;font-size:13px;font-weight:600;line-height:1.55;margin-bottom:12px;white-space:pre-wrap}
      .iv-summary.ok{background:linear-gradient(135deg,#d1fae5,#a7f3d0);color:#065f46;border:1px solid #6ee7b7}
      .iv-summary.err{background:linear-gradient(135deg,#fee2e2,#fecaca);color:#991b1b;border:1px solid #fca5a5}
      .iv-result h3{margin:8px 0;font-size:14px;color:#0f172a;font-weight:800}
      .iv-table-wrap{max-height:340px;overflow:auto;border:1px solid #e5e7eb;border-radius:10px}
      .iv-table{width:100%;border-collapse:separate;border-spacing:0;font-size:12px}
      .iv-table th{position:sticky;top:0;z-index:1;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;padding:9px 10px;text-align:left;font-weight:700;letter-spacing:0.02em;font-size:10.5px;text-transform:uppercase}
      .iv-table td{padding:8px 10px;border-bottom:1px solid #f1f5f9;color:#334155;vertical-align:middle}
      .iv-table tr:hover td{background:#f8fafc}
      .iv-thumb{width:88px;height:50px;border-radius:6px;object-fit:cover;background:#0f172a;display:block;border:1px solid #e5e7eb}
      .iv-num{font-weight:700;color:#0f172a;font-feature-settings:"tnum";white-space:nowrap}
      @media(max-width:980px){.iv-row.c4,.iv-row.c3{grid-template-columns:repeat(2,1fr)}}
      @media(max-width:680px){.iv-row.c4,.iv-row.c3,.iv-row.c2{grid-template-columns:1fr}.iv-actions{flex-direction:column;align-items:stretch}.iv-primary{width:100%}}
      .iv-info-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin:12px 0}
      .iv-info-card{background:linear-gradient(135deg,#fff,#f8fafc);border:1px solid #e5e7eb;border-radius:12px;padding:12px 14px;transition:all .2s ease}
      .iv-info-card:hover{transform:translateY(-2px);box-shadow:0 8px 16px -8px rgba(15,23,42,.16)}
      .iv-info-card .lbl{font-size:10px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px}
      .iv-info-card .val{font-size:22px;font-weight:900;color:#0f172a;letter-spacing:-0.02em;line-height:1.1}
      .iv-info-card .sub{font-size:11.5px;color:#94a3b8;font-weight:600;margin-top:2px}
      .iv-info-card.ok .val{color:#059669}
      .iv-info-card.warn .val{color:#d97706}
      .iv-info-card.danger .val{color:#dc2626}
      .iv-bar{display:inline-flex;flex-wrap:wrap;gap:6px;margin-top:8px}
      .iv-tag{font-size:10.5px;font-weight:800;padding:3px 9px;border-radius:999px;letter-spacing:0.04em;text-transform:uppercase}
      .iv-tag.ok{background:#d1fae5;color:#065f46}
      .iv-tag.warn{background:#fef3c7;color:#92400e}
      .iv-tag.danger{background:#fee2e2;color:#991b1b}
      .iv-row-item{display:grid;grid-template-columns:auto 1fr auto;gap:10px;align-items:center;background:#fff;border:1px solid #e5e7eb;border-radius:11px;padding:10px 12px;margin-bottom:6px;animation:dlgSlideUp .3s ease both;transition:all .2s ease}
      .iv-row-item:hover{transform:translateX(3px);box-shadow:0 6px 14px -8px rgba(15,23,42,.16)}
      .iv-row-item .ico{width:30px;height:30px;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;font-size:14px;font-weight:900;flex-shrink:0}
      .iv-row-item.ok .ico{background:linear-gradient(135deg,#d1fae5,#a7f3d0);color:#065f46}
      .iv-row-item.warn .ico{background:linear-gradient(135deg,#fef3c7,#fde68a);color:#92400e}
      .iv-row-item.danger .ico{background:linear-gradient(135deg,#fee2e2,#fecaca);color:#991b1b}
      .iv-row-item .grp{font-size:9.5px;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:#94a3b8}
      .iv-row-item .lbl{font-size:13px;font-weight:700;color:#0f172a}
      .iv-row-item .det{font-size:11.5px;color:#64748b;margin-top:1px}
      .iv-tip{background:linear-gradient(135deg,#fef3c7,#fde68a);border:1px solid #fcd34d;border-radius:11px;padding:10px 14px;font-size:12.5px;color:#78350f;line-height:1.55;margin:10px 0}
      .iv-quota-ring{position:relative;width:140px;height:140px;flex-shrink:0}
      .iv-quota-ring svg{width:100%;height:100%;transform:rotate(-90deg)}
      .iv-quota-center{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center}
      .iv-quota-pct{font-size:26px;font-weight:900;color:#0f172a;letter-spacing:-0.02em}
      .iv-quota-lbl{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.08em}
    </style>
    <script>
      function ivEscCell_(s){return String(s == null ? '' : s).replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);});}
      function ivRenderSheetTable(d, opts){
        opts = opts || {};
        var rows = (d && d.displayRows) || [];
        var headers = (d && d.headers) || [];
        if (!rows.length) {
          return '<div style="padding:18px;color:#64748b;font-weight:600;">'+ivEscCell_(opts.emptyText || 'Sheet chưa có dữ liệu hiển thị.')+'</div>';
        }
        rows = rows.slice().reverse();
        var maxCols = headers.length || (rows[0] && rows[0].values ? rows[0].values.length : 0);
        var html = '<div class="iv-table-wrap"><table class="iv-table"><thead><tr><th>Dòng</th>';
        for (var i=0;i<maxCols;i++) html += '<th>'+ivEscCell_(headers[i] || ('Cột '+(i+1)))+'</th>';
        html += '</tr></thead><tbody>';
        rows.forEach(function(r){
          html += '<tr><td class="iv-num">'+ivEscCell_(r.rowNumber || '')+'</td>';
          for (var i=0;i<maxCols;i++){
            var raw = (r.values && r.values[i] != null) ? String(r.values[i]) : '';
            var head = String(headers[i] || '').toLowerCase();
            var safe = ivEscCell_(raw);
            var cell = safe || '<span style="color:#cbd5e1;">–</span>';
            if (/^https?:\\/\\//i.test(raw)) {
              if (head.indexOf('thumb') >= 0 || /\\.(jpg|jpeg|png|webp)(\\?|$)/i.test(raw)) {
                cell = '<a target="_blank" rel="noopener" href="'+safe+'"><img class="iv-thumb" loading="lazy" src="'+safe+'"></a>';
              } else {
                cell = '<a target="_blank" rel="noopener" href="'+safe+'" style="color:#1d4ed8;font-weight:700;text-decoration:none;">'+safe+'</a>';
              }
            }
            html += '<td style="max-width:340px;white-space:normal;word-break:break-word;">'+cell+'</td>';
          }
          html += '</tr>';
        });
        html += '</tbody></table></div><div style="font-size:11.5px;color:#64748b;margin-top:8px;font-weight:600;">Hiển thị '+rows.length+' dòng mới nhất / Tổng <b>'+Number((d && d.totalRows) || 0).toLocaleString('vi-VN')+'</b> dòng có link trong Sheet <b>'+ivEscCell_((d && d.sheetName) || '')+'</b></div>';
        return html;
      }
      function ivMakeProgressId(prefix){
        return (prefix || 'task') + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
      }
      function ivStartProgressPoll(progressId, fallbackTotal, label){
        if (!progressId || !window.dlgProgress) return null;
        window.__activeProgressId = progressId;
        try { google.script.run.setActiveAdminProgress(progressId); } catch(e) {}
        var total = Math.max(1, parseInt(fallbackTotal, 10) || 1);
        window.dlgProgress.set(0, label || ('Đang xử lý 0/' + total));
        var timer = setInterval(function(){
          google.script.run
            .withSuccessHandler(function(p){
              if (!p) return;
              var pct = Number(p.pct || 0);
              var text = p.label || ('Đã xử lý ' + (p.done || 0) + '/' + (p.total || total));
              window.dlgProgress.set(pct, text);
              if (p.status === 'done' || p.status === 'error') clearInterval(timer);
            })
            .withFailureHandler(function(){})
            .getTaskProgress(progressId);
        }, 900);
        return timer;
      }
      function ivStopProgressPoll(timer){
        if (timer) clearInterval(timer);
      }
    </script>
  `;
}

function _ivSearchVideo_() {
  return _ivSharedStyles_() + `
    <div class="iv-form">
      <div class="iv-row c4">
        <label>Chủ đề <span style="color:#dc2626;">*</span><input id="svTopic" placeholder="psychology · human behavior · mindset"></label>
        <label>Số ngày gần đây<input id="svDays" type="number" value="90" min="1" max="365"></label>
        <label>Min views<input id="svMinViews" type="number" value="100000"></label>
        <label>Min duration (phút)<input id="svMinDur" type="number" value="7" min="0"></label>
      </div>
      <div class="iv-row c4">
        <label>Sắp xếp theo<select id="svOrder"><option value="viewCount" selected>Phổ biến nhất (Views)</option><option value="relevance">Liên quan nhất</option><option value="date">Mới nhất</option><option value="rating">Đánh giá cao nhất</option></select></label>
        <label>Số kết quả/quốc gia<input id="svMaxResults" type="number" value="50" min="5" max="50"></label>
        <label>Quốc gia ƯU TIÊN<input id="svCountries" value="US,GB" placeholder="VD: US,GB,CA"></label>
        <label>Quốc gia LOẠI BỎ<input id="svExcludeCountries" value="IN" placeholder="VD: IN,PK"></label>
      </div>
      <div class="iv-row c3">
        <label>Ngôn ngữ ưu tiên<select id="svLanguage"><option value="en" selected>Tiếng Anh (en)</option><option value="">Bất kỳ</option><option value="vi">Tiếng Việt</option><option value="es">Tiếng Tây Ban Nha</option><option value="fr">Tiếng Pháp</option><option value="ja">Tiếng Nhật</option><option value="ko">Tiếng Hàn</option></select></label>
        <label class="check"><input type="checkbox" id="svExcludeShorts" checked> Loại bỏ Shorts (&lt; 60s)</label>
        <label class="check"><input type="checkbox" id="svAutoUpdate" checked> Auto update metadata sau khi thêm</label>
      </div>
      <div class="iv-actions">
        <button id="svBtn" class="iv-primary" style="background:linear-gradient(135deg,#7c3aed,#ec4899);box-shadow:0 6px 14px -4px rgba(124,58,237,.42);" onclick="svRun()">🔍 BẮT ĐẦU TÌM</button>
        <span class="iv-hint">⚠️ <code style="background:#fff7ed;padding:1px 6px;border-radius:4px;border:1px solid #fdba74;">search.list</code> tốn ~100 quota/lần · sẽ thêm vào cuối Sheet ${SHEET_VIDEO}</span>
      </div>
    </div>
    <div id="svLoader" class="iv-loader" style="display:none;">
      <div class="iv-spin"></div>
      <div class="iv-spin-text">Đang tìm video...</div>
      <div class="iv-spin-sub" id="svSub">Gọi YouTube Data API · search.list + videos.list · ước lượng 20–60 giây</div>
      <div class="iv-spin-meter"><div></div></div>
    </div>
    <div id="svResult" class="iv-result" style="display:none;"></div>
    <script>
      function svFmt(n){ if (n==null||n==='') return '–'; var v=parseInt(n,10); return isNaN(v)?String(n):v.toLocaleString('vi-VN'); }
      function svEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function svFmtPub(p){ if (!p) return '–'; var s = String(p); return s.length > 40 ? s.substring(0,40)+'…' : s; }
      function svRun(){
        var topic = (document.getElementById('svTopic').value||'').trim();
        if (!topic){ alert('Vui lòng nhập chủ đề.'); document.getElementById('svTopic').focus(); return; }
        var minMin = parseInt(document.getElementById('svMinDur').value,10) || 3;
        var params = {
          topic: topic,
          daysAgo: parseInt(document.getElementById('svDays').value,10) || 90,
          minViews: parseInt(document.getElementById('svMinViews').value,10) || 0,
          minDurationMin: minMin,
          orderBy: document.getElementById('svOrder').value || 'viewCount',
          maxResults: Math.min(50, Math.max(5, parseInt(document.getElementById('svMaxResults').value,10) || 50)),
          includeCountries: (document.getElementById('svCountries').value||'').split(',').map(function(s){return s.trim().toUpperCase();}).filter(Boolean),
          excludeCountries: (document.getElementById('svExcludeCountries').value||'').split(',').map(function(s){return s.trim().toUpperCase();}).filter(Boolean),
          relevanceLanguage: document.getElementById('svLanguage').value,
          excludeShorts: document.getElementById('svExcludeShorts').checked,
          autoUpdate: document.getElementById('svAutoUpdate').checked,
          progressId: ivMakeProgressId('search_video')
        };
        var btn = document.getElementById('svBtn');
        btn.disabled = true;
        document.getElementById('svResult').style.display='none';
        document.getElementById('svLoader').style.display='flex';
        var progressPoll = ivStartProgressPoll(params.progressId, Math.max(1, params.includeCountries.length), 'Đang tìm video theo '+params.includeCountries.length+' quốc gia');
        var t0 = Date.now();
        var subEl = document.getElementById('svSub');
        var timer = setInterval(function(){
          var sec = Math.floor((Date.now()-t0)/1000);
          subEl.textContent = 'Đã chạy '+sec+'s · search.list + videos.list · params: '+params.maxResults+' results, min '+svFmt(params.minViews)+' views';
        }, 500);
        google.script.run
          .withSuccessHandler(function(res){
            clearInterval(timer); ivStopProgressPoll(progressPoll);
            document.getElementById('svLoader').style.display='none';
            btn.disabled = false;
            svShow(res);
            if (window.dlgProgress) { res.success ? window.dlgProgress.complete('Hoàn tất search') : window.dlgProgress.fail('Lỗi: '+res.message); }
            if (typeof loadStats === 'function') loadStats();
          })
          .withFailureHandler(function(err){
            clearInterval(timer); ivStopProgressPoll(progressPoll);
            document.getElementById('svLoader').style.display='none';
            btn.disabled = false;
            svShow({success:false, message: err.message || String(err)});
            if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
          })
          .executeSearchByTopic(params);
      }
      function svShow(res){
        var box = document.getElementById('svResult');
        box.style.display='block';
        var html = '<div class="iv-summary '+(res.success?'ok':'err')+'">'+(res.success?'✅ ':'❌ ')+svEsc(res.message||'(không có thông tin)')+'</div>';
        if (res.stats) {
          var s = res.stats;
          html += '<div class="iv-info-grid" style="margin:10px 0;">'
            + '<div class="iv-info-card"><div class="lbl">Ứng viên</div><div class="val">'+svFmt(s.found||0)+'</div><div class="sub">video tìm thấy</div></div>'
            + '<div class="iv-info-card"><div class="lbl">Ứng viên mới</div><div class="val">'+svFmt(s.newCandidates||0)+'</div><div class="sub">chưa có trong Sheet</div></div>'
            + '<div class="iv-info-card ok"><div class="lbl">Đạt lọc</div><div class="val">'+svFmt(s.passed||0)+'</div><div class="sub">video hợp lệ</div></div>'
            + '<div class="iv-info-card ok"><div class="lbl">Đã thêm</div><div class="val">'+svFmt(s.added||0)+'</div><div class="sub">link mới</div></div>'
            + '</div>';
          html += '<div class="iv-tip" style="background:#fff;border-color:#e5e7eb;color:#475569;">Lọc bỏ: <b>'+svFmt(s.tooFewViews||0)+'</b> views thấp · <b>'+svFmt(s.tooShort||0)+'</b> quá ngắn · <b>'+svFmt(s.isShort||0)+'</b> Shorts · <b>'+svFmt(s.excludedCountry||0)+'</b> quốc gia loại bỏ · <b>'+svFmt(s.isLive||0)+'</b> live/upcoming · <b>'+svFmt(s.missing||0)+'</b> thiếu dữ liệu.</div>';
        }
        html += '<h3>📥 Video gần nhất trong Sheet '+svEsc('${SHEET_VIDEO}')+'</h3><div id="svRecent">⏳ Đang đọc Sheet...</div>';
        box.innerHTML = html;
        google.script.run
          .withSuccessHandler(function(d){ svRender(d); })
          .withFailureHandler(function(err){ document.getElementById('svRecent').innerHTML = '<div style="padding:14px;color:#dc2626;font-weight:600;">Lỗi đọc Sheet: '+svEsc(err.message||err)+'</div>'; })
          .getRecentRowsForDisplay('video', 15);
      }
      function svRender(d){
        document.getElementById('svRecent').innerHTML = ivRenderSheetTable(d, {emptyText:'Sheet chưa có link video ở cột C để hiển thị.'});
      }
      document.getElementById('svTopic').focus();
      document.getElementById('svTopic').addEventListener('keydown', function(e){ if(e.key==='Enter') svRun(); });
    </script>
  `;
}

function _ivSearchChannel_() {
  return _ivSharedStyles_() + `
    <div class="iv-form">
      <div class="iv-row c4">
        <label>Chủ đề kênh <span style="color:#dc2626;">*</span><input id="scTopic" placeholder="psychology · self help · mindset"></label>
        <label>Min subscribers<input id="scMinSubs" type="number" value="10000"></label>
        <label>Min views/tháng<input id="scMinViewsMonth" type="number" value="500000"></label>
        <label>Số kết quả/quốc gia<input id="scMaxResults" type="number" value="50" min="5" max="50"></label>
      </div>
      <div class="iv-row c4">
        <label>Sắp xếp theo<select id="scOrder"><option value="viewCount" selected>Phổ biến (Views)</option><option value="relevance">Liên quan</option></select></label>
        <label>Quốc gia ƯU TIÊN<input id="scCountries" value="US,GB" placeholder="VD: US,GB,CA"></label>
        <label>Quốc gia LOẠI BỎ<input id="scExcludeCountries" value="IN" placeholder="VD: IN,PK"></label>
        <label>Ngôn ngữ<select id="scLanguage"><option value="en" selected>Tiếng Anh</option><option value="">Bất kỳ</option><option value="vi">Tiếng Việt</option></select></label>
      </div>
      <div class="iv-row c1">
        <label class="check"><input type="checkbox" id="scAutoUpdate" checked> Auto update metadata sau khi thêm</label>
      </div>
      <div class="iv-actions">
        <button id="scBtn" class="iv-primary" style="background:linear-gradient(135deg,#0f9d58,#10b981);box-shadow:0 6px 14px -4px rgba(15,157,88,.42);" onclick="scRun()">📺 BẮT ĐẦU TÌM KÊNH</button>
        <span class="iv-hint">⚠️ <code style="background:#fff7ed;padding:1px 6px;border-radius:4px;border:1px solid #fdba74;">search.list</code> tốn ~100 quota/lần · sẽ thêm vào cuối Sheet ${SHEET_CHANNEL}</span>
      </div>
    </div>
    <div id="scLoader" class="iv-loader" style="display:none;">
      <div class="iv-spin" style="border-top-color:#0f9d58;border-right-color:#10b981;"></div>
      <div class="iv-spin-text">Đang tìm kênh...</div>
      <div class="iv-spin-sub" id="scSub">Gọi YouTube Data API · search.list + channels.list · ước lượng 20–60s</div>
      <div class="iv-spin-meter"><div style="background:linear-gradient(90deg,transparent,#0f9d58,#10b981,transparent);"></div></div>
    </div>
    <div id="scResult" class="iv-result" style="display:none;"></div>
    <script>
      function scFmt(n){ if (n==null||n==='') return '–'; var v=parseInt(n,10); return isNaN(v)?String(n):v.toLocaleString('vi-VN'); }
      function scEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function scRun(){
        var topic = (document.getElementById('scTopic').value||'').trim();
        if (!topic){ alert('Vui lòng nhập chủ đề kênh.'); document.getElementById('scTopic').focus(); return; }
        var params = {
          topic: topic,
          minSubs: parseInt(document.getElementById('scMinSubs').value,10) || 0,
          minMonthlyViews: parseInt(document.getElementById('scMinViewsMonth').value,10) || 0,
          orderBy: document.getElementById('scOrder').value || 'viewCount',
          maxResults: Math.min(50, Math.max(5, parseInt(document.getElementById('scMaxResults').value,10) || 50)),
          includeCountries: (document.getElementById('scCountries').value||'').split(',').map(function(s){return s.trim().toUpperCase();}).filter(Boolean),
          excludeCountries: (document.getElementById('scExcludeCountries').value||'').split(',').map(function(s){return s.trim().toUpperCase();}).filter(Boolean),
          relevanceLanguage: document.getElementById('scLanguage').value,
          autoUpdate: document.getElementById('scAutoUpdate').checked,
          progressId: ivMakeProgressId('search_channel')
        };
        var btn = document.getElementById('scBtn'); btn.disabled = true;
        document.getElementById('scResult').style.display='none';
        document.getElementById('scLoader').style.display='flex';
        var progressPoll = ivStartProgressPoll(params.progressId, Math.max(1, params.includeCountries.length), 'Đang tìm kênh theo '+params.includeCountries.length+' quốc gia');
        var t0 = Date.now();
        var timer = setInterval(function(){ var sec = Math.floor((Date.now()-t0)/1000); document.getElementById('scSub').textContent = 'Đã chạy '+sec+'s · params: '+params.maxResults+' kết quả/quốc gia, min '+scFmt(params.minSubs)+' subs, min '+scFmt(params.minMonthlyViews)+' views/tháng'; }, 500);
        google.script.run
          .withSuccessHandler(function(res){
            clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('scLoader').style.display='none'; btn.disabled = false;
            scShow(res);
            if (window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất') : window.dlgProgress.fail('Lỗi: '+res.message); }
            if (typeof loadStats === 'function') loadStats();
          })
          .withFailureHandler(function(err){
            clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('scLoader').style.display='none'; btn.disabled = false;
            scShow({success:false, message: err.message || String(err)});
            if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
          })
          .executeSearchChannelsByTopic(params);
      }
      function scShow(res){
        var box = document.getElementById('scResult');
        box.style.display='block';
        var html = '<div class="iv-summary '+(res.success?'ok':'err')+'">'+(res.success?'✅ ':'❌ ')+scEsc(res.message||'(không có thông tin)')+'</div>';
        if (res.stats) {
          var s = res.stats;
          html += '<div class="iv-info-grid" style="margin:10px 0;">'
            + '<div class="iv-info-card"><div class="lbl">Ứng viên</div><div class="val">'+scFmt(s.found||0)+'</div><div class="sub">kênh tìm thấy</div></div>'
            + '<div class="iv-info-card"><div class="lbl">Ứng viên mới</div><div class="val">'+scFmt(s.newCandidates||0)+'</div><div class="sub">chưa có trong Sheet</div></div>'
            + '<div class="iv-info-card ok"><div class="lbl">Đạt lọc</div><div class="val">'+scFmt(s.passed||0)+'</div><div class="sub">kênh hợp lệ</div></div>'
            + '<div class="iv-info-card ok"><div class="lbl">Đã thêm</div><div class="val">'+scFmt(s.added||0)+'</div><div class="sub">link mới</div></div>'
            + '</div>';
          html += '<div class="iv-tip" style="background:#fff;border-color:#e5e7eb;color:#475569;">Lọc bỏ: <b>'+scFmt(s.tooFewSubs||0)+'</b> subs thấp · <b>'+scFmt(s.tooFewMonthlyViews||0)+'</b> views/tháng thấp · <b>'+scFmt(s.excludedCountry||0)+'</b> quốc gia loại bỏ · <b>'+scFmt(s.hidden||0)+'</b> ẩn subs · <b>'+scFmt(s.monthlyViewErrors||0)+'</b> lỗi tính views/tháng.</div>';
        }
        html += '<h3>📺 Kênh gần nhất trong Sheet</h3><div id="scRecent">⏳ Đang đọc Sheet...</div>';
        box.innerHTML = html;
        google.script.run
          .withSuccessHandler(function(d){ scRender(d); })
          .withFailureHandler(function(err){ document.getElementById('scRecent').innerHTML = '<div style="padding:14px;color:#dc2626;font-weight:600;">Lỗi đọc Sheet: '+scEsc(err.message||err)+'</div>'; })
          .getRecentRowsForDisplay('channel', 15);
      }
      function scRender(d){
        document.getElementById('scRecent').innerHTML = ivRenderSheetTable(d, {emptyText:'Sheet chưa có link kênh ở cột C để hiển thị.'});
      }
      document.getElementById('scTopic').focus();
      document.getElementById('scTopic').addEventListener('keydown', function(e){ if(e.key==='Enter') scRun(); });
    </script>
  `;
}

function _ivHealth_() {
  return _ivSharedStyles_() + `
    <div id="ivhRoot">
      <div id="ivhLoading" class="iv-loader"><div class="iv-spin" style="border-top-color:#10b981;border-right-color:#0ea5e9;"></div><div class="iv-spin-text">Đang chẩn đoán hệ thống...</div><div class="iv-spin-sub">Kiểm tra API key, sheet, OAuth, quota, nguồn subtitle</div><div class="iv-spin-meter"><div style="background:linear-gradient(90deg,transparent,#10b981,#0ea5e9,transparent);"></div></div></div>
      <div id="ivhBody" style="display:none;"></div>
    </div>
    <script>
      (function(){
        function esc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
        google.script.run.withSuccessHandler(function(d){
          var tone = d.dangerCount>0?'KÉM':(d.warnCount>0?'KHÁ':'TỐT');
          var color = d.dangerCount>0?'#dc2626':(d.warnCount>0?'#d97706':'#059669');
          var html = '<div class="iv-info-grid">'
            + '<div class="iv-info-card"><div class="lbl">Health Score</div><div class="val" style="color:'+color+';">'+d.score+'</div><div class="sub">Tình trạng: <b style="color:'+color+';">'+tone+'</b></div></div>'
            + '<div class="iv-info-card ok"><div class="lbl">OK</div><div class="val">'+d.okCount+'</div><div class="sub">mục đạt yêu cầu</div></div>'
            + '<div class="iv-info-card warn"><div class="lbl">Cảnh báo</div><div class="val">'+d.warnCount+'</div><div class="sub">có thể cải thiện</div></div>'
            + '<div class="iv-info-card danger"><div class="lbl">Nghiêm trọng</div><div class="val">'+d.dangerCount+'</div><div class="sub">cần xử lý gấp</div></div>'
            + '</div>'
            + '<div style="font-size:11.5px;color:#64748b;margin-bottom:10px;font-weight:600;">Lần kiểm tra cuối: '+esc(d.ts)+'</div>'
            + '<h3 style="margin:6px 0 10px;font-size:14px;color:#0f172a;font-weight:800;">📋 Chi tiết các mục</h3>';
          d.checks.forEach(function(c, i){
            var icon = c.severity==='ok'?'✓':(c.severity==='warn'?'!':'×');
            var tag = c.severity==='ok'?'OK':(c.severity==='warn'?'CẢNH BÁO':'NGHIÊM TRỌNG');
            html += '<div class="iv-row-item '+c.severity+'" style="animation-delay:'+(i*25)+'ms;"><div class="ico">'+icon+'</div><div><div class="grp">'+esc(c.group)+'</div><div class="lbl">'+esc(c.label)+'</div><div class="det">'+esc(c.detail)+'</div></div><span class="iv-tag '+c.severity+'">'+tag+'</span></div>';
          });
          html += '<div style="margin-top:14px;display:flex;gap:8px;"><button onclick="ivhReload()" class="iv-primary" style="background:linear-gradient(135deg,#10b981,#059669);">🔄 Kiểm tra lại</button></div>';
          document.getElementById('ivhLoading').style.display='none';
          var b = document.getElementById('ivhBody');
          b.style.display='block';
          b.innerHTML = html;
        }).withFailureHandler(function(err){
          document.getElementById('ivhLoading').innerHTML = '<div style="padding:30px;text-align:center;color:#dc2626;font-weight:700;">Lỗi: '+esc(err.message||err)+'</div>';
        }).getSystemHealthData();
      })();
      window.ivhReload = function(){
        document.getElementById('ivhBody').style.display='none';
        document.getElementById('ivhLoading').style.display='flex';
        google.script.run.withSuccessHandler(function(d){
          // Re-render via server view to keep DRY — easier: reload entire inline view
          if (typeof loadInlineView === 'function') loadInlineView('health', '🩺 Kiểm tra sức khỏe hệ thống');
        }).withFailureHandler(function(){}).getSystemHealthData();
      };
    </script>
  `;
}

function _ivQuota_() {
  return _ivSharedStyles_() + `
    <div id="ivqRoot">
      <div id="ivqLoading" class="iv-loader"><div class="iv-spin"></div><div class="iv-spin-text">Đang đọc snapshot quota...</div><div class="iv-spin-meter"><div></div></div></div>
      <div id="ivqBody" style="display:none;"></div>
    </div>
    <script>
      (function(){
        function esc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
        function fmt(n){ return Number(n).toLocaleString('vi-VN'); }
        var ivqTimer = null;
        function ivqRender(s){
          var color = s.level==='danger'?'#dc2626':(s.level==='warn'?'#d97706':'#7c3aed');
          var pctClamp = Math.min(100, s.percent);
          var dash = (326.7 - (pctClamp/100)*326.7).toFixed(2);
          var html = ''
            + '<div style="display:flex;gap:18px;align-items:center;background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:18px;box-shadow:0 6px 18px -10px rgba(15,23,42,.10);margin-bottom:14px;">'
            + '  <div class="iv-quota-ring"><svg viewBox="0 0 120 120"><circle cx="60" cy="60" r="52" fill="none" stroke="#eef2f7" stroke-width="14"/><circle cx="60" cy="60" r="52" fill="none" stroke="url(#ivqGrad)" stroke-width="14" stroke-linecap="round" stroke-dasharray="326.7" stroke-dashoffset="'+dash+'" style="transition:stroke-dashoffset 1s cubic-bezier(.2,.7,.2,1);"/><defs><linearGradient id="ivqGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#7c3aed"/><stop offset="100%" stop-color="#ec4899"/></linearGradient></defs></svg><div class="iv-quota-center"><div class="iv-quota-pct">'+s.percent.toFixed(1)+'%</div><div class="iv-quota-lbl">Đã dùng</div></div></div>'
            + '  <div style="flex:1;">'
            + '    <div style="font-size:12px;font-weight:700;color:'+color+';text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px;">'+(s.level==='danger'?'🚨 NGUY HIỂM':(s.level==='warn'?'⚠️ CẢNH BÁO':'✅ AN TOÀN'))+'</div>'
            + '    <div style="font-size:14px;color:#334155;font-weight:600;line-height:1.55;">'+esc(s.tip)+'</div>'
            + '    <div style="margin-top:10px;display:flex;flex-wrap:wrap;gap:8px;font-size:11.5px;font-weight:600;color:#64748b;"><span>📍 PT: '+esc(s.nowPT)+'</span><span>⏱️ Reset sau: <b>'+esc(s.resetIn)+'</b></span></div>'
            + '  </div>'
            + '</div>'
            + '<div class="iv-info-grid">'
            + '<div class="iv-info-card"><div class="lbl">Đã dùng</div><div class="val">'+fmt(s.used)+'</div><div class="sub">/ '+fmt(s.limit)+' units</div></div>'
            + '<div class="iv-info-card '+(s.remaining<=0?'danger':'ok')+'"><div class="lbl">Còn lại</div><div class="val">'+fmt(s.remaining)+'</div><div class="sub">units khả dụng</div></div>'
            + '<div class="iv-info-card"><div class="lbl">Số request</div><div class="val">'+fmt(s.requests)+'</div><div class="sub">đã ghi nhận</div></div>'
            + '<div class="iv-info-card"><div class="lbl">Avg cost</div><div class="val">'+s.avgCostPerReq+'</div><div class="sub">units / request</div></div>'
            + '</div>'
            + '<div class="iv-card" style="margin-top:6px;"><div style="display:flex;justify-content:space-between;font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:8px;"><span>Mức tiêu thụ</span><span style="color:#7c3aed;">'+s.percent.toFixed(1)+'%</span></div><div style="position:relative;height:14px;background:linear-gradient(180deg,#eef2f7,#e2e8f0);border-radius:8px;overflow:hidden;border:1px solid #d8dee7;"><div style="position:absolute;left:0;top:0;bottom:0;width:'+pctClamp+'%;background:linear-gradient(90deg,#10b981,#7c3aed,#ec4899);background-size:200% 100%;animation:dlgShimmer 2s linear infinite;border-radius:8px;transition:width 1s cubic-bezier(.2,.7,.2,1);"></div><div style="position:absolute;left:50%;top:0;bottom:0;width:1px;background:rgba(15,23,42,0.18);"></div><div style="position:absolute;left:80%;top:0;bottom:0;width:1px;background:rgba(220,38,38,0.4);"></div></div><div style="display:flex;justify-content:space-between;font-size:10.5px;color:#94a3b8;margin-top:5px;font-weight:600;"><span>0</span><span>50% (cảnh báo)</span><span>80% (nguy)</span><span>'+fmt(s.limit)+'</span></div></div>'
            + '<div class="iv-tip"><b>💡 Tham khảo chi phí:</b> <code>search.list</code> ~100 units · <code>videos.list</code>/<code>channels.list</code>/<code>playlistItems.list</code> ~1 unit. Mặc định '+fmt(s.limit)+' units/ngày/project (reset 0:00 PT).</div>';
          var b = document.getElementById('ivqBody');
          b.innerHTML = html;
        }
        function ivqLoad(){
          google.script.run.withSuccessHandler(function(s){
            document.getElementById('ivqLoading').style.display='none';
            document.getElementById('ivqBody').style.display='block';
            ivqRender(s);
          }).withFailureHandler(function(err){
            document.getElementById('ivqLoading').innerHTML = '<div style="padding:30px;text-align:center;color:#dc2626;font-weight:700;">Lỗi: '+esc(err.message||err)+'</div>';
          }).getQuotaSnapshot();
        }
        ivqLoad();
        ivqTimer = setInterval(ivqLoad, 10000);
      })();
    </script>
  `;
}

function _ivRange_(action) {
  const META = {
    videoRange: { sheet: SHEET_VIDEO, title: 'Cập nhật KHOẢNG dòng VIDEO', icon: '🎯', color: '#1a73e8', recent: 'video', etaPerRow: 0.7, hint: 'Sẽ cập nhật metadata video (tiêu đề, tags, views, VPH, duration, thumbnail, ngày đăng, mô tả).' },
    channelRange: { sheet: SHEET_CHANNEL, title: 'Cập nhật KHOẢNG dòng KÊNH', icon: '🎯', color: '#0f9d58', recent: 'channel', etaPerRow: 0.7, hint: 'Sẽ cập nhật metadata kênh (tên, subs, views/tháng, mô tả, country, category, monetization).' },
    subtitleRange: { sheet: SHEET_VIDEO, title: 'Lấy SUBTITLE theo KHOẢNG', icon: '🎯', color: '#0891b2', recent: 'video', etaPerRow: 1.5, hint: 'Lấy transcript cho từng video trong khoảng. Ưu tiên Supadata (B5) → YT-Transcript.io (B6) → fallback.' }
  };
  const m = META[action];
  if (!m) throw new Error('Range action không hỗ trợ: ' + action);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(m.sheet);
  const lastRow = sh ? sh.getLastRow() : 0;
  return _ivSharedStyles_() + `
    <div class="iv-form">
      <div style="margin-bottom:10px;font-size:13px;color:#475569;font-weight:600;">Sheet <b>${m.sheet}</b> · ${lastRow >= 2 ? 'có dữ liệu đến dòng <b>'+lastRow+'</b>' : '<span style="color:#dc2626;">chưa có dữ liệu để xử lý</span>'}</div>
      <div class="iv-row c2">
        <label>Khoảng dòng cần chạy <span style="color:#dc2626;">*</span><input id="rRange" value="${lastRow >= 2 ? '2-' + lastRow : ''}" placeholder="VD: 5-20 hoặc 10"></label>
        <label>Quick chọn<select id="rQuick" onchange="rApplyQuick(this.value)">
          <option value="">— Chọn nhanh —</option>
          ${lastRow >= 2 ? '<option value="2-'+lastRow+'">Toàn bộ ('+(lastRow-1)+' dòng)</option>' : ''}
          ${lastRow >= 12 ? '<option value="2-11">10 dòng đầu</option>' : ''}
          ${lastRow >= 52 ? '<option value="2-51">50 dòng đầu</option>' : ''}
          ${lastRow >= 102 ? '<option value="2-101">100 dòng đầu</option>' : ''}
          ${lastRow >= 11 ? '<option value="'+(lastRow-9)+'-'+lastRow+'">10 dòng cuối</option>' : ''}
          ${lastRow >= 51 ? '<option value="'+(lastRow-49)+'-'+lastRow+'">50 dòng cuối</option>' : ''}
        </select></label>
      </div>
      <div class="iv-tip" style="margin:8px 0 12px;">${m.hint}<br>Nhập <code>5-20</code> để chạy từ dòng 5 → 20, hoặc nhập <code>10</code> để chạy từ dòng 10 đến cuối.</div>
      <div class="iv-actions">
        <button id="rBtn" class="iv-primary" style="background:linear-gradient(135deg,${m.color},${m.color}d9);box-shadow:0 6px 14px -4px ${m.color}66;" onclick="rRun()">▶ BẮT ĐẦU</button>
        <span class="iv-hint">Khoảng lớn → tốn quota cao, có thể timeout 6 phút. Apps Script sẽ tự dừng nếu vượt giới hạn.</span>
      </div>
    </div>
    <div id="rLoader" class="iv-loader" style="display:none;"><div class="iv-spin" style="border-top-color:${m.color};border-right-color:${m.color};"></div><div class="iv-spin-text">Đang xử lý...</div><div class="iv-spin-sub" id="rSub">Khởi tạo...</div><div class="iv-spin-meter"><div></div></div></div>
    <div id="rResult" class="iv-result" style="display:none;"></div>
    <script>
      function rEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function rFmt(n){ if (n==null||n==='') return '–'; var v=parseInt(n,10); return isNaN(v)?String(n):v.toLocaleString('vi-VN'); }
      function rApplyQuick(v){ if (v) document.getElementById('rRange').value = v; }
      function rRun(){
        var raw = (document.getElementById('rRange').value||'').trim();
        if (!raw){ alert('Vui lòng nhập khoảng dòng.'); return; }
        var start, end;
        if (raw.indexOf('-') >= 0){ var p = raw.split('-'); start = parseInt(p[0],10); end = parseInt(p[1],10); }
        else { start = parseInt(raw,10); end = ${lastRow}; }
        if (!start || !end || start < 2 || end < start || end > ${lastRow}){ alert('Khoảng dòng không hợp lệ. Bắt đầu >= 2, kết thúc <= ${lastRow}.'); return; }
        var rows = end - start + 1;
        var btn = document.getElementById('rBtn'); btn.disabled = true;
        document.getElementById('rResult').style.display='none';
        document.getElementById('rLoader').style.display='flex';
        var progressId = ivMakeProgressId('range');
        var progressPoll = ivStartProgressPoll(progressId, rows, 'Đang xử lý 0/'+rows+' dòng');
        var t0 = Date.now();
        var timer = setInterval(function(){ var s = Math.floor((Date.now()-t0)/1000); document.getElementById('rSub').textContent = 'Đã chạy '+s+'s · '+rows+' dòng (' +start+'-'+end+ ') · sheet ${m.sheet}'; }, 500);
        google.script.run
          .withSuccessHandler(function(res){ clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('rLoader').style.display='none'; btn.disabled = false; rShow(res, rows); if (window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất '+rows+' dòng') : window.dlgProgress.fail('Lỗi: '+res.message); } if (typeof loadStats === 'function') loadStats(); })
          .withFailureHandler(function(err){ clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('rLoader').style.display='none'; btn.disabled = false; rShow({success:false, message: err.message || String(err)}, rows); if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err)); })
          .executeRangeAction({ action: '${action}', sheetName: '${m.sheet}', start: start, end: end, progressId: progressId });
      }
      function rShow(res, rows){
        var box = document.getElementById('rResult'); box.style.display='block';
        var html = '<div class="iv-summary '+(res.success?'ok':'err')+'">'+(res.success?'✅ ':'❌ ')+rEsc(res.message||'(không có thông tin)')+'</div>';
        html += '<h3>📥 Dữ liệu mới nhất trong Sheet '+rEsc('${m.sheet}')+'</h3><div id="rRecent">⏳ Đang đọc Sheet...</div>';
        box.innerHTML = html;
        if (res.success) google.script.run.withSuccessHandler(rRender${m.recent==='channel'?'Channel':''}).withFailureHandler(function(err){ document.getElementById('rRecent').innerHTML='<div style="padding:10px;color:#dc2626;">Lỗi đọc: '+rEsc(err.message||err)+'</div>'; }).getRecentRowsForDisplay('${m.recent}', 12);
      }
      function rRender(d){
        document.getElementById('rRecent').innerHTML = ivRenderSheetTable(d, {emptyText:'Sheet chưa có link video ở cột C để hiển thị.'});
      }
      function rRenderChannel(d){
        document.getElementById('rRecent').innerHTML = ivRenderSheetTable(d, {emptyText:'Sheet chưa có link kênh ở cột C để hiển thị.'});
      }
      var __rInp = document.getElementById('rRange'); if (__rInp){ __rInp.focus(); __rInp.addEventListener('keydown', function(e){ if(e.key==='Enter') rRun(); }); }
    </script>
  `;
}

function _ivSingleRow_(action) {
  const META = {
    videoSingle: { sheet: SHEET_VIDEO, title: 'Cập nhật 1 DÒNG VIDEO', color: '#1a73e8', recent: 'video' },
    channelSingle: { sheet: SHEET_CHANNEL, title: 'Cập nhật 1 DÒNG KÊNH', color: '#0f9d58', recent: 'channel' },
    subtitleSingle: { sheet: SHEET_VIDEO, title: 'Lấy SUBTITLE 1 DÒNG', color: '#0891b2', recent: 'video' }
  };
  const m = META[action];
  if (!m) throw new Error('Single action không hỗ trợ: ' + action);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(m.sheet);
  const lastRow = sh ? sh.getLastRow() : 0;
  return _ivSharedStyles_() + `
    <div class="iv-form">
      <div style="margin-bottom:10px;font-size:13px;color:#475569;font-weight:600;">Sheet <b>${m.sheet}</b> · ${lastRow >= 2 ? 'có dữ liệu đến dòng <b>'+lastRow+'</b>' : '<span style="color:#dc2626;">chưa có dữ liệu</span>'}</div>
      <div class="iv-row c2">
        <label>Số dòng cần chạy <span style="color:#dc2626;">*</span><input id="srRow" type="number" min="2" max="${lastRow}" placeholder="VD: 5"></label>
        <label>&nbsp;<button id="srBtn" class="iv-primary" style="background:linear-gradient(135deg,${m.color},${m.color}d9);box-shadow:0 6px 14px -4px ${m.color}66;height:42px;" onclick="srRun()">▶ CHẠY</button></label>
      </div>
      <div class="iv-tip" style="margin:6px 0 0;">Nhập đúng số dòng có link/dữ liệu cần xử lý. Hữu ích khi vừa thêm 1 link mới hoặc fix dòng lỗi.</div>
    </div>
    <div id="srLoader" class="iv-loader" style="display:none;"><div class="iv-spin" style="border-top-color:${m.color};border-right-color:${m.color};"></div><div class="iv-spin-text">Đang xử lý...</div><div class="iv-spin-sub" id="srSub">Khởi tạo...</div><div class="iv-spin-meter"><div></div></div></div>
    <div id="srResult" class="iv-result" style="display:none;"></div>
    <script>
      function srEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function srFmt(n){ if (n==null||n==='') return '–'; var v=parseInt(n,10); return isNaN(v)?String(n):v.toLocaleString('vi-VN'); }
      function srRun(){
        var row = parseInt(document.getElementById('srRow').value, 10);
        if (!row || row < 2 || row > ${lastRow}){ alert('Số dòng không hợp lệ. Phải từ 2 đến ${lastRow}.'); return; }
        var btn = document.getElementById('srBtn'); btn.disabled = true;
        document.getElementById('srResult').style.display='none';
        document.getElementById('srLoader').style.display='flex';
        var progressId = ivMakeProgressId('single');
        var progressPoll = ivStartProgressPoll(progressId, 1, 'Đang xử lý dòng '+row);
        var t0 = Date.now();
        var timer = setInterval(function(){ var s = Math.floor((Date.now()-t0)/1000); document.getElementById('srSub').textContent = 'Đã chạy '+s+'s · dòng '+row+' của '+ '${m.sheet}'; }, 500);
        google.script.run
          .withSuccessHandler(function(res){ clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('srLoader').style.display='none'; btn.disabled = false; srShow(res, row); if (window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất dòng '+row) : window.dlgProgress.fail('Lỗi: '+res.message); } if (typeof loadStats === 'function') loadStats(); })
          .withFailureHandler(function(err){ clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('srLoader').style.display='none'; btn.disabled = false; srShow({success:false, message: err.message || String(err)}, row); if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err)); })
          .executeSingleRowAction({ action: '${action}', sheetName: '${m.sheet}', row: row, progressId: progressId });
      }
      function srShow(res, row){
        var box = document.getElementById('srResult'); box.style.display='block';
        var html = '<div class="iv-summary '+(res.success?'ok':'err')+'">'+(res.success?'✅ ':'❌ ')+srEsc(res.message||'(không có thông tin)')+'</div><div style="font-size:12px;color:#64748b;font-weight:600;">Đã xử lý xong dòng <b>'+row+'</b> trong Sheet <b>${m.sheet}</b>.</div>';
        html += '<h3>📥 Dữ liệu mới nhất trong Sheet</h3><div id="srRecent">⏳ Đang đọc Sheet...</div>';
        box.innerHTML = html;
        if (res.success) google.script.run.withSuccessHandler(function(d){ document.getElementById('srRecent').innerHTML = ivRenderSheetTable(d, {emptyText:'Sheet chưa có link ở cột C để hiển thị.'}); }).withFailureHandler(function(err){ document.getElementById('srRecent').innerHTML='<div style="padding:10px;color:#dc2626;">Lỗi đọc: '+srEsc(err.message||err)+'</div>'; }).getRecentRowsForDisplay('${m.recent}', 12);
      }
      var __srInp = document.getElementById('srRow'); if (__srInp){ __srInp.focus(); __srInp.addEventListener('keydown', function(e){ if(e.key==='Enter') srRun(); }); }
    </script>
  `;
}

function _ivDirect_(action) {
  const META = {
    updateVideoSheet: { title: 'Cập nhật TẤT CẢ video', desc: 'Cập nhật full metadata cho TOÀN BỘ video trong Sheet ' + SHEET_VIDEO + '. Tốn quota cao — chỉ chạy khi cần.', danger: true, color: '#1a73e8', recent: 'video', etaSec: 60 },
    updateFastVideoSheet: { title: 'Cập nhật NHANH Views & VPH', desc: 'Chỉ làm mới views + VPH cho toàn bộ video. Ít quota, phù hợp dùng hằng ngày.', color: '#0ea5e9', recent: 'video', etaSec: 30 },
    updateChannelSheet: { title: 'Cập nhật TẤT CẢ kênh', desc: 'Cập nhật full metadata cho TOÀN BỘ kênh. Tốn quota cao.', danger: true, color: '#0f9d58', recent: 'channel', etaSec: 60 },
    updateFastChannelSheet: { title: 'Cập nhật NHANH Subs & Views/Tháng', desc: 'Chỉ làm mới subs + views/tháng. Ít quota.', color: '#22c55e', recent: 'channel', etaSec: 30 },
    fetchAllSubtitles: { title: 'Lấy Subtitle TẤT CẢ video', desc: 'Lấy transcript cho toàn bộ video. Ưu tiên Supadata (B5) → YT-Transcript.io (B6). Có thể chạy lâu nếu nhiều video.', danger: true, color: '#0891b2', recent: 'video', etaSec: 120 },
    retryFailedSubtitles: { title: 'Thử lại các dòng FAIL', desc: 'Quét cột Subtitle, retry các dòng có "[KHÔNG CÓ SUBTITLE]" hoặc "[LỖI]".', color: '#0891b2', recent: 'video', etaSec: 30 },
    clearSubtitleColumn: { title: 'Xóa Subtitle (Cột M)', desc: 'Xóa toàn bộ nội dung cột M trong Sheet ' + SHEET_VIDEO + '. KHÔNG THỂ HOÀN TÁC.', danger: true, color: '#dc2626', recent: 'video', etaSec: 5 },
    runFullAutoCleanup: { title: 'DỌN DẸP TỔNG HỢP', desc: 'Chạy chuỗi: xóa video > 3 tháng + trùng lặp + xóa kênh chết. Backup trước nếu cần.', danger: true, color: '#dc2626', recent: 'video', etaSec: 30 },
    checkAndCleanData: { title: 'Xóa video > 3 tháng & trùng lặp', desc: 'Xóa video cũ và video bị trùng. Yêu cầu cập nhật ngày đăng trước.', danger: true, color: '#dc2626', recent: 'video', etaSec: 15 },
    deleteLowViewVideos: { title: 'Xóa video < 20.000 views', desc: 'Giữ lại video có views >= 20.000. Yêu cầu cập nhật views trước.', danger: true, color: '#dc2626', recent: 'video', etaSec: 15 },
    updateAnalytics7Days: { title: 'Analytics KÊNH 7 ngày', desc: 'Báo cáo Analytics tổng quan cho 7 ngày qua. Yêu cầu OAuth đã kết nối.', color: '#4285F4', recent: 'analytics', etaSec: 10 },
    updateAnalytics28Days: { title: 'Analytics KÊNH 28 ngày', desc: 'Báo cáo Analytics tổng quan cho 28 ngày qua.', color: '#4285F4', recent: 'analytics', etaSec: 10 },
    updateAnalyticsLifetime: { title: 'Analytics KÊNH toàn thời gian', desc: 'Báo cáo Analytics tổng quan toàn thời gian.', color: '#4285F4', recent: 'analytics', etaSec: 10 },
    updateAnalyticsAll: { title: 'Analytics KÊNH TẤT CẢ', desc: 'Cập nhật cả 3 mốc 7d + 28d + lifetime liên tiếp.', color: '#4285F4', recent: 'analytics', etaSec: 30 },
    updateVideoAnalytics7Days: { title: 'Analytics TỪNG VIDEO 7 ngày', desc: 'Analytics chi tiết theo từng video, 7 ngày.', color: '#3b82f6', recent: 'videoAnalytics', etaSec: 30 },
    updateVideoAnalytics28Days: { title: 'Analytics TỪNG VIDEO 28 ngày', desc: 'Analytics chi tiết theo từng video, 28 ngày.', color: '#3b82f6', recent: 'videoAnalytics', etaSec: 30 },
    updateVideoAnalyticsLifetime: { title: 'Analytics TỪNG VIDEO toàn thời gian', desc: 'Analytics chi tiết theo từng video, toàn thời gian.', color: '#3b82f6', recent: 'videoAnalytics', etaSec: 30 },
    clearOfflineToken: { title: 'Xóa Token Analytics', desc: 'Xóa refresh token OAuth. Sau khi xóa, cần chạy lại "Cài đặt Access Token" để dùng Analytics.', danger: true, color: '#dc2626', etaSec: 1 }
  };
  const m = META[action];
  if (!m) throw new Error('Direct action không hỗ trợ: ' + action);
  return _ivSharedStyles_() + `
    <div class="iv-form" style="display:flex;gap:18px;align-items:center;">
      <div style="width:64px;height:64px;border-radius:14px;background:linear-gradient(135deg,${m.color}20,#fff);border:1px solid ${m.color}40;display:inline-flex;align-items:center;justify-content:center;font-size:30px;flex-shrink:0;box-shadow:inset 0 1px 0 #fff;">${m.danger ? '⚠️' : '▶️'}</div>
      <div style="flex:1;">
        <div style="font-size:18px;font-weight:800;color:#0f172a;margin-bottom:4px;letter-spacing:-0.01em;">${m.title}</div>
        <div style="font-size:13px;color:#475569;font-weight:500;line-height:1.55;">${m.desc}</div>
        <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
          ${m.danger ? '<span class="iv-tag danger">⚠️ HÀNH ĐỘNG NẶNG</span>' : ''}
          <span class="iv-tag" style="background:#f1f5f9;color:#475569;">⏱️ ETA ~ ${m.etaSec}s</span>
          <span class="iv-tag" style="background:#eef2ff;color:#3730a3;">🔧 Silent run · không pop modal</span>
        </div>
      </div>
    </div>
    ${m.danger ? '<div style="margin:10px 0;background:linear-gradient(135deg,#fee2e2,#fecaca);border:1px solid #fca5a5;border-radius:12px;padding:11px 14px;"><label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:12.5px;color:#991b1b;font-weight:700;"><input id="dCfm" type="checkbox" style="width:16px;height:16px;cursor:pointer;"> Tôi xác nhận muốn chạy hành động này. Một số tác vụ KHÔNG THỂ HOÀN TÁC.</label></div>' : ''}
    <div class="iv-actions" style="margin-top:14px;">
      <button id="dBtn" class="iv-primary" style="background:linear-gradient(135deg,${m.color},${m.color}d9);box-shadow:0 6px 14px -4px ${m.color}66;" onclick="dRun()">▶ CHẠY NGAY</button>
      <button class="iv-primary" style="background:#fff;color:#475569;border:1px solid #e5e7eb;box-shadow:none;" onclick="goHome()">← Hủy & quay lại</button>
    </div>
    <div id="dLoader" class="iv-loader" style="display:none;"><div class="iv-spin" style="border-top-color:${m.color};border-right-color:${m.color};"></div><div class="iv-spin-text">Đang chạy ${m.title}...</div><div class="iv-spin-sub" id="dSub">Khởi tạo silent runner</div><div class="iv-spin-meter"><div></div></div></div>
    <div id="dResult" class="iv-result" style="display:none;"></div>
    <script>
      function dEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function dFmt(n){ if (n==null||n==='') return '–'; var v=parseInt(n,10); return isNaN(v)?String(n):v.toLocaleString('vi-VN'); }
      function dRun(){
        ${m.danger ? 'if (!document.getElementById("dCfm").checked){ alert("Vui lòng tick xác nhận trước khi chạy."); return; }' : ''}
        var btn = document.getElementById('dBtn'); btn.disabled = true;
        document.getElementById('dResult').style.display='none';
        document.getElementById('dLoader').style.display='flex';
        var progressId = ivMakeProgressId('direct');
        var progressPoll = ivStartProgressPoll(progressId, 1, 'Đang chạy: ${m.title.replace(/'/g, "\\'")}');
        var t0 = Date.now();
        var timer = setInterval(function(){ var s = Math.floor((Date.now()-t0)/1000); document.getElementById('dSub').textContent = 'Đã chạy '+s+'s · ETA ~ ${m.etaSec}s'; }, 500);
        google.script.run
          .withSuccessHandler(function(res){ clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('dLoader').style.display='none'; btn.disabled = false; dShow(res); if (window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất') : window.dlgProgress.fail('Lỗi: '+res.message); } if (typeof loadStats === 'function') loadStats(); })
          .withFailureHandler(function(err){ clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('dLoader').style.display='none'; btn.disabled = false; dShow({success:false, message: err.message || String(err)}); if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err)); })
          .adminRunSilent('${action}', progressId);
      }
      function dShow(res){
        var box = document.getElementById('dResult'); box.style.display='block';
        var html = '<div class="iv-summary '+(res.success?'ok':'err')+'">'+(res.success?'✅ ':'❌ ')+dEsc(res.message||'(không có thông tin)')+'</div>';
        ${m.recent ? `html += '<h3>📥 Dữ liệu mới nhất</h3><div id="dRecent">⏳ Đang đọc Sheet...</div>'; box.innerHTML = html; if (res.success) google.script.run.withSuccessHandler(dRenderAny).withFailureHandler(function(err){ document.getElementById('dRecent').innerHTML='<div style="padding:10px;color:#dc2626;">Lỗi đọc: '+dEsc(err.message||err)+'</div>'; }).getRecentRowsForDisplay('${m.recent}', 12);` : 'box.innerHTML = html;'}
      }
      ${m.recent ? `function dRenderAny(d){ document.getElementById('dRecent').innerHTML = ivRenderSheetTable(d, {emptyText:'Sheet kết quả chưa có dữ liệu để hiển thị.'}); }` : ''}
    </script>
  `;
}

function _ivAI_() {
  return _ivSharedStyles_() + `
    <div class="iv-form" style="display:flex;gap:18px;align-items:center;">
      <div style="width:68px;height:68px;border-radius:16px;background:linear-gradient(135deg,#ede9fe,#fff);border:1px solid #c4b5fd;display:inline-flex;align-items:center;justify-content:center;font-size:32px;flex-shrink:0;box-shadow:inset 0 1px 0 #fff;">🤖</div>
      <div style="flex:1;">
        <div style="font-size:20px;font-weight:900;color:#111827;letter-spacing:-0.02em;">AI phân tích Sheet</div>
        <div style="margin-top:5px;color:#475569;font-size:13px;font-weight:600;line-height:1.55;">Phân tích dữ liệu bằng 9router, chọn model, sheet, cột, kết hợp lịch sử và lưu kết quả thành công vào sheet ẩn <b>${AI_MEMORY_SHEET}</b>.</div>
      </div>
    </div>
    <div class="iv-info-grid" style="margin-top:12px;">
      <div class="iv-info-card"><div class="lbl">Cài đặt</div><div class="val">B7</div><div class="sub">9router API key/token</div></div>
      <div class="iv-info-card"><div class="lbl">Dữ liệu</div><div class="val">Sheet</div><div class="sub">chọn toàn bộ hoặc từng cột</div></div>
      <div class="iv-info-card"><div class="lbl">Lịch sử</div><div class="val">Ẩn</div><div class="sub">xem/xóa/tái sử dụng</div></div>
    </div>
    <div class="iv-tip">AI Analyzer là cửa sổ lớn 3 tab để đủ không gian cấu hình, hoạt động và xem lịch sử. Bấm nút dưới đây để mở, dashboard sẽ tự đóng sau khi khởi chạy.</div>
    <div class="iv-actions">
      <button id="aiOpenBtn" class="iv-primary" style="background:linear-gradient(135deg,#7c3aed,#ec4899);box-shadow:0 6px 14px -4px rgba(124,58,237,.42);" onclick="aiOpen()">🤖 MỞ AI PHÂN TÍCH SHEET</button>
      <button class="iv-primary" style="background:#fff;color:#475569;border:1px solid #e5e7eb;box-shadow:none;" onclick="goHome()">← Quay lại</button>
    </div>
    <div id="aiStatus" class="iv-result" style="display:none;"></div>
    <script>
      function aiOpen(){
        var btn = document.getElementById('aiOpenBtn');
        btn.disabled = true;
        document.getElementById('aiStatus').style.display = 'block';
        document.getElementById('aiStatus').innerHTML = '<div class="iv-summary ok">⏳ Đang mở AI phân tích Sheet...</div>';
        if (window.dlgProgress) window.dlgProgress.start(2, 'Đang mở AI phân tích Sheet');
        google.script.run
          .withSuccessHandler(function(){
            if (window.dlgProgress) window.dlgProgress.complete('Đã mở AI phân tích Sheet');
            setTimeout(function(){ google.script.host.close(); }, 300);
          })
          .withFailureHandler(function(err){
            btn.disabled = false;
            document.getElementById('aiStatus').innerHTML = '<div class="iv-summary err">❌ Lỗi: '+String((err && err.message) || err)+'</div>';
            if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+String((err && err.message) || err));
          })
          .openAISheetAnalyzer();
      }
    </script>
  `;
}

function _ivCleanup_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHANNEL);
  const lastRow = sh ? sh.getLastRow() : 0;
  return _ivSharedStyles_() + `
    <div class="iv-form">
      <div style="margin-bottom:10px;font-size:13px;color:#475569;font-weight:600;">Sheet <b>${SHEET_CHANNEL}</b> · ${lastRow >= 2 ? '<b>'+(lastRow-1)+'</b> kênh' : '<span style="color:#dc2626;">chưa có dữ liệu</span>'}</div>
      <div class="iv-row c2">
        <label>Ngưỡng VIEWS/THÁNG tối thiểu để giữ lại<input id="cThresh" type="number" min="0" value="0"></label>
        <label>Quick chọn<select onchange="document.getElementById('cThresh').value = this.value;">
          <option value="0">0 (xóa kênh chết)</option>
          <option value="1000">1.000</option>
          <option value="10000">10.000</option>
          <option value="100000">100.000</option>
          <option value="500000">500.000</option>
        </select></label>
      </div>
      <div class="iv-tip" style="margin:8px 0 0;">Nhập <code>0</code> để xóa kênh có 0 views/tháng. Nhập <code>1000</code> để xóa kênh có views/tháng &lt;= 1000.</div>
      <div style="margin:10px 0;background:linear-gradient(135deg,#fee2e2,#fecaca);border:1px solid #fca5a5;border-radius:12px;padding:11px 14px;"><label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:12.5px;color:#991b1b;font-weight:700;"><input id="cCfm" type="checkbox" style="width:16px;height:16px;cursor:pointer;"> Tôi xác nhận: thao tác này XÓA HÀNG và KHÔNG THỂ HOÀN TÁC.</label></div>
      <div class="iv-actions">
        <button id="cBtn" class="iv-primary" style="background:linear-gradient(135deg,#dc2626,#b91c1c);box-shadow:0 6px 14px -4px rgba(220,38,38,.42);" onclick="cRun()">💀 DỌN DẸP</button>
        <button class="iv-primary" style="background:#fff;color:#475569;border:1px solid #e5e7eb;box-shadow:none;" onclick="goHome()">← Hủy & quay lại</button>
      </div>
    </div>
    <div id="cLoader" class="iv-loader" style="display:none;"><div class="iv-spin" style="border-top-color:#dc2626;border-right-color:#b91c1c;"></div><div class="iv-spin-text">Đang dọn dẹp...</div><div class="iv-spin-sub">Đọc cột views/tháng và xóa hàng phù hợp</div><div class="iv-spin-meter"><div style="background:linear-gradient(90deg,transparent,#dc2626,#fca5a5,transparent);"></div></div></div>
    <div id="cResult" class="iv-result" style="display:none;"></div>
    <script>
      function cRun(){
        var t = parseInt(document.getElementById('cThresh').value, 10);
        if (isNaN(t) || t < 0){ alert('Ngưỡng không hợp lệ.'); return; }
        if (!document.getElementById('cCfm').checked){ alert('Vui lòng tick xác nhận trước khi xóa.'); return; }
        var btn = document.getElementById('cBtn'); btn.disabled = true;
        document.getElementById('cResult').style.display='none';
        document.getElementById('cLoader').style.display='flex';
        if (window.dlgProgress) window.dlgProgress.start(5, 'Đang dọn dẹp kênh ≤ '+t+' views/tháng');
        google.script.run
          .withSuccessHandler(function(res){ document.getElementById('cLoader').style.display='none'; btn.disabled = false; document.getElementById('cResult').style.display='block'; document.getElementById('cResult').innerHTML = '<div class="iv-summary '+(res.success?'ok':'err')+'">'+(res.success?'✅ ':'❌ ')+(res.message||'')+'</div>'; if (window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất') : window.dlgProgress.fail('Lỗi: '+res.message); } if (typeof loadStats === 'function') loadStats(); })
          .withFailureHandler(function(err){ document.getElementById('cLoader').style.display='none'; btn.disabled = false; document.getElementById('cResult').style.display='block'; document.getElementById('cResult').innerHTML = '<div class="iv-summary err">❌ Lỗi: '+(err.message||err)+'</div>'; if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err)); })
          .executeCleanInactiveChannels(t);
      }
    </script>
  `;
}

function _ivOAuth_() {
  return _ivSharedStyles_() + `
    <div class="iv-form">
      <h3 style="margin:0 0 8px;font-size:15px;color:#0f172a;font-weight:800;">🔑 Cài đặt Access Token cho YouTube Analytics</h3>
      <p style="margin:0 0 12px;font-size:13px;color:#475569;line-height:1.6;">Cấp quyền OAuth để hệ thống đọc Analytics kênh của bạn. Token được lưu vĩnh viễn — chỉ cần làm 1 lần.</p>
      <ol style="font-size:13px;color:#334155;line-height:1.8;margin:0 0 14px;padding-left:22px;">
        <li>Đảm bảo đã điền <b>API KEY!B3</b> (Client ID) và <b>B4</b> (Client Secret).</li>
        <li>Bấm <b>Mở link cấp quyền</b> ở dưới — sẽ mở tab mới đến Google.</li>
        <li>Đăng nhập đúng tài khoản quản lý kênh YouTube + cho phép quyền.</li>
        <li>Sau khi cấp quyền, trình duyệt sẽ redirect đến <code>http://localhost/?code=...</code> (hiện lỗi "không thể truy cập" là bình thường).</li>
        <li>Copy <b>toàn bộ URL</b> trên thanh địa chỉ → dán vào ô bên dưới → bấm <b>Xác nhận</b>.</li>
      </ol>
      <div id="oAuthUrlBox" style="display:none;background:linear-gradient(135deg,#eff6ff,#dbeafe);border:1px solid #bfdbfe;border-radius:11px;padding:12px;margin-bottom:12px;">
        <div style="font-size:11px;font-weight:800;color:#1e40af;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">🔗 Link cấp quyền (đã sinh)</div>
        <a id="oAuthUrl" target="_blank" style="display:inline-block;padding:9px 14px;background:linear-gradient(135deg,#1e40af,#3b82f6);color:#fff;text-decoration:none;border-radius:9px;font-weight:700;font-size:12.5px;box-shadow:0 4px 10px -3px rgba(30,64,175,.4);">🚀 Mở link cấp quyền (tab mới)</a>
      </div>
      <div class="iv-row c1" style="grid-template-columns:1fr;">
        <label>URL trả về sau khi cấp quyền (chứa <code>code=</code>)<input id="oCallback" placeholder="http://localhost/?code=4/0Adeu...&scope=..."></label>
      </div>
      <div class="iv-actions">
        <button id="oGenBtn" class="iv-primary" style="background:linear-gradient(135deg,#1e40af,#3b82f6);box-shadow:0 6px 14px -4px rgba(30,64,175,.42);" onclick="oGen()">🔗 1. Sinh link cấp quyền</button>
        <button id="oConfirmBtn" class="iv-primary" style="background:linear-gradient(135deg,#10b981,#059669);box-shadow:0 6px 14px -4px rgba(16,185,129,.42);" onclick="oConfirm()" disabled>✓ 2. Xác nhận token</button>
      </div>
    </div>
    <div id="oResult" class="iv-result" style="display:none;"></div>
    <script>
      function oEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function oGen(){
        document.getElementById('oGenBtn').disabled = true;
        if (window.dlgProgress) window.dlgProgress.start(2, 'Đang đọc Client ID/Secret');
        google.script.run
          .withSuccessHandler(function(res){
            document.getElementById('oGenBtn').disabled = false;
            if (!res.success){ alert('Lỗi: ' + res.message); if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+res.message); return; }
            document.getElementById('oAuthUrl').href = res.url;
            document.getElementById('oAuthUrlBox').style.display = 'block';
            document.getElementById('oConfirmBtn').disabled = false;
            if (window.dlgProgress) window.dlgProgress.complete('Đã tạo link');
          })
          .withFailureHandler(function(err){ document.getElementById('oGenBtn').disabled = false; alert('Lỗi: '+(err.message||err)); if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err)); })
          .getOAuthAuthUrl();
      }
      function oConfirm(){
        var raw = (document.getElementById('oCallback').value||'').trim();
        if (!raw){ alert('Vui lòng dán URL trả về.'); return; }
        var m = raw.match(/[?&]code=([^&]+)/);
        if (!m){ alert('Không tìm thấy code= trong URL. Vui lòng kiểm tra lại.'); return; }
        var code = decodeURIComponent(m[1]);
        document.getElementById('oConfirmBtn').disabled = true;
        if (window.dlgProgress) window.dlgProgress.start(3, 'Đang đổi code lấy refresh token');
        google.script.run
          .withSuccessHandler(function(res){
            document.getElementById('oConfirmBtn').disabled = false;
            var box = document.getElementById('oResult'); box.style.display='block';
            box.innerHTML = '<div class="iv-summary '+(res && res.success !== false ?'ok':'err')+'">'+(res && res.success !== false ?'✅ Đã lưu refresh token. Bạn có thể bắt đầu dùng các báo cáo Analytics.':'❌ '+oEsc((res && res.message)||'Lỗi không rõ'))+'</div>';
            if (window.dlgProgress){ (res && res.success !== false) ? window.dlgProgress.complete('Đã kết nối') : window.dlgProgress.fail('Lỗi'); }
            if (typeof loadStats === 'function') loadStats();
          })
          .withFailureHandler(function(err){ document.getElementById('oConfirmBtn').disabled = false; var box = document.getElementById('oResult'); box.style.display='block'; box.innerHTML = '<div class="iv-summary err">❌ Lỗi: '+oEsc(err.message||err)+'</div>'; if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err)); })
          .processAuthCode(code);
      }
    </script>
  `;
}

function getOAuthAuthUrl_() {
  try {
    const clientId = getClientId();
    const scopes = 'https://www.googleapis.com/auth/yt-analytics.readonly https://www.googleapis.com/auth/youtube.readonly';
    const url = 'https://accounts.google.com/o/oauth2/v2/auth?'
      + 'client_id=' + encodeURIComponent(clientId)
      + '&redirect_uri=' + encodeURIComponent('http://localhost')
      + '&response_type=code'
      + '&access_type=offline'
      + '&prompt=consent'
      + '&scope=' + encodeURIComponent(scopes);
    return { success: true, url: url };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function _ivGuide_() {
  const guideSections = _getGuideSectionsArray_();
  return `
    <style>
      .gFullWrap{display:flex;flex-direction:column;gap:14px;height:100%}
      .gTop{display:grid;grid-template-columns:340px 1fr;gap:14px}
      .gSelect,.gSearch{width:100%;box-sizing:border-box;padding:11px 14px;border:1px solid #d1d5db;border-radius:12px;background:#fff;font-size:13.5px;color:#0f172a;box-shadow:0 1px 2px rgba(15,23,42,.04),inset 0 1px 0 rgba(255,255,255,.8);transition:all .2s ease;outline:none;font-family:inherit}
      .gSelect:hover,.gSearch:hover{border-color:#94a3b8}
      .gSelect:focus,.gSearch:focus{border-color:#6366f1;box-shadow:0 0 0 4px rgba(99,102,241,.14)}
      .gSearch::placeholder{color:#94a3b8}
      .gBody{display:grid;grid-template-columns:300px minmax(0,1fr);gap:14px;flex:1;min-height:0}
      .gNav{background:linear-gradient(180deg,#fff 0%,#f8fafc 100%);border:1px solid #e5e7eb;border-radius:14px;padding:10px;overflow:auto;max-height:calc(100vh - 240px);box-shadow:0 8px 24px -10px rgba(15,23,42,.10),inset 0 1px 0 #fff}
      .gNav::-webkit-scrollbar,.gContent::-webkit-scrollbar{width:8px;height:8px}
      .gNav::-webkit-scrollbar-thumb,.gContent::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#cbd5e1,#94a3b8);border-radius:8px}
      .gNav::-webkit-scrollbar-thumb:hover,.gContent::-webkit-scrollbar-thumb:hover{background:linear-gradient(180deg,#94a3b8,#64748b)}
      .gNav::-webkit-scrollbar-track,.gContent::-webkit-scrollbar-track{background:transparent}
      .gBtn{display:block;width:100%;text-align:left;padding:10px 12px;border:0;background:transparent;border-radius:10px;cursor:pointer;font-size:12.5px;color:#334155;font-weight:500;transition:all .18s ease;line-height:1.35;margin:1px 0;position:relative;font-family:inherit}
      .gBtn:hover{background:#eef2ff;color:#1e293b;transform:translateX(2px)}
      .gBtn.active{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);color:#fff;font-weight:700;box-shadow:0 4px 12px rgba(15,23,42,.32),inset 0 1px 0 rgba(255,255,255,.10)}
      .gBtn.active::before{content:"";position:absolute;left:0;top:8px;bottom:8px;width:3px;background:linear-gradient(180deg,#6366f1,#8b5cf6);border-radius:3px}
      .gGroup{font-size:10.5px;text-transform:uppercase;color:#64748b;font-weight:800;margin:14px 10px 6px;letter-spacing:0.08em;display:flex;align-items:center;gap:8px}
      .gGroup::after{content:"";flex:1;height:1px;background:linear-gradient(90deg,#e2e8f0,transparent)}
      .gContent{background:linear-gradient(180deg,#fff 0%,#fafbfd 100%);border:1px solid #e5e7eb;border-radius:14px;padding:24px 26px;overflow:auto;max-height:calc(100vh - 240px);box-shadow:0 10px 30px -12px rgba(15,23,42,.12),inset 0 1px 0 #fff;color:#1f2937;font-size:14px;line-height:1.65}
      .gContent h2{margin:0 0 14px;font-size:24px;color:#0f172a;font-weight:800;letter-spacing:-0.015em;background:linear-gradient(135deg,#0f172a 0%,#334155 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;padding-bottom:10px;border-bottom:2px solid #f1f5f9}
      .gContent h3{margin:18px 0 10px;font-size:16px;color:#1e293b;font-weight:700}
      .gContent p{margin:8px 0;color:#334155}
      .gContent ol,.gContent ul{padding-left:22px;margin:10px 0}
      .gContent li{margin:7px 0;color:#334155}
      .gContent li::marker{color:#6366f1;font-weight:700}
      .gContent b{color:#0f172a;font-weight:700}
      .gContent code{background:linear-gradient(180deg,#f1f5f9,#e2e8f0);border:1px solid #cbd5e1;border-radius:6px;padding:2px 7px;font-family:'SF Mono','Monaco','Cascadia Code','Roboto Mono',Consolas,monospace;font-size:12.5px;color:#0f172a;font-weight:600}
      .gContent a{color:#4f46e5;text-decoration:none;border-bottom:1px dashed #a5b4fc;transition:all .2s}
      .gContent a:hover{color:#3730a3;border-bottom-color:#4f46e5}
      .guideTable{border-collapse:separate;border-spacing:0;width:100%;font-size:13px;border-radius:10px;overflow:hidden;box-shadow:0 4px 12px rgba(15,23,42,.06);margin:12px 0}
      .guideTable th,.guideTable td{border:1px solid #e5e7eb;padding:10px 12px;text-align:left}
      .guideTable th{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);color:#fff;font-weight:700;letter-spacing:0.02em;border-color:#1e293b}
      .guideTable tr:nth-child(even) td{background:#f8fafc}
      .guideTable tr:hover td{background:#eef2ff}
      .gContent .note{background:linear-gradient(135deg,#ecfdf5 0%,#d1fae5 100%);border-left:4px solid #10b981;padding:12px 14px;border-radius:10px;margin:12px 0;color:#064e3b;box-shadow:0 2px 6px rgba(16,185,129,.10)}
      .gContent .warn{background:linear-gradient(135deg,#fff7ed 0%,#ffedd5 100%);border-left:4px solid #f97316;padding:12px 14px;border-radius:10px;margin:12px 0;color:#7c2d12;box-shadow:0 2px 6px rgba(249,115,22,.10)}
      .gContent .tip{background:linear-gradient(135deg,#eff6ff 0%,#dbeafe 100%);border-left:4px solid #3b82f6;padding:12px 14px;border-radius:10px;margin:12px 0;color:#1e3a8a;box-shadow:0 2px 6px rgba(59,130,246,.10)}
      .gContent .videoBox{background:linear-gradient(135deg,#f8fafc 0%,#eef2ff 100%);border:1px solid #e0e7ff;border-radius:14px;padding:14px;margin:14px 0;box-shadow:0 6px 18px -8px rgba(99,102,241,.20)}
      .gContent .links{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}
      .gContent .links a{display:inline-flex;align-items:center;gap:6px;background:linear-gradient(135deg,#eef2ff 0%,#e0e7ff 100%);color:#3730a3;text-decoration:none;border:1px solid #c7d2fe;border-radius:999px;padding:7px 13px;font-size:12px;font-weight:700;transition:all .2s ease;box-shadow:0 1px 2px rgba(79,70,229,.08)}
      .gContent .links a:hover{background:linear-gradient(135deg,#4f46e5 0%,#6366f1 100%);color:#fff;border-color:#4f46e5;transform:translateY(-1px);box-shadow:0 4px 10px rgba(79,70,229,.32)}
    </style>
    <div class="gFullWrap">
      <div class="gTop">
        <select id="gSelect" class="gSelect" onchange="gSelectGuide(this.value)"></select>
        <input id="gSearch" class="gSearch" oninput="gRenderNav()" placeholder="🔎 Tìm hướng dẫn theo chức năng, API key, OAuth, subtitle, analytics, troubleshooting...">
      </div>
      <div class="gBody">
        <div id="gNav" class="gNav"></div>
        <div id="gContent" class="gContent"></div>
      </div>
    </div>
    <script>
      var GUIDE_SECTIONS = ${JSON.stringify(guideSections)}; var gActive = 0;
      function gEsc(s){ return (s||'').replace(/[&<>"']/g, function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function gRenderSelect(){ document.getElementById('gSelect').innerHTML = GUIDE_SECTIONS.map(function(g,i){ return '<option value="'+i+'">'+gEsc(g.group+' - '+g.title)+'</option>'; }).join(''); }
      /* gFuzzyMatch: prefix-token matching cho ô tìm kiếm hướng dẫn (dialog 2) */
      function gFuzzyMatch(text, rawQ){
        if (!rawQ) return true;
        var t = (text||'').toLowerCase();
        var q = rawQ.toLowerCase().trim();
        if (!q) return true;
        if (t.indexOf(q) >= 0) return true;
        var tokens = q.split(/\s+/).filter(Boolean);
        var words = t.split(/[\s,.\-\/|]+/).filter(Boolean);
        if (tokens.length < 2) return words.some(function(w){ return w.indexOf(tokens[0])===0; });
        return tokens.every(function(tok){ return t.indexOf(tok)>=0 || words.some(function(w){ return w.indexOf(tok)===0; }); });
      }
      function gRenderNav(){ var q = (document.getElementById('gSearch').value||'').trim(); var last=''; var html = GUIDE_SECTIONS.map(function(g,i){return {g:g,i:i};}).filter(function(x){ return !q || gFuzzyMatch(x.g.group+' '+x.g.title+' '+x.g.html.replace(/<[^>]+>/g,' '), q); }).map(function(x){ var h = x.g.group !== last ? '<div class="gGroup">'+gEsc(x.g.group)+'</div>' : ''; last = x.g.group; return h+'<button class="gBtn '+(x.i===gActive?'active':'')+'" onclick="gSelectGuide('+x.i+')">'+gEsc(x.g.title)+'</button>'; }).join(''); document.getElementById('gNav').innerHTML = html || '<div style="padding:10px;color:#64748b;">Không tìm thấy hướng dẫn phù hợp.</div>'; }
      function gSelectGuide(i){ gActive = parseInt(i,10) || 0; document.getElementById('gSelect').value = String(gActive); var gc = document.getElementById('gContent'); gc.innerHTML = GUIDE_SECTIONS[gActive].html; gc.scrollTop = 0; gRenderNav(); if(window.dlgProgress){ window.dlgProgress.set(0,'Đang đọc: '+GUIDE_SECTIONS[gActive].title); } }
      gRenderSelect(); gSelectGuide(0);
      setTimeout(function(){ if(window.dlgProgress) window.dlgProgress.attachScroll('gContent'); }, 60);
      document.getElementById('gSearch').addEventListener('keydown', function(e){ if(e.key==='Escape'){ this.value=''; gRenderNav(); }});
    </script>
  `;
}

function _ivSubtitleGuide_() {
  return _ivSharedStyles_() + `
    <div class="iv-tip" style="margin:0 0 14px;background:linear-gradient(135deg,#fff7ed,#fed7aa);border-color:#fdba74;color:#7c2d12;">
      <b>⚠️ Vì sao cần cấu hình?</b><br>Apps Script chạy chung IP với hàng triệu người → YouTube hay chặn (HTTP 429). Cấu hình API miễn phí để có tỷ lệ thành công ~99%.
    </div>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:12px;margin-bottom:14px;">
      <div class="iv-card" style="border-left:4px solid #10b981;">
        <h3 style="margin:0 0 8px;font-size:14px;font-weight:800;color:#0f172a;">🌟 LỰA CHỌN 1: SUPADATA.AI <span style="font-size:10px;background:#d1fae5;color:#065f46;padding:2px 7px;border-radius:999px;margin-left:4px;">KHUYẾN NGHỊ</span></h3>
        <ul style="margin:0 0 8px;padding-left:18px;font-size:12.5px;color:#475569;line-height:1.7;">
          <li>✅ <b>Miễn phí 100 transcript/tháng</b> — không cần thẻ</li>
          <li>✅ Hỗ trợ AI fallback nếu video không có sub</li>
          <li>✅ Hỗ trợ trên 100+ ngôn ngữ</li>
        </ul>
        <ol style="margin:0;padding-left:18px;font-size:12.5px;color:#334155;line-height:1.7;">
          <li>Truy cập <a href="https://supadata.ai" target="_blank" style="color:#10b981;font-weight:700;">supadata.ai</a> → Đăng ký bằng Google</li>
          <li>Dashboard → Copy API Key</li>
          <li>Dán vào ô <b>B5</b> tại Sheet "<b>${SHEET_API_KEY}</b>"</li>
        </ol>
      </div>
      <div class="iv-card" style="border-left:4px solid #4285F4;">
        <h3 style="margin:0 0 8px;font-size:14px;font-weight:800;color:#0f172a;">🌟 LỰA CHỌN 2: YOUTUBE-TRANSCRIPT.IO</h3>
        <ul style="margin:0 0 8px;padding-left:18px;font-size:12.5px;color:#475569;line-height:1.7;">
          <li>✅ Miễn phí khi đăng ký tài khoản</li>
          <li>✅ Hỗ trợ batch 50 video/lần</li>
        </ul>
        <ol style="margin:0;padding-left:18px;font-size:12.5px;color:#334155;line-height:1.7;">
          <li>Truy cập <a href="https://www.youtube-transcript.io" target="_blank" style="color:#4285F4;font-weight:700;">youtube-transcript.io</a></li>
          <li>Tạo tài khoản → Profile → Copy API Token</li>
          <li>Dán vào ô <b>B6</b> tại Sheet "<b>${SHEET_API_KEY}</b>"</li>
        </ol>
      </div>
    </div>
    <div class="iv-tip" style="background:linear-gradient(135deg,#ecfdf5,#d1fae5);border-color:#6ee7b7;color:#065f46;">
      <b>💡 MẸO:</b> Cấu hình <b>CẢ HAI</b> để tăng độ ổn định. Hệ thống sẽ thử Supadata trước, fail thì fallback YT-Transcript.io, cuối cùng mới scrape (hay 429).
    </div>
    <h3 style="margin:14px 0 8px;font-size:14px;font-weight:800;color:#0f172a;">📋 Bố cục Sheet "${SHEET_API_KEY}"</h3>
    <table class="iv-table" style="margin-bottom:6px;">
      <thead><tr><th style="width:60px;">Ô</th><th>Nội dung</th></tr></thead>
      <tbody>
        <tr><td class="iv-num">B2</td><td><b style="color:#dc2626;">YouTube Data API Key (BẮT BUỘC)</b></td></tr>
        <tr><td class="iv-num">B3</td><td>Google OAuth Client ID</td></tr>
        <tr><td class="iv-num">B4</td><td>Google OAuth Client Secret</td></tr>
        <tr style="background:#fff7ed;"><td class="iv-num">B5</td><td><b>🆕 Supadata API Key</b> (Subtitle ưu tiên)</td></tr>
        <tr style="background:#fff7ed;"><td class="iv-num">B6</td><td><b>🆕 YouTube-Transcript.io Token</b> (Subtitle dự phòng)</td></tr>
        <tr style="background:#eef2ff;"><td class="iv-num">B7</td><td><b>9router API Key / Bearer Token</b> (cho AI)</td></tr>
      </tbody>
    </table>
  `;
}

function _ivFetch_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const channelSheet = ss.getSheetByName(SHEET_CHANNEL);
  if (!channelSheet) return _ivSharedStyles_() + '<div class="iv-card" style="padding:30px;text-align:center;color:#dc2626;font-weight:700;">❌ Không tìm thấy Sheet ' + SHEET_CHANNEL + '</div>';
  const lastChannelRow = channelSheet.getLastRow();
  if (lastChannelRow < 2) return _ivSharedStyles_() + '<div class="iv-card" style="padding:30px;text-align:center;color:#dc2626;font-weight:700;">❌ Sheet ' + SHEET_CHANNEL + ' chưa có link kênh ở cột C</div>';
  setupFetchVideoHistorySheet_();
  const channelPickerItems = channelSheet.getRange(2, 2, lastChannelRow - 1, 2).getDisplayValues()
    .map(function(r, idx){ return { row: idx + 2, name: (r[0] || '').toString().trim() || 'Kênh chưa có tên', link: (r[1] || '').toString().trim() }; })
    .filter(function(x){ return x.link; });
  const initialHistory = getFetchVideoHistoryItems_(30);
  return _ivSharedStyles_() + `
    <style>
      .ivf-picker-btn{width:100%;padding:11px 14px;border:1px solid #1a73e8;background:linear-gradient(180deg,#fff,#eff6ff);text-align:left;cursor:pointer;border-radius:10px;font-family:inherit;font-size:13px;font-weight:700;color:#1e40af;transition:all .2s ease}
      .ivf-picker-btn:hover{background:linear-gradient(180deg,#eff6ff,#dbeafe);box-shadow:0 4px 12px -4px rgba(26,115,232,.3)}
      .ivf-picker-popup{display:none;position:absolute;z-index:999;left:0;right:0;top:74px;background:#fff;border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 16px 40px rgba(15,23,42,.18);padding:12px;animation:dlgFadeIn .2s ease both}
      .ivf-picker-popup.open{display:block}
      .ivf-search{width:100%;box-sizing:border-box;padding:9px 12px;border:1px solid #cbd5e1;border-radius:8px;margin-bottom:8px;font-family:inherit;font-size:13px;outline:none;transition:all .2s ease}
      .ivf-search:focus{border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.16)}
      .ivf-quick-btns{display:flex;gap:8px;margin-bottom:8px;flex-wrap:wrap}
      .ivf-quick-btn{padding:7px 11px;border:1px solid #cbd5e1;background:#f8fafc;border-radius:8px;cursor:pointer;font-family:inherit;font-size:12px;font-weight:600;transition:all .15s ease}
      .ivf-quick-btn:hover{background:#eef2ff;border-color:#a5b4fc}
      .ivf-options{max-height:240px;overflow:auto;border-top:1px solid #eef2f7;padding-top:6px}
      .ivf-options::-webkit-scrollbar{width:8px}
      .ivf-options::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:8px}
      .ivf-opt{display:grid;grid-template-columns:18px minmax(0,1fr) minmax(150px,42%);gap:8px;align-items:center;padding:7px 6px;border-bottom:1px solid #f1f5f9;cursor:pointer;border-radius:6px;transition:background .15s}
      .ivf-opt:hover{background:#f8fafc}
      .ivf-opt b{font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#0f172a}
      .ivf-opt .lk{font-size:11px;color:#64748b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right}
      .ivf-stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:10px}
      .ivf-history-item{display:grid;grid-template-columns:18px 145px 1fr 110px;gap:10px;align-items:start;padding:9px 10px;border-radius:8px;background:#fff;border:1px solid #f1f5f9;margin-bottom:5px;font-size:12px;transition:background .15s}
      .ivf-history-item:hover{background:#f8fafc}
      .ivf-section-lbl{font-size:11px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.04em;margin-bottom:5px}
    </style>
    <div class="iv-form" style="position:relative;">
      <div style="font-size:13px;color:#475569;font-weight:600;margin-bottom:10px;">Sheet <b>${SHEET_CHANNEL}</b> · ${channelPickerItems.length} kênh có link · sẽ thêm video phù hợp vào cuối Sheet <b>${SHEET_VIDEO}</b></div>
      <div style="position:relative;">
        <div class="ivf-section-lbl">📺 Chọn kênh cần quét <span style="font-weight:600;color:#94a3b8;text-transform:none;letter-spacing:0;">(có thể chọn nhiều để tiết kiệm quota)</span></div>
        <button type="button" id="ivfPickerBtn" class="ivf-picker-btn" onclick="ivfTogglePicker(event)">Chưa chọn kênh nào</button>
        <div id="ivfPicker" class="ivf-picker-popup" onclick="event.stopPropagation()">
          <input id="ivfSearch" class="ivf-search" oninput="ivfRenderOpts()" placeholder="🔎 Tìm theo tên kênh hoặc link...">
          <div class="ivf-quick-btns">
            <button type="button" class="ivf-quick-btn" onclick="ivfSelectAllVisible(true)">✓ Chọn tất cả (đang lọc)</button>
            <button type="button" class="ivf-quick-btn" onclick="ivfSelectAllVisible(false)">○ Bỏ chọn (đang lọc)</button>
            <button type="button" class="ivf-quick-btn" onclick="ivfTogglePicker()">✕ Đóng</button>
          </div>
          <div id="ivfOptions" class="ivf-options"></div>
        </div>
      </div>
      <div class="iv-row c2" style="margin-top:14px;">
        <label>📌 Chủ đề lọc thêm <span style="font-weight:600;color:#94a3b8;text-transform:none;letter-spacing:0;">(có thể để trống)</span><input id="ivfTopic" type="text" placeholder="VD: psychology, mental health, self improvement"></label>
        <label>🔎 Cách khớp chủ đề<select id="ivfMatch"><option value="any" selected>Khớp ít nhất 1 từ/cụm từ</option><option value="all">Bắt buộc khớp tất cả từ/cụm từ</option></select></label>
      </div>
      <div class="iv-row c4" style="margin-top:6px;">
        <label>👁️ Views tối thiểu<input id="ivfMinViews" type="number" value="50000" min="0"></label>
        <label>📅 Mới trong vòng (ngày)<input id="ivfDays" type="number" value="90" min="1" max="90"></label>
        <label>⏱️ Thời lượng tối thiểu (HH:MM:SS)<input id="ivfMinDur" type="text" value="00:03:10"></label>
        <label class="check"><input type="checkbox" id="ivfAutoUpdate" checked> Auto update metadata sau khi thêm</label>
      </div>
      <div class="iv-actions" style="margin-top:14px;">
        <button id="ivfBtn" class="iv-primary" style="background:linear-gradient(135deg,#1a73e8,#1e40af);box-shadow:0 6px 14px -4px rgba(26,115,232,.42);" onclick="ivfRun()">🚀 BẮT ĐẦU QUÉT</button>
        <button class="iv-primary" style="background:#fff;color:#475569;border:1px solid #e5e7eb;box-shadow:none;" onclick="goHome()">← Hủy & quay lại</button>
        <span class="iv-hint">⚠️ <code style="background:#fff7ed;padding:1px 6px;border-radius:4px;border:1px solid #fdba74;">playlistItems.list</code> ~1u/kênh + <code style="background:#fff7ed;padding:1px 6px;border-radius:4px;border:1px solid #fdba74;">videos.list</code> ~1u/50 video</span>
      </div>
    </div>
    <div id="ivfLoader" class="iv-loader" style="display:none;">
      <div class="iv-spin" style="border-top-color:#1a73e8;border-right-color:#3b82f6;"></div>
      <div class="iv-spin-text">Đang quét video từ các kênh...</div>
      <div class="iv-spin-sub" id="ivfSub">Khởi tạo · gọi YouTube Data API · lọc theo bộ lọc</div>
      <div class="iv-spin-meter"><div style="background:linear-gradient(90deg,transparent,#1a73e8,#3b82f6,transparent);"></div></div>
    </div>
    <div id="ivfResult" class="iv-result" style="display:none;"></div>
    <div class="iv-card" style="margin-top:14px;">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap;">
        <h3 style="margin:0;font-size:14px;font-weight:800;color:#0f172a;">📜 Lịch sử lấy video trong kênh</h3>
        <div style="display:flex;gap:6px;flex-wrap:wrap;">
          <button class="ivf-quick-btn" onclick="ivfRefreshHistory()">🔄 Tải lại</button>
          <button class="ivf-quick-btn" style="border-color:#fecaca;background:#fff7f7;color:#b91c1c;" onclick="ivfDelSelectedHistory()">🗑️ Xóa đã chọn</button>
          <button class="ivf-quick-btn" style="border-color:#b91c1c;background:#b91c1c;color:#fff;" onclick="ivfClearHistory()">💀 Xóa toàn bộ</button>
        </div>
      </div>
      <div id="ivfHistoryBox" style="max-height:240px;overflow:auto;"></div>
    </div>
    <script>
      var IVF_CHANNELS = ${JSON.stringify(channelPickerItems)};
      var IVF_HISTORY = ${JSON.stringify(initialHistory)};
      var IVF_SELECTED = new Set();
      var IVF_LAST_ROW = ${lastChannelRow};
      function ivfFmt(n){ return (Number(n)||0).toLocaleString('vi-VN'); }
      function ivfEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
      function ivfTogglePicker(e){ if (e) e.stopPropagation(); document.getElementById('ivfPicker').classList.toggle('open'); ivfRenderOpts(); }
      /* ivfFuzzyMatch: prefix-token matching cho ô tìm kênh trong Fetch-Video picker */
      function ivfFuzzyMatch(text, rawQ){
        if (!rawQ) return true;
        var t = (text||'').toLowerCase();
        var q = rawQ.toLowerCase().trim();
        if (!q) return true;
        if (t.indexOf(q) >= 0) return true;
        var tokens = q.split(/\s+/).filter(Boolean);
        var words = t.split(/[\s,.\-\/|]+/).filter(Boolean);
        if (tokens.length < 2) return words.some(function(w){ return w.indexOf(tokens[0])===0; });
        return tokens.every(function(tok){ return t.indexOf(tok)>=0 || words.some(function(w){ return w.indexOf(tok)===0; }); });
      }
      function ivfRenderOpts(){
        var q = (document.getElementById('ivfSearch').value||'').toLowerCase().trim();
        var filtered = IVF_CHANNELS.filter(function(c){
          if (!q) return true;
          var haystack = c.name + ' ' + c.link + ' ' + String(c.row);
          return ivfFuzzyMatch(haystack, q);
        });
        document.getElementById('ivfOptions').innerHTML = filtered.length ? filtered.map(function(c){
          return '<label class="ivf-opt"><input type="checkbox" class="ivfPick" data-row="'+c.row+'" '+(IVF_SELECTED.has(c.row)?'checked':'')+' onchange="ivfToggleRow('+c.row+',this.checked)"><b>'+ivfEsc(c.name)+'</b><span class="lk">'+ivfEsc(c.link)+'</span></label>';
        }).join('') : '<div style="padding:14px;color:#64748b;font-weight:600;">Không có kênh phù hợp với "'+ivfEsc(q)+'"</div>';
        ivfUpdateBtn();
      }
      function ivfToggleRow(row, checked){ if (checked) IVF_SELECTED.add(row); else IVF_SELECTED.delete(row); ivfUpdateBtn(); }
      function ivfSelectAllVisible(checked){ Array.from(document.querySelectorAll('.ivfPick')).forEach(function(cb){ cb.checked = checked; ivfToggleRow(parseInt(cb.dataset.row,10), checked); }); ivfRenderOpts(); }
      function ivfUpdateBtn(){ var total = IVF_CHANNELS.length; var picked = IVF_SELECTED.size; document.getElementById('ivfPickerBtn').textContent = picked === 0 ? 'Chưa chọn kênh nào' : (picked === total ? '✓ Đang chọn toàn bộ '+total+' kênh có link' : '✓ Đã chọn '+picked+'/'+total+' kênh'); }
      function ivfRenderHistory(items){
        IVF_HISTORY = items || IVF_HISTORY || [];
        document.getElementById('ivfHistoryBox').innerHTML = IVF_HISTORY.length ? IVF_HISTORY.map(function(h){
          return '<label class="ivf-history-item"><input type="checkbox" class="ivfHistPick" value="'+h.id+'"><span style="color:#64748b;font-weight:600;">'+ivfEsc(h.time||'')+'</span><span><b style="color:#0f172a;">'+ivfEsc(h.channels||'')+'</b><br><span style="color:#64748b;">'+ivfEsc(h.message||'')+'</span></span><span style="text-align:right;color:#1a73e8;font-weight:700;">'+ivfEsc(h.stats||'')+'</span></label>';
        }).join('') : '<div style="color:#64748b;padding:14px;font-weight:600;text-align:center;">Chưa có lịch sử.</div>';
      }
      function ivfRefreshHistory(){ google.script.run.withSuccessHandler(function(res){ if (res.success) ivfRenderHistory(res.history || []); }).getFetchVideoHistory(); }
      function ivfDelSelectedHistory(){
        var ids = Array.from(document.querySelectorAll('.ivfHistPick:checked')).map(function(x){ return parseInt(x.value,10); });
        if (!ids.length){ alert('Vui lòng chọn lịch sử cần xóa.'); return; }
        google.script.run.withSuccessHandler(function(res){ alert(res.message); ivfRenderHistory(res.history || []); }).deleteFetchVideoHistory(ids);
      }
      function ivfClearHistory(){
        if (!confirm('Xóa toàn bộ lịch sử lấy video trong kênh?')) return;
        google.script.run.withSuccessHandler(function(res){ alert(res.message); ivfRenderHistory(res.history || []); }).clearFetchVideoHistory();
      }
      function ivfRenderResult(res){
        var s = res.stats || {};
        return '<div class="ivf-stats-grid">'
          + '<div class="iv-info-card"><div class="lbl">Kênh đã kiểm tra</div><div class="val">'+ivfFmt(s.channelsProcessed||0)+'</div><div class="sub">trong tổng '+ivfFmt(IVF_SELECTED.size||0)+' đã chọn</div></div>'
          + '<div class="iv-info-card"><div class="lbl">Video đã xét</div><div class="val">'+ivfFmt(s.videosChecked||0)+'</div><div class="sub">qua YouTube API</div></div>'
          + '<div class="iv-info-card ok"><div class="lbl">Đạt bộ lọc</div><div class="val">'+ivfFmt(s.passed||0)+'</div><div class="sub">video phù hợp</div></div>'
          + '<div class="iv-info-card ok"><div class="lbl">Đã thêm</div><div class="val">'+ivfFmt(s.added||0)+'</div><div class="sub">vào Sheet '+ivfEsc('${SHEET_VIDEO}')+'</div></div>'
          + '</div>'
          + '<div class="iv-summary '+(res.success?'ok':'err')+'">'+(res.success?'✅ ':'❌ ')+ivfEsc(res.message||'')+'</div>'
          + (res.success ? '<h3>📥 Video gần nhất trong Sheet</h3><div id="ivfRecent">⏳ Đang đọc Sheet...</div>' : '');
      }
      function ivfRenderRecent(d){
        document.getElementById('ivfRecent').innerHTML = ivRenderSheetTable(d, {emptyText:'Sheet chưa có link video ở cột C để hiển thị.'});
      }
      function ivfRun(){
        var pickedRows = Array.from(IVF_SELECTED).sort(function(a,b){ return a-b; });
        if (pickedRows.length === 0){ alert('Vui lòng chọn ít nhất 1 kênh để quét.'); ivfTogglePicker(); return; }
        var params = {
          startRow: 2,
          endRow: IVF_LAST_ROW,
          selectedRows: pickedRows,
          topicKeywords: document.getElementById('ivfTopic').value.trim(),
          minViews: Number.isNaN(parseInt(document.getElementById('ivfMinViews').value,10)) ? 50000 : parseInt(document.getElementById('ivfMinViews').value,10),
          daysBack: Number.isNaN(parseInt(document.getElementById('ivfDays').value,10)) ? 90 : parseInt(document.getElementById('ivfDays').value,10),
          minDuration: document.getElementById('ivfMinDur').value,
          matchMode: document.getElementById('ivfMatch').value,
          autoUpdate: document.getElementById('ivfAutoUpdate').checked,
          progressId: ivMakeProgressId('fetch_channel_videos')
        };
        var btn = document.getElementById('ivfBtn'); btn.disabled = true;
        document.getElementById('ivfResult').style.display='none';
        document.getElementById('ivfLoader').style.display='flex';
        var progressPoll = ivStartProgressPoll(params.progressId, pickedRows.length, 'Đang quét 0/'+pickedRows.length+' kênh');
        var t0 = Date.now();
        var timer = setInterval(function(){ var s = Math.floor((Date.now()-t0)/1000); document.getElementById('ivfSub').textContent = 'Đã chạy '+s+'s · '+pickedRows.length+' kênh · min '+ivfFmt(params.minViews)+' views · '+params.daysBack+' ngày'; }, 500);
        google.script.run
          .withSuccessHandler(function(res){
            clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('ivfLoader').style.display='none'; btn.disabled = false;
            document.getElementById('ivfResult').style.display='block';
            document.getElementById('ivfResult').innerHTML = ivfRenderResult(res);
            if (window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất') : window.dlgProgress.fail('Lỗi: '+(res.message||'')); }
            ivfRefreshHistory();
            if (res.success) google.script.run.withSuccessHandler(ivfRenderRecent).withFailureHandler(function(err){ document.getElementById('ivfRecent').innerHTML = '<div style="padding:10px;color:#dc2626;">Lỗi đọc: '+ivfEsc(err.message||err)+'</div>'; }).getRecentRowsForDisplay('video', 12);
            if (typeof loadStats === 'function') loadStats();
          })
          .withFailureHandler(function(err){
            clearInterval(timer); ivStopProgressPoll(progressPoll); document.getElementById('ivfLoader').style.display='none'; btn.disabled = false;
            document.getElementById('ivfResult').style.display='block';
            document.getElementById('ivfResult').innerHTML = '<div class="iv-summary err">❌ Lỗi: '+ivfEsc(err.message||err)+'</div>';
            if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
          })
          .executeFetchRecentVideosFromChannels(params);
      }
      ivfRenderOpts();
      ivfRenderHistory();
      // Click ngoài picker → đóng
      document.addEventListener('click', function(e){
        var picker = document.getElementById('ivfPicker');
        var btn = document.getElementById('ivfPickerBtn');
        if (picker && picker.classList.contains('open') && !picker.contains(e.target) && e.target !== btn && !btn.contains(e.target)) {
          picker.classList.remove('open');
        }
      });
    </script>
  `;
}

function getAdminDashboardStats_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usage = getYouTubeQuotaUsage_();
  const pct = usage.limit ? (usage.used / usage.limit) * 100 : 0;
  function rows(name){ const sh = ss.getSheetByName(name); return sh ? Math.max(0, sh.getLastRow() - 1) : 0; }
  let healthScore = null, healthSummary = '';
  try {
    const h = getSystemHealthData_();
    healthScore = h.score;
    healthSummary = h.okCount + ' OK / ' + h.warnCount + ' cảnh báo / ' + h.dangerCount + ' nghiêm trọng';
  } catch(e) { healthSummary = 'Lỗi: ' + e.message; }
  const props = PropertiesService.getDocumentProperties();
  const oauthOk = !!props.getProperty('YT_REFRESH_TOKEN');
  return {
    quotaUsed: usage.used,
    quotaLimit: usage.limit,
    quotaRemaining: Math.max(0, usage.limit - usage.used),
    quotaRequests: usage.requests,
    quotaPct: pct,
    healthScore: healthScore,
    healthSummary: healthSummary,
    videoRows: rows(SHEET_VIDEO),
    channelRows: rows(SHEET_CHANNEL),
    analyticsRows: rows(SHEET_ANALYTICS),
    videoAnalyticsRows: rows(SHEET_VIDEO_ANALYTICS),
    aiMemoryRows: rows(AI_MEMORY_SHEET),
    oauthOk: oauthOk,
    timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh', 'HH:mm:ss dd/MM/yyyy')
  };
}

function adminRun(actionName) {
  const ALLOWED = {
    showSystemBrandInfo: showSystemBrandInfo,
    showSystemHealthCheck: showSystemHealthCheck,
    showYouTubeQuotaUsage: showYouTubeQuotaUsage,
    updateVideoSheet: updateVideoSheet,
    updateRangeVideoSheet: updateRangeVideoSheet,
    updateSpecificVideoRow: updateSpecificVideoRow,
    updateFastVideoSheet: updateFastVideoSheet,
    updateChannelSheet: updateChannelSheet,
    updateRangeChannelSheet: updateRangeChannelSheet,
    updateSpecificChannelRow: updateSpecificChannelRow,
    updateFastChannelSheet: updateFastChannelSheet,
    fetchRecentVideosFromChannels: fetchRecentVideosFromChannels,
    openSearchByTopicDialog: openSearchByTopicDialog,
    openSearchChannelsByTopicDialog: openSearchChannelsByTopicDialog,
    runFullAutoCleanup: runFullAutoCleanup,
    checkAndCleanData: checkAndCleanData,
    deleteLowViewVideos: deleteLowViewVideos,
    cleanInactiveChannels: cleanInactiveChannels,
    showSubtitleApiGuide: showSubtitleApiGuide,
    fetchAllSubtitles: fetchAllSubtitles,
    fetchRangeSubtitles: fetchRangeSubtitles,
    fetchSpecificSubtitle: fetchSpecificSubtitle,
    retryFailedSubtitles: retryFailedSubtitles,
    clearSubtitleColumn: clearSubtitleColumn,
    setupOfflineToken: setupOfflineToken,
    clearOfflineToken: clearOfflineToken,
    updateAnalytics7Days: updateAnalytics7Days,
    updateAnalytics28Days: updateAnalytics28Days,
    updateAnalyticsLifetime: updateAnalyticsLifetime,
    updateAnalyticsAll: updateAnalyticsAll,
    updateVideoAnalytics7Days: updateVideoAnalytics7Days,
    updateVideoAnalytics28Days: updateVideoAnalytics28Days,
    updateVideoAnalyticsLifetime: updateVideoAnalyticsLifetime,
    openAISheetAnalyzer: openAISheetAnalyzer
  };
  const fn = ALLOWED[actionName];
  if (!fn) throw new Error('Action không hợp lệ: ' + actionName);
  fn();
  return { ok: true, action: actionName };
}

function openAdminPanel() {
  const CATEGORIES = [
    { id: 'home', icon: '🏠', label: 'Tổng quan', color: '#4f46e5' },
    { id: 'video', icon: '📹', label: 'Video', color: '#1a73e8' },
    { id: 'channel', icon: '📺', label: 'Kênh', color: '#0f9d58' },
    { id: 'search', icon: '🔍', label: 'Tìm kiếm', color: '#7c3aed' },
    { id: 'cleanup', icon: '🧹', label: 'Dọn dẹp', color: '#dc2626' },
    { id: 'subtitle', icon: '📝', label: 'Subtitle', color: '#0891b2' },
    { id: 'analytics', icon: '📊', label: 'Analytics', color: '#4285F4' },
    { id: 'ai', icon: '🤖', label: 'AI', color: '#7c3aed' }
  ];
  const ACTIONS = [
    { cat:'home', icon:'📘', title:'Trung tâm hướng dẫn', desc:'Tóm tắt 9 mục quan trọng: Tổng quan, API key, Quota, Video, Kênh, Tìm kiếm, Subtitle, OAuth, AI, Health.', fn:'showSystemBrandInfo', color:'#4f46e5', inline:'guide' },
    { cat:'home', icon:'🔑', title:'Cài đặt API Keys', desc:'Quản lý tập trung toàn bộ API key: YouTube Data, OAuth, Supadata, YT-Transcript.io, 9router. Điền và lưu trực tiếp vào sheet API KEY.', fn:'showSystemBrandInfo', color:'#0f9d58', inline:'apiKey' },
    { cat:'home', icon:'🩺', title:'Kiểm tra sức khỏe hệ thống', desc:'Chẩn đoán toàn bộ API key, OAuth, sheet, quota — Health Score 0–100.', fn:'showSystemHealthCheck', color:'#10b981', inline:'health' },
    { cat:'home', icon:'📊', title:'Quota Dashboard', desc:'Vòng tròn % quota, 5 stat-chip, ETA reset PT, auto-refresh 10s.', fn:'showYouTubeQuotaUsage', color:'#7c3aed', inline:'quota' },
    { cat:'video', icon:'🔄', title:'Cập nhật TẤT CẢ video', desc:'Cập nhật metadata toàn bộ video. Tốn nhiều quota — chỉ chạy khi cần thiết.', fn:'updateVideoSheet', color:'#1a73e8', danger:true, inline:'direct:updateVideoSheet' },
    { cat:'video', icon:'🎯', title:'Cập nhật theo KHOẢNG dòng', desc:'Khuyến nghị: chia khoảng nhỏ để tránh timeout & tiết kiệm quota.', fn:'updateRangeVideoSheet', color:'#1a73e8', inline:'range:videoRange' },
    { cat:'video', icon:'📍', title:'Cập nhật 1 DÒNG video', desc:'Cập nhật chính xác 1 dòng — dùng khi vừa thêm/sửa link.', fn:'updateSpecificVideoRow', color:'#1a73e8', inline:'single:videoSingle' },
    { cat:'video', icon:'⚡', title:'Cập nhật NHANH Views & VPH', desc:'Chỉ làm mới views + VPH, tốn ít quota, phù hợp dùng hằng ngày.', fn:'updateFastVideoSheet', color:'#0ea5e9', inline:'direct:updateFastVideoSheet' },
    { cat:'channel', icon:'🔄', title:'Cập nhật TẤT CẢ kênh', desc:'Cập nhật metadata toàn bộ kênh. Chạy theo khoảng khi danh sách lớn.', fn:'updateChannelSheet', color:'#0f9d58', danger:true, inline:'direct:updateChannelSheet' },
    { cat:'channel', icon:'🎯', title:'Cập nhật theo KHOẢNG kênh', desc:'Cập nhật kênh trong khoảng dòng cụ thể — kiểm soát quota & timeout.', fn:'updateRangeChannelSheet', color:'#0f9d58', inline:'range:channelRange' },
    { cat:'channel', icon:'📍', title:'Cập nhật 1 DÒNG kênh', desc:'Cập nhật chính xác 1 kênh.', fn:'updateSpecificChannelRow', color:'#0f9d58', inline:'single:channelSingle' },
    { cat:'channel', icon:'⚡', title:'Cập nhật NHANH Subs & Views/tháng', desc:'Chỉ làm mới subscribers + views/tháng, ít quota.', fn:'updateFastChannelSheet', color:'#22c55e', inline:'direct:updateFastChannelSheet' },
    { cat:'channel', icon:'📥', title:'Lấy video trong kênh (<3 tháng)', desc:'Quét link kênh ở cột C, lọc video theo thời lượng/views/chủ đề + lịch sử quét.', fn:'fetchRecentVideosFromChannels', color:'#16a34a', inline:'fetch' },
    { cat:'search', icon:'🎬', title:'Tìm VIDEO theo chủ đề', desc:'Dùng search.list (~100u/lần) — chạy có giới hạn để tiết kiệm quota.', fn:'openSearchByTopicDialog', color:'#7c3aed', danger:true, inline:'searchVideo' },
    { cat:'search', icon:'📺', title:'Tìm KÊNH theo chủ đề', desc:'Tìm kênh theo chủ đề, lọc subs/quốc gia/views.', fn:'openSearchChannelsByTopicDialog', color:'#7c3aed', danger:true, inline:'searchChannel' },
    { cat:'cleanup', icon:'🚀', title:'DỌN DẸP TỔNG HỢP', desc:'Chạy chuỗi: xóa video cũ/trùng + xóa kênh chết. Backup trước nếu cần.', fn:'runFullAutoCleanup', color:'#dc2626', danger:true, inline:'direct:runFullAutoCleanup' },
    { cat:'cleanup', icon:'🗑️', title:'Xóa video > 3 tháng & trùng', desc:'Xóa video cũ/trùng lặp. Yêu cầu cập nhật ngày đăng trước.', fn:'checkAndCleanData', color:'#dc2626', danger:true, inline:'direct:checkAndCleanData' },
    { cat:'cleanup', icon:'✂️', title:'Xóa video < 20.000 views', desc:'Giữ lại video có hiệu suất tối thiểu. Cập nhật views trước khi chạy.', fn:'deleteLowViewVideos', color:'#dc2626', danger:true, inline:'direct:deleteLowViewVideos' },
    { cat:'cleanup', icon:'💀', title:'Xóa kênh chết (theo ngưỡng)', desc:'Form chọn ngưỡng views/tháng tối thiểu để giữ kênh.', fn:'cleanInactiveChannels', color:'#dc2626', danger:true, inline:'cleanup' },
    { cat:'subtitle', icon:'ℹ️', title:'Hướng dẫn cấu hình API Subtitle', desc:'BẮT BUỘC ĐỌC trước khi dùng. Cấu hình B5/B6 để có 99% success.', fn:'showSubtitleApiGuide', color:'#0891b2', inline:'subtitleGuide' },
    { cat:'subtitle', icon:'📥', title:'Lấy Subtitle TẤT CẢ video', desc:'Lấy transcript cho toàn bộ video. Nên chạy theo khoảng nếu danh sách lớn.', fn:'fetchAllSubtitles', color:'#0891b2', danger:true, inline:'direct:fetchAllSubtitles' },
    { cat:'subtitle', icon:'🎯', title:'Lấy Subtitle theo KHOẢNG', desc:'Lấy transcript trong khoảng dòng cụ thể.', fn:'fetchRangeSubtitles', color:'#0891b2', inline:'range:subtitleRange' },
    { cat:'subtitle', icon:'📍', title:'Lấy Subtitle 1 DÒNG', desc:'Lấy transcript cho đúng 1 video.', fn:'fetchSpecificSubtitle', color:'#0891b2', inline:'single:subtitleSingle' },
    { cat:'subtitle', icon:'🔄', title:'Thử lại các dòng FAIL', desc:'Quét cột subtitle để chạy lại các dòng lỗi.', fn:'retryFailedSubtitles', color:'#0891b2', inline:'direct:retryFailedSubtitles' },
    { cat:'subtitle', icon:'🗑️', title:'Xóa Subtitle (Cột M)', desc:'Làm sạch cột transcript — backup trước nếu cần.', fn:'clearSubtitleColumn', color:'#dc2626', danger:true, inline:'direct:clearSubtitleColumn' },
    { cat:'analytics', icon:'🔑', title:'Cài đặt Access Token', desc:'Cấp quyền OAuth Analytics — chỉ cần làm 1 lần. Form 2 bước.', fn:'setupOfflineToken', color:'#4285F4', inline:'oauth' },
    { cat:'analytics', icon:'🗑️', title:'Xóa Token Analytics', desc:'Xóa refresh token để đổi tài khoản/kênh.', fn:'clearOfflineToken', color:'#dc2626', danger:true, inline:'direct:clearOfflineToken' },
    { cat:'analytics', icon:'📅', title:'Tổng quan KÊNH 7 ngày', desc:'Báo cáo Analytics cho 7 ngày qua.', fn:'updateAnalytics7Days', color:'#4285F4', inline:'direct:updateAnalytics7Days' },
    { cat:'analytics', icon:'📅', title:'Tổng quan KÊNH 28 ngày', desc:'Báo cáo Analytics cho 28 ngày qua.', fn:'updateAnalytics28Days', color:'#4285F4', inline:'direct:updateAnalytics28Days' },
    { cat:'analytics', icon:'📅', title:'Tổng quan KÊNH toàn thời gian', desc:'Báo cáo Analytics toàn thời gian.', fn:'updateAnalyticsLifetime', color:'#4285F4', inline:'direct:updateAnalyticsLifetime' },
    { cat:'analytics', icon:'🌐', title:'Tổng quan KÊNH TẤT CẢ', desc:'Cập nhật cả 7d + 28d + lifetime liên tiếp.', fn:'updateAnalyticsAll', color:'#4285F4', inline:'direct:updateAnalyticsAll' },
    { cat:'analytics', icon:'🎬', title:'Từng VIDEO 7 ngày', desc:'Analytics chi tiết theo từng video, 7 ngày.', fn:'updateVideoAnalytics7Days', color:'#3b82f6', inline:'direct:updateVideoAnalytics7Days' },
    { cat:'analytics', icon:'🎬', title:'Từng VIDEO 28 ngày', desc:'Analytics chi tiết theo từng video, 28 ngày.', fn:'updateVideoAnalytics28Days', color:'#3b82f6', inline:'direct:updateVideoAnalytics28Days' },
    { cat:'analytics', icon:'🎬', title:'Từng VIDEO toàn thời gian', desc:'Analytics chi tiết theo từng video, toàn thời gian.', fn:'updateVideoAnalyticsLifetime', color:'#3b82f6', inline:'direct:updateVideoAnalyticsLifetime' },
    { cat:'ai', icon:'🤖', title:'AI phân tích Sheet', desc:'9router OpenAI-compatible — phân tích theo cột/sheet, có lịch sử.', fn:'openAISheetAnalyzer', color:'#7c3aed', inline:'ai' }
  ];
  const stats = getAdminDashboardStats_();
  const html = HtmlService.createHtmlOutput(`
    <div id="adminRoot" style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;padding:14px;background:linear-gradient(135deg,#eef2ff 0%,#f8fafc 50%,#fdf4ff 100%);min-height:100vh;color:#0f172a;">
      ${buildDialogHeader_("🚀 YouTube Tools - Dương Huỳnh Team", "Quản trị toàn bộ chức năng từ một cửa sổ duy nhất. Click vào card để chạy chức năng tương ứng.", "#4f46e5")}
      <div class="adminWindowTools" title="Nút này nằm sát góc phải trong phần nội dung dialog; Apps Script không cho phép chèn vào thanh tiêu đề gốc cạnh dấu X.">
        <button class="adminDockBtn" onclick="minimizeAdminPanel()" title="Hạ bảng điều khiển xuống dock bên phải để dùng Sheet">▾ Hạ xuống</button>
      </div>
      <button id="adminMinDock" class="adminMinDock" onclick="restoreAdminPanel()" title="Bấm để mở lại bảng điều khiển">
        <div class="dockTop"><img src="${APP_LOGO_URL}" alt="logo"><div><b>YouTube Tools</b><span id="dockStatus">Sẵn sàng</span></div><strong id="dockPct">0%</strong></div>
        <div class="dockTrack"><span id="dockBar"></span></div>
        <div class="dockHint">Bấm để mở lại · Ctrl+M</div>
      </button>
      <div class="dlgFadeBody">
        <div class="adminTopBar">
          <div class="adminSearchBox">
            <span class="adminSearchIcon">🔎</span>
            <input id="adminSearch" placeholder="Tìm chức năng theo tên, mô tả, từ khóa..." class="adminSearchInput">
            <kbd class="adminKbd">Ctrl + K</kbd>
          </div>
          <div class="adminTopStats">
            <div class="adminMiniStat" id="msQuota"><div class="lbl">Quota hôm nay</div><div class="val">--</div><div class="bar"><span id="msQuotaBar"></span></div></div>
            <div class="adminMiniStat" id="msHealth"><div class="lbl">Health Score</div><div class="val">--</div><div class="bar"><span id="msHealthBar"></span></div></div>
            <div class="adminMiniStat"><div class="lbl">Sheets</div><div class="val" id="msSheets">--</div><div class="sub" id="msSheetsSub">Đang đọc...</div></div>
            <div class="adminMiniStat"><div class="lbl">OAuth Analytics</div><div class="val" id="msOauth">--</div><div class="sub" id="msOauthSub"></div></div>
          </div>
        </div>
        <div class="adminBody">
          <aside class="adminSide">
            <div class="adminSideTitle">DANH MỤC</div>
            <div id="adminNav"></div>
            <div class="adminSideFoot">
              <button class="adminQuickBtn" onclick="runAction('showSystemHealthCheck','Health Check')">🩺 Quick Health</button>
              <button class="adminQuickBtn" onclick="runAction('showYouTubeQuotaUsage','Quota')">📊 Quota</button>
              <button class="adminQuickBtn" onclick="loadStats()">🔄 Refresh stats</button>
            </div>
          </aside>
          <main class="adminMain" id="adminMain">
            <div id="adminHomeView">
              <div id="adminCatHeader" class="adminCatHeader"></div>
              <div id="adminGrid" class="adminGrid"></div>
              <div id="adminEmpty" class="adminEmpty" style="display:none;">Không tìm thấy chức năng nào khớp với từ khóa.</div>
            </div>
            <div id="adminInlineView" style="display:none;">
              <div class="iv-bar">
                <button class="iv-back" onclick="goHome()" title="Quay lại Bảng điều khiển (Esc)">← Quay lại</button>
                <div class="iv-bread">
                  <span class="iv-bread-root">🎛️ Bảng điều khiển</span>
                  <span class="iv-bread-sep">▸</span>
                  <span class="iv-bread-cur" id="ivBreadCur">...</span>
                </div>
                <button class="iv-refresh" onclick="reloadInlineView()" title="Tải lại view">🔄</button>
              </div>
              <div id="adminInlineContent"></div>
            </div>
            <div id="adminLoaderOverlay" class="adminLoaderOverlay" style="display:none;">
              <div class="adminLoaderCard">
                <div class="adminLoaderRing">
                  <svg viewBox="0 0 50 50"><circle cx="25" cy="25" r="20" fill="none" stroke="url(#alGrad)" stroke-width="5" stroke-linecap="round" stroke-dasharray="90 200" /><defs><linearGradient id="alGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#7c3aed"/><stop offset="100%" stop-color="#ec4899"/></linearGradient></defs></svg>
                </div>
                <div class="adminLoaderText" id="adminLoaderText">Đang khởi chạy...</div>
                <div class="adminLoaderSub" id="adminLoaderSub">Vui lòng chờ trong giây lát</div>
                <div class="adminLoaderMeter"><div></div></div>
              </div>
            </div>
          </main>
        </div>
        <footer class="adminFoot">
          <span id="adminTs">Cập nhật lúc: ${stats.timestamp}</span>
          <span class="adminFootRight">${ACTIONS.length} chức năng · ${CATEGORIES.length} nhóm · YouTube Tools - Dương Huỳnh Team</span>
        </footer>
      </div>
      <style>
        .adminWindowTools{position:fixed;top:8px;right:48px;z-index:9999;display:flex;justify-content:flex-end;pointer-events:none}
        .adminDockBtn{border:1px solid #c7d2fe;background:linear-gradient(135deg,#ffffff,#eef2ff);color:#3730a3;border-radius:999px;padding:8px 13px;font-size:12px;font-weight:800;cursor:pointer;font-family:inherit;box-shadow:0 6px 16px -10px rgba(79,70,229,.35)}
        .adminWindowTools .adminDockBtn{pointer-events:auto}
        .adminDockBtn:hover{background:linear-gradient(135deg,#4f46e5,#6366f1);color:#fff;border-color:#4f46e5}
        .adminMinDock{display:none;width:100%;border:1px solid #c7d2fe;background:linear-gradient(135deg,#ffffff,#eef2ff 65%,#fdf4ff);border-radius:15px;padding:9px 11px;text-align:left;cursor:pointer;font-family:inherit;box-shadow:0 14px 26px -16px rgba(15,23,42,.34),inset 0 1px 0 #fff;color:#0f172a;box-sizing:border-box;margin:0;overflow:hidden}
        .adminMinDock .dockTop{display:grid;grid-template-columns:auto 1fr auto;gap:8px;align-items:center}
        .adminMinDock img{width:30px;height:30px;border-radius:8px;object-fit:contain;background:#fff;border:1px solid #e5e7eb;box-shadow:0 4px 10px -6px rgba(15,23,42,.3)}
        .adminMinDock b{display:block;font-size:12.5px;line-height:1.1;color:#111827}
        .adminMinDock span{display:block;font-size:10.5px;color:#64748b;margin-top:1px;font-weight:650;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px}
        .adminMinDock strong{font-size:14px;color:#4f46e5;font-weight:900;font-feature-settings:"tnum"}
        .dockTrack{height:6px;background:#e2e8f0;border-radius:999px;overflow:hidden;margin-top:8px;border:1px solid #dbe4ff}
        .dockTrack span{display:block;height:100%;width:0;background:linear-gradient(90deg,#10b981,#4f46e5,#ec4899);border-radius:999px;transition:width .35s ease}
        .dockHint{font-size:9.5px;color:#94a3b8;font-weight:700;text-align:center;margin-top:5px}
        #adminRoot.adminMinimized{padding:6px !important;min-height:auto !important;background:transparent !important;overflow:hidden !important}
        #adminRoot.adminMinimized .dlgHeader,#adminRoot.adminMinimized .adminWindowTools,#adminRoot.adminMinimized .dlgFadeBody{display:none !important}
        #adminRoot.adminMinimized .adminMinDock{display:block}
        body:has(#adminRoot.adminMinimized){overflow:hidden !important}
        html:has(#adminRoot.adminMinimized){overflow:hidden !important}
        .adminTopBar{display:grid;grid-template-columns:minmax(280px,1fr) auto;gap:14px;align-items:center;background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:14px 16px;margin-bottom:14px;box-shadow:0 10px 26px -18px rgba(15,23,42,.18)}
        .adminSearchBox{position:relative;display:flex;align-items:center;background:linear-gradient(180deg,#f8fafc,#eef2ff);border:1px solid #d8dee7;border-radius:13px;padding:0 12px;transition:all .2s ease;min-width:0}
        .adminSearchBox:focus-within{border-color:#6366f1;box-shadow:0 0 0 4px rgba(99,102,241,.16);background:#fff}
        .adminSearchIcon{font-size:14px;margin-right:8px;opacity:.7}
        .adminSearchInput{flex:1;border:0;background:transparent;padding:11px 0;font-size:14px;color:#0f172a;outline:none}
        .adminKbd{font-size:10.5px;padding:2px 7px;border:1px solid #cbd5e1;border-radius:5px;background:#fff;color:#475569;font-family:'SF Mono',Consolas,monospace;font-weight:700}
        .adminTopStats{display:flex;gap:10px}
        .adminMiniStat{min-width:146px;background:linear-gradient(135deg,#fff,#f8fafc);border:1px solid #e5e7eb;border-radius:12px;padding:10px 12px;box-shadow:0 2px 6px -3px rgba(15,23,42,.08);transition:all .2s ease}
        .adminMiniStat:hover{transform:translateY(-2px);box-shadow:0 8px 16px -8px rgba(15,23,42,.16)}
        .adminMiniStat .lbl{font-size:10px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.06em}
        .adminMiniStat .val{font-size:18px;font-weight:900;color:#0f172a;margin:2px 0;letter-spacing:-0.015em}
        .adminMiniStat .sub{font-size:10.5px;color:#94a3b8;font-weight:600}
        .adminMiniStat .bar{height:5px;background:#e2e8f0;border-radius:4px;overflow:hidden;margin-top:5px}
        .adminMiniStat .bar span{display:block;height:100%;width:0;background:linear-gradient(90deg,#10b981,#7c3aed,#ec4899);background-size:200% 100%;animation:dlgShimmer 2.4s linear infinite;transition:width 1s cubic-bezier(.2,.7,.2,1);border-radius:4px}
        .adminBody{display:grid;grid-template-columns:228px minmax(0,1fr);gap:14px;margin-bottom:14px}
        .adminSide{background:linear-gradient(180deg,#fff,#f8fafc);border:1px solid #e5e7eb;border-radius:16px;padding:12px;box-shadow:0 10px 26px -18px rgba(15,23,42,.18);height:fit-content;position:sticky;top:0}
        .adminSideTitle{font-size:10px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.08em;margin:4px 8px 8px}
        .adminCat{display:flex;align-items:center;gap:10px;width:100%;text-align:left;padding:10px 12px;border:0;background:transparent;border-radius:10px;cursor:pointer;color:#334155;font-weight:600;font-size:13px;margin:2px 0;transition:all .2s ease;position:relative;font-family:inherit}
        .adminCat .ic{font-size:18px}
        .adminCat .cnt{margin-left:auto;font-size:10.5px;font-weight:800;color:#64748b;background:#f1f5f9;padding:2px 7px;border-radius:999px;letter-spacing:0.04em}
        .adminCat:hover{background:#eef2ff;color:#1e293b;transform:translateX(3px)}
        .adminCat.active{color:#fff;background:linear-gradient(135deg,#0f172a,#1e293b);box-shadow:0 6px 14px rgba(15,23,42,.30)}
        .adminCat.active .cnt{background:rgba(255,255,255,.18);color:#fff}
        .adminCat.active::before{content:"";position:absolute;left:-2px;top:9px;bottom:9px;width:4px;background:linear-gradient(180deg,#a78bfa,#ec4899);border-radius:4px;box-shadow:0 0 8px #a78bfa}
        .adminSideFoot{margin-top:12px;padding-top:10px;border-top:1px dashed #e2e8f0;display:flex;flex-direction:column;gap:5px}
        .adminQuickBtn{padding:7px 10px;border:1px solid #e5e7eb;background:#fff;border-radius:8px;cursor:pointer;font-size:11.5px;color:#334155;font-weight:600;text-align:left;transition:all .2s ease;font-family:inherit}
        .adminQuickBtn:hover{border-color:#a5b4fc;background:#eef2ff;color:#3730a3;transform:translateX(2px)}
        .adminMain{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:18px 20px;min-height:460px;box-shadow:0 10px 26px -18px rgba(15,23,42,.18)}
        .adminCatHeader{display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:2px solid #f1f5f9;margin-bottom:14px}
        .adminCatHeader .ic{width:42px;height:42px;border-radius:11px;display:inline-flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;box-shadow:0 4px 10px -4px rgba(15,23,42,.18)}
        .adminCatHeader .ttl{font-size:20px;font-weight:800;color:#0f172a;letter-spacing:-0.01em}
        .adminCatHeader .sub{font-size:12.5px;color:#64748b;font-weight:600;margin-top:1px}
        .adminGrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px}
        .adminCard{position:relative;background:linear-gradient(135deg,#ffffff,#fafbfd);border:1px solid #e5e7eb;border-radius:16px;padding:16px;cursor:pointer;transition:all .25s cubic-bezier(.2,.7,.2,1);overflow:hidden;animation:dlgSlideUp .35s cubic-bezier(.2,.7,.2,1) both;display:flex;flex-direction:column;min-height:132px;font-family:inherit}
        .adminCard::before{content:"";position:absolute;left:0;right:0;top:0;height:3px;background:linear-gradient(90deg,var(--cardColor,#1a73e8),var(--cardColor2,#6366f1));opacity:.85}
        .adminCard:hover{transform:translateY(-3px);box-shadow:0 14px 28px -14px rgba(15,23,42,.20),0 0 0 1px var(--cardColor,#1a73e8)}
        .adminCard:active{transform:translateY(-1px);transition:transform .08s}
        .adminCard .ch{display:flex;align-items:center;gap:10px;margin-bottom:6px}
        .adminCard .ch .ico{width:38px;height:38px;border-radius:10px;background:linear-gradient(135deg,var(--cardColor,#1a73e8)18,#fff);border:1px solid var(--cardColor,#1a73e8)33;display:inline-flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;box-shadow:inset 0 1px 0 #fff}
        .adminCard .ttl{font-size:13.5px;font-weight:800;color:#0f172a;line-height:1.3;letter-spacing:-0.01em;flex:1;min-width:0}
        .adminCard .desc{font-size:11.5px;color:#64748b;line-height:1.5;flex:1;font-weight:500}
        .adminCard .act{margin-top:10px;display:flex;align-items:center;justify-content:space-between;gap:8px}
        .adminCard .runBtn{padding:6px 12px;background:linear-gradient(135deg,var(--cardColor,#1a73e8),var(--cardColor,#1a73e8)d9);color:#fff;border:0;border-radius:8px;cursor:pointer;font-size:11.5px;font-weight:800;letter-spacing:0.02em;transition:all .2s ease;box-shadow:0 4px 10px -3px var(--cardColor,#1a73e8)66;font-family:inherit}
        .adminCard:hover .runBtn{filter:brightness(1.08);box-shadow:0 6px 14px -3px var(--cardColor,#1a73e8)80}
        .adminCard.danger .badge{background:#fee2e2;color:#991b1b;font-size:9.5px;font-weight:800;padding:2px 7px;border-radius:999px;letter-spacing:0.06em;text-transform:uppercase}
        .adminCard:not(.danger) .badge{display:none}
        .adminInlineTag{background:linear-gradient(135deg,#ddd6fe,#c4b5fd);color:#5b21b6;font-size:9.5px;font-weight:800;padding:2px 7px;border-radius:999px;letter-spacing:0.06em;text-transform:uppercase;border:1px solid #a78bfa66}
        .adminCard.has-inline::after{content:"";position:absolute;right:0;top:0;width:0;height:0;border-style:solid;border-width:0 28px 28px 0;border-color:transparent #a78bfa transparent transparent;opacity:.85}
        .adminEmpty{padding:60px 20px;text-align:center;color:#64748b;font-size:14px;font-weight:600;background:linear-gradient(135deg,#f8fafc,#fff);border:2px dashed #cbd5e1;border-radius:12px}
        .adminCat.cat-has-result .cnt{background:linear-gradient(135deg,#ddd6fe,#c4b5fd);color:#5b21b6;font-weight:900}
        .adminSearchCatLabel{margin-top:8px;}
        .adminFoot{display:flex;justify-content:space-between;align-items:center;gap:12px;background:#fff;border:1px solid #e5e7eb;border-radius:13px;padding:10px 14px;font-size:11.5px;color:#64748b;font-weight:600;box-shadow:0 4px 12px -8px rgba(15,23,42,.08)}
        .adminFootRight{color:#94a3b8}
        .adminMain{position:relative}
        @keyframes alSpin{to{transform:rotate(360deg)}}
        @keyframes alGlow{0%,100%{box-shadow:0 0 30px -8px rgba(124,58,237,.4)}50%{box-shadow:0 0 50px -4px rgba(236,72,153,.6)}}
        @keyframes alShim{0%{margin-left:-30%}100%{margin-left:100%}}
        .adminLoaderOverlay{position:absolute;inset:0;background:rgba(248,250,252,0.92);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);z-index:50;display:flex;align-items:center;justify-content:center;border-radius:14px;animation:dlgFadeIn .25s ease both}
        .adminLoaderCard{background:linear-gradient(135deg,#ffffff,#fafbfd);border:1px solid #e5e7eb;border-radius:16px;padding:32px 38px;display:flex;flex-direction:column;align-items:center;gap:12px;box-shadow:0 24px 60px -20px rgba(15,23,42,.30);min-width:340px;animation:alGlow 2.4s ease-in-out infinite}
        .adminLoaderRing{width:72px;height:72px;animation:alSpin .9s linear infinite;filter:drop-shadow(0 4px 12px rgba(124,58,237,.4))}
        .adminLoaderRing svg{width:100%;height:100%}
        .adminLoaderText{font-size:16px;font-weight:800;color:#0f172a;letter-spacing:-0.01em;text-align:center}
        .adminLoaderSub{font-size:12px;color:#64748b;font-weight:600;text-align:center;max-width:320px;line-height:1.5}
        .adminLoaderMeter{width:240px;height:5px;background:#e2e8f0;border-radius:6px;overflow:hidden;margin-top:6px;position:relative}
        .adminLoaderMeter > div{position:absolute;left:0;top:0;width:30%;height:100%;background:linear-gradient(90deg,transparent,#7c3aed,#ec4899,transparent);animation:alShim 1.4s ease-in-out infinite}
        .iv-bar{display:flex;align-items:center;gap:12px;padding-bottom:12px;margin-bottom:14px;border-bottom:2px solid #f1f5f9}
        .iv-back{padding:8px 14px;background:linear-gradient(135deg,#fff,#eef2ff);border:1px solid #c7d2fe;color:#3730a3;border-radius:10px;cursor:pointer;font-weight:800;font-size:12.5px;font-family:inherit;transition:all .2s ease;display:inline-flex;align-items:center;gap:4px;flex-shrink:0}
        .iv-back:hover{background:linear-gradient(135deg,#4f46e5,#6366f1);color:#fff;border-color:#4f46e5;transform:translateX(-3px);box-shadow:0 4px 12px rgba(79,70,229,.32)}
        .iv-bread{flex:1;display:flex;align-items:center;gap:8px;font-size:12.5px;color:#64748b;font-weight:600;min-width:0}
        .iv-bread-root{color:#94a3b8}
        .iv-bread-sep{color:#cbd5e1}
        .iv-bread-cur{color:#0f172a;font-weight:800;font-size:15px;letter-spacing:-0.01em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
        .iv-refresh{padding:8px 12px;background:#fff;border:1px solid #e5e7eb;color:#475569;border-radius:10px;cursor:pointer;font-size:14px;font-family:inherit;transition:all .2s ease;flex-shrink:0}
        .iv-refresh:hover{background:#eef2ff;color:#3730a3;border-color:#a5b4fc;transform:rotate(90deg)}
        body.iv-fullscreen .adminTopBar{display:none !important}
        body.iv-fullscreen .adminFoot{display:none !important}
        body.iv-fullscreen .adminSide{display:none !important}
        body.iv-fullscreen .adminBody{grid-template-columns:1fr !important;gap:0 !important}
        body.iv-fullscreen .adminMain{padding:14px 16px;min-height:calc(100vh - 150px)}
        body.iv-fullscreen .iv-bar{position:sticky;top:0;z-index:30;background:linear-gradient(180deg,#fff 0%,rgba(255,255,255,.96) 100%);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);padding:10px 0 12px;margin:-2px 0 14px;box-shadow:0 4px 12px -8px rgba(15,23,42,.08)}
        body.iv-fullscreen .iv-back{padding:10px 16px;font-size:13px}
        body.iv-fullscreen .iv-bread-cur{font-size:17px}
        @media(max-width:1050px){
          .adminTopBar{grid-template-columns:1fr}
          .adminTopStats{display:grid;grid-template-columns:repeat(2,minmax(0,1fr))}
          .adminMiniStat{min-width:0}
          .adminBody{grid-template-columns:1fr}
          .adminSide{position:relative}
          .adminSide #adminNav{display:flex;flex-wrap:wrap;gap:6px}
          .adminCat{width:auto;margin:0}
        }
        @media(max-width:720px){
          .adminTopStats{grid-template-columns:1fr}
          .adminGrid{grid-template-columns:1fr}
          .adminFoot{flex-direction:column;align-items:flex-start}
          .adminMain{padding:14px}
        }
      </style>
      <script>
        var ACTIONS = ${JSON.stringify(ACTIONS)};
        var CATEGORIES = ${JSON.stringify(CATEGORIES)};
        var INITIAL = ${JSON.stringify(stats)};
        var activeCat = 'home', searchQ = '', adminIsMinimized = false;
        function fmt(n){ return Number(n).toLocaleString('vi-VN'); }
        function esc(s){ return (s||'').replace(/[&<>"']/g, function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }
        function syncDockProgress(){
          var bar = document.getElementById('dlgProgressBar');
          var pct = document.getElementById('dlgProgressLabel');
          var st = document.getElementById('dlgProgressStatus');
          var dockBar = document.getElementById('dockBar');
          var dockPct = document.getElementById('dockPct');
          var dockStatus = document.getElementById('dockStatus');
          if (!dockBar || !bar) return;
          dockBar.style.width = bar.style.width || '0%';
          dockPct.textContent = pct ? pct.textContent : '0%';
          dockStatus.textContent = st ? st.textContent : 'Sẵn sàng';
        }
        function minimizeAdminPanel(){
          adminIsMinimized = true;
          syncDockProgress();
          document.getElementById('adminRoot').classList.add('adminMinimized');
          try { google.script.host.setWidth(420); google.script.host.setHeight(160); } catch(e) {}
        }
        function restoreAdminPanel(){
          adminIsMinimized = false;
          document.getElementById('adminRoot').classList.remove('adminMinimized');
          try { google.script.host.setWidth(1240); google.script.host.setHeight(820); } catch(e) {}
          setTimeout(syncDockProgress, 80);
        }
        setInterval(syncDockProgress, 700);
        function countCat(catId){
          if (searchQ){
            return ACTIONS.filter(function(a){ return a.cat===catId && matchSearch(a); }).length;
          }
          return ACTIONS.filter(function(a){ return a.cat===catId; }).length;
        }
        /* ── fuzzyTokenMatch: tách query thành từng token, mỗi token phải
           khớp prefix ít nhất 1 từ trong text (hoặc khớp substring phrase).
           Ví dụ: "t" → match "tìm","title","topic"; "sub" → match "subtitle" */
        function fuzzyTokenMatch(text, rawQ){
          if (!rawQ) return true;
          var t = (text||'').toLowerCase();
          var q = rawQ.toLowerCase().trim();
          if (!q) return true;
          // Kiểm tra substring phrase trước (giữ backward-compat)
          if (t.indexOf(q) >= 0) return true;
          // Token-prefix: mỗi token trong query phải có ít nhất 1 từ trong text bắt đầu bằng token đó
          var tokens = q.split(/\s+/).filter(Boolean);
          if (tokens.length < 2) {
            // Single token: tìm từ nào trong text bắt đầu bằng token
            var words = t.split(/[\s,.\-\/|]+/).filter(Boolean);
            return words.some(function(w){ return w.indexOf(tokens[0])===0; });
          }
          // Multi-token: ALL tokens phải match (AND logic)
          var words = t.split(/[\s,.\-\/|]+/).filter(Boolean);
          return tokens.every(function(tok){
            return t.indexOf(tok)>=0 || words.some(function(w){ return w.indexOf(tok)===0; });
          });
        }
        function matchSearch(a){
          if (!searchQ) return true;
          var haystack = (a.title||'') + ' ' + (a.desc||'') + ' ' + (a.fn||'') + ' ' + (a.cat||'');
          return fuzzyTokenMatch(haystack, searchQ);
        }
        function getCatLabel(catId){
          var c = CATEGORIES.find(function(x){return x.id===catId;});
          return c ? c.label : catId;
        }
        function getCatColor(catId){
          var c = CATEGORIES.find(function(x){return x.id===catId;});
          return c ? c.color : '#64748b';
        }
        function getCatIcon(catId){
          var c = CATEGORIES.find(function(x){return x.id===catId;});
          return c ? c.icon : '📋';
        }
        function renderNav(){
          var html = '';
          /* Khi đang search: highlight danh mục có kết quả, không active bất kỳ danh mục nào */
          CATEGORIES.forEach(function(c){
            var n = countCat(c.id);
            var isActive = searchQ ? false : (c.id===activeCat);
            var hasResult = searchQ && n > 0;
            html += '<button class="adminCat '+(isActive?'active':'')+(hasResult?' cat-has-result':'')+'" onclick="setCat(\\''+c.id+'\\')"><span class="ic">'+c.icon+'</span><span>'+esc(c.label)+'</span><span class="cnt">'+n+'</span></button>';
          });
          document.getElementById('adminNav').innerHTML = html;
        }
        function renderHeader(){
          if (searchQ){
            /* Khi tìm kiếm: hiển thị header tổng hợp tất cả danh mục */
            var total = ACTIONS.filter(matchSearch).length;
            var html = '<div class="ic" style="background:linear-gradient(135deg,#eef2ff,#fff);color:#4f46e5;border:1px solid #c7d2fe;font-size:20px;">🔎</div>'
              + '<div><div class="ttl" style="color:#4f46e5;">Kết quả tìm kiếm</div>'
              + '<div class="sub">Tìm thấy <b>'+total+'</b> chức năng khớp "<b>'+esc(searchQ)+'</b>" trong toàn bộ danh mục</div></div>';
            document.getElementById('adminCatHeader').innerHTML = html;
          } else {
            var c = CATEGORIES.find(function(x){return x.id===activeCat;}) || CATEGORIES[0];
            var html = '<div class="ic" style="background:linear-gradient(135deg,'+c.color+'2e,#fff);color:'+c.color+';border:1px solid '+c.color+'40;">'+c.icon+'</div>'
              + '<div><div class="ttl" style="color:'+c.color+';">'+esc(c.label)+'</div>'
              + '<div class="sub">'+countCat(activeCat)+' chức năng trong nhóm này</div></div>';
            document.getElementById('adminCatHeader').innerHTML = html;
          }
        }
        function renderGrid(){
          /* Khi search: hiển thị TẤT CẢ danh mục, nhóm theo category, không filter activeCat */
          var list;
          if (searchQ) {
            list = ACTIONS.filter(matchSearch);
          } else {
            list = ACTIONS.filter(function(a){ return a.cat===activeCat; });
          }
          if (!list.length){
            document.getElementById('adminGrid').innerHTML = '';
            document.getElementById('adminEmpty').style.display='block';
            document.getElementById('adminEmpty').innerHTML = searchQ
              ? '🔎 Không tìm thấy chức năng nào khớp với từ khóa <b>"'+esc(searchQ)+'"</b>.<br><span style="font-size:12px;font-weight:500;">Thử từ khóa khác: tên chức năng, mô tả hoặc nhóm (video, kênh, subtitle...)</span>'
              : 'Không có chức năng trong nhóm này.';
            return;
          }
          document.getElementById('adminEmpty').style.display='none';
          /* Khi search: render có nhãn danh mục để dễ phân biệt */
          var html = '';
          if (searchQ) {
            /* Nhóm theo cat */
            var grouped = {};
            var catOrder = [];
            list.forEach(function(a){
              if (!grouped[a.cat]){ grouped[a.cat]=[]; catOrder.push(a.cat); }
              grouped[a.cat].push(a);
            });
            catOrder.forEach(function(catId){
              var cColor = getCatColor(catId);
              var cIcon = getCatIcon(catId);
              var cLabel = getCatLabel(catId);
              html += '<div class="adminSearchCatLabel" style="grid-column:1/-1;display:flex;align-items:center;gap:8px;padding:4px 2px 2px;font-size:10px;font-weight:800;color:'+cColor+';text-transform:uppercase;letter-spacing:0.08em;border-bottom:1px solid '+cColor+'22;margin-bottom:2px;">'
                + '<span>'+cIcon+'</span><span>'+esc(cLabel)+'</span><span style="font-weight:500;color:#94a3b8;">('+grouped[catId].length+')</span>'
                + '</div>';
              grouped[catId].forEach(function(a, i){
                html += buildCard_(a, i);
              });
            });
          } else {
            list.forEach(function(a, i){ html += buildCard_(a, i); });
          }
          document.getElementById('adminGrid').innerHTML = html;
        }
        function buildCard_(a, i){
          var bg = a.color || '#1a73e8';
          var inlineAttr = a.inline ? ('\\''+a.inline+'\\'') : 'null';
          var safeTitle = esc(a.title).replace(/\\\\/g,"\\\\\\\\").replace(/\\'/g,"\\\\\\'");
          var inlineBadge = a.inline ? '<span class="adminInlineTag">⚡ INLINE</span>' : '';
          return '<div class="adminCard '+(a.danger?'danger':'')+(a.inline?' has-inline':'')+'" onclick="runAction(\\''+a.fn+'\\',\\''+safeTitle+'\\','+inlineAttr+')" style="--cardColor:'+bg+';--cardColor2:'+bg+'cc;animation-delay:'+(i*25)+'ms;">'
            + '<div class="ch"><div class="ico">'+a.icon+'</div><div class="ttl">'+esc(a.title)+'</div></div>'
            + '<div class="desc">'+esc(a.desc)+'</div>'
            + '<div class="act"><div style="display:flex;gap:5px;flex-wrap:wrap;">'+inlineBadge+'<span class="badge">⚠️ NẶNG</span></div><button class="runBtn" type="button">▶ '+(a.inline?'MỞ':'CHẠY')+'</button></div>'
            + '</div>';
        }
        function setCat(id){
          activeCat = id;
          /* Khi click danh mục → xóa search query để vào đúng danh mục */
          if (searchQ){ searchQ=''; document.getElementById('adminSearch').value=''; }
          renderNav(); renderHeader(); renderGrid();
        }
        function showLoader(text, sub){
          var ov = document.getElementById('adminLoaderOverlay');
          document.getElementById('adminLoaderText').textContent = text || 'Đang xử lý...';
          document.getElementById('adminLoaderSub').textContent = sub || 'Vui lòng chờ trong giây lát';
          ov.style.display = 'flex';
        }
        function hideLoader(){ document.getElementById('adminLoaderOverlay').style.display = 'none'; }
        var __currentInline = null, __currentInlineLabel = null;
        function loadInlineView(viewName, label){
          __currentInline = viewName; __currentInlineLabel = label;
          showLoader('Đang tải: '+label, 'Khởi tạo giao diện inline · '+viewName);
          if (window.dlgProgress) window.dlgProgress.start(2.5, 'Đang tải view: '+label);
          google.script.run
            .withSuccessHandler(function(html){
              hideLoader();
              if (window.dlgProgress) window.dlgProgress.complete('Đã tải: '+label);
              document.body.classList.add('iv-fullscreen');
              document.getElementById('adminHomeView').style.display='none';
              var iv = document.getElementById('adminInlineView');
              iv.style.display='block';
              document.getElementById('ivBreadCur').textContent = label;
              var ct = document.getElementById('adminInlineContent');
              ct.innerHTML = html;
              // Re-execute scripts injected via innerHTML
              ct.querySelectorAll('script').forEach(function(s){
                var ns = document.createElement('script');
                Array.from(s.attributes).forEach(function(a){ ns.setAttribute(a.name, a.value); });
                ns.textContent = s.textContent;
                s.parentNode.replaceChild(ns, s);
              });
            })
            .withFailureHandler(function(err){
              hideLoader();
              if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
              alert('Không tải được view: '+(err.message||err));
            })
            .getInlineView(viewName);
        }
        function goHome(){
          __currentInline = null; __currentInlineLabel = null;
          document.body.classList.remove('iv-fullscreen');
          document.getElementById('adminInlineView').style.display='none';
          document.getElementById('adminInlineContent').innerHTML='';
          document.getElementById('adminHomeView').style.display='block';
          if (window.dlgProgress) window.dlgProgress.reset();
          loadStats();
        }
        function reloadInlineView(){
          if (!__currentInline) return;
          loadInlineView(__currentInline, __currentInlineLabel);
        }
        function runAction(fn, label, inline){
          if (inline){
            loadInlineView(inline, label);
            return;
          }
          showLoader('Đang khởi chạy: '+label, 'Mở cửa sổ chức năng "'+label+'" · Bảng điều khiển sẽ tự đóng');
          if (window.dlgProgress) window.dlgProgress.start(2.5, 'Đang khởi chạy: '+label);
          google.script.run
            .withSuccessHandler(function(){
              if (window.dlgProgress) window.dlgProgress.complete('Đã khởi chạy: '+label);
              setTimeout(function(){ google.script.host.close(); }, 350);
            })
            .withFailureHandler(function(err){
              hideLoader();
              if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
              alert('Lỗi khi khởi chạy: '+(err.message||err));
            })
            .adminRun(fn);
        }
        function applyStats(s){
          // Quota
          var q = document.getElementById('msQuota');
          var qv = q.querySelector('.val');
          var qb = document.getElementById('msQuotaBar');
          qv.textContent = s.quotaPct.toFixed(1)+'%';
          qb.style.width = Math.min(100, s.quotaPct)+'%';
          var sub = q.querySelector('.sub'); if (sub) sub.remove();
          q.insertAdjacentHTML('beforeend', '<div class="sub">'+fmt(s.quotaUsed)+' / '+fmt(s.quotaLimit)+' units · '+fmt(s.quotaRequests)+' req</div>');
          // Health
          var h = document.getElementById('msHealth');
          var hv = h.querySelector('.val');
          var hb = document.getElementById('msHealthBar');
          hv.textContent = (s.healthScore!=null ? s.healthScore : '--');
          hb.style.width = Math.min(100, s.healthScore || 0)+'%';
          var hs = h.querySelector('.sub'); if (hs) hs.remove();
          h.insertAdjacentHTML('beforeend', '<div class="sub">'+esc(s.healthSummary)+'</div>');
          // Sheets
          document.getElementById('msSheets').textContent = fmt(s.videoRows + s.channelRows);
          document.getElementById('msSheetsSub').textContent = fmt(s.videoRows)+' video · '+fmt(s.channelRows)+' kênh';
          // OAuth
          document.getElementById('msOauth').textContent = s.oauthOk ? '✓ Sẵn sàng' : '✕ Chưa kết nối';
          document.getElementById('msOauth').style.color = s.oauthOk ? '#059669' : '#dc2626';
          document.getElementById('msOauthSub').textContent = s.oauthOk ? 'Refresh token đã lưu' : 'Cần chạy Cài đặt Access Token';
          document.getElementById('adminTs').textContent = 'Cập nhật lúc: '+s.timestamp;
        }
        function loadStats(){
          if (window.dlgProgress) window.dlgProgress.start(1.5, 'Đang làm mới stats');
          google.script.run.withSuccessHandler(function(s){
            applyStats(s);
            if (window.dlgProgress) window.dlgProgress.complete('Đã làm mới');
          }).withFailureHandler(function(err){
            if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
          }).getAdminDashboardStats();
        }
        document.getElementById('adminSearch').addEventListener('input', function(e){
          searchQ = (e.target.value||'').trim();
          renderNav(); renderHeader(); renderGrid();
        });
        document.addEventListener('keydown', function(e){
          if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='k'){ e.preventDefault(); if (adminIsMinimized) restoreAdminPanel(); document.getElementById('adminSearch').focus(); }
          if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='m'){ e.preventDefault(); adminIsMinimized ? restoreAdminPanel() : minimizeAdminPanel(); }
          if (e.key==='Escape'){
            if (adminIsMinimized){ restoreAdminPanel(); }
            else if (__currentInline){ goHome(); }
            else if (searchQ){ searchQ=''; document.getElementById('adminSearch').value=''; renderNav(); renderHeader(); renderGrid(); }
            else { minimizeAdminPanel(); }
          }
        });
        applyStats(INITIAL);
        renderNav(); renderHeader(); renderGrid();
      </script>
    </div>
  `).setWidth(1240).setHeight(820);
  SpreadsheetApp.getUi().showModelessDialog(html, '🎛️ Bảng điều khiển ');
}

