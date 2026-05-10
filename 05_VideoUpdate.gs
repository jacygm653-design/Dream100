// =========================================================================
// TÍNH NĂNG 1: XỬ LÝ SHEET VIDEO 
// =========================================================================

function updateFastVideoSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIDEO);
    if (!sheet) return ui.alert("Không tìm thấy Sheet: " + SHEET_VIDEO);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    executeVideoUpdate(sheet, 2, lastRow, 'fast');
    ui.alert(`✅ Hoàn tất Cập Nhật Nhanh!\nVPH đã được tính chuẩn xác theo công thức Rolling Window.\n\n${getYouTubeQuotaUsageMessage_()}`);
  } catch (error) {
    ui.alert(`❌ HỆ THỐNG GẶP LỖI:\n${error.message}`);
  }
}

function updateVideoSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIDEO);
    if (!sheet) return ui.alert("Không tìm thấy Sheet: " + SHEET_VIDEO);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    executeVideoUpdate(sheet, 2, lastRow, 'full');
    ui.alert(`✅ Hoàn tất cập nhật Video.\n\n${getYouTubeQuotaUsageMessage_()}`);
  } catch (error) {
    ui.alert(`❌ HỆ THỐNG GẶP LỖI:\n${error.message}`);
  }
}

function updateRangeVideoSheet() {
  openRangeActionDialog_("videoRange", SHEET_VIDEO, "Cập nhật Video theo KHOẢNG DÒNG", "Chọn khoảng dòng trên Sheet VIDEO để cập nhật metadata đầy đủ.", "#1a73e8");
}

/**
 * 🛠️ HÀM LÕI CẬP NHẬT VIDEO - VPH CHUẨN 100%
 * 
 * 📐 CÔNG THỨC VPH (ROLLING WINDOW):
 *    VPH = (V2 - V1) / ((t2 - t1) / 3600)
 *    Trong đó:
 *      V2 = Views hiện tại
 *      V1 = Views lần đo trước (lưu ở Cột O)
 *      (t2 - t1) = Khoảng thời gian giữa 2 lần đo (giây)
 *      → Quy đổi sang giờ bằng cách chia cho 3600
 * 
 * 📌 LOGIC:
 *   - Lần đầu chưa có timestamp (Cột N rỗng) → VPH = 0
 *   - Lần đầu chưa có views snapshot (Cột O rỗng) → VPH = 0
 *   - Khoảng cách < 1 phút → bỏ qua tránh nhiễu
 *   - viewDiff < 0 (views giảm do YouTube điều chỉnh) → VPH = 0
 */
function executeVideoUpdate(sheet, startRow, endRow, updateMode = 'full', progressId) {
  const apiKey = getApiKey();
  const numRows = endRow - startRow + 1;
  
  if (sheet.getMaxColumns() < COL_VIDEO.PREV_VIEWS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL_VIDEO.PREV_VIEWS - sheet.getMaxColumns());
  }
  
  if (!sheet.getRange(1, COL_VIDEO.LAST_UPDATE).getValue()) {
    sheet.getRange(1, COL_VIDEO.LAST_UPDATE).setValue("⏰ CẬP NHẬT LÚC")
      .setFontWeight("bold").setBackground("#4a86e8").setFontColor("white");
  }
  if (!sheet.getRange(1, COL_VIDEO.PREV_VIEWS).getValue()) {
    sheet.getRange(1, COL_VIDEO.PREV_VIEWS).setValue("📊 VIEWS LẦN ĐO TRƯỚC")
      .setFontWeight("bold").setBackground("#4a86e8").setFontColor("white");
  }
  if (!sheet.getRange(1, COL_VIDEO.SUBTITLE).getValue()) {
    sheet.getRange(1, COL_VIDEO.SUBTITLE).setValue("📝 SUBTITLE")
      .setFontWeight("bold").setBackground("#0f9d58").setFontColor("white");
  }
  
  const range = sheet.getRange(startRow, 1, numRows, COL_VIDEO.PREV_VIEWS);
  const data = range.getValues();
  
  const videoIds = [];
  for (let i = 0; i < data.length; i++) {
    const videoId = extractVideoId(data[i][COL_VIDEO.LINK - 1]);
    if (videoId) videoIds.push(videoId);
  }

  if (videoIds.length === 0) throw new Error("Không tìm thấy link video hợp lệ.");

  const videoUrls = [];
  for (let i = 0; i < videoIds.length; i += 50) {
    const batchIds = videoIds.slice(i, i + 50).join(',');
    videoUrls.push(`https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics,contentDetails&id=${batchIds}`);
  }
  const allVideoData = {};
  fetchAPIBatchWithRetry_(videoUrls, apiKey).forEach(json => {
    if (json.items) json.items.forEach(item => { allVideoData[item.id] = item; });
  });

  const channelIdSet = new Set();
  Object.values(allVideoData).forEach(v => {
    if (v.snippet && v.snippet.channelId) channelIdSet.add(v.snippet.channelId);
  });
  
  const scriptCache = CacheService.getScriptCache();
  const channelCountryMap = {};
  const uncachedChannelIds = [];
  
  channelIdSet.forEach(cid => {
    const cached = scriptCache.get('ch_country_' + cid);
    if (cached !== null && cached !== undefined) {
      channelCountryMap[cid] = cached === '__NULL__' ? '' : cached;
    } else {
      uncachedChannelIds.push(cid);
    }
  });
  
  if (uncachedChannelIds.length > 0) {
    const channelUrls = [];
    for (let i = 0; i < uncachedChannelIds.length; i += 50) {
      const batchIds = uncachedChannelIds.slice(i, i + 50).join(',');
      channelUrls.push(`https://www.googleapis.com/youtube/v3/channels?part=snippet&id=${batchIds}`);
    }
    try {
      fetchAPIBatchWithRetry_(channelUrls, apiKey).forEach(json => {
      if (json.items) {
        json.items.forEach(ch => {
          const country = (ch.snippet && ch.snippet.country) ? ch.snippet.country : '';
          channelCountryMap[ch.id] = country;
          try {
            scriptCache.put('ch_country_' + ch.id, country || '__NULL__', 21600);
          } catch(e) {}
        });
      }
      });
    } catch (e) {
      console.warn("Không lấy được country cho batch: " + e.message);
    }
  }

  const currentDate = new Date();
  const currentUpdateStr = Utilities.formatDate(currentDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");

  const rowsToWrap = [];

  // ── BATCH WRITE BUFFERS ─────────────────────────────────────────────────
  // fast-mode: buffer [views, vph, lastUpdate, prevViews] theo index dòng
  const fastBatchRows = [];   // { idx, views, vph, lastUpdate, prevViews }
  // error title buffer cho cả hai mode
  const errorTitleRows = [];  // { idx, title }
  // ────────────────────────────────────────────────────────────────────────

  for (let i = 0; i < data.length; i++) {
    const link = data[i][COL_VIDEO.LINK - 1]; 
    if (!link) {
      updateTaskProgress_(progressId, i + 1, data.length, "Đã cập nhật video " + (i + 1) + "/" + data.length);
      continue;
    }
    
    const videoId = extractVideoId(link);
    const actualRow = startRow + i; 

    if (videoId) {
      const video = allVideoData[videoId];
      if (video) {
        const snippet = video.snippet;
        const stats = video.statistics;
        const viewCount = parseInt(stats.viewCount || 0);
        
        const oldUpdateStr = data[i][COL_VIDEO.LAST_UPDATE - 1];
        const oldViewsRaw = data[i][COL_VIDEO.PREV_VIEWS - 1];
        let vphResult = 0; 
        let isFirstTime = false;

        if (oldUpdateStr && oldViewsRaw !== "" && oldViewsRaw !== null && oldViewsRaw !== undefined) {
          const oldTime = parseVietnameseDate(oldUpdateStr);
          const oldViews = unformatNumber(oldViewsRaw);
          
          if (oldTime instanceof Date && !isNaN(oldTime.getTime()) && oldViews >= 0) {
            const viewDiff = viewCount - oldViews;
            const timeDiffMs = currentDate.getTime() - oldTime.getTime();
            const timeDiffSeconds = timeDiffMs / 1000;
            const timeDiffHours = timeDiffSeconds / 3600;
            
            if (timeDiffSeconds >= 60 && viewDiff > 0) {
              vphResult = viewDiff / timeDiffHours;
            }
          }
        } else {
          isFirstTime = true;
        }

        const formattedVPH = formatNumber(Math.round(vphResult));
        const formattedViews = formatNumber(viewCount);

        if (updateMode === 'full') {
          const fullText = (snippet.title + " " + snippet.description).toLowerCase();
          data[i][COL_VIDEO.TITLE - 1] = snippet.title;                                
          data[i][COL_VIDEO.TAGS - 1] = snippet.tags ? snippet.tags.join(', ') : "";  
          data[i][COL_VIDEO.HASHTAGS - 1] = extractHashtags(fullText);                                      
          data[i][COL_VIDEO.VIEWS - 1] = formattedViews;                      
          data[i][COL_VIDEO.VPH - 1] = formattedVPH;                 
          
          const liveStatus = snippet.liveBroadcastContent || "none";
          if (liveStatus === "live") data[i][COL_VIDEO.DURATION - 1] = "Đang Live";
          else if (liveStatus === "upcoming") data[i][COL_VIDEO.DURATION - 1] = "Sắp chiếu";
          else data[i][COL_VIDEO.DURATION - 1] = parseISODuration(video.contentDetails ? video.contentDetails.duration : "");
          
          const thumbUrl = (snippet.thumbnails && snippet.thumbnails.high) ? snippet.thumbnails.high.url : '';
          data[i][COL_VIDEO.THUMBNAIL - 1] = thumbUrl ? `=IMAGE("${thumbUrl}")` : ""; 
          
          const channelCountry = channelCountryMap[snippet.channelId] || '';
          data[i][COL_VIDEO.PUBLISHED - 1] = snippet.publishedAt 
            ? formatPublishedAtMultiline(snippet.publishedAt, channelCountry)
            : "";
          rowsToWrap.push(actualRow);
          
          data[i][COL_VIDEO.DESCRIPTION - 1] = snippet.description || "";                         
          
          data[i][COL_VIDEO.LAST_UPDATE - 1] = currentUpdateStr;
          data[i][COL_VIDEO.PREV_VIEWS - 1] = viewCount;
          // ✅ KHÔNG ghi sheet ở đây — dồn vào batch write sau loop
          
        } else if (updateMode === 'fast') {
          // ✅ Buffer lại, batch write sau loop
          fastBatchRows.push({ i, formattedViews, formattedVPH, currentUpdateStr, viewCount });
        }

      } else {
        // Đánh dấu lỗi trong data để batch write cùng lúc
        data[i][COL_VIDEO.TITLE - 1] = "[Video bị Xóa hoặc Private]";
        errorTitleRows.push(i);
      }
    } else {
      data[i][COL_VIDEO.TITLE - 1] = "[Link sai định dạng]";
      errorTitleRows.push(i);
    }
    // ✅ Đã bỏ SpreadsheetApp.flush() thừa trong vòng lặp
    updateTaskProgress_(progressId, i + 1, data.length, "Đã cập nhật video " + (i + 1) + "/" + data.length);
  }

  // ── BATCH WRITE: ghi toàn bộ ra sheet chỉ 1 lần ──────────────────────
  if (updateMode === 'full') {
    // Ghi toàn bộ data (bao gồm cả dòng lỗi đã cập nhật title) 1 lần duy nhất
    sheet.getRange(startRow, 1, data.length, COL_VIDEO.PREV_VIEWS).setValues(data);
  } else if (updateMode === 'fast' && fastBatchRows.length > 0) {
    // Với fast-mode: chỉ 4 cột thay đổi — xây 4 mảng theo từng cột để setValues batch
    // Chuẩn bị 4 mảng song song theo index dòng
    const viewsArr  = data.map(r => [r[COL_VIDEO.VIEWS - 1]]);
    const vphArr    = data.map(r => [r[COL_VIDEO.VPH - 1]]);
    const luArr     = data.map(r => [r[COL_VIDEO.LAST_UPDATE - 1]]);
    const pvArr     = data.map(r => [r[COL_VIDEO.PREV_VIEWS - 1]]);
    // Áp giá trị mới từ buffer vào mảng
    fastBatchRows.forEach(({ i, formattedViews, formattedVPH, currentUpdateStr, viewCount }) => {
      viewsArr[i][0] = formattedViews;
      vphArr[i][0]   = formattedVPH;
      luArr[i][0]    = currentUpdateStr;
      pvArr[i][0]    = viewCount;
    });
    // 4 setValues thay vì N×4 setValue
    sheet.getRange(startRow, COL_VIDEO.VIEWS,       data.length, 1).setValues(viewsArr);
    sheet.getRange(startRow, COL_VIDEO.VPH,         data.length, 1).setValues(vphArr);
    sheet.getRange(startRow, COL_VIDEO.LAST_UPDATE, data.length, 1).setValues(luArr);
    sheet.getRange(startRow, COL_VIDEO.PREV_VIEWS,  data.length, 1).setValues(pvArr);
  }
  SpreadsheetApp.flush(); // ✅ Chỉ flush 1 lần duy nhất sau toàn bộ batch
  // ─────────────────────────────────────────────────────────────────────

  if (rowsToWrap.length > 0 && updateMode === 'full') {
    try {
      const minRow = Math.min(...rowsToWrap);
      const maxRow = Math.max(...rowsToWrap);
      const publishedRange = sheet.getRange(minRow, COL_VIDEO.PUBLISHED, maxRow - minRow + 1, 1);
      publishedRange.setWrap(true).setVerticalAlignment("middle");
    } catch(e) {
    }
  }
  
  recalculateSTT(SHEET_VIDEO);
}

function updateSpecificVideoRow() {
  openSingleRowActionDialog_("videoRow", SHEET_VIDEO, "Cập nhật Video theo dòng", "Nhập một dòng trên Sheet VIDEO để cập nhật metadata đầy đủ.", "#1a73e8");
}

