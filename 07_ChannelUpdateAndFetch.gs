// =========================================================================
// TÍNH NĂNG 2: XỬ LÝ SHEET KÊNH
// =========================================================================

function updateFastChannelSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHANNEL);
    if (!sheet) return;
    executeChannelUpdate(sheet, 2, sheet.getLastRow(), 'fast');
    ui.alert("✅ Đã cập nhật nhanh Subscribers & Views.\n\n" + getYouTubeQuotaUsageMessage_());
  } catch (error) { ui.alert(`❌ LỖI: ${error.message}`); }
}

function updateChannelSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHANNEL);
    if (!sheet) return;
    executeChannelUpdate(sheet, 2, sheet.getLastRow(), 'full');
    ui.alert("✅ Đã cập nhật xong toàn bộ Kênh.\n\n" + getYouTubeQuotaUsageMessage_());
  } catch (error) { ui.alert(`❌ LỖI: ${error.message}`); }
}

function updateRangeChannelSheet() {
  openRangeActionDialog_("channelRange", SHEET_CHANNEL, "Cập nhật Kênh theo KHOẢNG DÒNG", "Chọn khoảng dòng trên Sheet KÊNH để cập nhật metadata đầy đủ.", "#0f9d58");
}

function executeChannelUpdate(sheet, startRow, endRow, updateMode = 'full', progressId) {
  const apiKey = getApiKey();
  const numRows = endRow - startRow + 1;
  if (sheet.getMaxColumns() < COL_CHANNEL_MONETIZATION) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL_CHANNEL_MONETIZATION - sheet.getMaxColumns());
  }
  if (!sheet.getRange(1, COL_CHANNEL_MONETIZATION).getValue()) {
    sheet.getRange(1, COL_CHANNEL_MONETIZATION).setValue("KIẾM TIỀN (API)")
      .setFontWeight("bold").setBackground("#4a86e8").setFontColor("white");
  }
  const range = sheet.getRange(startRow, 1, numRows, COL_CHANNEL_MONETIZATION);
  const data = range.getValues();
  const channelCache = {}; 

  // ── BATCH WRITE BUFFERS cho fast-mode ───────────────────────────────────
  // fast-mode chỉ ghi 3 cột: 5 (subs), 7 (views4Weeks), 14 (monetization)
  // Chuẩn bị 3 mảng song song kích thước = data.length
  const fastSubsArr   = data.map(r => [r[4]]);            // cột 5 (index 4)
  const fastViewsArr  = data.map(r => [r[6]]);            // cột 7 (index 6)
  const fastMonetArr  = data.map(r => [r[COL_CHANNEL_MONETIZATION - 1]]); // cột 14
  let fastHasUpdate   = false;
  // ────────────────────────────────────────────────────────────────────────

  for (let i = 0; i < data.length; i++) {
    const channelLink = data[i][2]; 
    if (!channelLink) {
      updateTaskProgress_(progressId, i + 1, data.length, "Đã cập nhật kênh " + (i + 1) + "/" + data.length);
      continue;
    }

    let channelData;
    if (channelCache[channelLink]) channelData = channelCache[channelLink];
    else {
      channelData = getChannelDataFromAPI(channelLink, apiKey);
      if (channelData) channelCache[channelLink] = channelData;
    }
    
    if (channelData && channelData.found) {
      if (updateMode === 'full') {
        data[i][1] = channelData.title; data[i][3] = channelData.keywords; data[i][4] = formatNumber(channelData.subscribers); 
        data[i][5] = channelData.vidsPerMonth; data[i][6] = formatNumber(channelData.views4Weeks); data[i][7] = channelData.avgDuration; 
        data[i][8] = channelData.commonUploadHour; data[i][9] = channelData.mainTopic; data[i][10] = channelData.description; 
        data[i][11] = channelData.country; data[i][12] = channelData.category; data[i][13] = channelData.monetizationStatus;
        // ✅ KHÔNG ghi ngay — dồn vào batch write sau loop
      } else if (updateMode === 'fast') {
        // ✅ Buffer vào mảng, batch write sau loop
        fastSubsArr[i][0]  = formatNumber(channelData.subscribers);
        fastViewsArr[i][0] = formatNumber(channelData.views4Weeks);
        fastMonetArr[i][0] = channelData.monetizationStatus;
        fastHasUpdate = true;
      }
    } else {
      // Đánh dấu lỗi trong data để batch write cùng lúc
      data[i][1] = "[Không tìm thấy Kênh]";
    }
    // ✅ Đã bỏ SpreadsheetApp.flush() thừa trong vòng lặp
    updateTaskProgress_(progressId, i + 1, data.length, "Đã cập nhật kênh " + (i + 1) + "/" + data.length);
  }

  // ── BATCH WRITE: ghi ra sheet chỉ 1 lần ─────────────────────────────
  if (updateMode === 'full') {
    sheet.getRange(startRow, 1, data.length, COL_CHANNEL_MONETIZATION).setValues(data);
  } else if (updateMode === 'fast' && fastHasUpdate) {
    sheet.getRange(startRow, 5,                    data.length, 1).setValues(fastSubsArr);
    sheet.getRange(startRow, 7,                    data.length, 1).setValues(fastViewsArr);
    sheet.getRange(startRow, COL_CHANNEL_MONETIZATION, data.length, 1).setValues(fastMonetArr);
  }
  SpreadsheetApp.flush(); // ✅ Chỉ flush 1 lần duy nhất sau toàn bộ batch
  // ─────────────────────────────────────────────────────────────────────
  recalculateSTT(SHEET_CHANNEL);
}

function updateSpecificChannelRow() {
  openSingleRowActionDialog_("channelRow", SHEET_CHANNEL, "Cập nhật Kênh theo dòng", "Nhập một dòng trên Sheet KÊNH để cập nhật metadata đầy đủ.", "#0f9d58");
}

// =========================================================================
// LẤY LINKS TỪ TAB VIDEO (< 3 THÁNG) - LỌC THEO CHỦ ĐỀ & VIEWS
// =========================================================================
function fetchRecentVideosFromChannels() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const channelSheet = ss.getSheetByName(SHEET_CHANNEL);
  const videoSheet = ss.getSheetByName(SHEET_VIDEO);
  if (!channelSheet || !videoSheet) return ui.alert("❌ Lỗi: Không tìm thấy Sheet Kênh hoặc Sheet Video.");

  const lastChannelRow = channelSheet.getLastRow();
  if (lastChannelRow < 2) return ui.alert(`Sheet "${SHEET_CHANNEL}" chưa có link kênh ở cột C.`);
  setupFetchVideoHistorySheet_();
  const channelPickerItems = channelSheet.getRange(2, 2, lastChannelRow - 1, 2).getDisplayValues()
    .map((r, idx) => ({
      row: idx + 2,
      name: (r[0] || "").toString().trim() || "Kênh chưa có tên",
      link: (r[1] || "").toString().trim()
    }))
    .filter(item => item.link);

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif; padding:18px; line-height:1.45; color:#202124; background:#f8fafc;">
      ${buildDialogHeader_("LẤY VIDEO TRONG KÊNH", "Đọc link kênh từ cột C của Sheet KÊNH, sau đó lọc video theo thời gian, thời lượng và views.", "#1a73e8")}

      <table style="width:100%; border-collapse:collapse; background:#ffffff; border:1px solid #dfe3eb; border-radius:8px; overflow:hidden;">
        <tr><td colspan="2" style="padding:8px 0;">
          <b>📺 Chọn kênh cần quét</b> <span style="font-size:11px;color:#666;">(có thể chọn nhiều kênh để tiết kiệm quota)</span>
        </td></tr>
        <tr><td colspan="2">
          <div style="position:relative;">
            <button type="button" onclick="toggleChannelPicker()" id="channelPickerBtn"
              style="width:100%;padding:10px;border:2px solid #1a73e8;border-radius:6px;background:#fff;text-align:left;cursor:pointer;">
              Chưa chọn kênh nào
            </button>
            <div id="channelPicker" style="display:none;position:absolute;z-index:999;left:0;right:0;top:44px;background:#fff;border:1px solid #cbd5e1;border-radius:8px;box-shadow:0 10px 30px rgba(15,23,42,.16);padding:10px;">
              <input id="channelSearch" oninput="renderChannelOptions()" placeholder="Tìm theo tên kênh..."
                style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #cbd5e1;border-radius:6px;margin-bottom:8px;">
              <div style="display:flex;gap:8px;margin-bottom:8px;">
                <button type="button" onclick="selectVisibleChannels(true)" style="padding:6px 10px;border:1px solid #cbd5e1;background:#f8fafc;border-radius:6px;cursor:pointer;">Chọn tất cả</button>
                <button type="button" onclick="selectVisibleChannels(false)" style="padding:6px 10px;border:1px solid #cbd5e1;background:#f8fafc;border-radius:6px;cursor:pointer;">Bỏ chọn</button>
              </div>
              <div id="channelOptions" style="max-height:230px;overflow:auto;border-top:1px solid #eef2f7;padding-top:6px;"></div>
            </div>
          </div>
        </td></tr>

        <tr><td colspan="2" style="padding:8px 0;"><b>📌 Chủ đề lọc thêm</b> <span style="font-size:11px; color:#666;">(có thể để trống)</span></td></tr>
        <tr><td colspan="2"><input id="topicKeywords" type="text" placeholder="VD: psychology, mental health, self improvement"
          style="width:100%; padding:9px; box-sizing:border-box; border:2px solid #1a73e8; border-radius:4px;"></td></tr>

        <tr><td style="width:50%; padding:10px 5px 5px 0;">
          <b>👁️ Views tối thiểu:</b><br>
          <input id="minViews" type="number" value="50000" min="0" style="width:95%; padding:7px; border:1px solid #ccc; border-radius:4px;">
        </td><td style="padding:10px 0 5px 5px;">
          <b>📅 Mới trong vòng (ngày):</b><br>
          <input id="daysBack" type="number" value="90" min="1" max="90" style="width:95%; padding:7px; border:1px solid #ccc; border-radius:4px;">
        </td></tr>

        <tr><td style="width:50%; padding:10px 5px 5px 0;">
          <b>⏱️ Thời lượng tối thiểu:</b><br>
          <input id="minDuration" type="text" value="00:03:10" style="width:95%; padding:7px; border:1px solid #ccc; border-radius:4px;">
        </td><td style="padding:10px 0 5px 5px;">
          <b>📍 Phạm vi quét:</b><br>
          <div id="selectedChannelHint" style="padding:8px;border:1px solid #dfe3eb;border-radius:4px;background:#f8fafc;font-size:12px;color:#475569;">Chỉ quét các kênh đã tick trong danh sách.</div>
          <input id="rowRange" type="hidden" value="2-${lastChannelRow}">
        </td></tr>

        <tr><td colspan="2" style="padding:10px 0 5px;">
          <b>🔎 Cách khớp chủ đề:</b><br>
          <select id="matchMode" style="width:100%; padding:7px; border:1px solid #ccc; border-radius:4px;">
            <option value="any" selected>Khớp ít nhất 1 từ/cụm từ</option>
            <option value="all">Bắt buộc khớp tất cả từ/cụm từ</option>
          </select>
        </td></tr>

        <tr><td colspan="2" style="padding:10px 0 5px;">
          <input type="checkbox" id="autoUpdate" checked> <label for="autoUpdate">Tự động cập nhật metadata sau khi thêm link</label>
        </td></tr>
      </table>

      <div style="margin-top:16px; text-align:center;">
        <button onclick="runFetch()" id="btnRun" style="padding:12px 28px; background:#1a73e8; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold;">🚀 BẮT ĐẦU QUÉT</button>
        <button onclick="google.script.host.close()" style="padding:12px 18px; background:#9aa0a6; color:white; border:none; border-radius:6px; cursor:pointer; margin-left:8px;">HỦY</button>
      </div>

      <div id="status" style="margin-top:14px; padding:12px; background:#ffffff; border:1px solid #dfe3eb; border-radius:8px; min-height:34px; font-size:13px;"></div>
      <div style="margin-top:12px;background:#fff;border:1px solid #dfe3eb;border-radius:8px;padding:10px;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
          <b>Lịch sử lấy video trong kênh</b>
          <div>
            <button onclick="refreshFetchHistory()" style="padding:6px 10px;border:1px solid #cbd5e1;background:#fff;border-radius:6px;cursor:pointer;">Tải lại</button>
            <button onclick="deleteSelectedFetchHistory()" style="padding:6px 10px;border:1px solid #fecaca;background:#fff7f7;color:#b91c1c;border-radius:6px;cursor:pointer;">Xóa đã chọn</button>
            <button onclick="clearAllFetchHistory()" style="padding:6px 10px;border:1px solid #b91c1c;background:#b91c1c;color:#fff;border-radius:6px;cursor:pointer;">Xóa toàn bộ</button>
          </div>
        </div>
        <div id="fetchHistoryBox" style="max-height:180px;overflow:auto;margin-top:8px;font-size:12px;"></div>
      </div>

      <script>
        const CHANNELS = ${JSON.stringify(channelPickerItems)};
        let FETCH_HISTORY = ${JSON.stringify(getFetchVideoHistoryItems_(30))};
        const selectedRows = new Set();
        function fmt(n) { return (Number(n) || 0).toLocaleString('vi-VN'); }
        function escapeHtml(s) {
          return (s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
        }
        function toggleChannelPicker() {
          const el = document.getElementById('channelPicker');
          el.style.display = el.style.display === 'none' ? 'block' : 'none';
          renderChannelOptions();
        }
        /* chFuzzyMatch: prefix-token matching cho ô tìm kênh trong Fetch-Video-in-Channel picker */
        function chFuzzyMatch(text, rawQ) {
          if (!rawQ) return true;
          const t = (text || '').toLowerCase();
          const q = rawQ.toLowerCase().trim();
          if (!q) return true;
          if (t.includes(q)) return true;
          const tokens = q.split(/\s+/).filter(Boolean);
          const words = t.split(/[\s,.\-\/|]+/).filter(Boolean);
          if (tokens.length < 2) return words.some(w => w.startsWith(tokens[0]));
          return tokens.every(tok => t.includes(tok) || words.some(w => w.startsWith(tok)));
        }
        function renderChannelOptions() {
          const q = (document.getElementById('channelSearch').value || '').toLowerCase().trim();
          const filtered = CHANNELS.filter(c => {
            if (!q) return true;
            return chFuzzyMatch(c.name + ' ' + c.link + ' ' + String(c.row), q);
          });
          document.getElementById('channelOptions').innerHTML = filtered.length ? filtered.map(c =>
            '<label style="display:grid;grid-template-columns:18px minmax(0,1fr) minmax(160px,42%);gap:8px;align-items:center;padding:7px;border-bottom:1px solid #f1f5f9;cursor:pointer;">' +
            '<input type="checkbox" class="channelPick" data-row="' + c.row + '" ' + (selectedRows.has(c.row) ? 'checked' : '') + ' onchange="toggleChannelRow(' + c.row + ', this.checked)">' +
            '<b style="font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">' + escapeHtml(c.name) + '</b>' +
            '<span style="font-size:11px;color:#64748b;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right;">' + escapeHtml(c.link) + '</span>' +
            '</label>'
          ).join('') : '<div style="padding:10px;color:#64748b;">Không có kênh phù hợp.</div>';
          updateChannelPickerButton();
        }
        function toggleChannelRow(row, checked) {
          if (checked) selectedRows.add(row);
          else selectedRows.delete(row);
          updateChannelPickerButton();
        }
        function selectVisibleChannels(checked) {
          Array.from(document.querySelectorAll('.channelPick')).forEach(cb => {
            cb.checked = checked;
            toggleChannelRow(parseInt(cb.dataset.row, 10), checked);
          });
          renderChannelOptions();
        }
        function updateChannelPickerButton() {
          const total = CHANNELS.length;
          const picked = selectedRows.size;
          document.getElementById('channelPickerBtn').textContent = picked === 0 ? 'Chưa chọn kênh nào' : (picked === total ? 'Đang chọn toàn bộ ' + total + ' kênh có link' : 'Đã chọn ' + picked + '/' + total + ' kênh');
          document.getElementById('selectedChannelHint').textContent = picked === 0 ? 'Chỉ quét các kênh đã tick trong danh sách.' : 'Sẽ quét ' + picked + ' kênh đã chọn.';
        }
        function parseRowRange(text) {
          const raw = (text || '').trim();
          if (!raw) return { startRow: 2, endRow: ${lastChannelRow} };
          if (raw.indexOf('-') >= 0) {
            const parts = raw.split('-');
            return { startRow: parseInt(parts[0].trim()), endRow: parseInt(parts[1].trim()) };
          }
          return { startRow: parseInt(raw), endRow: ${lastChannelRow} };
        }
        function renderResult(res) {
          const s = res.stats || {};
          const cards = [
            ['Kênh đã kiểm tra', fmt(s.channelsProcessed || 0)],
            ['Video đã xét', fmt(s.videosChecked || 0)],
            ['Đạt bộ lọc', fmt(s.passed || 0)],
            ['Đã thêm', fmt(s.added || 0)]
          ].map(c => '<div style="display:inline-block; width:23%; min-width:112px; box-sizing:border-box; margin:4px; padding:10px; border:1px solid #dfe3eb; border-radius:8px; background:#f8fafc; text-align:center;"><div style="font-size:11px; color:#5f6368;">' + c[0] + '</div><div style="font-size:20px; color:#1a73e8; font-weight:bold;">' + c[1] + '</div></div>').join('');
          return '<div>' + cards + '</div><div style="white-space:pre-line; margin-top:8px;">' + res.message + '</div>';
        }
        function renderFetchHistory(items) {
          FETCH_HISTORY = items || FETCH_HISTORY || [];
          document.getElementById('fetchHistoryBox').innerHTML = FETCH_HISTORY.length ? FETCH_HISTORY.map(h =>
            '<label style="display:grid;grid-template-columns:18px 145px 1fr 90px;gap:8px;align-items:start;border-bottom:1px solid #f1f5f9;padding:7px 0;">' +
            '<input type="checkbox" class="fetchHistoryPick" value="' + h.id + '">' +
            '<span style="color:#64748b;">' + escapeHtml(h.time || '') + '</span>' +
            '<span><b>' + escapeHtml(h.channels || '') + '</b><br><span style="color:#64748b;">' + escapeHtml(h.message || '') + '</span></span>' +
            '<span style="text-align:right;color:#1a73e8;">' + escapeHtml(h.stats || '') + '</span>' +
            '</label>'
          ).join('') : '<div style="color:#64748b;padding:8px 0;">ChÆ°a cÃ³ lá»‹ch sá»­.</div>';
        }
        function pickedFetchHistoryIds() {
          return Array.from(document.querySelectorAll('.fetchHistoryPick:checked')).map(x => parseInt(x.value, 10));
        }
        function refreshFetchHistory() {
          google.script.run.withSuccessHandler(res => {
            if (res.success) renderFetchHistory(res.history || []);
          }).getFetchVideoHistory();
        }
        function deleteSelectedFetchHistory() {
          const ids = pickedFetchHistoryIds();
          if (!ids.length) return alert('Vui lÃ²ng chá»n lá»‹ch sá»­ cáº§n xÃ³a.');
          google.script.run.withSuccessHandler(res => {
            alert(res.message);
            renderFetchHistory(res.history || []);
          }).deleteFetchVideoHistory(ids);
        }
        function clearAllFetchHistory() {
          if (!confirm('XÃ³a toÃ n bá»™ lá»‹ch sá»­ láº¥y video trong kÃªnh?')) return;
          google.script.run.withSuccessHandler(res => {
            alert(res.message);
            renderFetchHistory(res.history || []);
          }).clearFetchVideoHistory();
        }
        function runFetch() {
          const topicKeywords = document.getElementById('topicKeywords').value.trim();
          const rowRange = parseRowRange(document.getElementById('rowRange').value);
          const pickedRows = Array.from(selectedRows).sort((a,b) => a - b);
          if (pickedRows.length === 0) {
            alert('Vui lòng chọn ít nhất 1 kênh để quét.');
            return;
          }
          if (!rowRange.startRow || !rowRange.endRow || rowRange.startRow < 2 || rowRange.endRow < rowRange.startRow) {
            alert('Khoảng dòng kênh không hợp lệ.');
            return;
          }
          const params = {
            startRow: rowRange.startRow,
            endRow: Math.min(rowRange.endRow, ${lastChannelRow}),
            selectedRows: pickedRows,
            topicKeywords: topicKeywords,
            minViews: Number.isNaN(parseInt(document.getElementById('minViews').value)) ? 50000 : parseInt(document.getElementById('minViews').value),
            daysBack: Number.isNaN(parseInt(document.getElementById('daysBack').value)) ? 90 : parseInt(document.getElementById('daysBack').value),
            minDuration: document.getElementById('minDuration').value,
            matchMode: document.getElementById('matchMode').value,
            autoUpdate: document.getElementById('autoUpdate').checked
          };
          document.getElementById('btnRun').disabled = true;
          document.getElementById('btnRun').innerHTML = '⏳ Đang quét...';
          document.getElementById('status').innerHTML = '🔄 Đang đọc video từ các kênh và lọc theo dữ liệu YouTube API...';
          google.script.run
            .withSuccessHandler(function(res) {
              document.getElementById('status').innerHTML = (res.success ? '✅ ' : '❌ ') + renderResult(res);
              document.getElementById('btnRun').innerHTML = res.success ? '✅ HOÀN TẤT' : '🔄 THỬ LẠI';
              document.getElementById('btnRun').disabled = res.success;
              refreshFetchHistory();
            })
            .withFailureHandler(function(err) {
              document.getElementById('status').innerHTML = '❌ Lỗi: ' + err.message;
              document.getElementById('btnRun').disabled = false;
              document.getElementById('btnRun').innerHTML = '🔄 THỬ LẠI';
            })
            .executeFetchRecentVideosFromChannels(params);
        }
        renderChannelOptions();
        renderFetchHistory();
      </script>
    </div>
  `).setWidth(680).setHeight(620);

  ui.showModalDialog(html, 'Lấy video trong kênh');
}

function normalizeTopicText_(text) {
  return (text || "")
    .toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s#_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseTopicTerms_(topicKeywords) {
  return (topicKeywords || "")
    .toString()
    .split(',')
    .map(t => normalizeTopicText_(t))
    .filter(t => t);
}

function videoMatchesTopic_(video, topicTerms, matchMode) {
  if (!topicTerms || topicTerms.length === 0) return true;
  const snippet = video.snippet || {};
  const haystack = normalizeTopicText_([
    snippet.title || "",
    snippet.description || "",
    snippet.tags ? snippet.tags.join(' ') : ""
  ].join(' '));
  if (!haystack) return false;
  return matchMode === 'all'
    ? topicTerms.every(term => haystack.includes(term))
    : topicTerms.some(term => haystack.includes(term));
}

function parseDurationToSeconds_(durationValue) {
  if (durationValue === null || durationValue === undefined || durationValue === "") return 0;
  const text = durationValue.toString().trim();
  const iso = text.match(/^PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?$/i);
  if (iso) {
    return (parseInt(iso[1] || 0, 10) * 3600) + (parseInt(iso[2] || 0, 10) * 60) + parseInt(iso[3] || 0, 10);
  }
  const parts = text.split(':').map(p => parseInt(p, 10));
  if (parts.some(p => isNaN(p))) return parseInt(text, 10) || 0;
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  if (parts.length === 2) return parts[0] * 60 + parts[1];
  return parts[0] || 0;
}

function hideSheetIfPossible_(sheet) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (sheet && !sheet.isSheetHidden() && ss.getSheets().length > 1) sheet.hideSheet();
  } catch (e) {}
}

function setupFetchVideoHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FETCH_VIDEO_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(FETCH_VIDEO_HISTORY_SHEET);
    sheet.getRange(1, 1, 1, 10).setValues([[
      "TIME", "CHANNEL_ROWS", "CHANNELS", "TOPIC", "DAYS", "MIN_DURATION", "MIN_VIEWS", "STATS", "MESSAGE", "CONFIG"
    ]]).setFontWeight("bold").setBackground("#1f2937").setFontColor("white");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 10);
  }
  hideSheetIfPossible_(sheet);
  return sheet;
}

function getFetchVideoHistoryItems_(limit) {
  const sheet = setupFetchVideoHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const take = Math.min(limit || 30, lastRow - 1);
  const values = sheet.getRange(lastRow - take + 1, 1, take, 10).getDisplayValues();
  return values.reverse().map((r, idx) => ({
    id: lastRow - idx,
    time: r[0],
    rows: r[1],
    channels: r[2],
    topic: r[3],
    days: r[4],
    minDuration: r[5],
    minViews: r[6],
    stats: r[7],
    message: r[8],
    config: r[9]
  }));
}

function deleteRowsByIds_(sheet, ids) {
  const rows = (ids || []).map(id => parseInt(id, 10)).filter(id => id >= 2 && id <= sheet.getLastRow());
  return deleteRowsInBlocks_(sheet, rows);
}

function getFetchVideoHistory() {
  try {
    return { success: true, history: getFetchVideoHistoryItems_(30) };
  } catch (e) {
    return { success: false, message: e.message, history: [] };
  }
}

function deleteFetchVideoHistory(ids) {
  try {
    const sheet = setupFetchVideoHistorySheet_();
    const deleted = deleteRowsByIds_(sheet, ids);
    return { success: true, message: `Đã xóa ${deleted} lịch sử.`, history: getFetchVideoHistoryItems_(30) };
  } catch (e) {
    return { success: false, message: e.message, history: [] };
  }
}

function clearFetchVideoHistory() {
  try {
    const sheet = setupFetchVideoHistorySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    return { success: true, message: "Đã xóa toàn bộ lịch sử lấy video trong kênh.", history: [] };
  } catch (e) {
    return { success: false, message: e.message, history: [] };
  }
}

function executeFetchRecentVideosFromChannels(params) {
  const stats = {
    channelsProcessed: 0,
    videosChecked: 0,
    passed: 0,
    added: 0,
    duplicate: 0,
    tooOld: 0,
    tooShort: 0,
    tooFewViews: 0,
    topicMismatch: 0,
    missingDetails: 0
  };

  try {
    const topicTerms = parseTopicTerms_(params.topicKeywords);

    const minViews = params.minViews === undefined ? 50000 : Math.max(0, parseInt(params.minViews, 10) || 0);
    const daysBack = params.daysBack === undefined ? 90 : Math.min(90, Math.max(1, parseInt(params.daysBack, 10) || 90));
    const minDurationSeconds = params.minDuration === undefined ? 190 : Math.max(0, parseDurationToSeconds_(params.minDuration) || 0);
    const matchMode = params.matchMode === 'all' ? 'all' : 'any';
    const rangeParams = { start: parseInt(params.startRow, 10), end: parseInt(params.endRow, 10) };
    if (!rangeParams.start || !rangeParams.end || rangeParams.start < 2 || rangeParams.end < rangeParams.start) {
      return { success: false, message: "Khoảng dòng không hợp lệ.", stats };
    }

    const apiKey = getApiKey();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const channelSheet = ss.getSheetByName(SHEET_CHANNEL);
    const videoSheet = ss.getSheetByName(SHEET_VIDEO);

    if (!channelSheet || !videoSheet) {
      return { success: false, message: "Không tìm thấy Sheet Kênh hoặc Sheet Video.", stats };
    }

    const selectedRows = Array.isArray(params.selectedRows)
      ? params.selectedRows.map(r => parseInt(r, 10)).filter(r => r >= 2 && r <= channelSheet.getLastRow())
      : [];
    const uniqueRows = [...new Set(selectedRows)].sort((a, b) => a - b);
    const selectedRowSet = new Set(uniqueRows);
    const readStart = uniqueRows.length > 0 ? Math.min(...uniqueRows) : rangeParams.start;
    const readEnd = uniqueRows.length > 0 ? Math.max(...uniqueRows) : rangeParams.end;
    const channelRows = channelSheet.getRange(readStart, 2, readEnd - readStart + 1, 2).getDisplayValues();
    const channelEntries = channelRows
      .map((r, idx) => ({ row: readStart + idx, name: r[0], link: r[1] }))
      .filter(item => uniqueRows.length === 0 || selectedRowSet.has(item.row));
    if (channelEntries.length === 0) {
      return { success: false, message: "Chưa chọn kênh nào để quét.", stats };
    }
    initTaskProgress_(params.progressId, channelEntries.length, "Đang quét video trong kênh 0/" + channelEntries.length);

    const cutoffDate = new Date(); 
    cutoffDate.setDate(cutoffDate.getDate() - daysBack);

    const existingSet = new Set();
    const lastVideoRow = videoSheet.getLastRow();
    if (lastVideoRow >= 2) {
      const existingLinks = videoSheet.getRange(2, 3, lastVideoRow - 1, 1).getValues();
      existingLinks.forEach(row => {
         const id = extractVideoId(row[0]);
         if(id) existingSet.add(id);
      });
    }

    const candidateIds = [];
    const runLabel = uniqueRows.length > 0 ? `Đã chọn ${uniqueRows.length} kênh` : `Dòng ${rangeParams.start}-${rangeParams.end}`;
    const channelNames = channelEntries.filter(c => c.link).map(c => c.name || ("Kênh " + c.row));
    const selectedRowLabel = channelEntries.filter(c => c.link).map(c => c.row).join(", ");
    
    for (let i = 0; i < channelEntries.length; i++) {
      const urlStr = channelEntries[i].link;
      if (!urlStr) {
        updateTaskProgress_(params.progressId, i + 1, channelEntries.length, "Đã quét kênh " + (i + 1) + "/" + channelEntries.length);
        continue;
      }

      let videoTabPlaylistId;
      try {
        videoTabPlaylistId = getVideoTabPlaylistId(urlStr, apiKey);
      } catch (e) {
        updateTaskProgress_(params.progressId, i + 1, channelEntries.length, "Đã quét kênh " + (i + 1) + "/" + channelEntries.length);
        continue;
      }
      
      if (!videoTabPlaylistId) {
        updateTaskProgress_(params.progressId, i + 1, channelEntries.length, "Đã quét kênh " + (i + 1) + "/" + channelEntries.length);
        continue;
      }

      let pageToken = "";
      let keepFetching = true;

      while(keepFetching) {
          let apiUrl = `https://www.googleapis.com/youtube/v3/playlistItems?part=snippet,contentDetails&playlistId=${videoTabPlaylistId}&maxResults=50`;
          if (pageToken) apiUrl += `&pageToken=${pageToken}`;

          let json;
          try {
            json = fetchAPIWithRetry(apiUrl, apiKey);
          } catch (e) {
            console.warn(`Đã bỏ qua kênh dòng ${channelEntries[i].row} do lỗi lấy playlist: ${urlStr} | ${e.message}`);
            break; 
          }

          if (!json || !json.items || json.items.length === 0) break;

          for (let j = 0; j < json.items.length; j++) {
              const item = json.items[j];
              const publishedRaw = (item.contentDetails && item.contentDetails.videoPublishedAt) || (item.snippet && item.snippet.publishedAt);
              const pubDate = new Date(publishedRaw);

              if (pubDate > cutoffDate) {
                  const vId = item.contentDetails.videoId;
                  if (!existingSet.has(vId)) {
                      candidateIds.push(vId);
                      existingSet.add(vId);
                  } else {
                      stats.duplicate++;
                  }
              } else {
                  stats.tooOld++;
                  keepFetching = false; 
                  break;
              }
          }
          
          pageToken = json.nextPageToken;
          if (!pageToken) keepFetching = false;
      }
      stats.channelsProcessed++;
      updateTaskProgress_(params.progressId, i + 1, channelEntries.length, "Đã quét kênh " + (i + 1) + "/" + channelEntries.length);
    }

    const newVideoLinks = [];
    for (let i = 0; i < candidateIds.length; i += 50) {
      const batchIds = candidateIds.slice(i, i + 50).join(',');
      const json = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics,contentDetails&id=${batchIds}`, apiKey);
      const detailMap = {};
      if (json.items) json.items.forEach(v => { detailMap[v.id] = v; });

      for (const vId of candidateIds.slice(i, i + 50)) {
        stats.videosChecked++;
        const video = detailMap[vId];
        if (!video) { stats.missingDetails++; continue; }

        const views = parseInt((video.statistics && video.statistics.viewCount) || 0);
        if (views <= minViews) { stats.tooFewViews++; continue; }

        const durationSeconds = parseDurationToSeconds_(video.contentDetails && video.contentDetails.duration);
        if (durationSeconds <= minDurationSeconds) { stats.tooShort++; continue; }

        if (!videoMatchesTopic_(video, topicTerms, matchMode)) {
          stats.topicMismatch++;
          continue;
        }

        newVideoLinks.push([`https://www.youtube.com/watch?v=${vId}`]);
      }
    }
    stats.passed = newVideoLinks.length;

    if (newVideoLinks.length > 0) {
       const lastCRow = getLastNonEmptyRowInColumnFast_(videoSheet, 3, 2);
       const startWriteRow = Math.max(lastCRow + 1, 2);

       videoSheet.getRange(startWriteRow, 3, newVideoLinks.length, 1).setValues(newVideoLinks);
       stats.added = newVideoLinks.length;
       recalculateSTT(SHEET_VIDEO);

       let updateMsg = "";
       if (params.autoUpdate) {
         try {
           const endWriteRow = startWriteRow + newVideoLinks.length - 1;
           executeVideoUpdate(videoSheet, startWriteRow, endWriteRow, 'full');
           updateMsg = `\n- Đã cập nhật đầy đủ metadata cho video mới.`;
         } catch (e) {
           updateMsg = `\n- Link đã thêm nhưng cập nhật metadata bị lỗi: ${e.message}`;
         }
       }
       
       const message = `QUÉT HOÀN TẤT (${runLabel})\n\n` +
           `- Điều kiện: mới trong ${daysBack} ngày, duration > ${formatSecondsToMMSS(minDurationSeconds)}, views > ${formatNumber(minViews)}${topicTerms.length ? `, chủ đề: ${params.topicKeywords}` : ""}\n` +
           `- Đã thêm ${newVideoLinks.length} video vào cuối Sheet Video.${updateMsg}\n\n` +
           `Lọc bỏ: ${stats.tooFewViews} views thấp, ${stats.tooShort} quá ngắn, ${stats.topicMismatch} không khớp chủ đề, ${stats.duplicate} trùng, ${stats.missingDetails} thiếu dữ liệu.`;
       setupFetchVideoHistorySheet_().appendRow([
         Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss"),
         selectedRowLabel,
         channelNames.join(", "),
         params.topicKeywords || "",
         daysBack,
         formatSecondsToMMSS(minDurationSeconds),
         minViews,
         `Thêm ${stats.added}/${stats.videosChecked}`,
         message,
         JSON.stringify({ rows: uniqueRows, minViews, daysBack, minDuration: params.minDuration, matchMode })
       ]);
       finishTaskProgress_(params.progressId, true, "Hoàn tất quét: thêm " + stats.added + " video");
       return {
         success: true,
         stats,
         message
       };
    } else {
       const message = `QUÉT HOÀN TẤT (${runLabel}) nhưng không có video đạt yêu cầu.\n\n` +
           `Điều kiện: mới trong ${daysBack} ngày, duration > ${formatSecondsToMMSS(minDurationSeconds)}, views > ${formatNumber(minViews)}${topicTerms.length ? `, chủ đề: ${params.topicKeywords}` : ""}\n` +
           `Lọc bỏ: ${stats.tooFewViews} views thấp, ${stats.tooShort} quá ngắn, ${stats.topicMismatch} không khớp chủ đề, ${stats.duplicate} trùng, ${stats.tooOld} quá cũ, ${stats.missingDetails} thiếu dữ liệu.`;
       setupFetchVideoHistorySheet_().appendRow([
         Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss"),
         selectedRowLabel,
         channelNames.join(", "),
         params.topicKeywords || "",
         daysBack,
         formatSecondsToMMSS(minDurationSeconds),
         minViews,
         `Thêm 0/${stats.videosChecked}`,
         message,
         JSON.stringify({ rows: uniqueRows, minViews, daysBack, minDuration: params.minDuration, matchMode })
       ]);
       finishTaskProgress_(params.progressId, true, "Hoàn tất quét: không có video mới");
       return {
         success: false,
         stats,
         message
       };
    }

  } catch (error) {
    finishTaskProgress_(params && params.progressId, false, "Lỗi: " + error.message);
    return { success: false, message: `HỆ THỐNG GẶP LỖI:\n${error.message}`, stats };
  }
}

function getVideoTabPlaylistId(urlStr, apiKey) {
  urlStr = urlStr.toString().trim().replace(/\/$/, ""); 
  let endpoints = [];
  if (urlStr.length === 24 && urlStr.startsWith("UC")) endpoints.push(`id=${urlStr}`);
  else if (urlStr.startsWith("@")) endpoints.push(`forHandle=${urlStr.substring(1)}`);
  else if (urlStr.includes("/channel/UC")) { const m = urlStr.match(/\/channel\/(UC[\w-]{22})/); if (m) endpoints.push(`id=${m[1]}`); }
  else if (urlStr.includes("@")) { const m = urlStr.match(/@([\w.-]+)/); if (m) endpoints.push(`forHandle=${m[1]}`); }
  else if (urlStr.includes("/user/")) { const m = urlStr.match(/\/user\/([\w.-]+)/); if (m) endpoints.push(`forUsername=${m[1]}`); }
  else if (urlStr.includes("/c/")) { const m = urlStr.match(/\/c\/([\w.-]+)/); if (m) { endpoints.push(`forHandle=@${m[1]}`); endpoints.push(`forUsername=${m[1]}`); } }
  else { const m = urlStr.match(/youtube\.com\/([^/?&#]+)/); if (m && !['watch', 'shorts', 'playlist', 'results', 'live', 'feed'].includes(m[1])) { endpoints.push(`forHandle=@${m[1]}`); endpoints.push(`forHandle=${m[1]}`); endpoints.push(`forUsername=${m[1]}`); } }
  
  if (endpoints.length === 0) return null;

  for (let ep of endpoints) {
    try {
      let tempJson = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/channels?part=id&${ep}`, apiKey);
      if (tempJson && tempJson.items && tempJson.items.length > 0) { 
          let channelId = tempJson.items[0].id;
          if (channelId && channelId.startsWith("UC")) {
              return "UULF" + channelId.substring(2);
          }
      }
    } catch (e) {
      continue;
    }
  }
  return null;
}

