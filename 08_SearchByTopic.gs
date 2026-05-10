// =========================================================================
// TÌM VIDEO/KÊNH THEO CHỦ ĐỀ
// =========================================================================

/**
 * 🔍 Mở dialog HTML để user nhập chủ đề + tham số tìm kiếm
 */
function openSearchByTopicDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 18px; line-height: 1.45; color:#202124; background:#f8fafc;">
      ${buildDialogHeader_("TÌM VIDEO YOUTUBE THEO CHỦ ĐỀ", "Tìm video theo dữ liệu YouTube API, lọc theo điều kiện bên dưới và nối thêm vào cuối Sheet VIDEO.", "#1a73e8")}
      
      <table style="width:100%; border-collapse: collapse; background:#ffffff; border:1px solid #dfe3eb; border-radius:8px; overflow:hidden;">
        <tr><td colspan="2" style="padding:8px 0;"><b>📌 CHỦ ĐỀ TÌM KIẾM</b> <span style="color:red;">*</span></td></tr>
        <tr><td colspan="2"><input id="topic" type="text" placeholder="VD: dark psychology, true crime stories, history documentary..." 
          style="width:100%; padding:8px; box-sizing:border-box; border:2px solid #4285F4; border-radius:4px;"></td></tr>
        
        <tr><td style="width:50%; padding:10px 5px 5px 0;">
          <b>👁️ Min views:</b><br>
          <input id="minViews" type="number" value="100000" style="width:95%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td><td style="padding:10px 0 5px 5px;">
          <b>⏱️ Min duration (phút):</b><br>
          <input id="minDurationMin" type="number" value="7" style="width:95%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td></tr>
        
        <tr><td style="padding:10px 5px 5px 0;">
          <b>📅 Trong vòng (ngày):</b><br>
          <input id="daysAgo" type="number" value="90" style="width:95%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td><td style="padding:10px 0 5px 5px;">
          <b>📊 Sắp xếp theo:</b><br>
          <select id="orderBy" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            <option value="viewCount" selected>Phổ biến nhất (Views)</option>
            <option value="relevance">Liên quan nhất</option>
            <option value="date">Mới nhất</option>
            <option value="rating">Đánh giá cao nhất</option>
          </select>
        </td></tr>
        
        <tr><td colspan="2" style="padding:10px 0 5px;">
          <b>🌍 Quốc gia ƯU TIÊN tìm kiếm</b> <span style="font-size:11px; color:#666;">(mã ISO 2 chữ, cách nhau bằng dấu phẩy):</span><br>
          <input id="includeCountries" type="text" value="US,GB" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;" placeholder="VD: US,GB,CA">
        </td></tr>
        
        <tr><td colspan="2" style="padding:5px 0;">
          <b>🚫 Quốc gia LOẠI BỎ</b> <span style="font-size:11px; color:#666;">(loại video từ kênh có quốc gia này):</span><br>
          <input id="excludeCountries" type="text" value="IN" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;" placeholder="VD: IN,PK">
        </td></tr>
        
        <tr><td style="padding:10px 5px 5px 0;">
          <b>🌐 Ngôn ngữ ưu tiên:</b><br>
          <select id="relevanceLanguage" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            <option value="en" selected>Tiếng Anh (en)</option>
            <option value="">Bất kỳ</option>
            <option value="vi">Tiếng Việt</option>
            <option value="es">Tiếng Tây Ban Nha</option>
            <option value="fr">Tiếng Pháp</option>
            <option value="ja">Tiếng Nhật</option>
            <option value="ko">Tiếng Hàn</option>
          </select>
        </td><td style="padding:10px 0 5px 5px;">
          <b>📦 Số kết quả/quốc gia:</b><br>
          <input id="maxResults" type="number" value="50" min="5" max="50" style="width:95%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td></tr>
        
        <tr><td colspan="2" style="padding:10px 0 5px;">
          <input type="checkbox" id="excludeShorts" checked> <label for="excludeShorts">Loại bỏ Shorts (< 60s)</label> &nbsp;&nbsp;
          <input type="checkbox" id="autoUpdate" checked> <label for="autoUpdate">Tự động fill metadata sau khi tìm xong</label>
        </td></tr>
      </table>
      
      <div style="margin-top:20px; text-align:center;">
        <button onclick="runSearch()" id="btnSearch" style="padding:12px 30px; background:#4285F4; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">🚀 BẮT ĐẦU TÌM KIẾM</button>
        <button onclick="google.script.host.close()" style="padding:12px 20px; background:#9aa0a6; color:white; border:none; border-radius:6px; cursor:pointer; margin-left:10px;">❌ HỦY</button>
      </div>
      
      <div id="status" style="margin-top:15px; padding:12px; background:#ffffff; border:1px solid #dfe3eb; border-radius:8px; min-height:34px; font-size:13px;"></div>
      
      <p style="margin-top:15px; padding:8px; background:#fff3cd; border-left:4px solid #f29900; font-size:12px;">
        <b>⚠️ Lưu ý quota:</b> Mỗi search tốn 100 quota units, mỗi quốc gia 1 search. Quota mặc định YouTube là 10,000/ngày.<br>
        Nếu tìm 2 nước (US+GB) → tốn ~210 units cho lần tìm.
      </p>
      
      <script>
        function fmt(n) {
          return (Number(n) || 0).toLocaleString('vi-VN');
        }
        function renderResult(res) {
          const s = res.stats || {};
          const cards = [
            ['Tìm được', fmt(s.found || 0)],
            ['Chưa có trong Sheet', fmt(s.newCandidates || 0)],
            ['Đạt bộ lọc', fmt(s.passed || 0)],
            ['Đã thêm', fmt(s.added || 0)]
          ].map(c => '<div style="display:inline-block; width:23%; min-width:110px; box-sizing:border-box; margin:4px; padding:10px; border:1px solid #dfe3eb; border-radius:8px; background:#f8fafc; text-align:center;"><div style="font-size:11px; color:#5f6368;">' + c[0] + '</div><div style="font-size:20px; color:#1a73e8; font-weight:bold;">' + c[1] + '</div></div>').join('');
          return '<div>' + cards + '</div><div style="white-space:pre-line; margin-top:8px;">' + res.message + '</div>';
        }
        function runSearch() {
          const topic = document.getElementById('topic').value.trim();
          if (!topic) { alert('Vui lòng nhập chủ đề!'); return; }
          
          const params = {
            topic: topic,
            minViews: parseInt(document.getElementById('minViews').value) || 0,
            minDurationMin: parseInt(document.getElementById('minDurationMin').value) || 0,
            daysAgo: parseInt(document.getElementById('daysAgo').value) || 90,
            orderBy: document.getElementById('orderBy').value,
            includeCountries: document.getElementById('includeCountries').value.split(',').map(s => s.trim().toUpperCase()).filter(s => s),
            excludeCountries: document.getElementById('excludeCountries').value.split(',').map(s => s.trim().toUpperCase()).filter(s => s),
            relevanceLanguage: document.getElementById('relevanceLanguage').value,
            maxResults: Math.min(50, Math.max(5, parseInt(document.getElementById('maxResults').value) || 50)),
            excludeShorts: document.getElementById('excludeShorts').checked,
            autoUpdate: document.getElementById('autoUpdate').checked
          };
          
          document.getElementById('btnSearch').disabled = true;
          document.getElementById('btnSearch').innerHTML = '⏳ Đang tìm kiếm...';
          document.getElementById('status').innerHTML = '🔄 Đang gọi YouTube API, vui lòng đợi...';
          
          google.script.run
            .withSuccessHandler(function(res) {
              if (res.success) {
                document.getElementById('status').innerHTML = '✅ ' + renderResult(res);
                document.getElementById('btnSearch').innerHTML = '✅ HOÀN TẤT';
              } else {
                document.getElementById('status').innerHTML = '❌ ' + renderResult(res);
                document.getElementById('btnSearch').disabled = false;
                document.getElementById('btnSearch').innerHTML = '🔄 THỬ LẠI';
              }
            })
            .withFailureHandler(function(err) {
              document.getElementById('status').innerHTML = '❌ Lỗi: ' + err.message;
              document.getElementById('btnSearch').disabled = false;
              document.getElementById('btnSearch').innerHTML = '🔄 THỬ LẠI';
            })
            .executeSearchByTopic(params);
        }
      </script>
    </div>
  `).setWidth(680).setHeight(760);
  
  ui.showModalDialog(html, '🔍 Tìm Video Theo Chủ Đề');
}

/**
 * 🎯 LÕI XỬ LÝ TÌM KIẾM - được gọi từ HTML dialog
 * @param {object} params - Tham số từ form
 * @return {object} { success, message }
 */
function executeSearchByTopic(params) {
  try {
    const apiKey = getApiKey();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const videoSheet = ss.getSheetByName(SHEET_VIDEO);
    if (!videoSheet) return { success: false, message: "Không tìm thấy Sheet Video" };
    const resultStats = { found: 0, newCandidates: 0, passed: 0, added: 0 };
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - (params.daysAgo || 90));
    const publishedAfter = cutoffDate.toISOString();
    
    const existingIds = new Set();
    const lastRow = videoSheet.getLastRow();
    if (lastRow >= 2) {
      const links = videoSheet.getRange(2, 3, lastRow - 1, 1).getValues();
      links.forEach(r => { 
        const id = extractVideoId(r[0]); 
        if (id) existingIds.add(id);
      });
    }
    
    const includeCountries = (params.includeCountries && params.includeCountries.length > 0) 
      ? params.includeCountries : ["US"];
    const progressId = params.progressId;
    let progressDone = 0;
    let progressTotal = Math.max(1, includeCountries.length);
    initTaskProgress_(progressId, progressTotal, "Đang tìm video theo quốc gia 0/" + includeCountries.length);
    
    const allCandidateIds = new Set();
    let totalFromSearch = 0;
    
    const orderMap = {
      "viewCount": "viewCount",
      "relevance": "relevance",
      "date": "date",
      "rating": "rating"
    };
    const order = orderMap[params.orderBy] || "viewCount";
    
    let videoDurationParam = "any";
    if (params.minDurationMin >= 20) videoDurationParam = "long";
    else if (params.minDurationMin >= 4) videoDurationParam = "medium";
    
    for (const region of includeCountries) {
      try {
        let url = `https://www.googleapis.com/youtube/v3/search?part=snippet&type=video`;
        url += `&q=${encodeURIComponent(params.topic)}`;
        url += `&maxResults=${params.maxResults}`;
        url += `&order=${order}`;
        url += `&publishedAfter=${publishedAfter}`;
        url += `&regionCode=${region}`;
        url += `&videoDuration=${videoDurationParam}`;
        if (params.relevanceLanguage) {
          url += `&relevanceLanguage=${params.relevanceLanguage}`;
        }
        
        const json = fetchAPIWithRetry(url, apiKey);
        if (json.items) {
          json.items.forEach(item => {
            if (item.id && item.id.videoId) {
              allCandidateIds.add(item.id.videoId);
              totalFromSearch++;
            }
          });
        }
      } catch (e) {
        console.warn(`Lỗi search region ${region}: ${e.message}`);
      }
      progressDone++;
      updateTaskProgress_(progressId, progressDone, progressTotal, "Đã tìm video " + progressDone + "/" + includeCountries.length + " quốc gia");
    }
    
    if (allCandidateIds.size === 0) {
      return { success: false, message: `Không tìm thấy video nào với chủ đề "${params.topic}". Thử thay từ khóa khác.`, stats: resultStats };
    }
    resultStats.found = allCandidateIds.size;
    
    const newCandidateIds = [...allCandidateIds].filter(id => !existingIds.has(id));
    resultStats.newCandidates = newCandidateIds.length;
    if (newCandidateIds.length === 0) {
      return { success: false, message: `Tìm được ${allCandidateIds.size} video nhưng tất cả đã có sẵn trong Sheet.`, stats: resultStats };
    }
    
    const videoDetails = {};
    const detailBatches = Math.ceil(newCandidateIds.length / 50);
    progressTotal = includeCountries.length + detailBatches + 1;
    for (let i = 0; i < newCandidateIds.length; i += 50) {
      const batch = newCandidateIds.slice(i, i + 50).join(',');
      const url = `https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics,contentDetails&id=${batch}`;
      const json = fetchAPIWithRetry(url, apiKey);
      if (json.items) {
        json.items.forEach(v => { videoDetails[v.id] = v; });
      }
      progressDone++;
      updateTaskProgress_(progressId, progressDone, progressTotal, "Đã tải metadata video batch " + Math.min(detailBatches, Math.ceil(i / 50) + 1) + "/" + detailBatches);
    }
    
    const channelIds = new Set();
    Object.values(videoDetails).forEach(v => {
      if (v.snippet && v.snippet.channelId) channelIds.add(v.snippet.channelId);
    });
    
    const channelCountryMap = {};
    const channelIdArr = [...channelIds];
    for (let i = 0; i < channelIdArr.length; i += 50) {
      const batch = channelIdArr.slice(i, i + 50).join(',');
      const url = `https://www.googleapis.com/youtube/v3/channels?part=snippet&id=${batch}`;
      try {
        const json = fetchAPIWithRetry(url, apiKey);
        if (json.items) {
          json.items.forEach(c => {
            channelCountryMap[c.id] = (c.snippet && c.snippet.country) ? c.snippet.country.toUpperCase() : '';
          });
        }
      } catch(e) {}
    }
    
    const minDurationSec = (params.minDurationMin || 0) * 60;
    const minViews = params.minViews || 0;
    const excludeCountries = (params.excludeCountries || []).map(c => c.toUpperCase());
    
    const passedVideoIds = [];
    let filterStats = { 
      tooFewViews: 0, 
      tooShort: 0, 
      isShort: 0,
      excludedCountry: 0, 
      isLive: 0, 
      missing: 0
    };
    
    for (const vId of newCandidateIds) {
      const v = videoDetails[vId];
      if (!v) { filterStats.missing++; continue; }
      
      if (v.snippet.liveBroadcastContent && v.snippet.liveBroadcastContent !== "none") {
        filterStats.isLive++;
        continue;
      }
      
      const views = parseInt(v.statistics.viewCount || 0);
      if (views < minViews) { filterStats.tooFewViews++; continue; }
      
      const durationStr = v.contentDetails ? v.contentDetails.duration : "";
      const dMatch = durationStr.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
      let durationSec = 0;
      if (dMatch) {
        durationSec = parseInt(dMatch[1]||0)*3600 + parseInt(dMatch[2]||0)*60 + parseInt(dMatch[3]||0);
      }
      if (durationSec < minDurationSec) { filterStats.tooShort++; continue; }
      
      if (params.excludeShorts && durationSec < 60) {
        filterStats.isShort++;
        continue;
      }
      
      const channelCountry = channelCountryMap[v.snippet.channelId] || '';
      if (channelCountry && excludeCountries.includes(channelCountry)) {
        filterStats.excludedCountry++;
        continue;
      }
      
      passedVideoIds.push(vId);
    }
    resultStats.passed = passedVideoIds.length;
    Object.assign(resultStats, filterStats);
    progressDone++;
    updateTaskProgress_(progressId, progressDone, progressTotal, "Đã lọc xong " + newCandidateIds.length + " video ứng viên");
    
    if (passedVideoIds.length === 0) {
      let msg = `Tìm được ${allCandidateIds.size} video, lọc xong KHÔNG có video nào đạt yêu cầu.\n\n`;
      msg += `Chi tiết loại bỏ:\n`;
      msg += `• ${filterStats.tooFewViews} video views < ${minViews}\n`;
      msg += `• ${filterStats.tooShort} video duration < ${params.minDurationMin}p\n`;
      if (filterStats.isShort > 0) msg += `• ${filterStats.isShort} video Shorts\n`;
      if (filterStats.excludedCountry > 0) msg += `• ${filterStats.excludedCountry} video từ kênh excluded\n`;
      if (filterStats.isLive > 0) msg += `• ${filterStats.isLive} video live\n`;
      return { success: false, message: msg, stats: resultStats };
    }
    
    const cValues = videoSheet.getRange("C:C").getValues();
    let lastCRow = 0;
    for (let i = cValues.length - 1; i >= 0; i--) {
      if (cValues[i][0] && cValues[i][0].toString().trim() !== "") {
        lastCRow = i + 1;
        break;
      }
    }
    const startWriteRow = Math.max(lastCRow + 1, 2);
    
    const newLinks = passedVideoIds.map(id => [`https://www.youtube.com/watch?v=${id}`]);
    videoSheet.getRange(startWriteRow, 3, newLinks.length, 1).setValues(newLinks);
    resultStats.added = newLinks.length;
    Object.assign(resultStats, filterStats);
    recalculateSTT(SHEET_VIDEO);
    
    let updateMsg = "";
    if (params.autoUpdate) {
      try {
        const endWriteRow = startWriteRow + newLinks.length - 1;
        executeVideoUpdate(videoSheet, startWriteRow, endWriteRow, 'full');
        updateMsg = ` Đã cập nhật đầy đủ metadata.`;
      } catch (e) {
        updateMsg = ` Tuy nhiên khi cập nhật metadata bị lỗi: ${e.message}. Bạn có thể chạy "Cập nhật theo khoảng dòng" thủ công.`;
      }
    }
    finishTaskProgress_(progressId, true, "Hoàn tất tìm video: thêm " + newLinks.length + " link");
    
    return {
      success: true,
      stats: resultStats,
      message: 
        `Tìm thấy ${allCandidateIds.size} video → Lọc còn ${passedVideoIds.length} video phù hợp.\n` +
        `📥 Đã thêm ${newLinks.length} link mới vào dòng ${startWriteRow}-${startWriteRow + newLinks.length - 1}.${updateMsg}\n\n` +
        `Lọc bỏ: ${filterStats.tooFewViews} views thấp, ${filterStats.tooShort} quá ngắn, ${filterStats.excludedCountry} bị exclude.`
    };
    
  } catch (e) {
    finishTaskProgress_(params && params.progressId, false, "Lỗi: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 🔍 Mở dialog tìm KÊNH theo chủ đề (tương tự nhưng cho channel)
 */
function openSearchChannelsByTopicDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 18px; line-height: 1.45; color:#202124; background:#f8fafc;">
      ${buildDialogHeader_("TÌM KÊNH YOUTUBE THEO CHỦ ĐỀ", "Tìm kênh theo YouTube API, lọc subscriber, quốc gia và views/tháng trước khi nối vào Sheet KÊNH.", "#0f9d58")}
      
      <table style="width:100%; border-collapse: collapse; background:#ffffff; border:1px solid #dfe3eb; border-radius:8px; overflow:hidden;">
        <tr><td colspan="2" style="padding:8px 0;"><b>📌 CHỦ ĐỀ TÌM KIẾM</b> <span style="color:red;">*</span></td></tr>
        <tr><td colspan="2"><input id="topic" type="text" placeholder="VD: dark psychology, history channel..." 
          style="width:100%; padding:8px; box-sizing:border-box; border:2px solid #0f9d58; border-radius:4px;"></td></tr>
        
        <tr><td style="width:50%; padding:10px 5px 5px 0;">
          <b>👥 Min subscribers:</b><br>
          <input id="minSubs" type="number" value="10000" style="width:95%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td><td style="padding:10px 0 5px 5px;">
          <b>👁️ Min views/tháng:</b><br>
          <input id="minMonthlyViews" type="number" value="500000" style="width:95%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td></tr>

        <tr><td style="padding:10px 5px 5px 0;">
          <b>📊 Sắp xếp theo:</b><br>
          <select id="orderBy" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            <option value="viewCount" selected>Phổ biến (Views)</option>
            <option value="relevance">Liên quan</option>
          </select>
        </td><td style="padding:10px 0 5px 5px;">
          <b>📦 Số kết quả/quốc gia:</b><br>
          <input id="maxResults" type="number" value="50" min="5" max="50" style="width:95%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td></tr>
        
        <tr><td colspan="2" style="padding:10px 0 5px;">
          <b>🌍 Quốc gia ƯU TIÊN:</b><br>
          <input id="includeCountries" type="text" value="US,GB" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;" placeholder="VD: US,GB,CA">
        </td></tr>
        
        <tr><td colspan="2" style="padding:5px 0;">
          <b>🚫 Quốc gia LOẠI BỎ:</b><br>
          <input id="excludeCountries" type="text" value="IN" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
        </td></tr>
        
        <tr><td colspan="2" style="padding:10px 0 5px;">
          <b>🌐 Ngôn ngữ:</b><br>
          <select id="relevanceLanguage" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            <option value="en" selected>Tiếng Anh</option>
            <option value="">Bất kỳ</option>
            <option value="vi">Tiếng Việt</option>
          </select>
        </td></tr>
        
        <tr><td colspan="2" style="padding:10px 0 5px;">
          <input type="checkbox" id="autoUpdate" checked> <label for="autoUpdate">Tự động fill metadata sau khi tìm xong</label>
        </td></tr>
      </table>
      
      <div style="margin-top:20px; text-align:center;">
        <button onclick="runSearch()" id="btnSearch" style="padding:12px 30px; background:#0f9d58; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px;">🚀 BẮT ĐẦU TÌM KÊNH</button>
        <button onclick="google.script.host.close()" style="padding:12px 20px; background:#9aa0a6; color:white; border:none; border-radius:6px; cursor:pointer; margin-left:10px;">❌ HỦY</button>
      </div>
      
      <div id="status" style="margin-top:15px; padding:12px; background:#ffffff; border:1px solid #dfe3eb; border-radius:8px; min-height:34px; font-size:13px;"></div>
      
      <script>
        function fmt(n) {
          return (Number(n) || 0).toLocaleString('vi-VN');
        }
        function renderResult(res) {
          const s = res.stats || {};
          const cards = [
            ['Tìm được', fmt(s.found || 0)],
            ['Chưa có trong Sheet', fmt(s.newCandidates || 0)],
            ['Đạt bộ lọc', fmt(s.passed || 0)],
            ['Đã thêm', fmt(s.added || 0)]
          ].map(c => '<div style="display:inline-block; width:23%; min-width:110px; box-sizing:border-box; margin:4px; padding:10px; border:1px solid #dfe3eb; border-radius:8px; background:#f8fafc; text-align:center;"><div style="font-size:11px; color:#5f6368;">' + c[0] + '</div><div style="font-size:20px; color:#0f9d58; font-weight:bold;">' + c[1] + '</div></div>').join('');
          return '<div>' + cards + '</div><div style="white-space:pre-line; margin-top:8px;">' + res.message + '</div>';
        }
        function runSearch() {
          const topic = document.getElementById('topic').value.trim();
          if (!topic) { alert('Vui lòng nhập chủ đề!'); return; }
          
          const params = {
            topic: topic,
            minSubs: parseInt(document.getElementById('minSubs').value) || 0,
            minMonthlyViews: parseInt(document.getElementById('minMonthlyViews').value) || 0,
            orderBy: document.getElementById('orderBy').value,
            includeCountries: document.getElementById('includeCountries').value.split(',').map(s => s.trim().toUpperCase()).filter(s => s),
            excludeCountries: document.getElementById('excludeCountries').value.split(',').map(s => s.trim().toUpperCase()).filter(s => s),
            relevanceLanguage: document.getElementById('relevanceLanguage').value,
            maxResults: Math.min(50, Math.max(5, parseInt(document.getElementById('maxResults').value) || 50)),
            autoUpdate: document.getElementById('autoUpdate').checked
          };
          
          document.getElementById('btnSearch').disabled = true;
          document.getElementById('btnSearch').innerHTML = '⏳ Đang tìm kiếm...';
          document.getElementById('status').innerHTML = '🔄 Đang gọi YouTube API, vui lòng đợi...';
          
          google.script.run
            .withSuccessHandler(function(res) {
              if (res.success) {
                document.getElementById('status').innerHTML = '✅ ' + renderResult(res);
                document.getElementById('btnSearch').innerHTML = '✅ HOÀN TẤT';
              } else {
                document.getElementById('status').innerHTML = '❌ ' + renderResult(res);
                document.getElementById('btnSearch').disabled = false;
                document.getElementById('btnSearch').innerHTML = '🔄 THỬ LẠI';
              }
            })
            .withFailureHandler(function(err) {
              document.getElementById('status').innerHTML = '❌ Lỗi: ' + err.message;
              document.getElementById('btnSearch').disabled = false;
              document.getElementById('btnSearch').innerHTML = '🔄 THỬ LẠI';
            })
            .executeSearchChannelsByTopic(params);
        }
      </script>
    </div>
  `).setWidth(680).setHeight(700);
  
  ui.showModalDialog(html, '📺 Tìm Kênh Theo Chủ Đề');
}

/**
 * Ước tính views/tháng bằng cách cộng viewCount của video đăng trong 28 ngày gần nhất.
 * Dữ liệu lấy trực tiếp từ YouTube Data API public statistics, không dùng số tự nhập.
 */
function estimateChannelMonthlyViewsFromUploads_(channel, apiKey) {
  const playlistId = channel && channel.contentDetails && channel.contentDetails.relatedPlaylists
    ? channel.contentDetails.relatedPlaylists.uploads
    : null;
  if (!playlistId) return { views: 0, videos: 0 };

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 28);

  const videoIds = [];
  let pageToken = "";
  let keepFetching = true;
  let pageCount = 0;

  while (keepFetching && pageCount < 10) {
    let url = `https://www.googleapis.com/youtube/v3/playlistItems?part=snippet,contentDetails&playlistId=${playlistId}&maxResults=50`;
    if (pageToken) url += `&pageToken=${pageToken}`;
    const json = fetchAPIWithRetry(url, apiKey);
    pageCount++;

    if (!json.items || json.items.length === 0) break;

    for (const item of json.items) {
      const publishedRaw = (item.contentDetails && item.contentDetails.videoPublishedAt) || (item.snippet && item.snippet.publishedAt);
      const publishedAt = publishedRaw ? new Date(publishedRaw) : null;
      if (publishedAt && publishedAt < cutoff) {
        keepFetching = false;
        break;
      }
      if (item.contentDetails && item.contentDetails.videoId) {
        videoIds.push(item.contentDetails.videoId);
      }
    }

    pageToken = json.nextPageToken || "";
    if (!pageToken) keepFetching = false;
  }

  let views = 0;
  for (let i = 0; i < videoIds.length; i += 50) {
    const batch = videoIds.slice(i, i + 50).join(',');
    const json = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/videos?part=statistics&id=${batch}`, apiKey);
    if (json.items) {
      json.items.forEach(v => {
        views += parseInt((v.statistics && v.statistics.viewCount) || 0);
      });
    }
  }

  return { views, videos: videoIds.length };
}

/**
 * 🎯 LÕI XỬ LÝ TÌM KÊNH THEO CHỦ ĐỀ
 */
function executeSearchChannelsByTopic(params) {
  try {
    const apiKey = getApiKey();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const channelSheet = ss.getSheetByName(SHEET_CHANNEL);
    if (!channelSheet) return { success: false, message: "Không tìm thấy Sheet Kênh" };
    const resultStats = { found: 0, newCandidates: 0, passed: 0, added: 0 };
    
    const existingChannels = new Set();
    const lastRow = channelSheet.getLastRow();
    if (lastRow >= 2) {
      const links = channelSheet.getRange(2, 3, lastRow - 1, 1).getValues();
      for (const r of links) {
        const u = r[0] ? r[0].toString().trim() : "";
        if (!u) continue;
        const m = u.match(/\/channel\/(UC[\w-]{22})/);
        if (m) existingChannels.add(m[1].toLowerCase());
        else existingChannels.add(u.toLowerCase());
      }
    }
    
    const includeCountries = (params.includeCountries && params.includeCountries.length > 0) 
      ? params.includeCountries : ["US"];
    const progressId = params.progressId;
    let progressDone = 0;
    let progressTotal = Math.max(1, includeCountries.length);
    initTaskProgress_(progressId, progressTotal, "Đang tìm kênh theo quốc gia 0/" + includeCountries.length);
    
    const allCandidateChannelIds = new Set();
    
    for (const region of includeCountries) {
      try {
        let url = `https://www.googleapis.com/youtube/v3/search?part=snippet&type=channel`;
        url += `&q=${encodeURIComponent(params.topic)}`;
        url += `&maxResults=${params.maxResults}`;
        url += `&order=${params.orderBy || 'relevance'}`;
        url += `&regionCode=${region}`;
        if (params.relevanceLanguage) {
          url += `&relevanceLanguage=${params.relevanceLanguage}`;
        }
        
        const json = fetchAPIWithRetry(url, apiKey);
        if (json.items) {
          json.items.forEach(item => {
            if (item.id && item.id.channelId) {
              allCandidateChannelIds.add(item.id.channelId);
            }
          });
        }
      } catch (e) {
        console.warn(`Lỗi search kênh region ${region}: ${e.message}`);
      }
      progressDone++;
      updateTaskProgress_(progressId, progressDone, progressTotal, "Đã tìm kênh " + progressDone + "/" + includeCountries.length + " quốc gia");
    }
    
    if (allCandidateChannelIds.size === 0) {
      return { success: false, message: `Không tìm thấy kênh nào với chủ đề "${params.topic}".`, stats: resultStats };
    }
    resultStats.found = allCandidateChannelIds.size;
    
    const newChannelIds = [...allCandidateChannelIds].filter(id => !existingChannels.has(id.toLowerCase()));
    resultStats.newCandidates = newChannelIds.length;
    if (newChannelIds.length === 0) {
      return { success: false, message: `Tìm thấy ${allCandidateChannelIds.size} kênh nhưng tất cả đã có sẵn.`, stats: resultStats };
    }
    
    const channelDetails = {};
    const detailBatches = Math.ceil(newChannelIds.length / 50);
    progressTotal = includeCountries.length + detailBatches + newChannelIds.length + 1;
    for (let i = 0; i < newChannelIds.length; i += 50) {
      const batch = newChannelIds.slice(i, i + 50).join(',');
      const url = `https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics,contentDetails&id=${batch}`;
      try {
        const json = fetchAPIWithRetry(url, apiKey);
        if (json.items) json.items.forEach(c => { channelDetails[c.id] = c; });
      } catch(e) {}
      progressDone++;
      updateTaskProgress_(progressId, progressDone, progressTotal, "Đã tải metadata kênh batch " + Math.min(detailBatches, Math.ceil(i / 50) + 1) + "/" + detailBatches);
    }
    
    const minSubs = params.minSubs || 0;
    const minMonthlyViews = params.minMonthlyViews === undefined ? 500000 : (params.minMonthlyViews || 0);
    const excludeCountries = (params.excludeCountries || []).map(c => c.toUpperCase());
    const passedChannels = [];
    let filterStats = { tooFewSubs: 0, tooFewMonthlyViews: 0, excludedCountry: 0, missing: 0, hidden: 0, monthlyViewErrors: 0 };
    
    for (const cId of newChannelIds) {
      const c = channelDetails[cId];
      if (!c) { filterStats.missing++; continue; }
      
      if (c.statistics.hiddenSubscriberCount === true) {
      } else {
        const subs = parseInt(c.statistics.subscriberCount || 0);
        if (subs < minSubs) { filterStats.tooFewSubs++; continue; }
      }
      
      const country = (c.snippet && c.snippet.country) ? c.snippet.country.toUpperCase() : '';
      if (country && excludeCountries.includes(country)) {
        filterStats.excludedCountry++;
        continue;
      }

      if (minMonthlyViews > 0) {
        try {
          const monthly = estimateChannelMonthlyViewsFromUploads_(c, apiKey);
          if (monthly.views < minMonthlyViews) {
            filterStats.tooFewMonthlyViews++;
            continue;
          }
        } catch (e) {
          filterStats.monthlyViewErrors++;
          continue;
        }
      }
      
      passedChannels.push(cId);
      progressDone++;
      updateTaskProgress_(progressId, progressDone, progressTotal, "Đã lọc kênh " + progressDone + "/" + progressTotal);
    }
    resultStats.passed = passedChannels.length;
    updateTaskProgress_(progressId, progressTotal - 1, progressTotal, "Đã lọc xong " + newChannelIds.length + " kênh ứng viên");
    
    if (passedChannels.length === 0) {
      Object.assign(resultStats, filterStats);
      return { 
        success: false, 
        message: `Tìm ${allCandidateChannelIds.size} kênh, KHÔNG có kênh đạt yêu cầu.\n` +
                 `Lọc bỏ: ${filterStats.tooFewSubs} subs thấp, ${filterStats.tooFewMonthlyViews} views/tháng thấp, ${filterStats.excludedCountry} bị exclude, ${filterStats.monthlyViewErrors} lỗi tính views/tháng.`,
        stats: resultStats
      };
    }
    
    const lastCRow = getLastNonEmptyRowInColumnFast_(channelSheet, 3, 2);
    const startWriteRow = Math.max(lastCRow + 1, 2);
    
    const newLinks = passedChannels.map(id => [`https://www.youtube.com/channel/${id}`]);
    channelSheet.getRange(startWriteRow, 3, newLinks.length, 1).setValues(newLinks);
    resultStats.added = newLinks.length;
    Object.assign(resultStats, filterStats);
    recalculateSTT(SHEET_CHANNEL);
    
    let updateMsg = "";
    if (params.autoUpdate) {
      try {
        const endWriteRow = startWriteRow + newLinks.length - 1;
        executeChannelUpdate(channelSheet, startWriteRow, endWriteRow, 'full');
        updateMsg = ` Đã cập nhật đầy đủ metadata.`;
      } catch (e) {
        updateMsg = ` Lỗi khi cập nhật: ${e.message}.`;
      }
    }
    finishTaskProgress_(progressId, true, "Hoàn tất tìm kênh: thêm " + newLinks.length + " link");
    
    return {
      success: true,
      stats: resultStats,
      message: 
        `Tìm ${allCandidateChannelIds.size} kênh → Lọc còn ${passedChannels.length} kênh.\n` +
        `📥 Đã thêm vào dòng ${startWriteRow}-${startWriteRow + newLinks.length - 1}.${updateMsg}\n\n` +
        `Lọc bỏ: ${filterStats.tooFewSubs} subs thấp, ${filterStats.tooFewMonthlyViews} views/tháng thấp, ${filterStats.excludedCountry} bị exclude.`
    };
    
  } catch (e) {
    finishTaskProgress_(params && params.progressId, false, "Lỗi: " + e.message);
    return { success: false, message: e.message };
  }
}


