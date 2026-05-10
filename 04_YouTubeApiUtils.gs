function showDailyQuotaReminder_(ui) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const today = getQuotaDateKey_();
    if (props.getProperty('YT_QUOTA_REMINDER_DATE') === today) return;
    props.setProperty('YT_QUOTA_REMINDER_DATE', today);
    ui.alert("📊 Nhắc quota YouTube API hằng ngày", getYouTubeQuotaUsageMessage_(), ui.ButtonSet.OK);
  } catch (e) {}
}

function fetchAPIWithRetry(url, apiKey) {
  const cache = CacheService.getScriptCache();
  const cacheKey = url.replace(/[^a-zA-Z0-9_-]/g, '').substring(0, 200);
  
  const cachedData = cache.get(cacheKey);
  if (cachedData) return JSON.parse(cachedData);

  const fullUrl = url + (url.includes('?') ? '&' : '?') + 'key=' + apiKey;
  let lastErrorMsg = "";
  
  for (let i = 0; i < 4; i++) { 
    try {
      recordYouTubeQuota_(fullUrl);
      const response = UrlFetchApp.fetch(fullUrl, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      
      if (json.error) {
        const errReason = (json.error.errors && json.error.errors.length > 0) ? json.error.errors[0].reason : '';
        lastErrorMsg = json.error.message;
        
        if (errReason === 'quotaExceeded') {
           throw new Error(`[QUOTA EXCEEDED] Bạn đã dùng hết Quota API trong ngày hôm nay.\n${getYouTubeQuotaUsageMessage_()}`);
        }
        
        if (errReason === 'rateLimitExceeded' || errReason === 'userRateLimitExceeded' || response.getResponseCode() === 429) {
          Utilities.sleep((i + 1) * 2000); 
          continue; 
        }
        
        throw new Error(`[API ERROR] ${json.error.message}`);
      }
      
      if (json.items) {
        try { cache.put(cacheKey, response.getContentText(), 1800); } catch(e) {}
      }
      
      return json; 
    } catch (e) {
      if (e.message.includes("[QUOTA EXCEEDED]") || e.message.includes("[API ERROR]")) {
        throw e; 
      }
      if (i === 3) throw new Error(`Không thể kết nối máy chủ Google: ${lastErrorMsg || e.message}`);
      Utilities.sleep((i + 1) * 1500);
    }
  }
}

function fetchAPIBatchWithRetry_(urls, apiKey) {
  const cache = CacheService.getScriptCache();
  const results = new Array(urls.length);
  const pending = [];

  urls.forEach((url, index) => {
    const cacheKey = url.replace(/[^a-zA-Z0-9_-]/g, '').substring(0, 200);
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      results[index] = JSON.parse(cachedData);
    } else {
      const fullUrl = url + (url.includes('?') ? '&' : '?') + 'key=' + apiKey;
      pending.push({ index, url, fullUrl, cacheKey });
    }
  });

  for (let attempt = 0; pending.length && attempt < 4; attempt++) {
    const requests = pending.map(item => {
      recordYouTubeQuota_(item.fullUrl);
      return { url: item.fullUrl, muteHttpExceptions: true };
    });
    const responses = UrlFetchApp.fetchAll(requests);
    const retry = [];

    responses.forEach((response, responseIndex) => {
      const item = pending[responseIndex];
      let json;
      try {
        json = JSON.parse(response.getContentText());
      } catch (e) {
        if (attempt < 3) retry.push(item);
        else throw new Error("Không thể đọc phản hồi Google API: " + e.message);
        return;
      }

      if (json.error) {
        const errReason = (json.error.errors && json.error.errors.length > 0) ? json.error.errors[0].reason : "";
        if (errReason === "quotaExceeded") {
          throw new Error(`[QUOTA EXCEEDED] Bạn đã dùng hết Quota API trong ngày hôm nay.\n${getYouTubeQuotaUsageMessage_()}`);
        }
        if (errReason === "rateLimitExceeded" || errReason === "userRateLimitExceeded" || response.getResponseCode() === 429) {
          if (attempt < 3) retry.push(item);
          else throw new Error("Không thể kết nối máy chủ Google: " + json.error.message);
          return;
        }
        throw new Error(`[API ERROR] ${json.error.message}`);
      }

      if (json.items) {
        try { cache.put(item.cacheKey, response.getContentText(), 1800); } catch (e) {}
      }
      results[item.index] = json;
    });

    pending.length = 0;
    pending.push.apply(pending, retry);
    if (pending.length) Utilities.sleep((attempt + 1) * 1500);
  }

  return results;
}

function getLastNonEmptyRowInColumnFast_(sheet, column, minRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < minRow) return 0;
  const values = sheet.getRange(minRow, column, lastRow - minRow + 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] !== null && values[i][0] !== undefined) return minRow + i;
  }
  return 0;
}

function deleteRowsInBlocks_(sheet, rows) {
  const uniqueRows = [...new Set(rows)].sort((a, b) => b - a);
  let deleted = 0;
  for (let i = 0; i < uniqueRows.length;) {
    const endRow = uniqueRows[i];
    let startRow = endRow;
    i++;
    while (i < uniqueRows.length && uniqueRows[i] === startRow - 1) {
      startRow = uniqueRows[i];
      i++;
    }
    sheet.deleteRows(startRow, endRow - startRow + 1);
    deleted += endRow - startRow + 1;
  }
  return deleted;
}

function writeSparseColumnValues_(sheet, startRow, column, values) {
  let written = 0;
  for (let i = 0; i < values.length;) {
    if (values[i] === null || values[i] === undefined) {
      i++;
      continue;
    }
    const blockStart = i;
    const block = [];
    while (i < values.length && values[i] !== null && values[i] !== undefined) {
      block.push([values[i]]);
      i++;
    }
    sheet.getRange(startRow + blockStart, column, block.length, 1).setValues(block);
    written += block.length;
  }
  return written;
}

function writeRowValueMapToColumn_(sheet, column, rowValueMap) {
  const rows = Object.keys(rowValueMap).map(r => parseInt(r, 10)).filter(r => r > 0).sort((a, b) => a - b);
  let written = 0;
  for (let i = 0; i < rows.length;) {
    const startRow = rows[i];
    const block = [[rowValueMap[startRow]]];
    let prev = startRow;
    i++;
    while (i < rows.length && rows[i] === prev + 1) {
      block.push([rowValueMap[rows[i]]]);
      prev = rows[i];
      i++;
    }
    sheet.getRange(startRow, column, block.length, 1).setValues(block);
    written += block.length;
  }
  return written;
}

