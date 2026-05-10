// =========================================================================
// CLIENT BOOTSTRAP FOR CLONED SPREADSHEETS
// Add the master Apps Script Library with identifier: YTTools
// Keep this file in every cloned spreadsheet project.
// =========================================================================

function getTaskProgress(progressId) {
  return YTTools.getTaskProgress(progressId);
}

function setActiveAdminProgress(progressId) {
  return YTTools.setActiveAdminProgress(progressId);
}

function getActiveAdminProgress() {
  return YTTools.getActiveAdminProgress();
}

function getApiConfig(type) {
  var config = getClientApiConfigMap_();
  if (type === 'API_KEY') {
    if (!config.API_KEY) throw new Error('API Key dang trong. Vui long dien vao API KEY!B2 trong file nhan ban nay.');
    return config.API_KEY;
  }
  if (type === 'CLIENT_ID') {
    if (!config.CLIENT_ID) throw new Error('Google OAuth Client ID dang trong. Vui long dien vao API KEY!B3 trong file nhan ban nay.');
    return config.CLIENT_ID;
  }
  if (type === 'CLIENT_SECRET') {
    if (!config.CLIENT_SECRET) throw new Error('Google OAuth Client Secret dang trong. Vui long dien vao API KEY!B4 trong file nhan ban nay.');
    return config.CLIENT_SECRET;
  }
  return config[type] || '';
}

function getApiKey() {
  return getApiConfig('API_KEY');
}

function getClientId() {
  return getApiConfig('CLIENT_ID');
}

function getClientSecret() {
  return getApiConfig('CLIENT_SECRET');
}

function getSupadataKey() {
  return getApiConfig('SUPADATA_KEY');
}

function getYTTranscriptIoToken() {
  return getApiConfig('YT_TRANSCRIPT_IO_TOKEN');
}

function getAI9RouterApiKey() {
  return getApiConfig('AI_9ROUTER_KEY');
}

function getRapidApiTranscriptKey() {
  return getApiConfig('RAPIDAPI_TRANSCRIPT_KEY');
}

function getRapidApiTranscriptHost() {
  return getApiConfig('RAPIDAPI_TRANSCRIPT_HOST');
}

function getRapidApiTranscriptEndpoint() {
  return getApiConfig('RAPIDAPI_TRANSCRIPT_ENDPOINT');
}

function getApifyToken() {
  return getApiConfig('APIFY_TOKEN');
}

function getAssemblyAIApiKey() {
  return getApiConfig('ASSEMBLYAI_API_KEY');
}

function getAssemblyAIAudioUrlTemplate() {
  return getApiConfig('ASSEMBLYAI_AUDIO_URL_TEMPLATE');
}

function getYTTranscriptApiBridgeUrl() {
  return getApiConfig('YT_TRANSCRIPT_API_BRIDGE_URL');
}

function getApiKeyConfig() {
  return getClientApiKeyConfig_();
}

function saveApiKeyConfig(params) {
  return saveClientApiKeyConfig_(params);
}

function clearApiKeysForTemplate() {
  return clearClientApiKeysForTemplate_();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 YouTube Tools')
    .addItem('🎛️  MỞ BẢNG ĐIỀU KHIỂN', 'openAdminPanel')
    .addItem('🔑 Cài đặt API Keys', 'openClientApiKeyManager')
    .addItem('🔄 Kiểm tra cập nhật hệ thống', 'showSystemUpdateCenter')
    .addSeparator()
    .addItem('📘 Trung tâm hướng dẫn hệ thống', 'showSystemBrandInfo')
    .addItem('🩺 Kiểm tra sức khỏe hệ thống', 'showSystemHealthCheck')
    .addItem('📊 Quota Dashboard', 'showYouTubeQuotaUsage')
    .addSeparator()
    .addItem('🧹 Xóa API key khỏi file mẫu', 'clearApiKeysForTemplate')
    .addToUi();
}

function showSystemBrandInfo() {
  return YTTools.showSystemBrandInfo();
}

function promptForRange(ui, sheetName) {
  return YTTools.promptForRange(ui, sheetName);
}

function executeRangeAction(params) {
  return YTTools.executeRangeAction(params);
}

function executeSingleRowAction(params) {
  return YTTools.executeSingleRowAction(params);
}

function showYouTubeQuotaUsage() {
  return YTTools.showYouTubeQuotaUsage();
}

function showSystemHealthCheck() {
  return YTTools.showSystemHealthCheck();
}

function getInlineView(viewName) {
  if (viewName === 'apiKey') return getClientApiKeyInlineView_();
  return YTTools.getInlineView(viewName);
}

function getRecentRowsForDisplay(type, count) {
  return YTTools.getRecentRowsForDisplay(type, count);
}

function adminRunSilent(actionName, progressId) {
  return YTTools.adminRunSilent(actionName, progressId);
}

function getOAuthAuthUrl() {
  return YTTools.getOAuthAuthUrl();
}

function getSystemHealthData() {
  return YTTools.getSystemHealthData();
}

function getQuotaSnapshot() {
  return YTTools.getQuotaSnapshot();
}

function getAdminDashboardStats() {
  return YTTools.getAdminDashboardStats();
}

function adminRun(actionName) {
  return YTTools.adminRun(actionName);
}

function openAdminPanel() {
  return YTTools.openAdminPanel();
}

function fetchAPIWithRetry(url, apiKey) {
  return YTTools.fetchAPIWithRetry(url, apiKey);
}

function updateFastVideoSheet() {
  return YTTools.updateFastVideoSheet();
}

function updateVideoSheet() {
  return YTTools.updateVideoSheet();
}

function updateRangeVideoSheet() {
  return YTTools.updateRangeVideoSheet();
}

function executeVideoUpdate(sheet, startRow, endRow, updateMode = 'full', progressId) {
  return YTTools.executeVideoUpdate(sheet, startRow, endRow, updateMode, progressId);
}

function updateSpecificVideoRow() {
  return YTTools.updateSpecificVideoRow();
}

function getYouTubeTranscript(videoId) {
  return YTTools.getYouTubeTranscript(videoId);
}

function fetchTranscriptViaSupadata(videoId, apiKey) {
  return YTTools.fetchTranscriptViaSupadata(videoId, apiKey);
}

function pollSupadataJob(jobId, apiKey) {
  return YTTools.pollSupadataJob(jobId, apiKey);
}

function fetchTranscriptViaYTTranscriptIo(videoId, token) {
  return YTTools.fetchTranscriptViaYTTranscriptIo(videoId, token);
}

function fetchTranscriptViaRapidApi(videoId, apiKey, host, endpoint) {
  return YTTools.fetchTranscriptViaRapidApi(videoId, apiKey, host, endpoint);
}

function fetchTranscriptViaApify(videoId, token) {
  return YTTools.fetchTranscriptViaApify(videoId, token);
}

function fetchTranscriptViaYoutubeTranscriptApiBridge(videoId, bridgeUrl) {
  return YTTools.fetchTranscriptViaYoutubeTranscriptApiBridge(videoId, bridgeUrl);
}

function fetchTranscriptViaAssemblyAI(videoId, apiKey, audioUrlTemplate) {
  return YTTools.fetchTranscriptViaAssemblyAI(videoId, apiKey, audioUrlTemplate);
}

function fetchTranscriptViaTimedText(videoId) {
  return YTTools.fetchTranscriptViaTimedText(videoId);
}

function fetchTranscriptViaInnerTube(videoId) {
  return YTTools.fetchTranscriptViaInnerTube(videoId);
}

function fetchTranscriptViaScrape(videoId) {
  return YTTools.fetchTranscriptViaScrape(videoId);
}

function parseTranscriptXml(xml) {
  return YTTools.parseTranscriptXml(xml);
}

function decodeHtmlEntities(str) {
  return YTTools.decodeHtmlEntities(str);
}

function fetchAllSubtitles() {
  return YTTools.fetchAllSubtitles();
}

function fetchRangeSubtitles() {
  return YTTools.fetchRangeSubtitles();
}

function fetchSpecificSubtitle() {
  return YTTools.fetchSpecificSubtitle();
}

function executeSubtitleFetch(sheet, startRow, endRow, suppressAlert, progressId) {
  return YTTools.executeSubtitleFetch(sheet, startRow, endRow, suppressAlert, progressId);
}

function showSubtitleApiGuide() {
  return YTTools.showSubtitleApiGuide();
}

function retryFailedSubtitles() {
  return YTTools.retryFailedSubtitles();
}

function clearSubtitleColumn() {
  return YTTools.clearSubtitleColumn();
}

function updateFastChannelSheet() {
  return YTTools.updateFastChannelSheet();
}

function updateChannelSheet() {
  return YTTools.updateChannelSheet();
}

function updateRangeChannelSheet() {
  return YTTools.updateRangeChannelSheet();
}

function executeChannelUpdate(sheet, startRow, endRow, updateMode = 'full', progressId) {
  return YTTools.executeChannelUpdate(sheet, startRow, endRow, updateMode, progressId);
}

function updateSpecificChannelRow() {
  return YTTools.updateSpecificChannelRow();
}

function fetchRecentVideosFromChannels() {
  return YTTools.fetchRecentVideosFromChannels();
}

function getFetchVideoHistory() {
  return YTTools.getFetchVideoHistory();
}

function deleteFetchVideoHistory(ids) {
  return YTTools.deleteFetchVideoHistory(ids);
}

function clearFetchVideoHistory() {
  return YTTools.clearFetchVideoHistory();
}

function executeFetchRecentVideosFromChannels(params) {
  return YTTools.executeFetchRecentVideosFromChannels(params);
}

function getVideoTabPlaylistId(urlStr, apiKey) {
  return YTTools.getVideoTabPlaylistId(urlStr, apiKey);
}

function openSearchByTopicDialog() {
  return YTTools.openSearchByTopicDialog();
}

function executeSearchByTopic(params) {
  return YTTools.executeSearchByTopic(params);
}

function openSearchChannelsByTopicDialog() {
  return YTTools.openSearchChannelsByTopicDialog();
}

function executeSearchChannelsByTopic(params) {
  return YTTools.executeSearchChannelsByTopic(params);
}

function checkAndCleanData() {
  return YTTools.checkAndCleanData();
}

function deleteLowViewVideos() {
  return YTTools.deleteLowViewVideos();
}

function cleanInactiveChannels() {
  return YTTools.cleanInactiveChannels();
}

function executeCleanInactiveChannels(threshold) {
  return YTTools.executeCleanInactiveChannels(threshold);
}

function runFullAutoCleanup() {
  return YTTools.runFullAutoCleanup();
}

function recalculateSTT(sheetName) {
  return YTTools.recalculateSTT(sheetName);
}

function getCountryTimezoneInfo(countryCode, dateUTC) {
  return YTTools.getCountryTimezoneInfo(countryCode, dateUTC);
}

function formatPublishedAtMultiline(publishedAtISO, countryCode) {
  return YTTools.formatPublishedAtMultiline(publishedAtISO, countryCode);
}

function parseVietnameseDate(dateVal) {
  return YTTools.parseVietnameseDate(dateVal);
}

function unformatNumber(num) {
  return YTTools.unformatNumber(num);
}

function formatNumber(num) {
  return YTTools.formatNumber(num);
}

function extractVideoId(url) {
  return YTTools.extractVideoId(url);
}

function extractHashtags(text) {
  return YTTools.extractHashtags(text);
}

function parseISODuration(duration) {
  return YTTools.parseISODuration(duration);
}

function getChannelDataFromAPI(url, apiKey) {
  return YTTools.getChannelDataFromAPI(url, apiKey);
}

function getVideoCategoryName(categoryId, regionCode, apiKey) {
  return YTTools.getVideoCategoryName(categoryId, regionCode, apiKey);
}

function setupOfflineToken() {
  return YTTools.setupOfflineToken();
}

function processAuthCode(authCode) {
  return YTTools.processAuthCode(authCode);
}

function clearOfflineToken() {
  return YTTools.clearOfflineToken();
}

function getFreshAccessToken() {
  return YTTools.getFreshAccessToken();
}

function setupAnalyticsSheet() {
  return YTTools.setupAnalyticsSheet();
}

function updateAnalyticsRow(timeframeIndex) {
  return YTTools.updateAnalyticsRow(timeframeIndex);
}

function updateAnalytics7Days() {
  return YTTools.updateAnalytics7Days();
}

function updateAnalytics28Days() {
  return YTTools.updateAnalytics28Days();
}

function updateAnalyticsLifetime() {
  return YTTools.updateAnalyticsLifetime();
}

function updateAnalyticsAll() {
  return YTTools.updateAnalyticsAll();
}

function processVideoAnalytics(timeframeIndex) {
  return YTTools.processVideoAnalytics(timeframeIndex);
}

function updateVideoAnalytics7Days() {
  return YTTools.updateVideoAnalytics7Days();
}

function updateVideoAnalytics28Days() {
  return YTTools.updateVideoAnalytics28Days();
}

function updateVideoAnalyticsLifetime() {
  return YTTools.updateVideoAnalyticsLifetime();
}

function formatSecondsToMMSS(seconds) {
  return YTTools.formatSecondsToMMSS(seconds);
}

function openAISheetAnalyzer() {
  return YTTools.openAISheetAnalyzer();
}

function saveAIAnalyzerSettings(params) {
  return YTTools.saveAIAnalyzerSettings(params);
}

function getAIAnalysisHistory() {
  return YTTools.getAIAnalysisHistory();
}

function deleteAIAnalysisHistory(ids) {
  return YTTools.deleteAIAnalysisHistory(ids);
}

function clearAIAnalysisHistory() {
  return YTTools.clearAIAnalysisHistory();
}

function executeAISheetAnalysis(params) {
  return YTTools.executeAISheetAnalysis(params);
}

function getSystemVersion() {
  return YTTools.getSystemVersion();
}

function saveSystemUpdateManifestUrl(url) {
  return YTTools.saveSystemUpdateManifestUrl(url);
}

function getSystemUpdateStatus() {
  return YTTools.getSystemUpdateStatus();
}

function applySystemLibraryUpdate() {
  return YTTools.applySystemLibraryUpdate();
}

function showSystemUpdateCenter() {
  return YTTools.showSystemUpdateCenter();
}

// =========================================================================
// LOCAL API KEY MANAGER FOR CLONED SPREADSHEETS
// These functions intentionally run in the cloned spreadsheet project so
// API keys are read from and saved to this file's own "API KEY" sheet.
// =========================================================================

var CLIENT_API_KEY_SHEET_NAME = 'API KEY';

function getClientApiFieldMap_() {
  return {
    youtubeDataKey: { row: 2, label: 'YouTube Data API Key', key: 'API_KEY', required: true },
    clientId: { row: 3, label: 'Google OAuth Client ID', key: 'CLIENT_ID', required: false },
    clientSecret: { row: 4, label: 'Google OAuth Client Secret', key: 'CLIENT_SECRET', required: false },
    supadataKey: { row: 5, label: 'Supadata API Key', key: 'SUPADATA_KEY', required: false },
    ytTranscriptToken: { row: 6, label: 'YouTube-Transcript.io Token', key: 'YT_TRANSCRIPT_IO_TOKEN', required: false },
    nineRouterKey: { row: 7, label: '9router API Key / Bearer Token', key: 'AI_9ROUTER_KEY', required: false },
    rapidApiTranscriptKey: { row: 8, label: 'RapidAPI YouTube Transcript Key', key: 'RAPIDAPI_TRANSCRIPT_KEY', required: false },
    rapidApiTranscriptHost: { row: 9, label: 'RapidAPI YouTube Transcript Host', key: 'RAPIDAPI_TRANSCRIPT_HOST', required: false },
    rapidApiTranscriptEndpoint: { row: 10, label: 'RapidAPI YouTube Transcript Endpoint', key: 'RAPIDAPI_TRANSCRIPT_ENDPOINT', required: false },
    apifyToken: { row: 11, label: 'Apify API Token', key: 'APIFY_TOKEN', required: false },
    assemblyAiApiKey: { row: 12, label: 'AssemblyAI API Key', key: 'ASSEMBLYAI_API_KEY', required: false },
    assemblyAiAudioUrlTemplate: { row: 13, label: 'AssemblyAI Audio URL Template', key: 'ASSEMBLYAI_AUDIO_URL_TEMPLATE', required: false },
    ytTranscriptApiBridgeUrl: { row: 14, label: 'youtube-transcript-api Bridge URL', key: 'YT_TRANSCRIPT_API_BRIDGE_URL', required: false }
  };
}

function getClientApiSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CLIENT_API_KEY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CLIENT_API_KEY_SHEET_NAME);
    setupClientApiKeySheet_(sheet);
  }
  return sheet;
}

function setupClientApiKeySheet_(sheet) {
  var map = getClientApiFieldMap_();
  var rows = [['MODE', 'KEY']];
  Object.keys(map).forEach(function(field) {
    rows.push([map[field].label, '']);
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#1f4e8c').setFontColor('#ffffff');
  sheet.getRange('A2:A14').setFontWeight('bold').setBackground('#155bd6').setFontColor('#ffffff');
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 620);
}

function getClientApiConfigMap_() {
  var sheet = getClientApiSheet_();
  var vals = sheet.getRange('B2:B14').getValues().map(function(r) {
    return (r[0] || '').toString().trim();
  });
  return {
    API_KEY: vals[0],
    CLIENT_ID: vals[1],
    CLIENT_SECRET: vals[2],
    SUPADATA_KEY: vals[3],
    YT_TRANSCRIPT_IO_TOKEN: vals[4],
    AI_9ROUTER_KEY: vals[5],
    RAPIDAPI_TRANSCRIPT_KEY: vals[6],
    RAPIDAPI_TRANSCRIPT_HOST: vals[7],
    RAPIDAPI_TRANSCRIPT_ENDPOINT: vals[8],
    APIFY_TOKEN: vals[9],
    ASSEMBLYAI_API_KEY: vals[10],
    ASSEMBLYAI_AUDIO_URL_TEMPLATE: vals[11],
    YT_TRANSCRIPT_API_BRIDGE_URL: vals[12]
  };
}

function getClientApiKeyConfig_() {
  try {
    var sheet = getClientApiSheet_();
    var labels = sheet.getRange('A2:A14').getDisplayValues().map(function(r) { return r[0] || ''; });
    var vals = sheet.getRange('B2:B14').getValues().map(function(r) { return (r[0] || '').toString().trim(); });
    return {
      success: true,
      keys: {
        youtubeDataKey: vals[0],
        clientId: vals[1],
        clientSecret: vals[2],
        supadataKey: vals[3],
        ytTranscriptToken: vals[4],
        nineRouterKey: vals[5],
        rapidApiTranscriptKey: vals[6],
        rapidApiTranscriptHost: vals[7],
        rapidApiTranscriptEndpoint: vals[8],
        apifyToken: vals[9],
        assemblyAiApiKey: vals[10],
        assemblyAiAudioUrlTemplate: vals[11],
        ytTranscriptApiBridgeUrl: vals[12]
      },
      labels: labels,
      sourceSpreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId()
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveClientApiKeyConfig_(params) {
  try {
    params = params || {};
    var sheet = getClientApiSheet_();
    var map = getClientApiFieldMap_();
    var existing = sheet.getRange('A2:B14').getValues();
    Object.keys(map).forEach(function(field) {
      if (params[field] === undefined || params[field] === null) return;
      var meta = map[field];
      var idx = meta.row - 2;
      existing[idx][0] = existing[idx][0] || meta.label;
      existing[idx][1] = params[field].toString().trim();
    });
    sheet.getRange('A2:B14').setValues(existing);
    return { success: true, message: 'Da luu API key vao sheet API KEY cua file nhan ban nay.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function clearClientApiKeysForTemplate_() {
  try {
    var sheet = getClientApiSheet_();
    var map = getClientApiFieldMap_();
    var labels = Object.keys(map).map(function(field) { return [map[field].label]; });
    sheet.getRange('A2:A14').setValues(labels);
    sheet.getRange('B2:B14').clearContent();
    return { success: true, message: 'Da xoa API key trong file nay. Nguoi dung nhan ban se tu nhap key rieng.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getClientApiKeyInlineView_() {
  var cfg = getClientApiKeyConfig_();
  var keys = cfg.success ? cfg.keys : {};
  var fields = [
    ['youtubeDataKey', 'B2', 'YouTube Data API Key', true, 'Dung cho Video, Kenh, Tim kiem.'],
    ['clientId', 'B3', 'Google OAuth Client ID', false, 'Dung cho Analytics OAuth.'],
    ['clientSecret', 'B4', 'Google OAuth Client Secret', false, 'Dung cho Analytics OAuth.'],
    ['supadataKey', 'B5', 'Supadata API Key', false, 'Nguon transcript uu tien.'],
    ['ytTranscriptToken', 'B6', 'YouTube-Transcript.io Token', false, 'Nguon transcript du phong.'],
    ['nineRouterKey', 'B7', '9router API Key / Bearer Token', false, 'Dung cho AI phan tich Sheet.'],
    ['rapidApiTranscriptKey', 'B8', 'RapidAPI YouTube Transcript Key', false, 'Key RapidAPI youtube-transcript1.'],
    ['rapidApiTranscriptHost', 'B9', 'RapidAPI YouTube Transcript Host', false, 'X-RapidAPI-Host lay tu RapidAPI.'],
    ['rapidApiTranscriptEndpoint', 'B10', 'RapidAPI YouTube Transcript Endpoint', false, 'Endpoint co the dung {VIDEO_ID} hoac {URL}.'],
    ['apifyToken', 'B11', 'Apify API Token', false, 'Token Apify actor youtube-transcript-extractor.'],
    ['assemblyAiApiKey', 'B12', 'AssemblyAI API Key', false, 'API key AssemblyAI.'],
    ['assemblyAiAudioUrlTemplate', 'B13', 'AssemblyAI Audio URL Template', false, 'URL media truc tiep/template cho AssemblyAI.'],
    ['ytTranscriptApiBridgeUrl', 'B14', 'youtube-transcript-api Bridge URL', false, 'URL web service tu host cho python package.']
  ];
  return `
    <style>
      .clientApiWrap{font-family:Arial,sans-serif;color:#0f172a;display:grid;gap:12px}
      .clientApiHead{background:#eef6ff;border:1px solid #bfdbfe;border-radius:10px;padding:12px}
      .clientApiTitle{font-size:18px;font-weight:900;margin-bottom:4px}
      .clientApiSub{font-size:12.5px;color:#475569;line-height:1.45}
      .clientApiGrid{display:grid;gap:9px}
      .clientApiField{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px}
      .clientApiTop{display:flex;gap:8px;align-items:center;margin-bottom:7px}
      .clientApiCell{background:#0f172a;color:#fff;border-radius:6px;padding:2px 7px;font-size:10px;font-weight:900}
      .clientApiLabel{font-size:12.5px;font-weight:800;flex:1}
      .clientApiReq{font-size:10px;font-weight:900;color:#b91c1c;background:#fee2e2;border-radius:999px;padding:2px 7px}
      .clientApiInput{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;padding:9px 10px;font-size:13px}
      .clientApiDesc{font-size:11.5px;color:#64748b;margin-top:5px}
      .clientApiActions{display:flex;gap:8px;align-items:center;position:sticky;bottom:0;background:#f8fafc;padding:10px 0}
      .clientApiBtn{border:0;border-radius:9px;padding:10px 14px;background:#16a34a;color:#fff;font-weight:900;cursor:pointer}
      .clientApiGhost{border:1px solid #cbd5e1;border-radius:9px;padding:10px 14px;background:#fff;color:#334155;font-weight:800;cursor:pointer}
      .clientApiMsg{display:none;border-radius:9px;padding:10px 12px;font-size:12.5px;font-weight:700}
    </style>
    <div class="clientApiWrap">
      <div class="clientApiHead">
        <div class="clientApiTitle">Cai dat API Keys</div>
        <div class="clientApiSub">Man hinh nay chay local trong <b>Bootstrap_Client.gs</b> va luu truc tiep vao sheet <b>API KEY</b> cua file nhan ban hien tai. Khong ghi vao file goc/Library.</div>
      </div>
      <div id="clientApiMsg" class="clientApiMsg"></div>
      <div class="clientApiGrid">
        ${fields.map(function(f) {
          var id = f[0], cell = f[1], label = f[2], required = f[3], desc = f[4];
          var value = (keys[id] || '').toString().replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
          return '<div class="clientApiField">'
            + '<div class="clientApiTop"><span class="clientApiCell">' + cell + '</span><span class="clientApiLabel">' + label + '</span>' + (required ? '<span class="clientApiReq">BAT BUOC</span>' : '') + '</div>'
            + '<input class="clientApiInput" id="api_' + id + '" type="password" value="' + value + '" placeholder="Nhap gia tri cho ' + cell + '">'
            + '<div class="clientApiDesc">' + desc + '</div>'
            + '</div>';
        }).join('')}
      </div>
      <div class="clientApiActions">
        <button class="clientApiBtn" onclick="clientApiSave()">Luu vao API KEY</button>
        <button class="clientApiGhost" onclick="clientApiToggle()">Hien/An key</button>
      </div>
    </div>
    <script>
      var clientApiShown = false;
      var clientApiFields = ${JSON.stringify(fields.map(function(f) { return f[0]; }))};
      function clientApiMessage(text, ok){
        var el = document.getElementById('clientApiMsg');
        el.style.display = 'block';
        el.style.background = ok ? '#dcfce7' : '#fee2e2';
        el.style.color = ok ? '#166534' : '#991b1b';
        el.textContent = text;
      }
      function clientApiToggle(){
        clientApiShown = !clientApiShown;
        clientApiFields.forEach(function(id){
          var input = document.getElementById('api_' + id);
          if (input) input.type = clientApiShown ? 'text' : 'password';
        });
      }
      function clientApiSave(){
        var params = {};
        clientApiFields.forEach(function(id){
          var input = document.getElementById('api_' + id);
          params[id] = input ? input.value : '';
        });
        google.script.run
          .withSuccessHandler(function(res){
            clientApiMessage((res && res.message) || 'Da luu.', !!(res && res.success));
          })
          .withFailureHandler(function(err){
            clientApiMessage('Loi: ' + ((err && err.message) || err), false);
          })
          .saveApiKeyConfig(params);
      }
    </script>
  `;
}

function openClientApiKeyManager() {
  var html = HtmlService
    .createHtmlOutput(getClientApiKeyInlineView_())
    .setWidth(820)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cai dat API Keys');
}

