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
  return YTTools.getApiConfig(type);
}

function getApiKey() {
  return YTTools.getApiKey();
}

function getClientId() {
  return YTTools.getClientId();
}

function getClientSecret() {
  return YTTools.getClientSecret();
}

function getSupadataKey() {
  return YTTools.getSupadataKey();
}

function getYTTranscriptIoToken() {
  return YTTools.getYTTranscriptIoToken();
}

function getAI9RouterApiKey() {
  return YTTools.getAI9RouterApiKey();
}

function getRapidApiTranscriptKey() {
  return YTTools.getRapidApiTranscriptKey();
}

function getRapidApiTranscriptHost() {
  return YTTools.getRapidApiTranscriptHost();
}

function getRapidApiTranscriptEndpoint() {
  return YTTools.getRapidApiTranscriptEndpoint();
}

function getApifyToken() {
  return YTTools.getApifyToken();
}

function getAssemblyAIApiKey() {
  return YTTools.getAssemblyAIApiKey();
}

function getAssemblyAIAudioUrlTemplate() {
  return YTTools.getAssemblyAIAudioUrlTemplate();
}

function getYTTranscriptApiBridgeUrl() {
  return YTTools.getYTTranscriptApiBridgeUrl();
}

function getApiKeyConfig() {
  return YTTools.getApiKeyConfig();
}

function saveApiKeyConfig(params) {
  return YTTools.saveApiKeyConfig(params);
}

function clearApiKeysForTemplate() {
  return YTTools.clearApiKeysForTemplate();
}

function onOpen() {
  return YTTools.onOpen();
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

