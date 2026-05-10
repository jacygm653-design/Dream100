// =========================================================================
// SYSTEM UPDATE CENTER
// =========================================================================

const SYSTEM_VERSION = "2026.05.10.1";
const SYSTEM_UPDATE_MANIFEST_URL_PROP = "SYSTEM_UPDATE_MANIFEST_URL";
const SYSTEM_UPDATE_LOG_SHEET = "SYSTEM UPDATE LOG";

function getSystemVersion() {
  return SYSTEM_VERSION;
}

function getSystemUpdateManifestUrl_() {
  const props = PropertiesService.getDocumentProperties();
  return (props.getProperty(SYSTEM_UPDATE_MANIFEST_URL_PROP) || "").trim();
}

function saveSystemUpdateManifestUrl(url) {
  try {
    url = (url || "").toString().trim();
    const props = PropertiesService.getDocumentProperties();
    if (url) props.setProperty(SYSTEM_UPDATE_MANIFEST_URL_PROP, url);
    else props.deleteProperty(SYSTEM_UPDATE_MANIFEST_URL_PROP);
    return { success: true, message: url ? "Đã lưu URL manifest update." : "Đã xóa URL manifest update.", manifestUrl: url };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function compareSystemVersions_(a, b) {
  const pa = (a || "").toString().split(/[^\d]+/).filter(Boolean).map(Number);
  const pb = (b || "").toString().split(/[^\d]+/).filter(Boolean).map(Number);
  const n = Math.max(pa.length, pb.length);
  for (let i = 0; i < n; i++) {
    const va = pa[i] || 0;
    const vb = pb[i] || 0;
    if (va > vb) return 1;
    if (va < vb) return -1;
  }
  return 0;
}

function fetchSystemUpdateManifest_(manifestUrl) {
  if (!manifestUrl) {
    return {
      success: false,
      message: "Chưa cấu hình URL manifest update. Hãy lưu URL manifest JSON trước.",
      manifest: null
    };
  }
  try {
    const resp = UrlFetchApp.fetch(manifestUrl, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code < 200 || code >= 300) {
      return { success: false, message: "Không tải được manifest. HTTP " + code + ": " + body.slice(0, 300), manifest: null };
    }
    const manifest = JSON.parse(body);
    return { success: true, message: "Đã tải manifest.", manifest };
  } catch (e) {
    return { success: false, message: "Lỗi đọc manifest: " + e.message, manifest: null };
  }
}

function setupSystemUpdateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SYSTEM_UPDATE_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_UPDATE_LOG_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([[
      "TIME", "CURRENT_VERSION", "LATEST_VERSION", "STATUS", "MANIFEST_URL", "MESSAGE", "RAW_MANIFEST"
    ]]).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    try { hideSheetIfPossible_(sheet); } catch (e) {}
  }
  return sheet;
}

function logSystemUpdateCheck_(status) {
  try {
    const sheet = setupSystemUpdateLogSheet_();
    sheet.appendRow([
      Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss"),
      status.currentVersion || "",
      status.latestVersion || "",
      status.updateAvailable ? "UPDATE_AVAILABLE" : (status.success ? "UP_TO_DATE" : "ERROR"),
      status.manifestUrl || "",
      status.message || "",
      status.manifest ? JSON.stringify(status.manifest) : ""
    ]);
  } catch (e) {}
}

function getSystemUpdateStatus() {
  const manifestUrl = getSystemUpdateManifestUrl_();
  const currentVersion = getSystemVersion();
  const fetched = fetchSystemUpdateManifest_(manifestUrl);
  const status = {
    success: fetched.success,
    currentVersion,
    latestVersion: "",
    updateAvailable: false,
    manifestUrl,
    message: fetched.message,
    manifest: fetched.manifest,
    libraryIdentifier: "YTTools",
    updateMode: "Apps Script Library",
    checkedAt: Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss")
  };

  if (fetched.success && fetched.manifest) {
    const manifest = fetched.manifest;
    status.latestVersion = (manifest.version || manifest.latestVersion || "").toString();
    status.releaseNotes = manifest.releaseNotes || manifest.notes || "";
    status.libraryScriptId = manifest.libraryScriptId || "";
    status.libraryVersion = manifest.libraryVersion || manifest.versionNumber || "";
    if (!status.latestVersion) {
      status.success = false;
      status.message = "Manifest thiếu trường version/latestVersion.";
    } else {
      status.updateAvailable = compareSystemVersions_(status.latestVersion, currentVersion) > 0;
      status.message = status.updateAvailable
        ? "Có bản cập nhật mới: " + status.latestVersion
        : "Bạn đang dùng phiên bản mới nhất.";
    }
  }

  logSystemUpdateCheck_(status);
  return status;
}

function showSystemUpdateCenter() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:16px;background:#f8fafc;color:#0f172a;line-height:1.45;">
      ${buildDialogHeader_("Trung tâm cập nhật hệ thống", "Kiểm tra version Library gốc và hướng dẫn cập nhật file nhân bản.", "#2563eb")}
      <div class="dlgFadeBody" style="display:grid;gap:12px;">
        <div style="background:#fff;border:1px solid #dbe4f0;border-radius:10px;padding:12px;">
          <div style="font-size:12px;font-weight:800;color:#475569;margin-bottom:6px;">URL manifest update</div>
          <div style="display:flex;gap:8px;">
            <input id="manifestUrl" value="${getSystemUpdateManifestUrl_().replace(/"/g, "&quot;")}" placeholder="https://raw.githubusercontent.com/.../manifest.json" style="flex:1;padding:10px;border:1px solid #cbd5e1;border-radius:8px;font-size:13px;">
            <button onclick="saveManifest()" style="padding:10px 14px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-weight:800;cursor:pointer;">Lưu</button>
          </div>
        </div>

        <div id="statusBox" style="background:#fff;border:1px solid #dbe4f0;border-radius:10px;padding:12px;white-space:pre-wrap;font-size:13px;">Bấm "Kiểm tra cập nhật" để xem trạng thái.</div>

        <div style="display:flex;justify-content:flex-end;gap:8px;">
          <button onclick="checkUpdate()" id="checkBtn" style="padding:10px 16px;border:0;border-radius:8px;background:#2563eb;color:#fff;font-weight:800;cursor:pointer;">Kiểm tra cập nhật</button>
          <button onclick="google.script.host.close()" style="padding:10px 16px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;color:#334155;font-weight:700;cursor:pointer;">Đóng</button>
        </div>

        <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:12px;font-size:12.5px;color:#1e3a8a;">
          <b>Cách cập nhật bản nhân bản dùng Library:</b><br>
          1. Mở Apps Script của file nhân bản.<br>
          2. Vào mục <b>Thư viện</b> ở thanh trái.<br>
          3. Chọn Library identifier <b>YTTools</b>.<br>
          4. Chọn version mới theo manifest rồi bấm lưu.<br>
          5. Reload Google Sheet để menu nhận code mới.<br><br>
          Hệ thống này không ghi đè dữ liệu Sheet, API key hoặc lịch sử của người dùng.
        </div>
      </div>
      <script>
        function setBox(text, ok){
          var el = document.getElementById('statusBox');
          el.style.borderColor = ok ? '#86efac' : '#fecaca';
          el.style.background = ok ? '#f0fdf4' : '#fff7f7';
          el.textContent = text;
        }
        function saveManifest(){
          var url = document.getElementById('manifestUrl').value.trim();
          google.script.run.withSuccessHandler(function(res){
            setBox(res.message || '', !!res.success);
          }).withFailureHandler(function(err){
            setBox('Lỗi: ' + (err.message || err), false);
          }).saveSystemUpdateManifestUrl(url);
        }
        function checkUpdate(){
          document.getElementById('checkBtn').disabled = true;
          setBox('Đang kiểm tra...', true);
          google.script.run.withSuccessHandler(function(s){
            document.getElementById('checkBtn').disabled = false;
            var text = '';
            text += 'Phiên bản hiện tại: ' + (s.currentVersion || '-') + '\\n';
            text += 'Phiên bản mới nhất: ' + (s.latestVersion || '-') + '\\n';
            text += 'Trạng thái: ' + (s.updateAvailable ? 'CÓ BẢN CẬP NHẬT' : (s.success ? 'ĐÃ MỚI NHẤT' : 'LỖI')) + '\\n';
            text += 'Library identifier: ' + (s.libraryIdentifier || 'YTTools') + '\\n';
            if (s.libraryVersion) text += 'Library version cần chọn: ' + s.libraryVersion + '\\n';
            if (s.libraryScriptId) text += 'Library Script ID: ' + s.libraryScriptId + '\\n';
            text += '\\n' + (s.message || '');
            if (s.releaseNotes) text += '\\n\\nRelease notes:\\n' + s.releaseNotes;
            setBox(text, !!s.success);
          }).withFailureHandler(function(err){
            document.getElementById('checkBtn').disabled = false;
            setBox('Lỗi: ' + (err.message || err), false);
          }).getSystemUpdateStatus();
        }
      </script>
    </div>
  `).setWidth(760).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, "Trung tâm cập nhật hệ thống");
}
