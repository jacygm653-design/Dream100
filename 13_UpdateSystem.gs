// =========================================================================
// SYSTEM UPDATE CENTER
// =========================================================================

const SYSTEM_VERSION = "2026.05.10.5";
const SYSTEM_DEFAULT_MANIFEST_URL = "https://raw.githubusercontent.com/jacygm653-design/Dream100/main/manifest.json";
const SYSTEM_UPDATE_MANIFEST_URL_PROP = "SYSTEM_UPDATE_MANIFEST_URL";
const SYSTEM_UPDATE_LOG_SHEET = "SYSTEM UPDATE LOG";
const SYSTEM_LIBRARY_IDENTIFIER = "YTTools";

function getSystemVersion() {
  return SYSTEM_VERSION;
}

function getSystemUpdateManifestUrl_() {
  const props = PropertiesService.getDocumentProperties();
  return (props.getProperty(SYSTEM_UPDATE_MANIFEST_URL_PROP) || SYSTEM_DEFAULT_MANIFEST_URL).trim();
}

function saveSystemUpdateManifestUrl(url) {
  try {
    url = (url || "").toString().trim();
    const props = PropertiesService.getDocumentProperties();
    if (url && url !== SYSTEM_DEFAULT_MANIFEST_URL) props.setProperty(SYSTEM_UPDATE_MANIFEST_URL_PROP, url);
    else props.deleteProperty(SYSTEM_UPDATE_MANIFEST_URL_PROP);
    return {
      success: true,
      message: url ? "Da luu URL manifest update." : "Da khoi phuc URL manifest mac dinh.",
      manifestUrl: getSystemUpdateManifestUrl_()
    };
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
    return { success: false, message: "Chua cau hinh URL manifest update.", manifest: null };
  }
  try {
    const resp = UrlFetchApp.fetch(manifestUrl, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code < 200 || code >= 300) {
      return { success: false, message: "Khong tai duoc manifest. HTTP " + code + ": " + body.slice(0, 300), manifest: null };
    }
    const manifest = JSON.parse(body);
    return { success: true, message: "Da tai manifest.", manifest };
  } catch (e) {
    return { success: false, message: "Loi doc manifest: " + e.message, manifest: null };
  }
}

function setupSystemUpdateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SYSTEM_UPDATE_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_UPDATE_LOG_SHEET);
    sheet.getRange(1, 1, 1, 8).setValues([[
      "TIME", "CURRENT_VERSION", "LATEST_VERSION", "STATUS", "MANIFEST_URL", "LIBRARY_VERSION", "MESSAGE", "RAW_MANIFEST"
    ]]).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    try { hideSheetIfPossible_(sheet); } catch (e) {}
  }
  return sheet;
}

function logSystemUpdateEvent_(status) {
  try {
    const sheet = setupSystemUpdateLogSheet_();
    sheet.appendRow([
      Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss"),
      status.currentVersion || "",
      status.latestVersion || "",
      status.status || (status.updateAvailable ? "UPDATE_AVAILABLE" : (status.success ? "UP_TO_DATE" : "ERROR")),
      status.manifestUrl || "",
      status.libraryVersion || "",
      status.message || "",
      status.manifest ? JSON.stringify(status.manifest) : ""
    ]);
  } catch (e) {}
}

function getCurrentLibraryDependencyVersion_() {
  try {
    const scriptId = ScriptApp.getScriptId();
    const content = getAppsScriptProjectContent_(scriptId);
    const manifestFile = findAppsScriptManifestFile_(content.files);
    if (!manifestFile) return "";
    const manifest = JSON.parse(manifestFile.source || "{}");
    const libraries = manifest.dependencies && manifest.dependencies.libraries;
    if (!Array.isArray(libraries)) return "";
    const dep = libraries.find(function(lib) {
      return lib.userSymbol === SYSTEM_LIBRARY_IDENTIFIER;
    });
    return dep ? String(dep.version || "") : "";
  } catch (e) {
    return "";
  }
}

function getSystemUpdateStatus() {
  const manifestUrl = getSystemUpdateManifestUrl_();
  const currentVersion = getSystemVersion();
  const fetched = fetchSystemUpdateManifest_(manifestUrl);
  const status = {
    success: fetched.success,
    status: fetched.success ? "CHECKED" : "ERROR",
    currentVersion,
    latestVersion: "",
    currentLibraryVersion: getCurrentLibraryDependencyVersion_(),
    updateAvailable: false,
    manifestUrl,
    message: fetched.message,
    manifest: fetched.manifest,
    libraryIdentifier: SYSTEM_LIBRARY_IDENTIFIER,
    updateMode: "Apps Script Library dependency update",
    checkedAt: Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss")
  };

  if (fetched.success && fetched.manifest) {
    const manifest = fetched.manifest;
    status.latestVersion = (manifest.version || manifest.latestVersion || "").toString();
    status.releaseNotes = manifest.releaseNotes || manifest.notes || "";
    status.libraryScriptId = manifest.libraryScriptId || "";
    status.libraryVersion = (manifest.libraryVersion || manifest.versionNumber || "").toString();
    if (!status.latestVersion) {
      status.success = false;
      status.status = "ERROR";
      status.message = "Manifest thieu truong version/latestVersion.";
    } else if (!status.libraryScriptId || !status.libraryVersion) {
      status.success = false;
      status.status = "ERROR";
      status.message = "Manifest thieu libraryScriptId hoac libraryVersion.";
    } else {
      status.updateAvailable = compareSystemVersions_(status.latestVersion, currentVersion) > 0 ||
        (!!status.currentLibraryVersion && status.currentLibraryVersion !== status.libraryVersion);
      status.status = status.updateAvailable ? "UPDATE_AVAILABLE" : "UP_TO_DATE";
      status.message = status.updateAvailable
        ? "Co ban cap nhat moi: " + status.latestVersion + " / Library version " + status.libraryVersion
        : "Ban dang dung phien ban moi nhat.";
    }
  }

  logSystemUpdateEvent_(status);
  return status;
}

function getAppsScriptProjectContent_(scriptId) {
  const resp = UrlFetchApp.fetch("https://script.googleapis.com/v1/projects/" + encodeURIComponent(scriptId) + "/content", {
    method: "get",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("Apps Script API getContent HTTP " + code + ": " + body.slice(0, 500));
  }
  return JSON.parse(body);
}

function updateAppsScriptProjectContent_(scriptId, files) {
  const resp = UrlFetchApp.fetch("https://script.googleapis.com/v1/projects/" + encodeURIComponent(scriptId) + "/content", {
    method: "put",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ files: files }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("Apps Script API updateContent HTTP " + code + ": " + body.slice(0, 800));
  }
  return JSON.parse(body);
}

function findAppsScriptManifestFile_(files) {
  return (files || []).find(function(file) {
    return file.name === "appsscript" && file.type === "JSON";
  });
}

function applySystemLibraryUpdate() {
  const status = getSystemUpdateStatus();
  if (!status.success) return status;
  if (!status.updateAvailable) {
    status.status = "UP_TO_DATE";
    status.message = "Khong co ban cap nhat moi de ap dung.";
    return status;
  }

  const manifest = status.manifest || {};
  const targetLibraryId = manifest.libraryScriptId;
  const targetLibraryVersion = String(manifest.libraryVersion || manifest.versionNumber || "");
  const scriptId = ScriptApp.getScriptId();

  if (scriptId === targetLibraryId) {
    status.success = false;
    status.status = "MASTER_LIBRARY";
    status.message = "Day la project Library goc. Khong can tu cap nhat dependency YTTools trong chinh Library goc.";
    logSystemUpdateEvent_(status);
    return status;
  }

  try {
    const content = getAppsScriptProjectContent_(scriptId);
    const files = content.files || [];
    let manifestFile = findAppsScriptManifestFile_(files);
    if (!manifestFile) {
      manifestFile = { name: "appsscript", type: "JSON", source: "{}" };
      files.push(manifestFile);
    }

    const appsscript = JSON.parse(manifestFile.source || "{}");
    appsscript.timeZone = appsscript.timeZone || Session.getScriptTimeZone() || "Asia/Ho_Chi_Minh";
    appsscript.exceptionLogging = appsscript.exceptionLogging || "STACKDRIVER";
    appsscript.dependencies = appsscript.dependencies || {};
    appsscript.dependencies.libraries = appsscript.dependencies.libraries || [];

    let lib = appsscript.dependencies.libraries.find(function(item) {
      return item.userSymbol === SYSTEM_LIBRARY_IDENTIFIER || item.libraryId === targetLibraryId;
    });
    if (!lib) {
      lib = { userSymbol: SYSTEM_LIBRARY_IDENTIFIER, libraryId: targetLibraryId, version: targetLibraryVersion };
      appsscript.dependencies.libraries.push(lib);
    }
    lib.userSymbol = SYSTEM_LIBRARY_IDENTIFIER;
    lib.libraryId = targetLibraryId;
    lib.version = targetLibraryVersion;
    delete lib.developmentMode;

    manifestFile.source = JSON.stringify(appsscript, null, 2);
    updateAppsScriptProjectContent_(scriptId, files);

    const props = PropertiesService.getDocumentProperties();
    props.setProperty("SYSTEM_LAST_APPLIED_VERSION", status.latestVersion);
    props.setProperty("SYSTEM_LAST_APPLIED_LIBRARY_VERSION", targetLibraryVersion);

    status.success = true;
    status.status = "UPDATED";
    status.currentLibraryVersion = targetLibraryVersion;
    status.updateAvailable = false;
    status.message = "Da cap nhat Library " + SYSTEM_LIBRARY_IDENTIFIER + " len version " + targetLibraryVersion + ". Hay reload Google Sheet de nap code moi.";
    logSystemUpdateEvent_(status);
    return status;
  } catch (e) {
    status.success = false;
    status.status = "UPDATE_FAILED";
    status.message = "Cap nhat that bai: " + e.message + "\n\nCan bat Apps Script API va them OAuth scope script.projects trong appsscript.json cua project nhan ban.";
    logSystemUpdateEvent_(status);
    return status;
  }
}

function getInlineSystemUpdateView_() {
  return `
    <div style="font-family:Arial,sans-serif;color:#0f172a;">
      <div id="sysUpdateBox" style="background:#fff;border:1px solid #dbe4f0;border-radius:10px;padding:12px;white-space:pre-wrap;font-size:13px;">Dang kiem tra cap nhat...</div>
      <div style="display:flex;gap:8px;margin-top:12px;flex-wrap:wrap;">
        <button id="sysCheckBtn" onclick="sysCheckUpdate()" class="miniBtn" style="padding:10px 14px;border:0;border-radius:8px;background:#2563eb;color:#fff;font-weight:800;cursor:pointer;">Kiem tra lai</button>
        <button id="sysApplyBtn" onclick="sysApplyUpdate()" class="miniBtn" style="display:none;padding:10px 14px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-weight:800;cursor:pointer;">Cap nhat ngay</button>
        <button onclick="sysOpenDocs()" class="miniBtn" style="padding:10px 14px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;color:#334155;font-weight:800;cursor:pointer;">Huong dan neu cap nhat loi</button>
      </div>
      <div style="margin-top:12px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:12px;font-size:12.5px;color:#1e3a8a;line-height:1.55;">
        Nut <b>Cap nhat ngay</b> chi cap nhat dependency Library <b>YTTools</b> trong project nhan ban. Du lieu Sheet, API key va lich su nguoi dung khong bi ghi de.
      </div>
      <script>
        function sysSetBox(text, ok){
          var el = document.getElementById('sysUpdateBox');
          el.style.borderColor = ok ? '#86efac' : '#fecaca';
          el.style.background = ok ? '#f0fdf4' : '#fff7f7';
          el.textContent = text;
        }
        function sysRenderStatus(s){
          var text = '';
          text += 'Manifest: ' + (s.manifestUrl || '-') + '\\n';
          text += 'Phien ban hien tai: ' + (s.currentVersion || '-') + '\\n';
          text += 'Phien ban moi nhat: ' + (s.latestVersion || '-') + '\\n';
          text += 'Library hien tai: ' + (s.currentLibraryVersion || '-') + '\\n';
          text += 'Library can dung: ' + (s.libraryVersion || '-') + '\\n';
          text += 'Trang thai: ' + (s.updateAvailable ? 'CO BAN CAP NHAT' : (s.success ? 'DA MOI NHAT' : 'LOI')) + '\\n\\n';
          text += s.message || '';
          if (s.releaseNotes) text += '\\n\\nRelease notes:\\n' + s.releaseNotes;
          sysSetBox(text, !!s.success);
          document.getElementById('sysApplyBtn').style.display = s.updateAvailable && s.success ? 'inline-block' : 'none';
        }
        function sysCheckUpdate(){
          document.getElementById('sysCheckBtn').disabled = true;
          sysSetBox('Dang kiem tra cap nhat...', true);
          google.script.run.withSuccessHandler(function(s){
            document.getElementById('sysCheckBtn').disabled = false;
            sysRenderStatus(s);
          }).withFailureHandler(function(err){
            document.getElementById('sysCheckBtn').disabled = false;
            sysSetBox('Loi: ' + (err.message || err), false);
          }).getSystemUpdateStatus();
        }
        function sysApplyUpdate(){
          if (!confirm('Cap nhat Library YTTools len version moi? Sau khi xong can reload Google Sheet.')) return;
          document.getElementById('sysApplyBtn').disabled = true;
          sysSetBox('Dang cap nhat Library dependency...', true);
          google.script.run.withSuccessHandler(function(s){
            document.getElementById('sysApplyBtn').disabled = false;
            sysRenderStatus(s);
            if (s.success && s.status === 'UPDATED') alert('Da cap nhat. Hay reload Google Sheet de nap code moi.');
          }).withFailureHandler(function(err){
            document.getElementById('sysApplyBtn').disabled = false;
            sysSetBox('Loi: ' + (err.message || err), false);
          }).applySystemLibraryUpdate();
        }
        function sysOpenDocs(){
          alert('Neu cap nhat loi: bat Apps Script API trong Google Cloud project, them OAuth scope https://www.googleapis.com/auth/script.projects vao appsscript.json cua project nhan ban, sau do chay lai va cap quyen.');
        }
        sysCheckUpdate();
      </script>
    </div>
  `;
}

function showSystemUpdateCenter() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:16px;background:#f8fafc;color:#0f172a;line-height:1.45;">
      ${buildDialogHeader_("Trung tam cap nhat he thong", "Kiem tra va cap nhat Library YTTools tu manifest GitHub.", "#2563eb")}
      <div class="dlgFadeBody" style="display:grid;gap:12px;">
        <div style="background:#fff;border:1px solid #dbe4f0;border-radius:10px;padding:12px;">
          <div style="font-size:12px;font-weight:800;color:#475569;margin-bottom:6px;">URL manifest update</div>
          <div style="display:flex;gap:8px;">
            <input id="manifestUrl" value="${getSystemUpdateManifestUrl_().replace(/"/g, "&quot;")}" placeholder="${SYSTEM_DEFAULT_MANIFEST_URL}" style="flex:1;padding:10px;border:1px solid #cbd5e1;border-radius:8px;font-size:13px;">
            <button onclick="saveManifest()" style="padding:10px 14px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-weight:800;cursor:pointer;">Luu</button>
          </div>
        </div>
        ${getInlineSystemUpdateView_()}
      </div>
      <script>
        function saveManifest(){
          var url = document.getElementById('manifestUrl').value.trim();
          google.script.run.withSuccessHandler(function(res){
            alert(res.message || '');
            if (typeof sysCheckUpdate === 'function') sysCheckUpdate();
          }).withFailureHandler(function(err){
            alert('Loi: ' + (err.message || err));
          }).saveSystemUpdateManifestUrl(url);
        }
      </script>
    </div>
  `).setWidth(820).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, "Trung tam cap nhat he thong");
}

