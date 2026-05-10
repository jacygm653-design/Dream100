// =========================================================================
// HÀM LẤY CẤU HÌNH TỪ SHEET "API KEY"
// =========================================================================
function getApiConfigMap_() {
  const cache = getSafeCache_();
  try {
    const cached = cache.get(API_CONFIG_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_API_KEY);
  if (!sheet) {
    throw new Error(`Không tìm thấy Sheet "${SHEET_API_KEY}". Vui lòng tạo Sheet này để cấu hình hệ thống.`);
  }

  const vals = sheet.getRange("B2:B14").getValues().map(r => (r[0] || "").toString().trim());
  const config = {
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
  try { cache.put(API_CONFIG_CACHE_KEY, JSON.stringify(config), 300); } catch (e) {}
  return config;
}

function clearApiConfigCache_() {
  try { getSafeCache_().remove(API_CONFIG_CACHE_KEY); } catch (e) {}
}

function getApiConfig(type) {
  const config = getApiConfigMap_();

  if (type === 'API_KEY') {
    const key = config.API_KEY;
    if (!key) throw new Error(`API Key đang trống! Vui lòng điền vào ô B2 tại Sheet "${SHEET_API_KEY}".`);
    return key;
  }
  if (type === 'CLIENT_ID') {
    const id = config.CLIENT_ID;
    if (!id) throw new Error(`CLIENT ID đang trống! Vui lòng điền vào ô B3 tại Sheet "${SHEET_API_KEY}".`);
    return id;
  }
  if (type === 'CLIENT_SECRET') {
    const secret = config.CLIENT_SECRET;
    if (!secret) throw new Error(`CLIENT SECRET đang trống! Vui lòng điền vào ô B4 tại Sheet "${SHEET_API_KEY}".`);
    return secret;
  }
  if (type === 'SUPADATA_KEY') {
    return config.SUPADATA_KEY;
  }
  if (type === 'YT_TRANSCRIPT_IO_TOKEN') {
    return config.YT_TRANSCRIPT_IO_TOKEN;
  }
  if (type === 'AI_9ROUTER_KEY') {
    return config.AI_9ROUTER_KEY;
  }
  if (type === 'RAPIDAPI_TRANSCRIPT_KEY') {
    return config.RAPIDAPI_TRANSCRIPT_KEY;
  }
  if (type === 'RAPIDAPI_TRANSCRIPT_HOST') {
    return config.RAPIDAPI_TRANSCRIPT_HOST;
  }
  if (type === 'RAPIDAPI_TRANSCRIPT_ENDPOINT') {
    return config.RAPIDAPI_TRANSCRIPT_ENDPOINT;
  }
  if (type === 'APIFY_TOKEN') {
    return config.APIFY_TOKEN;
  }
  if (type === 'ASSEMBLYAI_API_KEY') {
    return config.ASSEMBLYAI_API_KEY;
  }
  if (type === 'ASSEMBLYAI_AUDIO_URL_TEMPLATE') {
    return config.ASSEMBLYAI_AUDIO_URL_TEMPLATE;
  }
  if (type === 'YT_TRANSCRIPT_API_BRIDGE_URL') {
    return config.YT_TRANSCRIPT_API_BRIDGE_URL;
  }
}

function getApiKey() { return getApiConfig('API_KEY'); }
function getClientId() { return getApiConfig('CLIENT_ID'); }
function getClientSecret() { return getApiConfig('CLIENT_SECRET'); }
function getSupadataKey() { return getApiConfig('SUPADATA_KEY'); }
function getYTTranscriptIoToken() { return getApiConfig('YT_TRANSCRIPT_IO_TOKEN'); }
function getAI9RouterApiKey() { return getApiConfig('AI_9ROUTER_KEY'); }
function getRapidApiTranscriptKey() { return getApiConfig('RAPIDAPI_TRANSCRIPT_KEY'); }
function getRapidApiTranscriptHost() { return getApiConfig('RAPIDAPI_TRANSCRIPT_HOST'); }
function getRapidApiTranscriptEndpoint() { return getApiConfig('RAPIDAPI_TRANSCRIPT_ENDPOINT'); }
function getApifyToken() { return getApiConfig('APIFY_TOKEN'); }
function getAssemblyAIApiKey() { return getApiConfig('ASSEMBLYAI_API_KEY'); }
function getAssemblyAIAudioUrlTemplate() { return getApiConfig('ASSEMBLYAI_AUDIO_URL_TEMPLATE'); }
function getYTTranscriptApiBridgeUrl() { return getApiConfig('YT_TRANSCRIPT_API_BRIDGE_URL'); }

function setAI9RouterApiKey_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_API_KEY);
  if (!sheet) throw new Error(`Không tìm thấy Sheet "${SHEET_API_KEY}" để lưu API key AI.`);
  const row = sheet.getRange("A7:B7").getValues()[0];
  if (!row[0]) row[0] = "9router API Key / Bearer Token";
  row[1] = (key || "").toString().trim();
  sheet.getRange("A7:B7").setValues([row]);
  clearApiConfigCache_();
}

// =========================================================================
// API KEY MANAGER — đọc / ghi sheet API KEY
// =========================================================================
function getApiKeyConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_API_KEY);
    if (!sheet) {
      // Tự tạo sheet nếu chưa có
      sheet = ss.insertSheet(SHEET_API_KEY);
      _setupApiKeySheet_(sheet);
    }
    const labels = sheet.getRange("A2:A14").getDisplayValues().map(r => r[0] || "");
    const vals   = sheet.getRange("B2:B14").getValues().map(r => (r[0] || "").toString().trim());
    return {
      success: true,
      keys: {
        youtubeDataKey:   vals[0],
        clientId:         vals[1],
        clientSecret:     vals[2],
        supadataKey:      vals[3],
        ytTranscriptToken: vals[4],
        nineRouterKey:    vals[5],
        rapidApiTranscriptKey: vals[6],
        rapidApiTranscriptHost: vals[7],
        rapidApiTranscriptEndpoint: vals[8],
        apifyToken: vals[9],
        assemblyAiApiKey: vals[10],
        assemblyAiAudioUrlTemplate: vals[11],
        ytTranscriptApiBridgeUrl: vals[12]
      },
      labels: {
        youtubeDataKey:   labels[0] || "YouTube Data API Key",
        clientId:         labels[1] || "Google OAuth Client ID",
        clientSecret:     labels[2] || "Google OAuth Client Secret",
        supadataKey:      labels[3] || "Supadata API Key",
        ytTranscriptToken: labels[4] || "YouTube-Transcript.io Token",
        nineRouterKey:    labels[5] || "9router API Key / Bearer Token",
        rapidApiTranscriptKey: labels[6] || "RapidAPI YouTube Transcript Key",
        rapidApiTranscriptHost: labels[7] || "RapidAPI YouTube Transcript Host",
        rapidApiTranscriptEndpoint: labels[8] || "RapidAPI YouTube Transcript Endpoint",
        apifyToken: labels[9] || "Apify API Token",
        assemblyAiApiKey: labels[10] || "AssemblyAI API Key",
        assemblyAiAudioUrlTemplate: labels[11] || "AssemblyAI Audio URL Template",
        ytTranscriptApiBridgeUrl: labels[12] || "youtube-transcript-api Bridge URL"
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveApiKeyConfig(params) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_API_KEY);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_API_KEY);
      _setupApiKeySheet_(sheet);
    }
    const keyMap = {
      youtubeDataKey:    { row: 2, label: "YouTube Data API Key" },
      clientId:          { row: 3, label: "Google OAuth Client ID" },
      clientSecret:      { row: 4, label: "Google OAuth Client Secret" },
      supadataKey:       { row: 5, label: "Supadata API Key" },
      ytTranscriptToken: { row: 6, label: "YouTube-Transcript.io Token" },
      nineRouterKey:     { row: 7, label: "9router API Key / Bearer Token" },
      rapidApiTranscriptKey: { row: 8, label: "RapidAPI YouTube Transcript Key" },
      rapidApiTranscriptHost: { row: 9, label: "RapidAPI YouTube Transcript Host" },
      rapidApiTranscriptEndpoint: { row: 10, label: "RapidAPI YouTube Transcript Endpoint" },
      apifyToken: { row: 11, label: "Apify API Token" },
      assemblyAiApiKey: { row: 12, label: "AssemblyAI API Key" },
      assemblyAiAudioUrlTemplate: { row: 13, label: "AssemblyAI Audio URL Template" },
      ytTranscriptApiBridgeUrl: { row: 14, label: "youtube-transcript-api Bridge URL" }
    };
    const saved = [];
    const skipped = [];
    const existing = sheet.getRange("A2:B14").getValues();
    Object.keys(keyMap).forEach(field => {
      const val = (params[field] !== undefined && params[field] !== null) ? params[field].toString().trim() : null;
      if (val === null) return; // không gửi lên → bỏ qua
      const meta = keyMap[field];
      const idx = meta.row - 2;
      if (!existing[idx][0]) existing[idx][0] = meta.label;
      existing[idx][1] = val;
      if (val) saved.push(meta.label);
      else skipped.push(meta.label + " (đã xóa)");
    });
    sheet.getRange("A2:B14").setValues(existing);
    clearApiConfigCache_();
    const msg = saved.length
      ? "Đã lưu: " + saved.join(", ") + (skipped.length ? " | Đã xóa: " + skipped.join(", ") : "")
      : "Không có thay đổi nào được lưu.";
    return { success: true, message: msg };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _setupApiKeySheet_(sheet) {
  const headers = [
    ["Loại Key / Cấu hình", "Giá trị"],
    ["YouTube Data API Key", ""],
    ["Google OAuth Client ID", ""],
    ["Google OAuth Client Secret", ""],
    ["Supadata API Key", ""],
    ["YouTube-Transcript.io Token", ""],
    ["9router API Key / Bearer Token", ""],
    ["RapidAPI YouTube Transcript Key", ""],
    ["RapidAPI YouTube Transcript Host", ""],
    ["RapidAPI YouTube Transcript Endpoint", ""],
    ["Apify API Token", ""],
    ["AssemblyAI API Key", ""],
    ["AssemblyAI Audio URL Template", ""],
    ["youtube-transcript-api Bridge URL", ""]
  ];
  sheet.getRange(1, 1, headers.length, 2).setValues(headers);
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
  sheet.getRange("A2:A14").setFontWeight("bold").setBackground("#f1f5f9");
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 400);
}

function _ivApiKey_() {
  const cfg = getApiKeyConfig();
  const keys = cfg.success ? cfg.keys : {};
  const SHEET_NAME = SHEET_API_KEY;
  return _ivSharedStyles_() + `
    <style>
      .ak-wrap{display:flex;flex-direction:column;gap:14px}
      .ak-header{display:flex;align-items:center;gap:14px;background:linear-gradient(135deg,#f0fdf4,#dcfce7);border:1px solid #86efac;border-radius:16px;padding:16px 20px;margin-bottom:4px}
      .ak-header-icon{width:56px;height:56px;border-radius:14px;background:linear-gradient(135deg,#22c55e,#16a34a);display:flex;align-items:center;justify-content:center;font-size:26px;flex-shrink:0;box-shadow:0 6px 14px -4px rgba(34,197,94,.42)}
      .ak-title{font-size:20px;font-weight:900;color:#14532d;letter-spacing:-0.015em;line-height:1.2}
      .ak-sub{font-size:12.5px;color:#166534;font-weight:600;margin-top:3px;line-height:1.5}
      .ak-grid{display:grid;gap:10px}
      .ak-field{background:#fff;border:1px solid #e5e7eb;border-radius:13px;padding:14px 16px;transition:all .2s ease;position:relative}
      .ak-field:hover{border-color:#86efac;box-shadow:0 4px 12px -6px rgba(34,197,94,.2)}
      .ak-field.focused{border-color:#22c55e;box-shadow:0 0 0 4px rgba(34,197,94,.14)}
      .ak-field-header{display:flex;align-items:center;gap:8px;margin-bottom:8px}
      .ak-field-cell{font-size:10px;font-weight:900;color:#fff;background:linear-gradient(135deg,#0f172a,#1e293b);padding:2px 8px;border-radius:6px;letter-spacing:0.04em;flex-shrink:0}
      .ak-field-label{font-size:12.5px;font-weight:800;color:#0f172a;flex:1}
      .ak-field-status{font-size:10px;font-weight:800;padding:2px 8px;border-radius:999px;letter-spacing:0.04em;text-transform:uppercase;flex-shrink:0}
      .ak-field-status.filled{background:#d1fae5;color:#065f46}
      .ak-field-status.empty{background:#fef3c7;color:#92400e}
      .ak-field-status.required{background:#fee2e2;color:#991b1b}
      .ak-input-wrap{position:relative;display:flex;gap:8px;align-items:center}
      .ak-input{flex:1;padding:10px 44px 10px 12px;border:1px solid #d1d5db;border-radius:9px;font-size:13px;color:#0f172a;font-family:inherit;outline:none;transition:all .2s ease;background:#fafafa;width:100%;box-sizing:border-box}
      .ak-input:focus{border-color:#22c55e;box-shadow:0 0 0 4px rgba(34,197,94,.14);background:#fff}
      .ak-eye{position:absolute;right:10px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;font-size:16px;padding:2px 4px;color:#64748b;line-height:1;transition:color .15s}
      .ak-eye:hover{color:#0f172a}
      .ak-desc{font-size:11px;color:#64748b;font-weight:600;margin-top:5px;line-height:1.45}
      .ak-desc b{color:#0f172a}
      .ak-actions{display:flex;align-items:center;gap:12px;margin-top:6px;flex-wrap:wrap}
      .ak-save-btn{padding:12px 28px;background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff;border:0;border-radius:11px;cursor:pointer;font-size:14px;font-weight:800;font-family:inherit;box-shadow:0 6px 14px -4px rgba(34,197,94,.46);transition:all .2s ease;letter-spacing:0.02em}
      .ak-save-btn:hover{filter:brightness(1.07);transform:translateY(-1px)}
      .ak-save-btn:disabled{opacity:.55;cursor:not-allowed;transform:none}
      .ak-reload-btn{padding:11px 18px;background:#fff;border:1px solid #d1d5db;color:#334155;border-radius:11px;cursor:pointer;font-size:13px;font-weight:700;font-family:inherit;transition:all .2s ease}
      .ak-reload-btn:hover{border-color:#22c55e;color:#16a34a;background:#f0fdf4}
      .ak-tip{background:linear-gradient(135deg,#eff6ff,#dbeafe);border:1px solid #93c5fd;border-radius:11px;padding:10px 14px;font-size:12px;color:#1e40af;line-height:1.55;font-weight:600}
      .ak-result{border-radius:11px;padding:11px 14px;font-size:13px;font-weight:700;line-height:1.5;display:none;animation:dlgSlideUp .3s ease both}
      .ak-result.ok{background:linear-gradient(135deg,#d1fae5,#a7f3d0);color:#065f46;border:1px solid #6ee7b7}
      .ak-result.err{background:linear-gradient(135deg,#fee2e2,#fecaca);color:#991b1b;border:1px solid #fca5a5}
      .ak-security-note{display:flex;align-items:flex-start;gap:8px;background:linear-gradient(135deg,#fef9c3,#fef08a);border:1px solid #fde047;border-radius:11px;padding:10px 14px;font-size:11.5px;color:#713f12;line-height:1.5;font-weight:600}
    </style>
    <div class="ak-wrap">
      <div class="ak-header">
        <div class="ak-header-icon">🔑</div>
        <div>
          <div class="ak-title">Cài đặt API Keys</div>
          <div class="ak-sub">Quản lý tập trung toàn bộ API key hệ thống · Lưu thẳng vào Sheet "<b>${SHEET_NAME}</b>"<br>Tất cả giá trị được ẩn đi khi nhập — chỉ bạn mới thấy</div>
        </div>
      </div>

      <div class="ak-security-note">
        <span style="font-size:18px;flex-shrink:0;">🔒</span>
        <div><b>Bảo mật:</b> API key được lưu vào Sheet Google Sheets của bạn, mã hoá bởi Google. Hệ thống không truyền key ra ngoài ngoài các API call chính thức của từng dịch vụ. Không chia sẻ file Sheet với người lạ.</div>
      </div>

      <div class="ak-grid" id="akGrid">
        <!-- Sẽ render bằng JS -->
      </div>

      <div class="ak-result" id="akResult"></div>

      <div class="ak-actions">
        <button class="ak-save-btn" id="akSaveBtn" onclick="akSave()">💾 Lưu tất cả thay đổi</button>
        <button class="ak-reload-btn" onclick="akReload()">🔄 Tải lại giá trị từ Sheet</button>
        <button class="ak-reload-btn" onclick="akRevealAll()" id="akRevealBtn">👁 Hiện tất cả</button>
      </div>

      <div class="ak-tip">
        <b>📋 Thứ tự ưu tiên điền:</b><br>
        1️⃣ <b>B2</b> — YouTube Data API Key (BẮT BUỘC cho mọi tính năng Video/Kênh)<br>
        2️⃣ <b>B3 + B4</b> — OAuth Client ID + Secret (cần cho Analytics)<br>
        3️⃣ <b>B5:B6 + B8:B14</b> — Các nguồn Subtitle/Transcript dự phòng<br>
        4️⃣ <b>B7</b> — 9router key (cần cho AI phân tích Sheet)
      </div>
    </div>

    <script>
      (function(){
        var FIELDS = [
          { id:'youtubeDataKey',    cell:'B2', label:'YouTube Data API Key',            required:true,  desc:'<b>BẮT BUỘC</b> — Dùng cho toàn bộ tính năng Video, Kênh, Tìm kiếm. Tạo tại <b>Google Cloud Console → APIs → YouTube Data API v3 → Credentials</b>.' },
          { id:'clientId',          cell:'B3', label:'Google OAuth Client ID',           required:false, desc:'Cần cho YouTube Analytics. Tạo OAuth Client ID loại <b>Desktop app</b> tại Google Cloud Console.' },
          { id:'clientSecret',      cell:'B4', label:'Google OAuth Client Secret',       required:false, desc:'Đi kèm với Client ID ở trên. Lấy từ cùng credential vừa tạo.' },
          { id:'supadataKey',       cell:'B5', label:'Supadata API Key',                 required:false, desc:'Ưu tiên cho Subtitle. Đăng ký tại <a href="https://supadata.ai" target="_blank" style="color:#1d4ed8;font-weight:700;">supadata.ai</a> → Dashboard → API Keys.' },
          { id:'ytTranscriptToken', cell:'B6', label:'YouTube-Transcript.io Token',      required:false, desc:'Dự phòng cho Subtitle. Đăng ký tại <a href="https://www.youtube-transcript.io" target="_blank" style="color:#1d4ed8;font-weight:700;">youtube-transcript.io</a> → Profile.' },
          { id:'nineRouterKey',     cell:'B7', label:'9router API Key / Bearer Token',   required:false, desc:'Cần cho AI phân tích Sheet. Cài 9router (<code>npm install -g 9router</code>) → dashboard → copy key.' },
          { id:'rapidApiTranscriptKey', cell:'B8', label:'RapidAPI YouTube Transcript Key', required:false, desc:'Key RapidAPI cho API <a href="https://rapidapi.com/thisisgazzar/api/youtube-transcript1" target="_blank" style="color:#1d4ed8;font-weight:700;">youtube-transcript1</a>. Cần điền thêm Host và Endpoint ở B9:B10 từ trang Endpoints của RapidAPI.' },
          { id:'rapidApiTranscriptHost', cell:'B9', label:'RapidAPI YouTube Transcript Host', required:false, desc:'Giá trị <b>X-RapidAPI-Host</b> lấy đúng từ RapidAPI Hub. Không tự đoán nếu trang API đổi host.' },
          { id:'rapidApiTranscriptEndpoint', cell:'B10', label:'RapidAPI YouTube Transcript Endpoint', required:false, desc:'URL endpoint từ RapidAPI, có thể dùng biến <code>{VIDEO_ID}</code> hoặc <code>{URL}</code>. Ví dụ dạng đầy đủ: <code>https://.../...?id={VIDEO_ID}</code>.' },
          { id:'apifyToken', cell:'B11', label:'Apify API Token', required:false, desc:'Token Apify cho Actor <a href="https://apify.com/akash9078/youtube-transcript-extractor" target="_blank" style="color:#1d4ed8;font-weight:700;">akash9078/youtube-transcript-extractor</a>. Hệ thống gọi API v2 run-sync-get-dataset-items.' },
          { id:'assemblyAiApiKey', cell:'B12', label:'AssemblyAI API Key', required:false, desc:'API key AssemblyAI. Chỉ dùng khi có URL audio/video công khai trực tiếp ở B13; AssemblyAI docs không nhận trang YouTube watch như media file.' },
          { id:'assemblyAiAudioUrlTemplate', cell:'B13', label:'AssemblyAI Audio URL Template', required:false, desc:'URL file audio/video công khai hoặc template chứa <code>{VIDEO_ID}</code>/<code>{URL}</code>. Bỏ trống nếu không có nguồn media trực tiếp.' },
          { id:'ytTranscriptApiBridgeUrl', cell:'B14', label:'youtube-transcript-api Bridge URL', required:false, desc:'URL web service tự host dùng Python package <a href="https://pypi.org/project/youtube-transcript-api/" target="_blank" style="color:#1d4ed8;font-weight:700;">youtube-transcript-api</a>. Hệ thống POST JSON <code>{videoId,languages}</code>.' }
        ];
        var initialVals = ${JSON.stringify(keys)};
        var currentVals = Object.assign({}, initialVals);
        var revealed = {};

        function akEsc(s){ return String(s||'').replace(/[&<>"']/g,function(c){return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]);}); }

        function renderGrid(){
          var html = '';
          FIELDS.forEach(function(f){
            var val = currentVals[f.id] || '';
            var filled = val.length > 0;
            var statusClass = filled ? 'filled' : (f.required ? 'required' : 'empty');
            var statusText = filled ? '✓ Đã điền' : (f.required ? '⚠ BẮT BUỘC' : '○ Tuỳ chọn');
            html += '<div class="ak-field" id="akField_'+f.id+'">'
              + '<div class="ak-field-header">'
              + '<span class="ak-field-cell">'+akEsc(f.cell)+'</span>'
              + '<span class="ak-field-label">'+akEsc(f.label)+'</span>'
              + '<span class="ak-field-status '+statusClass+'">'+statusText+'</span>'
              + '</div>'
              + '<div class="ak-input-wrap">'
              + '<input class="ak-input" id="akInp_'+f.id+'" type="password" placeholder="Dán API key vào đây (ẩn khi nhập)..." autocomplete="off" value="'+akEsc(val)+'" oninput="akOnInput(\\''+f.id+'\\')" onfocus="akFocus(\\''+f.id+'\\')" onblur="akBlur(\\''+f.id+'\\')">'
              + '<button class="ak-eye" onclick="akToggle(\\''+f.id+'\\')" id="akEye_'+f.id+'" title="Hiện/ẩn key">👁</button>'
              + '</div>'
              + '<div class="ak-desc">'+f.desc+'</div>'
              + '</div>';
          });
          document.getElementById('akGrid').innerHTML = html;
        }

        function akOnInput(id){
          var val = document.getElementById('akInp_'+id).value;
          currentVals[id] = val;
          var f = FIELDS.find(function(x){return x.id===id;});
          var el = document.getElementById('akField_'+id);
          var st = el.querySelector('.ak-field-status');
          if (val){
            st.className='ak-field-status filled';
            st.textContent='✓ Đã điền';
          } else {
            st.className='ak-field-status '+(f.required?'required':'empty');
            st.textContent=f.required?'⚠ BẮT BUỘC':'○ Tuỳ chọn';
          }
        }

        function akFocus(id){
          document.getElementById('akField_'+id).classList.add('focused');
        }
        function akBlur(id){
          document.getElementById('akField_'+id).classList.remove('focused');
        }

        function akToggle(id){
          var inp = document.getElementById('akInp_'+id);
          var eye = document.getElementById('akEye_'+id);
          if (inp.type==='password'){
            inp.type='text'; eye.textContent='🙈'; revealed[id]=true;
          } else {
            inp.type='password'; eye.textContent='👁'; delete revealed[id];
          }
        }

        window.akRevealAll = function(){
          var btn = document.getElementById('akRevealBtn');
          var allRevealed = Object.keys(revealed).length === FIELDS.length;
          FIELDS.forEach(function(f){
            var inp = document.getElementById('akInp_'+f.id);
            var eye = document.getElementById('akEye_'+f.id);
            if (!inp) return;
            if (allRevealed){
              inp.type='password'; eye.textContent='👁'; delete revealed[f.id];
            } else {
              inp.type='text'; eye.textContent='🙈'; revealed[f.id]=true;
            }
          });
          btn.textContent = allRevealed ? '👁 Hiện tất cả' : '🙈 Ẩn tất cả';
        };

        window.akReload = function(){
          if (window.dlgProgress) window.dlgProgress.start(2, 'Đang tải lại API key từ Sheet');
          google.script.run
            .withSuccessHandler(function(cfg){
              if (!cfg.success){ akShowResult('Lỗi: '+(cfg.message||'Không xác định'), true); return; }
              currentVals = Object.assign({}, cfg.keys || {});
              Object.keys(revealed).forEach(function(id){ delete revealed[id]; });
              renderGrid();
              if (window.dlgProgress) window.dlgProgress.complete('Đã tải lại');
              akShowResult('✅ Đã tải lại giá trị mới nhất từ Sheet "'+akEsc("${SHEET_NAME}")+'"', false);
            })
            .withFailureHandler(function(err){
              if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
              akShowResult('Lỗi: '+(err.message||err), true);
            })
            .getApiKeyConfig();
        };

        window.akSave = function(){
          /* Thu thập giá trị từ input (không re-read currentVals để tránh stale closure) */
          var params = {};
          FIELDS.forEach(function(f){
            var inp = document.getElementById('akInp_'+f.id);
            params[f.id] = inp ? inp.value : currentVals[f.id] || '';
          });
          var btn = document.getElementById('akSaveBtn');
          btn.disabled = true;
          btn.textContent = '⏳ Đang lưu...';
          if (window.dlgProgress) window.dlgProgress.start(2, 'Đang lưu API keys vào Sheet');
          google.script.run
            .withSuccessHandler(function(res){
              btn.disabled = false;
              btn.textContent = '💾 Lưu tất cả thay đổi';
              if (res.success){
                if (window.dlgProgress) window.dlgProgress.complete('Đã lưu API keys');
                akShowResult('✅ ' + akEsc(res.message), false);
                /* Cập nhật currentVals */
                FIELDS.forEach(function(f){
                  var inp = document.getElementById('akInp_'+f.id);
                  if (inp) currentVals[f.id] = inp.value;
                });
                /* Refresh health nếu có */
                if (typeof loadStats === 'function') setTimeout(loadStats, 300);
              } else {
                if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+akEsc(res.message));
                akShowResult('❌ ' + akEsc(res.message), true);
              }
            })
            .withFailureHandler(function(err){
              btn.disabled = false;
              btn.textContent = '💾 Lưu tất cả thay đổi';
              if (window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
              akShowResult('❌ Lỗi: '+akEsc(err.message||String(err)), true);
            })
            .saveApiKeyConfig(params);
        };

        function akShowResult(msg, isErr){
          var el = document.getElementById('akResult');
          el.className = 'ak-result '+(isErr?'err':'ok');
          el.innerHTML = msg;
          el.style.display = 'block';
          setTimeout(function(){ if (!isErr) el.style.display='none'; }, 6000);
        }

        /* Keyboard shortcut Ctrl+S để lưu */
        document.addEventListener('keydown', function(e){
          if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='s'){
            e.preventDefault();
            akSave();
          }
        });

        renderGrid();
      })();
    </script>
  `;
}

