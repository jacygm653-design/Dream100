// =========================================================================
// AI SHEET ANALYZER - 9ROUTER
// =========================================================================
function getAIRouterModels_() {
  return [
    "ag/gemini-3.1-pro-high",
    "ag/gemini-3.1-pro-low",
    "ag/gemini-3-flash",
    "ag/claude-sonnet-4-6",
    "ag/claude-opus-4-6-thinking",
    "ag/gpt-oss-120b-medium",
    "gc/gemini-3.1-pro-preview",
    "gc/gemini-3.1-flash-lite-preview",
    "gc/gemini-3-flash-preview"
  ];
}

function setupAIMemorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(AI_MEMORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(AI_MEMORY_SHEET);
    sheet.getRange(1, 1, 1, 9).setValues([[
      "TIME", "PROVIDER", "MODEL", "SOURCE_SHEET", "COLUMNS", "ROWS", "QUESTION", "ANSWER", "CONFIG"
    ]]).setFontWeight("bold").setBackground("#1f2937").setFontColor("white");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 9);
  }
  hideSheetIfPossible_(sheet);
  return sheet;
}

function getAIHistoryItems_(limit) {
  const sheet = setupAIMemorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const take = Math.min(limit || 30, lastRow - 1);
  const values = sheet.getRange(lastRow - take + 1, 1, take, 9).getDisplayValues();
  return values.reverse().map((r, idx) => ({
    id: lastRow - idx,
    time: r[0],
    provider: r[1],
    model: r[2],
    sheet: r[3],
    columns: r[4],
    rows: r[5],
    question: r[6],
    answer: r[7],
    config: r[8]
  }));
}

function getAISheetAnalyzerContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  let savedApiKey = "";
  try { savedApiKey = getAI9RouterApiKey(); } catch(e) { savedApiKey = ""; }
  const hiddenHelperSheets = [AI_MEMORY_SHEET, FETCH_VIDEO_HISTORY_SHEET];
  const sheets = ss.getSheets()
    .filter(sh => hiddenHelperSheets.indexOf(sh.getName()) === -1)
    .map(sh => {
      const lastColumn = sh.getLastColumn();
      const headers = lastColumn > 0 ? sh.getRange(1, 1, 1, lastColumn).getValues()[0].map((h, i) => ({
        index: i + 1,
        label: h ? h.toString() : `Cột ${i + 1}`,
        letter: columnToLetter_(i + 1)
      })) : [];
      return {
        name: sh.getName(),
        rows: sh.getLastRow(),
        columns: lastColumn,
        headers
      };
    });
  return {
    logoUrl: AI_LOGO_URL,
    sheets,
    routerModels: getAIRouterModels_(),
    history: getAIHistoryItems_(30),
    defaults: {
      routerBaseUrl: props.getProperty('AI_9ROUTER_BASE_URL') || 'http://localhost:20128/v1',
      model: props.getProperty('AI_MODEL') || 'ag/gemini-3-flash',
      maxRows: parseInt(props.getProperty('AI_MAX_ROWS') || '300', 10) || 300,
      hasApiKey: !!savedApiKey
    }
  };
}

function openAISheetAnalyzer() {
  const ctx = getAISheetAnalyzerContext_();
  setupAIMemorySheet_();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Inter,Arial,sans-serif;background:#f4f7fb;color:#111827;padding:18px;line-height:1.45;">
      ${buildDialogHeader_("AI phân tích Sheet", "9router OpenAI-compatible. Cài đặt endpoint, phân tích dữ liệu, xem lại lịch sử trong cùng một giao diện.", "#111827")}

      <div style="display:flex;gap:8px;margin-bottom:12px;">
        <button class="tabBtn" onclick="showTab('settings')" id="tab-settings">Cài đặt</button>
        <button class="tabBtn" onclick="showTab('activity')" id="tab-activity">Hoạt động</button>
        <button class="tabBtn" onclick="showTab('history')" id="tab-history">Lịch sử</button>
      </div>

      <style>
        .tabBtn{padding:10px 16px;border:1px solid #cbd5e1;background:#fff;border-radius:8px;cursor:pointer;font-weight:700}
        .tabBtn.active{background:#111827;color:#fff;border-color:#111827}
        .panel{display:none;background:#fff;border:1px solid #d8dee9;border-radius:10px;padding:14px}
        .panel.active{display:block}
        label b{font-size:13px}
        input,select,textarea{font-family:inherit}
        .analysisProgress{margin-top:14px;background:#f8fafc;border:1px solid #d8dee9;border-radius:10px;padding:12px}
        .progressTrack{height:12px;background:#e5e7eb;border-radius:999px;overflow:hidden}
        .progressFill{height:100%;width:0%;background:linear-gradient(90deg,#2563eb,#16a34a);border-radius:999px;transition:width .35s ease}
        .progressMeta{display:flex;justify-content:space-between;margin-top:8px;font-size:12px;color:#475569}
        .answerBox{white-space:pre-wrap;margin-top:12px;background:#fff;border:1px solid #cbd5e1;border-radius:10px;padding:14px;min-height:220px;max-height:420px;overflow:auto;font-size:13px;line-height:1.55}
        .answerBox.ok{border-color:#86efac;background:#f7fff9}
        .answerBox.err{border-color:#fecaca;background:#fff7f7}
        .memoryGrid{display:block;margin-top:10px}
        .memoryPanel{border:1px solid #d8dee9;border-radius:8px;background:#fff;padding:10px}
        .memoryPicker{display:none;position:static;margin-top:8px;background:#fff;border:1px solid #cbd5e1;border-radius:8px;box-shadow:none;padding:10px}
        .memoryPickerBody{display:grid;grid-template-columns:minmax(0,1.2fr) minmax(300px,.8fr);gap:10px;align-items:start}
        .memoryOptionsList{max-height:310px;overflow:auto;border-top:1px solid #eef2f7;padding-top:6px}
        .memoryPreviewBox{margin-top:8px;max-height:335px;overflow:auto;font-size:12px;color:#475569;background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;padding:8px}
        .miniBtn{padding:7px 10px;border:1px solid #cbd5e1;background:#fff;border-radius:6px;cursor:pointer}
        .dangerBtn{padding:7px 10px;border:1px solid #fecaca;background:#fff7f7;color:#b91c1c;border-radius:6px;cursor:pointer}
      </style>

      <div id="panel-settings" class="panel">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
          <label><b>9router base URL</b><input id="routerBaseUrl" value="${ctx.defaults.routerBaseUrl}" style="width:100%;box-sizing:border-box;padding:9px;border:1px solid #cbd5e1;border-radius:6px;"></label>
          <label><b>Model mặc định</b><select id="defaultModel" style="width:100%;padding:9px;border:1px solid #cbd5e1;border-radius:6px;"></select></label>
        </div>
        <label style="display:block;margin-top:10px;"><b>API key / Bearer token nếu endpoint yêu cầu</b><input id="apiKey" type="password" placeholder="${ctx.defaults.hasApiKey ? 'Đã có key trong API KEY!B7. Nhập key mới nếu muốn thay thế.' : 'Nhập key để lưu vào API KEY!B7'}" style="width:100%;box-sizing:border-box;padding:9px;border:1px solid #cbd5e1;border-radius:6px;"></label>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:10px;">
          <label><b>Dòng tối đa mặc định</b><input id="defaultMaxRows" type="number" value="${ctx.defaults.maxRows}" min="1" max="2000" style="width:100%;box-sizing:border-box;padding:9px;border:1px solid #cbd5e1;border-radius:6px;"></label>
          <label><b>Sheet ghi nhớ</b><input value="${AI_MEMORY_SHEET}" disabled style="width:100%;box-sizing:border-box;padding:9px;border:1px solid #cbd5e1;border-radius:6px;background:#f3f4f6;"></label>
        </div>
        <div style="margin-top:12px;font-size:12px;color:#5f6368;background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;padding:10px;">
          9router được gọi qua endpoint OpenAI-compatible <b>/v1/chat/completions</b>. Nếu Apps Script không truy cập được localhost, hãy nhập public URL/tunnel của 9router.
        </div>
        <div style="margin-top:14px;text-align:center;">
          <button onclick="saveSettings()" style="padding:11px 24px;background:#111827;color:#fff;border:none;border-radius:8px;font-weight:bold;cursor:pointer;">Lưu cài đặt</button>
        </div>
      </div>

      <div id="panel-activity" class="panel">
        <div style="display:grid;grid-template-columns:1fr 1fr 120px 120px;gap:10px;">
          <label><b>Model phân tích</b><select id="model" style="width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:6px;"></select></label>
          <label><b>Sheet nguồn</b><select id="sheetName" onchange="renderColumns()" style="width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:6px;"></select></label>
          <label><b>Dòng bắt đầu</b><input id="startRow" type="number" value="1" min="1" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #cbd5e1;border-radius:6px;"></label>
          <label><b>Tối đa dòng</b><input id="maxRows" type="number" value="${ctx.defaults.maxRows}" min="1" max="2000" style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #cbd5e1;border-radius:6px;"></label>
        </div>
        <div style="margin-top:10px;border:1px solid #d8dee9;border-radius:8px;padding:10px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
            <b>Chọn cột phân tích</b>
            <label style="font-size:12px;"><input type="checkbox" id="allColumns" checked onchange="renderColumns()"> Toàn bộ sheet</label>
          </div>
          <div id="columnsBox" style="max-height:135px;overflow:auto;display:grid;grid-template-columns:repeat(3,1fr);gap:6px;font-size:12px;"></div>
        </div>
        <label style="display:block;margin-top:10px;"><b>Câu hỏi phân tích</b><textarea id="question" rows="5" placeholder="VD: Phân tích các dòng có tiềm năng nhất, rủi ro dữ liệu, điểm bất thường, và đề xuất hành động." style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:8px;"></textarea></label>
        <div style="display:flex;gap:14px;align-items:center;margin-top:10px;font-size:13px;">
          <label><input type="checkbox" id="useMemory" onchange="toggleMemoryPicker()"> Kết hợp lịch sử đã phân tích</label>
          <label><input type="checkbox" id="saveConfig" checked> Lưu cấu hình sau khi chạy</label>
        </div>
        <div id="memoryPickerWrap" class="memoryGrid" style="display:none;">
          <div class="memoryPanel">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
              <b>Chọn lịch sử cần kết hợp</b>
              <span style="font-size:12px;color:#64748b;">Tick nhiều lịch sử, xem nội dung ngay trong cùng khung.</span>
            </div>
            <div style="position:relative;margin-top:8px;">
              <button type="button" id="memoryPickerBtn" onclick="toggleMemoryDropdown()" class="miniBtn" style="width:100%;text-align:left;">Chưa chọn lịch sử nào</button>
              <div id="memoryPicker" class="memoryPicker">
                <input id="memorySearch" oninput="renderMemoryOptions()" placeholder="Tìm theo model, sheet, câu hỏi..." style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #cbd5e1;border-radius:6px;margin-bottom:8px;">
                <div class="memoryPickerBody">
                  <div>
                    <b style="font-size:12px;color:#334155;">Danh sách lịch sử</b>
                    <div id="memoryOptions" class="memoryOptionsList"></div>
                  </div>
                  <div>
                    <b style="font-size:12px;color:#334155;">Xem trước lịch sử đã chọn</b>
                    <div id="memoryPreview" class="memoryPreviewBox">Chưa chọn lịch sử.</div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div style="margin-top:14px;text-align:center;">
          <button id="runBtn" onclick="runAnalysis()" style="padding:12px 28px;background:#111827;color:#fff;border:none;border-radius:8px;font-weight:bold;cursor:pointer;">Phân tích</button>
          <button onclick="google.script.host.close()" style="padding:12px 18px;background:#9ca3af;color:#fff;border:none;border-radius:8px;margin-left:8px;cursor:pointer;">Đóng</button>
        </div>
        <div class="analysisProgress">
          <div class="progressTrack"><div id="progressFill" class="progressFill"></div></div>
          <div class="progressMeta"><span id="progressText">Sẵn sàng phân tích.</span><b id="progressPct">0%</b></div>
        </div>
        <div id="status" class="answerBox">Kết quả phân tích sẽ hiển thị tại đây.</div>
      </div>

      <div id="panel-history" class="panel">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
          <b>Lịch sử phân tích gần đây</b>
          <div style="display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end;">
            <button onclick="refreshHistory()" class="miniBtn">Tải lại</button>
            <button onclick="deleteSelectedAIHistory()" class="dangerBtn">Xóa đã chọn</button>
            <button onclick="clearAllAIHistory()" style="padding:7px 10px;border:1px solid #b91c1c;background:#b91c1c;color:#fff;border-radius:6px;cursor:pointer;">Xóa toàn bộ</button>
          </div>
        </div>
        <div id="historyBox" style="max-height:520px;overflow:auto;"></div>
      </div>

      <script>
        const CTX = ${JSON.stringify(ctx)};
        function byId(id){ return document.getElementById(id); }
        let progressTimer = null;
        let progressValue = 0;
        const selectedMemoryIds = new Set();
        function setProgress(value, text){
          progressValue = Math.max(0, Math.min(100, value));
          byId('progressFill').style.width = progressValue + '%';
          byId('progressPct').textContent = progressValue + '%';
          if (text) byId('progressText').textContent = text;
        }
        function startProgress(){
          clearInterval(progressTimer);
          setProgress(8, 'Đang chuẩn bị dữ liệu Sheet...');
          progressTimer = setInterval(() => {
            const next = progressValue < 35 ? progressValue + 7 : (progressValue < 78 ? progressValue + 3 : progressValue + 1);
            if (next < 92) setProgress(next, next < 45 ? 'Đang đọc và đóng gói dữ liệu...' : 'Đang gửi yêu cầu phân tích tới 9router...');
          }, 650);
        }
        function finishProgress(success, text){
          clearInterval(progressTimer);
          setProgress(success ? 100 : 100, text || (success ? 'Hoàn tất phân tích.' : 'Phân tích lỗi.'));
        }
        function renderAnswer(res){
          const box = byId('status');
          box.classList.toggle('ok', !!res.success);
          box.classList.toggle('err', !res.success);
          box.textContent = (res.success ? 'OK: ' : 'LỖI: ') + res.message + (res.answer ? '\\n\\n' + res.answer : '');
        }
        function showTab(name){
          ['settings','activity','history'].forEach(t => {
            byId('panel-' + t).classList.toggle('active', t === name);
            byId('tab-' + t).classList.toggle('active', t === name);
          });
        }
        function fillModels(){
          const opts = CTX.routerModels.map(m => '<option value="' + m + '">' + m + '</option>').join('');
          byId('model').innerHTML = opts;
          byId('defaultModel').innerHTML = opts;
          if (CTX.routerModels.includes(CTX.defaults.model)) {
            byId('model').value = CTX.defaults.model;
            byId('defaultModel').value = CTX.defaults.model;
          }
        }
        function renderSheets(){
          byId('sheetName').innerHTML = CTX.sheets.map(s => '<option value="' + s.name.replace(/"/g,'&quot;') + '">' + s.name + ' (' + s.rows + 'x' + s.columns + ')</option>').join('');
          renderColumns();
        }
        function renderColumns(){
          const sheet = CTX.sheets.find(s => s.name === byId('sheetName').value) || CTX.sheets[0];
          const disabled = byId('allColumns').checked;
          byId('columnsBox').innerHTML = (sheet ? sheet.headers : []).map(h =>
            '<label style="display:block;padding:5px;border:1px solid #e5e7eb;border-radius:6px;background:#f9fafb;"><input class="colPick" type="checkbox" value="' + h.index + '" ' + (disabled ? 'disabled checked' : '') + '> ' + h.letter + ': ' + h.label + '</label>'
          ).join('');
        }
        function pickedColumns(){
          if (byId('allColumns').checked) return [];
          return Array.from(document.querySelectorAll('.colPick:checked')).map(x => parseInt(x.value));
        }
        function syncSelectedMemoryIds(){
          const valid = new Set((CTX.history || []).map(h => parseInt(h.id, 10)));
          Array.from(selectedMemoryIds).forEach(id => { if (!valid.has(id)) selectedMemoryIds.delete(id); });
        }
        function toggleMemoryPicker(){
          const enabled = byId('useMemory').checked;
          byId('memoryPickerWrap').style.display = enabled ? 'block' : 'none';
          if (!enabled) byId('memoryPicker').style.display = 'none';
          renderMemoryOptions();
          renderMemoryPreview();
        }
        function toggleMemoryDropdown(){
          const el = byId('memoryPicker');
          el.style.display = el.style.display === 'none' ? 'block' : 'none';
          renderMemoryOptions();
        }
        function updateMemoryButton(){
          const count = selectedMemoryIds.size;
          byId('memoryPickerBtn').textContent = count ? ('Đã chọn ' + count + ' lịch sử') : 'Chưa chọn lịch sử nào';
        }
        function toggleMemoryId(id, checked){
          id = parseInt(id, 10);
          if (checked) selectedMemoryIds.add(id);
          else selectedMemoryIds.delete(id);
          updateMemoryButton();
          renderMemoryPreview();
        }
        function renderMemoryOptions(){
          syncSelectedMemoryIds();
          const q = ((byId('memorySearch') && byId('memorySearch').value) || '').toLowerCase().trim();
          const data = (CTX.history || []).filter(h => {
            const hay = [h.time,h.model,h.sheet,h.question,h.answer].join(' ').toLowerCase();
            return !q || hay.indexOf(q) >= 0;
          });
          byId('memoryOptions').innerHTML = data.length ? data.map(h =>
            '<label style="display:grid;grid-template-columns:18px minmax(0,1fr);gap:8px;align-items:start;padding:7px;border-bottom:1px solid #f1f5f9;cursor:pointer;">' +
            '<input type="checkbox" value="' + h.id + '" ' + (selectedMemoryIds.has(parseInt(h.id,10)) ? 'checked' : '') + ' onchange="toggleMemoryId(this.value,this.checked)">' +
            '<span><b>' + escapeHtml(h.question || '') + '</b><br><span style="color:#64748b;">' + escapeHtml((h.time || '') + ' • ' + (h.model || '') + ' • ' + (h.sheet || '')) + '</span></span>' +
            '</label>'
          ).join('') : '<div style="padding:8px;color:#64748b;">Không có lịch sử phù hợp.</div>';
          updateMemoryButton();
        }
        function renderMemoryPreview(){
          syncSelectedMemoryIds();
          const selected = (CTX.history || []).filter(h => selectedMemoryIds.has(parseInt(h.id, 10)));
          byId('memoryPreview').innerHTML = selected.length ? selected.map(h =>
            '<div style="border:1px solid #e5e7eb;border-radius:8px;background:#f8fafc;padding:8px;margin-bottom:8px;">' +
            '<div style="font-size:11px;color:#64748b;">' + escapeHtml((h.time || '') + ' • ' + (h.model || '') + ' • ' + (h.sheet || '')) + '</div>' +
            '<div style="font-weight:700;margin-top:4px;">' + escapeHtml(h.question || '') + '</div>' +
            '<div style="margin-top:5px;white-space:pre-wrap;">' + escapeHtml((h.answer || '').slice(0, 900)) + (h.answer && h.answer.length > 900 ? '...' : '') + '</div>' +
            '</div>'
          ).join('') : 'Chưa chọn lịch sử.';
        }
        function renderHistory(items){
          const data = items || CTX.history || [];
          byId('historyBox').innerHTML = data.length ? data.map((h, i) =>
            '<div style="border:1px solid #e5e7eb;border-radius:8px;background:#f9fafb;padding:10px;margin-bottom:8px;">' +
            '<div style="font-size:12px;color:#6b7280;">' + h.time + ' • ' + h.model + ' • ' + h.sheet + ' • ' + h.rows + ' dòng</div>' +
            '<div style="font-weight:bold;margin-top:4px;">' + escapeHtml(h.question || '') + '</div>' +
            '<details style="margin-top:6px;"><summary>Xem kết quả</summary><pre style="white-space:pre-wrap;font-family:Arial,sans-serif;background:#fff;border:1px solid #e5e7eb;border-radius:6px;padding:8px;">' + escapeHtml(h.answer || '') + '</pre></details>' +
            '<button onclick="reuseQuestion(' + i + ')" style="margin-top:6px;padding:6px 10px;border:1px solid #cbd5e1;background:#fff;border-radius:6px;cursor:pointer;">Dùng lại câu hỏi</button>' +
            '</div>'
          ).join('') : '<div style="color:#6b7280;">Chưa có lịch sử.</div>';
        }
        function renderHistoryWithActions(items){
          const data = items || CTX.history || [];
          byId('historyBox').innerHTML = data.length ? data.map((h, i) =>
            '<div style="border:1px solid #e5e7eb;border-radius:8px;background:#f9fafb;padding:10px;margin-bottom:8px;">' +
            '<div style="display:grid;grid-template-columns:22px minmax(0,1fr);gap:8px;align-items:start;">' +
            '<input type="checkbox" class="aiHistoryPick" value="' + h.id + '">' +
            '<div>' +
            '<div style="font-size:12px;color:#6b7280;">' + escapeHtml((h.time || '') + ' • ' + (h.model || '') + ' • ' + (h.sheet || '') + ' • ' + (h.rows || '') + ' dòng') + '</div>' +
            '<div style="font-weight:bold;margin-top:4px;">' + escapeHtml(h.question || '') + '</div>' +
            '<details style="margin-top:6px;"><summary>Xem kết quả</summary><pre style="white-space:pre-wrap;font-family:Arial,sans-serif;background:#fff;border:1px solid #e5e7eb;border-radius:6px;padding:8px;">' + escapeHtml(h.answer || '') + '</pre></details>' +
            '<button onclick="reuseQuestion(' + i + ')" class="miniBtn" style="margin-top:6px;">Dùng lại câu hỏi</button>' +
            '</div></div></div>'
          ).join('') : '<div style="color:#6b7280;">Chưa có lịch sử.</div>';
        }
        renderHistory = renderHistoryWithActions;
        function pickedAIHistoryIds(){
          return Array.from(document.querySelectorAll('.aiHistoryPick:checked')).map(x => parseInt(x.value, 10));
        }
        function updateAIHistoryAfterChange(res){
          alert(res.message);
          CTX.history = res.history || [];
          syncSelectedMemoryIds();
          renderHistory(CTX.history);
          renderMemoryOptions();
          renderMemoryPreview();
        }
        function deleteSelectedAIHistory(){
          const ids = pickedAIHistoryIds();
          if (!ids.length) return alert('Vui lòng chọn lịch sử AI cần xóa.');
          google.script.run.withSuccessHandler(updateAIHistoryAfterChange).deleteAIAnalysisHistory(ids);
        }
        function clearAllAIHistory(){
          if (!confirm('Xóa toàn bộ lịch sử AI phân tích Sheet?')) return;
          google.script.run.withSuccessHandler(updateAIHistoryAfterChange).clearAIAnalysisHistory();
        }
        function escapeHtml(s){ return (s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
        function reuseQuestion(i){
          const item = (CTX.history || [])[i];
          if (!item) return;
          byId('question').value = item.question || '';
          showTab('activity');
        }
        function refreshHistory(){
          google.script.run.withSuccessHandler(res => {
            if (res.success) {
              CTX.history = res.history || [];
              syncSelectedMemoryIds();
              renderHistory(CTX.history);
              renderMemoryOptions();
              renderMemoryPreview();
            }
          }).getAIAnalysisHistory();
        }
        function saveSettings(){
          google.script.run.withSuccessHandler(res => {
            alert(res.message);
          }).saveAIAnalyzerSettings({
            routerBaseUrl: byId('routerBaseUrl').value.trim(),
            model: byId('defaultModel').value,
            maxRows: parseInt(byId('defaultMaxRows').value) || 300,
            apiKey: byId('apiKey').value.trim()
          });
        }
        function runAnalysis(){
          const q = byId('question').value.trim();
          if (!q) { alert('Vui lòng nhập câu hỏi phân tích.'); return; }
          if (byId('useMemory').checked && selectedMemoryIds.size === 0) {
            alert('Vui lòng chọn ít nhất 1 lịch sử cần kết hợp, hoặc bỏ tick kết hợp lịch sử.');
            return;
          }
          byId('runBtn').disabled = true;
          byId('status').classList.remove('ok','err');
          byId('status').textContent = 'Đang phân tích. Vui lòng giữ cửa sổ này mở cho đến khi hoàn tất...';
          startProgress();
          const params = {
            model: byId('model').value,
            routerBaseUrl: byId('routerBaseUrl').value.trim(),
            apiKey: byId('apiKey').value.trim(),
            sheetName: byId('sheetName').value,
            startRow: parseInt(byId('startRow').value) || 1,
            maxRows: parseInt(byId('maxRows').value) || 300,
            allColumns: byId('allColumns').checked,
            columns: pickedColumns(),
            question: q,
            useMemory: byId('useMemory').checked,
            memoryIds: Array.from(selectedMemoryIds),
            saveConfig: byId('saveConfig').checked
          };
          google.script.run.withSuccessHandler(res => {
            byId('runBtn').disabled = false;
            finishProgress(!!res.success, res.success ? 'Hoàn tất phân tích và đã lưu lịch sử.' : 'Phân tích lỗi. Xem chi tiết bên dưới.');
            renderAnswer(res);
            if (res.success) refreshHistory();
          }).withFailureHandler(err => {
            byId('runBtn').disabled = false;
            finishProgress(false, 'Phân tích lỗi. Xem chi tiết bên dưới.');
            renderAnswer({ success:false, message: err.message || err });
          }).executeAISheetAnalysis(params);
        }
        fillModels();
        renderSheets();
        renderHistory();
        renderMemoryOptions();
        renderMemoryPreview();
        showTab('settings');
      </script>
    </div>
  `).setWidth(1100).setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'AI phân tích Sheet');
}

function columnToLetter_(column) {
  let temp = "";
  while (column > 0) {
    const mod = (column - 1) % 26;
    temp = String.fromCharCode(65 + mod) + temp;
    column = Math.floor((column - mod) / 26);
  }
  return temp;
}

function buildSelectedSheetData_(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(params.sheetName);
  if (!sheet) throw new Error("Không tìm thấy sheet: " + params.sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0) throw new Error("Sheet không có dữ liệu.");

  const startRow = Math.max(1, parseInt(params.startRow, 10) || 1);
  const maxRows = Math.min(2000, Math.max(1, parseInt(params.maxRows, 10) || 300));
  const rowCount = Math.max(0, Math.min(maxRows, lastRow - startRow + 1));
  if (rowCount === 0) throw new Error("Không có dòng nào trong phạm vi đã chọn.");

  const values = sheet.getRange(startRow, 1, rowCount, lastColumn).getDisplayValues();
  const selectedColumns = params.allColumns ? [] : (params.columns || []).filter(c => c >= 1 && c <= lastColumn);
  const indexes = selectedColumns.length > 0 ? selectedColumns.map(c => c - 1) : values[0].map((_, i) => i);
  const headers = indexes.map(i => `${columnToLetter_(i + 1)}:${values[0][i] || 'Cột ' + (i + 1)}`);
  const rows = values.map(r => indexes.map(i => (r[i] || "").toString().replace(/\s+/g, " ").trim()).join("\t"));
  return {
    tsv: headers.join("\t") + "\n" + rows.join("\n"),
    rowCount,
    columnLabels: headers,
    sourceRange: `${params.sheetName}!${startRow}:${startRow + rowCount - 1}`
  };
}

function getAIMemoryText_(limit) {
  const sheet = setupAIMemorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "";
  const take = Math.min(limit || 5, lastRow - 1);
  const values = sheet.getRange(lastRow - take + 1, 1, take, 8).getDisplayValues();
  return values.map(r => `TIME: ${r[0]}\nMODEL: ${r[2]}\nSHEET: ${r[3]}\nQUESTION: ${r[6]}\nANSWER: ${r[7]}`).join("\n\n---\n\n");
}

function getAIMemoryTextByIds_(ids) {
  const sheet = setupAIMemorySheet_();
  const lastRow = sheet.getLastRow();
  const rows = [...new Set((ids || []).map(id => parseInt(id, 10)).filter(id => id >= 2 && id <= lastRow))].sort((a, b) => a - b);
  if (!rows.length) return "";
  return rows.map(row => {
    const r = sheet.getRange(row, 1, 1, 8).getDisplayValues()[0];
    return `TIME: ${r[0]}\nMODEL: ${r[2]}\nSHEET: ${r[3]}\nQUESTION: ${r[6]}\nANSWER: ${r[7]}`;
  }).join("\n\n---\n\n");
}

function normalizeBaseUrl_(baseUrl) {
  return (baseUrl || "").toString().trim().replace(/\/+$/, "");
}

function aiContentToText_(value) {
  if (value === null || value === undefined) return "";
  if (Array.isArray(value)) {
    return value.map(item => {
      if (typeof item === "string") return item;
      if (item && typeof item === "object") return item.text || item.content || "";
      return "";
    }).join("");
  }
  if (typeof value === "object") return value.text || value.content || JSON.stringify(value);
  return value.toString();
}

function extractAIContentFromJson_(json) {
  if (!json) return "";
  if (json.error) {
    throw new Error(json.error.message || JSON.stringify(json.error));
  }
  if (json.choices && json.choices.length) {
    const choice = json.choices[0];
    if (choice.message && choice.message.content) return aiContentToText_(choice.message.content);
    if (choice.delta && choice.delta.content) return aiContentToText_(choice.delta.content);
    if (choice.text) return aiContentToText_(choice.text);
  }
  if (json.output_text) return aiContentToText_(json.output_text);
  if (json.content) return aiContentToText_(json.content);
  return "";
}

function parse9RouterResponse_(text) {
  const raw = (text || "").toString();
  const trimmed = raw.trim();
  if (!trimmed) throw new Error("9router khong tra ve noi dung phan tich.");

  if (trimmed.indexOf("data:") === 0 || trimmed.indexOf("\ndata:") >= 0) {
    const parts = trimmed.split(/\r?\n/);
    let content = "";
    let lastError = "";
    parts.forEach(line => {
      const clean = line.trim();
      if (!clean || clean.indexOf("data:") !== 0) return;
      const payload = clean.replace(/^data:\s*/, "");
      if (!payload || payload === "[DONE]") return;
      try {
        const item = JSON.parse(payload);
        content += extractAIContentFromJson_(item);
      } catch (e) {
        lastError = e.message;
      }
    });
    if (content) return content;
    throw new Error("9router tra ve SSE nhung khong co noi dung hop le." + (lastError ? " Chi tiet: " + lastError : ""));
  }

  try {
    const json = JSON.parse(trimmed);
    const content = extractAIContentFromJson_(json);
    return content || trimmed;
  } catch (e) {
    throw new Error("9router khong tra ve JSON hop le. Phan dau response: " + trimmed.slice(0, 700));
  }
}

function call9RouterChat_(params, messages) {
  const base = normalizeBaseUrl_(params.routerBaseUrl || "http://localhost:20128/v1");
  const endpoint = base.endsWith("/v1") ? `${base}/chat/completions` : `${base}/v1/chat/completions`;
  const headers = { "Content-Type": "application/json" };
  let savedKey = "";
  try { savedKey = getAI9RouterApiKey(); } catch(e) { savedKey = ""; }
  const apiKey = (params.apiKey || savedKey || "").toString().trim();
  if (apiKey) headers.Authorization = "Bearer " + apiKey;
  const resp = UrlFetchApp.fetch(endpoint, {
    method: "post",
    headers,
    payload: JSON.stringify({ model: params.model, messages, temperature: 0.2, stream: false }),
    muteHttpExceptions: true
  });
  const text = resp.getContentText();
  if (resp.getResponseCode() >= 300) {
    try {
      parse9RouterResponse_(text);
    } catch (e) {
      throw new Error("Loi 9router HTTP " + resp.getResponseCode() + ": " + e.message);
    }
    throw new Error("Loi 9router HTTP " + resp.getResponseCode() + ": " + text.slice(0, 700));
  }
  return parse9RouterResponse_(text);
}

function saveAIAnalyzerSettings(params) {
  try {
    const routerModels = getAIRouterModels_();
    if (!routerModels.includes(params.model)) throw new Error("Model 9router không hợp lệ.");
    const props = PropertiesService.getDocumentProperties();
    props.setProperty('AI_MODEL', params.model);
    props.setProperty('AI_9ROUTER_BASE_URL', params.routerBaseUrl || "");
    props.setProperty('AI_MAX_ROWS', String(params.maxRows || 300));
    if (params.apiKey) setAI9RouterApiKey_(params.apiKey);
    return { success: true, message: "Đã lưu cài đặt 9router." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAIAnalysisHistory() {
  try {
    return { success: true, history: getAIHistoryItems_(30) };
  } catch (e) {
    return { success: false, message: e.message, history: [] };
  }
}

function deleteAIAnalysisHistory(ids) {
  try {
    const sheet = setupAIMemorySheet_();
    const deleted = deleteRowsByIds_(sheet, ids);
    return { success: true, message: `Đã xóa ${deleted} lịch sử AI.`, history: getAIHistoryItems_(30) };
  } catch (e) {
    return { success: false, message: e.message, history: [] };
  }
}

function clearAIAnalysisHistory() {
  try {
    const sheet = setupAIMemorySheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    return { success: true, message: "Đã xóa toàn bộ lịch sử AI.", history: [] };
  } catch (e) {
    return { success: false, message: e.message, history: [] };
  }
}

function executeAISheetAnalysis(params) {
  try {
    if (!params || !params.question) throw new Error("Thiếu câu hỏi phân tích.");
    const routerModels = getAIRouterModels_();
    if (!routerModels.includes(params.model)) throw new Error("Model 9router không hợp lệ.");

    const data = buildSelectedSheetData_(params);
    const memory = params.useMemory ? getAIMemoryTextByIds_(params.memoryIds || []) : "";
    const system = "Bạn là chuyên gia phân tích dữ liệu Google Sheet. Chỉ kết luận từ dữ liệu được cung cấp. Nếu thiếu dữ liệu, nói rõ thiếu gì. Trả lời bằng tiếng Việt, có cấu trúc, ưu tiên insight hành động.";
    const user = `NGUỒN DỮ LIỆU: ${data.sourceRange}\nCỘT: ${data.columnLabels.join(", ")}\nSỐ DÒNG: ${data.rowCount}\n\n${memory ? "TRÍ NHỚ PHÂN TÍCH GẦN ĐÂY:\n" + memory + "\n\n" : ""}DỮ LIỆU TSV:\n${data.tsv}\n\nCÂU HỎI:\n${params.question}`;
    const messages = [{ role: "system", content: system }, { role: "user", content: user }];
    const answer = call9RouterChat_(params, messages);

    const memorySheet = setupAIMemorySheet_();
    memorySheet.appendRow([
      Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss"),
      "9router",
      params.model,
      params.sheetName,
      data.columnLabels.join(", "),
      data.rowCount,
      params.question,
      answer,
      JSON.stringify({ routerBaseUrl: params.routerBaseUrl, sourceRange: data.sourceRange })
    ]);

    if (params.apiKey) setAI9RouterApiKey_(params.apiKey);

    if (params.saveConfig) {
      const props = PropertiesService.getDocumentProperties();
      props.setProperty('AI_MODEL', params.model);
      props.setProperty('AI_9ROUTER_BASE_URL', params.routerBaseUrl || "");
      props.setProperty('AI_MAX_ROWS', String(params.maxRows || 300));
    }

    return { success: true, message: `Đã phân tích ${data.rowCount} dòng từ ${data.sourceRange}. Kết quả đã lưu vào sheet "${AI_MEMORY_SHEET}".`, answer };
  } catch (e) {
    return {
      success: false,
      message: e.message + (e.message.includes("localhost") ? "\n\nLưu ý: Apps Script chạy trên máy chủ Google, nên localhost thường không truy cập được máy của bạn. Hãy dùng public URL/tunnel cho 9router." : "")
    };
  }
}
