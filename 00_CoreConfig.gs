// =========================================================================
// YOUTUBE TOOLS
// =========================================================================

const SHEET_VIDEO = "DREAM > 100 VIDEO < 3 THÁNG";
const SHEET_CHANNEL = "DREAM >100 KÊNH THEO CHỦ ĐỀ";
const SHEET_ANALYTICS = "ANALYTICS KÊNH CỦA TÔI"; 
const SHEET_VIDEO_ANALYTICS = "ANALYTICS TỪNG VIDEO"; 
const SHEET_API_KEY = "API KEY";
const YOUTUBE_DAILY_QUOTA_LIMIT = 10000;
const COL_CHANNEL_MONETIZATION = 14;
const AI_MEMORY_SHEET = "AI ANALYSIS MEMORY";
const FETCH_VIDEO_HISTORY_SHEET = "FETCH VIDEO CHANNEL HISTORY";
const APP_LOGO_URL = "https://i.ibb.co/TM3vVTDD/logo.png";
const APP_LOGO_LINK = "https://ibb.co/3mZrDs55";
const AI_LOGO_URL = APP_LOGO_URL;
const TASK_PROGRESS_PREFIX = "yt_tools_progress_";
const API_CONFIG_CACHE_KEY = "yt_tools_api_config_v1";
const TASK_PROGRESS_MIN_WRITE_MS = 700;
const TASK_PROGRESS_LAST_WRITE_ = {};

function getSafeCache_() {
  try {
    return CacheService.getDocumentCache() || CacheService.getScriptCache();
  } catch (e) {
    return CacheService.getScriptCache();
  }
}

function initTaskProgress_(progressId, total, label) {
  if (!progressId) return;
  const payload = {
    id: progressId,
    total: Math.max(1, parseInt(total, 10) || 1),
    done: 0,
    pct: 0,
    label: label || "Đang xử lý...",
    status: "running",
    updatedAt: new Date().toISOString()
  };
  TASK_PROGRESS_LAST_WRITE_[progressId] = Date.now();
  getSafeCache_().put(TASK_PROGRESS_PREFIX + progressId, JSON.stringify(payload), 21600);
}

function updateTaskProgress_(progressId, done, total, label) {
  if (!progressId) return;
  total = Math.max(1, parseInt(total, 10) || 1);
  done = Math.max(0, Math.min(total, parseInt(done, 10) || 0));
  const now = Date.now();
  const isFinal = done >= total;
  const lastWrite = TASK_PROGRESS_LAST_WRITE_[progressId] || 0;
  if (!isFinal && now - lastWrite < TASK_PROGRESS_MIN_WRITE_MS) return;
  const payload = {
    id: progressId,
    total: total,
    done: done,
    pct: Math.round((done / total) * 1000) / 10,
    label: label || ("Đã xử lý " + done + "/" + total),
    status: done >= total ? "done" : "running",
    updatedAt: new Date().toISOString()
  };
  TASK_PROGRESS_LAST_WRITE_[progressId] = now;
  getSafeCache_().put(TASK_PROGRESS_PREFIX + progressId, JSON.stringify(payload), 21600);
}

function finishTaskProgress_(progressId, success, label) {
  if (!progressId) return;
  let payload = null;
  try {
    const cached = getSafeCache_().get(TASK_PROGRESS_PREFIX + progressId);
    if (cached) payload = JSON.parse(cached);
  } catch(e) {}
  payload = payload || { id: progressId, total: 1, done: 0 };
  payload.done = success ? payload.total : payload.done;
  payload.pct = success ? 100 : Math.round(((payload.done || 0) / Math.max(1, payload.total || 1)) * 1000) / 10;
  payload.status = success ? "done" : "error";
  payload.label = label || (success ? "Hoàn tất" : "Đã lỗi");
  payload.updatedAt = new Date().toISOString();
  TASK_PROGRESS_LAST_WRITE_[progressId] = Date.now();
  getSafeCache_().put(TASK_PROGRESS_PREFIX + progressId, JSON.stringify(payload), 21600);
}

function getTaskProgress(progressId) {
  if (!progressId) return null;
  const cached = getSafeCache_().get(TASK_PROGRESS_PREFIX + progressId);
  if (!cached) return null;
  try { return JSON.parse(cached); } catch(e) { return null; }
}

function setActiveAdminProgress(progressId) {
  const props = PropertiesService.getDocumentProperties();
  if (progressId) props.setProperty("ACTIVE_ADMIN_PROGRESS_ID", progressId);
  else props.deleteProperty("ACTIVE_ADMIN_PROGRESS_ID");
  return { success: true };
}

function getActiveAdminProgress() {
  const progressId = PropertiesService.getDocumentProperties().getProperty("ACTIVE_ADMIN_PROGRESS_ID");
  if (!progressId) return { progressId: "", progress: null };
  return { progressId: progressId, progress: getTaskProgress(progressId) };
}

function buildDialogHeader_(title, subtitle, accentColor) {
  const color = accentColor || "#1a73e8";
  return `
    <style>
      @keyframes dlgFadeIn{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:none}}
      @keyframes dlgSlideUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
      @keyframes dlgPulse{0%,100%{opacity:.55}50%{opacity:1}}
      @keyframes dlgShimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
      @keyframes dlgGlow{0%,100%{box-shadow:0 0 0 ${color}40}50%{box-shadow:0 0 18px ${color}80}}
      @keyframes dlgRipple{from{transform:scale(.4);opacity:.6}to{transform:scale(2.4);opacity:0}}
      @keyframes dlgSpin{to{transform:rotate(360deg)}}
      .dlgHeader{animation:dlgFadeIn .4s cubic-bezier(.2,.7,.2,1) both}
      .dlgProgressWrap{position:relative;height:10px;border-radius:8px;overflow:hidden;background:linear-gradient(180deg,#eef2f7,#e2e8f0);margin-top:14px;border:1px solid #d8dee7;box-shadow:inset 0 1px 3px rgba(15,23,42,.08)}
      .dlgProgressBar{position:absolute;left:0;top:0;bottom:0;width:0%;background:linear-gradient(90deg,${color} 0%,${color}cc 35%,#ffffffaa 50%,${color}cc 65%,${color} 100%);background-size:220% 100%;animation:dlgShimmer 1.5s linear infinite;transition:width .4s cubic-bezier(.2,.7,.2,1);border-radius:8px;box-shadow:0 0 10px ${color}99}
      .dlgProgressBar.idle{width:100%;opacity:.22;animation:dlgPulse 2.4s ease-in-out infinite}
      .dlgProgressBar.error{background:linear-gradient(90deg,#ef4444,#dc2626 50%,#ef4444);background-size:200% 100%;animation:dlgShimmer 1.5s linear infinite;box-shadow:0 0 10px rgba(239,68,68,.6);width:100%}
      .dlgProgressBar.done{background:linear-gradient(90deg,#10b981,#059669 50%,#10b981);background-size:200% 100%;animation:dlgShimmer 1.5s linear infinite;box-shadow:0 0 10px rgba(16,185,129,.6);width:100%}
      .dlgProgressLabel{position:absolute;right:8px;top:50%;transform:translateY(-50%);font-size:10px;font-weight:800;color:#0f172a;background:rgba(255,255,255,.92);padding:1px 7px;border-radius:5px;line-height:1.4;letter-spacing:.02em;box-shadow:0 1px 2px rgba(15,23,42,.08)}
      .dlgProgressMeta{display:flex;justify-content:space-between;font-size:11px;color:#64748b;margin-top:6px;font-weight:600;letter-spacing:.02em}
      .dlgProgressMeta span{display:inline-flex;align-items:center;gap:4px}
      .dlgProgressDot{display:inline-block;width:7px;height:7px;border-radius:50%;background:${color};box-shadow:0 0 6px ${color}aa;animation:dlgPulse 1.4s ease-in-out infinite}
      button:not([disabled]){transition:transform .12s ease,box-shadow .2s ease,filter .2s ease;position:relative;overflow:hidden}
      button:not([disabled]):hover{filter:brightness(1.07);transform:translateY(-1px);box-shadow:0 6px 14px rgba(15,23,42,.12)}
      button:not([disabled]):active{transform:translateY(0);filter:brightness(.95)}
      button[disabled]{opacity:.55;cursor:not-allowed}
      input,select,textarea{transition:border-color .18s ease,box-shadow .2s ease,background .18s ease}
      input:focus,select:focus,textarea:focus{outline:none;box-shadow:0 0 0 4px ${color}26}
      .dlgFadeBody{animation:dlgSlideUp .45s cubic-bezier(.2,.7,.2,1) .05s both}
    </style>
    <div class="dlgHeader" style="position:relative;background:linear-gradient(135deg,#ffffff 0%,#f8fafc 50%,#f1f5f9 100%);border:1px solid #e5e7eb;border-radius:16px;padding:18px 22px 14px 26px;margin-bottom:16px;box-shadow:0 10px 30px -12px rgba(15,23,42,0.18),0 4px 10px -4px rgba(15,23,42,0.06),inset 0 1px 0 rgba(255,255,255,0.9);overflow:hidden;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;">
      <div style="position:absolute;left:0;top:0;bottom:0;width:6px;background:linear-gradient(180deg,${color} 0%,${color}cc 50%,${color}66 100%);box-shadow:2px 0 8px ${color}33;"></div>
      <div style="position:absolute;right:-40px;top:-40px;width:180px;height:180px;background:radial-gradient(circle,${color}1a 0%,transparent 70%);pointer-events:none;animation:dlgPulse 4s ease-in-out infinite;"></div>
      <div style="display:flex;gap:16px;align-items:center;position:relative;z-index:1;">
        <a href="${APP_LOGO_LINK}" target="_blank" style="display:flex;align-items:center;text-decoration:none;flex-shrink:0;">
          <span style="display:inline-flex;align-items:center;justify-content:center;width:64px;height:64px;border-radius:16px;background:radial-gradient(circle at 30% 30%,${color}2e,#ffffff 75%);border:1px solid ${color}40;box-shadow:0 6px 16px ${color}33,inset 0 1px 0 #ffffff;transition:transform .25s ease;" onmouseover="this.style.transform='scale(1.06) rotate(-3deg)'" onmouseout="this.style.transform='none'">
            <img src="${APP_LOGO_URL}" alt="logo" border="0" style="width:50px;height:50px;object-fit:contain;border-radius:11px;background:#fff;">
          </span>
        </a>
        <div style="min-width:0;flex:1;">
          <h2 style="background:linear-gradient(90deg,${color} 0%,${color}d9 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;color:${color};margin:0 0 6px;font-size:24px;line-height:1.2;font-weight:800;letter-spacing:-0.015em;">${title}</h2>
          <div style="font-size:13.5px;color:#475569;line-height:1.55;font-weight:500;">${subtitle || ""}</div>
        </div>
      </div>
      <div class="dlgProgressWrap" id="dlgProgressWrap">
        <div class="dlgProgressBar idle" id="dlgProgressBar"></div>
        <div class="dlgProgressLabel" id="dlgProgressLabel">Sẵn sàng</div>
      </div>
      <div class="dlgProgressMeta"><span><i class="dlgProgressDot"></i><span id="dlgProgressStatus">Chờ thao tác</span></span><span id="dlgProgressTime">0s</span></div>
    </div>
    <script>
      (function(){
        if (window.dlgProgress) return;
        var bar = document.getElementById('dlgProgressBar');
        var label = document.getElementById('dlgProgressLabel');
        var stat = document.getElementById('dlgProgressStatus');
        var timeEl = document.getElementById('dlgProgressTime');
        var rafId = null, startTs = 0, etaSec = 0, lockedDone = false;
        var scrollEl = null, scrollHandler = null;
        function fmt(sec){ if(sec<60) return sec.toFixed(1)+'s'; var m=Math.floor(sec/60); return m+'m'+Math.round(sec%60)+'s'; }
        function tick(){
          if (lockedDone) return;
          var elapsed = (Date.now()-startTs)/1000;
          if (timeEl) timeEl.textContent = fmt(elapsed);
          var pct = etaSec>0 ? Math.min(95, (elapsed/etaSec)*92) : 0;
          if (bar) bar.style.width = pct.toFixed(1)+'%';
          if (label) label.textContent = pct.toFixed(0)+'%';
          rafId = requestAnimationFrame(tick);
        }
        function clearAnim(){ if(rafId){cancelAnimationFrame(rafId);rafId=null;} }
        window.dlgProgress = {
          start: function(eta, statusText){
            clearAnim(); lockedDone = false;
            if (bar){ bar.classList.remove('idle','error','done'); bar.style.width='0%'; }
            startTs = Date.now(); etaSec = Math.max(1, eta||3);
            if (stat) stat.textContent = statusText || 'Đang xử lý...';
            if (label) label.textContent = '0%';
            tick();
          },
          set: function(pct, statusText){
            clearAnim();
            if (bar) bar.classList.remove('idle','error','done');
            pct = Math.max(0, Math.min(100, pct));
            if (bar) bar.style.width = pct.toFixed(1)+'%';
            if (label) label.textContent = pct.toFixed(0)+'%';
            if (statusText && stat) stat.textContent = statusText;
          },
          complete: function(statusText){
            clearAnim(); lockedDone = true;
            if (bar){ bar.classList.remove('idle','error'); bar.classList.add('done'); bar.style.width='100%'; }
            if (label) label.textContent = '100%';
            if (stat) stat.textContent = statusText || 'Hoàn tất';
            if (timeEl) timeEl.textContent = fmt((Date.now()-startTs)/1000);
          },
          fail: function(statusText){
            clearAnim(); lockedDone = true;
            if (bar){ bar.classList.remove('idle','done'); bar.classList.add('error'); }
            if (stat) stat.textContent = statusText || 'Đã lỗi';
            if (timeEl) timeEl.textContent = fmt((Date.now()-startTs)/1000);
          },
          reset: function(){
            clearAnim(); lockedDone = false;
            if (bar){ bar.classList.remove('error','done'); bar.classList.add('idle'); bar.style.width='100%'; }
            if (label) label.textContent = 'Sẵn sàng';
            if (stat) stat.textContent = 'Chờ thao tác';
            if (timeEl) timeEl.textContent = '0s';
          },
          attachScroll: function(elId){
            var doAttach = function(){
              if (scrollEl && scrollHandler) scrollEl.removeEventListener('scroll', scrollHandler);
              scrollEl = document.getElementById(elId);
              if (!scrollEl) return setTimeout(doAttach, 120);
              if (bar) bar.classList.remove('idle','error','done');
              if (stat) stat.textContent = 'Tiến độ đọc';
              scrollHandler = function(){
                var max = scrollEl.scrollHeight - scrollEl.clientHeight;
                var pct = max>0 ? (scrollEl.scrollTop/max)*100 : 0;
                pct = Math.max(0, Math.min(100, pct));
                if (bar) bar.style.width = pct.toFixed(1)+'%';
                if (label) label.textContent = pct.toFixed(0)+'%';
                if (timeEl) timeEl.textContent = pct>=99 ? 'Đã đọc hết' : 'Đang đọc';
              };
              scrollEl.addEventListener('scroll', scrollHandler);
              scrollHandler();
            };
            doAttach();
          }
        };
      })();
    </script>
  `;
}

const COL_VIDEO = {
  STT: 1,
  TITLE: 2,
  LINK: 3,
  TAGS: 4,
  HASHTAGS: 5,
  VIEWS: 6,
  VPH: 7,
  DURATION: 8,
  THUMBNAIL: 9,
  PUBLISHED: 10,
  DESCRIPTION: 11,
  RESERVED: 12,
  SUBTITLE: 13,
  LAST_UPDATE: 14,
  PREV_VIEWS: 15
};

