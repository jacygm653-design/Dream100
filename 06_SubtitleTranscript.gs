// =========================================================================
// SUBTITLE / TRANSCRIPT
// =========================================================================

/**
 * 📝 LẤY TRANSCRIPT TỪ YOUTUBE - HỆ THỐNG NHIỀU LỚP DỰ PHÒNG
 * 
 * Mỗi lớp được thử tuần tự, nếu lớp trước fail thì dùng lớp sau:
 *   ▶ LỚP 1: Supadata.ai API (100 req/tháng FREE - tỷ lệ thành công ~99%)
 *   ▶ LỚP 2: youtube-transcript.io API (FREE với token - tỷ lệ ~95%)
 *   ▶ LỚP 3: RapidAPI youtube-transcript1 (nếu cấu hình endpoint từ RapidAPI)
 *   ▶ LỚP 4: Apify Actor youtube-transcript-extractor
 *   ▶ LỚP 5: youtube-transcript-api bridge tự host
 *   ▶ LỚP 6: AssemblyAI STT nếu có direct media URL
 *   ▶ LỚP 7+: YouTube timedtext / InnerTube / Scrape
 * 
 * @param {string} videoId - ID video YouTube (11 ký tự)
 * @return {object} { success, text, language, error, source }
 */
function getYouTubeTranscript(videoId) {
  const result = { success: false, text: "", language: "", error: "", source: "" };
  const errors = [];
  
  try {
    const supadataKey = getSupadataKey();
    if (supadataKey) {
      const r = fetchTranscriptViaSupadata(videoId, supadataKey);
      if (r.success) { 
        r.source = "Supadata.ai"; 
        return r; 
      }
      errors.push("Supadata: " + r.error);
    } else {
      errors.push("Supadata: Chưa cấu hình key (B5)");
    }
  } catch (e) {
    errors.push("Supadata: " + e.message);
  }
  
  try {
    const ytIoToken = getYTTranscriptIoToken();
    if (ytIoToken) {
      const r = fetchTranscriptViaYTTranscriptIo(videoId, ytIoToken);
      if (r.success) { 
        r.source = "youtube-transcript.io"; 
        return r; 
      }
      errors.push("YT-Transcript.io: " + r.error);
    } else {
      errors.push("YT-Transcript.io: Chưa cấu hình token (B6)");
    }
  } catch (e) {
    errors.push("YT-Transcript.io: " + e.message);
  }

  try {
    const rapidKey = getRapidApiTranscriptKey();
    const rapidHost = getRapidApiTranscriptHost();
    const rapidEndpoint = getRapidApiTranscriptEndpoint();
    if (rapidKey && rapidHost && rapidEndpoint) {
      const r = fetchTranscriptViaRapidApi(videoId, rapidKey, rapidHost, rapidEndpoint);
      if (r.success) {
        r.source = "RapidAPI youtube-transcript1";
        return r;
      }
      errors.push("RapidAPI: " + r.error);
    } else {
      errors.push("RapidAPI: Chưa cấu hình đủ B8:B10");
    }
  } catch (e) {
    errors.push("RapidAPI: " + e.message);
  }

  try {
    const apifyToken = getApifyToken();
    if (apifyToken) {
      const r = fetchTranscriptViaApify(videoId, apifyToken);
      if (r.success) {
        r.source = "Apify youtube-transcript-extractor";
        return r;
      }
      errors.push("Apify: " + r.error);
    } else {
      errors.push("Apify: Chưa cấu hình token (B11)");
    }
  } catch (e) {
    errors.push("Apify: " + e.message);
  }

  try {
    const bridgeUrl = getYTTranscriptApiBridgeUrl();
    if (bridgeUrl) {
      const r = fetchTranscriptViaYoutubeTranscriptApiBridge(videoId, bridgeUrl);
      if (r.success) {
        r.source = "youtube-transcript-api bridge";
        return r;
      }
      errors.push("youtube-transcript-api: " + r.error);
    } else {
      errors.push("youtube-transcript-api: Chưa cấu hình bridge URL (B14)");
    }
  } catch (e) {
    errors.push("youtube-transcript-api: " + e.message);
  }

  try {
    const assemblyKey = getAssemblyAIApiKey();
    const audioTemplate = getAssemblyAIAudioUrlTemplate();
    if (assemblyKey && audioTemplate) {
      const r = fetchTranscriptViaAssemblyAI(videoId, assemblyKey, audioTemplate);
      if (r.success) {
        r.source = "AssemblyAI";
        return r;
      }
      errors.push("AssemblyAI: " + r.error);
    } else {
      errors.push("AssemblyAI: Chưa cấu hình đủ B12:B13");
    }
  } catch (e) {
    errors.push("AssemblyAI: " + e.message);
  }
  
  try {
    const r = fetchTranscriptViaTimedText(videoId);
    if (r.success) { 
      r.source = "timedtext"; 
      return r; 
    }
    errors.push("Timedtext: " + r.error);
  } catch (e) {
    errors.push("Timedtext: " + e.message);
  }
  
  try {
    const r = fetchTranscriptViaInnerTube(videoId);
    if (r.success) {
      r.source = "innertube";
      return r;
    }
    errors.push("InnerTube: " + r.error);
  } catch (e) {
    errors.push("InnerTube: " + e.message);
  }

  try {
    const r = fetchTranscriptViaScrape(videoId);
    if (r.success) { 
      r.source = "watch-page"; 
      return r; 
    }
    errors.push("Scrape: " + r.error);
  } catch (e) {
    errors.push("Scrape: " + e.message);
  }
  
  result.error = errors.join(" | ");
  return result;
}

// ═══════════════════════════════════════════════════════════════
// LỚP 1: SUPADATA.AI API
// ═══════════════════════════════════════════════════════════════
function fetchTranscriptViaSupadata(videoId, apiKey) {
  const result = { success: false, text: "", language: "", error: "" };
  try {
    const url = `https://api.supadata.ai/v1/transcript?url=https://www.youtube.com/watch?v=${videoId}&text=true`;
    const resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "x-api-key": apiKey },
      muteHttpExceptions: true
    });
    
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    
    if (code === 200) {
      const json = JSON.parse(body);
      if (json.content && typeof json.content === 'string' && json.content.trim()) {
        result.success = true;
        result.text = json.content.trim();
        result.language = json.lang || "unknown";
        return result;
      }
      if (Array.isArray(json.content)) {
        result.success = true;
        result.text = json.content.map(s => s.text).join(' ').trim();
        result.language = json.lang || "unknown";
        return result;
      }
      result.error = "Phản hồi rỗng từ Supadata";
      return result;
    }
    
    if (code === 202) {
      const jobJson = JSON.parse(body);
      if (jobJson.jobId) {
        return pollSupadataJob(jobJson.jobId, apiKey);
      }
    }
    
    if (code === 401 || code === 403) {
      result.error = "Supadata API Key sai hoặc hết quota";
      return result;
    }
    
    if (code === 429) {
      // Retry 1 lan sau 3 giay khi rate limit
      Utilities.sleep(3000);
      try {
        const retry = UrlFetchApp.fetch(url, {
          method: "get",
          headers: { "x-api-key": apiKey },
          muteHttpExceptions: true
        });
        const retryCode = retry.getResponseCode();
        const retryBody = retry.getContentText();
        if (retryCode === 200) {
          const json = JSON.parse(retryBody);
          if (json.content && typeof json.content === 'string' && json.content.trim()) {
            result.success = true;
            result.text = json.content.trim();
            result.language = json.lang || "unknown";
            return result;
          }
          if (Array.isArray(json.content)) {
            result.success = true;
            result.text = json.content.map(function(s){ return s.text; }).join(' ').trim();
            result.language = json.lang || "unknown";
            return result;
          }
        }
        if (retryCode === 202) {
          const jobJson = JSON.parse(retryBody);
          if (jobJson.jobId) return pollSupadataJob(jobJson.jobId, apiKey);
        }
      } catch(re) {}
      result.error = "Supadata rate limit - chờ vài giây (retry thất bại)";
      return result;
    }
    
    try {
      const errJson = JSON.parse(body);
      result.error = errJson.error || errJson.message || `HTTP ${code}`;
    } catch(e) {
      result.error = `HTTP ${code}`;
    }
    return result;
    
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

/**
 * Poll job khi Supadata trả 202 (video dài cần xử lý AI)
 */
function pollSupadataJob(jobId, apiKey) {
  const result = { success: false, text: "", language: "", error: "" };
  const maxAttempts = 10;
  
  for (let i = 0; i < maxAttempts; i++) {
    Utilities.sleep(3000);
    try {
      const resp = UrlFetchApp.fetch(`https://api.supadata.ai/v1/transcript/${jobId}`, {
        method: "get",
        headers: { "x-api-key": apiKey },
        muteHttpExceptions: true
      });
      
      if (resp.getResponseCode() === 200) {
        const json = JSON.parse(resp.getContentText());
        if (json.status === "completed" && json.content) {
          result.success = true;
          result.text = typeof json.content === 'string' ? json.content : 
                        (Array.isArray(json.content) ? json.content.map(s => s.text).join(' ') : "");
          result.language = json.lang || "unknown";
          return result;
        }
        if (json.status === "failed") {
          result.error = "Job xử lý thất bại: " + (json.error || "không rõ");
          return result;
        }
      }
    } catch (e) {
    }
  }
  
  result.error = "Job timeout sau " + maxAttempts + " lần thử";
  return result;
}

// ═══════════════════════════════════════════════════════════════
// LỚP 2: YOUTUBE-TRANSCRIPT.IO API
// ═══════════════════════════════════════════════════════════════
function fetchTranscriptViaYTTranscriptIo(videoId, token) {
  const result = { success: false, text: "", language: "", error: "" };
  try {
    const resp = UrlFetchApp.fetch("https://www.youtube-transcript.io/api/transcripts", {
      method: "post",
      contentType: "application/json",
      headers: { 
        "Authorization": "Basic " + token
      },
      payload: JSON.stringify({ ids: [videoId] }),
      muteHttpExceptions: true
    });
    
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    
    if (code === 200) {
      const json = JSON.parse(body);
      let videoEntry = null;
      if (Array.isArray(json)) {
        videoEntry = json.find(v => v.id === videoId) || json[0];
      } else if (json[videoId]) {
        videoEntry = json[videoId];
      }
      
      if (!videoEntry) {
        result.error = "Không tìm thấy entry cho video này";
        return result;
      }
      
      let tracks = videoEntry.tracks || videoEntry.transcripts || [];
      if (!Array.isArray(tracks) || tracks.length === 0) {
        if (Array.isArray(videoEntry.transcript)) {
          result.success = true;
          result.text = videoEntry.transcript.map(t => t.text || "").join(' ').trim();
          result.language = videoEntry.language || "unknown";
          return result;
        }
        result.error = "Không có track nào";
        return result;
      }
      
      const priority = ['vi', 'en', 'en-US', 'en-GB'];
      let chosen = null;
      for (let lang of priority) {
        chosen = tracks.find(t => (t.language === lang || t.languageCode === lang || t.lang === lang));
        if (chosen) break;
      }
      if (!chosen) chosen = tracks[0];
      
      let segments = chosen.transcript || chosen.segments || chosen.captions || [];
      if (!Array.isArray(segments) || segments.length === 0) {
        result.error = "Track rỗng";
        return result;
      }
      
      result.success = true;
      result.text = segments.map(s => s.text || s.content || "").join(' ').replace(/\s+/g, ' ').trim();
      result.language = chosen.language || chosen.languageCode || "unknown";
      return result;
    }
    
    if (code === 401 || code === 403) {
      result.error = "Token sai hoặc hết quota";
      return result;
    }

    if (code === 402) {
      result.error = "HTTP 402: Not enough credits to fetch these transcripts";
      return result;
    }
    
    if (code === 429) {
      result.error = "Rate limit - chờ 10s";
      return result;
    }
    
    result.error = `HTTP ${code}: ${body.substring(0, 100)}`;
    return result;
    
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

function normalizeTranscriptText_(text) {
  return (text || "").toString().replace(/\s+/g, " ").trim();
}

function transcriptTextFromJson_(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return normalizeTranscriptText_(value);
  if (Array.isArray(value)) {
    return normalizeTranscriptText_(value.map(item => {
      if (typeof item === "string") return item;
      if (!item || typeof item !== "object") return "";
      return item.text || item.content || item.transcript || item.caption || item.sentence || "";
    }).join(" "));
  }
  if (typeof value === "object") {
    const direct = value.text || value.content || value.transcript || value.caption || value.fullText || value.full_text;
    if (direct) return transcriptTextFromJson_(direct);
    const candidates = [
      value.segments,
      value.captions,
      value.subtitles,
      value.items,
      value.data,
      value.result,
      value.results,
      value.transcripts
    ];
    for (let i = 0; i < candidates.length; i++) {
      const text = transcriptTextFromJson_(candidates[i]);
      if (text) return text;
    }
  }
  return "";
}

function fetchTranscriptViaRapidApi(videoId, apiKey, host, endpoint) {
  const result = { success: false, text: "", language: "", error: "" };
  try {
    const videoUrl = "https://www.youtube.com/watch?v=" + videoId;
    let url = endpoint.toString().trim()
      .replace(/\{VIDEO_ID\}/g, encodeURIComponent(videoId))
      .replace(/\{URL\}/g, encodeURIComponent(videoUrl));
    if (!/^https?:\/\//i.test(url)) {
      url = "https://" + host.replace(/\/+$/, "") + "/" + url.replace(/^\/+/, "");
    }
    if (url.indexOf("{") >= 0) {
      result.error = "Endpoint RapidAPI còn biến chưa thay thế. Dùng {VIDEO_ID} hoặc {URL}.";
      return result;
    }

    const resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "X-RapidAPI-Key": apiKey,
        "X-RapidAPI-Host": host
      },
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);
      const text = transcriptTextFromJson_(json);
      if (text) {
        result.success = true;
        result.text = text;
        result.language = json.language || json.languageCode || json.lang || "unknown";
        return result;
      }
      result.error = "RapidAPI trả JSON nhưng không tìm thấy nội dung transcript.";
      return result;
    }
    result.error = "HTTP " + code + ": " + body.slice(0, 180);
    return result;
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

function fetchTranscriptViaApify(videoId, token) {
  const result = { success: false, text: "", language: "", error: "" };
  try {
    const url = "https://api.apify.com/v2/acts/akash9078~youtube-transcript-extractor/run-sync-get-dataset-items";
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + token },
      payload: JSON.stringify({
        videoUrl: "https://www.youtube.com/watch?v=" + videoId
      }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);
      const text = transcriptTextFromJson_(json);
      if (text) {
        result.success = true;
        result.text = text;
        result.language = (Array.isArray(json) && json[0] && (json[0].language || json[0].language_code)) || "unknown";
        return result;
      }
      result.error = "Apify chạy xong nhưng dataset không có transcript.";
      return result;
    }
    result.error = "HTTP " + code + ": " + body.slice(0, 180);
    return result;
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

function fetchTranscriptViaYoutubeTranscriptApiBridge(videoId, bridgeUrl) {
  const result = { success: false, text: "", language: "", error: "" };
  try {
    const url = bridgeUrl.toString().trim();
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ videoId: videoId, languages: ["vi", "en"] }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);
      const text = transcriptTextFromJson_(json);
      if (text) {
        result.success = true;
        result.text = text;
        result.language = json.language || json.language_code || json.lang || "unknown";
        return result;
      }
      result.error = "Bridge trả JSON nhưng không có transcript.";
      return result;
    }
    result.error = "HTTP " + code + ": " + body.slice(0, 180);
    return result;
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

function buildAssemblyAIAudioUrl_(videoId, template) {
  const videoUrl = "https://www.youtube.com/watch?v=" + videoId;
  const audioUrl = template.toString().trim()
    .replace(/\{VIDEO_ID\}/g, encodeURIComponent(videoId))
    .replace(/\{URL\}/g, encodeURIComponent(videoUrl));
  return /^https?:\/\//i.test(audioUrl) && audioUrl.indexOf("{") < 0 ? audioUrl : "";
}

function fetchTranscriptViaAssemblyAI(videoId, apiKey, audioUrlTemplate) {
  const result = { success: false, text: "", language: "", error: "" };
  try {
    const audioUrl = buildAssemblyAIAudioUrl_(videoId, audioUrlTemplate);
    if (!audioUrl) {
      result.error = "B13 phải là direct audio/video URL công khai hoặc template hợp lệ.";
      return result;
    }
    const createResp = UrlFetchApp.fetch("https://api.assemblyai.com/v2/transcript", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": apiKey },
      payload: JSON.stringify({
        audio_url: audioUrl,
        speech_models: ["universal-3-pro", "universal-2"],
        language_detection: true
      }),
      muteHttpExceptions: true
    });
    const createCode = createResp.getResponseCode();
    const createBody = createResp.getContentText();
    if (createCode < 200 || createCode >= 300) {
      result.error = "Submit HTTP " + createCode + ": " + createBody.slice(0, 180);
      return result;
    }
    const created = JSON.parse(createBody);
    if (!created.id) {
      result.error = "AssemblyAI không trả transcript id.";
      return result;
    }
    for (let i = 0; i < 20; i++) {
      Utilities.sleep(3000);
      const pollResp = UrlFetchApp.fetch("https://api.assemblyai.com/v2/transcript/" + created.id, {
        method: "get",
        headers: { "Authorization": apiKey },
        muteHttpExceptions: true
      });
      const pollCode = pollResp.getResponseCode();
      const pollBody = pollResp.getContentText();
      if (pollCode < 200 || pollCode >= 300) {
        result.error = "Poll HTTP " + pollCode + ": " + pollBody.slice(0, 180);
        return result;
      }
      const item = JSON.parse(pollBody);
      if (item.status === "completed" && item.text) {
        result.success = true;
        result.text = normalizeTranscriptText_(item.text);
        result.language = item.language_code || "unknown";
        return result;
      }
      if (item.status === "error") {
        result.error = item.error || "AssemblyAI xử lý lỗi.";
        return result;
      }
    }
    result.error = "Timeout chờ AssemblyAI hoàn tất.";
    return result;
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

// ═══════════════════════════════════════════════════════════════
// LỚP 3: TIMEDTEXT API TRỰC TIẾP
// ═══════════════════════════════════════════════════════════════
function fetchTranscriptViaTimedText(videoId) {
  const result = { success: false, text: "", language: "", error: "" };
  
  // Mo rong lang list - bao gom nhieu ngon ngu pho bien
  const langs = ['vi', 'en', 'en-US', 'en-GB', 'es', 'pt', 'id', 'ko', 'ja', 'fr', 'de', 'zh-Hans', 'zh-Hant', 'ar', 'hi', 'ru', 'th'];
  
  const headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept-Language': 'vi-VN,vi;q=0.9,en;q=0.8',
    'Accept': 'text/xml,application/xml,*/*'
  };
  
  const endpoints = [
    'https://www.youtube.com/api/timedtext',
    'https://video.google.com/timedtext'
  ];
  
  for (let endpoint of endpoints) {
    for (let lang of langs) {
      try {
        const variants = [
          `?lang=${lang}&v=${videoId}`,
          `?lang=${lang}&v=${videoId}&fmt=srv3`,
          `?lang=${lang}&v=${videoId}&kind=asr`,
          `?lang=${lang}&v=${videoId}&kind=asr&fmt=srv3`,
        ];
        
        for (let variant of variants) {
          const url = endpoint + variant;
          const resp = UrlFetchApp.fetch(url, { 
            muteHttpExceptions: true, 
            headers: headers 
          });
          
          if (resp.getResponseCode() === 200) {
            const xml = resp.getContentText();
            if (xml && xml.trim() && xml.includes('<text')) {
              const parsed = parseTranscriptXml(xml);
              if (parsed && parsed.length > 10) {
                result.success = true;
                result.text = parsed;
                result.language = lang;
                return result;
              }
            }
          }
        }
      } catch (e) {
        continue;
      }
    }
  }

  // Fallback: thu timedtext voi tlang=en de dich tu dong bat ky ngon ngu nao
  try {
    const tlangUrl = `https://www.youtube.com/api/timedtext?v=${videoId}&tlang=en&kind=asr`;
    const tlangResp = UrlFetchApp.fetch(tlangUrl, { muteHttpExceptions: true, headers: headers });
    if (tlangResp.getResponseCode() === 200) {
      const xml = tlangResp.getContentText();
      if (xml && xml.includes('<text')) {
        const parsed = parseTranscriptXml(xml);
        if (parsed && parsed.length > 10) {
          result.success = true;
          result.text = parsed;
          result.language = 'auto-en';
          return result;
        }
      }
    }
  } catch(e) {}
  
  result.error = "Không có timedtext khả dụng cho video";
  return result;

// =============================================================
// LOP 4b: INNERTUBE API (youtube.com/youtubei/v1/player)
// Nguon manh nhat - lay captionTracks tu API chinh thuc YouTube
// =============================================================
function fetchTranscriptViaInnerTube(videoId) {
  const result = { success: false, text: "", language: "", error: "" };
  try {
    const playerPayload = {
      videoId: videoId,
      context: {
        client: {
          clientName: "WEB",
          clientVersion: "2.20240101.00.00",
          hl: "en",
          gl: "US"
        }
      }
    };
    const playerResp = UrlFetchApp.fetch("https://www.youtube.com/youtubei/v1/player?key=AIzaSyAO_FJ2SlqU8Q4STEHLGCilw_Y9_11qcW8", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(playerPayload),
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Origin": "https://www.youtube.com",
        "Referer": "https://www.youtube.com/watch?v=" + videoId
      }
    });
    if (playerResp.getResponseCode() !== 200) {
      result.error = "InnerTube player HTTP " + playerResp.getResponseCode();
      return result;
    }
    const playerJson = JSON.parse(playerResp.getContentText());
    const captions = playerJson &&
                     playerJson.captions &&
                     playerJson.captions.playerCaptionsTracklistRenderer &&
                     playerJson.captions.playerCaptionsTracklistRenderer.captionTracks;
    if (!captions || captions.length === 0) {
      result.error = "InnerTube: Video khong co caption track";
      return result;
    }
    const priorityLangs = ["vi", "en", "en-US", "en-GB", "es", "pt", "id", "ko", "ja"];
    let chosenTrack = null;
    for (let lang of priorityLangs) {
      chosenTrack = captions.find(function(t) { return t.languageCode === lang; });
      if (chosenTrack) break;
    }
    if (!chosenTrack) chosenTrack = captions[0];
    let baseUrl = chosenTrack.baseUrl;
    if (!baseUrl) {
      result.error = "InnerTube: Track khong co baseUrl";
      return result;
    }
    baseUrl = baseUrl + (baseUrl.includes("?") ? "&" : "?") + "fmt=srv3";
    const transcriptResp = UrlFetchApp.fetch(baseUrl, {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://www.youtube.com/watch?v=" + videoId
      }
    });
    if (transcriptResp.getResponseCode() !== 200) {
      result.error = "InnerTube: HTTP " + transcriptResp.getResponseCode() + " khi tai transcript";
      return result;
    }
    const xml = transcriptResp.getContentText();
    if (!xml || !xml.includes("<")) {
      result.error = "InnerTube: Transcript XML rong";
      return result;
    }
    const parsed = parseTranscriptXml(xml);
    if (!parsed || parsed.length < 10) {
      result.error = "InnerTube: Transcript parse ra rong";
      return result;
    }
    result.success = true;
    result.text = parsed;
    result.language = chosenTrack.languageCode || "unknown";
    return result;
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

}

// ═══════════════════════════════════════════════════════════════
// LỚP 4: SCRAPE TRANG WATCH
// ═══════════════════════════════════════════════════════════════
function fetchTranscriptViaScrape(videoId) {
  const result = { success: false, text: "", language: "", error: "" };
  
  try {
    const watchUrl = `https://www.youtube.com/watch?v=${videoId}&hl=vi`;
    const watchResp = UrlFetchApp.fetch(watchUrl, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Accept-Language': 'vi-VN,vi;q=0.9,en;q=0.8',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9'
      }
    });
    
    const code = watchResp.getResponseCode();
    if (code === 429) {
      result.error = "IP bị YouTube chặn (HTTP 429) - cấu hình API key Supadata để fix";
      return result;
    }
    if (code !== 200) {
      result.error = "HTTP " + code;
      return result;
    }
    
    const html = watchResp.getContentText();
    
    if (html.includes('"status":"ERROR"') || html.includes('Video unavailable')) {
      result.error = "Video bị xóa/private";
      return result;
    }
    
    // Regex nang cap: bat captionTracks JSON day du hon, tranh bi cat som
    let captionMatch = html.match(/"captionTracks":(\[\{[\s\S]*?\}\])/);
    if (!captionMatch) {
      // Thu regex don gian hon de bat truong hop JSON ngan
      captionMatch = html.match(/"captionTracks":(\[.*?\])/s);
    }
    if (!captionMatch) {
      result.error = "Không có captionTracks (video không bật phụ đề)";
      return result;
    }
    
    let tracks;
    try {
      tracks = JSON.parse(captionMatch[1]);
    } catch(e) {
      result.error = "Parse captionTracks lỗi";
      return result;
    }
    
    if (!tracks || tracks.length === 0) {
      result.error = "Tracks trống";
      return result;
    }
    
    const priorityLangs = ['vi', 'en', 'en-US', 'en-GB'];
    let chosenTrack = null;
    for (let lang of priorityLangs) {
      chosenTrack = tracks.find(t => t.languageCode === lang);
      if (chosenTrack) break;
    }
    if (!chosenTrack) chosenTrack = tracks[0];
    
    let baseUrl = chosenTrack.baseUrl;
    if (!baseUrl) {
      result.error = "Track không có baseUrl";
      return result;
    }
    baseUrl = baseUrl.replace(/\\u0026/g, '&').replace(/\\\//g, '/');
    
    const transcriptResp = UrlFetchApp.fetch(baseUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });
    
    if (transcriptResp.getResponseCode() !== 200) {
      result.error = "HTTP " + transcriptResp.getResponseCode() + " khi tải transcript";
      return result;
    }
    
    const xml = transcriptResp.getContentText();
    if (!xml || xml.trim() === "") {
      result.error = "Transcript rỗng";
      return result;
    }
    
    const transcript = parseTranscriptXml(xml);
    if (!transcript || transcript.trim() === "") {
      result.error = "Transcript parse rỗng";
      return result;
    }
    
    result.success = true;
    result.text = transcript;
    result.language = chosenTrack.languageCode || "unknown";
    return result;
    
  } catch (e) {
    result.error = e.message;
    return result;
  }
}

/**
 * Parse XML transcript của YouTube thành plain text
 * Format: <text start="0.0" dur="2.5">Hello world</text>
 */
function parseTranscriptXml(xml) {
  if (!xml) return "";
  
  const textMatches = xml.match(/<text[^>]*>([\s\S]*?)<\/text>/g);
  if (!textMatches) return "";
  
  const lines = [];
  for (let m of textMatches) {
    const inner = m.replace(/<text[^>]*>/, '').replace(/<\/text>/, '');
    const decoded = decodeHtmlEntities(inner)
      .replace(/\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (decoded) lines.push(decoded);
  }
  
  return lines.join(' ');
}

/**
 * Decode HTML entities trong transcript
 */
function decodeHtmlEntities(str) {
  if (!str) return "";
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, function(m, code) {
      return String.fromCharCode(parseInt(code));
    })
    .replace(/&#x([0-9a-fA-F]+);/g, function(m, code) {
      return String.fromCharCode(parseInt(code, 16));
    });
}

/**
 * 📥 Lấy Subtitle TẤT CẢ video
 */
function fetchAllSubtitles() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIDEO);
  if (!sheet) return ui.alert("Không tìm thấy Sheet: " + SHEET_VIDEO);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ui.alert("Sheet không có dữ liệu.");
  
  const confirmRes = ui.alert(
    '📝 Xác nhận lấy Subtitle',
    `Bạn sắp lấy subtitle cho ${lastRow - 1} video.\n\n⚠️ Quá trình này có thể mất vài phút (mỗi video ~2-3 giây).\n\nKết quả lưu vào Cột M.\n\nTiếp tục?`,
    ui.ButtonSet.YES_NO
  );
  if (confirmRes !== ui.Button.YES) return;
  
  executeSubtitleFetch(sheet, 2, lastRow);
}

/**
 * 📥 Lấy Subtitle THEO KHOẢNG DÒNG
 */
function fetchRangeSubtitles() {
  openRangeActionDialog_("subtitleRange", SHEET_VIDEO, "Lấy Subtitle theo KHOẢNG DÒNG", "Chọn khoảng dòng trên Sheet VIDEO để lấy transcript/subtitle.", "#0f9d58");
}

/**
 * 📥 Lấy Subtitle THEO 1 DÒNG
 */
function fetchSpecificSubtitle() {
  openSingleRowActionDialog_("subtitleRow", SHEET_VIDEO, "Lấy Subtitle theo dòng", "Nhập một dòng trên Sheet VIDEO để lấy transcript/subtitle.", "#0f9d58");
}

/**
 * Lõi xử lý lấy subtitle - dùng chung cho 3 hàm trên
 * 🆕 Phiên bản nâng cấp:
 *   - Tự động chọn nguồn phù hợp (Supadata → YT-Transcript.io → Timedtext → Scrape)
 *   - Sleep thông minh: API có key chỉ cần 300ms, scrape phải 2000ms
 *   - Hiển thị nguồn lấy được trong báo cáo
 *   - Phát hiện rate-limit để dừng kịp thời
 */
function executeSubtitleFetch(sheet, startRow, endRow, suppressAlert, progressId) {
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet.getRange(1, COL_VIDEO.SUBTITLE).getValue()) {
    sheet.getRange(1, COL_VIDEO.SUBTITLE).setValue("📝 SUBTITLE")
      .setFontWeight("bold").setBackground("#0f9d58").setFontColor("white");
  }
  
  let hasSupadataKey = false;
  let hasYTIoToken = false;
  let hasPremiumTranscriptSource = false;
  try { hasSupadataKey = !!getSupadataKey(); } catch(e) {}
  try { hasYTIoToken = !!getYTTranscriptIoToken(); } catch(e) {}
  try { hasPremiumTranscriptSource = !!(getRapidApiTranscriptKey() && getRapidApiTranscriptHost() && getRapidApiTranscriptEndpoint()) || !!getApifyToken() || !!getYTTranscriptApiBridgeUrl() || !!(getAssemblyAIApiKey() && getAssemblyAIAudioUrlTemplate()); } catch(e) {}
  
  if (!hasSupadataKey && !hasYTIoToken && !hasPremiumTranscriptSource) {
    const warning = ui.alert(
      '⚠️ CẢNH BÁO QUAN TRỌNG',
      'Bạn CHƯA cấu hình API key cho Subtitle (B5/B6 trong Sheet "API KEY")!\n\n' +
      'Hệ thống sẽ dùng phương pháp scrape trực tiếp - HAY BỊ LỖI 429 (YouTube chặn IP của Apps Script).\n\n' +
      'KHUYẾN NGHỊ: Cấu hình ít nhất 1 nguồn API như Supadata, Apify, RapidAPI, AssemblyAI hoặc bridge youtube-transcript-api\n' +
      '(Xem hướng dẫn tại Menu → 📝 SUBTITLE → ℹ️ Hướng dẫn cấu hình API)\n\n' +
      'Vẫn tiếp tục với phương pháp cũ?',
      ui.ButtonSet.YES_NO
    );
    if (warning !== ui.Button.YES) return;
  }
  
  const numRows = endRow - startRow + 1;
  const linkRange = sheet.getRange(startRow, COL_VIDEO.LINK, numRows, 1).getValues();
  
  let successCount = 0;
  let failCount = 0;
  let skipCount = 0;
  let consecutive429 = 0;
  const failedRows = [];
  const sourceCount = {};
  
  const sleepMs = (hasSupadataKey || hasPremiumTranscriptSource) ? 300 : (hasYTIoToken ? 500 : 2500);

  // ── BATCH WRITE BUFFER ────────────────────────────────────────────────
  // subtitleBatch[i] = giá trị cần ghi vào cột SUBTITLE tại linkRange index i
  // Dùng null làm sentinel "chưa thay đổi" để tránh ghi đè dòng skip
  const subtitleBatch = new Array(linkRange.length).fill(null);
  let earlyStop = false;       // flag dừng sớm khi 429
  let earlyStopRow = -1;       // dòng ngay sau điểm dừng để ghi warning
  // ─────────────────────────────────────────────────────────────────────
  
  for (let i = 0; i < linkRange.length; i++) {
    const link = linkRange[i][0];
    const actualRow = startRow + i;
    
    if (!link || link.toString().trim() === "") {
      skipCount++;
      updateTaskProgress_(progressId, i + 1, linkRange.length, "Đã xử lý subtitle " + (i + 1) + "/" + linkRange.length);
      continue;
    }
    
    const videoId = extractVideoId(link);
    if (!videoId) {
      failCount++;
      failedRows.push({ row: actualRow, reason: "Link sai định dạng" });
      subtitleBatch[i] = "[LỖI: Link sai định dạng]"; // ✅ Buffer
      updateTaskProgress_(progressId, i + 1, linkRange.length, "Đã xử lý subtitle " + (i + 1) + "/" + linkRange.length);
      continue;
    }
    
    try {
      const transcript = getYouTubeTranscript(videoId);
      
      if (transcript.success && transcript.text) {
        let finalText = transcript.text;
        if (finalText.length > 49500) {
          finalText = finalText.substring(0, 49500) + "...[truncated]";
        }
        
        subtitleBatch[i] = finalText; // ✅ Buffer thay vì ghi ngay
        successCount++;
        consecutive429 = 0;
        
        const src = transcript.source || "unknown";
        sourceCount[src] = (sourceCount[src] || 0) + 1;
        
      } else {
        failCount++;
        const errMsg = transcript.error || "Không rõ lý do";
        failedRows.push({ row: actualRow, reason: errMsg });
        subtitleBatch[i] = "[KHÔNG CÓ SUBTITLE: " + errMsg + "]"; // ✅ Buffer
        
        if (errMsg.includes("429") || errMsg.includes("rate limit") || errMsg.includes("IP bị YouTube chặn")) {
          consecutive429++;
        } else {
          consecutive429 = 0;
        }
        
        if (consecutive429 >= 5) {
          // Ghi cảnh báo vào dòng kế tiếp (nếu còn trong range)
          if (i + 1 < linkRange.length) {
            earlyStopRow = i + 1;
          }
          earlyStop = true;
          // Ghi batch đã tích lũy trước khi alert để user thấy tiến độ.
          writeSparseColumnValues_(sheet, startRow, COL_VIDEO.SUBTITLE, subtitleBatch.slice(0, i + 1));
          if (earlyStopRow >= 0) {
            sheet.getRange(startRow + earlyStopRow, COL_VIDEO.SUBTITLE).setValue("[ĐÃ DỪNG: Quá nhiều lỗi 429 liên tiếp]");
          }
          SpreadsheetApp.flush();
          ui.alert(
            '🛑 ĐÃ TỰ ĐỘNG DỪNG',
            'Phát hiện ' + consecutive429 + ' lỗi 429 liên tiếp.\n\n' +
            'YouTube đã tạm chặn IP của Apps Script. Để fix triệt để:\n' +
            '1. Cấu hình Supadata / Apify / RapidAPI / bridge youtube-transcript-api\n' +
            '2. Điền key vào Sheet "API KEY" theo hướng dẫn\n' +
            '3. Thử lại\n\n' +
            'Đã xử lý: ' + (i + 1) + '/' + linkRange.length + ' video.',
            ui.ButtonSet.OK
          );
          break;
        }
      }
    } catch (e) {
      failCount++;
      failedRows.push({ row: actualRow, reason: e.message });
      subtitleBatch[i] = "[LỖI: " + e.message + "]"; // ✅ Buffer
    }
    
    // ✅ Đã bỏ SpreadsheetApp.flush() thừa trong vòng lặp
    Utilities.sleep(sleepMs);
    updateTaskProgress_(progressId, i + 1, linkRange.length, "Đã xử lý subtitle " + (i + 1) + "/" + linkRange.length);
  }

  // ── BATCH WRITE: ghi tất cả kết quả ra sheet 1 lần ───────────────────
  if (!earlyStop) {
    // Ghi tất cả dòng đã xử lý theo block liên tiếp (bỏ qua null = dòng skip)
    writeSparseColumnValues_(sheet, startRow, COL_VIDEO.SUBTITLE, subtitleBatch);
    SpreadsheetApp.flush(); // ✅ Chỉ flush 1 lần duy nhất
  }
  // (nếu earlyStop thì đã flush bên trong vòng lặp tại điểm dừng)
  // ─────────────────────────────────────────────────────────────────────
  
  let report = `📊 BÁO CÁO LẤY SUBTITLE:\n\n`;
  report += `✅ Thành công: ${successCount}\n`;
  report += `❌ Thất bại: ${failCount}\n`;
  report += `⏭️ Bỏ qua (dòng trống): ${skipCount}\n`;
  report += `📍 Tổng: ${successCount + failCount + skipCount}\n\n`;
  
  if (Object.keys(sourceCount).length > 0) {
    report += `🔧 NGUỒN LẤY ĐƯỢC:\n`;
    for (let src in sourceCount) {
      report += `• ${src}: ${sourceCount[src]} video\n`;
    }
    report += '\n';
  }
  
  if (failedRows.length > 0) {
    report += `🔍 CHI TIẾT VIDEO KHÔNG LẤY ĐƯỢC:\n`;
    const showLimit = Math.min(failedRows.length, 15);
    for (let i = 0; i < showLimit; i++) {
      let reason = failedRows[i].reason;
      if (reason.length > 100) reason = reason.substring(0, 100) + "...";
      report += `• Dòng ${failedRows[i].row}: ${reason}\n`;
    }
    if (failedRows.length > 15) {
      report += `... và ${failedRows.length - 15} dòng khác (xem chi tiết tại Cột M)\n`;
    }
  }
  
  if (!suppressAlert) ui.alert(report);
  return report;
}

/**
 * 🆕 Hiển thị hướng dẫn cấu hình API key cho Subtitle
 */
function showSubtitleApiGuide() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial; padding: 15px; line-height: 1.6;">
      ${buildDialogHeader_("HƯỚNG DẪN CẤU HÌNH API SUBTITLE", "Cấu hình nguồn lấy transcript ổn định hơn cho hệ thống YouTube Tools.", "#0f9d58")}
      <div id="subtitleGuideBody" class="dlgFadeBody" style="max-height:430px;overflow:auto;padding-right:6px;">
      <p style="background:#fff3cd; padding:10px; border-left:4px solid #f29900; font-size:13px; margin: 10px 0;">
        <b>⚠️ Vì sao cần cấu hình?</b><br>
        Apps Script chạy chung IP với hàng triệu người → YouTube hay chặn (HTTP 429).<br>
        Cấu hình API miễn phí để có tỷ lệ thành công ~99%.
      </p>
      
      <h3 style="color:#4285F4; margin-top: 20px;">🌟 LỰA CHỌN 1: SUPADATA.AI (KHUYẾN NGHỊ)</h3>
      <ul style="font-size:13px;">
        <li>✅ <b>Miễn phí 100 transcript/tháng</b> - không cần thẻ tín dụng</li>
        <li>✅ Hỗ trợ AI fallback nếu video không có sub</li>
        <li>✅ Hỗ trợ trên 100+ ngôn ngữ</li>
      </ul>
      <ol style="font-size:13px;">
        <li>Truy cập: <a href="https://supadata.ai" target="_blank"><b>https://supadata.ai</b></a></li>
        <li>Đăng ký bằng Google (30 giây)</li>
        <li>Vào Dashboard → Copy API Key</li>
        <li>Dán vào ô <b>B5</b> tại Sheet "<b>API KEY</b>"</li>
      </ol>
      
      <h3 style="color:#4285F4; margin-top: 20px;">🌟 LỰA CHỌN 2: YOUTUBE-TRANSCRIPT.IO</h3>
      <ul style="font-size:13px;">
        <li>✅ Miễn phí khi đăng ký tài khoản</li>
        <li>✅ Hỗ trợ batch 50 video/lần</li>
      </ul>
      <ol style="font-size:13px;">
        <li>Truy cập: <a href="https://www.youtube-transcript.io" target="_blank"><b>https://www.youtube-transcript.io</b></a></li>
        <li>Tạo tài khoản → Profile → Copy API Token</li>
        <li>Dán vào ô <b>B6</b> tại Sheet "<b>API KEY</b>"</li>
      </ol>
      
      <h3 style="color:#4285F4; margin-top: 20px;">🌟 LỰA CHỌN 3: RAPIDAPI - YOUTUBE TRANSCRIPT</h3>
      <ul style="font-size:13px;">
        <li>Nguồn: <a href="https://rapidapi.com/thisisgazzar/api/youtube-transcript1" target="_blank"><b>rapidapi.com/thisisgazzar/api/youtube-transcript1</b></a></li>
        <li>Điền <b>B8</b> = X-RapidAPI-Key, <b>B9</b> = X-RapidAPI-Host, <b>B10</b> = endpoint từ tab Endpoints.</li>
        <li>Endpoint nên chứa <code>{VIDEO_ID}</code> hoặc <code>{URL}</code>; hệ thống sẽ tự thay bằng video hiện tại.</li>
        <li><b>Lưu ý:</b> RapidAPI không expose endpoint trong HTML đọc tự động ở đây, nên phải copy đúng từ nguồn RapidAPI để tránh bịa đường dẫn.</li>
      </ul>

      <h3 style="color:#4285F4; margin-top: 20px;">🌟 LỰA CHỌN 4: APIFY ACTOR</h3>
      <ul style="font-size:13px;">
        <li>Nguồn: <a href="https://apify.com/akash9078/youtube-transcript-extractor" target="_blank"><b>akash9078/youtube-transcript-extractor</b></a></li>
        <li>Docs Apify API v2: <a href="https://docs.apify.com/api/v2" target="_blank"><b>docs.apify.com/api/v2</b></a></li>
        <li>Điền <b>B11</b> = Apify API token. Hệ thống gọi Actor bằng endpoint đồng bộ <code>run-sync-get-dataset-items</code>.</li>
        <li>Theo trang Actor, input chính là <code>videoUrl</code> và có thể truyền <code>language</code>.</li>
      </ul>

      <h3 style="color:#4285F4; margin-top: 20px;">🌟 LỰA CHỌN 5: ASSEMBLYAI</h3>
      <ul style="font-size:13px;">
        <li>Nguồn docs: <a href="https://www.assemblyai.com/docs/api-reference/overview" target="_blank"><b>AssemblyAI API Reference</b></a></li>
        <li>Điền <b>B12</b> = API key, <b>B13</b> = URL audio/video công khai trực tiếp hoặc template có <code>{VIDEO_ID}</code>/<code>{URL}</code>.</li>
        <li><b>Quan trọng:</b> AssemblyAI transcribe media file bằng <code>audio_url</code>; docs nêu lỗi thường gặp khi URL là webpage thay vì file audio/video. Vì vậy không dùng trực tiếp link YouTube watch nếu không phải media file.</li>
        <li>Hệ thống submit <code>/v2/transcript</code>, poll <code>/v2/transcript/{id}</code>, dùng <code>speech_models: ["universal-3-pro","universal-2"]</code>.</li>
      </ul>

      <h3 style="color:#4285F4; margin-top: 20px;">🌟 LỰA CHỌN 6: PYTHON youtube-transcript-api</h3>
      <ul style="font-size:13px;">
        <li>Nguồn: <a href="https://pypi.org/project/youtube-transcript-api/" target="_blank"><b>PyPI youtube-transcript-api</b></a> / <a href="https://github.com/jdepoix/youtube-transcript-api" target="_blank"><b>GitHub</b></a></li>
        <li>Apps Script không chạy Python trực tiếp, nên cần deploy bridge riêng (Cloud Run, VPS, Apps Script Web App trung gian gọi server Python...).</li>
        <li>Điền <b>B14</b> = URL bridge. Hệ thống POST JSON: <code>{"videoId":"...","languages":["vi","en"]}</code>.</li>
        <li>Bridge nên trả JSON có <code>text</code> hoặc <code>transcript</code> hoặc mảng segment <code>{text,start,duration}</code>.</li>
      </ul>

      <p style="background:#e8f5e9; padding:10px; border-left:4px solid #0f9d58; font-size:13px; margin: 15px 0;">
        <b>💡 THỨ TỰ FALLBACK:</b><br>
        Supadata → YouTube-Transcript.io → RapidAPI → Apify → youtube-transcript-api bridge → AssemblyAI → timedtext → InnerTube → scrape watch page.<br>
        Cấu hình càng nhiều nguồn thì tỷ lệ thành công càng cao, nhưng vẫn phụ thuộc video có caption, quyền truy cập, quota và trạng thái API từng nhà cung cấp.
      </p>
      
      <h3 style="color:#d32f2f; margin-top: 20px;">📋 BỐ CỤC SHEET "API KEY":</h3>
      <table style="border-collapse: collapse; width: 100%; font-size:13px;">
        <tr style="background: #4285F4; color:white;">
          <th style="border:1px solid #ccc; padding:6px;">Ô</th>
          <th style="border:1px solid #ccc; padding:6px;">Nội dung</th>
        </tr>
        <tr><td style="border:1px solid #ccc; padding:6px;"><b>B2</b></td><td style="border:1px solid #ccc; padding:6px;">YouTube Data API Key (bắt buộc)</td></tr>
        <tr><td style="border:1px solid #ccc; padding:6px;"><b>B3</b></td><td style="border:1px solid #ccc; padding:6px;">Google OAuth Client ID</td></tr>
        <tr><td style="border:1px solid #ccc; padding:6px;"><b>B4</b></td><td style="border:1px solid #ccc; padding:6px;">Google OAuth Client Secret</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B5</b></td><td style="border:1px solid #ccc; padding:6px;"><b>🆕 Supadata API Key</b> (cho Subtitle)</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B6</b></td><td style="border:1px solid #ccc; padding:6px;"><b>🆕 YouTube-Transcript.io Token</b> (cho Subtitle)</td></tr>
        <tr style="background:#eef2ff;"><td style="border:1px solid #ccc; padding:6px;"><b>B7</b></td><td style="border:1px solid #ccc; padding:6px;"><b>9router API Key / Bearer Token</b> (cho AI phân tích Sheet)</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B8</b></td><td style="border:1px solid #ccc; padding:6px;">RapidAPI YouTube Transcript Key</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B9</b></td><td style="border:1px solid #ccc; padding:6px;">RapidAPI YouTube Transcript Host</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B10</b></td><td style="border:1px solid #ccc; padding:6px;">RapidAPI endpoint, hỗ trợ <code>{VIDEO_ID}</code>/<code>{URL}</code></td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B11</b></td><td style="border:1px solid #ccc; padding:6px;">Apify API Token</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B12</b></td><td style="border:1px solid #ccc; padding:6px;">AssemblyAI API Key</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B13</b></td><td style="border:1px solid #ccc; padding:6px;">AssemblyAI direct audio/video URL template</td></tr>
        <tr style="background:#fff3cd;"><td style="border:1px solid #ccc; padding:6px;"><b>B14</b></td><td style="border:1px solid #ccc; padding:6px;">youtube-transcript-api Bridge URL</td></tr>
      </table>
      
      </div>
      <div style="text-align:center; margin-top:16px;">
        <button onclick="google.script.host.close()" style="padding:10px 30px; background:linear-gradient(135deg,#4285F4,#3367d6); color:white; border:none; border-radius:8px; cursor:pointer; font-weight:bold; box-shadow:0 4px 12px rgba(66,133,244,.3);">ĐÓNG</button>
      </div>
      <script>
        setTimeout(function(){ if(window.dlgProgress) window.dlgProgress.attachScroll('subtitleGuideBody'); }, 60);
        document.addEventListener('keydown', function(e){ if(e.key==='Escape') google.script.host.close(); });
      </script>
    </div>
  `).setWidth(600).setHeight(720);
  ui.showModalDialog(html, '📝 Hướng dẫn cấu hình API Subtitle');
}

/**
 * 🔄 Thử lại CHỈ các dòng đã FAIL trước đó
 * Quét cột M, dòng nào có dấu "[KHÔNG CÓ SUBTITLE..." hoặc "[LỖI..." → thử lại
 */
function retryFailedSubtitles() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIDEO);
  if (!sheet) return ui.alert("Không tìm thấy Sheet: " + SHEET_VIDEO);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ui.alert("Sheet trống.");
  
  const subtitleData = sheet.getRange(2, COL_VIDEO.SUBTITLE, lastRow - 1, 1).getValues();
  const failedRows = [];
  
  for (let i = 0; i < subtitleData.length; i++) {
    const cellValue = subtitleData[i][0];
    if (cellValue && typeof cellValue === 'string') {
      if (cellValue.startsWith("[KHÔNG CÓ SUBTITLE") || 
          cellValue.startsWith("[LỖI") || 
          cellValue.startsWith("[ĐÃ DỪNG")) {
        failedRows.push(i + 2);
      }
    }
  }
  
  if (failedRows.length === 0) {
    return ui.alert("✅ Không có dòng nào FAIL cần thử lại.\n\nTất cả các ô ở Cột M đều có dữ liệu hợp lệ.");
  }
  
  const confirm = ui.alert(
    '🔄 Thử lại các dòng FAIL',
    `Phát hiện ${failedRows.length} dòng có lỗi cần thử lại:\n` +
    `Dòng: ${failedRows.slice(0, 10).join(', ')}${failedRows.length > 10 ? '...' : ''}\n\n` +
    'Tiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  
  if (!sheet.getRange(1, COL_VIDEO.SUBTITLE).getValue()) {
    sheet.getRange(1, COL_VIDEO.SUBTITLE).setValue("📝 SUBTITLE")
      .setFontWeight("bold").setBackground("#0f9d58").setFontColor("white");
  }
  
  let hasSupadataKey = false;
  let hasYTIoToken = false;
  let hasPremiumTranscriptSource = false;
  try { hasSupadataKey = !!getSupadataKey(); } catch(e) {}
  try { hasYTIoToken = !!getYTTranscriptIoToken(); } catch(e) {}
  try { hasPremiumTranscriptSource = !!(getRapidApiTranscriptKey() && getRapidApiTranscriptHost() && getRapidApiTranscriptEndpoint()) || !!getApifyToken() || !!getYTTranscriptApiBridgeUrl() || !!(getAssemblyAIApiKey() && getAssemblyAIAudioUrlTemplate()); } catch(e) {}
  const sleepMs = (hasSupadataKey || hasPremiumTranscriptSource) ? 300 : (hasYTIoToken ? 500 : 2500);
  
  let successCount = 0;
  let stillFailCount = 0;
  const sourceCount = {};
  const stillFailedRows = [];

  // ── BATCH READ links cho tất cả failedRows ────────────────────────────
  // Dùng lại lastRow đã khai báo ở đầu hàm (tránh redeclaration conflict)
  const allLinks = lastRow >= 2
    ? sheet.getRange(2, COL_VIDEO.LINK, lastRow - 1, 1).getValues()
    : [];
  // Mảng buffer kết quả: rowNum -> value cần ghi
  const retryBatch = {}; // key = actualRow, value = string
  // ─────────────────────────────────────────────────────────────────────
  
  for (let actualRow of failedRows) {
    // Đọc từ allLinks đã tải sẵn thay vì getRange từng dòng
    const linkIdx = actualRow - 2;
    const link = (linkIdx >= 0 && linkIdx < allLinks.length) ? allLinks[linkIdx][0] : null;
    if (!link) continue;
    
    const videoId = extractVideoId(link);
    if (!videoId) {
      stillFailCount++;
      stillFailedRows.push({ row: actualRow, reason: "Link sai định dạng" });
      continue;
    }
    
    try {
      const transcript = getYouTubeTranscript(videoId);
      if (transcript.success && transcript.text) {
        let finalText = transcript.text;
        if (finalText.length > 49500) finalText = finalText.substring(0, 49500) + "...[truncated]";
        retryBatch[actualRow] = finalText; // ✅ Buffer
        successCount++;
        const src = transcript.source || "unknown";
        sourceCount[src] = (sourceCount[src] || 0) + 1;
      } else {
        stillFailCount++;
        stillFailedRows.push({ row: actualRow, reason: transcript.error });
        retryBatch[actualRow] = "[KHÔNG CÓ SUBTITLE: " + transcript.error + "]"; // ✅ Buffer
      }
    } catch (e) {
      stillFailCount++;
      stillFailedRows.push({ row: actualRow, reason: e.message });
      retryBatch[actualRow] = "[LỖI retry: " + e.message + "]"; // ✅ Buffer
    }
    
    // ✅ Đã bỏ SpreadsheetApp.flush() thừa trong vòng lặp
    Utilities.sleep(sleepMs);
  }

  // ── BATCH WRITE: ghi tất cả kết quả retry 1 lần ──────────────────────
  writeRowValueMapToColumn_(sheet, COL_VIDEO.SUBTITLE, retryBatch);
  if (Object.keys(retryBatch).length > 0) SpreadsheetApp.flush(); // ✅ 1 flush duy nhất
  // ─────────────────────────────────────────────────────────────────────
  
  let report = `🔄 BÁO CÁO THỬ LẠI:\n\n`;
  report += `✅ Khắc phục thành công: ${successCount}\n`;
  report += `❌ Vẫn fail: ${stillFailCount}\n`;
  report += `📍 Tổng đã thử lại: ${failedRows.length}\n\n`;
  
  if (Object.keys(sourceCount).length > 0) {
    report += `🔧 NGUỒN LẤY ĐƯỢC:\n`;
    for (let src in sourceCount) report += `• ${src}: ${sourceCount[src]} video\n`;
    report += '\n';
  }
  
  if (stillFailedRows.length > 0) {
    report += `🔍 VIDEO VẪN CHƯA LẤY ĐƯỢC:\n`;
    const showLimit = Math.min(stillFailedRows.length, 10);
    for (let i = 0; i < showLimit; i++) {
      let reason = stillFailedRows[i].reason || "";
      if (reason.length > 100) reason = reason.substring(0, 100) + "...";
      report += `• Dòng ${stillFailedRows[i].row}: ${reason}\n`;
    }
    if (stillFailedRows.length > 10) {
      report += `... và ${stillFailedRows.length - 10} dòng khác\n`;
    }
  }
  
  ui.alert(report);
}

/**
 * 🗑️ Xóa toàn bộ Cột M (Subtitle)
 */
function clearSubtitleColumn() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '🗑️ Xác nhận xóa',
    'Bạn có chắc muốn xóa TOÀN BỘ dữ liệu Subtitle ở Cột M?\n\nKhông thể hoàn tác!',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VIDEO);
  if (!sheet) return ui.alert("Không tìm thấy Sheet.");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  sheet.getRange(2, COL_VIDEO.SUBTITLE, lastRow - 1, 1).clearContent();
  ui.alert("✅ Đã xóa toàn bộ Subtitle ở Cột M.");
}

