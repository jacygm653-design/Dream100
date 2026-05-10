// =========================================================================
// HỆ THỐNG MÚI GIỜ THEO QUỐC GIA
// =========================================================================

/**
 * Map mã quốc gia → IANA timezone + thành phố tiêu biểu
 * Mục đích: Biết video được đăng vào giờ địa phương nào của tác giả.
 * 
 * Lưu ý: Các quốc gia rộng như US, RU, AU, CA có nhiều múi giờ.
 * Mặc định lấy múi giờ phổ biến nhất (Eastern cho US, Moscow cho RU...).
 * Có thể tinh chỉnh thêm dựa trên ngôn ngữ kênh sau này.
 */
const COUNTRY_TIMEZONE_MAP = {
  "VN": { tz: "Asia/Ho_Chi_Minh", city: "Hồ Chí Minh, Việt Nam" },
  "TH": { tz: "Asia/Bangkok", city: "Bangkok, Thái Lan" },
  "ID": { tz: "Asia/Jakarta", city: "Jakarta, Indonesia" },
  "PH": { tz: "Asia/Manila", city: "Manila, Philippines" },
  "MY": { tz: "Asia/Kuala_Lumpur", city: "Kuala Lumpur, Malaysia" },
  "SG": { tz: "Asia/Singapore", city: "Singapore" },
  "MM": { tz: "Asia/Yangon", city: "Yangon, Myanmar" },
  "KH": { tz: "Asia/Phnom_Penh", city: "Phnom Penh, Campuchia" },
  "LA": { tz: "Asia/Vientiane", city: "Vientiane, Lào" },
  "BN": { tz: "Asia/Brunei", city: "Bandar Seri Begawan, Brunei" },
  "TL": { tz: "Asia/Dili", city: "Dili, Đông Timor" },

  "JP": { tz: "Asia/Tokyo", city: "Tokyo, Nhật Bản" },
  "KR": { tz: "Asia/Seoul", city: "Seoul, Hàn Quốc" },
  "KP": { tz: "Asia/Pyongyang", city: "Bình Nhưỡng, Triều Tiên" },
  "CN": { tz: "Asia/Shanghai", city: "Thượng Hải, Trung Quốc" },
  "TW": { tz: "Asia/Taipei", city: "Đài Bắc, Đài Loan" },
  "HK": { tz: "Asia/Hong_Kong", city: "Hồng Kông" },
  "MO": { tz: "Asia/Macau", city: "Macau" },
  "MN": { tz: "Asia/Ulaanbaatar", city: "Ulaanbaatar, Mông Cổ" },

  "IN": { tz: "Asia/Kolkata", city: "Mumbai/New Delhi, Ấn Độ" },
  "PK": { tz: "Asia/Karachi", city: "Karachi, Pakistan" },
  "BD": { tz: "Asia/Dhaka", city: "Dhaka, Bangladesh" },
  "LK": { tz: "Asia/Colombo", city: "Colombo, Sri Lanka" },
  "NP": { tz: "Asia/Kathmandu", city: "Kathmandu, Nepal" },
  "BT": { tz: "Asia/Thimphu", city: "Thimphu, Bhutan" },
  "MV": { tz: "Indian/Maldives", city: "Malé, Maldives" },
  "AF": { tz: "Asia/Kabul", city: "Kabul, Afghanistan" },

  "AE": { tz: "Asia/Dubai", city: "Dubai, UAE" },
  "SA": { tz: "Asia/Riyadh", city: "Riyadh, Saudi Arabia" },
  "IL": { tz: "Asia/Jerusalem", city: "Tel Aviv, Israel" },
  "TR": { tz: "Europe/Istanbul", city: "Istanbul, Thổ Nhĩ Kỳ" },
  "IR": { tz: "Asia/Tehran", city: "Tehran, Iran" },
  "IQ": { tz: "Asia/Baghdad", city: "Baghdad, Iraq" },
  "QA": { tz: "Asia/Qatar", city: "Doha, Qatar" },
  "KW": { tz: "Asia/Kuwait", city: "Kuwait City, Kuwait" },
  "JO": { tz: "Asia/Amman", city: "Amman, Jordan" },
  "LB": { tz: "Asia/Beirut", city: "Beirut, Liban" },
  "SY": { tz: "Asia/Damascus", city: "Damascus, Syria" },

  "US": { tz: "America/Chicago", city: "Chicago/Central US" }, 
  "CA": { tz: "America/Toronto", city: "Toronto, Canada (ET)" },
  "MX": { tz: "America/Mexico_City", city: "Mexico City, Mexico" },
  "GT": { tz: "America/Guatemala", city: "Guatemala City, Guatemala" },
  "CU": { tz: "America/Havana", city: "La Habana, Cuba" },
  "DO": { tz: "America/Santo_Domingo", city: "Santo Domingo, Dominican" },
  "PR": { tz: "America/Puerto_Rico", city: "San Juan, Puerto Rico" },
  "JM": { tz: "America/Jamaica", city: "Kingston, Jamaica" },
  "HT": { tz: "America/Port-au-Prince", city: "Port-au-Prince, Haiti" },
  "PA": { tz: "America/Panama", city: "Panama City, Panama" },
  "CR": { tz: "America/Costa_Rica", city: "San José, Costa Rica" },
  "HN": { tz: "America/Tegucigalpa", city: "Tegucigalpa, Honduras" },
  "SV": { tz: "America/El_Salvador", city: "San Salvador, El Salvador" },
  "NI": { tz: "America/Managua", city: "Managua, Nicaragua" },

  "BR": { tz: "America/Sao_Paulo", city: "São Paulo, Brazil" },
  "AR": { tz: "America/Argentina/Buenos_Aires", city: "Buenos Aires, Argentina" },
  "CL": { tz: "America/Santiago", city: "Santiago, Chile" },
  "CO": { tz: "America/Bogota", city: "Bogotá, Colombia" },
  "PE": { tz: "America/Lima", city: "Lima, Peru" },
  "VE": { tz: "America/Caracas", city: "Caracas, Venezuela" },
  "EC": { tz: "America/Guayaquil", city: "Quito, Ecuador" },
  "BO": { tz: "America/La_Paz", city: "La Paz, Bolivia" },
  "PY": { tz: "America/Asuncion", city: "Asunción, Paraguay" },
  "UY": { tz: "America/Montevideo", city: "Montevideo, Uruguay" },
  "GY": { tz: "America/Guyana", city: "Georgetown, Guyana" },
  "SR": { tz: "America/Paramaribo", city: "Paramaribo, Suriname" },

  "GB": { tz: "Europe/London", city: "London, Anh Quốc" },
  "IE": { tz: "Europe/Dublin", city: "Dublin, Ireland" },
  "FR": { tz: "Europe/Paris", city: "Paris, Pháp" },
  "DE": { tz: "Europe/Berlin", city: "Berlin, Đức" },
  "IT": { tz: "Europe/Rome", city: "Rome, Italia" },
  "ES": { tz: "Europe/Madrid", city: "Madrid, Tây Ban Nha" },
  "PT": { tz: "Europe/Lisbon", city: "Lisbon, Bồ Đào Nha" },
  "NL": { tz: "Europe/Amsterdam", city: "Amsterdam, Hà Lan" },
  "BE": { tz: "Europe/Brussels", city: "Brussels, Bỉ" },
  "CH": { tz: "Europe/Zurich", city: "Zurich, Thụy Sĩ" },
  "AT": { tz: "Europe/Vienna", city: "Vienna, Áo" },
  "PL": { tz: "Europe/Warsaw", city: "Warsaw, Ba Lan" },
  "CZ": { tz: "Europe/Prague", city: "Praha, Cộng hòa Séc" },
  "SK": { tz: "Europe/Bratislava", city: "Bratislava, Slovakia" },
  "HU": { tz: "Europe/Budapest", city: "Budapest, Hungary" },
  "RO": { tz: "Europe/Bucharest", city: "Bucharest, Romania" },
  "BG": { tz: "Europe/Sofia", city: "Sofia, Bulgaria" },
  "GR": { tz: "Europe/Athens", city: "Athens, Hy Lạp" },
  "SE": { tz: "Europe/Stockholm", city: "Stockholm, Thụy Điển" },
  "NO": { tz: "Europe/Oslo", city: "Oslo, Na Uy" },
  "DK": { tz: "Europe/Copenhagen", city: "Copenhagen, Đan Mạch" },
  "FI": { tz: "Europe/Helsinki", city: "Helsinki, Phần Lan" },
  "IS": { tz: "Atlantic/Reykjavik", city: "Reykjavik, Iceland" },
  "UA": { tz: "Europe/Kyiv", city: "Kyiv, Ukraine" },
  "RU": { tz: "Europe/Moscow", city: "Moscow, Nga" },
  "BY": { tz: "Europe/Minsk", city: "Minsk, Belarus" },
  "RS": { tz: "Europe/Belgrade", city: "Belgrade, Serbia" },
  "HR": { tz: "Europe/Zagreb", city: "Zagreb, Croatia" },
  "SI": { tz: "Europe/Ljubljana", city: "Ljubljana, Slovenia" },
  "EE": { tz: "Europe/Tallinn", city: "Tallinn, Estonia" },
  "LV": { tz: "Europe/Riga", city: "Riga, Latvia" },
  "LT": { tz: "Europe/Vilnius", city: "Vilnius, Lithuania" },

  "EG": { tz: "Africa/Cairo", city: "Cairo, Ai Cập" },
  "ZA": { tz: "Africa/Johannesburg", city: "Johannesburg, Nam Phi" },
  "NG": { tz: "Africa/Lagos", city: "Lagos, Nigeria" },
  "KE": { tz: "Africa/Nairobi", city: "Nairobi, Kenya" },
  "MA": { tz: "Africa/Casablanca", city: "Casablanca, Morocco" },
  "DZ": { tz: "Africa/Algiers", city: "Algiers, Algeria" },
  "TN": { tz: "Africa/Tunis", city: "Tunis, Tunisia" },
  "ET": { tz: "Africa/Addis_Ababa", city: "Addis Ababa, Ethiopia" },
  "GH": { tz: "Africa/Accra", city: "Accra, Ghana" },
  "TZ": { tz: "Africa/Dar_es_Salaam", city: "Dar es Salaam, Tanzania" },
  "UG": { tz: "Africa/Kampala", city: "Kampala, Uganda" },
  "ZW": { tz: "Africa/Harare", city: "Harare, Zimbabwe" },

  "AU": { tz: "Australia/Sydney", city: "Sydney, Úc" },
  "NZ": { tz: "Pacific/Auckland", city: "Auckland, New Zealand" },
  "FJ": { tz: "Pacific/Fiji", city: "Suva, Fiji" },
  "PG": { tz: "Pacific/Port_Moresby", city: "Port Moresby, Papua New Guinea" }
};

/**
 * 🆕 Lấy thông tin múi giờ + offset theo quốc gia + thời điểm cụ thể
 * Tự động xử lý DST (Daylight Saving Time) bằng Intl API
 * 
 * @param {string} countryCode - Mã ISO 3166-1 alpha-2 (vd: "US", "VN")
 * @param {Date} dateUTC - Thời điểm cần tra cứu (ảnh hưởng DST mùa đông/hè)
 * @return {object} { tz, city, offsetHours, offsetStr, abbr, found }
 */
function getCountryTimezoneInfo(countryCode, dateUTC) {
  const result = {
    tz: "UTC",
    city: "Không xác định",
    offsetHours: 0,
    offsetStr: "UTC ±0",
    abbr: "UTC",
    found: false
  };

  if (!countryCode) return result;
  
  const code = countryCode.toString().toUpperCase().trim();
  const info = COUNTRY_TIMEZONE_MAP[code];
  
  if (!info) {
    result.city = "Quốc gia: " + code + " (chưa hỗ trợ)";
    return result;
  }
  
  result.tz = info.tz;
  result.city = info.city;
  result.found = true;
  
  try {
    const date = dateUTC instanceof Date ? dateUTC : new Date(dateUTC);
    
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: info.tz,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
      timeZoneName: 'short'
    });
    
    const parts = formatter.formatToParts(date);
    let abbr = "";
    let year = 0, month = 0, day = 0, hour = 0, min = 0, sec = 0;
    
    for (const p of parts) {
      if (p.type === 'timeZoneName') abbr = p.value;
      else if (p.type === 'year') year = parseInt(p.value);
      else if (p.type === 'month') month = parseInt(p.value);
      else if (p.type === 'day') day = parseInt(p.value);
      else if (p.type === 'hour') hour = parseInt(p.value === '24' ? '0' : p.value);
      else if (p.type === 'minute') min = parseInt(p.value);
      else if (p.type === 'second') sec = parseInt(p.value);
    }
    
    const localAsUtc = Date.UTC(year, month - 1, day, hour, min, sec);
    const utcMs = date.getTime();
    const offsetMs = localAsUtc - utcMs;
    const offsetHours = offsetMs / (1000 * 60 * 60);
    
    result.offsetHours = offsetHours;
    result.abbr = abbr || info.tz;
    
    const sign = offsetHours >= 0 ? '+' : '-';
    const absHours = Math.abs(offsetHours);
    const hoursPart = Math.floor(absHours);
    const minsPart = Math.round((absHours - hoursPart) * 60);
    
    if (minsPart === 0) {
      result.offsetStr = `UTC ${sign}${hoursPart}`;
    } else {
      result.offsetStr = `UTC ${sign}${hoursPart}:${minsPart.toString().padStart(2, '0')}`;
    }
    
  } catch (e) {
    result.offsetStr = "UTC (lỗi tính offset: " + e.message + ")";
  }
  
  return result;
}

/**
 * 🆕 Format chuỗi 3 dòng cho cột "NGÀY ĐĂNG":
 *   - Sioux Falls, Nam Dakota: UTC -5 / Central Daylight Time (CDT)
 *   - Giờ Việt Nam: 09/05/2026 15:48:56
 *   - Giờ video đăng theo khu vực: 09/05/2026 03:48:56
 * 
 * @param {string} publishedAtISO - ISO timestamp từ YouTube API (UTC)
 * @param {string} countryCode - Mã quốc gia ISO của kênh
 * @return {string} Chuỗi 3 dòng đã format
 */
function formatPublishedAtMultiline(publishedAtISO, countryCode) {
  if (!publishedAtISO) return "";
  
  const dateUTC = new Date(publishedAtISO);
  if (isNaN(dateUTC.getTime())) return "";
  
  const vnTimeStr = Utilities.formatDate(dateUTC, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");
  
  let line1 = "";
  let line3 = "";
  
  if (countryCode) {
    const tzInfo = getCountryTimezoneInfo(countryCode, dateUTC);
    
    if (tzInfo.found) {
      const localTimeStr = Utilities.formatDate(dateUTC, tzInfo.tz, "dd/MM/yyyy HH:mm:ss");
      line1 = `- ${tzInfo.city}: ${tzInfo.offsetStr} / ${tzInfo.abbr}`;
      line3 = `- Giờ video đăng theo khu vực: ${localTimeStr}`;
    } else {
      line1 = `- ${tzInfo.city}`;
      line3 = `- Giờ video đăng theo khu vực: (không xác định)`;
    }
  } else {
    line1 = `- Quốc gia kênh: chưa khai báo`;
    line3 = `- Giờ video đăng theo khu vực: (không xác định)`;
  }
  
  return `${line1}\n- Giờ Việt Nam: ${vnTimeStr}\n${line3}`;
}


function parseVietnameseDate(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) return dateVal;
  try {
    const str = dateVal.toString().trim();
    const parts = str.split(' ');
    if (parts.length >= 2) {
      const dateParts = parts[0].split('/');
      const timeParts = parts[1].split(':');
      if (dateParts.length === 3 && timeParts.length >= 2) {
        return new Date(dateParts[2], dateParts[1] - 1, dateParts[0], timeParts[0], timeParts[1], timeParts[2] || 0);
      }
    }
  } catch(e) {}
  return null;
}

function unformatNumber(num) {
  if (!num || num === "") return 0;
  return parseInt(num.toString().replace(/[^0-9]/g, '')) || 0;
}

function formatNumber(num) {
  if (num === null || num === undefined || num === "") return "0";
  return Number(num).toLocaleString('vi-VN');
}

function extractVideoId(url) {
  if (!url) return null;
  const str = url.toString().trim();
  const match = str.match(/(?:https?:\/\/)?(?:www\.)?(?:youtube\.com\/(?:[^\/\n\s]+\/\S+\/|(?:v|e(?:mbed)?)\/|shorts\/|live\/|\S*?[?&]v=)|youtu\.be\/)([a-zA-Z0-9_-]{11})/);
  if (match) return match[1];
  return (str.length === 11 && /^[a-zA-Z0-9_-]{11}$/.test(str)) ? str : null;
}

function extractHashtags(text) {
  if (!text) return "";
  const matches = text.match(/#[a-zA-Z0-9_À-ỹ]+/g);
  return matches ? [...new Set(matches)].join(', ') : "";
}

function parseISODuration(duration) {
  if (!duration || typeof duration !== 'string') return "00:00:00";
  const match = duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return "00:00:00";
  const h = (match[1] || "00").padStart(2, '0');
  const m = (match[2] || "00").padStart(2, '0');
  const s = (match[3] || "00").padStart(2, '0');
  return `${h}:${m}:${s}`;
}

function inferPublicMonetizationStatus_(channel, recentVideos) {
  const snippet = channel.snippet || {};
  const stats = channel.statistics || {};
  const status = channel.status || {};
  const branding = channel.brandingSettings || {};
  const subs = parseInt(stats.subscriberCount || 0, 10) || 0;
  const videoCount = parseInt(stats.videoCount || 0, 10) || 0;
  const publicSignals = [];
  const blockers = [];

  if (status.madeForKids === true || status.selfDeclaredMadeForKids === true) {
    blockers.push("madeForKids");
  }
  if (status.longUploadsStatus && status.longUploadsStatus !== "allowed") {
    blockers.push("longUploads=" + status.longUploadsStatus);
  }
  if (subs >= 1000) publicSignals.push(">=1.000 subs");
  if ((recentVideos || []).some(v => parseInt((v.statistics && v.statistics.viewCount) || 0, 10) >= 10000)) {
    publicSignals.push("có video gần đây >=10.000 views");
  }
  if (branding.channel && branding.channel.trackingAnalyticsAccountId) {
    publicSignals.push("có trackingAnalyticsAccountId");
  }
  if (videoCount >= 3) publicSignals.push("có nhiều video public");

  if (blockers.length > 0) {
    return "Không thể xác minh 100% qua API công khai; có tín hiệu hạn chế: " + blockers.join(", ");
  }
  if (publicSignals.length > 0) {
    return "Không thể xác minh 100% qua API công khai; tín hiệu công khai: " + publicSignals.join("; ");
  }
  return "Không thể xác minh 100% qua API công khai; chưa có tín hiệu đủ mạnh";
}

function getChannelDataFromAPI(url, apiKey) {
  let urlStr = url.toString().trim().replace(/\/$/, ""); 
  let endpoints = [];
  if (urlStr.length === 24 && urlStr.startsWith("UC")) endpoints.push(`id=${urlStr}`);
  else if (urlStr.startsWith("@")) endpoints.push(`forHandle=${urlStr.substring(1)}`);
  else if (urlStr.includes("/channel/UC")) { const m = urlStr.match(/\/channel\/(UC[\w-]{22})/); if (m) endpoints.push(`id=${m[1]}`); }
  else if (urlStr.includes("@")) { const m = urlStr.match(/@([\w.-]+)/); if (m) endpoints.push(`forHandle=${m[1]}`); }
  else if (urlStr.includes("/user/")) { const m = urlStr.match(/\/user\/([\w.-]+)/); if (m) endpoints.push(`forUsername=${m[1]}`); }
  else if (urlStr.includes("/c/")) { const m = urlStr.match(/\/c\/([\w.-]+)/); if (m) { endpoints.push(`forHandle=@${m[1]}`); endpoints.push(`forUsername=${m[1]}`); } }
  else { const m = urlStr.match(/youtube\.com\/([^/?&#]+)/); if (m && !['watch', 'shorts', 'playlist', 'results', 'live', 'feed'].includes(m[1])) { endpoints.push(`forHandle=@${m[1]}`); endpoints.push(`forHandle=${m[1]}`); endpoints.push(`forUsername=${m[1]}`); } }
  if (endpoints.length === 0) return { found: false };
  let json = null;
  for (let ep of endpoints) {
    try {
      let tempJson = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics,topicDetails,contentDetails,brandingSettings&${ep}`, apiKey);
      if (tempJson.items && tempJson.items.length > 0) { json = tempJson; break; }
    } catch (e) {
      continue;
    }
  }
  if (!json) return { found: false };
  
  const ch = json.items[0]; 
  const brand = ch.brandingSettings || {}; 
  const stats = ch.statistics || {}; 
  const snippet = ch.snippet || {};
  const topicDetails = ch.topicDetails || {};
  const recentVideosForSignals = [];
  
  let views4Weeks = 0; 
  let vidsPerMonth = 0; 
  let totalSecs = 0; 
  let validCount = 0;
  let videoCategoryId = null;
  const uploadHours = [];
  
  const playlistId = ch.contentDetails && ch.contentDetails.relatedPlaylists ? ch.contentDetails.relatedPlaylists.uploads : null;
  if (playlistId) {
    try {
      const plJson = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/playlistItems?part=contentDetails&playlistId=${playlistId}&maxResults=50`, apiKey);
      if (plJson.items && plJson.items.length > 0) {
        const vids = fetchAPIWithRetry(`https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics,contentDetails&id=${plJson.items.map(it=>it.contentDetails.videoId).join(',')}`, apiKey);
        const fourWeeksAgo = new Date(); fourWeeksAgo.setDate(fourWeeksAgo.getDate() - 28);
        const thirtyDaysAgo = new Date(); thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        const channelTz = snippet.country ? (COUNTRY_TIMEZONE_MAP[snippet.country.toUpperCase()] || null) : null;
        
        vids.items.forEach(v => {
          if (v.snippet.liveBroadcastContent !== "none") return;
          recentVideosForSignals.push(v);
          const pubDate = new Date(v.snippet.publishedAt);
          if (pubDate >= fourWeeksAgo) views4Weeks += parseInt(v.statistics.viewCount || 0);
          if (pubDate >= thirtyDaysAgo) vidsPerMonth++;
          
          const match = (v.contentDetails.duration || "").match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
          if (match) { 
            totalSecs += (parseInt(match[1]||0)*3600 + parseInt(match[2]||0)*60 + parseInt(match[3]||0)); 
            validCount++; 
          }
          
          if (!videoCategoryId && v.snippet.categoryId) {
            videoCategoryId = v.snippet.categoryId;
          }
          
          if (channelTz) {
            try {
              const hourStr = Utilities.formatDate(pubDate, channelTz.tz, "HH");
              uploadHours.push(parseInt(hourStr));
            } catch(e) {}
          } else {
            try {
              const hourStr = Utilities.formatDate(pubDate, "Asia/Ho_Chi_Minh", "HH");
              uploadHours.push(parseInt(hourStr));
            } catch(e) {}
          }
        });
      }
    } catch(e) {}
  }
  
  let categoryName = "";
  if (videoCategoryId) {
    const region = snippet.country || "US";
    categoryName = getVideoCategoryName(videoCategoryId, region, apiKey);
  }
  
  let mainTopic = "";
  if (topicDetails.topicCategories && topicDetails.topicCategories.length > 0) {
    const topics = topicDetails.topicCategories.slice(0, 3).map(url => {
      const parts = url.split('/wiki/');
      return parts.length > 1 ? decodeURIComponent(parts[1]).replace(/_/g, ' ') : "";
    }).filter(t => t);
    mainTopic = topics.join(' / ');
  }
  
  let commonUploadHour = "";
  if (uploadHours.length > 0) {
    const hourFreq = {};
    uploadHours.forEach(h => { hourFreq[h] = (hourFreq[h] || 0) + 1; });
    let maxFreq = 0, topHour = -1;
    for (const h in hourFreq) {
      if (hourFreq[h] > maxFreq) { maxFreq = hourFreq[h]; topHour = parseInt(h); }
    }
    if (topHour >= 0) {
      commonUploadHour = `${topHour.toString().padStart(2, '0')}h (${maxFreq}/${uploadHours.length} video)`;
    }
  }
  
  const avgSecs = validCount ? Math.floor(totalSecs/validCount) : 0;
  return { 
    found: true, 
    title: snippet.title, 
    subscribers: stats.subscriberCount, 
    views4Weeks, 
    vidsPerMonth, 
    avgDuration: `${Math.floor(avgSecs/3600).toString().padStart(2,'0')}:${Math.floor((avgSecs%3600)/60).toString().padStart(2,'0')}:${(avgSecs%60).toString().padStart(2,'0')}`, 
    description: snippet.description, 
    country: snippet.country, 
    keywords: brand.channel ? brand.channel.keywords : "",
    category: categoryName,
    mainTopic: mainTopic,
    commonUploadHour: commonUploadHour,
    monetizationStatus: inferPublicMonetizationStatus_(ch, recentVideosForSignals)
  };
}

/**
 * 🆕 Lấy tên category từ categoryId (có cache)
 * @param {string} categoryId - vd: "10", "22"
 * @param {string} regionCode - vd: "US", "VN"
 * @param {string} apiKey
 * @return {string} Tên category, vd: "Music", "People & Blogs"
 */
function getVideoCategoryName(categoryId, regionCode, apiKey) {
  if (!categoryId) return "";
  
  const cache = CacheService.getScriptCache();
  const cacheKey = `vc_${regionCode}_${categoryId}`;
  
  const cached = cache.get(cacheKey);
  if (cached) return cached;
  
  try {
    const url = `https://www.googleapis.com/youtube/v3/videoCategories?part=snippet&regionCode=${regionCode}`;
    const json = fetchAPIWithRetry(url, apiKey);
    
    if (json.items) {
      let foundName = "";
      for (const cat of json.items) {
        const name = (cat.snippet && cat.snippet.title) || "";
        try {
          cache.put(`vc_${regionCode}_${cat.id}`, name, 21600);
        } catch(e) {}
        if (cat.id === categoryId) foundName = name;
      }
      return foundName;
    }
  } catch (e) {
    const fallbackMap = {
      "1": "Film & Animation", "2": "Autos & Vehicles", "10": "Music",
      "15": "Pets & Animals", "17": "Sports", "19": "Travel & Events",
      "20": "Gaming", "22": "People & Blogs", "23": "Comedy",
      "24": "Entertainment", "25": "News & Politics", "26": "Howto & Style",
      "27": "Education", "28": "Science & Technology", "29": "Nonprofits & Activism"
    };
    return fallbackMap[categoryId] || `Category ${categoryId}`;
  }
  
  return "";
}

