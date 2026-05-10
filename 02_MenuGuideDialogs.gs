// =========================================================================
// MENU
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 YouTube Tools')
    .addItem('📘 Trung tâm hướng dẫn hệ thống', 'showSystemBrandInfo')
    .addItem('🎛️  MỞ BẢNG ĐIỀU KHIỂN', 'openAdminPanel')
    .addItem('🔄 Kiểm tra cập nhật hệ thống', 'showSystemUpdateCenter')
    .addToUi();
  showDailyQuotaReminder_(ui);
}

// =========================================================================
// HÀM HỖ TRỢ LẤY KHOẢNG DÒNG TỪ NGƯỜI DÙNG
// =========================================================================
function _getGuideSectionsArray_() {
  return [
    { group: "Bắt đầu", title: "Tổng quan hệ thống", html: `
      <h2>Tổng quan hệ thống</h2>
      <p><b>YouTube Tools - Dương Huỳnh Team</b> là bộ công cụ chạy trong Google Sheets bằng Apps Script, giúp khai thác dữ liệu YouTube Data API, YouTube Analytics API và phân tích bằng AI thông qua 9router.</p>
      <h3>📂 Các sheet chính</h3>
      <table class="guideTable"><tr><th>Sheet</th><th>Vai trò</th></tr><tr><td><b>${SHEET_VIDEO}</b></td><td>Danh sách video, link ở cột C, metadata ở các cột B–N</td></tr><tr><td><b>${SHEET_CHANNEL}</b></td><td>Danh sách kênh, link ở cột C, metadata ở các cột B–N</td></tr><tr><td><b>${SHEET_API_KEY}</b></td><td>Tất cả key/token: B2 (YouTube Data), B3/B4 (OAuth), B5 (Supadata), B6 (YT-Transcript.io), B7 (9router)</td></tr><tr><td><b>${SHEET_ANALYTICS}</b></td><td>Báo cáo Analytics tổng quan kênh (7d/28d/lifetime)</td></tr><tr><td><b>${SHEET_VIDEO_ANALYTICS}</b></td><td>Analytics chi tiết theo từng video</td></tr></table>
      <h3>🔁 Quy trình chuẩn lần đầu</h3>
      <ol>
        <li>Vào sheet <b>${SHEET_API_KEY}</b> và điền đầy đủ key/token theo đúng ô.</li>
        <li>Mở Trung tâm hướng dẫn này, đọc các tab <b>API key</b> và <b>Sheet</b> để hiểu cách điền.</li>
        <li>Chạy thử <b>1 dòng</b> hoặc một <b>khoảng nhỏ</b> (ví dụ <code>2-5</code>) để xác nhận key hoạt động.</li>
        <li>Sau khi ổn, mới chạy toàn bộ.</li>
        <li>Sau tác vụ lớn, mở menu <b>📊 Xem Quota YouTube API hôm nay</b> để theo dõi quota còn lại.</li>
      </ol>
      <h3>⚙️ Module có sẵn</h3>
      <ul>
        <li><b>VIDEO</b>: cập nhật full / khoảng / theo dòng / nhanh views &amp; VPH.</li>
        <li><b>KÊNH</b>: cập nhật full / khoảng / theo dòng / nhanh subs &amp; views/tháng.</li>
        <li><b>Tìm kiếm</b>: theo chủ đề (video) hoặc theo chủ đề (kênh), lọc đa điều kiện.</li>
        <li><b>Lấy video trong kênh</b> (&lt; 3 tháng) với điều kiện thời lượng/views/chủ đề.</li>
        <li><b>Subtitle</b>: 4 nguồn fallback (Supadata → YT-Transcript.io → timedtext → scrape).</li>
        <li><b>Dọn dẹp</b>: xóa video cũ/trùng, xóa video &lt; 20.000 views, xóa kênh chết, dọn tổng hợp.</li>
        <li><b>Analytics</b>: kết nối OAuth, báo cáo tổng quan kênh và từng video (7d/28d/lifetime).</li>
        <li><b>AI phân tích Sheet</b>: gọi 9router OpenAI-compatible, có lịch sử ở sheet ẩn <b>${AI_MEMORY_SHEET}</b>.</li>
      </ul>
      <div class="tip"><b>💡 Gợi ý:</b> Quota mặc định YouTube Data API là <b>${YOUTUBE_DAILY_QUOTA_LIMIT.toLocaleString()} units/ngày/project</b>. <code>search.list</code> tốn ~100 units, <code>videos.list</code>/<code>channels.list</code> tốn ~1 unit. Ưu tiên dùng list theo ID, tránh dùng search trừ khi cần.</div>
      <div class="note"><b>🔒 An toàn:</b> không chia sẻ API key, không đưa key lên nơi công khai, nên giới hạn API key theo đúng API cần dùng (Restrict key → chỉ chọn YouTube Data API v3).</div>
      <div class="warn"><b>⚠️ Lưu ý:</b> không đóng file Google Sheets khi script đang chạy. Với danh sách lớn, hãy chia khoảng dòng để tránh timeout 6 phút của Apps Script.</div>
      <div class="videoBox"><div style="font-size:20px;font-weight:900;background:linear-gradient(135deg,#4f46e5 0%,#7c3aed 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;color:#4f46e5;margin-bottom:12px;letter-spacing:0.01em;">🎬 Hướng dẫn lấy tất cả API key cần thiết</div><div style="position:relative;width:100%;padding-top:56.25%;border-radius:12px;overflow:hidden;background:#0f172a;box-shadow:0 12px 32px -10px rgba(15,23,42,0.32);"><iframe src="https://www.youtube.com/embed/Y2yFS9XLhCA" title="Hướng dẫn lấy tất cả API key cần thiết" style="position:absolute;inset:0;width:100%;height:100%;border:0;" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe></div></div>` },
    { group: "Bắt đầu", title: "🩺 Kiểm tra sức khỏe hệ thống", html: `
      <h2>🩺 Kiểm tra sức khỏe hệ thống</h2>
      <p>Chức năng chẩn đoán toàn diện, kiểm tra cùng lúc <b>tất cả nguồn cấu hình</b>: API key, OAuth, sheet, quota, nguồn subtitle. Mở từ menu: <b>🚀 YouTube Tools → 🩺 Kiểm tra sức khỏe hệ thống</b>.</p>
      <h3>📋 Các mục được kiểm tra</h3>
      <table class="guideTable"><tr><th>Nhóm</th><th>Mục</th><th>Tiêu chí</th></tr>
        <tr><td>Sheet</td><td>${SHEET_VIDEO}, ${SHEET_CHANNEL}, ${SHEET_API_KEY}, ${SHEET_ANALYTICS}, ${SHEET_VIDEO_ANALYTICS}, ${AI_MEMORY_SHEET}</td><td>Tồn tại + số dòng dữ liệu</td></tr>
        <tr><td>API Key</td><td>B2 / B3 / B4 / B5 / B6 / B7</td><td>Có giá trị, độ dài ký tự</td></tr>
        <tr><td>Subtitle</td><td>Phủ sóng nguồn</td><td>Phải có ít nhất 1 trong B5/B6</td></tr>
        <tr><td>OAuth</td><td>Refresh token</td><td>Đã cấp quyền hay chưa</td></tr>
        <tr><td>Quota</td><td>Tỷ lệ dùng hôm nay</td><td>&lt;50% OK · 50–80% Cảnh báo · ≥80% Nghiêm trọng</td></tr>
        <tr><td>AI</td><td>9router base URL</td><td>localhost (cảnh báo) hay public URL (OK)</td></tr>
      </table>
      <h3>🎯 Health Score</h3>
      <p>Điểm sức khỏe được tính theo công thức: <code>(số mục OK + số cảnh báo × 0.5) / tổng mục × 100</code>. Hiển thị bằng vòng tròn gradient và 3 badge thống kê (OK / Cảnh báo / Nghiêm trọng).</p>
      <div class="tip"><b>💡 Khi nào nên chạy?</b><ul><li>Sau khi vừa setup hệ thống lần đầu.</li><li>Sau khi đổi API key/OAuth.</li><li>Khi gặp lỗi không rõ nguyên nhân.</li><li>Định kỳ hằng tuần để rà soát cấu hình.</li></ul></div>
      <div class="note"><b>📊 Theo dõi quota trực quan:</b> menu <b>📊 Xem Quota YouTube API hôm nay</b> giờ là dashboard động với vòng tròn %, breakdown chi tiết, ETA reset Pacific, auto-refresh 10 giây và gợi ý tối ưu theo mức tiêu thụ.</div>` },
    { group: "Bắt đầu", title: "Quy trình hằng ngày khuyến nghị", html: `
      <h2>Quy trình hằng ngày khuyến nghị</h2>
      <p>Quy trình tối ưu để duy trì dữ liệu mới mà vẫn tiết kiệm quota.</p>
      <h3>🌅 Buổi sáng</h3>
      <ol>
        <li>Chạy <b>Cập nhật NHANH Views &amp; VPH (Sheet VIDEO)</b> — chỉ làm mới views &amp; VPH, tốn ít quota.</li>
        <li>Chạy <b>Cập nhật NHANH Subs &amp; Views/Tháng (Sheet KÊNH)</b> để theo dõi subs/tháng.</li>
        <li>Mở <b>📊 Xem Quota YouTube API hôm nay</b> để xác định ngân sách quota còn lại trong ngày.</li>
      </ol>
      <h3>🌞 Buổi trưa</h3>
      <ol>
        <li>Nếu cần thêm dữ liệu mới: <b>Lấy links video trong kênh (&lt; 3 tháng)</b> với khoảng dòng nhỏ.</li>
        <li>Cập nhật metadata video mới thêm bằng <b>Cập nhật THEO KHOẢNG DÒNG</b>.</li>
        <li>Chạy <b>Lấy Subtitle THEO KHOẢNG DÒNG</b> cho khoảng dòng vừa thêm (ưu tiên Supadata B5).</li>
      </ol>
      <h3>🌙 Cuối ngày</h3>
      <ol>
        <li>Chạy <b>Analytics TỔNG QUAN KÊNH</b> nếu đã kết nối OAuth.</li>
        <li>Chạy <b>AI phân tích Sheet</b> cho phần dữ liệu hôm nay (chọn đúng cột để tiết kiệm context).</li>
        <li>Kiểm tra quota tổng đã dùng — nếu &gt; 80% thì giảm tải tác vụ ngày hôm sau.</li>
      </ol>
      <div class="tip"><b>💡 Mẹo:</b> Đặt cố định một bộ 3 tab: VIDEO, KÊNH, API KEY. Mọi tác vụ đều quay về sheet này nên tránh đổi tên sheet — code đọc đúng tên hằng số.</div>
      <div class="warn"><b>⚠️ Tránh:</b> chạy đồng thời nhiều tác vụ nặng (search + update full) trong cùng một ngày sẽ vượt quota nhanh.</div>` },
    { group: "Quota", title: "Tối ưu quota & chi phí", html: `
      <h2>Tối ưu quota &amp; chi phí</h2>
      <p>YouTube Data API v3 mặc định cấp <b>${YOUTUBE_DAILY_QUOTA_LIMIT.toLocaleString()} units/ngày</b> cho mỗi project. Hệ thống đã ước lượng và ghi nhận chi phí từng request qua <code>estimateYouTubeQuotaCost_</code>.</p>
      <h3>💰 Bảng chi phí thực tế trong code</h3>
      <table class="guideTable"><tr><th>Endpoint</th><th>Chi phí</th><th>Khi nào dùng</th></tr><tr><td><code>search.list</code></td><td>~100 units</td><td>Tìm kiếm theo chủ đề (video / kênh)</td></tr><tr><td><code>videos.list</code></td><td>~1 unit</td><td>Lấy metadata video theo ID</td></tr><tr><td><code>channels.list</code></td><td>~1 unit</td><td>Lấy metadata kênh theo ID</td></tr><tr><td><code>playlistItems.list</code></td><td>~1 unit</td><td>Lấy danh sách video trong kênh</td></tr><tr><td><code>videoCategories.list</code></td><td>~1 unit</td><td>Lấy tên category</td></tr></table>
      <h3>🎯 Chiến lược giảm quota</h3>
      <ol>
        <li><b>Tránh lạm dụng search</b>: 1 lần search = 100 update bằng list. Nếu đã có ID/URL, dùng update thẳng.</li>
        <li><b>Chạy theo khoảng dòng</b>: tránh chạy full khi sheet đã &gt; 200 dòng.</li>
        <li><b>Dùng Cập nhật NHANH</b> thay cho full khi chỉ cần views/VPH/subs.</li>
        <li><b>Tách project nhiều API key</b>: mỗi project Google Cloud có quota riêng 10.000 units. Có thể luân phiên key qua sheet API KEY.</li>
        <li><b>Yêu cầu tăng quota</b>: trong Google Cloud → IAM &amp; Admin → Quotas → YouTube Data API v3 → Edit Quotas (cần lý do thuyết phục).</li>
      </ol>
      <h3>🧾 Chi phí ngoài YouTube API</h3>
      <ul>
        <li><b>Supadata (B5)</b>: tính theo gói trên dash.supadata.ai.</li>
        <li><b>YouTube-Transcript.io (B6)</b>: tối đa 50 video ID/request, có rate limit theo gói.</li>
        <li><b>9router (B7)</b>: tự cấu hình provider; chi phí phụ thuộc model bạn chọn (xem dashboard 9router).</li>
        <li><b>Analytics OAuth</b>: không tính quota Data API, có quota Analytics API riêng (rất rộng cho mục đích đọc).</li>
      </ul>
      <div class="tip"><b>💡 Theo dõi:</b> Mở menu <b>📊 Xem Quota YouTube API hôm nay</b> bất cứ lúc nào để biết units đã tiêu — số liệu cập nhật real-time qua <code>PropertiesService</code>.</div>` },
    { group: "Khắc phục lỗi", title: "Khắc phục lỗi thường gặp", html: `
      <h2>Khắc phục lỗi thường gặp</h2>
      <h3>🔑 Lỗi liên quan API key</h3>
      <ul>
        <li><b>"API key not valid"</b>: kiểm tra <b>API KEY!B2</b> — phải copy nguyên không có khoảng trắng đầu/cuối. Đảm bảo đã <b>Enable</b> YouTube Data API v3.</li>
        <li><b>"This API key is not authorized"</b>: key đang bị Restrict. Vào Cloud Console → Credentials → mở key → API restrictions → chọn <b>YouTube Data API v3</b>.</li>
        <li><b>"quotaExceeded"</b>: hết quota ngày, đợi reset 0:00 PT (theo Google) hoặc đổi sang project khác.</li>
      </ul>
      <h3>🔐 Lỗi OAuth Analytics</h3>
      <ul>
        <li><b>"redirect_uri_mismatch"</b>: trong OAuth client phải có redirect URI <code>http://localhost</code> (đúng schema/host).</li>
        <li><b>"access_denied"</b>: tài khoản chưa được thêm vào <b>Test users</b> khi app đang trong chế độ Testing.</li>
        <li><b>"invalid_grant"</b>: refresh token hết hạn hoặc bị revoke. Chạy lại <b>Xóa Token Analytics</b> rồi <b>Cài đặt Access Token</b>.</li>
        <li><b>Doanh thu không trả về</b>: tài khoản không có quyền tài chính trên kênh, hoặc kênh chưa bật kiếm tiền.</li>
      </ul>
      <h3>📝 Lỗi Subtitle</h3>
      <ul>
        <li><b>Hay 429/timeout với timedtext/scrape</b>: cấu hình <b>API KEY!B5</b> (Supadata) — ổn định hơn.</li>
        <li><b>Supadata báo lỗi auth</b>: kiểm tra header <code>x-api-key</code> — code dùng đúng định danh, key chỉ cần copy đúng.</li>
        <li><b>YT-Transcript.io rate limit</b>: chia khoảng dòng nhỏ hơn (≤ 30 dòng/lần). Code đã giới hạn 50 video ID/request theo doc.</li>
        <li><b>Subtitle trống nhưng video có CC</b>: dùng menu <b>Thử lại các dòng FAIL</b> sau khi cấu hình thêm B5 hoặc B6.</li>
      </ul>
      <h3>⏱️ Lỗi timeout / chạy chậm</h3>
      <ul>
        <li><b>Apps Script timeout 6 phút</b>: chia khoảng dòng. Ví dụ <code>1-100</code> rồi <code>101-200</code>.</li>
        <li><b>Sheet có nhiều công thức nặng</b>: tạm copy sang sheet khác chạy, sau đó dán lại.</li>
        <li><b>Không thấy progress</b>: mở <b>View → Executions</b> trong Apps Script editor để xem log.</li>
      </ul>
      <h3>🤖 Lỗi AI phân tích Sheet</h3>
      <ul>
        <li><b>Không gọi được localhost</b>: Apps Script chạy trên server Google nên không truy cập <code>http://localhost:20128</code>. Dùng public URL/tunnel (ngrok/cloudflared) và dán vào <b>9router base URL</b> trong tab Cài đặt.</li>
        <li><b>Lỗi 401</b>: token sai/hết hạn. Cập nhật <b>API KEY!B7</b> hoặc tab Cài đặt.</li>
        <li><b>Phản hồi cắt giữa chừng</b>: chọn cột chính xác thay vì toàn sheet để giảm context; chọn model có context lớn hơn.</li>
      </ul>
      <div class="warn"><b>⚠️ Khi mọi cách đều fail:</b> kiểm tra <code>Apps Script → Triggers</code> có trigger lỗi nào treo, kiểm tra <code>Executions</code> để đọc stack trace cụ thể.</div>` },
    { group: "API key", title: "YouTube Data API Key - API KEY!B2", html: `
      <h2>YouTube Data API Key - API KEY!B2</h2><p>Dùng cho cập nhật video/kênh, tìm kiếm theo chủ đề, lấy video trong kênh và đọc metadata công khai.</p>
      <ol><li>Mở <a href="https://console.cloud.google.com/" target="_blank">Google Cloud Console</a>.</li><li>Chọn project hiện có hoặc bấm <b>New project</b>.</li><li>Vào <b>APIs & Services</b> → <b>Library</b>.</li><li>Tìm <b>YouTube Data API v3</b> và bấm <b>Enable</b>.</li><li>Vào <b>APIs & Services</b> → <b>Credentials</b>.</li><li>Bấm <b>Create credentials</b> → <b>API key</b>.</li><li>Copy API key vừa tạo và dán vào <b>API KEY!B2</b>.</li><li>Khuyến nghị: mở key vừa tạo → <b>API restrictions</b> → <b>Restrict key</b> → chỉ chọn <b>YouTube Data API v3</b> → <b>Save</b>.</li></ol>
      <div class="links"><a target="_blank" href="https://developers.google.com/youtube/v3/getting-started">YouTube Data API</a><a target="_blank" href="https://cloud.google.com/docs/authentication/api-keys">Google Cloud API keys</a><a target="_blank" href="https://cloud.google.com/service-usage/docs/enabled-service">Enable APIs</a></div>
      <div class="videoBox"><b>Video hướng dẫn lấy YouTube Data API v3:</b><div style="position:relative;width:100%;padding-top:56.25%;margin-top:10px;border-radius:10px;overflow:hidden;background:#0f172a;"><iframe src="https://www.youtube.com/embed/dM6UufLcvZI" title="Video hướng dẫn lấy YouTube Data API v3" style="position:absolute;inset:0;width:100%;height:100%;border:0;" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe></div></div>` },
    { group: "API key", title: "OAuth Client ID/Secret - API KEY!B3:B4", html: `
      <h2>OAuth Client ID/Secret - API KEY!B3:B4</h2><p>Dùng cho nhóm <b>Kết nối kênh Analytics</b>. Code hiện dùng redirect URI <code>http://localhost</code>.</p>
      <ol><li>Trong Google Cloud, chọn project đã dùng cho YouTube API.</li><li>Bật <b>YouTube Analytics API</b> trong <b>APIs & Services</b> → <b>Library</b>.</li><li>Vào <b>OAuth consent screen</b> hoặc <b>Google Auth Platform</b>.</li><li>Điền tên app, email hỗ trợ, email nhà phát triển.</li><li>Thêm scopes: <code>https://www.googleapis.com/auth/yt-analytics.readonly</code> và <code>https://www.googleapis.com/auth/youtube.readonly</code>.</li><li>Nếu app đang testing, thêm email của bạn vào <b>Test users</b>.</li><li>Vào <b>Credentials</b> → <b>Create credentials</b> → <b>OAuth client ID</b>.</li><li>Chọn loại có Client ID và Client Secret. Nếu chọn <b>Web application</b>, thêm redirect URI <code>http://localhost</code>.</li><li>Copy <b>Client ID</b> vào <b>API KEY!B3</b>, copy <b>Client Secret</b> vào <b>API KEY!B4</b>.</li><li>Chạy menu <b>12. KẾT NỐI KÊNH ANALYTICS</b> → <b>Cài đặt Access Token</b>.</li></ol>
      <div class="warn">Google có thể chỉ hiển thị/download Client Secret lúc tạo OAuth client. Hãy lưu ngay.</div>
      <div class="links"><a target="_blank" href="https://developers.google.com/workspace/guides/configure-oauth-consent">OAuth consent</a><a target="_blank" href="https://support.google.com/cloud/answer/6158849">OAuth clients</a><a target="_blank" href="https://developers.google.com/youtube/analytics">YouTube Analytics API</a></div><div class="videoBox"><b>Video hướng dẫn lấy OAuth Client ID/Secret:</b><div style="position:relative;width:100%;padding-top:56.25%;margin-top:10px;border-radius:10px;overflow:hidden;background:#0f172a;"><iframe src="https://www.youtube.com/embed/ci8PGyXbzgE" title="Video hướng dẫn lấy OAuth Client ID/Secret" style="position:absolute;inset:0;width:100%;height:100%;border:0;" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe></div></div>` },
    { group: "API key", title: "Supadata API Key - API KEY!B5", html: `
      <h2>Supadata API Key - API KEY!B5</h2><p>Dùng cho subtitle/transcript. Code gọi <code>GET https://api.supadata.ai/v1/transcript</code> với header <code>x-api-key</code>.</p>
      <ol><li>Mở <a href="https://dash.supadata.ai" target="_blank">dash.supadata.ai</a>.</li><li>Đăng ký hoặc đăng nhập.</li><li>Chọn gói sử dụng phù hợp.</li><li>Trong dashboard, tìm mục API key.</li><li>Copy API key.</li><li>Dán vào <b>API KEY!B5</b>.</li><li>Chạy thử <b>Lấy Subtitle THEO DÒNG</b> với 1 video.</li></ol>
      <div class="links"><a target="_blank" href="https://supadata.ai/documentation/getting-started">Supadata Getting Started</a><a target="_blank" href="https://supadata.ai/documentation/get-transcript">Supadata Transcript API</a></div>` },
    { group: "API key", title: "YouTube-Transcript.io Token - API KEY!B6", html: `
      <h2>YouTube-Transcript.io Token - API KEY!B6</h2><p>Dùng làm nguồn transcript dự phòng. Code gọi <code>POST https://www.youtube-transcript.io/api/transcripts</code> với header <code>Authorization: Basic &lt;token&gt;</code>.</p>
      <ol><li>Mở <a href="https://www.youtube-transcript.io" target="_blank">youtube-transcript.io</a>.</li><li>Đăng ký hoặc đăng nhập.</li><li>Vào profile/dashboard.</li><li>Tìm <b>API token</b> hoặc <b>Basic API token</b>.</li><li>Copy token.</li><li>Dán vào <b>API KEY!B6</b>.</li><li>Chạy thử <b>Lấy Subtitle THEO DÒNG</b>.</li></ol>
      <div class="note">Tài liệu API nêu giới hạn tối đa 50 video ID mỗi request và có rate limit. Hệ thống chạy theo dòng/khoảng để giảm lỗi.</div>
      <div class="links"><a target="_blank" href="https://www.youtube-transcript.io/api">YouTube-Transcript.io API</a></div>` },
    { group: "API key", title: "9router API Key / Bearer Token - API KEY!B7", html: `
      <h2>9router API Key / Bearer Token - API KEY!B7</h2><p>Dùng cho <b>AI phân tích Sheet</b>. Hệ thống gọi endpoint OpenAI-compatible <code>/v1/chat/completions</code>; mặc định <code>http://localhost:20128/v1</code>.</p>
      <ol><li>Cài 9router: <code>npm install -g 9router</code>.</li><li>Chạy <code>9router</code> để mở dashboard.</li><li>Kết nối provider theo dashboard. 9router hỗ trợ OAuth hoặc API key theo từng provider.</li><li>Copy API key/token nếu dashboard yêu cầu.</li><li>Dán vào <b>API KEY!B7</b> hoặc nhập trong tab <b>Cài đặt</b> của <b>AI phân tích Sheet</b>.</li><li>Nếu Apps Script không gọi được localhost, dùng public URL/tunnel và điền vào <b>9router base URL</b>.</li></ol>
      <div class="links"><a target="_blank" href="https://github.com/decolua/9router">GitHub 9router</a><a target="_blank" href="https://9router.com/">9router website</a></div><div class="videoBox"><b>Video hướng dẫn sử dụng 9router:</b><div style="position:relative;width:100%;padding-top:56.25%;margin-top:10px;border-radius:10px;overflow:hidden;background:#0f172a;"><iframe src="https://www.youtube.com/embed/MqXTWzaFtu8" title="Video hướng dẫn sử dụng 9router" style="position:absolute;inset:0;width:100%;height:100%;border:0;" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe></div></div>` },
    { group: "Sheet", title: "Sheet API KEY - điền đúng ô", html: `
      <h2>Sheet API KEY</h2><table class="guideTable"><tr><th>Ô</th><th>Nội dung</th><th>Dùng cho</th></tr><tr><td>B2</td><td>YouTube Data API Key</td><td>Video, kênh, tìm kiếm, lấy video trong kênh</td></tr><tr><td>B3</td><td>Google OAuth Client ID</td><td>Analytics</td></tr><tr><td>B4</td><td>Google OAuth Client Secret</td><td>Analytics</td></tr><tr><td>B5</td><td>Supadata API Key</td><td>Subtitle ưu tiên</td></tr><tr><td>B6</td><td>YouTube-Transcript.io Token</td><td>Subtitle dự phòng</td></tr><tr><td>B7</td><td>9router API Key / Bearer Token</td><td>AI phân tích Sheet</td></tr></table><p>Nếu đổi key, thay giá trị đúng ô; hệ thống đọc lại ở lần chạy tiếp theo.</p>` },
    { group: "Video", title: "Cập nhật TẤT CẢ (Sheet VIDEO)", html: `<h2>Cập nhật TẤT CẢ (Sheet VIDEO)</h2><ol><li>Yêu cầu <b>API KEY!B2</b>.</li><li>Sheet VIDEO có link ở cột C từ dòng 2.</li><li>Chạy menu này để cập nhật tiêu đề, tags, hashtags, views, VPH, thời lượng, thumbnail, ngày đăng, mô tả.</li><li>Không đóng file khi Apps Script đang chạy.</li></ol><div class="warn">Dữ liệu lớn nên chạy theo khoảng dòng để tránh timeout/quota.</div>` },
    { group: "Video", title: "Cập nhật THEO KHOẢNG DÒNG (Sheet VIDEO)", html: `<h2>Cập nhật THEO KHOẢNG DÒNG (Sheet VIDEO)</h2><ol><li>Nhập <code>5-20</code> để chạy dòng 5 đến 20.</li><li>Nhập <code>10</code> để chạy từ dòng 10 đến dòng cuối.</li><li>Đây là cách khuyến nghị khi sheet lớn.</li></ol>` },
    { group: "Video", title: "Cập nhật THEO DÒNG (Sheet VIDEO)", html: `<h2>Cập nhật THEO DÒNG (Sheet VIDEO)</h2><ol><li>Dùng khi vừa thêm/sửa 1 link video.</li><li>Nhập số dòng, ví dụ <code>5</code>.</li><li>Hệ thống chỉ cập nhật đúng dòng đó.</li></ol>` },
    { group: "Kênh", title: "Cập nhật TẤT CẢ (Sheet KÊNH)", html: `<h2>Cập nhật TẤT CẢ (Sheet KÊNH)</h2><ol><li>Link kênh nằm ở cột C sheet <b>${SHEET_CHANNEL}</b>.</li><li>Hệ thống cập nhật tên kênh, tags, subscribers, video/tháng, views/tháng, chủ đề, mô tả, quốc gia, category và trạng thái kiếm tiền ước lượng.</li><li>Danh sách lớn nên chạy theo khoảng dòng.</li></ol><div class="note">YouTube không có API công khai xác nhận tuyệt đối trạng thái kiếm tiền của mọi kênh bất kỳ; hệ thống chỉ đánh giá theo dữ liệu/tín hiệu có thể kiểm chứng.</div>` },
    { group: "Kênh", title: "Cập nhật THEO KHOẢNG DÒNG (Sheet KÊNH)", html: `<h2>Cập nhật THEO KHOẢNG DÒNG (Sheet KÊNH)</h2><ol><li>Nhập khoảng <code>5-20</code>.</li><li>Hệ thống chỉ cập nhật kênh trong khoảng đó.</li><li>Dùng để kiểm soát quota và tránh timeout.</li></ol>` },
    { group: "Kênh", title: "Cập nhật THEO DÒNG (Sheet KÊNH)", html: `<h2>Cập nhật THEO DÒNG (Sheet KÊNH)</h2><ol><li>Nhập số dòng có link kênh ở cột C.</li><li>Cập nhật metadata cho đúng kênh đó.</li><li>Dùng để sửa nhanh một dòng lỗi.</li></ol>` },
    { group: "Kênh", title: "Lấy links video trong kênh (< 3 tháng)", html: `<h2>Lấy links video trong kênh (&lt; 3 tháng)</h2><ol><li>Nguồn: tên kênh cột B, link kênh cột C sheet <b>${SHEET_CHANNEL}</b>.</li><li>Tick chọn một hoặc nhiều kênh trong dropdown.</li><li>Có thể nhập chủ đề lọc thêm.</li><li>Mặc định: mới trong 90 ngày, thời lượng tối thiểu 00:03:10, views tối thiểu 50.000.</li><li>Chọn cách khớp chủ đề rồi bấm <b>Bắt đầu quét</b>.</li><li>Video đạt điều kiện được thêm vào cuối sheet VIDEO.</li><li>Lịch sử lưu ở sheet ẩn <b>${FETCH_VIDEO_HISTORY_SHEET}</b>.</li></ol>` },
    { group: "Tìm kiếm", title: "Tìm VIDEO theo chủ đề", html: `<h2>Tìm VIDEO theo chủ đề</h2><ol><li>Nhập chủ đề rõ ràng, ví dụ <code>psychology human behavior</code>.</li><li>Chỉnh min views, min duration, số ngày gần đây và các điều kiện có trong giao diện.</li><li>Hệ thống tìm ứng viên bằng YouTube Data API rồi lọc theo dữ liệu thật.</li><li>Kết quả hợp lệ thêm vào sheet VIDEO và có thông báo số video tìm được/thêm mới.</li></ol><div class="warn">Endpoint search thường tốn quota cao, hãy chạy có giới hạn.</div>` },
    { group: "Tìm kiếm", title: "Tìm KÊNH theo chủ đề", html: `<h2>Tìm KÊNH theo chủ đề</h2><ol><li>Nhập chủ đề kênh cần tìm.</li><li>Mặc định lọc views/tháng tối thiểu 500.000 nếu không đổi.</li><li>Có thể chỉnh số lượng kênh, subscribers, views/tháng, quốc gia hoặc điều kiện khác.</li><li>Kết quả ghi vào sheet KÊNH và thông báo số kênh tìm được/thêm mới.</li></ol>` },
    { group: "Dọn dẹp", title: "DỌN DẸP TỔNG HỢP", html: `<h2>DỌN DẸP TỔNG HỢP</h2><ol><li>Nên tạo bản sao file trước nếu dữ liệu quan trọng.</li><li>Chạy sau khi metadata đã được cập nhật mới.</li><li>Hệ thống chạy nhiều bước dọn video/kênh liên tiếp.</li></ol>` },
    { group: "Dọn dẹp", title: "Xóa video > 3 tháng & trùng lặp", html: `<h2>Xóa video &gt; 3 tháng & trùng lặp</h2><ol><li>Kiểm tra cột ngày đăng trước khi chạy.</li><li>Chạy sau khi cập nhật metadata video.</li><li>Hệ thống loại video cũ và link trùng.</li></ol>` },
    { group: "Dọn dẹp", title: "Xóa video < 20.000 views", html: `<h2>Xóa video &lt; 20.000 views</h2><ol><li>Cập nhật views trước.</li><li>Chạy để giữ lại video có hiệu suất tối thiểu.</li><li>Kiểm tra lại sheet sau khi chạy.</li></ol>` },
    { group: "Dọn dẹp", title: "Xóa kênh chết (0 views/tháng)", html: `<h2>Xóa kênh chết (0 views/tháng)</h2><ol><li>Cập nhật views/tháng của sheet KÊNH trước.</li><li>Hệ thống loại kênh không có views/tháng theo dữ liệu đang có.</li><li>Kiểm tra dữ liệu cũ/thiếu trước khi xóa hàng loạt.</li></ol>` },
    { group: "Cập nhật nhanh", title: "Cập nhật NHANH Views & VPH (Sheet VIDEO)", html: `<h2>Cập nhật NHANH Views & VPH</h2><ol><li>Dùng để làm mới views hiện tại và VPH.</li><li>Nhẹ hơn cập nhật full metadata.</li><li>Phù hợp theo dõi hằng ngày.</li></ol>` },
    { group: "Cập nhật nhanh", title: "Cập nhật NHANH Subs & Views/Tháng (Sheet KÊNH)", html: `<h2>Cập nhật NHANH Subs & Views/Tháng</h2><ol><li>Làm mới subscribers và views/tháng.</li><li>Không làm lại toàn bộ mô tả/chủ đề nếu không cần.</li><li>Phù hợp lọc kênh tiềm năng nhanh.</li></ol>` },
    { group: "Quota", title: "Xem Quota YouTube API hôm nay", html: `<h2>Xem Quota YouTube API hôm nay</h2><ol><li>Mở sau tác vụ lớn như tìm kiếm hoặc cập nhật toàn bộ.</li><li>Nếu quota gần hết, chuyển sang khoảng dòng nhỏ.</li><li>YouTube Data API thường có quota mặc định 10.000 units/ngày cho project đã bật API.</li></ol><div class="links"><a target="_blank" href="https://developers.google.com/youtube/v3/getting-started">Nguồn quota YouTube Data API</a></div>` },
    { group: "AI", title: "AI phân tích Sheet", html: `<h2>AI phân tích Sheet</h2><ol><li>Điền 9router key ở <b>API KEY!B7</b> hoặc tab <b>Cài đặt</b>.</li><li>Chọn model, sheet nguồn, toàn bộ sheet hoặc các cột cần phân tích.</li><li>Nhập câu hỏi cụ thể.</li><li>Nếu muốn dùng lịch sử, tick <b>Kết hợp lịch sử đã phân tích</b>, chọn lịch sử và xem preview.</li><li>Kết quả thành công lưu vào sheet ẩn <b>${AI_MEMORY_SHEET}</b>; lỗi không lưu.</li></ol>` },
    { group: "Subtitle", title: "Hướng dẫn cấu hình API Subtitle", html: `<h2>Hướng dẫn cấu hình API Subtitle</h2><ol><li>Ưu tiên cấu hình <b>API KEY!B5</b> Supadata.</li><li>Nếu có thêm <b>API KEY!B6</b> YouTube-Transcript.io thì ổn định hơn.</li><li>Nếu không có B5/B6, hệ thống fallback timedtext/scrape nhưng dễ lỗi 429.</li></ol>` },
    { group: "Subtitle", title: "Lấy Subtitle TẤT CẢ video", html: `<h2>Lấy Subtitle TẤT CẢ video</h2><ol><li>Sheet VIDEO có link ở cột C.</li><li>Nên cấu hình B5/B6 trước.</li><li>Chạy để lấy transcript toàn bộ video.</li></ol><div class="warn">Video nhiều nên chạy theo khoảng dòng.</div>` },
    { group: "Subtitle", title: "Lấy Subtitle THEO KHOẢNG DÒNG", html: `<h2>Lấy Subtitle THEO KHOẢNG DÒNG</h2><ol><li>Nhập khoảng <code>5-20</code>.</li><li>Hệ thống lấy transcript cho video trong khoảng đó.</li><li>Khuyến nghị khi danh sách dài.</li></ol>` },
    { group: "Subtitle", title: "Lấy Subtitle THEO DÒNG", html: `<h2>Lấy Subtitle THEO DÒNG</h2><ol><li>Nhập số dòng có link video.</li><li>Hệ thống chỉ lấy subtitle cho video đó.</li><li>Dùng để kiểm tra API key hoặc sửa một dòng lỗi.</li></ol>` },
    { group: "Subtitle", title: "Thử lại các dòng FAIL", html: `<h2>Thử lại các dòng FAIL</h2><ol><li>Hệ thống quét cột subtitle để tìm dòng lỗi.</li><li>Chỉ chạy lại các dòng fail trước đó.</li><li>Nên cấu hình B5/B6 trước khi thử lại.</li></ol>` },
    { group: "Subtitle", title: "Xóa Subtitle (Cột M)", html: `<h2>Xóa Subtitle (Cột M)</h2><ol><li>Dùng khi muốn làm sạch transcript hiện có.</li><li>Backup nếu cần giữ dữ liệu cũ.</li><li>Sau khi xóa có thể lấy lại subtitle.</li></ol>` },
    { group: "Analytics", title: "Kết nối kênh Analytics - Cài đặt Access Token", html: `<h2>Kết nối kênh Analytics</h2><ol><li>Điền B3/B4.</li><li>Chạy <b>Cài đặt Access Token</b>.</li><li>Mở link cấp quyền.</li><li>Đăng nhập đúng tài khoản quản lý kênh YouTube.</li><li>Cho phép quyền YouTube readonly và YouTube Analytics readonly.</li><li>Copy code hoặc URL redirect về <code>http://localhost</code> và dán lại vào hộp thoại.</li><li>Hệ thống lưu refresh token để dùng lại.</li></ol>` },
    { group: "Analytics", title: "Xóa Token Analytics", html: `<h2>Xóa Token Analytics</h2><ol><li>Dùng khi muốn đổi tài khoản/kênh.</li><li>Sau khi xóa, báo cáo Analytics yêu cầu kết nối lại.</li><li>Chạy lại <b>Cài đặt Access Token</b>.</li></ol>` },
    { group: "Analytics", title: "Analytics TỔNG QUAN KÊNH", html: `<h2>Analytics TỔNG QUAN KÊNH</h2><ol><li>Yêu cầu đã kết nối Analytics.</li><li>Chọn 7 ngày, 28 ngày, toàn thời gian hoặc cập nhật tất cả.</li><li>Dữ liệu trả về sheet <b>${SHEET_ANALYTICS}</b>.</li><li>Nếu doanh thu không trả về, có thể do quyền hoặc metric không có trong trường hợp hiện tại.</li></ol>` },
    { group: "Analytics", title: "Analytics TỪNG VIDEO", html: `<h2>Analytics TỪNG VIDEO</h2><ol><li>Yêu cầu kết nối Analytics.</li><li>Chọn 7 ngày, 28 ngày hoặc toàn thời gian.</li><li>Dữ liệu ghi vào <b>${SHEET_VIDEO_ANALYTICS}</b> khi module chi tiết được triển khai đầy đủ.</li></ol>` }
  ];
}

function showSystemBrandInfo() {
  const guideSections = _getGuideSectionsArray_();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:18px;background:#f8fafc;color:#202124;line-height:1.55;">
      ${buildDialogHeader_("Trung tâm hướng dẫn hệ thống", "Hướng dẫn từng bước cho API key, OAuth, sheet dữ liệu và toàn bộ chức năng trong YouTube Tools.", "#111827")}
      <style>
        *{box-sizing:border-box;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif}
        .guideTop{display:grid;grid-template-columns:320px 1fr;gap:14px;margin-bottom:14px}
        .guideSelect,.guideSearch{width:100%;box-sizing:border-box;padding:11px 14px;border:1px solid #d1d5db;border-radius:12px;background:#ffffff;font-size:13.5px;color:#0f172a;box-shadow:0 1px 2px rgba(15,23,42,0.04),inset 0 1px 0 rgba(255,255,255,0.8);transition:all .2s ease;outline:none}
        .guideSelect:hover,.guideSearch:hover{border-color:#94a3b8}
        .guideSelect:focus,.guideSearch:focus{border-color:#6366f1;box-shadow:0 0 0 4px rgba(99,102,241,0.14)}
        .guideSearch::placeholder{color:#94a3b8}
        .guideBody{display:grid;grid-template-columns:280px minmax(0,1fr);gap:14px}
        .guideNav{background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%);border:1px solid #e5e7eb;border-radius:14px;padding:10px;max-height:640px;overflow:auto;box-shadow:0 8px 24px -10px rgba(15,23,42,0.10),inset 0 1px 0 #ffffff}
        .guideNav::-webkit-scrollbar,.guideContent::-webkit-scrollbar{width:8px;height:8px}
        .guideNav::-webkit-scrollbar-thumb,.guideContent::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#cbd5e1,#94a3b8);border-radius:8px}
        .guideNav::-webkit-scrollbar-thumb:hover,.guideContent::-webkit-scrollbar-thumb:hover{background:linear-gradient(180deg,#94a3b8,#64748b)}
        .guideNav::-webkit-scrollbar-track,.guideContent::-webkit-scrollbar-track{background:transparent}
        .guideBtn{display:block;width:100%;text-align:left;padding:10px 12px;border:0;background:transparent;border-radius:10px;cursor:pointer;font-size:12.5px;color:#334155;font-weight:500;transition:all .18s ease;line-height:1.35;margin:1px 0;position:relative}
        .guideBtn:hover{background:#eef2ff;color:#1e293b;transform:translateX(2px)}
        .guideBtn.active{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);color:#fff;font-weight:700;box-shadow:0 4px 12px rgba(15,23,42,0.32),inset 0 1px 0 rgba(255,255,255,0.10)}
        .guideBtn.active::before{content:"";position:absolute;left:0;top:8px;bottom:8px;width:3px;background:linear-gradient(180deg,#6366f1,#8b5cf6);border-radius:3px}
        .guideGroup{font-size:10.5px;text-transform:uppercase;color:#64748b;font-weight:800;margin:14px 10px 6px;letter-spacing:0.08em;display:flex;align-items:center;gap:8px}
        .guideGroup::after{content:"";flex:1;height:1px;background:linear-gradient(90deg,#e2e8f0,transparent)}
        .guideContent{background:linear-gradient(180deg,#ffffff 0%,#fafbfd 100%);border:1px solid #e5e7eb;border-radius:14px;padding:24px 26px;max-height:640px;overflow:auto;box-shadow:0 10px 30px -12px rgba(15,23,42,0.12),inset 0 1px 0 #ffffff;color:#1f2937;font-size:14px;line-height:1.65}
        .guideContent h2{margin:0 0 14px;font-size:24px;color:#0f172a;font-weight:800;letter-spacing:-0.015em;background:linear-gradient(135deg,#0f172a 0%,#334155 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;padding-bottom:10px;border-bottom:2px solid #f1f5f9}
        .guideContent h3{margin:18px 0 10px;font-size:16px;color:#1e293b;font-weight:700}
        .guideContent p{margin:8px 0;color:#334155}
        .guideContent ol,.guideContent ul{padding-left:22px;margin:10px 0}
        .guideContent li{margin:7px 0;color:#334155}
        .guideContent li::marker{color:#6366f1;font-weight:700}
        .guideContent b{color:#0f172a;font-weight:700}
        .guideContent code{background:linear-gradient(180deg,#f1f5f9,#e2e8f0);border:1px solid #cbd5e1;border-radius:6px;padding:2px 7px;font-family:'SF Mono','Monaco','Cascadia Code','Roboto Mono',Consolas,monospace;font-size:12.5px;color:#0f172a;font-weight:600}
        .guideContent a{color:#4f46e5;text-decoration:none;border-bottom:1px dashed #a5b4fc;transition:all .2s}
        .guideContent a:hover{color:#3730a3;border-bottom-color:#4f46e5}
        .guideTable{border-collapse:separate;border-spacing:0;width:100%;font-size:13px;border-radius:10px;overflow:hidden;box-shadow:0 4px 12px rgba(15,23,42,0.06);margin:12px 0}
        .guideTable th,.guideTable td{border:1px solid #e5e7eb;padding:10px 12px;text-align:left}
        .guideTable th{background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);color:#fff;font-weight:700;letter-spacing:0.02em;border-color:#1e293b}
        .guideTable tr:nth-child(even) td{background:#f8fafc}
        .guideTable tr:hover td{background:#eef2ff}
        .note{background:linear-gradient(135deg,#ecfdf5 0%,#d1fae5 100%);border-left:4px solid #10b981;padding:12px 14px;border-radius:10px;margin:12px 0;color:#064e3b;box-shadow:0 2px 6px rgba(16,185,129,0.10)}
        .warn{background:linear-gradient(135deg,#fff7ed 0%,#ffedd5 100%);border-left:4px solid #f97316;padding:12px 14px;border-radius:10px;margin:12px 0;color:#7c2d12;box-shadow:0 2px 6px rgba(249,115,22,0.10)}
        .tip{background:linear-gradient(135deg,#eff6ff 0%,#dbeafe 100%);border-left:4px solid #3b82f6;padding:12px 14px;border-radius:10px;margin:12px 0;color:#1e3a8a;box-shadow:0 2px 6px rgba(59,130,246,0.10)}
        .videoBox{background:linear-gradient(135deg,#f8fafc 0%,#eef2ff 100%);border:1px solid #e0e7ff;border-radius:14px;padding:14px;margin:14px 0;box-shadow:0 6px 18px -8px rgba(99,102,241,0.20)}
        .links{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}
        .links a{display:inline-flex;align-items:center;gap:6px;background:linear-gradient(135deg,#eef2ff 0%,#e0e7ff 100%);color:#3730a3;text-decoration:none;border:1px solid #c7d2fe;border-radius:999px;padding:7px 13px;font-size:12px;font-weight:700;transition:all .2s ease;box-shadow:0 1px 2px rgba(79,70,229,0.08)}
        .links a:hover{background:linear-gradient(135deg,#4f46e5 0%,#6366f1 100%);color:#fff;border-color:#4f46e5;transform:translateY(-1px);box-shadow:0 4px 10px rgba(79,70,229,0.32)}
      </style>
      <div class="guideTop"><select id="guideSelect" class="guideSelect" onchange="selectGuide(this.value)"></select><input id="guideSearch" class="guideSearch" oninput="renderGuideNav()" placeholder="Tìm hướng dẫn theo chức năng, API key, OAuth, subtitle, analytics..."></div>
      <div class="guideBody"><div id="guideNav" class="guideNav"></div><div id="guideContent" class="guideContent"></div></div>
      <script>
        const GUIDE_SECTIONS = ${JSON.stringify(guideSections)}; let activeGuide = 0;
        function esc(s){return (s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
        function renderGuideSelect(){document.getElementById('guideSelect').innerHTML=GUIDE_SECTIONS.map((g,i)=>'<option value="'+i+'">'+esc(g.group+' - '+g.title)+'</option>').join('');}
        /* guideFuzzyMatch: prefix-token matching cho ô tìm kiếm hướng dẫn */
        function guideFuzzyMatch(text, rawQ){
          if (!rawQ) return true;
          const t = (text||'').toLowerCase();
          const q = rawQ.toLowerCase().trim();
          if (!q) return true;
          if (t.indexOf(q) >= 0) return true;
          const tokens = q.split(/\s+/).filter(Boolean);
          const words = t.split(/[\s,.\-\/|]+/).filter(Boolean);
          if (tokens.length < 2) return words.some(w => w.indexOf(tokens[0])===0);
          return tokens.every(tok => t.indexOf(tok)>=0 || words.some(w => w.indexOf(tok)===0));
        }
        function renderGuideNav(){const q=(document.getElementById('guideSearch').value||'').trim();let last='';const html=GUIDE_SECTIONS.map((g,i)=>({g,i})).filter(x=>!q||guideFuzzyMatch(x.g.group+' '+x.g.title+' '+x.g.html.replace(/<[^>]+>/g,' '),q)).map(x=>{const h=x.g.group!==last?'<div class="guideGroup">'+esc(x.g.group)+'</div>':'';last=x.g.group;return h+'<button class="guideBtn '+(x.i===activeGuide?'active':'')+'" onclick="selectGuide('+x.i+')">'+esc(x.g.title)+'</button>';}).join('');document.getElementById('guideNav').innerHTML=html||'<div style="padding:10px;color:#64748b;">Không tìm thấy hướng dẫn phù hợp.</div>';}
        function selectGuide(i){activeGuide=parseInt(i,10)||0;document.getElementById('guideSelect').value=String(activeGuide);var gc=document.getElementById('guideContent');gc.innerHTML=GUIDE_SECTIONS[activeGuide].html;gc.scrollTop=0;renderGuideNav();if(window.dlgProgress){window.dlgProgress.set(0,'Đang đọc: '+GUIDE_SECTIONS[activeGuide].title);}}
        renderGuideSelect();selectGuide(0);
        setTimeout(function(){ if(window.dlgProgress) window.dlgProgress.attachScroll('guideContent'); }, 60);
        document.getElementById('guideSearch').addEventListener('keydown', function(e){ if(e.key==='Escape'){ this.value=''; renderGuideNav(); }});
      </script>
    </div>
  `).setWidth(1180).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, "Trung tâm hướng dẫn hệ thống");
}
function promptForRange(ui, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) { ui.alert("Không tìm thấy Sheet: " + sheetName); return null; }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert("Sheet hiện tại không có dữ liệu để cập nhật."); return null; }

  const response = ui.prompt(
    '⚙️ Cập nhật theo KHOẢNG DÒNG',
    `Sheet hiện tại có dữ liệu đến dòng ${lastRow}.\n\n👉 CÁCH NHẬP:\n- Nhập khoảng: "5-20" (Chạy từ dòng 5 đến 20)\n- Nhập 1 số: "10" (Chạy từ dòng 10 đến dòng cuối cùng)`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const text = response.getResponseText().trim();
    let startRow, endRow;

    if (text.includes('-')) {
      const parts = text.split('-');
      startRow = parseInt(parts[0].trim());
      endRow = parseInt(parts[1].trim());
    } else {
      startRow = parseInt(text);
      endRow = lastRow;
    }

    if (isNaN(startRow) || startRow < 2) { ui.alert("❌ Số dòng bắt đầu không hợp lệ (Phải từ dòng 2 trở đi)."); return null; }
    if (isNaN(endRow) || endRow < startRow) { ui.alert("❌ Số dòng kết thúc không hợp lệ."); return null; }

    if (endRow > lastRow) endRow = lastRow;
    return { start: startRow, end: endRow };
  }
  return null;
}

// =========================================================================
// LÕI XỬ LÝ API - CACHE & RETRY
// =========================================================================
function openRangeActionDialog_(action, sheetName, title, subtitle, accentColor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return SpreadsheetApp.getUi().alert("Không tìm thấy Sheet: " + sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return SpreadsheetApp.getUi().alert("Sheet hiện tại không có dữ liệu để xử lý.");
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:18px;background:#f8fafc;color:#202124;line-height:1.45;">
      ${buildDialogHeader_(title, subtitle, accentColor)}
      <div style="background:#fff;border:1px solid #dfe3eb;border-radius:10px;padding:14px;">
        <div style="font-size:13px;color:#4b5563;margin-bottom:10px;">Sheet hiện tại có dữ liệu đến dòng <b>${lastRow}</b>.</div>
        <label style="display:block;font-weight:bold;margin-bottom:6px;">Khoảng dòng cần chạy</label>
        <input id="rangeText" value="2-${lastRow}" placeholder="VD: 5-20 hoặc 10" style="width:100%;box-sizing:border-box;padding:11px;border:2px solid ${accentColor || "#1a73e8"};border-radius:8px;font-size:14px;">
        <div style="font-size:12px;color:#6b7280;margin-top:8px;">Nhập <b>5-20</b> để chạy từ dòng 5 đến 20, hoặc nhập <b>10</b> để chạy từ dòng 10 đến dòng cuối.</div>
        <div id="status" style="display:none;margin-top:12px;background:#f8fafc;border:1px solid #dfe3eb;border-radius:8px;padding:10px;font-size:13px;white-space:pre-wrap;"></div>
      </div>
      <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px;">
        <button onclick="google.script.host.close()" style="padding:10px 18px;background:#9ca3af;color:#fff;border:none;border-radius:8px;cursor:pointer;">Hủy</button>
        <button id="runBtn" onclick="run()" style="padding:10px 22px;background:${accentColor || "#1a73e8"};color:#fff;border:none;border-radius:8px;font-weight:bold;cursor:pointer;">Bắt đầu</button>
      </div>
      <script>
        function parseRange(raw){
          raw = (raw || '').trim();
          if (!raw) throw new Error('Vui lòng nhập khoảng dòng.');
          if (raw.indexOf('-') >= 0) {
            const p = raw.split('-');
            return { start: parseInt(p[0].trim(), 10), end: parseInt(p[1].trim(), 10) };
          }
          return { start: parseInt(raw, 10), end: ${lastRow} };
        }
        function setStatus(text, isErr){
          const el = document.getElementById('status');
          el.style.display = 'block';
          el.style.borderColor = isErr ? '#fecaca' : '#bfdbfe';
          el.style.background = isErr ? '#fff7f7' : '#eff6ff';
          el.textContent = text;
        }
        function run(){
          let range;
          try { range = parseRange(document.getElementById('rangeText').value); }
          catch(e) { setStatus(e.message, true); if(window.dlgProgress) window.dlgProgress.fail(e.message); return; }
          if (!range.start || !range.end || range.start < 2 || range.end < range.start || range.end > ${lastRow}) {
            setStatus('Khoảng dòng không hợp lệ. Dòng bắt đầu phải từ 2 và dòng kết thúc không vượt quá ${lastRow}.', true);
            if(window.dlgProgress) window.dlgProgress.fail('Khoảng dòng không hợp lệ');
            return;
          }
          var rows = (range.end - range.start + 1);
          var eta = Math.max(2, rows * 0.6);
          document.getElementById('runBtn').disabled = true;
          setStatus('Đang xử lý dòng ' + range.start + ' đến ' + range.end + ' (' + rows + ' dòng)...', false);
          if(window.dlgProgress) window.dlgProgress.start(eta, 'Đang xử lý ' + rows + ' dòng (' + range.start + '-' + range.end + ')');
          google.script.run.withSuccessHandler(function(res){
            setStatus((res.success ? 'Hoàn tất: ' : 'Lỗi: ') + res.message, !res.success);
            document.getElementById('runBtn').disabled = !!res.success;
            if(window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất ' + rows + ' dòng') : window.dlgProgress.fail('Lỗi: ' + res.message); }
          }).withFailureHandler(function(err){
            setStatus('Lỗi: ' + (err.message || err), true);
            document.getElementById('runBtn').disabled = false;
            if(window.dlgProgress) window.dlgProgress.fail('Lỗi: ' + (err.message || err));
          }).executeRangeAction({ action: "${action}", sheetName: "${sheetName}", start: range.start, end: range.end });
        }
        document.getElementById('rangeText').addEventListener('keydown', function(e){ if(e.key==='Enter'){ e.preventDefault(); run(); } });
      </script>
    </div>
  `).setWidth(620).setHeight(490);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function openSingleRowActionDialog_(action, sheetName, title, subtitle, accentColor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return SpreadsheetApp.getUi().alert("Không tìm thấy Sheet: " + sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return SpreadsheetApp.getUi().alert("Sheet hiện tại không có dữ liệu để xử lý.");
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:18px;background:#f8fafc;color:#202124;line-height:1.45;">
      ${buildDialogHeader_(title, subtitle, accentColor)}
      <div style="background:#fff;border:1px solid #dfe3eb;border-radius:10px;padding:14px;">
        <div style="font-size:13px;color:#4b5563;margin-bottom:10px;">Sheet hiện tại có dữ liệu đến dòng <b>${lastRow}</b>.</div>
        <label style="display:block;font-weight:bold;margin-bottom:6px;">Số dòng cần chạy</label>
        <input id="rowText" type="number" min="2" max="${lastRow}" placeholder="VD: 5" style="width:100%;box-sizing:border-box;padding:11px;border:2px solid ${accentColor || "#1a73e8"};border-radius:8px;font-size:14px;">
        <div id="status" style="display:none;margin-top:12px;background:#f8fafc;border:1px solid #dfe3eb;border-radius:8px;padding:10px;font-size:13px;white-space:pre-wrap;"></div>
      </div>
      <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px;">
        <button onclick="google.script.host.close()" style="padding:10px 18px;background:#9ca3af;color:#fff;border:none;border-radius:8px;cursor:pointer;">Hủy</button>
        <button id="runBtn" onclick="run()" style="padding:10px 22px;background:${accentColor || "#1a73e8"};color:#fff;border:none;border-radius:8px;font-weight:bold;cursor:pointer;">Bắt đầu</button>
      </div>
      <script>
        function setStatus(text, isErr){
          const el = document.getElementById('status');
          el.style.display = 'block';
          el.style.borderColor = isErr ? '#fecaca' : '#bfdbfe';
          el.style.background = isErr ? '#fff7f7' : '#eff6ff';
          el.textContent = text;
        }
        function run(){
          const row = parseInt(document.getElementById('rowText').value, 10);
          if (!row || row < 2 || row > ${lastRow}) {
            setStatus('Số dòng không hợp lệ. Vui lòng nhập từ 2 đến ${lastRow}.', true);
            if(window.dlgProgress) window.dlgProgress.fail('Số dòng không hợp lệ');
            return;
          }
          document.getElementById('runBtn').disabled = true;
          setStatus('Đang xử lý dòng ' + row + '...', false);
          if(window.dlgProgress) window.dlgProgress.start(2, 'Đang xử lý dòng ' + row);
          google.script.run.withSuccessHandler(function(res){
            setStatus((res.success ? 'Hoàn tất: ' : 'Lỗi: ') + res.message, !res.success);
            document.getElementById('runBtn').disabled = !!res.success;
            if(window.dlgProgress){ res.success ? window.dlgProgress.complete('Hoàn tất dòng ' + row) : window.dlgProgress.fail('Lỗi: ' + res.message); }
          }).withFailureHandler(function(err){
            setStatus('Lỗi: ' + (err.message || err), true);
            document.getElementById('runBtn').disabled = false;
            if(window.dlgProgress) window.dlgProgress.fail('Lỗi: ' + (err.message || err));
          }).executeSingleRowAction({ action: "${action}", sheetName: "${sheetName}", row: row });
        }
        document.getElementById('rowText').addEventListener('keydown', function(e){ if(e.key==='Enter'){ e.preventDefault(); run(); } });
      </script>
    </div>
  `).setWidth(560).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function executeRangeAction(params) {
  try {
    const action = params && params.action;
    const sheetName = params && params.sheetName;
    const start = parseInt(params.start, 10);
    const end = parseInt(params.end, 10);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error("Không tìm thấy Sheet: " + sheetName);
    if (!start || !end || start < 2 || end < start || end > sheet.getLastRow()) throw new Error("Khoảng dòng không hợp lệ.");
    const totalRows = end - start + 1;
    initTaskProgress_(params.progressId, totalRows, "Chuẩn bị xử lý " + totalRows + " dòng");
    if (action === "videoRange") {
      executeVideoUpdate(sheet, start, end, "full", params.progressId);
      finishTaskProgress_(params.progressId, true, "Hoàn tất " + totalRows + " dòng video");
      return { success: true, message: `Đã cập nhật Video dòng ${start}-${end}.\n\n${getYouTubeQuotaUsageMessage_()}` };
    }
    if (action === "channelRange") {
      executeChannelUpdate(sheet, start, end, "full", params.progressId);
      finishTaskProgress_(params.progressId, true, "Hoàn tất " + totalRows + " dòng kênh");
      return { success: true, message: `Đã cập nhật Kênh dòng ${start}-${end}.\n\n${getYouTubeQuotaUsageMessage_()}` };
    }
    if (action === "subtitleRange") {
      const report = executeSubtitleFetch(sheet, start, end, true, params.progressId);
      finishTaskProgress_(params.progressId, true, "Hoàn tất lấy subtitle " + totalRows + " dòng");
      return { success: true, message: report || `Đã lấy Subtitle dòng ${start}-${end}.` };
    }
    throw new Error("Action không hợp lệ.");
  } catch (e) {
    finishTaskProgress_(params && params.progressId, false, "Lỗi: " + e.message);
    return { success: false, message: e.message };
  }
}

function executeSingleRowAction(params) {
  try {
    const action = params && params.action;
    const sheetName = params && params.sheetName;
    const row = parseInt(params.row, 10);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error("Không tìm thấy Sheet: " + sheetName);
    if (!row || row < 2 || row > sheet.getLastRow()) throw new Error("Số dòng không hợp lệ.");
    initTaskProgress_(params.progressId, 1, "Đang xử lý dòng " + row);
    if (action === "videoRow" || action === "videoSingle") {
      executeVideoUpdate(sheet, row, row, "full", params.progressId);
      finishTaskProgress_(params.progressId, true, "Hoàn tất dòng " + row);
      return { success: true, message: `Đã cập nhật Video dòng ${row}.` };
    }
    if (action === "channelRow" || action === "channelSingle") {
      executeChannelUpdate(sheet, row, row, "full", params.progressId);
      finishTaskProgress_(params.progressId, true, "Hoàn tất dòng " + row);
      return { success: true, message: `Đã cập nhật Kênh dòng ${row}.` };
    }
    if (action === "subtitleRow" || action === "subtitleSingle") {
      const report = executeSubtitleFetch(sheet, row, row, true, params.progressId);
      finishTaskProgress_(params.progressId, true, "Hoàn tất dòng " + row);
      return { success: true, message: report || `Đã lấy Subtitle dòng ${row}.` };
    }
    throw new Error("Action không hợp lệ.");
  } catch (e) {
    finishTaskProgress_(params && params.progressId, false, "Lỗi: " + e.message);
    return { success: false, message: e.message };
  }
}

function getQuotaDateKey_() {
  return Utilities.formatDate(new Date(), "America/Los_Angeles", "yyyy-MM-dd");
}

function estimateYouTubeQuotaCost_(url) {
  if (!url || !url.includes("youtube/v3/")) return 0;
  const match = url.match(/youtube\/v3\/([a-zA-Z]+)/);
  if (!match) return 1;
  const resource = match[1];
  if (resource === "search") return 100;
  if (resource === "videos" || resource === "channels" || resource === "playlistItems" || resource === "videoCategories") return 1;
  return 1;
}

function getYouTubeQuotaUsage_() {
  const props = PropertiesService.getDocumentProperties();
  const today = getQuotaDateKey_();
  const storedDate = props.getProperty('YT_QUOTA_DATE');
  if (storedDate !== today) {
    props.setProperty('YT_QUOTA_DATE', today);
    props.setProperty('YT_QUOTA_USED', '0');
    props.setProperty('YT_QUOTA_REQUESTS', '0');
    return { date: today, used: 0, requests: 0, limit: YOUTUBE_DAILY_QUOTA_LIMIT };
  }
  return {
    date: today,
    used: parseInt(props.getProperty('YT_QUOTA_USED') || '0', 10) || 0,
    requests: parseInt(props.getProperty('YT_QUOTA_REQUESTS') || '0', 10) || 0,
    limit: YOUTUBE_DAILY_QUOTA_LIMIT
  };
}

function recordYouTubeQuota_(url) {
  const cost = estimateYouTubeQuotaCost_(url);
  if (cost <= 0) return;
  const usage = getYouTubeQuotaUsage_();
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('YT_QUOTA_USED', String(usage.used + cost));
  props.setProperty('YT_QUOTA_REQUESTS', String(usage.requests + 1));
}

function getYouTubeQuotaUsageMessage_() {
  const usage = getYouTubeQuotaUsage_();
  const percent = usage.limit ? Math.min(999, (usage.used / usage.limit) * 100) : 0;
  const remaining = Math.max(0, usage.limit - usage.used);
  return `Quota YouTube API hôm nay (reset theo giờ Pacific): ${formatNumber(usage.used)}/${formatNumber(usage.limit)} units (${percent.toFixed(1)}%). Còn khoảng ${formatNumber(remaining)} units. Số request đã ghi nhận: ${formatNumber(usage.requests)}.`;
}

function getQuotaSnapshot_() {
  const usage = getYouTubeQuotaUsage_();
  const percent = usage.limit ? Math.min(999, (usage.used / usage.limit) * 100) : 0;
  const remaining = Math.max(0, usage.limit - usage.used);
  const nowPT = Utilities.formatDate(new Date(), "America/Los_Angeles", "yyyy-MM-dd HH:mm:ss");
  const ptDate = new Date(Utilities.formatDate(new Date(), "America/Los_Angeles", "yyyy-MM-dd'T'HH:mm:ss"));
  const secondsToMidnightPT = Math.max(0, 86400 - (ptDate.getHours() * 3600 + ptDate.getMinutes() * 60 + ptDate.getSeconds()));
  const hoursLeft = Math.floor(secondsToMidnightPT / 3600);
  const minsLeft = Math.floor((secondsToMidnightPT % 3600) / 60);
  let level = 'safe', tip = 'Quota dồi dào, có thể chạy tự do.';
  if (percent >= 80) { level = 'danger'; tip = 'Quota gần cạn — tránh dùng search.list, ưu tiên Cập nhật NHANH hoặc theo khoảng nhỏ.'; }
  else if (percent >= 50) { level = 'warn'; tip = 'Quota ở mức trung bình — chia khoảng nhỏ và tránh chạy search lan man.'; }
  return {
    used: usage.used,
    limit: usage.limit,
    remaining: remaining,
    requests: usage.requests,
    percent: percent,
    date: usage.date,
    nowPT: nowPT,
    resetIn: hoursLeft + 'h ' + minsLeft + 'm',
    secondsToReset: secondsToMidnightPT,
    level: level,
    tip: tip,
    avgCostPerReq: usage.requests > 0 ? (usage.used / usage.requests).toFixed(2) : '0.00'
  };
}

function showYouTubeQuotaUsage() {
  const snap = getQuotaSnapshot_();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:18px;background:#f8fafc;color:#202124;line-height:1.5;">
      ${buildDialogHeader_("📊 Quota YouTube Data API hôm nay", "Theo dõi units đã dùng / còn lại theo giờ Pacific (reset 0:00 PT mỗi ngày).", "#7c3aed")}
      <div class="dlgFadeBody">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
          <div style="background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:18px;text-align:center;box-shadow:0 6px 18px -10px rgba(15,23,42,0.18);">
            <div style="position:relative;width:160px;height:160px;margin:0 auto;">
              <svg id="ringSvg" viewBox="0 0 120 120" style="width:100%;height:100%;transform:rotate(-90deg);">
                <circle cx="60" cy="60" r="52" fill="none" stroke="#eef2f7" stroke-width="14"/>
                <circle id="ringFg" cx="60" cy="60" r="52" fill="none" stroke="url(#ringGrad)" stroke-width="14" stroke-linecap="round" stroke-dasharray="326.7" stroke-dashoffset="326.7" style="transition:stroke-dashoffset 1s cubic-bezier(.2,.7,.2,1);"/>
                <defs><linearGradient id="ringGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#7c3aed"/><stop offset="100%" stop-color="#ec4899"/></linearGradient></defs>
              </svg>
              <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;">
                <div id="ringPct" style="font-size:30px;font-weight:900;color:#0f172a;letter-spacing:-0.02em;">0%</div>
                <div id="ringLabel" style="font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;">Đã dùng</div>
              </div>
            </div>
            <div id="quotaTip" style="margin-top:14px;font-size:12.5px;color:#475569;font-weight:600;line-height:1.5;padding:10px 12px;border-radius:10px;background:#f8fafc;border:1px solid #e2e8f0;"></div>
          </div>
          <div style="display:flex;flex-direction:column;gap:10px;">
            <div id="cUsed" class="qChip"></div>
            <div id="cRemain" class="qChip"></div>
            <div id="cReq" class="qChip"></div>
            <div id="cAvg" class="qChip"></div>
            <div id="cReset" class="qChip"></div>
          </div>
        </div>
        <div style="margin-top:14px;background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:14px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;font-size:12px;font-weight:700;color:#475569;letter-spacing:0.04em;text-transform:uppercase;">
            <span>Mức tiêu thụ</span><span id="barPct" style="color:#7c3aed;">0%</span>
          </div>
          <div style="position:relative;height:14px;background:linear-gradient(180deg,#eef2f7,#e2e8f0);border-radius:8px;overflow:hidden;border:1px solid #d8dee7;">
            <div id="barFill" style="position:absolute;left:0;top:0;bottom:0;width:0%;background:linear-gradient(90deg,#10b981,#7c3aed,#ec4899);background-size:200% 100%;animation:dlgShimmer 2s linear infinite;border-radius:8px;transition:width 1s cubic-bezier(.2,.7,.2,1);"></div>
            <div style="position:absolute;left:50%;top:0;bottom:0;width:1px;background:rgba(15,23,42,0.18);"></div>
            <div style="position:absolute;left:80%;top:0;bottom:0;width:1px;background:rgba(220,38,38,0.4);"></div>
          </div>
          <div style="display:flex;justify-content:space-between;font-size:10.5px;color:#94a3b8;margin-top:5px;font-weight:600;">
            <span>0</span><span>50% (mốc cảnh báo)</span><span>80% (mốc nguy)</span><span>${YOUTUBE_DAILY_QUOTA_LIMIT.toLocaleString()}</span>
          </div>
        </div>
        <div style="margin-top:14px;background:linear-gradient(135deg,#fef3c7 0%,#fde68a 100%);border:1px solid #fcd34d;border-radius:12px;padding:12px 14px;font-size:12.5px;color:#78350f;line-height:1.55;">
          <b>💡 Tham khảo chi phí:</b> <code style="background:#fff7ed;padding:1px 6px;border-radius:4px;">search.list</code> ~100 units · <code style="background:#fff7ed;padding:1px 6px;border-radius:4px;">videos.list</code>/<code style="background:#fff7ed;padding:1px 6px;border-radius:4px;">channels.list</code>/<code style="background:#fff7ed;padding:1px 6px;border-radius:4px;">playlistItems.list</code> ~1 unit. Mặc định mỗi project Google Cloud có ${YOUTUBE_DAILY_QUOTA_LIMIT.toLocaleString()} units/ngày.
        </div>
      </div>
      <div style="display:flex;justify-content:space-between;gap:8px;margin-top:14px;align-items:center;">
        <div id="lastUpdate" style="font-size:11px;color:#94a3b8;font-weight:600;">Đang tải...</div>
        <div style="display:flex;gap:8px;">
          <button id="refreshBtn" onclick="loadSnapshot(true)" style="padding:9px 16px;background:#fff;border:1px solid #d1d5db;border-radius:8px;cursor:pointer;font-weight:700;color:#0f172a;">🔄 Làm mới</button>
          <button onclick="google.script.host.close()" style="padding:9px 18px;background:linear-gradient(135deg,#7c3aed,#ec4899);color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:700;box-shadow:0 4px 12px rgba(124,58,237,.3);">Đóng</button>
        </div>
      </div>
      <style>
        .qChip{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:11px 14px;display:flex;justify-content:space-between;align-items:center;gap:10px;box-shadow:0 4px 10px -6px rgba(15,23,42,.10);transition:all .2s ease}
        .qChip:hover{transform:translateY(-1px);box-shadow:0 8px 18px -6px rgba(15,23,42,.16)}
        .qChip .lbl{font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;font-weight:700}
        .qChip .val{font-size:18px;font-weight:800;color:#0f172a;letter-spacing:-0.01em}
        .qChip .ico{font-size:20px;flex-shrink:0}
        .qChip.danger .val{color:#dc2626}
        .qChip.warn .val{color:#d97706}
        .qChip.ok .val{color:#059669}
      </style>
      <script>
        function fmt(n){ return Number(n).toLocaleString('vi-VN'); }
        function chip(id, ico, lbl, val, tone){ var el=document.getElementById(id); el.className='qChip '+(tone||''); el.innerHTML='<span class="ico">'+ico+'</span><div style="flex:1;"><div class="lbl">'+lbl+'</div><div class="val">'+val+'</div></div>'; }
        function loadSnapshot(manual){
          if (manual && window.dlgProgress) window.dlgProgress.start(1, 'Đang làm mới snapshot');
          google.script.run.withSuccessHandler(function(s){
            var ringFg = document.getElementById('ringFg');
            var pctClamp = Math.min(100, s.percent);
            ringFg.style.strokeDashoffset = (326.7 - (pctClamp/100)*326.7).toFixed(2);
            document.getElementById('ringPct').textContent = s.percent.toFixed(1)+'%';
            document.getElementById('barFill').style.width = Math.min(100, s.percent)+'%';
            document.getElementById('barPct').textContent = s.percent.toFixed(1)+'%';
            var tone = s.level==='danger'?'danger':(s.level==='warn'?'warn':'ok');
            chip('cUsed','📤','Đã dùng', fmt(s.used)+' units', tone);
            chip('cRemain','💎','Còn lại', fmt(s.remaining)+' units', s.remaining<=0?'danger':'ok');
            chip('cReq','📡','Số request', fmt(s.requests),'');
            chip('cAvg','⚖️','Avg cost/request', s.avgCostPerReq+' units','');
            chip('cReset','⏱️','Reset sau', s.resetIn,'ok');
            document.getElementById('quotaTip').innerHTML = (s.level==='danger'?'🚨 ':(s.level==='warn'?'⚠️ ':'✅ '))+s.tip;
            document.getElementById('lastUpdate').textContent = 'Cập nhật lúc '+s.nowPT+' (PT)';
            if (manual && window.dlgProgress) window.dlgProgress.complete('Đã làm mới');
          }).withFailureHandler(function(err){
            document.getElementById('lastUpdate').textContent = 'Lỗi: '+(err.message||err);
            if (manual && window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
          }).getQuotaSnapshot();
        }
        loadSnapshot(false);
        setInterval(function(){ loadSnapshot(false); }, 10000);
        document.addEventListener('keydown', function(e){ if(e.key==='Escape') google.script.host.close(); });
      </script>
    </div>
  `).setWidth(720).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, "📊 Quota YouTube API hôm nay");
}

function getSystemHealthData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apiSheet = ss.getSheetByName(SHEET_API_KEY);
  function readCell(addr){ try { return apiSheet ? apiSheet.getRange(addr).getValue().toString().trim() : ''; } catch(e){ return ''; } }
  const props = PropertiesService.getDocumentProperties();
  const checks = [];
  // Sheets
  const sheets = [SHEET_VIDEO, SHEET_CHANNEL, SHEET_API_KEY, SHEET_ANALYTICS, SHEET_VIDEO_ANALYTICS, AI_MEMORY_SHEET];
  sheets.forEach(function(name){
    const sh = ss.getSheetByName(name);
    const rows = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
    checks.push({
      group: 'Sheet',
      label: name,
      ok: !!sh,
      detail: sh ? (rows + ' dòng dữ liệu') : 'Chưa tồn tại — sẽ tự tạo khi cần',
      severity: sh ? 'ok' : (name === SHEET_API_KEY ? 'danger' : 'warn')
    });
  });
  // API keys
  const b2 = readCell('B2'), b3 = readCell('B3'), b4 = readCell('B4'), b5 = readCell('B5'), b6 = readCell('B6'), b7 = readCell('B7');
  checks.push({ group: 'API Key', label: 'YouTube Data API (B2)', ok: !!b2, detail: b2 ? ('Đã cấu hình • ' + b2.length + ' ký tự') : 'BẮT BUỘC — chưa điền', severity: b2 ? 'ok' : 'danger' });
  checks.push({ group: 'API Key', label: 'OAuth Client ID (B3)', ok: !!b3, detail: b3 ? ('Đã cấu hình • ' + b3.length + ' ký tự') : 'Trống — Analytics sẽ không hoạt động', severity: b3 ? 'ok' : 'warn' });
  checks.push({ group: 'API Key', label: 'OAuth Client Secret (B4)', ok: !!b4, detail: b4 ? ('Đã cấu hình • ' + b4.length + ' ký tự') : 'Trống — Analytics sẽ không hoạt động', severity: b4 ? 'ok' : 'warn' });
  checks.push({ group: 'API Key', label: 'Supadata Key (B5)', ok: !!b5, detail: b5 ? ('Đã cấu hình • ' + b5.length + ' ký tự') : 'Trống — sẽ fallback sang nguồn khác', severity: b5 ? 'ok' : 'warn' });
  checks.push({ group: 'API Key', label: 'YT-Transcript.io Token (B6)', ok: !!b6, detail: b6 ? ('Đã cấu hình • ' + b6.length + ' ký tự') : 'Trống — chỉ là phương án dự phòng', severity: b6 ? 'ok' : 'warn' });
  checks.push({ group: 'API Key', label: '9router Key/Token (B7)', ok: !!b7, detail: b7 ? ('Đã cấu hình • ' + b7.length + ' ký tự') : 'Trống — AI phân tích Sheet sẽ không gọi được', severity: b7 ? 'ok' : 'warn' });
  // Subtitle source coverage
  const subOk = !!(b5 || b6);
  checks.push({ group: 'Subtitle', label: 'Có ít nhất 1 nguồn ổn định (B5 hoặc B6)', ok: subOk, detail: subOk ? 'OK' : 'Sẽ fallback timedtext/scrape — dễ lỗi 429', severity: subOk ? 'ok' : 'warn' });
  // OAuth refresh token
  const refresh = props.getProperty('YT_REFRESH_TOKEN');
  checks.push({ group: 'OAuth', label: 'Refresh token Analytics', ok: !!refresh, detail: refresh ? 'Đã cấp quyền — Analytics sẵn sàng' : 'Chưa kết nối — chạy menu "Cài đặt Access Token"', severity: refresh ? 'ok' : 'warn' });
  // Quota
  const usage = getYouTubeQuotaUsage_();
  const pct = usage.limit ? (usage.used / usage.limit) * 100 : 0;
  checks.push({ group: 'Quota', label: 'YouTube Data quota hôm nay', ok: pct < 80, detail: usage.used + '/' + usage.limit + ' units (' + pct.toFixed(1) + '%) • ' + usage.requests + ' request', severity: pct >= 80 ? 'danger' : (pct >= 50 ? 'warn' : 'ok') });
  // 9router base URL
  const routerBase = props.getProperty('AI_9ROUTER_BASE_URL') || 'http://localhost:20128/v1';
  checks.push({ group: 'AI', label: '9router base URL', ok: true, detail: routerBase, severity: routerBase.indexOf('localhost') >= 0 ? 'warn' : 'ok' });
  // Summary
  const total = checks.length;
  const okCount = checks.filter(function(c){ return c.severity === 'ok'; }).length;
  const warnCount = checks.filter(function(c){ return c.severity === 'warn'; }).length;
  const dangerCount = checks.filter(function(c){ return c.severity === 'danger'; }).length;
  const score = Math.round((okCount + warnCount * 0.5) / total * 100);
  return { checks: checks, total: total, okCount: okCount, warnCount: warnCount, dangerCount: dangerCount, score: score, ts: new Date().toLocaleString('vi-VN') };
}

function showSystemHealthCheck() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:18px;background:#f8fafc;color:#202124;line-height:1.5;">
      ${buildDialogHeader_("🩺 Kiểm tra sức khỏe hệ thống", "Chẩn đoán toàn bộ API key, sheet, OAuth, quota và nguồn subtitle.", "#0f9d58")}
      <div class="dlgFadeBody">
        <div style="display:grid;grid-template-columns:auto 1fr;gap:16px;align-items:center;background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:18px;box-shadow:0 6px 18px -10px rgba(15,23,42,.16);margin-bottom:14px;">
          <div style="position:relative;width:120px;height:120px;">
            <svg viewBox="0 0 120 120" style="width:100%;height:100%;transform:rotate(-90deg);">
              <circle cx="60" cy="60" r="52" fill="none" stroke="#eef2f7" stroke-width="12"/>
              <circle id="hScore" cx="60" cy="60" r="52" fill="none" stroke="url(#hGrad)" stroke-width="12" stroke-linecap="round" stroke-dasharray="326.7" stroke-dashoffset="326.7" style="transition:stroke-dashoffset 1.2s cubic-bezier(.2,.7,.2,1);"/>
              <defs><linearGradient id="hGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#10b981"/><stop offset="100%" stop-color="#0ea5e9"/></linearGradient></defs>
            </svg>
            <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;">
              <div id="hScoreText" style="font-size:26px;font-weight:900;color:#0f172a;letter-spacing:-0.02em;">--</div>
              <div style="font-size:10px;font-weight:700;color:#64748b;letter-spacing:0.08em;text-transform:uppercase;">Health</div>
            </div>
          </div>
          <div>
            <div style="font-size:18px;font-weight:800;color:#0f172a;margin-bottom:6px;" id="hSummary">Đang kiểm tra...</div>
            <div style="display:flex;gap:8px;flex-wrap:wrap;" id="hBadges"></div>
            <div style="font-size:11.5px;color:#64748b;margin-top:8px;" id="hTs"></div>
          </div>
        </div>
        <div id="hList" style="display:flex;flex-direction:column;gap:8px;max-height:380px;overflow:auto;padding-right:6px;"></div>
      </div>
      <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px;">
        <button id="refreshBtn" onclick="loadHealth(true)" style="padding:10px 18px;background:#fff;border:1px solid #d1d5db;border-radius:8px;cursor:pointer;font-weight:700;color:#0f172a;">🔄 Kiểm tra lại</button>
        <button onclick="google.script.run.showSystemBrandInfo();google.script.host.close();" style="padding:10px 18px;background:#eef2ff;border:1px solid #c7d2fe;color:#3730a3;border-radius:8px;cursor:pointer;font-weight:700;">📘 Mở hướng dẫn</button>
        <button onclick="google.script.host.close()" style="padding:10px 18px;background:linear-gradient(135deg,#0f9d58,#059669);color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:700;box-shadow:0 4px 12px rgba(15,157,88,.3);">Đóng</button>
      </div>
      <style>
        .hRow{display:grid;grid-template-columns:auto 1fr auto;gap:12px;align-items:center;background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:11px 14px;box-shadow:0 2px 6px -3px rgba(15,23,42,.1);transition:all .2s ease;animation:dlgSlideUp .35s cubic-bezier(.2,.7,.2,1) both}
        .hRow:hover{transform:translateX(3px);box-shadow:0 6px 14px -6px rgba(15,23,42,.16)}
        .hRow .ico{width:32px;height:32px;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
        .hRow.ok .ico{background:linear-gradient(135deg,#d1fae5,#a7f3d0);color:#065f46}
        .hRow.warn .ico{background:linear-gradient(135deg,#fef3c7,#fde68a);color:#92400e}
        .hRow.danger .ico{background:linear-gradient(135deg,#fee2e2,#fecaca);color:#991b1b}
        .hRow .grp{font-size:10px;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;color:#94a3b8}
        .hRow .lbl{font-size:13.5px;font-weight:700;color:#0f172a;margin-top:1px}
        .hRow .det{font-size:12px;color:#64748b;margin-top:2px}
        .hRow .badge{font-size:10.5px;font-weight:800;padding:3px 9px;border-radius:999px;letter-spacing:0.04em;text-transform:uppercase}
        .hRow.ok .badge{background:#d1fae5;color:#065f46}
        .hRow.warn .badge{background:#fef3c7;color:#92400e}
        .hRow.danger .badge{background:#fee2e2;color:#991b1b}
        .hBadge{font-size:11px;font-weight:800;padding:3px 10px;border-radius:999px;letter-spacing:0.04em;text-transform:uppercase}
      </style>
      <script>
        function loadHealth(manual){
          if (manual && window.dlgProgress) window.dlgProgress.start(2, 'Đang chẩn đoán hệ thống');
          google.script.run.withSuccessHandler(function(d){
            document.getElementById('hScore').style.strokeDashoffset = (326.7 - (d.score/100)*326.7).toFixed(2);
            document.getElementById('hScoreText').textContent = d.score;
            var tone = d.dangerCount>0?'KÉM':(d.warnCount>0?'KHÁ':'TỐT');
            var summaryColor = d.dangerCount>0?'#dc2626':(d.warnCount>0?'#d97706':'#059669');
            document.getElementById('hSummary').innerHTML = 'Tình trạng: <span style="color:'+summaryColor+';">'+tone+'</span>';
            var badges = '';
            badges += '<span class="hBadge" style="background:#d1fae5;color:#065f46;">✅ '+d.okCount+' OK</span>';
            badges += '<span class="hBadge" style="background:#fef3c7;color:#92400e;">⚠️ '+d.warnCount+' Cảnh báo</span>';
            badges += '<span class="hBadge" style="background:#fee2e2;color:#991b1b;">🚨 '+d.dangerCount+' Nghiêm trọng</span>';
            document.getElementById('hBadges').innerHTML = badges;
            document.getElementById('hTs').textContent = 'Lần kiểm tra cuối: '+d.ts;
            var html = '';
            d.checks.forEach(function(c, i){
              var icon = c.severity==='ok'?'✓':(c.severity==='warn'?'!':'×');
              var badge = c.severity==='ok'?'OK':(c.severity==='warn'?'CẢNH BÁO':'NGHIÊM TRỌNG');
              html += '<div class="hRow '+c.severity+'" style="animation-delay:'+(i*30)+'ms;"><div class="ico">'+icon+'</div><div><div class="grp">'+c.group+'</div><div class="lbl">'+c.label+'</div><div class="det">'+c.detail+'</div></div><span class="badge">'+badge+'</span></div>';
            });
            document.getElementById('hList').innerHTML = html;
            if (manual && window.dlgProgress) window.dlgProgress.complete('Hoàn tất • Score: '+d.score);
          }).withFailureHandler(function(err){
            document.getElementById('hSummary').innerHTML = '<span style="color:#dc2626;">Lỗi: '+(err.message||err)+'</span>';
            if (manual && window.dlgProgress) window.dlgProgress.fail('Lỗi: '+(err.message||err));
          }).getSystemHealthData();
        }
        loadHealth(false);
        if (window.dlgProgress) window.dlgProgress.start(2, 'Đang chẩn đoán hệ thống');
        document.addEventListener('keydown', function(e){ if(e.key==='Escape') google.script.host.close(); });
      </script>
    </div>
  `).setWidth(780).setHeight(740);
  SpreadsheetApp.getUi().showModalDialog(html, "🩺 Kiểm tra sức khỏe hệ thống");
}

