(function(global){
  const SPREADSHEET_ID = '1EvCTYTus1CGRuokNlAqF8ucqeKmwRLwUpKrZiq0bPGU';
  const SPREADSHEET_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;
  const LOCAL_API_KEY = 'hoahoc10_apps_script_url';
  const GH_TOKEN_KEY = 'hoahoc10_gh_token';
  const REPO_OWNER = 'Dthainek';
  const REPO_NAME = 'hoahoc10';
  const CONFIG_PATH = 'sheets_config.json';

  const APPS_SCRIPT_CODE = `function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = data.game === 'LatManhGhep' ? 'LatManhGhep' : 'ToHinhDoanChu';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonOut({ ok: false, error: 'Sheet not found' });

    if (data.action === 'reset') {
      var lastRow = sheet.getLastRow();
      var deleted = 0;
      var targetSet = data.setTitle ? String(data.setTitle) : '';
      if (lastRow > 1) {
        if (!targetSet) {
          deleted = lastRow - 1;
          sheet.deleteRows(2, lastRow - 1);
        } else {
          var vals = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
          for (var r = vals.length - 1; r >= 0; r--) {
            if (String(vals[r][3]) === targetSet) {
              sheet.deleteRow(r + 2);
              deleted++;
            }
          }
        }
      }
      return jsonOut({ ok: true, reset: true, deleted: deleted });
    }

    var stt = sheet.getLastRow();
    sheet.appendRow([
      stt,
      data.name || '',
      data.className || '',
      data.setTitle || '',
      data.startTime || '',
      data.endTime || '',
      Number(data.totalTime) || 0,
      data.correctCount || '',
      Number(data.score) || 0
    ]);
    return jsonOut({ ok: true, stt: stt });
  } catch (err) {
    return jsonOut({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    var game = (e && e.parameter && e.parameter.game) || 'ToHinhDoanChu';
    var setTitle = (e && e.parameter && e.parameter.setTitle) || '';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(game);
    if (!sheet) return jsonOut({ ok: false, rows: [] });

    var values = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!r[1]) continue;
      if (setTitle && String(r[3]) !== setTitle) continue;
      list.push({
        name: r[1],
        className: r[2],
        setTitle: r[3],
        startTime: r[4],
        endTime: r[5],
        totalTime: Number(r[6]) || 0,
        correctCount: String(r[7]),
        score: Number(r[8]) || 0
      });
    }
    list.sort(function(a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.totalTime - b.totalTime;
    });
    return jsonOut({ ok: true, rows: list.slice(0, 30) });
  } catch (err) {
    return jsonOut({ ok: false, rows: [], error: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}`;

  let appsScriptUrl = '';
  let modalsInjected = false;

  async function initConfig(){
    const params = new URLSearchParams(window.location.search);
    const paramUrl = params.get('api');
    if(paramUrl && paramUrl.startsWith('https://script.google.com/')){
      localStorage.setItem(LOCAL_API_KEY, paramUrl);
    }

    try {
      const res = await fetch('../sheets_config.json?t=' + Date.now());
      if(res.ok){
        const cfg = await res.json();
        if(cfg.appsScriptUrl) appsScriptUrl = cfg.appsScriptUrl.trim();
      }
    } catch(e){}

    if(!appsScriptUrl){
      appsScriptUrl = (localStorage.getItem(LOCAL_API_KEY) || '').trim();
    }
  }

  function pad(n){ return String(n).padStart(2, '0'); }
  function formatVNTime(d){
    const dt = d instanceof Date ? d : new Date(d);
    return `${pad(dt.getDate())}/${pad(dt.getMonth()+1)}/${dt.getFullYear()} ${pad(dt.getHours())}:${pad(dt.getMinutes())}:${pad(dt.getSeconds())}`;
  }

  function calcQuestionScore(isRight, timeRemainingSec){
    if(!isRight) return 0;
    const bonus = Math.max(0, Math.min(30, Math.round(timeRemainingSec)));
    return 100 + bonus;
  }

  async function submitScore(payload){
    const url = appsScriptUrl || (localStorage.getItem(LOCAL_API_KEY) || '').trim();
    if(!url) return { ok: false, reason: 'not_configured' };
    try {
      const res = await fetch(url, {
        method: 'POST',
        body: JSON.stringify(payload)
      });
      if(res.ok){
        const data = await res.json();
        return { ok: !!data.ok, data };
      }
      return { ok: false, reason: 'http_' + res.status };
    } catch(e){
      // Fallback no-cors trong trường hợp trình duyệt chặn đọc response redirect từ Apps Script
      try {
        await fetch(url, {
          method: 'POST',
          mode: 'no-cors',
          body: JSON.stringify(payload)
        });
        return { ok: true, mode: 'no-cors' };
      } catch(err2){
        return { ok: false, reason: err2.message };
      }
    }
  }

  async function fetchLeaderboard(game, setTitle){
    const url = appsScriptUrl || (localStorage.getItem(LOCAL_API_KEY) || '').trim();
    if(!url) return { ok: false, reason: 'not_configured', rows: [] };
    try {
      const q = new URLSearchParams({ game });
      if(setTitle) q.set('setTitle', setTitle);
      const res = await fetch(url + '?' + q.toString());
      if(!res.ok) throw new Error('HTTP ' + res.status);
      return await res.json();
    } catch(e){
      return { ok: false, reason: e.message, rows: [] };
    }
  }

  function ensureModals(){
    if(modalsInjected) return;
    modalsInjected = true;

    const style = document.createElement('style');
    style.textContent = `
      .sh-modal{position:fixed;inset:0;background:rgba(15,23,42,0.72);backdrop-filter:blur(4px);display:none;align-items:center;justify-content:center;z-index:220}
      .sh-modal.show{display:flex}
      .sh-box{background:#ffffff;border-radius:18px;width:min(760px,95vw);max-height:90vh;overflow:auto;padding:24px;color:#1e293b;box-shadow:0 24px 48px rgba(0,0,0,0.28)}
      .sh-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid #e2e8f0}
      .sh-head h3{margin:0;font-size:19px;font-weight:800;color:#0f172a}
      .sh-close{border:1px solid #cbd5e1;background:#f8fafc;border-radius:8px;padding:6px 12px;font-weight:700;cursor:pointer}
      .sh-table{width:100%;border-collapse:collapse;font-size:14px;margin-top:10px}
      .sh-table th,.sh-table td{padding:9px 10px;border-bottom:1px solid #e2e8f0;text-align:left}
      .sh-table th{background:#f1f5f9;font-weight:800;color:#334155;font-size:13px}
      .sh-table tr:hover td{background:#f8fafc}
      .sh-code{width:100%;height:160px;font-family:monospace;font-size:12px;padding:10px;border:1px solid #cbd5e1;border-radius:8px;background:#0f172a;color:#e2e8f0;resize:vertical}
      .sh-inp{width:100%;padding:9px 12px;border:1.5px solid #cbd5e1;border-radius:8px;font:inherit;font-size:14px;margin-top:4px}
      .sh-btn{border:1px solid #cbd5e1;background:#f8fafc;color:#1e293b;border-radius:8px;padding:8px 14px;font:inherit;font-size:13px;font-weight:700;cursor:pointer}
      .sh-btn.pri{background:#2563eb;color:#fff;border-color:#1d4ed8}
      .sh-btn.acc{background:#16a34a;color:#fff;border-color:#15803d}
      .sh-step{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:12px 14px;margin-bottom:12px;font-size:13.5px;line-height:1.5}
    `;
    document.head.appendChild(style);

    const wrap = document.createElement('div');
    wrap.innerHTML = `
      <!-- Modal Bảng Xếp Hạng Online -->
      <div class="sh-modal" id="shLeaderboardModal">
        <div class="sh-box">
          <div class="sh-head">
            <h3>🏆 Bảng Xếp Hạng Online</h3>
            <button class="sh-close" id="shLbClose">✖ Đóng</button>
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:8px">
            <div style="display:flex;gap:8px;align-items:center">
              <label style="font-size:13px;font-weight:700">Hiển thị:</label>
              <select id="shLbFilter" style="padding:5px 8px;border-radius:6px;border:1px solid #cbd5e1;font:inherit;font-size:13px">
                <option value="all">Tất cả bộ câu hỏi</option>
                <option value="current">Chỉ bộ câu hỏi hiện tại</option>
              </select>
              <button class="sh-btn" id="shLbRefresh">🔄 Làm mới</button>
            </div>
            <a href="${SPREADSHEET_URL}" target="_blank" style="color:#16a34a;font-size:13px;font-weight:700;text-decoration:none">📊 Mở Google Sheet gốc ↗</a>
          </div>
          <div id="shLbContent" style="min-height:120px;overflow-x:auto">Đang tải...</div>
        </div>
      </div>

      <!-- Modal Cài Đặt Kết Nối Google Sheets -->
      <div class="sh-modal" id="shSetupModal">
        <div class="sh-box">
          <div class="sh-head">
            <h3>📊 Kết nối Google Sheets (Lưu điểm Online)</h3>
            <button class="sh-close" id="shSetupClose">✖ Đóng</button>
          </div>
          <div class="sh-step">
            <b>Bước 1:</b> Mở file Google Sheet đã tạo sẵn trên Drive của bạn:
            <div style="margin-top:6px">
              <a href="${SPREADSHEET_URL}" target="_blank" class="sh-btn acc" style="display:inline-block;text-decoration:none">📊 Mở Google Sheet: Bảng Điểm Online ↗</a>
            </div>
          </div>
          <div class="sh-step">
            <b>Bước 2:</b> Trong Google Sheet, vào menu <b>Tiện ích mở rộng (Extensions) → Apps Script</b>, xóa hết chữ cũ và dán đoạn mã dưới đây rồi bấm Lưu:
            <div style="margin:8px 0">
              <textarea class="sh-code" id="shCodeArea" readonly></textarea>
            </div>
            <button class="sh-btn pri" id="shCopyCode">📋 Sao chép mã Apps Script</button>
          </div>
          <div class="sh-step">
            <b>Bước 3:</b> Bấm <b>Triển khai (Deploy) → Triển khai mới (New deployment)</b> → Chọn loại <b>Ứng dụng web (Web app)</b> → Mục <i>Ai có quyền truy cập (Who has access)</i> chọn <b>Bất kỳ ai (Anyone)</b> → Bấm <b>Triển khai</b> rồi dán đường link Web App (<code>https://script.google.com/macros/s/.../exec</code>) vào ô dưới đây:
            <input class="sh-inp" id="shApiInput" placeholder="https://script.google.com/macros/s/.../exec">
            <input class="sh-inp" type="password" id="shGhTokenInput" placeholder="GitHub Token (nếu muốn lưu vĩnh viễn lên web cho cả lớp dùng)" style="margin-top:8px">
            <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:10px">
              <button class="sh-btn pri" id="shSaveLocalBtn">💾 Lưu URL trên máy này</button>
              <button class="sh-btn acc" id="shPublishGhBtn">🚀 Xuất bản URL lên GitHub (Cả lớp cùng lưu)</button>
            </div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(wrap);

    document.getElementById('shCodeArea').value = APPS_SCRIPT_CODE;
    document.getElementById('shLbClose').onclick = () => document.getElementById('shLeaderboardModal').classList.remove('show');
    document.getElementById('shSetupClose').onclick = () => document.getElementById('shSetupModal').classList.remove('show');
    document.getElementById('shCopyCode').onclick = () => {
      navigator.clipboard.writeText(APPS_SCRIPT_CODE);
      document.getElementById('shCopyCode').textContent = '✅ Đã sao chép!';
      setTimeout(() => document.getElementById('shCopyCode').textContent = '📋 Sao chép mã Apps Script', 2500);
    };

    document.getElementById('shSaveLocalBtn').onclick = () => {
      const val = document.getElementById('shApiInput').value.trim();
      localStorage.setItem(LOCAL_API_KEY, val);
      appsScriptUrl = val;
      alert(val ? 'Đã lưu URL Google Apps Script vào máy này!' : 'Đã xóa URL.');
      document.getElementById('shSetupModal').classList.remove('show');
    };

    document.getElementById('shPublishGhBtn').onclick = async () => {
      const val = document.getElementById('shApiInput').value.trim();
      const token = document.getElementById('shGhTokenInput').value.trim();
      if(!val) return alert('Vui lòng dán URL Web App của Google Apps Script!');
      if(!token) return alert('Vui lòng nhập GitHub Token để cập nhật lên GitHub!');
      localStorage.setItem(LOCAL_API_KEY, val);
      localStorage.setItem(GH_TOKEN_KEY, token);
      appsScriptUrl = val;

      const btn = document.getElementById('shPublishGhBtn');
      btn.disabled = true; btn.textContent = '⏳ Đang xuất bản...';
      try {
        let sha = null;
        const g = await fetch(`https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${CONFIG_PATH}`, {
          headers: {'Authorization': `Bearer ${token}`, 'Accept': 'application/vnd.github.v3+json'}
        });
        if(g.ok) sha = (await g.json()).sha;

        const bodyObj = { spreadsheetId: SPREADSHEET_ID, spreadsheetUrl: SPREADSHEET_URL, appsScriptUrl: val };
        const bytes = new TextEncoder().encode(JSON.stringify(bodyObj, null, 2));
        let bin = ''; bytes.forEach(b => bin += String.fromCharCode(b));

        const p = await fetch(`https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${CONFIG_PATH}`, {
          method: 'PUT',
          headers: {'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'Accept': 'application/vnd.github.v3+json'},
          body: JSON.stringify({ message: 'Cap nhat URL Google Apps Script luu diem online', content: btoa(bin), sha: sha || undefined })
        });
        if(!p.ok) throw new Error((await p.json()).message || 'Lỗi GitHub');
        alert('🎉 Đã lưu URL kết nối Google Sheets lên GitHub! Sau ~1 phút toàn bộ học sinh sẽ tự động lưu điểm vào Sheet.');
        document.getElementById('shSetupModal').classList.remove('show');
      } catch(e){
        alert('Không thể xuất bản lên GitHub: ' + e.message);
      } finally {
        btn.disabled = false; btn.textContent = '🚀 Xuất bản URL lên GitHub (Cả lớp cùng lưu)';
      }
    };
  }

  let currentLbGame = 'ToHinhDoanChu';
  let currentLbSet = '';

  async function renderLeaderboardInto(el, gameName, setTitle, highlightName){
    ensureModals();
    el.innerHTML = '<p style="text-align:center;color:#64748b">⏳ Đang tải bảng xếp hạng từ Google Sheets...</p>';
    const res = await fetchLeaderboard(gameName, setTitle || '');
    if(!res.ok){
      if(res.reason === 'not_configured'){
        el.innerHTML = `<div style="text-align:center;padding:16px;color:#b45309;background:#fffbeb;border-radius:10px;font-size:13px">
          ⚠️ Chưa cấu hình link Google Apps Script.<br><br>
          <button class="sh-btn pri" onclick="SheetsOnline.openSetup();">⚙️ Kết nối Google Sheets</button>
        </div>`;
      } else {
        el.innerHTML = `<p style="color:#dc2626;text-align:center;font-size:13px">Không tải được bảng xếp hạng: ${res.reason}</p>`;
      }
      return;
    }
    if(!res.rows || res.rows.length === 0){
      el.innerHTML = '<p style="text-align:center;color:#64748b;padding:16px;font-size:13px">Chưa có lượt chơi nào được lưu trên Google Sheets.</p>';
      return;
    }
    const esc = s => String(s||'').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
    let html = `<table class="sh-table"><thead><tr><th>Hạng</th><th>Họ và tên</th><th>Lớp</th><th>Bộ đề</th><th>Bắt đầu</th><th>Thời gian</th><th>Đúng</th><th>Điểm</th></tr></thead><tbody>`;
    res.rows.forEach((r, idx) => {
      const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : (idx + 1);
      const isMe = highlightName && String(r.name).trim() === String(highlightName).trim();
      html += `<tr style="${isMe ? 'background:#fef3c7;font-weight:700' : ''}">
        <td style="font-weight:800;text-align:center">${medal}</td>
        <td style="font-weight:700">${esc(r.name)}</td>
        <td>${esc(r.className)}</td>
        <td style="font-size:12px;color:#475569">${esc(r.setTitle)}</td>
        <td style="font-size:12px;color:#64748b">${esc(r.startTime)}</td>
        <td>${r.totalTime}s</td>
        <td><b>${esc(r.correctCount)}</b></td>
        <td style="font-weight:800;color:#2563eb">${r.score}</td>
      </tr>`;
    });
    html += '</tbody></table>';
    el.innerHTML = html;
  }

  async function loadLeaderboardUI(){
    const el = document.getElementById('shLbContent');
    const mode = document.getElementById('shLbFilter').value;
    await renderLeaderboardInto(el, currentLbGame, mode === 'current' ? currentLbSet : '', '');
  }

  function openLeaderboard(gameName, setTitle){
    ensureModals();
    currentLbGame = gameName;
    currentLbSet = setTitle || '';
    document.getElementById('shLbFilter').onchange = loadLeaderboardUI;
    document.getElementById('shLbRefresh').onclick = loadLeaderboardUI;
    document.getElementById('shLeaderboardModal').classList.add('show');
    loadLeaderboardUI();
  }

  function openSetup(){
    ensureModals();
    document.getElementById('shApiInput').value = appsScriptUrl || localStorage.getItem(LOCAL_API_KEY) || '';
    document.getElementById('shGhTokenInput').value = localStorage.getItem(GH_TOKEN_KEY) || '';
    document.getElementById('shSetupModal').classList.add('show');
  }

  async function resetLeaderboard(gameName, setTitle){
    const url = appsScriptUrl || (localStorage.getItem(LOCAL_API_KEY) || '').trim();
    if(!url) return { ok: false, reason: 'not_configured' };
    try {
      const res = await fetch(url, {
        method: 'POST',
        body: JSON.stringify({ action: 'reset', game: gameName || 'ToHinhDoanChu', setTitle: setTitle || '' })
      });
      if(res.ok){
        const data = await res.json();
        if(data && data.ok && data.reset){
          return { ok: true, deleted: data.deleted || 0 };
        }
        return { ok: false, reason: 'needs_script_update' };
      }
      return { ok: false, reason: 'http_' + res.status };
    } catch(e){
      return { ok: false, reason: e.message || 'network_error' };
    }
  }

  global.SheetsOnline = {
    SPREADSHEET_URL,
    initConfig,
    formatVNTime,
    calcQuestionScore,
    submitScore,
    resetLeaderboard,
    renderLeaderboardInto,
    openLeaderboard,
    openSetup,
    isConfigured: () => !!(appsScriptUrl || localStorage.getItem(LOCAL_API_KEY))
  };
})(window);
