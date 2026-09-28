'use strict';

/* ============================================================
   設定
   ============================================================ */
const EXPORT_PREFIX = '売場スキャン２'; // Android版と同じ出力ファイル名にそろえる
const FORMATS = ['ean_13', 'ean_8', 'upc_a', 'upc_e'];
const PROCESS_COOLDOWN_MS = 2000;
const CENTER_ZONE_RATIO = 0.30;
const SCAN_INTERVAL_MS = 80;
const RESULT_SHOW_MS = 2500;

// 認識枠の色
const COLOR_BLUE = '#3BA7FF';
const COLOR_GREEN = '#2FE08A';
const COLOR_RED = '#FF4D4F';
const COLOR_ORANGE = '#FFB020';

const KEY_SESSION = 'uriba_session';
const KEY_DEVICE_ID = 'uriba_device_id';
const KEY_MASTER = 'uriba_master';

const $ = (id) => document.getElementById(id);

/* ============================================================
   保存（ブラウザ内：localStorage）
   ============================================================ */
function loadJson(key) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null; // 破損データは無視
  }
}

function saveJson(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (e) {
    return false;
  }
}

/** 端末ID：見間違え防止のため I,1,O,0,Q を除外した英数字2桁 */
function getOrCreateDeviceId() {
  const ID_CHARS = 'ABCDEFGHJKLMNPRSTUVWXYZ23456789';
  let id = null;
  try { id = localStorage.getItem(KEY_DEVICE_ID); } catch (e) { /* 無視 */ }
  if (id) return id;
  id = '';
  for (let i = 0; i < 2; i++) id += ID_CHARS[Math.floor(Math.random() * ID_CHARS.length)];
  try { localStorage.setItem(KEY_DEVICE_ID, id); } catch (e) { /* 無視 */ }
  return id;
}

/** スキャン済みデータ（1件ごとに保存し、再読み込み後も再開できる） */
const session = {
  records: [],
  scanned: new Set(),
  load() {
    const arr = loadJson(KEY_SESSION);
    this.records = Array.isArray(arr) ? arr.filter((r) => r && r.jan) : [];
    this.scanned = new Set(this.records.map((r) => r.jan));
  },
  isScanned(jan) { return this.scanned.has(jan); },
  add(jan, name) {
    this.records.push({ jan, name, ts: Date.now() });
    this.scanned.add(jan);
    saveJson(KEY_SESSION, this.records);
  },
  clear() {
    this.records = [];
    this.scanned.clear();
    saveJson(KEY_SESSION, this.records);
  },
};

/* ============================================================
   マスター（転送.xlsx を端末内で読み込む。サーバーには送らない）
   保存形式: { items: {JAN: [商品名, 在売価]}, file, loadedAt }
   ============================================================ */
let master = null;

/** 転送.xlsx の全シートから D列=JAN, E列=品番名, F列=在売価 を読み取る（重複JANは先勝ち） */
function parseMaster(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const items = {};
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
    for (const r of rows) {
      if (!r || r[3] == null) continue;
      const jan = (typeof r[3] === 'number' ? String(Math.round(r[3])) : String(r[3])).trim();
      if (!/^\d{8,14}$/.test(jan)) continue; // 見出し行など
      if (items[jan]) continue;
      const name = r[4] == null ? '' : String(r[4]).trim();
      let price = null;
      if (typeof r[5] === 'number') price = Math.round(r[5]);
      else if (r[5] != null && /^\d+$/.test(String(r[5]).trim())) price = Number(String(r[5]).trim());
      items[jan] = [name, price];
    }
  }
  return items;
}

function lookup(jan) {
  const v = master && master.items[jan];
  return v ? { name: v[0], price: v[1] } : null;
}

function masterCount() {
  return master ? Object.keys(master.items).length : 0;
}

function updateMasterStatus() {
  const btn = $('btnMaster');
  if (!master) {
    btn.textContent = 'マスター未読込（タップして転送.xlsxを選択）';
    btn.classList.add('missing');
    return;
  }
  const d = new Date(master.loadedAt);
  btn.textContent = `マスター ${masterCount().toLocaleString('ja-JP')}件（${d.getMonth() + 1}/${d.getDate()} 読込）`;
  btn.classList.remove('missing');
}

function pickMasterFile() {
  $('masterFile').value = '';
  $('masterFile').click();
}

async function onMasterFileSelected(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  showToast('マスターを読み込み中…', 10000);
  try {
    const items = parseMaster(await file.arrayBuffer());
    const count = Object.keys(items).length;
    if (count === 0) {
      showDialog('JANコードが見つかりませんでした。\n転送.xlsx（D列にJAN、E列に品番名、F列に在売価）を選んでください。', [{ label: 'OK' }]);
      hideToast();
      return;
    }
    master = { items, file: file.name, loadedAt: Date.now() };
    if (!saveJson(KEY_MASTER, master)) {
      showDialog('マスターを端末に保存できませんでした。\n（プライベートブラウズでは保存できません）\n今回の起動中だけ使えます。', [{ label: 'OK' }]);
    }
    updateMasterStatus();
    refreshCover();
    showToast(`マスターを読み込みました（${count.toLocaleString('ja-JP')}件）`);
  } catch (err) {
    hideToast();
    showDialog(`マスターの読み込みに失敗しました。\n(${err.message})`, [{ label: 'OK' }]);
  }
}

/* ============================================================
   音・バイブ
   ============================================================ */
let audioCtx = null;

function unlockAudio() {
  try {
    // iPhoneのマナーモードでも鳴るようにする（Safari 16.4+）
    if (navigator.audioSession) navigator.audioSession.type = 'playback';
    if (!audioCtx) {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (AC) audioCtx = new AC();
    }
    if (audioCtx && audioCtx.state !== 'running') audioCtx.resume();
  } catch (e) { /* 音が出なくても業務は継続 */ }
}

function playTone(freq, ms, delayMs = 0) {
  if (!audioCtx) return;
  try {
    const t0 = audioCtx.currentTime + delayMs / 1000;
    const t1 = t0 + ms / 1000;
    const osc = audioCtx.createOscillator();
    const gain = audioCtx.createGain();
    osc.type = 'sine';
    osc.frequency.value = freq;
    // なだらかなフェードアウトでクリック音を防ぐ
    gain.gain.setValueAtTime(0.6, t0);
    gain.gain.linearRampToValueAtTime(0.0001, t1);
    osc.connect(gain).connect(audioCtx.destination);
    osc.start(t0);
    osc.stop(t1 + 0.02);
  } catch (e) { /* 無視 */ }
}

const feedback = {
  success() { playTone(1760, 90); vibrate([100]); },
  warn() { playTone(220, 160); vibrate([200, 100, 200]); },
  // 少し低めの2連音で「ポン・ポン」
  notFound() { playTone(440, 80); playTone(330, 80, 120); },
};

function vibrate(pattern) {
  try { if (navigator.vibrate) navigator.vibrate(pattern); } catch (e) { /* iPhoneは非対応 */ }
}

/* ============================================================
   カメラ
   ============================================================ */
const video = $('video');
const overlay = $('overlay');
let stream = null;
let track = null;
let detector = null;
let cameraStarted = false;
let cameraError = false;
let scanning = false;

let mode = 'individual'; // individual | bulk
let zoomedIn = false;
let torchOn = false;

async function createDetector() {
  // Android の Chrome は標準の BarcodeDetector を使う。iPhone等は ZXing(WebAssembly) を使う
  if ('BarcodeDetector' in window) {
    try {
      const supported = await window.BarcodeDetector.getSupportedFormats();
      if (FORMATS.every((f) => supported.includes(f))) {
        return new window.BarcodeDetector({ formats: FORMATS });
      }
    } catch (e) { /* ZXingへフォールバック */ }
  }
  const api = window.BarcodeDetectionAPI;
  api.prepareZXingModule({
    overrides: {
      locateFile: (path, prefix) =>
        path.endsWith('.wasm') ? new URL(`vendor/${path}`, location.href).href : prefix + path,
    },
  });
  return new api.BarcodeDetector({ formats: FORMATS });
}

async function startCamera() {
  cameraError = false;
  refreshCover();
  try {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      throw new Error('このブラウザはカメラに対応していません（HTTPSで開いてください）');
    }
    if (!detector) detector = await createDetector();
    stream = await navigator.mediaDevices.getUserMedia({
      audio: false,
      video: {
        facingMode: { ideal: 'environment' },
        width: { ideal: 1920 },
        height: { ideal: 1080 },
      },
    });
    track = stream.getVideoTracks()[0];
    video.srcObject = stream;
    await video.play();

    const caps = getCaps();
    if (caps.focusMode && caps.focusMode.includes('continuous')) {
      track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(() => {});
    }
    // 端末が再起動した場合に備えてズーム／ライトの状態を反映し直す
    torchOn = false;
    $('btnTorch').classList.remove('warn');
    applyZoom();

    cameraStarted = true;
    refreshCover();
    requestWakeLock();
    if (!scanning) {
      scanning = true;
      scanLoop();
    }
  } catch (e) {
    console.error(e);
    cameraError = true;
    refreshCover(e && e.name === 'NotAllowedError'
      ? 'カメラの使用が許可されていません。\nブラウザの設定でカメラを許可してから再試行してください。'
      : null);
  }
}

function getCaps() {
  try {
    return (track && track.getCapabilities) ? track.getCapabilities() : {};
  } catch (e) {
    return {};
  }
}

function applyZoom() {
  const caps = getCaps();
  const inner = $('cameraInner');
  if (caps.zoom) {
    // 端末のズーム機能が使える場合
    inner.classList.remove('digital-zoom');
    const target = zoomedIn ? Math.min(caps.zoom.max, 2) : Math.max(caps.zoom.min, 1);
    track.applyConstraints({ advanced: [{ zoom: target }] }).catch(() => {});
  } else {
    // 使えない端末（iPhone等）は画面上の拡大で代用
    inner.classList.toggle('digital-zoom', zoomedIn);
  }
}

async function scanLoop() {
  if (!scanning) return;
  if (cameraStarted && video.readyState >= 2 && !document.hidden && video.videoWidth > 0) {
    try {
      const barcodes = await detector.detect(video);
      processBarcodes(barcodes, video.videoWidth, video.videoHeight);
    } catch (e) { /* 1フレームの失敗は無視 */ }
  }
  setTimeout(scanLoop, SCAN_INTERVAL_MS);
}

/* ============================================================
   照合ロジック
   ============================================================ */
const lastProcessedTs = new Map();

function colorFor(jan) {
  if (!jan) return COLOR_BLUE;
  if (!lookup(jan)) return COLOR_RED;
  if (session.isScanned(jan)) return COLOR_ORANGE;
  // クールダウン中か判定（直前にスキャン成功したものは緑にする）
  const lastTs = lastProcessedTs.get(jan);
  if (lastTs != null && Date.now() - lastTs < PROCESS_COOLDOWN_MS) return COLOR_GREEN;
  return COLOR_BLUE;
}

function processBarcodes(barcodes, frameWidth, frameHeight) {
  drawOverlay(barcodes, frameWidth, frameHeight);
  if (barcodes.length === 0 || !$('dialog').hidden) return;

  if (mode === 'bulk') {
    barcodes.forEach((b) => b.rawValue && handleDetectedJan(b.rawValue));
    return;
  }

  // 個別モード：中心付近の1件のみ採用
  const cx = frameWidth / 2;
  const cy = frameHeight / 2;
  const marginX = frameWidth * CENTER_ZONE_RATIO;
  const marginY = frameHeight * CENTER_ZONE_RATIO;
  let best = null;
  let bestDist = Infinity;
  for (const b of barcodes) {
    const r = b.boundingBox;
    if (!r) continue;
    const bx = r.x + r.width / 2;
    const by = r.y + r.height / 2;
    if (Math.abs(bx - cx) < marginX && Math.abs(by - cy) < marginY) {
      const d = Math.hypot(bx - cx, by - cy);
      if (d < bestDist) {
        bestDist = d;
        best = b;
      }
    }
  }
  if (best && best.rawValue) handleDetectedJan(best.rawValue);
}

function handleDetectedJan(rawJan) {
  const jan = rawJan.trim();
  const now = Date.now();
  const cooldownAt = lastProcessedTs.get(jan);
  if (cooldownAt != null && now - cooldownAt < PROCESS_COOLDOWN_MS) return; // クールダウン中は完全無視
  lastProcessedTs.set(jan, now);

  const item = lookup(jan);
  if (!item) {
    // マスター未登録 -> 通知のみ行う（保存はしない）
    feedback.notFound();
    showResult(jan, null);
    return;
  }

  showResult(jan, item);

  if (session.isScanned(jan)) {
    // 二重登録
    feedback.warn();
    showWarn(`二重登録済み：${item.name}`);
    return;
  }

  // 新規登録
  session.add(jan, item.name);
  updateCount();
  feedback.success();
}

/* ============================================================
   画面表示更新
   ============================================================ */
function drawOverlay(barcodes, vw, vh) {
  const dpr = window.devicePixelRatio || 1;
  const cw = overlay.clientWidth;
  const ch = overlay.clientHeight;
  if (overlay.width !== Math.round(cw * dpr) || overlay.height !== Math.round(ch * dpr)) {
    overlay.width = Math.round(cw * dpr);
    overlay.height = Math.round(ch * dpr);
  }
  const ctx = overlay.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cw, ch);
  // video は object-fit: cover なので同じ換算をする
  const scale = Math.max(cw / vw, ch / vh);
  const ox = (cw - vw * scale) / 2;
  const oy = (ch - vh * scale) / 2;
  for (const b of barcodes) {
    const r = b.boundingBox;
    if (!r) continue;
    const color = colorFor((b.rawValue || '').trim());
    const x = ox + r.x * scale;
    const y = oy + r.y * scale;
    const w = r.width * scale;
    const h = r.height * scale;
    ctx.lineWidth = 3;
    ctx.strokeStyle = color;
    ctx.fillStyle = color + '33';
    ctx.beginPath();
    if (ctx.roundRect) ctx.roundRect(x, y, w, h, 6); else ctx.rect(x, y, w, h);
    ctx.fill();
    ctx.stroke();
  }
}

function updateCount() {
  $('countNum').textContent = String(session.records.length);
}

let resultTimer = null;
let warnTimer = null;
let toastTimer = null;

/** マスターにあれば「◯」と商品名・在売価（税込）、なければ「該当なし」を表示する */
function showResult(jan, item) {
  clearTimeout(resultTimer);
  const mark = $('resultMark');
  if (item) {
    mark.textContent = '◯';
    mark.classList.remove('none');
    $('resultName').textContent = `${jan}　${item.name}`;
    if (item.price != null) {
      $('resultPrice').textContent = `¥${item.price.toLocaleString('ja-JP')}`;
      $('resultPriceRow').hidden = false;
    } else {
      $('resultPriceRow').hidden = true;
    }
  } else {
    mark.textContent = '該当なし';
    mark.classList.add('none');
    $('resultName').textContent = jan;
    $('resultPriceRow').hidden = true;
  }
  $('resultPanel').classList.add('show');
  resultTimer = setTimeout(() => $('resultPanel').classList.remove('show'), RESULT_SHOW_MS);
}

function showWarn(text) {
  clearTimeout(warnTimer);
  $('warnText').textContent = text;
  $('warnText').classList.add('show');
  warnTimer = setTimeout(() => $('warnText').classList.remove('show'), RESULT_SHOW_MS);
}

function showToast(text, ms = 2200) {
  clearTimeout(toastTimer);
  $('toast').textContent = text;
  $('toast').hidden = false;
  toastTimer = setTimeout(hideToast, ms);
}

function hideToast() {
  clearTimeout(toastTimer);
  $('toast').hidden = true;
}

/**
 * ダイアログ表示。buttons: [{label, primary?, action?}]
 * ボタンを押すとダイアログを閉じてから action を実行する（action内で次のダイアログを出せる）
 */
function showDialog(message, buttons) {
  $('dialogMessage').textContent = message;
  const box = $('dialogButtons');
  box.innerHTML = '';
  buttons.forEach((b, i) => {
    const el = document.createElement('button');
    el.className = 'btn ' + (b.primary || (i === 0 && !buttons.some((x) => x.primary)) ? 'primary' : 'secondary');
    el.textContent = b.label;
    el.addEventListener('click', () => {
      $('dialog').hidden = true;
      if (b.action) b.action();
    });
    box.appendChild(el);
  });
  $('dialog').hidden = false;
}

/** カメラ上の案内（マスター未読込／開始前／カメラエラー） */
let coverAction = null;
function refreshCover(errorText) {
  const cover = $('cover');
  const sub = $('btnCoverSub');
  sub.hidden = true;
  if (!master) {
    $('coverText').textContent = 'マスターデータが読み込まれていません。\n転送.xlsx を選んでください。\n（ファイルはこの端末の中だけで使われ、外部には送信されません）';
    $('btnCover').textContent = '📂 転送.xlsx を選ぶ';
    coverAction = pickMasterFile;
  } else if (cameraError) {
    $('coverText').textContent = errorText || 'カメラを起動できませんでした。\n権限設定をご確認のうえ、再試行してください。';
    $('btnCover').textContent = '再試行';
    coverAction = () => { unlockAudio(); startCamera(); };
  } else if (!cameraStarted) {
    $('coverText').textContent = `マスター ${masterCount().toLocaleString('ja-JP')}件 読込済み`;
    $('btnCover').textContent = '📷 スキャン開始';
    coverAction = () => { unlockAudio(); startCamera(); };
  } else {
    cover.hidden = true;
    return;
  }
  cover.hidden = false;
}

/* ============================================================
   Excel出力
   ============================================================ */
function pad(n) { return String(n).padStart(2, '0'); }

function buildFileName() {
  const d = new Date();
  const ts = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  return `${EXPORT_PREFIX}_${deviceId}_${ts}.xlsx`;
}

/** A列にJANを文字列で並べた .xlsx（Android版と同じ形式） */
function buildXlsx(janValues) {
  const ws = XLSX.utils.aoa_to_sheet(janValues.map((j) => [{ t: 's', v: j }]));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
}

function handleExport() {
  if (session.records.length === 0) {
    showToast('出力できるスキャンデータがありません。');
    return;
  }
  let file;
  const fileName = buildFileName();
  try {
    const data = buildXlsx(session.records.map((r) => r.jan));
    file = new File([data], fileName, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (e) {
    showDialog(`Excel出力に失敗しました。時間をおいて再度お試しください。\n(${e.message})`, [{ label: 'OK' }]);
    return;
  }

  let canShare = false;
  try { canShare = !!(navigator.canShare && navigator.canShare({ files: [file] })); } catch (e) { /* 非対応 */ }

  const buttons = [];
  if (canShare) {
    buttons.push({
      label: '📤 共有・保存（ドライブ等へ）',
      primary: true,
      action: async () => {
        try {
          await navigator.share({ files: [file], title: fileName });
          afterExport(fileName);
        } catch (e) {
          if (e.name !== 'AbortError') showToast('共有できませんでした。ダウンロードをお試しください。');
        }
      },
    });
  }
  buttons.push({
    label: '⬇️ ダウンロード',
    primary: !canShare,
    action: () => {
      const url = URL.createObjectURL(file);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
      afterExport(fileName);
    },
  });
  buttons.push({ label: 'キャンセル' });
  showDialog(`Excelファイルを作成しました（${session.records.length}件）。\n${fileName}`, buttons);
}

function afterExport(fileName) {
  showDialog(`出力が完了しました（${fileName}）。\nスキャンデータをクリアして新しく始めますか？`, [
    { label: 'クリアする', primary: true, action: () => { session.clear(); updateCount(); } },
    { label: 'そのまま続ける' },
  ]);
}

function handleReset() {
  showDialog('現在のスキャンデータをすべて消去します。よろしいですか？', [
    { label: 'リセットする', primary: true, action: () => { session.clear(); updateCount(); } },
    { label: 'キャンセル' },
  ]);
}

function handleMasterButton() {
  if (!master) {
    pickMasterFile();
    return;
  }
  const d = new Date(master.loadedAt);
  const when = `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  showDialog(`マスター：${masterCount().toLocaleString('ja-JP')}件\n${master.file}\n（${when} 読込）`, [
    { label: '📂 転送.xlsx を読み込み直す', primary: true, action: pickMasterFile },
    { label: '閉じる' },
  ]);
}

/* ============================================================
   画面スリープ防止
   ============================================================ */
let wakeLock = null;
async function requestWakeLock() {
  try {
    if ('wakeLock' in navigator && !document.hidden) wakeLock = await navigator.wakeLock.request('screen');
  } catch (e) { /* 非対応 */ }
}

/* ============================================================
   起動
   ============================================================ */
const deviceId = getOrCreateDeviceId();

function init() {
  $('deviceIdLabel').textContent = `端末ID: ${deviceId}`;

  session.load();
  updateCount();
  master = loadJson(KEY_MASTER);
  if (master && !master.items) master = null;
  updateMasterStatus();
  refreshCover();

  $('btnCover').addEventListener('click', () => coverAction && coverAction());
  $('masterFile').addEventListener('change', onMasterFileSelected);
  $('btnMaster').addEventListener('click', handleMasterButton);
  $('btnExport').addEventListener('click', handleExport);
  $('btnReset').addEventListener('click', handleReset);

  $('btnMode').addEventListener('click', () => {
    mode = mode === 'individual' ? 'bulk' : 'individual';
    const isBulk = mode === 'bulk';
    $('btnMode').textContent = isBulk ? '一括' : '個別';
    $('btnMode').classList.toggle('active', isBulk);
    $('reticle').classList.toggle('hidden', isBulk);
  });

  $('btnZoom').addEventListener('click', () => {
    zoomedIn = !zoomedIn;
    $('btnZoom').textContent = zoomedIn ? '2x' : '1x';
    $('btnZoom').classList.toggle('active', zoomedIn);
    applyZoom();
  });

  $('btnTorch').addEventListener('click', () => {
    const caps = getCaps();
    if (!track || !caps.torch) {
      showToast('この端末はライト制御に対応していません。');
      return;
    }
    torchOn = !torchOn;
    track.applyConstraints({ advanced: [{ torch: torchOn }] }).catch(() => {
      torchOn = false;
      $('btnTorch').classList.remove('warn');
    });
    $('btnTorch').classList.toggle('warn', torchOn);
  });

  // どこかをタップしたら音を使えるようにする（iPhoneは操作後でないと音が出ない）
  document.addEventListener('pointerdown', unlockAudio);

  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) {
      if (cameraStarted) {
        // 別アプリから戻った時にカメラが止まっていたら再開する
        if (!track || track.readyState === 'ended') startCamera();
        else video.play().catch(() => {});
      }
      requestWakeLock();
    }
  });

  if (session.records.length > 0) {
    showDialog(`前回の作業データが${session.records.length}件残っています。\n引き継ぎますか？`, [
      { label: '引き継ぐ', primary: true },
      { label: 'リセットする', action: () => { session.clear(); updateCount(); } },
    ]);
  }

  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
}

init();
