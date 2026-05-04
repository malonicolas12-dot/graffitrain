// ── IndexedDB ────────────────────────────────────────────────────────────────

const DB_NAME = 'GraffitrainDB';
const STORE_NAME = 'photos';
let db;

function initDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = e => {
      const database = e.target.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        const store = database.createObjectStore(STORE_NAME, { keyPath: 'id' });
        store.createIndex('trainName', 'trainName', { unique: false });
      }
    };
    req.onsuccess = e => { db = e.target.result; resolve(); };
    req.onerror   = e => reject(e.target.error);
  });
}

function compressImage(file) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        const MAX = 1200;
        let w = img.width, h = img.height;
        if (w > MAX) { h = Math.round(h * MAX / w); w = MAX; }
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        canvas.toBlob(resolve, 'image/jpeg', 0.7);
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

function getPhotosForTrain(trainName) {
  return new Promise(resolve => {
    const tx  = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).index('trainName').getAll(trainName);
    req.onsuccess = () => resolve(req.result);
  });
}

function storePhoto(trainName, blob) {
  return new Promise(resolve => {
    const id = `${trainName}_${Date.now()}_${Math.random()}`;
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put({ id, trainName, blob, timestamp: Date.now() });
    tx.oncomplete = resolve;
  });
}

function deletePhotoDB(id) {
  return new Promise(resolve => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(id);
    tx.oncomplete = resolve;
  });
}

async function deleteAllPhotosForTrain(trainName) {
  const photos = await getPhotosForTrain(trainName);
  return new Promise(resolve => {
    const tx    = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    photos.forEach(p => store.delete(p.id));
    tx.oncomplete = resolve;
  });
}

function blobToBase64(blob) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.readAsDataURL(blob);
  });
}

// ── Données ──────────────────────────────────────────────────────────────────

function loadTrains() {
  const v3 = localStorage.getItem("graffitrain_v3");
  if (v3) return JSON.parse(v3).map(t => ({ surface: "", ...t }));

  const v2 = localStorage.getItem("graffitrain_v2");
  if (v2) return JSON.parse(v2).map(t => ({ name: t.name, status: t.status, surface: "" }));

  return Array.from({ length: 52 }, (_, i) => ({
    name: `CC${String(i + 1).padStart(2, "0")}`,
    status: "PAS VU",
    surface: ""
  }));
}

let trains = loadTrains();

function saveData() {
  localStorage.setItem("graffitrain_v3", JSON.stringify(trains));
}

// ── Logique ──────────────────────────────────────────────────────────────────

function getCardClass(status) {
  if (status === "OK")      return "ok-card";
  if (status === "GRAFFITÉ") return "graffiti-card";
  if (status === "SALE")    return "sale-card";
  return "pasvu-card";
}

function count(status) {
  return trains.filter(t => t.status === status).length;
}

async function addTrain() {
  const input = document.getElementById("newTrain");
  const name  = input.value.trim().toUpperCase();
  if (!name) return alert("Nom du train obligatoire");
  if (trains.some(t => t.name === name)) return alert("Ce train existe déjà");
  trains.push({ name, status: "PAS VU" });
  input.value = "";
  saveData();
  await render();
}

async function deleteTrain(name) {
  if (!confirm(`Supprimer ${name} ?`)) return;
  await deleteAllPhotosForTrain(name);
  trains = trains.filter(t => t.name !== name);
  saveData();
  await render();
}

async function setStatus(name, status) {
  const train = trains.find(t => t.name === name);
  if (train.status === status) {
    train.status = "PAS VU";
    await deleteAllPhotosForTrain(name);
  } else {
    if (status !== "GRAFFITÉ") await deleteAllPhotosForTrain(name);
    train.status = status;
  }
  saveData();
  await render();
}

function setSurface(name, value) {
  const train = trains.find(t => t.name === name);
  train.surface = value;
  saveData();
}

async function addPhoto(trainName, input) {
  const files    = Array.from(input.files);
  if (!files.length) return;
  const existing = await getPhotosForTrain(trainName);
  if (existing.length >= 2) return alert("Maximum 2 photos par train");
  const slots = 2 - existing.length;
  for (const file of files.slice(0, slots)) {
    const blob = await compressImage(file);
    await storePhoto(trainName, blob);
  }
  await render();
}

async function deletePhoto(id) {
  await deletePhotoDB(id);
  await render();
}

async function resetAll() {
  if (!confirm("Démarrer une nouvelle tournée ? Tous les statuts, surfaces et photos seront remis à zéro.")) return;
  for (const train of trains) {
    await deleteAllPhotosForTrain(train.name);
    train.status = "PAS VU";
    train.surface = "";
  }
  saveData();
  await render();
}

// ── Export Excel ─────────────────────────────────────────────────────────────

function getWeekNumber(date) {
  const d      = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

async function exportExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet    = workbook.addWorksheet("Rapport Graffitrain");
  const now      = new Date();
  const date     = now.toLocaleDateString("fr-FR");
  const week     = getWeekNumber(now);

  sheet.columns = [
    { key: "train",       width: 15 },
    { key: "statut",      width: 18 },
    { key: "surface",     width: 14 },
    { key: "smartracer",  width: 20 },
    { key: "photo1",      width: 35 },
    { key: "photo2",      width: 35 }
  ];

  sheet.mergeCells("A1:F1");
  sheet.getCell("A1").value     = "Rapport Graffitrain — Pôle Propreté Ligne 6";
  sheet.getCell("A1").font      = { bold: true, size: 18 };
  sheet.getCell("A1").alignment = { horizontal: "center" };

  sheet.mergeCells("A2:F2");
  sheet.getCell("A2").value     = `Semaine ${week} - Date : ${date}`;
  sheet.getCell("A2").font      = { bold: true };
  sheet.getCell("A2").alignment = { horizontal: "center" };

  sheet.getRow(4).values    = ["Train", "Statut", "Surface (m²)", "SMARTRACER", "Photo 1", "Photo 2"];
  sheet.getRow(4).alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(4).height    = 22;
  sheet.getRow(4).eachCell(cell => {
    cell.font   = { bold: true, color: { argb: "000000" } };
    cell.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF" } };
    cell.border = { top: { style: "thin", color: { argb: "000000" } }, left: { style: "thin", color: { argb: "000000" } }, bottom: { style: "thin", color: { argb: "000000" } }, right: { style: "thin", color: { argb: "000000" } } };
  });

  let rowNumber = 5;
  for (const train of trains) {
    const photos = await getPhotosForTrain(train.name);
    const row    = sheet.getRow(rowNumber);
    const surface = train.status === "GRAFFITÉ" && train.surface ? `${train.surface} m²` : "";
    row.values   = [train.name, train.status, surface, "", "", ""];

    let fillColor = "AAAAAA";
    if (train.status === "OK")       fillColor = "00B050";
    if (train.status === "GRAFFITÉ") fillColor = "D71920";
    if (train.status === "SALE")     fillColor = "7B4A24";
    if (train.status === "PAS VU")   fillColor = "AAAAAA";

    row.eachCell(cell => {
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
      cell.font      = { bold: true, color: { argb: "000000" } };
      cell.border    = { top: { style: "thin", color: { argb: "000000" } }, left: { style: "thin", color: { argb: "000000" } }, bottom: { style: "thin", color: { argb: "000000" } }, right: { style: "thin", color: { argb: "000000" } } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    sheet.getRow(rowNumber).height = photos.length ? 120 : 24;

    for (let i = 0; i < Math.min(photos.length, 2); i++) {
      const base64   = await blobToBase64(photos[i].blob);
      const imageId  = workbook.addImage({ base64, extension: "jpeg" });
      sheet.addImage(imageId, {
        tl:  { col: 4 + i, row: rowNumber - 1 },
        ext: { width: 180, height: 110 }
      });
    }
    rowNumber++;
  }

  const buffer    = await workbook.xlsx.writeBuffer();
  const cleanDate = date.replaceAll("/", "-");
  const fileName  = `Rapport_Graffitrain_S${week}_${cleanDate}.xlsx`;
  const blob      = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  return { blob, fileName, week, date };
}

function setLoading(btn, loading) {
  btn.disabled    = loading;
  btn.style.opacity = loading ? "0.6" : "1";
}

async function downloadExcel(btn) {
  try {
    if (btn) setLoading(btn, true);
    const { blob, fileName } = await exportExcel();
    saveAs(blob, fileName);
  } catch(e) {
    alert("Erreur : " + e.message);
  } finally {
    if (btn) setLoading(btn, false);
  }
}

async function shareExcel(btn) {
  try {
    if (btn) setLoading(btn, true);
    const { blob, fileName, week, date } = await exportExcel();

    // 1. Télécharger le fichier
    saveAs(blob, fileName);

    // 2. Ouvrir un email pré-rempli
    const subject = encodeURIComponent(`Rapport Graffitrain — Semaine ${week}`);
    const body    = encodeURIComponent(`Bonjour,\n\nVeuillez trouver ci-joint le rapport de contrôle des trains de la ligne 6 pour la semaine ${week} du ${date}.\n\nCordialement,\nPôle Propreté — Ligne 6`);
    window.open(`mailto:?subject=${subject}&body=${body}`);

    alert(`Le fichier "${fileName}" a été téléchargé.\nJoignez-le à l'email qui vient de s'ouvrir.`);
  } catch(e) {
    alert("Erreur : " + e.message);
  } finally {
    if (btn) setLoading(btn, false);
  }
}

// ── Rendu ────────────────────────────────────────────────────────────────────

let activeObjectURLs = [];

async function render() {
  activeObjectURLs.forEach(url => URL.revokeObjectURL(url));
  activeObjectURLs = [];

  trains.sort((a, b) => a.name.localeCompare(b.name));

  const photoMap = {};
  for (const train of trains) {
    if (train.status === "GRAFFITÉ") {
      const photos = await getPhotosForTrain(train.name);
      photoMap[train.name] = photos.map(p => {
        const url = URL.createObjectURL(p.blob);
        activeObjectURLs.push(url);
        return { id: p.id, url };
      });
    }
  }

  const isLight = document.body.classList.contains("light");

  document.getElementById("app").innerHTML = `
    <div class="app">
      <div class="header-wrap">
        <h1>
          <div class="h1-top">
            <img class="h1-logo" src="logo.svg" alt="Ligne 6">
            <span class="h1-title">Graffitrain</span>
          </div>
          <span class="h1-sub">Pôle Propreté — Ligne 6</span>
        </h1>
        <button id="theme-btn" class="theme-toggle" onclick="toggleTheme()">
          ${isLight ? "🌙 Mode sombre" : "☀️ Mode clair"}
        </button>
      </div>

      <div class="stats">
        <div class="stat">✅ OK<br>${count("OK")}</div>
        <div class="stat">🎨 Graffité<br>${count("GRAFFITÉ")}</div>
        <div class="stat">🧽 Sale<br>${count("SALE")}</div>
        <div class="stat">👀 Pas vu<br>${count("PAS VU")}</div>
      </div>

      <div class="top-actions">
        <button class="main" onclick="resetAll()">Nouvelle tournée</button>
        <button class="export" onclick="downloadExcel(this)">⬇️ Télécharger</button>
        <button class="share" onclick="shareExcel(this)">📤 Partager</button>
      </div>

      <div class="add-box">
        <input id="newTrain" placeholder="Ajouter un train ex : CC53">
        <button class="add" onclick="addTrain()">Ajouter</button>
      </div>

      ${trains.map(train => {
        const photos = photoMap[train.name] || [];
        return `
          <div class="train ${getCardClass(train.status)}">
            <div class="train-top">
              <div>
                <div class="train-title">${train.name}</div>
                <div class="badge">${train.status}</div>
              </div>
              <button class="delete-train" onclick="deleteTrain('${train.name}')">Supprimer</button>
            </div>

            <div class="buttons">
              <button class="ok      ${train.status === 'OK'       ? 'active' : ''}" onclick="setStatus('${train.name}', 'OK')">OK</button>
              <button class="graffiti ${train.status === 'GRAFFITÉ' ? 'active' : ''}" onclick="setStatus('${train.name}', 'GRAFFITÉ')">Graffité</button>
              <button class="sale    ${train.status === 'SALE'     ? 'active' : ''}" onclick="setStatus('${train.name}', 'SALE')">Sale</button>
            </div>

            ${train.status === "GRAFFITÉ" ? `
              <div class="surface-box">
                <label class="surface-label">Surface estimée</label>
                <div class="surface-input-wrap">
                  <input type="number" min="0" step="0.5" placeholder="0"
                         class="surface-input"
                         value="${train.surface}"
                         oninput="setSurface('${train.name}', this.value)">
                  <span class="surface-unit">m²</span>
                </div>
              </div>
              <input type="file" accept="image/*" capture="environment" multiple
                     id="photo-cam-${train.name}" style="display:none"
                     onchange="addPhoto('${train.name}', this)">
              <input type="file" accept="image/*" multiple
                     id="photo-gal-${train.name}" style="display:none"
                     onchange="addPhoto('${train.name}', this)">
              <button class="photo-btn" onclick="showPhotoMenu('${train.name}')">
                📷 Ajouter photo${photos.length > 0 ? ` (${photos.length}/2)` : ""}
              </button>
              <div class="photo-list">
                ${photos.map(({ id, url }) => `
                  <div>
                    <img class="photo" src="${url}">
                    <button class="delete" onclick="deletePhoto('${id}')">Supprimer photo</button>
                  </div>
                `).join("")}
              </div>
            ` : ""}
          </div>
        `;
      }).join("")}
    </div>
  `;
}

// ── Menu photo ───────────────────────────────────────────────────────────────

function showPhotoMenu(trainName) {
  const existing = document.getElementById('photo-menu');
  if (existing) existing.remove();

  const menu = document.createElement('div');
  menu.id = 'photo-menu';
  menu.innerHTML = `
    <div class="photo-menu-overlay" onclick="closePhotoMenu()"></div>
    <div class="photo-menu-sheet">
      <button class="photo-menu-btn" onclick="closePhotoMenu(); document.getElementById('photo-cam-${trainName}').click()">📷 Prendre une photo</button>
      <button class="photo-menu-btn" onclick="closePhotoMenu(); document.getElementById('photo-gal-${trainName}').click()">🖼️ Choisir dans la galerie</button>
      <button class="photo-menu-cancel" onclick="closePhotoMenu()">Annuler</button>
    </div>
  `;
  document.body.appendChild(menu);
}

function closePhotoMenu() {
  const menu = document.getElementById('photo-menu');
  if (menu) menu.remove();
}

// ── Thème ────────────────────────────────────────────────────────────────────

function applyTheme() {
  const light = localStorage.getItem("graffitrain_theme") === "light";
  document.body.classList.toggle("light", light);
}

function toggleTheme() {
  const isLight = document.body.classList.toggle("light");
  localStorage.setItem("graffitrain_theme", isLight ? "light" : "dark");
  document.getElementById("theme-btn").textContent = isLight ? "🌙 Mode sombre" : "☀️ Mode clair";
}

// ── Init ─────────────────────────────────────────────────────────────────────

applyTheme();
initDB().then(() => render());
