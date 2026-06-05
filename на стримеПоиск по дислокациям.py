<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>🏛 Поиск по дислокациям ФССП</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root {
  --bg:#0d1b2a;--bg2:#1a2d42;--bg3:#243447;
  --accent:#f5c518;--text:#e8f0fe;--text2:#a0b4c8;
  --border:#2e4a6a;--green-light:#b8e8c9;
  --operator-border:#f5c518;--card-bg:#1a2d42;
  --header-bg:#0a1520;--btn-danger:#c0392b;
  --btn-ok:#27ae60;--btn-blue:#2980b9;
  --shadow:0 4px 24px rgba(0,0,0,0.5);
  --radius:10px;
  --font:'Segoe UI',Arial,sans-serif;
  --fs-base:18px;--fs-sm:16px;--fs-xs:14px;--fs-lg:21px;--fs-xl:26px;
  --highlight-bg:#ffe066;--highlight-color:#1a1a1a;
}
.light-theme{
  --bg:#f0f4f8;--bg2:#ffffff;--bg3:#e2e8f0;
  --accent:#1a56db;--text:#1a202c;--text2:#4a5568;--border:#cbd5e0;
  --operator-border:#1a56db;--card-bg:#ffffff;--header-bg:#e2e8f0;
  --btn-danger:#e53e3e;--btn-ok:#38a169;--btn-blue:#3182ce;
  --shadow:0 4px 24px rgba(0,0,0,0.12);
}
*{box-sizing:border-box;margin:0;padding:0;}
html{font-size:var(--fs-base);scroll-behavior:smooth;}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:100px;}

/* ── FIXED TOP ── */
#fixed-top{position:fixed;top:0;left:0;right:0;z-index:1000;background:var(--header-bg);border-bottom:2px solid var(--border);box-shadow:var(--shadow);}

#results-bar{
  background:var(--bg2);padding:10px 16px;
  font-size:var(--fs-lg);font-weight:700;min-height:48px;
  display:flex;align-items:center;justify-content:center;flex-wrap:wrap;gap:8px;
  border-bottom:1px solid var(--border);color:var(--text);text-align:center;
}
#results-bar .badge{
  background:var(--accent);color:#1a1a1a;
  border-radius:8px;padding:8px 20px;
  font-size:var(--fs-lg);font-weight:800;
  cursor:pointer;user-select:none;border:none;
  transition:filter .15s,transform .15s;
  min-height:44px;display:inline-flex;align-items:center;
}
#results-bar .badge:hover{filter:brightness(1.18);transform:translateY(-2px);}

/* ── TOOLBAR ── */
#toolbar{
  display:flex;align-items:center;justify-content:center;
  flex-wrap:wrap;gap:8px;padding:8px 12px;background:var(--header-bg);
}
.toolbar-btn{
  padding:8px 16px;border:none;border-radius:8px;
  font-size:var(--fs-sm);font-weight:700;cursor:pointer;
  display:inline-flex;align-items:center;gap:6px;
  transition:all .18s;white-space:nowrap;color:#fff;
}
.toolbar-btn:hover{filter:brightness(1.18);transform:translateY(-1px);}
/* Кнопка Поиск — вдвое больше */
.btn-search{
  background:var(--accent);color:#1a1a1a!important;
  font-size:calc(var(--fs-lg) + 2px);padding:14px 32px;
  border-radius:10px;font-weight:900;
}
.btn-excel{background:var(--btn-ok);}
.btn-reset{background:#555;}
.btn-theme{background:var(--bg3);color:var(--text)!important;border:1px solid var(--border);}
.btn-settings{background:#8e44ad;}
.btn-expand-all{background:#16a085;}
.btn-collapse-all{background:#555;}

#stats-bar{
  padding:6px 12px;font-size:var(--fs-sm);color:var(--text2);
  background:var(--bg3);border-bottom:1px solid var(--border);
  display:flex;gap:16px;flex-wrap:wrap;justify-content:center;
}
.stat-item strong{color:var(--accent);}

/* ── CONTENT ── */
#content{padding:12px 10px;}
#blocks-container{display:flex;flex-direction:column;align-items:center;}

/* ── OPERATOR BLOCK ── */
.op-block{
  background:var(--card-bg);border:3px solid var(--operator-border);
  border-radius:var(--radius);margin-bottom:16px;overflow:hidden;
  transition:box-shadow .2s;width:100%;max-width:900px;
}
.op-block:hover{box-shadow:0 0 18px rgba(245,197,24,.2);}
.op-block.highlighted{background:var(--highlight-bg)!important;border-color:#e67e22!important;}
.op-block.highlighted .op-header{background:#e67e22!important;color:#fff!important;}
.op-block.highlighted .op-header *{color:#fff!important;}

.op-header{
  background:var(--bg3);padding:12px 18px;
  display:flex;align-items:center;justify-content:center;
  flex-wrap:wrap;gap:8px;cursor:pointer;user-select:none;
  border-bottom:2px solid var(--operator-border);text-align:center;
}
.op-header:hover{filter:brightness(1.08);}
.op-num{font-size:var(--fs-xl);font-weight:900;color:var(--accent);margin-right:8px;min-width:38px;}
.op-code{font-size:var(--fs-lg);font-weight:800;color:var(--accent);}
.op-name{font-size:var(--fs-lg);font-weight:700;color:var(--text);}
.op-toggle{font-size:22px;color:var(--text2);margin-left:auto;transition:transform .2s;}
.op-toggle.open{transform:rotate(180deg);}
.op-body{padding:12px 18px 18px;}
.op-body.collapsed{display:none;}

.subj-block{margin-bottom:14px;text-align:center;}
.subj-title{
  font-size:var(--fs-base);font-weight:800;color:var(--accent);
  text-transform:uppercase;letter-spacing:.5px;
  border-bottom:1px solid var(--border);
  padding-bottom:4px;margin-bottom:6px;display:inline-block;width:100%;
}
.osp-list{list-style:none;padding-left:0;display:flex;flex-direction:column;align-items:center;}
.osp-item{
  font-size:var(--fs-sm);color:var(--text2);
  padding:4px 10px;border-radius:4px;margin:2px 0;
  display:flex;align-items:center;gap:7px;max-width:700px;width:100%;
}
.osp-item::before{content:'▸';color:var(--accent);font-size:14px;flex-shrink:0;}
.osp-item.highlighted{background:var(--highlight-bg);color:var(--highlight-color)!important;font-weight:700;}
.subj-title.highlighted{background:var(--highlight-bg);color:var(--highlight-color);padding:4px 12px;border-radius:4px;}

/* ── EXPAND/COLLAPSE BAR ── */
#expand-all-bar{display:flex;justify-content:center;gap:10px;padding:8px 10px 4px;}

/* ── SCROLLBAR ── */
#scrollbar-track{
  position:fixed;right:10px;top:50%;transform:translateY(-50%);
  display:flex;flex-direction:column;align-items:center;gap:6px;
  z-index:900;background:var(--bg2);
  border:1px solid var(--border);border-radius:16px;
  padding:5px 4px;box-shadow:var(--shadow);
}
.scroll-btn{
  background:var(--bg3);border:1px solid var(--border);
  color:var(--text);border-radius:50%;width:36px;height:36px;
  display:flex;align-items:center;justify-content:center;
  cursor:pointer;font-size:20px;transition:.15s;user-select:none;
}
.scroll-btn:hover{background:var(--accent);color:#1a1a1a;}
#scroll-rail{width:10px;height:130px;background:var(--bg3);border-radius:10px;position:relative;cursor:pointer;border:1px solid var(--border);}
#scroll-thumb{width:10px;height:32px;background:var(--accent);border-radius:10px;position:absolute;left:0;top:0;cursor:grab;}

/* ── MODALS ── */
.modal-overlay{
  position:fixed;inset:0;background:rgba(0,0,0,.75);
  z-index:2000;display:flex;align-items:center;justify-content:center;padding:10px;
}
.modal-overlay.hidden{display:none;}
.modal-box{
  background:var(--bg2);border:2px solid var(--border);
  border-radius:14px;padding:26px;max-width:680px;width:100%;
  box-shadow:var(--shadow);max-height:90vh;overflow-y:auto;position:relative;
  text-align:center;
}
.modal-title{font-size:var(--fs-xl);font-weight:800;color:var(--accent);margin-bottom:18px;}
.modal-close{
  position:absolute;top:12px;right:16px;
  background:none;border:none;color:var(--text2);font-size:26px;cursor:pointer;
}
.modal-close:hover{color:var(--btn-danger);}

/* Search modal */
.search-type-btns{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px;justify-content:center;}
.type-btn{
  padding:10px 20px;border-radius:9px;border:2px solid var(--border);
  background:var(--bg3);color:var(--text);
  font-size:var(--fs-base);font-weight:700;cursor:pointer;transition:.18s;
}
.type-btn.active{background:var(--accent);color:#1a1a1a;border-color:var(--accent);}
.type-btn:hover:not(.active){border-color:var(--accent);}
#search-input{
  width:100%;padding:14px 18px;border-radius:9px;
  border:2px solid var(--border);background:var(--bg3);
  color:var(--text);font-size:var(--fs-xl);margin-bottom:18px;
  outline:none;font-family:var(--font);text-align:center;
}
#search-input:focus{border-color:var(--accent);}
.search-actions{display:flex;gap:10px;flex-wrap:wrap;justify-content:center;}

/* Detail modal */
#detail-modal .modal-box{max-width:750px;}

/* Settings modal */
#settings-modal .modal-box{max-width:750px;}
.settings-section{margin-bottom:20px;border:1px solid var(--border);border-radius:10px;padding:16px;}
.settings-section-title{font-size:var(--fs-base);font-weight:800;color:var(--accent);margin-bottom:12px;}
.color-grid{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;}
.color-swatch{
  width:40px;height:40px;border-radius:50%;cursor:pointer;
  border:3px solid transparent;transition:.18s;
}
.color-swatch:hover,.color-swatch.selected{border-color:var(--accent);transform:scale(1.18);}
.font-grid{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;}
.font-btn{
  padding:8px 16px;border-radius:8px;border:2px solid var(--border);
  background:var(--bg3);color:var(--text);font-size:var(--fs-sm);
  cursor:pointer;transition:.18s;
}
.font-btn:hover,.font-btn.selected{border-color:var(--accent);background:var(--accent);color:#1a1a1a;}
.btn-save-settings{background:var(--btn-ok);margin-top:12px;font-size:var(--fs-base);padding:12px 28px;}
.btn-preview{background:var(--btn-blue);margin-top:8px;font-size:var(--fs-sm);padding:8px 18px;}

/* ── RABBIT ── */
#rabbit-wrap{position:fixed;bottom:20px;right:20px;z-index:1500;display:flex;flex-direction:column;align-items:flex-end;gap:8px;}
#rabbit-btn{
  width:58px;height:58px;background:var(--accent);
  border-radius:50%;display:flex;align-items:center;justify-content:center;
  font-size:34px;cursor:pointer;border:none;
  box-shadow:var(--shadow);animation:hop 3s ease-in-out infinite;
}
@keyframes hop{0%,90%,100%{transform:translateY(0)}50%{transform:translateY(-13px)}}
#rabbit-popup{
  background:var(--bg2);border:2px solid var(--accent);
  border-radius:14px;padding:18px;width:320px;max-width:93vw;
  box-shadow:var(--shadow);font-size:var(--fs-sm);color:var(--text);
  display:none;position:relative;text-align:left;
}
#rabbit-popup.open{display:block;}
#rabbit-popup h4{color:var(--accent);margin-bottom:12px;font-size:var(--fs-base);}
#rabbit-popup ol{padding-left:20px;}
#rabbit-popup li{margin-bottom:8px;line-height:1.55;}
#rabbit-timer{color:var(--text2);font-size:var(--fs-xs);margin-top:10px;text-align:right;}
#rabbit-close-btn{position:absolute;top:8px;right:12px;background:none;border:none;color:var(--text2);font-size:20px;cursor:pointer;}

/* ── RESPONSIVE ── */
@media(max-width:600px){
  :root{--fs-base:15px;--fs-sm:13px;--fs-lg:17px;--fs-xl:21px;}
  #toolbar{gap:5px;padding:5px 6px;}
  .toolbar-btn{padding:6px 10px;font-size:13px;}
  .btn-search{font-size:16px;padding:10px 18px;}
  #scrollbar-track{display:none;}
  .op-block{max-width:100%;}
}
@media(max-width:380px){
  :root{--fs-base:14px;--fs-sm:12px;--fs-lg:15px;--fs-xl:19px;}
  .toolbar-btn span{display:none;}
  .btn-search span{display:inline!important;}
}
</style>
</head>
<body>

<!-- FIXED TOP -->
<div id="fixed-top">
  <div id="results-bar">
    <span id="results-label">🔎 Результаты по запросу:</span>
    <span id="result-badges"></span>
  </div>
  <div id="toolbar">
    <button class="toolbar-btn btn-search" onclick="openSearch()" title="F2">🔍 <span>Поиск (F2)</span></button>
    <button class="toolbar-btn btn-excel"  onclick="exportExcel()">📊 <span>Excel</span></button>
    <button class="toolbar-btn btn-reset"  onclick="resetSearch()">♻️ <span>Сброс</span></button>
    <button class="toolbar-btn btn-settings" onclick="openSettings()">🎨 <span>Настройки</span></button>
    <button class="toolbar-btn btn-theme"  onclick="toggleTheme()">🌗 <span id="theme-label">☀️ Светлая</span></button>
  </div>
  <div id="stats-bar">
    <span>👥 Операторов: <strong id="stat-ops">0</strong></span>
    <span>🗺 Субъектов: <strong id="stat-subj">0</strong></span>
    <span>🏛 ОСП/РОСП: <strong id="stat-osp">0</strong></span>
    <span id="stat-source"></span>
  </div>
</div>

<div id="content">
  <div id="expand-all-bar">
    <button class="toolbar-btn btn-expand-all"  onclick="expandAll()">📂 Развернуть все</button>
    <button class="toolbar-btn btn-collapse-all" onclick="collapseAll()">📁 Свернуть все</button>
  </div>
  <div id="blocks-container"></div>
</div>

<!-- SCROLLBAR -->
<div id="scrollbar-track">
  <div class="scroll-btn" onclick="window.scrollTo({top:0,behavior:'smooth'})">▲</div>
  <div id="scroll-rail"><div id="scroll-thumb"></div></div>
  <div class="scroll-btn" onclick="window.scrollTo({top:document.body.scrollHeight,behavior:'smooth'})">▼</div>
</div>

<!-- SEARCH MODAL -->
<div id="search-modal" class="modal-overlay hidden">
  <div class="modal-box">
    <div class="modal-title">🔍 Параметры поиска</div>
    <button class="modal-close" onclick="closeSearch()">✕</button>
    <div class="search-type-btns">
      <button class="type-btn active" id="btn-type-subj" onclick="setSearchType('subj')">📍 По Субъекту РФ</button>
      <button class="type-btn"        id="btn-type-osp"  onclick="setSearchType('osp')">🏛 По РОСП/ОСП</button>
      <button class="type-btn"        id="btn-type-any"  onclick="setSearchType('any')">🔁 Любое совпадение</button>
    </div>
    <input type="text" id="search-input" placeholder="Введите запрос..."
           autocomplete="off" onkeydown="searchKeydown(event)"/>
    <div class="search-actions">
      <button class="toolbar-btn btn-search" onclick="doSearch()">🔍 Найти</button>
      <button class="toolbar-btn btn-reset"  onclick="resetSearch();closeSearch()">♻️ Сброс</button>
      <button class="toolbar-btn" style="background:#555" onclick="closeSearch()">✕ Закрыть</button>
    </div>
  </div>
</div>

<!-- DETAIL MODAL -->
<div id="detail-modal" class="modal-overlay hidden">
  <div class="modal-box">
    <div class="modal-title" id="detail-title">Оператор</div>
    <button class="modal-close" onclick="closeDetail()">✕</button>
    <div id="detail-content"></div>
  </div>
</div>

<!-- SETTINGS MODAL -->
<div id="settings-modal" class="modal-overlay hidden">
  <div class="modal-box">
    <div class="modal-title">🎨 Настройки интерфейса</div>
    <button class="modal-close" onclick="closeSettings()">✕</button>

    <div class="settings-section">
      <div class="settings-section-title">🖌 Цвет акцента (10 вариантов)</div>
      <div class="color-grid" id="color-grid"></div>
    </div>

    <div class="settings-section">
      <div class="settings-section-title">🔤 Шрифт интерфейса (10 вариантов)</div>
      <div class="font-grid" id="font-grid"></div>
    </div>

    <div class="settings-section">
      <div class="settings-section-title">📏 Размер шрифта</div>
      <div style="display:flex;align-items:center;gap:12px;justify-content:center;flex-wrap:wrap;">
        <label style="font-size:var(--fs-sm)">Базовый размер:</label>
        <input type="range" id="font-size-range" min="14" max="26" value="18"
               style="width:180px;" oninput="previewFontSize(this.value)">
        <span id="font-size-label" style="font-weight:700;color:var(--accent);min-width:40px;">18px</span>
      </div>
    </div>

    <div class="settings-section">
      <div class="settings-section-title">✨ Цвет подсветки найденных строк</div>
      <div class="color-grid" id="highlight-grid"></div>
    </div>

    <div style="text-align:center;margin-top:12px;display:flex;gap:10px;justify-content:center;flex-wrap:wrap;">
      <button class="toolbar-btn btn-preview" onclick="applySettingsPreview()">👁 Предпросмотр</button>
      <button class="toolbar-btn btn-save-settings" onclick="saveSettings()">💾 Сохранить изменения</button>
    </div>
    <p style="font-size:var(--fs-xs);color:var(--text2);margin-top:10px;text-align:center;line-height:1.6;">
      После «Сохранить изменения» нажмите <b>Ctrl+S</b> (или «Сохранить как…») в браузере,<br>
      чтобы сохранить HTML-файл с новыми настройками постоянно.
    </p>
  </div>
</div>

<!-- RABBIT -->
<div id="rabbit-wrap">
  <div id="rabbit-popup">
    <button id="rabbit-close-btn" onclick="closeRabbit()">✕</button>
    <h4>🐰 Справка по работе с приложением</h4>
    <ol>
      <li>🔍 Нажмите <b>«Поиск»</b> или клавишу <b>F2</b> — откроются параметры поиска.</li>
      <li>📍 Выберите тип: <b>По Субъекту РФ</b>, <b>По РОСП/ОСП</b> или <b>Любое совпадение</b>.</li>
      <li>⌨️ Введите запрос → нажмите <b>«Найти»</b> или <b>Enter</b>. Совпадения подсветятся цветом.</li>
      <li>🏷 В шапке «Результаты» — бейджи операторов. <b>Нажмите на бейдж</b> — откроется карточка.</li>
      <li>📂 Нажимайте на заголовок блока оператора — он раскроется/свернётся.</li>
      <li>🎨 Кнопка <b>«Настройки»</b> — выбор цвета, шрифта, размера, сохранение в HTML.</li>
      <li>📊 Кнопка <b>«Excel»</b> — выгрузка всех данных в файл.</li>
      <li>♻️ <b>«Сброс»</b> — сбросить результаты поиска.</li>
    </ol>
    <div id="rabbit-timer"></div>
  </div>
  <button id="rabbit-btn" onclick="toggleRabbit()" title="Помощник">🐰</button>
</div>

<script>
// ══════════════════════════════════════════════════════════════════
//  ПОЛНАЯ БАЗА ДАННЫХ (из всех трёх файлов xlsx — объединено)
// ══════════════════════════════════════════════════════════════════
const RAW_DATA = [
  {num:1,code:'822',name:'Курбонов Зафар',subjects:[
    {subj:'Новосибирская область',osps:[]},
    {subj:'Кировская область',osps:[]},
    {subj:'г. Москва',osps:['Алтуфьевский ОСП УФССП по Москве','Бабушкинский ОСП УФССП по Москве','Гагаринский ОСП УФССП России по Москве']}
  ]},
  {num:2,code:'823',name:'Максимов Александр',subjects:[
    {subj:'Омская область',osps:[]},
    {subj:'Республика Крым',osps:[]},
    {subj:'Камчатский Край',osps:[]},
    {subj:'г. Москва',osps:['Головинский ОСП УФССП по Москве','ГУФССП России по г.Москве','Даниловский ОСП УФССП по Москве','Тропарево-Никулинский ОСП УФССП по Москве']}
  ]},
  {num:3,code:'825',name:'Даньшина Ирина',subjects:[
    {subj:'Республика Татарстан',osps:[]},
    {subj:'Тульская область',osps:[]},
    {subj:'г. Москва',osps:['ОСП по ЦАО №2 г.Москвы УФССП по Москве','ОСП по Юго-Восточному АО УФССП России по Москве','Царицынский ОСП УФССП по Москве']}
  ]},
  {num:4,code:'826',name:'Хмелевская Ольга',subjects:[
    {subj:'Кемеровская область',osps:[]},
    {subj:'Курганская область',osps:[]},
    {subj:'Томская область',osps:[]},
    {subj:'г. Москва',osps:['Кунцевский ОСП УФССП России по Москве','Люблинский ОСП УФССП по Москве','МОСП по ОИП УФССП России по Москве','Чертановский ОСП УФССП России по Москве']}
  ]},
  {num:5,code:'457',name:'Сайфутдинова Регина',subjects:[
    {subj:'Республика Удмуртия',osps:[]},
    {subj:'г. Санкт-Петербург',osps:['Адмиралтейский РОСП УФССП по Санкт-Петербургу','Василеостровский РОСП УФССП по Санкт-Петербургу','Волковский ОСП Фрунзенского района УФССП по Санкт-Петербургу','Восточный ОСП Приморского района УФССП по г.Санкт-Петербургу','Выборгский РОСП УФССП России по Санкт-Петербургу','Западный ОСП УФССП по Санкт-Петербургу','Калининский РОСП УФССП по Санкт-Петербургу','Кировский РОСП УФССП по Санкт-Петербургу','Колпинский РОСП УФССП по Санкт-Петербургу','Красносельский РОСП УФССП по Санкт-Петербургу','Купчинский ОСП по Фрунзенскому району УФССП по Санкт-Петербургу','Ладожский ОСП Красногвардейского района УФССП по Санкт-Петербургу']},
    {subj:'г. Москва',osps:['ОСП по ЮЗАО г.Москвы УФССП России по Москве','Останкинский ОСП УФССП по Москве','Перовский РОСП УФССП по Москве']},
    {subj:'Ростовская область',osps:['Мартыновское районное ОСП','ОСП по Миллеровскому и Тарасовскому районам','ОСП по Морозовскому и Милютинскому районам','Мясниковское районное ОСП','Неклиновское районное ОСП','Новочеркасский городской отдел судебных приставов','Новошахтинское городское ОСП','ОСП по Обливскому и Советскому районам','Орловское районное ОСП','Песчанокопское районное ОСП','Пролетарское районное ОСП','Р.Несветайское районное ОСП','Сальское районное ОСП','Семикаракорское районное ОСП','Тацинское районное ОСП','Целинское районное ОСП','Цимлянское районное ОСП','Чертковское районное ОСП','Отдел судебных приставов по г.Шахты и Октябрьскому району','ОСП по Шолоховскому и Верхнедонскому районам','Октябрьское районное ОСП г.Ростова-на-Дону','Первомайское районное ОСП г.Ростова-на-Дону','Пролетарское районное ОСП г.Ростова-на-Дону','Советское районное ОСП г.Ростова-на-Дону']}
  ]},
  {num:6,code:'491',name:'Попков Илья',subjects:[
    {subj:'Свердловская область',osps:[]},
    {subj:'г. Москва',osps:['Дмитровский ОСП УФССП по Москве','Зюзинский ОСП УФССП по Москве','Измайловский РОСП УФССП России по Москве','Черемушкинский ОСП УФССП России по Москве']},
    {subj:'Краснодарский край',osps:['Апшеронский РОСП УФССП','Белореченский РОСП УФССП','Горячеключевский ГОСП УФССП','Динской РОСП УФССП по Краснодарскому краю','Каневский РОСП УФССП по Краснодарскому краю']}
  ]},
  {num:7,code:'803',name:'Волосникова Ольга',subjects:[
    {subj:'Московская область',osps:['Воскресенский РОСП УФССП России по Московской области','Домодедовский ГОСП УФССП по Московской области','Дубненский ГОСП УФССП по Московской области','Егорьевский РОСП УФССП по Московской области','Жуковский ГОСП УФССП по Московской области','Зарайский РОСП УФССП по Московской области','Каширский РОСП УФССП по Московской области','Коломенский РОСП УФССП по Московской области','Королевский ГОСП УФССП по Московской области','Ленинский РОСП УФССП по Московской области','Озерский РОСП УФССП по Московской области','Подольский РОСП УФССП России по Московской области','Раменский РОСП УФССП России по Московской области','Рузский РОСП УФССП России по Московской области','Солнечногорский РОСП УФССП России по Московской области','Ступинский РОСП УФССП по Московской области','Шатурский РОСП УФССП по Московской области','Щелковский РОСП УФССП России по Московской области']},
    {subj:'г. Москва',osps:['Дорогомиловский ОСП УФССП по Москве','Коптевский ОСП УФССП по Москве','ОСП по Зеленоградскому АО УФССП по Москве','ОСП по Новомосковскому АО УФССП России по Москве','Савеловский ОСП УФССП России по Москве']},
    {subj:'Иркутская область',osps:['Ангарский РОСП УФССП по Иркутской области','Нижнеудинский РОСП УФССП по Иркутской области']},
    {subj:'Краснодарский край',osps:['Калининский РОСП','Красноармейский РОСП','Крымский РОСП','Курганинский РОСП','Кущевский РОСП']},
    {subj:'Ростовская область',osps:['Азовское районное ОСП','Аксайское районное ОСП','Багаевское районное ОСП','Батайское городское ОСП','Белокалитвинское районное ОСП','Веселовское районное ОСП','Отдел СП по г.Волгодонску и Волгодонскому району','Гуковское городское ОСП','Донецкое городское ОСП','Дубовское районное ОСП','Егорлыкское районное ОСП','Зверевское городское ОСП','ОСП по Заветинскому и Ремонтненскому районам','ОСП по Зерноградскому и Кагальницкому районам','Зимовниковское районное ОСП','Каменское районное ОСП','ОСП по Кашарскому и Боковскому районам','ОСП по Константиновскому и Усть-Донецкому районам','Красносулинское районное ОСП','ОСП по Куйбышевскому и Матвеево-Курганскому районам','Таганрогский городской отдел СП','Ворошиловское районное ОСП г.Ростова-на-Дону','Железнодорожное районное ОСП г.Ростова-на-Дону','Кировское районное ОСП г.Ростова-на-Дону']}
  ]},
  {num:8,code:'461',name:'Мунтанион Катя',subjects:[
    {subj:'Республика Башкортостан',osps:['Октябрьский РОСП г.Уфы УФССП по Республике Башкортостан','Орджоникидзевский РОСП г.Уфы УФССП по Республике Башкортостан','Стерлибашевский РОСП УФССП по Республике Башкортостан','Туймазинский МОСП УФССП по Республике Башкортостан','Уфимский РОСП УФССП по Республике Башкортостан','Федоровский РОСП УФССП по Республике Башкортостан','Чишминский РОСП УФССП по Республике Башкортостан','Абзелиловский РОСП УФССП по Республике Башкортостан','Агидельский ГОСП УФССП по Республике Башкортостан','Альшеевский РОСП УФССП по Республике Башкортостан','Аургазинский РОСП УФССП по Республике Башкортостан','Баймакский МОСП УФССП по Республике Башкортостан','Бакалинский РОСП УФССП по Республике Башкортостан','Белебеевский МОСП УФССП по Республике Башкортостан','Белокатайский РОСП УФССП по Республике Башкортостан','Бижбулякский РОСП УФССП по Республике Башкортостан','Бирский МОСП УФССП по Республике Башкортостан','Благоварский РОСП УФССП по Республике Башкортостан','Благовещенский МОСП УФССП по Республике Башкортостан','Буздякский РОСП УФССП по Республике Башкортостан','Бураевский РОСП УФССП по Республике Башкортостан','Ермекеевский РОСП УФССП по Республике Башкортостан','Зилаирский РОСП УФССП по Республике Башкортостан','Иглинский РОСП УФССП по Республике Башкортостан','Ишимбайский МОСП УФССП по Республике Башкортостан','Калининский РОСП г.Уфы УФССП по Республике Башкортостан','Калтасинский РОСП УФССП по Республике Башкортостан','Караидельский РОСП УФССП по Республике Башкортостан','Кармаскалинский РОСП УФССП по Республике Башкортостан','Кировский РОСП г.Уфы УФССП по Республике Башкортостан','Кугарчинский РОСП УФССП по Республике Башкортостан','Ленинский РОСП г.Уфы УФССП по Республике Башкортостан','Мелеузовский МОСП УФССП по Республике Башкортостан','Мечетлинский РОСП УФССП по Республике Башкортостан','Миякинский РОСП УФССП по Республике Башкортостан','Мишкинский РОСП УФССП по Республике Башкортостан','Нефтекамский МОСП УФССП по Республике Башкортостан','Октябрьский ГОСП УФССП России по Республике Башкортостан','Салаватский ГОСП УФССП по Республике Башкортостан','Салаватский РОСП УФССП по Республике Башкортостан','Советский РОСП г.Уфы УФССП по Республике Башкортостан','Татышлинский РОСП УФССП по Республике Башкортостан','Хайбуллинский РОСП УФССП по Республике Башкортостан','Чекмагушевский РОСП УФССП России по Республике Башкортостан','Шаранский РОСП УФССП по Республике Башкортостан','Янаульский МОСП УФССП по Республике Башкортостан','Белорецкий МОСП УФССП по Республике Башкортостан','Сибайский ГОСП УФССП по Республике Башкортостан','Стерлитамакский ГОСП УФССП по Республике Башкортостан','Стерлитамакский РОСП УФССП России по Республике Башкортостан','Учалинский МОСП УФССП по Республике Башкортостан']},
    {subj:'Челябинская область',osps:['Советский РОСП','Копейский ГОСП','Сосновский РОСП','Калининский РОСП']},
    {subj:'г. Москва',osps:['ОСП по Северо-Западному АО УФССП по Москве','ОСП по Троицкому району АО УФССП по Москве','ОСП по ЦАО №1 г.Москвы УФССП России по Москве']},
    {subj:'Республика Карелия',osps:[]}
  ]},
    {num:9,code:'481',name:'Неджар Наталья',subjects:[
    {subj:'Новгородская область',osps:[]},
    {subj:'Приморский край',osps:[]},
    {subj:'Тверская область',osps:['Бологовский РОСП','Заволжский РОСП','Зубцовский РОСП','Калининский РОСП','Кимрский РОСП','Лихославельский РОСП','ОСП по Вышневолоцкому','ОСП по Западнодвинскому','ОСП по Краснохолмскому','ОСП по Нелидовскому','ОСП по Осташковскому','Селижаровский РОСП','Торжокский РОСП','Удомельский РОСП']},
    {subj:'Белгородская область',osps:['Валуйский РОСП','Грайворонский РОСП','МОСП по ИОИП','Яковлевский РОСП']}
  ]},
  {num:10,code:'488',name:'Кара Анна',subjects:[
    {subj:'Тамбовская область',osps:[]},
    {subj:'Московская область',osps:['Долгопрудненский РОСП','Красногорский РОСП','Лобненский ГОСП','Мытищинский РОСП','Реутовский ГОСП','Сергиево-Посадский РОСП','Талдомский РОСП','Чеховский РОСП','Химкинский РОСП','Электростальский ГОСП']},
    {subj:'Иркутская область',osps:['Братский МОСП по ОПИ','Казачинско-Ленский РОСП','Катангский РОСП','Киренский РОСП','Куйтунский РОСП','Ленинский ОСП','МОСП по ИОИП УФССП','Нижнеилимский РОСП','ОСП по ВАШ по г.Иркутску','ОСП по г.Саянску и Зиминскому району','Падунский ОСП г.Братска','Правобережный ОСП','Тулунский РОСП','Усть-Илимский РОСП','Усть-Кутский РОСП','Усольский РОСП','Свердловский ОСП','Слюдянский РОСП','Тайшетский РОСП','Чунский РОСП','Шелеховский РОСП']},
    {subj:'Белгородская область',osps:['Белгородский РОСП','Борисовский РОСП','Новооскольский РОСП','Чернянский РОСП']},
    {subj:'Красноярский край',osps:['МОСП по Ачинску, Ачинскому и Большеулуйскому районам','ОСП по ИИДАП по г.Красноярску','ОСП №2 по Советскому району г.Красноярска','ОСП по Абанскому району','ОСП по Богучанскому району','ОСП по г.Игарка','ОСП по г.Минусинску и Минусинскому району','ОСП по г.Сосновоборску','ОСП по Кежемскому району','ОСП по Козульскому району','ОСП по Новоселовскому району','ОСП по Рыбинскому району и г.Бородино']}
  ]},
  {num:11,code:'485',name:'Захарова Катя',subjects:[
    {subj:'Республика Тыва',osps:[]},
    {subj:'Псковская область',osps:[]},
    {subj:'Смоленская область',osps:[]},
    {subj:'Красноярский край',osps:['ОСП по Балахтинскому району','ОСП по г.Дивногорску','ОСП по г.Канск и Канскому району','ОСП по Железнодорожному району г.Красноярска','ОСП по Каратузскому','ОСП по Курагинскому району','ОСП по Назаровскому району','ОСП по Шушенскому району']},
    {subj:'Белгородская область',osps:['ОСП по Алексеевскому и Красненскому','ОСП по Краснояружскому и Ракитянскому','Ровеньский РОСП','Старооскольский РОСП']}
  ]},
  {num:12,code:'487',name:'Турушева Марина',subjects:[
    {subj:'Республика Мордовия',osps:[]},
    {subj:'Тверская область',osps:['Дублянский РОСП','Калязинский РОСП','Конаковский РОСП','Кувшиновский РОСП','Московский РОСП','ОСП по Алексинскому','ОСП по Бежецкому','ОСП по Весьегонскому','ОСП по Кашинскому','ОСП по Максатихинскому','ОСП по Московскому району','Пролетарский РОСП','Ремешковский РОСП','Ржевский РОСП','Старицкий РОСП','Торопецкий РОСП','Центральный РОСП']},
    {subj:'Красноярский край',osps:['МОСП по ИОИП','ОСП по Березовскому району','ОСП по Бирилюсскому району','ОСП по Боготольскому району','ОСП по Дзержинскому району','ОСП по Идринскому району','ОСП по Иланскому району','ОСП по Ленинскому району г.Красноярска','ОСП по Мотыгинскому району','ОСП по Свердловскому району г.Красноярска','ОСП по Сухобузимскому району','ОСП по Тасеевскому району']},
    {subj:'Белгородская область',osps:['Волоконовский РОСП','Красногвардейский РОСП','Щебекинский РОСП']}
  ]},
  {num:13,code:'489',name:'Реншлер Вера',subjects:[
    {subj:'Рязанская область',osps:[]},
    {subj:'Мурманская область',osps:[]},
    {subj:'Республика Ингушетия',osps:[]},
    {subj:'Красноярский край',osps:['ОСП по Октябрьскому району г.Красноярска','ОСП по г.Железногорску','ОСП по району Талнах г.Норильска','ОСП по Ужурскому району','ОСП по Саянскому району','ОСП по Уярскому и Партизанскому районам','ОСП по Ермаковскому району','ОСП по Туруханскому району','ОСП по Пировскому району','ОСП Байкитский Эвенкийского района']},
    {subj:'Белгородская область',osps:['Губкинский РОСП','Ивнянский РОСП','Корочанский РОСП','Прохоровский РОСП']}
  ]},
  {num:15,code:'478',name:'Пятков Сергей',subjects:[
    {subj:'Ульяновская область',osps:[]},
    {subj:'Республика Калмыкия',osps:[]},
    {subj:'Астраханская область',osps:[]}
  ]},
  {num:16,code:'483',name:'Камаева Ирина',subjects:[
    {subj:'Республика Саха (Якутия)',osps:[]},
    {subj:'Калужская область',osps:[]},
    {subj:'Красноярский край',osps:['МОСП по г.Норильску','ОСП по г.Зеленогорску','ОСП по Емельяновскому району','ОСП по Енисейскому району','ОСП по Ирбейскому району','ОСП по Кировскому району г.Красноярска','ОСП по Нижнеингашскому району','ОСП по Таймырскому району','Хатангский ОСП по Таймырскому Долгано-Ненецкому району']}
  ]},
  {num:17,code:'612',name:'Сергеева Альбина',subjects:[
    {subj:'ХМАО',osps:[]},
    {subj:'Красноярский край',osps:['ОСП по г.Лесосибирску','ОСП №1 по Советскому району г.Красноярска','МОСП по г.Шарыпово и Шарыповскому району','ОСП №3 по Советскому району г.Красноярска','ОСП по Центральному району г.Красноярска','ОСП по Северо-Енисейскому району','ОСП по Казачинскому району','ОСП по Большемуртинскому району','ОСП по Манскому району','ОСП по Тюхтетскому району','ОСП по Краснотуранскому району']}
  ]},
  {num:18,code:'420',name:'Щербакова Настя',subjects:[
    {subj:'Ленинградская область',osps:[]},
    {subj:'Республика Бурятия',osps:[]},
    {subj:'Ставропольский край',osps:[]},
    {subj:'Запорожская область',osps:[]}
  ]},
    {num:19,code:'401',name:'Татьяна Савватеева',subjects:[
    {subj:'Курская область',osps:[]},
    {subj:'Самарская область',osps:['ОСП по Большечерниговскому району','ОСП по Волжскому району','ОСП по Железнодорожному району г.Самары','ОСП по Кировскому району','ОСП по Красноармейскому району']},
    {subj:'Краснодарский край',osps:['ОСП по Западному округу г.Краснодара','ОСП по Кавказскому району и г.Кропоткину','ОСП по Карасунскому округу г.Краснодара','ОСП по Крыловскому и Павловскому районам','ОСП по Прикубанскому округу','ОСП по Центральному округу г.Краснодара']},
    {subj:'Саратовская область',osps:['Ершовский РОСП','Заводской РОСП']},
    {subj:'Хабаровский край',osps:['ОСП по Амурскому району','ОСП по Ванинскому району','ОСП по Верхнебуреинскому району','ОСП по Вяземскому району','ОСП по Кировскому району']}
  ]},
  {num:20,code:'456',name:'Эля Муллагалиева',subjects:[
    {subj:'Ярославская область',osps:[]},
    {subj:'Хабаровский край',osps:['ОСП по Индустриальному району г.Хабаровска','ОСП по Комсомольскому району','ОСП по Краснофлотскому району','ОСП по Ленинскому району','ОСП по Нанайскому району','ОСП по Николаевскому району','ОСП по Октябрьскому району','ОСП по Облученскому району']},
    {subj:'Пензенская область',osps:['Белинский РОСП','Бессоновский РОСП','Железнодорожный РОСП г.Пензы','Зареченский ГОСП','Кузнецкий МОСП','Лунинский РОСП','Неверкинский РОСП','Октябрьский РОСП г.Пензы','ОСП по Нижнеломовскому и Наровчатскому районам','Первомайский РОСП г.Пензы']}
  ]},
  {num:21,code:'459',name:'Альбина Пономарева',subjects:[
    {subj:'Архангельская область',osps:[]},
    {subj:'Вологодская область',osps:[]},
    {subj:'Сахалинская область',osps:[]},
    {subj:'Саратовская область',osps:['Энгельский РОСП']}
  ]},
  {num:22,code:'812',name:'Виктория Свиридова',subjects:[
    {subj:'Брянская область',osps:[]},
    {subj:'Еврейская АО',osps:[]},
    {subj:'Пензенская область',osps:['Иссинский РОСП','Городнищенский РОСП','Каменский РОСП','Кондольский РОСП','Ленинский РОСП Пензы','Мокшанский РОСП','Никольский РОСП','ОСП по Башмаковскому и Пачелмскому районам','ОСП по Бековскому и Тамалинскому районам','ОСП по Колышлейскому и Малосердобинскому районам','ОСП по Лопатинскому и Камешкирскому районам','Пензенский РОСП','Спасский РОСП','Сердобский РОСП','Сосновоборский РОСП','Шемышейский РОСП']},
    {subj:'Хабаровский край',osps:['ОСП по Аяно-Майскому району','ОСП по Бикинскому району','ОСП по Биробиджанскому району','ОСП по г.Комсомольск-на-Амуре №1','ОСП по г.Комсомольск-на-Амуре №2','ОСП по Железнодорожному району г.Хабаровска','ОСП по Охотскому району','ОСП по району им.Лазо','ОСП по Смидовичскому району','ОСП по Советско-Гаванскому району','ОСП по Солнечному району','ОСП по Ульчскому району','ОСП по Хабаровскому району','ОСП по Центральному району г.Хабаровска']},
    {subj:'Краснодарский край',osps:['Абинский РОСП','Адлерский РОСП','Анапский ГОСП','Армавирский ГОСП','Геленджикский ГОСП','Гулькевичский РОСП','Ейский РОСП']}
  ]},
  {num:23,code:'471',name:'Тамара Матвиенко',subjects:[
    {subj:'Амурская область',osps:[]},
    {subj:'Калининградская область',osps:[]},
    {subj:'Саратовская область',osps:['Балаковский РОСП']},
    {subj:'Краснодарский край',osps:['Новороссийский ГОСП','Темрюкский РОСП','Лабинский ГОСП','Лазаревский РОСП г.Сочи','Ленинградский РОСП','Мостовский РОСП','Новокубанский РОСП','ОСП по Белоглинскому и Новопокровскому районам','ОСП по г.Кореновску и Выселковскому району','ОСП по г.Тимашевску и Брюховецкому району']}
  ]},
  {num:24,code:'472',name:'Елена Клопот',subjects:[
    {subj:'Забайкальский край',osps:[]},
    {subj:'Краснодарский край',osps:['Крымский РОСП','ОСП по Щербиновскому и Староминскому районам','Отрадненский РОСП','Приморско-Ахтарский РОСП','Северский РОСП','Славянский ГОСП г.Славянск-на-Кубани','Тбилисский РОСП','Тихорецкий РОСП','Туапсинский РОСП','Успенский РОСП','Усть-Лабинский РОСП','Хостинский ОСП г.Сочи','Центральный РОСП г.Сочи']},
    {subj:'Челябинская область',osps:['РОСП Металлургического р-на','РОСП Курчатовского р-на','Ашинский ГОСП','Красноармейский РОСП','Миасский ГОСП','ОСП по г.Златоусту и Кусинскому р-ну','ОСП по Катав-Ивановскому р-ну и г.Усть-Катаву','Саткинский ГОСП','Трехгорный ГОСП','Уйский РОСП','Чебаркульский ГОСП']}
  ]},
  {num:25,code:'479',name:'Лиля Абдуллина',subjects:[
    {subj:'Тюменская область',osps:[]},
    {subj:'Ивановская область',osps:[]}
  ]},
  {num:26,code:'403',name:'Светлана Сидоренко',subjects:[
    {subj:'Алтайский край',osps:[]}
  ]},
  {num:27,code:'815',name:'Иванова Надежда',subjects:[
    {subj:'Чувашская Республика',osps:[]},
    {subj:'Республика Коми',osps:[]},
    {subj:'Самарская область',osps:['ОСП по г.Новокуйбышевску','ОСП по г.Чапаевску','ОСП №1 по г.Сызрани','ОСП по Красноярскому району с.Красный Яр','ОСП по Куйбышевскому району г.Самары','ОСП по Ленинскому району г.Самары','ОСП по Октябрьскому району г.Самары','ОСП №2 по г.Сызрани','ОСП по Хворостянскому району','ОСП по Промышленному району','ОСП по Самарскому району г.Самары','ОСП по Сергиевскому району','ОСП по Советскому району','ОСП по Центральному району г.Тольятти','ОСП по Комсомольскому району г.Тольятти','ОСП по Автозаводскому району №1 г.Тольятти','ОСП по Автозаводскому району №2 г.Тольятти']},
    {subj:'Херсонская область',osps:[]}
  ]},
  {num:28,code:'466',name:'Шайхова Лия',subjects:[
    {subj:'Владимирская область',osps:[]},
    {subj:'Республика Марий Эл',osps:[]},
    {subj:'Оренбургская область',osps:[]}
  ]},
  {num:29,code:'431',name:'Гусев Борис',subjects:[
    {subj:'Липецкая область',osps:[]},
    {subj:'Костромская область',osps:[]}
  ]},
  {num:30,code:'432',name:'Товмасян Андраник',subjects:[
    {subj:'Чеченская Республика',osps:[]},
    {subj:'Волгоградская область',osps:[]}
  ]},
  {num:31,code:'437',name:'Валеев Денис',subjects:[
    {subj:'Воронежская область',osps:['Аннинский РОСП','Богучарский РОСП','Грибановский РОСП','Бутурлиновский РОСП','Верхнехавский РОСП','Верхнемамонский РОСП','Железнодорожный РОСП','Коминтерновский РОСП','Левобережный РОСП','Новоусманский РОСП','ОСП по г.Нововоронежу и Каширскому району','Рамонский РОСП','Семилукский РОСП','Советский РОСП г.Воронежа','Таловский РОСП','Центральный РОСП','Контемировский РОСП','Павловский РОСП','Бобровский РОСП']},
    {subj:'Кабардино-Балкарская Республика',osps:[]},
    {subj:'Магаданская область',osps:[]}
  ]},
  {num:32,code:'446',name:'Уткин Сергей',subjects:[
    {subj:'Саратовская область',osps:['Александрово-Гайский РОСП','Волжский РОСП г.Саратова','Вольский МОСП','Воскресенский РОСП','Дергачевский РОСП','Духовницкий РОСП','Калининский РОСП','Кировский РОСП г.Саратова','Красноармейский РОСП','Краснокутский РОСП','Краснопартизанский РОСП','Ленинский РОСП №1 г.Саратова','Ленинский РОСП №2 г.Саратова','Новобурасский РОСП','Лысогорский РОСП']},
    {subj:'Республика Алтай',osps:['Кош-Агачский РОСП УФССП по Республике Алтай (649780, Кош-Агачский р-н, с.Кош-Агач, ул.Чуйская, д.8)','Онгудайский РОСП УФССП по Республике Алтай (649440, с.Онгудай, ул.Советская, д.164)','ОСП по г.Горно-Алтайск УФССП по Республике Алтай (649007, г.Горно-Алтайск, ул.Ленина, д.226)','ОСП по Майминскому и Чойскому районам УФССП по Республике Алтай (649100, Майминский р-н, с.Майма, ул.Заводская, д.52)','Улаганский РОСП УФССП по Республике Алтай (649750, с.Улаган, ул.Кокышева, д.7)','Усть-Канский РОСП УФССП по Республике Алтай (649450, с.Усть-Кан, ул.Первомайская, д.1Б)','Усть-Коксинский РОСП УФССП по Республике Алтай (649490, с.Усть-Кокса, ул.Харитошкина, д.13)','Чемальский РОСП УФССП по Республике Алтай (649240, с.Чемал, ул.Пчелкина, д.74, корп.А)']},
    {subj:'Республика Адыгея',osps:[]}
  ]},
  {num:33,code:'436',name:'Латышов Роман',subjects:[
    {subj:'г. Севастополь',osps:[]},
    {subj:'Челябинская область',osps:['Аргаяшский РОСП','Еманжелинский РОСП','Еткульский РОСП','Каслинский РОСП','Коркинский ГОСП','ОСП по г.Кыштыму и г.Карабашу','Озерский ГОСП','Октябрьский РОСП','Пластовский ГОСП','Снежинский ГОСП','Троицкий ГОСП','Южноуральский ГОСП','Увельский ГОСП','Верхнеуфалейский РОСП','Брединский РОСП','Верхнеуральский РОСП','Кизильский РОСП','Нязепетровский РОСП','Чесменский РОСП']},
    {subj:'Саратовская область',osps:['Петровский РОСП','ОСП по Пугачевскому и Ивантеевскому районам','Марксовский РОСП','Новоузенский РОСП','Озинское РОСП','Октябрьский РОСП г.Саратова','ОСП по Аткарскому и Екатериновскому районам','ОСП по Базарно-Карабулакскому и Балтайскому районам','ОСП по Балашовскому, Романовскому и Самойловскому районам','ОСП по Ртищевскому, Аркадакскому и Турковскому районам','Питерский РОСП','Перелюбский РОСП','Ровенский РОСП','Саратовский РОСП','Советский РОСП','Татищевский РОСП','Федоровский РОСП','Фрунзенский РОСП г.Саратова','Хвалынский РОСП']},
    {subj:'Республика Северная Осетия-Алания',osps:[]}
  ]},
  {num:34,code:'439',name:'Фамильцева Наталья',subjects:[
    {subj:'Республика Дагестан',osps:[]},
    {subj:'Белгородская область',osps:['ОСП по г.Белгород']},
    {subj:'г. Санкт-Петербург',osps:['Левобережный ОСП Невского района','Московский РОСП УФССП','МРОСП по ИОИП УФССП Питер','ОСП по Кронштадскому и Курортному районам','ОСП по Центральному району','Петроградский РОСП УФССП','Петродворцовый РОСП УФССП','Полюстровский ОСП Красногвардейского района','Правобережный ОСП Невского района УФССП','Пушкинский РОСП УФССП','Фрунзенский РОСП УФССП']},
    {subj:'Ямало-Ненецкий АО',osps:['ОСП по г.Губкинскому УФССП по ЯНАО (629830, г.Губкинский, мкр.7, д.1)','ОСП по г.Лабытнанги УФССП России по ЯНАО (629400, г.Лабытнанги, ул.Школьная, д.8)','ОСП по г.Муравленко УФССП по ЯНАО (629602, г.Муравленко, ул.Губкина, д.15)','ОСП по г.Надыму и Надымскому району УФССП по ЯНАО (629730, г.Надым, ул.Зверева, д.3/2)','ОСП по г.Новый Уренгой УФССП по ЯНАО (629300, г.Новый Уренгой, ул.26 Съезда КПСС, д.4)','ОСП по г.Ноябрьску УФССП по ЯНАО (629800, г.Ноябрьск, ул.Мира пр-кт, д.55/57)','ОСП по Красноселькупскому району УФССП России по ЯНАО (629380, п.Красноселькуп, ул.Геологоразведчиков, д.7)','ОСП по Пуровскому району УФССП по ЯНАО (629850, г.Тарко-Сале, ул.Е.Колесниковой, д.7, кор.3)','ОСП по Тазовскому району УФССП по ЯНАО (629350, п.Тазовский, мкр.Геологов, д.8)']}
  ]},
  {num:35,code:'433',name:'Литвиненко Татьяна',subjects:[
    {subj:'Самарская область',osps:['ОСП по Больше-Глушицкому району','ОСП по Борскому и Богатовскому районам','ОСП по Исаклинскому и Шенталинскому районам','ОСП по Кинель-Черкасского района УФССП','ОСП по Кошкинскому району УФССП','ОСП по Нефтегорскому и Алексеевскому районам','ОСП по Пестравскому району','ОСП по Приволжскому району','ОСП по Безенчукскому району','ОСП по г.Отрадному','ОСП по г.Похвистнево','ОСП по Елховскому району','ОСП по Кинельскому району','ОСП по Челно-Вершинскому району','ОСП по г.Жигулевску','ОСП по г.Октябрьску','ОСП по Ставропольскому району УФССП России по Самарской области','ОСП по Камышлинскому и Клявлинскому району','ОСП по Красноглинскому району','ОСП по Шигонскому району']},
    {subj:'Республика Хакасия',osps:[]}
  ]},
  {num:36,code:'447',name:'Пелих Евгений',subjects:[
    {subj:'Орловская область',osps:[]},
    {subj:'Карачаево-Черкесская Республика',osps:[]}
  ]},
  {num:37,code:'807',name:'Окунькова Татьяна',subjects:[
    {subj:'Московская область',osps:['Волоколамский, Лотошинский и Шаховской РОСП УФССП по Московской области','Дмитровский РОСП УФССП по Московской области','Истринский РОСП УФССП по Московской области','Клинский РОСП УФССП по Московской области','Луховицкий РОСП УФССП по Московской области','Наро-Фоминский РОСП УФССП по Московской области','Орехово-Зуевский РОСП УФССП по Московской области','ОСП по Балашихинскому району и г.Железнодорожный по Московской области','Павлово-Посадский РОСП УФССП России по Московской области','Пушкинский РОСП УФССП России по Московской области','Серебряно-Прудский РОСП УФССП по Московской области','Серпуховский РОСП УФССП по Московской области','Люберецкий РОСП УФССП по Московской области','Можайский РОСП УФССП по Московской области','Ногинский РОСП УФССП по Московской области','Одинцовский РОСП УФССП по Московской области']},
    {subj:'Иркутская область',osps:['ОСП по Бодайбинскому и Мамско-Чуйскому районам УФССП по Иркутской области','ОСП по Жигаловскому и Качугскому районам УФССП по Иркутской области','ОСП по Заларинскому, Балаганскому и Нукутскому районам УФССП по Иркутской области','ОСП по Осинскому, Боханскому и Усть-Удинскому районам УФССП по Иркутской области','ОСП по Черемховскому и Аларскому районам УФССП по Иркутской области','ОСП по Эхирит-Булагатскому, Баяндаевскому и Ольхонскому р-нам УФССП по Иркутской области']}
  ]},
  {num:38,code:'463',name:'Сокова Соня',subjects:[
    {subj:'Нижегородская область',osps:[]}
  ]}
];

// ══════════════════════════════════════════════
//  НАСТРОЙКИ
// ══════════════════════════════════════════════
const ACCENT_COLORS = [
  {name:'Золото',    val:'#f5c518'},
  {name:'Синий',     val:'#2980b9'},
  {name:'Зелёный',   val:'#27ae60'},
  {name:'Красный',   val:'#e74c3c'},
  {name:'Фиолет.',   val:'#8e44ad'},
  {name:'Оранжев.',  val:'#e67e22'},
  {name:'Бирюзов.',  val:'#16a085'},
  {name:'Розовый',   val:'#e91e8c'},
  {name:'Серебро',   val:'#bdc3c7'},
  {name:'Лайм',      val:'#a8e063'}
];
const HIGHLIGHT_COLORS = [
  {name:'Жёлтый',   val:'#ffe066',text:'#1a1a1a'},
  {name:'Зелёный',  val:'#b8e8c9',text:'#155724'},
  {name:'Оранжев.', val:'#ffe5b4',text:'#7a3800'},
  {name:'Голубой',  val:'#b8d8ff',text:'#003580'},
  {name:'Розовый',  val:'#ffd6e0',text:'#7a003a'},
  {name:'Лаванда',  val:'#e8d5ff',text:'#3a006b'},
  {name:'Мята',     val:'#ccf2e8',text:'#004d34'},
  {name:'Персик',   val:'#ffe4cc',text:'#7a3800'},
  {name:'Лимон',    val:'#f7ffa0',text:'#3a3a00'},
  {name:'Белый',    val:'#ffffff', text:'#1a1a1a'}
];
const FONTS = [
  {name:'Segoe UI',  val:"'Segoe UI', Arial, sans-serif"},
  {name:'Arial',     val:'Arial, sans-serif'},
  {name:'Georgia',   val:'Georgia, serif'},
  {name:'Tahoma',    val:'Tahoma, sans-serif'},
  {name:'Verdana',   val:'Verdana, sans-serif'},
  {name:'Courier',   val:"'Courier New', monospace"},
  {name:'Palatino',  val:"'Palatino Linotype', serif"},
  {name:'Trebuchet', val:"'Trebuchet MS', sans-serif"},
  {name:'Impact',    val:'Impact, sans-serif'},
  {name:'Comic',     val:"'Comic Sans MS', cursive"}
];

let DB = [];
let searchType = 'subj';
let isLightTheme = false;
let rabbitTimer = null;
let currentSettings = {
  accent: '#f5c518',
  font: "'Segoe UI', Arial, sans-serif",
  fontSize: 18,
  highlight: '#ffe066',
  highlightText: '#1a1a1a'
};
let selectedAccent = null;
let selectedFont = null;
let selectedHighlight = null;

// ══════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════
function init() {
  const saved = localStorage.getItem('disl_db');
  if (saved) { try { DB = JSON.parse(saved); } catch(e) { DB = clone(RAW_DATA); } }
  else DB = clone(RAW_DATA);

  const savedTheme = localStorage.getItem('disl_theme');
  if (savedTheme === 'light') { document.documentElement.classList.add('light-theme'); isLightTheme = true; document.getElementById('theme-label').textContent = '🌙 Тёмная'; }

  const savedSettings = localStorage.getItem('disl_settings');
  if (savedSettings) { try { currentSettings = JSON.parse(savedSettings); } catch(e){} }
  applySettingsToDom(currentSettings);

  buildColorGrids();
  renderAll();
  updateStats();
  adjustContentTop();
}

function clone(d) { return JSON.parse(JSON.stringify(d)); }

function adjustContentTop() {
  const h = document.getElementById('fixed-top').offsetHeight;
  document.getElementById('content').style.marginTop = (h + 8) + 'px';
}
window.addEventListener('resize', adjustContentTop);

// ══════════════════════════════════════════════
//  APPLY SETTINGS
// ══════════════════════════════════════════════
function applySettingsToDom(s) {
  const root = document.documentElement;
  root.style.setProperty('--accent', s.accent);
  root.style.setProperty('--operator-border', s.accent);
  root.style.setProperty('--font', s.font);
  root.style.setProperty('--fs-base', s.fontSize + 'px');
  root.style.setProperty('--fs-sm',   (s.fontSize - 2) + 'px');
  root.style.setProperty('--fs-lg',   (s.fontSize + 3) + 'px');
  root.style.setProperty('--fs-xl',   (s.fontSize + 8) + 'px');
  root.style.setProperty('--highlight-bg', s.highlight);
  root.style.setProperty('--highlight-color', s.highlightText);
  document.body.style.fontFamily = s.font;
}

function buildColorGrids() {
  // Accent colors
  const ag = document.getElementById('color-grid');
  ag.innerHTML = '';
  ACCENT_COLORS.forEach(c => {
    const sw = document.createElement('div');
    sw.className = 'color-swatch' + (c.val === currentSettings.accent ? ' selected' : '');
    sw.style.background = c.val;
    sw.title = c.name;
    sw.onclick = () => { selectedAccent = c.val; document.querySelectorAll('#color-grid .color-swatch').forEach(x=>x.classList.remove('selected')); sw.classList.add('selected'); };
    ag.appendChild(sw);
  });
  // Highlight colors
  const hg = document.getElementById('highlight-grid');
  hg.innerHTML = '';
  HIGHLIGHT_COLORS.forEach(c => {
    const sw = document.createElement('div');
    sw.className = 'color-swatch' + (c.val === currentSettings.highlight ? ' selected' : '');
    sw.style.background = c.val;
    sw.style.border = '3px solid #888';
    sw.title = c.name;
    sw.onclick = () => { selectedHighlight = c; document.querySelectorAll('#highlight-grid .color-swatch').forEach(x=>x.classList.remove('selected')); sw.classList.add('selected'); };
    hg.appendChild(sw);
  });
  // Fonts
  const fg = document.getElementById('font-grid');
  fg.innerHTML = '';
  FONTS.forEach(f => {
    const btn = document.createElement('button');
    btn.className = 'font-btn' + (f.val === currentSettings.font ? ' selected' : '');
    btn.textContent = f.name;
    btn.style.fontFamily = f.val;
    btn.onclick = () => { selectedFont = f.val; document.querySelectorAll('.font-btn').forEach(x=>x.classList.remove('selected')); btn.classList.add('selected'); };
    fg.appendChild(btn);
  });
  // Font size range
  document.getElementById('font-size-range').value = currentSettings.fontSize;
  document.getElementById('font-size-label').textContent = currentSettings.fontSize + 'px';
}

function previewFontSize(val) {
  document.getElementById('font-size-label').textContent = val + 'px';
}

function applySettingsPreview() {
  const preview = {
    accent: selectedAccent || currentSettings.accent,
    font: selectedFont || currentSettings.font,
    fontSize: parseInt(document.getElementById('font-size-range').value),
    highlight: selectedHighlight ? selectedHighlight.val : currentSettings.highlight,
    highlightText: selectedHighlight ? selectedHighlight.text : currentSettings.highlightText
  };
  applySettingsToDom(preview);
}

function saveSettings() {
  const newSettings = {
    accent: selectedAccent || currentSettings.accent,
    font: selectedFont || currentSettings.font,
    fontSize: parseInt(document.getElementById('font-size-range').value),
    highlight: selectedHighlight ? selectedHighlight.val : currentSettings.highlight,
    highlightText: selectedHighlight ? selectedHighlight.text : currentSettings.highlightText
  };
  currentSettings = newSettings;
  applySettingsToDom(currentSettings);
  localStorage.setItem('disl_settings', JSON.stringify(currentSettings));
  closeSettings();
  alert('✅ Настройки сохранены!\n\nДля постоянного сохранения в HTML-файле:\nнажмите Ctrl+S в браузере (или «Сохранить страницу как...»).');
}

function openSettings() {
  selectedAccent = null; selectedFont = null; selectedHighlight = null;
  buildColorGrids();
  document.getElementById('settings-modal').classList.remove('hidden');
}
function closeSettings() { document.getElementById('settings-modal').classList.add('hidden'); }

// ══════════════════════════════════════════════
//  RENDER
// ══════════════════════════════════════════════
function renderAll(highlighted) {
  const c = document.getElementById('blocks-container');
  c.innerHTML = '';
  DB.forEach(op => c.appendChild(buildOpBlock(op, highlighted)));
  updateStats();
}

function buildOpBlock(op, highlighted) {
  const div = document.createElement('div');
  div.className = 'op-block';
  div.id = 'op-' + op.code;

  let isHit = false;
  if (highlighted) {
    const q = highlighted.q.toLowerCase();
    const t = highlighted.type;
    op.subjects.forEach(s => {
      if ((t==='subj'||t==='any') && s.subj.toLowerCase().includes(q)) isHit = true;
      if ((t==='osp' ||t==='any') && s.osps.some(o => o.toLowerCase().includes(q))) isHit = true;
    });
    if (t==='any' && (op.code.includes(q) || op.name.toLowerCase().includes(q))) isHit = true;
  }
  if (isHit) div.classList.add('highlighted');

  const toggleSpan = document.createElement('span');
  toggleSpan.className = 'op-toggle';
  toggleSpan.textContent = '▾';

  const header = document.createElement('div');
  header.className = 'op-header';
  header.innerHTML = `<span class="op-num">${op.num}.</span><span class="op-code">ОСП ${op.code}</span>${op.name ? `<span class="op-name">( ${op.name} )</span>` : ''}`;
  header.appendChild(toggleSpan);
  div.appendChild(header);

  const body = document.createElement('div');
  body.className = 'op-body collapsed';

  op.subjects.forEach(s => {
    const sb = document.createElement('div');
    sb.className = 'subj-block';
    const st = document.createElement('div');
    st.className = 'subj-title';
    st.textContent = s.subj;
    if (highlighted) {
      const q = highlighted.q.toLowerCase(), t = highlighted.type;
      if ((t==='subj'||t==='any') && s.subj.toLowerCase().includes(q)) st.classList.add('highlighted');
    }
    sb.appendChild(st);
    if (s.osps && s.osps.length > 0) {
      const ul = document.createElement('ul');
      ul.className = 'osp-list';
      s.osps.forEach(o => {
        const li = document.createElement('li');
        li.className = 'osp-item';
        li.textContent = o;
        if (highlighted) {
          const q = highlighted.q.toLowerCase(), t = highlighted.type;
          if ((t==='osp'||t==='any') && o.toLowerCase().includes(q)) li.classList.add('highlighted');
        }
        ul.appendChild(li);
      });
      sb.appendChild(ul);
    }
    body.appendChild(sb);
  });
  div.appendChild(body);

  header.addEventListener('click', () => {
    const collapsed = body.classList.toggle('collapsed');
    toggleSpan.classList.toggle('open', !collapsed);
  });
  if (isHit) { body.classList.remove('collapsed'); toggleSpan.classList.add('open'); }
  return div;
}

function expandAll()  { document.querySelectorAll('.op-body').forEach(b=>b.classList.remove('collapsed')); document.querySelectorAll('.op-toggle').forEach(t=>t.classList.add('open')); }
function collapseAll(){ document.querySelectorAll('.op-body').forEach(b=>b.classList.add('collapsed'));    document.querySelectorAll('.op-toggle').forEach(t=>t.classList.remove('open')); }

function updateStats() {
  let subj=0,osp=0;
  DB.forEach(op => { subj+=op.subjects.length; op.subjects.forEach(s=>osp+=s.osps.length); });
  document.getElementById('stat-ops').textContent  = DB.length;
  document.getElementById('stat-subj').textContent = subj;
  document.getElementById('stat-osp').textContent  = osp;
}

// ══════════════════════════════════════════════
//  SEARCH
// ══════════════════════════════════════════════
function openSearch() { document.getElementById('search-modal').classList.remove('hidden'); setTimeout(()=>document.getElementById('search-input').focus(),80); }
function closeSearch(){ document.getElementById('search-modal').classList.add('hidden'); }
function setSearchType(t) { searchType=t; ['subj','osp','any'].forEach(x=>document.getElementById('btn-type-'+x).classList.toggle('active',x===t)); }
function searchKeydown(e){ if(e.key==='Enter')doSearch(); if(e.key==='Escape')closeSearch(); }

function doSearch() {
  const q = document.getElementById('search-input').value.trim();
  if (!q) { resetSearch(); closeSearch(); return; }
  const highlighted = {q, type: searchType};
  renderAll(highlighted);

  const badges = [];
  DB.forEach(op => {
    let hit = false;
    op.subjects.forEach(s => {
      if ((searchType==='subj'||searchType==='any') && s.subj.toLowerCase().includes(q.toLowerCase())) hit=true;
      if ((searchType==='osp' ||searchType==='any') && s.osps.some(o=>o.toLowerCase().includes(q.toLowerCase()))) hit=true;
    });
    if (searchType==='any' && (op.code.includes(q)||op.name.toLowerCase().includes(q.toLowerCase()))) hit=true;
    if (hit) badges.push(op.code);
  });

  const rb = document.getElementById('result-badges');
  if (badges.length) {
    rb.innerHTML = badges.map(code => {
      const op = DB.find(x=>x.code===code);
      return `<button class="badge" onclick="openDetail('${code}')" title="Карточка оператора">ОСП ${code}${op&&op.name?' — '+op.name:''}</button>`;
    }).join('');
  } else {
    rb.innerHTML = '<span style="color:var(--text2);font-weight:400;font-size:var(--fs-base)">Ничего не найдено</span>';
  }
  closeSearch();
  setTimeout(() => { const first=document.querySelector('.op-block.highlighted'); if(first)first.scrollIntoView({behavior:'smooth',block:'start'}); }, 120);
  adjustContentTop();
}

function resetSearch() {
  document.getElementById('search-input').value='';
  document.getElementById('result-badges').innerHTML='';
  renderAll();
  adjustContentTop();
}

// ══════════════════════════════════════════════
//  DETAIL MODAL
// ══════════════════════════════════════════════
function openDetail(code) {
  const op = DB.find(x=>x.code===code);
  if (!op) return;
  document.getElementById('detail-title').textContent = `ОСП ${op.code}${op.name?' — '+op.name:''}`;
  const container = document.getElementById('detail-content');
  container.innerHTML = '';
  op.subjects.forEach(s => {
    const sb = document.createElement('div');
    sb.className = 'subj-block';
    const st = document.createElement('div');
    st.className = 'subj-title';
    st.textContent = '📍 ' + s.subj;
    sb.appendChild(st);
    if (s.osps && s.osps.length > 0) {
      const ul = document.createElement('ul');
      ul.className = 'osp-list';
      s.osps.forEach(o => { const li=document.createElement('li'); li.className='osp-item'; li.textContent=o; ul.appendChild(li); });
      sb.appendChild(ul);
    } else {
      const note = document.createElement('div');
      note.style.cssText='font-size:var(--fs-sm);color:var(--text2);padding:4px 8px;';
      note.textContent='— субъект без детализации ОСП';
      sb.appendChild(note);
    }
    container.appendChild(sb);
  });
  document.getElementById('detail-modal').classList.remove('hidden');
}
function closeDetail() { document.getElementById('detail-modal').classList.add('hidden'); }

// ══════════════════════════════════════════════
//  EXPORT EXCEL
// ══════════════════════════════════════════════
function exportExcel() {
  const rows=[['№','Код','ФИО оператора','Субъект РФ','ОСП/РОСП']];
  DB.forEach(op => {
    op.subjects.forEach(s => {
      if (!s.osps.length) rows.push([op.num,op.code,op.name,s.subj,'']);
      else s.osps.forEach(o=>rows.push([op.num,op.code,op.name,s.subj,o]));
    });
  });
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:5},{wch:8},{wch:24},{wch:32},{wch:60}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Дислокации');
  XLSX.writeFile(wb,'Дислокации_экспорт.xlsx');
}

// ══════════════════════════════════════════════
//  THEME
// ══════════════════════════════════════════════
function toggleTheme() {
  isLightTheme=!isLightTheme;
  document.documentElement.classList.toggle('light-theme',isLightTheme);
  document.getElementById('theme-label').textContent=isLightTheme?'🌙 Тёмная':'☀️ Светлая';
  localStorage.setItem('disl_theme',isLightTheme?'light':'dark');
}

// ══════════════════════════════════════════════
//  SCROLLBAR
// ══════════════════════════════════════════════
function updateScrollThumb() {
  const rail=document.getElementById('scroll-rail'),thumb=document.getElementById('scroll-thumb');
  if(!rail||!thumb)return;
  const railH=rail.offsetHeight,docH=document.documentElement.scrollHeight-window.innerHeight;
  const ratio=docH>0?window.scrollY/docH:0;
  const thumbH=Math.max(22,railH*(window.innerHeight/document.documentElement.scrollHeight));
  thumb.style.height=thumbH+'px'; thumb.style.top=(ratio*(railH-thumbH))+'px';
}
window.addEventListener('scroll',updateScrollThumb);
let dragging=false,dragStartY=0,dragStartScroll=0;
document.getElementById('scroll-thumb').addEventListener('mousedown',e=>{dragging=true;dragStartY=e.clientY;dragStartScroll=window.scrollY;e.preventDefault();});
document.addEventListener('mousemove',e=>{if(!dragging)return;const rail=document.getElementById('scroll-rail'),docH=document.documentElement.scrollHeight-window.innerHeight;window.scrollTo(0,dragStartScroll+(e.clientY-dragStartY)/rail.offsetHeight*docH);});
document.addEventListener('mouseup',()=>dragging=false);

// ══════════════════════════════════════════════
//  RABBIT (10 сек)
// ══════════════════════════════════════════════
function toggleRabbit(){ document.getElementById('rabbit-popup').classList.contains('open')?closeRabbit():openRabbit(); }
function openRabbit()  { document.getElementById('rabbit-popup').classList.add('open'); startRabbitTimer(); }
function closeRabbit() { document.getElementById('rabbit-popup').classList.remove('open'); clearInterval(rabbitTimer); document.getElementById('rabbit-timer').textContent=''; }
function startRabbitTimer() {
  clearInterval(rabbitTimer); let sec=10;
  document.getElementById('rabbit-timer').textContent=`Окно закроется через ${sec} сек...`;
  rabbitTimer=setInterval(()=>{ sec--; if(sec<=0){closeRabbit();return;} document.getElementById('rabbit-timer').textContent=`Окно закроется через ${sec} сек...`; },1000);
}

// ══════════════════════════════════════════════
//  KEYBOARD
// ══════════════════════════════════════════════
document.addEventListener('keydown',e=>{
  if(e.key==='F2'){e.preventDefault();openSearch();}
  if(e.key==='Escape'){closeSearch();closeDetail();closeSettings();}
});

['search-modal','detail-modal','settings-modal'].forEach(id=>{
  document.getElementById(id).addEventListener('click',e=>{ if(e.target.id===id) document.getElementById(id).classList.add('hidden'); });
});

window.addEventListener('DOMContentLoaded',()=>{ init(); setTimeout(updateScrollThumb,300); });
</script>
</body>
</html>
