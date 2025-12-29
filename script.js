/* v156: 2〜6列の幅差（余り/不足）を等分で配分する
   delta = (テーブルの表示幅 - 1列目 - 基準合計) / 5
   ただし、縮小しすぎて入力が潰れないように最小幅を考慮してクランプする。
*/
(function(){
  function px(n){ return Math.round(n*100)/100; }

  function updateEmployeeCols(){
    const table = document.querySelector('#settingsEmployees .employee-table-frame table.settings-table.employee-table-fixed');
    if(!table) return;

    const cs = getComputedStyle(table);

    // 基準幅
    const drag = parseFloat(cs.getPropertyValue('--emp-drag')) || 28;
    const bases = [
      parseFloat(cs.getPropertyValue('--emp-c2')) || 80,
      parseFloat(cs.getPropertyValue('--emp-c3')) || 80,
      parseFloat(cs.getPropertyValue('--emp-c4')) || 140,
      parseFloat(cs.getPropertyValue('--emp-c5')) || 80,
      parseFloat(cs.getPropertyValue('--emp-c6')) || 60,
    ];

    // 最小幅（実務的に入力UIが潰れない程度。必要なら調整してください）
    const mins = [60, 60, 100, 60, 60];

    // テーブル表示幅（border込みの見た目幅を採用）
    const rect = table.getBoundingClientRect();
    const tableW = rect.width;

    // deltaを等分計算
    const baseSum = bases.reduce((a,b)=>a+b,0);
    let diff = tableW - drag - baseSum;
    let delta = diff / bases.length;

    // ただし、マイナス方向で最小幅を割り込むなら、割り込まない範囲までクランプ
    // （全列同じdeltaを維持しつつ、限界を超えたらそこで止める）
    let minAllowedDelta = Infinity;
    for(let i=0;i<bases.length;i++){
      minAllowedDelta = Math.min(minAllowedDelta, mins[i] - bases[i]);
    }
    if(delta < minAllowedDelta) delta = minAllowedDelta;

    // デカすぎる増加も抑制（極端に広い画面で不自然になりすぎないように）
    // 上限を外したいならこの行をコメントアウト
    const maxAllowedDelta = 240; // 例：+240px/列まで
    if(delta > maxAllowedDelta) delta = maxAllowedDelta;

    table.style.setProperty('--emp-delta', px(delta) + 'px');
  }

  // 初期化＆リサイズ追従
  window.addEventListener('resize', updateEmployeeCols, {passive:true});
  // タブ表示直後など、レイアウト確定後にもう一回
  setTimeout(updateEmployeeCols, 0);
  setTimeout(updateEmployeeCols, 200);

  // 可能なら ResizeObserver でコンテナ幅変化にも追従
  if('ResizeObserver' in window){
    const ro = new ResizeObserver(()=>updateEmployeeCols());
    const container = document.querySelector('#settingsEmployees .employee-col-left');
    if(container) ro.observe(container);
    ro.observe(document.documentElement);
  }
})();

/* ===========
  このファイルは「構文エラーが出ない完全版」をsandbox上で生成したものです。
  以前のコピペで途切れていた場合、こちらに差し替えるだけで Unexpected end of input は解消します。

const STORAGE_KEY = "workScheduleApp_v2";
let appState=null;

let currentYearMonth=null;
let currentStoreId=null;
let currentSchedule=null;
let currentStoreSchedule=null;

let ctrlDown=false;
let isSupportMode=false;

let dragSelecting=false; // 範囲選択中（ドラッグ中）
let pointerDown=false;   // 左クリック押下中
let didDrag=false;       // 実際にセルをまたいでドラッグしたか
let dragJustEnded=false; // ドラッグ直後の click で選択が潰れるのを防ぐ
let dragAdditive=false;  // Ctrl/Metaドラッグ：既存選択に追加
let lastMouseDownAdditive=false; // 直近mousedownがCtrl/Metaだったか（focusイベント対策）
let ctrlTogglePending=false; // Ctrl/Meta時：クリックならトグル、ドラッグなら矩形（判定待ち）
let ctrlPendingCell=null;
let dragRangeMode='replace'; // 'replace' | 'add' | 'remove'
let dragRangeModeStart='replace'; // デバッグ用：開始時のモード
let dragStartX=0, dragStartY=0; // ドラッグ開始座標
let dragLastRectKeys=null; // Ctrlドラッグ中の矩形キー集合（確定用）
let shiftAnchorCell=null; // Shift+矢印 範囲選択の起点
let suppressShiftAnchorReset=false; // Shift範囲選択中の anchor リセット抑止

let dragStartCell=null;
let dragMoveCount=0;
let dragLastHit=null;
let dragEventLog=[]; // {t,type,buttons,button,code,key,target,td}
function logDragEvent(ev){
  dragEventLog.push(ev);
  if(dragEventLog.length>40) dragEventLog.shift();
}

let selectedCells=[];

let selectedKeySet=new Set();
let preserveSelectionOnRender=false;
let preserveSelectionKeys=null;
let preserveSelectionActiveKey=null;
let activeCell=null;


// ===== Clear selection / focus =====
function clearScheduleCellFocus(){
  try{
    const table = document.getElementById("scheduleTable");
    if(table){
      // 選択強調（cell-selected / cell-multi）と、ドラッグプレビューをまとめて解除
      table.querySelectorAll("td.cell-selected, td.cell-multi, td.cell-preview-add, td.cell-preview-remove")
        .forEach(td=>{
          td.classList.remove("cell-selected","cell-multi","cell-preview-add","cell-preview-remove");
        });
    }
  }catch(_e){}
  try{ selectedCells = []; }catch(_e){}
  try{ selectedKeySet = new Set(); }catch(_e){}
  try{ activeCell = null; }catch(_e){}
  try{ shiftAnchorCell = null; }catch(_e){}
  try{ dragLastRectKeys = null; }catch(_e){}
  try{
    // 念のため：フォーカスが残っていると見た目が残る環境があるので解除
    if(document.activeElement && document.activeElement.closest?.("#scheduleTable")){
      document.activeElement.blur();
    }
  }catch(_e){}
  try{
    if(typeof renderDebugInfo==="function") renderDebugInfo();
  }catch(_e){}
}


function isClickInsideScheduleOrKeypad(target){
  try{
    if(!target) return false;
    if(target.closest?.("#scheduleTable")) return true;
    if(target.closest?.("#keypad")) return true; // keypad clicks should not clear selection
    return false;
  }catch(_e){ return false; }
}

// click outside schedule => clear focus
document.addEventListener("mousedown",(e)=>{
  try{
    if((selectedKeySet && selectedKeySet.size>0) || activeCell){
      if(!isClickInsideScheduleOrKeypad(e.target)){
        clearScheduleCellFocus();
      }
    }
  }catch(_e){}
}, true);

// Esc => clear focus
window.addEventListener("keydown",(e)=>{
  try{
    if(e.key==="Escape"){
      clearScheduleCellFocus();
    }
  }catch(_e){}
}, true);

// ===== Keyboard Debug =====
const __KEY_DEBUG__ = { enabled:true, max:80, events:[] };
function safeCellKeyMaybe(cell){
  try{
    if(!cell) return null;
    const td = (cell.tagName==="TD")?cell:cell.closest?.("td");
    return td ? cellKey(td) : null;
  }catch(_){ return null; }
}
function dbgKey(stage, e, extra){
  try{
    if(!__KEY_DEBUG__.enabled) return;
    const entry = {
      t: Date.now(),
      stage,
      key: e?.key,
      code: e?.code,
      shift: !!e?.shiftKey,
      ctrl: !!e?.ctrlKey,
      meta: !!e?.metaKey,
      alt: !!e?.altKey,
      composing: !!e?.isComposing,
      defaultPrevented: !!e?.defaultPrevented,
      target: e?.target?.tagName || null,
      activeEl: document.activeElement?.tagName || null,
      activeCell: safeCellKeyMaybe((()=>{try{return activeCell;}catch(_){return null;}})()),
      anchor: safeCellKeyMaybe((()=>{try{return shiftAnchorCell;}catch(_){return null;}})()),
      extra: extra || null
    };
    __KEY_DEBUG__.events.push(entry);
    if(__KEY_DEBUG__.events.length > __KEY_DEBUG__.max){
      __KEY_DEBUG__.events.splice(0, __KEY_DEBUG__.events.length-__KEY_DEBUG__.max);
    }
    try{
      if(typeof dragEventLog!=="undefined" && Array.isArray(dragEventLog)){
        dragEventLog.push({t:Date.now(), type:"key", stage, key:e?.key, code:e?.code, shift:!!e?.shiftKey, ctrl:!!e?.ctrlKey, meta:!!e?.metaKey, alt:!!e?.altKey, composing:!!e?.isComposing, prevented:!!e?.defaultPrevented, extra: extra||null});
        if(dragEventLog.length>200) dragEventLog.splice(0, dragEventLog.length-200);
      }
    }catch(_){ }
    if(typeof renderDebugInfo==="function") renderDebugInfo();
  }catch(_){}
}
window.__KEY_DEBUG__ = __KEY_DEBUG__;
// capture all keydown for troubleshooting
window.addEventListener("keydown",(e)=>{ dbgKey("cap-all", e, null); }, true);

// ===== Settings DnD Debug (always visible in Console) =====
// 「何も表示されない」= dragstart/drop が発火していない可能性が高いため、
// ネイティブDnDイベントをグローバルに捕捉してコンソールに出す。
window.addEventListener("dragstart", (e)=>{
  try{ console.debug("[dnd] dragstart", {tag:e.target?.tagName, cls:e.target?.className, id:e.target?.id}); }catch(_){ }
}, true);
window.addEventListener("dragover", (e)=>{
  try{ console.debug("[dnd] dragover", {tag:e.target?.tagName, cls:e.target?.className, id:e.target?.id}); }catch(_){ }
}, true);
window.addEventListener("drop", (e)=>{
  try{ console.debug("[dnd] drop", {tag:e.target?.tagName, cls:e.target?.className, id:e.target?.id}); }catch(_){ }
}, true);

let dirty=false;
let pendingNavigationTarget=null;
let pendingFocus=null;

let imeAlertTimer=null;

const dom={};

document.addEventListener("DOMContentLoaded",()=>{
  cacheDom();
  initState();
  initHeader();
  initSidebar();
  initModals();
  initDebugTools();
  loadInitialContext();
  renderAll();
  markDropZones();
});

// v119: markDropZones() が未定義だと初期化で停止し、設定画面DnDが動かない。
// ここでは安全なno-opとして定義しておき、必要な初期化は各パネル描画時に行う。
function markDropZones(){
  try{
    // 右側パネルは refreshEmployeeOrderPanel() 内でdropzone生成＋DnD初期化される
    // ここは ReferenceError を避けるための保険。初期化関数があれば呼ぶ。
    if(typeof ensureEmployeeOrderDnDInitialized==="function") ensureEmployeeOrderDnDInitialized();
  }catch(_e){}
}

function cacheDom(){
  dom.storeSelect=document.getElementById("storeSelect");
  dom.monthInput=document.getElementById("monthInput");
  dom.autoSaveToggle=document.getElementById("autoSaveToggle");
  dom.saveButton=document.getElementById("saveButton");  dom.exportButton=document.getElementById("exportButton");
  dom.settingsButton=document.getElementById("settingsButton");
dom.supportStoreSelect=document.getElementById("supportStoreSelect");
  dom.modeBadge=document.getElementById("modeBadge");
  dom.keypad=document.getElementById("keypad");
  dom.imeAlert=document.getElementById("imeAlert");
  dom.scheduleTitle=document.getElementById("scheduleTitle");
  dom.scheduleInner=document.getElementById("scheduleInner");
  dom.schedulePlaceholder=document.getElementById("schedulePlaceholder");
  dom.alertList=document.getElementById("alertList");

  dom.settingsModalBackdrop=document.getElementById("settingsModalBackdrop");
  dom.settingsCloseButton=document.getElementById("settingsCloseButton");
  dom.settingsSaveButton=document.getElementById("settingsSaveButton");
  dom.settingsCancelButton=document.getElementById("settingsCancelButton");
  dom.storeTableBody=document.getElementById("storeTableBody");
  dom.employeeTableBody=document.getElementById("employeeTableBody");
  dom.employeeOrderStoreSelect=document.getElementById("employeeOrderStoreSelect");
  dom.employeeOrderList=document.getElementById("employeeOrderList");
  // 右ペイン：店舗セレクト切替で即再描画
  dom.employeeOrderStoreSelect?.addEventListener("change", ()=>{ refreshEmployeeOrderPanel(); });
  // 店舗タブの変更を従業員タブの所属店舗プルダウンに即反映（設定モーダル内で完結させる）
  let __storeChangeRaf = null;
  const __onStoreChange = ()=>{
    if(__storeChangeRaf) cancelAnimationFrame(__storeChangeRaf);
    __storeChangeRaf = requestAnimationFrame(()=>{refreshEmployeeStoreOptionsFromSettings(); refreshEmployeeOrderPanel();});
  };
  dom.storeTableBody.addEventListener("input", __onStoreChange);
  dom.storeTableBody.addEventListener("change", __onStoreChange);

  dom.workCodeTableBody=document.getElementById("workCodeTableBody");
  dom.leaveCodeTableBody=document.getElementById("leaveCodeTableBody");
  dom.holidayYearInput=document.getElementById("holidayYearInput");
  dom.holidayFetchButton=document.getElementById("holidayFetchButton");
  dom.holidayTableBody=document.getElementById("holidayTableBody");
  dom.addHolidayButton=document.getElementById("addHolidayButton");
  dom.addStoreButton=document.getElementById("addStoreButton");
  dom.addEmployeeButton=document.getElementById("addEmployeeButton");
  dom.employeeUnassignedOnlyToggle=document.getElementById("employeeUnassignedOnlyToggle");
  dom.employeeUnassignedHint=document.getElementById("employeeUnassignedHint");
  dom.employeeUnassignedOnlyToggle?.addEventListener("click", ()=>{
    dom.employeeUnassignedOnlyToggle.classList.toggle("on");
    applyEmployeeLeftFilter();
  });
  dom.addLeaveCodeButton=document.getElementById("addLeaveCodeButton");

  dom.unsavedModalBackdrop=document.getElementById("unsavedModalBackdrop");
  dom.unsavedDiscardButton=document.getElementById("unsavedDiscardButton");
  dom.unsavedSaveButton=document.getElementById("unsavedSaveButton");
  dom.unsavedCancelButton=document.getElementById("unsavedCancelButton");

  dom.newScheduleModalBackdrop=document.getElementById("newScheduleModalBackdrop");
  dom.newScheduleMessage=document.getElementById("newScheduleMessage");
  dom.newScheduleWithSettingsButton=document.getElementById("newScheduleWithSettingsButton");
  dom.newScheduleDirectButton=document.getElementById("newScheduleDirectButton");
  dom.newScheduleCancelButton=document.getElementById("newScheduleCancelButton");

  dom.debugButton=document.getElementById("debugButton");
  dom.debugSection=document.getElementById("debugSection");
  dom.debugInfo=document.getElementById("debugInfo");
  dom.debugDragLog=document.getElementById("debugDragLog");
  dom.dragHint=document.getElementById('dragHint');
  dom.dragHintText=document.getElementById('dragHintText');
  dom.dragHintPill=document.getElementById('dragHintPill');
  dom.dbgExport=document.getElementById("dbgExport");
  dom.dbgCopyState=document.getElementById("dbgCopyState");
  dom.dbgFixSnapshot=document.getElementById("dbgFixSnapshot");
  dom.dbgReset=document.getElementById("dbgReset");
  dom.dbgClickInspectToggle=document.getElementById("dbgClickInspectToggle");
  dom.debugClickInfo=document.getElementById("debugClickInfo");
}

function initState(){
  const saved=localStorage.getItem(STORAGE_KEY);
  if(saved){
    try{ appState=JSON.parse(saved); }
    catch(e){ console.error("保存データ読込失敗。初期化します。",e); appState=createDefaultState(); }
  }else{
    appState=createDefaultState();
  }
  if(!appState.pendingNightCarryovers) appState.pendingNightCarryovers={};
  // v68: 入力モードの既定（クリックで切替 + Ctrlで一時切替）
  if(!appState.settings) appState.settings={};
  if(!("inputModeBase" in appState.settings)) appState.settings.inputModeBase="work";
  persistState();
}

function createDefaultState(){
  const workCodes=[
    {code:"A",key:"7",color:"#e11d48",bgColor:lightenColor("#e11d48",0.86)},
    {code:"B",key:"8",color:"#1976d2",bgColor:lightenColor("#1976d2",0.86)},
    {code:"C",key:"9",color:"#f59e0b",bgColor:lightenColor("#f59e0b",0.88)},
    {code:"D",key:"4",color:"#16a34a",bgColor:lightenColor("#16a34a",0.88)},
    {code:"E",key:"5",color:"#ca8a04",bgColor:lightenColor("#ca8a04",0.90)},
    {code:"F",key:"6",color:"#7c3aed",bgColor:lightenColor("#7c3aed",0.86)},
  ];
  const leaveCodes=[
    {code:"特",key:"7"},{code:"調",key:"8"},{code:"公",key:"9"},{code:"年",key:"4"},{code:"病気",key:"5"},{code:"休",key:"6"},
  ];
  return {
    settings:{ autoSave:true, lastStoreId:null, lastYearMonth:null, inputModeBase:"work" },
    templateMaster:{
      nextStoreId:1,nextEmployeeId:1,
      stores:[],
      employees:[],
      workCodes, leaveCodes,
      holidays:{}
    },
    schedules:{},
    pendingNightCarryovers:{}
  };
}

function persistState(){ localStorage.setItem(STORAGE_KEY, JSON.stringify(appState)); }

/* ===== Header ===== */
function initHeader(){
  dom.monthInput.addEventListener("change", ()=>{
    const ym=dom.monthInput.value||null;
    if(!ym) return;
    navigateTo(currentStoreId, ym);
  });

  dom.storeSelect.addEventListener("change", ()=>{
    const storeId=Number(dom.storeSelect.value||"0")||null;
    navigateTo(storeId, currentYearMonth);
  });

  dom.autoSaveToggle.addEventListener("click", ()=>{
    const newVal=!appState.settings.autoSave;
    appState.settings.autoSave=newVal;
    updateAutoSaveToggleUI();
    if(newVal){
      persistState();
      dirty=false;
      updateSaveButton();
    }
  });

  dom.saveButton.addEventListener("click", ()=>{
    if(!dirty) return;
    persistState();
    dirty=false;
    updateSaveButton();
  });
  dom.exportButton.addEventListener("click", ()=>{
    const schedule=getCurrentSchedule();
    if(!schedule) return;
    openPrintWindowForMonth(schedule,getStoreIdsInSchedule(schedule));
  });

  dom.settingsButton.addEventListener("click", ()=>openSettingsModal(null));


  window.addEventListener("beforeunload",(e)=>{
    if(!appState.settings.autoSave && dirty){
      e.preventDefault();
      e.returnValue="";
    }
  });
}

/* ===== Sidebar / Keys ===== */
function initSidebar(){
  dom.supportStoreSelect.addEventListener("change", ()=>{
    const val=dom.supportStoreSelect.value;
    isSupportMode = val !== "self";
    renderKeypad();
  });

  
  
  // v68: サイドバー「コード種別」をクリックで切替（Ctrl長押しの一時切替と共存）
  try{
    dom.modeBadge?.setAttribute("role","button");
    dom.modeBadge?.setAttribute("tabindex","0");
    dom.modeBadge?.addEventListener("click",(e)=>{
      // Ctrl/Meta押下中は「一時切替」用なので、クリック切替は無効にして衝突を防ぐ
      if(e.ctrlKey || e.metaKey || ctrlDown) return;
      const cur=getBaseInputMode();
      const next=(cur==="work") ? "leave" : "work";
      appState.settings.inputModeBase=next;
      persistState();
      renderKeypad();
    });
    dom.modeBadge?.addEventListener("keydown",(e)=>{
      if(e.key!=="Enter" && e.key!==" ") return;
      e.preventDefault();
      dom.modeBadge.click();
    });
  }catch(_e){}

// WS_SHIFT_ARROW_WINDOW: Shift+矢印を最優先で捕捉（他のkeydownハンドラに奪われないよう window/capture で処理）
  window.addEventListener("keydown",(e)=>{
    dbgKey("wshift-enter", e, null);
    if(e.isComposing || e.key==="Process"){ dbgKey("wshift-exit", e, {reason:"composing"}); return; }
    if(!e.shiftKey || e.ctrlKey || e.metaKey){ dbgKey("wshift-exit", e, {reason:"modifiers"}); return; }
    if(!(e.key==="ArrowLeft"||e.key==="ArrowRight"||e.key==="ArrowUp"||e.key==="ArrowDown")){ dbgKey("wshift-exit", e, {reason:"not-arrow"}); return; }

    // activeCell が無い場合は activeElement から推測（contenteditable対策）
    if(!activeCell){
      const td = document.activeElement?.closest?.("td");
      if(td){ setActiveCell(td); }
    }
    if(!activeCell){ dbgKey("wshift-exit", e, {reason:"no-activeCell"}); return; }

    const activeTd = (activeCell.tagName==="TD") ? activeCell : activeCell.closest?.("td");
    if(activeTd?.dataset?.rowType==='note'){ dbgKey('wshift-exit', e, {reason:'note-row'}); return; }
    if(!activeTd){ dbgKey("wshift-exit", e, {reason:"active-not-td"}); return; }

    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();
    dbgKey("wshift-stop", e, null);

    if(!shiftAnchorCell) shiftAnchorCell = activeTd;
    const anchorTd = (shiftAnchorCell.tagName==="TD") ? shiftAnchorCell : shiftAnchorCell.closest?.("td");
    if(!anchorTd){ shiftAnchorCell = activeTd; }

    const rowType = (shiftAnchorCell.tagName==="TD" ? shiftAnchorCell : activeTd).dataset.rowType;
    const maxDay = getVisibleMaxDay();
    const rowList = getRowIndexList(rowType);

    const curRow = parseInt(activeTd.dataset.rowIndex,10);
    const curDay = parseInt(activeTd.dataset.day,10);
    let r = curRow;
    let d = curDay;

    if(e.key==="ArrowLeft")  d = Math.max(1, d-1);
    if(e.key==="ArrowRight") d = Math.min(maxDay, d+1);

    if(e.key==="ArrowUp"||e.key==="ArrowDown"){
      const idx = rowList.indexOf(curRow);
      if(idx !== -1){
        const nextIdx = e.key==="ArrowUp" ? Math.max(0, idx-1) : Math.min(rowList.length-1, idx+1);
        r = rowList[nextIdx];
      }else{
        r = rowList[0];
      }
    }

    const next = getCellByRowDay(rowType, r, d);
    dbgKey("wshift-move", e, {rowType, curRow, curDay, nextRow:r, nextDay:d, nextFound:!!next, rowListLen:rowList.length, maxDay});
    if(!next){ dbgKey("wshift-exit", e, {reason:"next-null"}); return; }

    const rectKeys = buildRectKeys(shiftAnchorCell, next);
    dbgKey("wshift-rect", e, {rectSize: rectKeys.size, anchor: cellKey(shiftAnchorCell), next: cellKey(next)});
    suppressShiftAnchorReset=true;
    pointerDown = false;
    applySelectionKeys(rectKeys, next);
    try{ next.focus({preventScroll:true}); }catch(_){ try{ next.focus(); }catch(__){} }
        suppressShiftAnchorReset=false;
dbgKey("wshift-done", e, null);
  }, true);

document.addEventListener("keyup",(e)=>{ if(e.key==="Shift"){ shiftAnchorCell=null; } });

// Ctrlキーで「業務↔休暇」パレット切替（押下中のみ休暇モード）
document.addEventListener("keydown",(e)=>{
  // 右Ctrl/左Ctrlどちらも
  if(e.key==="Control"){
    if(!ctrlDown){
      ctrlDown=true;
      dbgKey("ctrl-down", e, null);
      try{ renderKeypad(); }catch(_){}
    }
  }
}, true);

document.addEventListener("keyup",(e)=>{
  if(e.key==="Control"){
    if(ctrlDown){
      ctrlDown=false;
      dbgKey("ctrl-up", e, null);
      try{ renderKeypad(); }catch(_){}
    }
  }
}, true);

// ブラウザ外にフォーカスが移った場合など：取りこぼし防止
window.addEventListener("blur", ()=>{
  if(ctrlDown){
    ctrlDown=false;
    try{ renderKeypad(); }catch(_){}
  }
  if(typeof shiftAnchorCell!=="undefined"){ shiftAnchorCell=null; }
}, true);
document.addEventListener("keydown",(e)=>{
    dbgKey("shift-enter", e, null);
    if(e.isComposing || e.key==="Process"){ dbgKey("shift-exit", e, {reason:"composing"}); return; }
    if(!e.shiftKey || e.ctrlKey || e.metaKey){ dbgKey("shift-exit", e, {reason:"modifiers"}); return; }
    if(!(e.key==="ArrowLeft"||e.key==="ArrowRight"||e.key==="ArrowUp"||e.key==="ArrowDown")){ dbgKey("shift-exit", e, {reason:"not-arrow"}); return; }

    // NumLock OFF のテンキーは入力用途と衝突するため対象外
    if(typeof e.code==="string" && e.code.startsWith("Numpad")){ dbgKey("shift-exit", e, {reason:"numpad"}); return; }

    // activeCell が未設定の場合、フォーカス要素から推測（contenteditable対策）
    if(!activeCell){
      const td = document.activeElement?.closest?.("td");
      if(td){ activeCell = td; td.classList.add('cell-selected'); }
    }
    if(!activeCell){ dbgKey("shift-exit", e, {reason:"no-activeCell"}); return; }

    const activeTd = (activeCell.tagName==="TD") ? activeCell : activeCell.closest?.("td");
    if(!activeTd){ dbgKey("shift-exit", e, {reason:"active-not-td"}); return; }

    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();
    dbgKey("shift-stop", e, null);

    if(!shiftAnchorCell) shiftAnchorCell = activeTd;

    const baseTd = (shiftAnchorCell.tagName==="TD") ? shiftAnchorCell : shiftAnchorCell.closest?.("td");
    const rowType = baseTd.dataset.rowType;
    const maxDay = getVisibleMaxDay();

    const rowList = getRowIndexList(rowType);
    const curRow = parseInt(activeTd.dataset.rowIndex,10);
    const curDay = parseInt(activeTd.dataset.day,10);

    let r = curRow;
    let d = curDay;

    if(e.key==="ArrowLeft")  d = Math.max(1, d-1);
    if(e.key==="ArrowRight") d = Math.min(maxDay, d+1);

    if(e.key==="ArrowUp"||e.key==="ArrowDown"){
      const idx = rowList.indexOf(curRow);
      if(idx !== -1){
        const nextIdx = e.key==="ArrowUp" ? Math.max(0, idx-1) : Math.min(rowList.length-1, idx+1);
        r = rowList[nextIdx];
      }else{
        r = rowList[0];
      }
    }

    const next = getCellByRowDay(rowType, r, d);
    dbgKey("shift-move", e, {rowType, curRow, curDay, nextRow:r, nextDay:d, nextFound:!!next, rowListLen:rowList.length, maxDay});
    if(!next){ dbgKey("shift-exit", e, {reason:"next-null"}); return; }

    const rectKeys = buildRectKeys(baseTd, next);
    dbgKey("shift-rect", e, {rectSize: rectKeys.size, anchor: cellKey(baseTd), next: cellKey(next)});

    suppressShiftAnchorReset=true;
    pointerDown = false;
    applySelectionKeys(rectKeys, next);
    try{ next.focus({preventScroll:true}); }catch(_){ try{ next.focus(); }catch(__){} }

        suppressShiftAnchorReset=false;
dbgKey("shift-done", e, null);
  }, true);
renderKeypad();

  // ドラッグ範囲選択中にセル内テキストが選択されるのを防止
  document.addEventListener("selectstart", (e)=>{
    if(dragSelecting){ e.preventDefault(); }
  });

  // 範囲選択：安定版（mouse主体）。段ずれは開始段へスナップ。表外クリックで状態リセット。
  const resolveCellAtPoint = (x,y)=>{
    const table = document.getElementById("scheduleTable");
    if(!table) return null;
    const els = (document.elementsFromPoint ? document.elementsFromPoint(x, y) : []);
    for(const el of els){
      const td = el && el.closest ? el.closest('td[data-day][data-row-type][data-emp-id]') : null;
      if(td && table.contains(td)) return td;
    }
    const el = document.elementFromPoint(x, y);
    const td = el && el.closest ? el.closest('td[data-day][data-row-type][data-emp-id]') : null;
    return (td && table.contains(td)) ? td : null;
  };

  const startDragOnCell = (td, e, typeTag)=>{
    if(!td) return;
    if(td.dataset.rowType==="support") return;

    // 2段目(note)は複数セル選択/範囲選択を禁止：単一セル選択のみ
    if(td.dataset.rowType==="note"){
      try{ e.preventDefault(); }catch(_){}
      // クリック=単一選択。Ctrlでも加算/除外しない
      dragAdditive = false;
      lastMouseDownAdditive = false;
      ctrlTogglePending = false;
      ctrlPendingCell = null;
      dragRangeMode = "replace";
      dragRangeModeStart = "replace";
      dragBaseSelectionKeys = null;
      clearDragPreview();
      setActiveCell(td,false);
      logDragEvent({t:Date.now(),type:typeTag+"(note-single)",button:e.button,buttons:e.buttons,ctrl:!!(e.ctrlKey||e.metaKey),meta:!!e.metaKey,target:e.target?.tagName,td:{rt:td.dataset.rowType,emp:td.dataset.empId,day:td.dataset.day}});
      return;
    }


    dragAdditive = !!(e.ctrlKey || e.metaKey || ctrlDown);
    lastMouseDownAdditive = dragAdditive;

    // Ctrlドラッグのモード：開始セルが選択済みなら除外(remove)、未選択なら追加(add)
    const k = cellKey(td);
    dragRangeMode = dragAdditive ? (selectedKeySet.has(k) ? "remove" : "add") : "replace";
    dragRangeModeStart = dragRangeMode;

    // Ctrlドラッグでは、開始時点の選択をスナップショットとして保持（ドラッグ中のプレビュー用）
    dragBaseSelectionKeys = dragAdditive ? new Set(selectedKeySet) : null;

    pointerDown=true;
    didDrag=false;
    dragSelecting=false;
    dragStartCell=td;
    dragStartX=e.clientX; dragStartY=e.clientY;
    dragMoveCount=0;
    dragLastHit=null;

    try{ e.preventDefault(); }catch(_){}

    // Ctrl/Metaの場合はクリックかドラッグかを mouseup まで保留
    ctrlTogglePending = dragAdditive;
    ctrlPendingCell = dragAdditive ? td : null;

    // ctrl click preview (確定は mouseup)
    if(dragAdditive && dragRangeMode==='remove'){
      clearDragPreview();

    shiftAnchorCell=null;
      td.classList.add('cell-preview-remove');
      setDragHint('remove', 1);
    }else if(dragAdditive && dragRangeMode==='add'){
      clearDragPreview();
      // 追加候補は未選択セルのみ
      if(!selectedKeySet.has(cellKey(td))) td.classList.add('cell-preview-add');
      setDragHint('add', selectedKeySet.has(cellKey(td)) ? 0 : 1);
    }

    if(dragAdditive){
      // Ctrl操作では selection の増減は mouseup/drag確定時に行う。activeは維持。
      // ただし段違いのドラッグ開始を防ぐため、選択がある場合は段を合わせる
      if(selectedCells.length>0){
        const baseType = selectedCells[0].dataset.rowType;
        if(td.dataset.rowType !== baseType) return;
      }
    }else{
      // 通常クリックはアクティブを更新し、選択を置換
      setActiveCell(td,false);
    }

    logDragEvent({t:Date.now(),type:typeTag,button:e.button,buttons:e.buttons,ctrl:e.ctrlKey,meta:e.metaKey,target:e.target?.tagName,td:{rt:td.dataset.rowType,emp:td.dataset.empId,day:td.dataset.day}});
  };

  const handleDragMove = (e, typeTag)=>{
    // Ctrl/Metaを押しながらドラッグ中は既存選択に追加
    dragAdditive = dragAdditive || !!(e.ctrlKey || e.metaKey || ctrlDown);
    if(dragAdditive && ctrlTogglePending && dragStartCell){
      // 開始セルから動いた時点で「ドラッグ」とみなしてトグル保留を解除
      // （クリックなら mouseup でトグルする）
      const tdNow = resolveCellAtPoint(e.clientX, e.clientY);
      if(tdNow && tdNow!==dragStartCell){ ctrlTogglePending=false; ctrlPendingCell=null; }
    }
    // Ctrl/Metaを押しながらドラッグ中は既存選択に追加
    dragAdditive = dragAdditive || !!(e.ctrlKey || e.metaKey || ctrlDown);
    if(!pointerDown || !dragStartCell) return;
    dragAdditive = dragAdditive || !!(e.ctrlKey || e.metaKey || ctrlDown);
    // moveログ
    logDragEvent({t:Date.now(),type:`${typeTag}(move)`,buttons:e.buttons,ctrl:e.ctrlKey,meta:e.metaKey,target:e.target?.tagName});
    // moveログ（Ctrlドラッグが無視されていないか確認するため常に出す）
    logDragEvent({t:Date.now(),type:`${typeTag}(move)`,buttons:e.buttons,ctrl:e.ctrlKey,meta:e.metaKey,target:e.target?.tagName});

    const table = document.getElementById("scheduleTable");
    if(!table) return;

    let td = resolveCellAtPoint(e.clientX, e.clientY);
    if(!td){
      logDragEvent({t:Date.now(),type:`${typeTag}(no-td)`,buttons:e.buttons,target:e.target?.tagName});
      return;
    }
    if(td.dataset.rowType==="support") return;

    const startType = dragStartCell.dataset.rowType;
    const hitType = td.dataset.rowType;

    if(hitType !== startType){
      const empId = td.dataset.empId;
      const day = td.dataset.day;
      const snapped = table.querySelector(`td[data-row-type="${startType}"][data-emp-id="${CSS.escape(empId)}"][data-day="${CSS.escape(day)}"]`);
      if(snapped) td = snapped;
      else return;
    }

    dragMoveCount++;
    dragLastHit = { rowType: td.dataset.rowType, empId: td.dataset.empId||null, day: td.dataset.day||null };

    // セル境界を跨がなくても、一定距離動いたらドラッグとして扱う（Ctrl除外/追加の安定化）
    const movedPx = Math.abs(e.clientX - dragStartX) + Math.abs(e.clientY - dragStartY);
    if(!didDrag && td === dragStartCell && movedPx < 6) return;

    if(!didDrag){
      didDrag=true;
      dragSelecting=true;
      if(dom.scheduleInner) dom.scheduleInner.classList.add("drag-selecting");
    }

    updateDragSelection(td);
    if(dragAdditive && dragStartCell){ activeCell = dragStartCell; }
    logDragEvent({t:Date.now(),type:typeTag,buttons:e.buttons,ctrl:e.ctrlKey,meta:e.metaKey,target:e.target?.tagName,td:{rt:td.dataset.rowType,emp:td.dataset.empId,day:td.dataset.day}});
  };

  const endDrag = (e, typeTag)=>{
    // drag/click終了時はプレビューを必ず消す
    // （確定後の描画は選択クラスで表現）
    
    if(!pointerDown) return;

    const wasDragging = dragSelecting || didDrag;
    const movedPx = e ? (Math.abs((e.clientX||0) - dragStartX) + Math.abs((e.clientY||0) - dragStartY)) : 0;

    // Ctrlクリック（ドラッグ無し）: トグルで除外/追加（微小移動はクリック扱い）
    if(dragAdditive && ctrlPendingCell && (!wasDragging || movedPx < 6)){
      toggleCellInSelection(ctrlPendingCell);
      clearDragPreview();
    }

    // Ctrlドラッグ（確定）: プレビューを確定して適用
    if(dragAdditive && wasDragging && dragBaseSelectionKeys && dragLastRectKeys){
      const base = dragBaseSelectionKeys;
      const rect = dragLastRectKeys;

      let committed = new Set(base);

      if(dragRangeMode === "add"){
        rect.forEach(k=>committed.add(k));
      }else if(dragRangeMode === "remove"){
        rect.forEach(k=>{ if(base.has(k)) committed.delete(k); });
      }

      if(committed.size===0){
        committed.add(cellKey(dragStartCell));
      }
      applySelectionKeys(committed, dragStartCell);
      clearDragPreview();
    }

        clearDragPreview();

        if(activeCell){ try{ activeCell.focus({preventScroll:true}); }catch(_){ try{ activeCell.focus(); }catch(__){} } }

    pointerDown=false;
    didDrag=false;
    dragSelecting=false;
    dragStartCell=null;
    dragBaseSelectionKeys=null;
    dragLastRectKeys=null;
    ctrlTogglePending=false;
    ctrlPendingCell=null;
    dragAdditive=false;

    if(dom.scheduleInner) dom.scheduleInner.classList.remove("drag-selecting");
    if(wasDragging){
      dragJustEnded=true;
      setTimeout(()=>{ dragJustEnded=false; }, 0);
    }
    logDragEvent({t:Date.now(),type:typeTag,button:e?.button,buttons:e?.buttons,ctrl:e?.ctrlKey,meta:e?.metaKey,target:e?.target?.tagName});
  };

  // 表外クリックで強制リセット（再現していた「外を押すと復帰」を自動化）
  document.addEventListener("mousedown",(e)=>{
    const table = document.getElementById("scheduleTable");
    if(table && !table.contains(e.target)){
      endDrag(e, "md(out)");
    }
  }, true);

  // 開始（セル上）
  document.addEventListener("mousedown",(e)=>{
    // 設定モーダル等のUI上では勤務表セルのドラッグ選択を開始しない（selectが開かない等の干渉を防ぐ）
    if(dom.settingsModalBackdrop && dom.settingsModalBackdrop.classList.contains("visible") && dom.settingsModalBackdrop.contains(e.target)){
      return;
    }
    // フォーム要素上も干渉回避（ネイティブ挙動優先）
    const t = e.target;
    if(t && t.closest && t.closest("select, input, textarea, button, label")) return;

    if(e.button!==0) return;
    const td = resolveCellAtPoint(e.clientX, e.clientY);
    if(!td) return;
    startDragOnCell(td, e, "md");
  }, true);

  // 移動・終了
  document.addEventListener("mousemove",(e)=>{ handleDragMove(e, "mm"); }, true);
  document.addEventListener("mouseup",(e)=>{ endDrag(e, "mu"); }, true);



  document.addEventListener("selectstart", (e)=>{
    if(dragSelecting){
      e.preventDefault();
    }
  });
}


function getBaseInputMode(){ return (appState?.settings?.inputModeBase==="leave") ? "leave" : "work"; }
function getInputMode(){
  const base=getBaseInputMode();
  if(ctrlDown) return base==="work" ? "leave" : "work";
  return base;
}

// テンキー/数字キー入力を安定して拾う（NumLock OFF のナビゲーションキーも対応）
function normalizeDigitKey(e){
  if(!e) return null;
  if(typeof e.key === "string" && /^\d$/.test(e.key)) return e.key;
  if(typeof e.code === "string"){
    const m = e.code.match(/^Numpad(\d)$/);
    if(m) return m[1];
  }
  // NumLock OFF の場合、テンキーがナビゲーションキーとして送られることがある
  // ※通常の十字キー（ArrowUpなど）まで数字扱いにするとセル移動と衝突するため、
  //   codeがNumpad系のときだけ数字へ正規化する
  const map = {
    Home:"7", ArrowUp:"8", PageUp:"9",
    ArrowLeft:"4", Clear:"5", ArrowRight:"6",
    End:"1", ArrowDown:"2", PageDown:"3",
    Insert:"0"
  };
  if(typeof e.code === "string" && e.code.startsWith("Numpad") && map[e.key]) return map[e.key];
  return null;
}
// WS_GLOBAL_DIGIT_INPUT: 数字/テンキー(1-9)でコード入力（2段目(note)では通常入力を優先）
document.addEventListener("keydown",(e)=>{
  const digit = normalizeDigitKey(e);
  if(!digit) return;
  // 押しっぱなし連打は無視
  if(e.repeat) return;
  // IME合成中は無視
  if(e.isComposing || e.key==="Process") return;

  // 入力フォームや設定ダイアログ中は無視
  const t = e.target;
  if(t && (t.tagName==="INPUT"||t.tagName==="TEXTAREA"||t.tagName==="SELECT")) return;

  // アクティブセルが main のときだけコード入力として扱う（note は自由入力を優先）
  if(!activeCell) return;
  const td = (activeCell.tagName==="TD") ? activeCell : activeCell.closest?.("td");
  if(!td) return;
  if(td.dataset.rowType!=="main") return;

  // Ctrl+数字 のブラウザ既定(タブ切替等)を抑止
  try{ e.preventDefault(); }catch(_){}
  try{ e.stopPropagation(); }catch(_){}
  try{ e.stopImmediatePropagation(); }catch(_){}

  dbgKey("digit-enter", e, {digit});
  try{
    handleNumericInput(Number(digit));
    dbgKey("digit-apply", e, {digit});
  }catch(err){
    dbgKey("digit-error", e, {digit, msg:String(err)});
  }
}, true);



/* ===== Modals ===== */
let settingsAfterSaveCallback=null;
let settingsScope="auto"; // auto: 開いている年月が作成済みなら current、未作成なら template
let settingsTempStoreId = -1; // 設定モーダル内で未保存の新規店舗に付与する仮ID（負数）
function initModals(){
  dom.settingsCloseButton.addEventListener("click", closeSettingsModal);
  dom.settingsCancelButton.addEventListener("click", closeSettingsModal);
  dom.settingsSaveButton.addEventListener("click", ()=>{
    saveSettingsFromUI();

    // 追加: 「この変更を新規作成時にも反映」がONなら、現在年月のマスタをテンプレにも反映する
    try{
      const schedule = getCurrentSchedule();
      const propagateCb = document.getElementById("propagateToTemplateCheckbox");
      const canPropagate = !!schedule && isCurrentYearMonthLatest() && propagateCb && !propagateCb.disabled && propagateCb.checked;
      if(canPropagate){
        appState.templateMaster = deepClone(schedule.masterSnapshot);
        console.log("[settings] propagate current masterSnapshot -> templateMaster");
      }
    }catch(e){
      console.warn("[settings] propagate failed", e);
    }

    persistState();
    refreshAfterMasterChange();
    closeSettingsModal();
    if(settingsAfterSaveCallback){ const cb=settingsAfterSaveCallback; settingsAfterSaveCallback=null; cb(); }
  });


  document.querySelectorAll(".modal-tab").forEach(tab=>{
    tab.addEventListener("click", ()=>switchSettingsTab(tab));
  });

  dom.addStoreButton.addEventListener("click", ()=>addStoreRow());
  dom.addEmployeeButton.addEventListener("click", ()=>addEmployeeRow());
  dom.addLeaveCodeButton.addEventListener("click", ()=>addLeaveCodeRow());

  // Holidays settings
  if(dom.addHolidayButton){
    dom.addHolidayButton.addEventListener("click", ()=>addHolidayRow());
  }
  if(dom.holidayYearInput){
    dom.holidayYearInput.addEventListener("change", ()=>{ renderHolidaySettingsTable(); });
  }
  if(dom.holidayFetchButton){
    dom.holidayFetchButton.addEventListener("click", async ()=>{
      const year = Number(dom.holidayYearInput?.value)||null;
      if(!year){ alert("年を入力してください"); return; }
      const master=getSettingsMasterTarget();
      if(!master){ alert('この年月の勤務表が未作成です。'); return; }
      await ensureHolidayMasterForYear(master, year, {forceFetch:true});
      renderHolidaySettingsTable();
    });
  }


  dom.unsavedDiscardButton.addEventListener("click", ()=>{
    closeUnsavedModal();
    const target=pendingNavigationTarget;
    pendingNavigationTarget=null;
    if(target){
      dirty=false;
      updateSaveButton();
      performNavigation(target.storeId, target.yearMonth);
    }
  });
  dom.unsavedSaveButton.addEventListener("click", ()=>{
    persistState();
    dirty=false;
    updateSaveButton();
    closeUnsavedModal();
    const target=pendingNavigationTarget;
    pendingNavigationTarget=null;
    if(target) performNavigation(target.storeId, target.yearMonth);
  });
  dom.unsavedCancelButton.addEventListener("click", ()=>{
    closeUnsavedModal();
    pendingNavigationTarget=null;
  });

  dom.newScheduleWithSettingsButton.addEventListener("click", async ()=>{
    closeNewScheduleModal();
    openSettingsModal(async ()=>{ await createScheduleForCurrentYearMonth(); renderAll(); }, "template");
  });
  dom.newScheduleDirectButton.addEventListener("click", async ()=>{
    closeNewScheduleModal();
    await createScheduleForCurrentYearMonth();
    renderAll();
  });
  dom.newScheduleCancelButton.addEventListener("click", ()=>{
    closeNewScheduleModal();
    if(appState.settings.lastYearMonth){
      dom.monthInput.value=appState.settings.lastYearMonth;
      currentYearMonth=appState.settings.lastYearMonth;
    }
  });
}

/* ===== Debug Tools ===== */
let debugVisible=false;
function initDebugTools(){
  if(dom.debugButton){
    dom.debugButton.addEventListener("click", toggleDebugPanel);
  }
  document.addEventListener("keydown",(e)=>{
    if(e.ctrlKey && e.shiftKey && (e.key==="D" || e.key==="d")){
      e.preventDefault();
      toggleDebugPanel();
    }
  });

  if(dom.dbgExport) dom.dbgExport.addEventListener("click", exportAllData);
  if(dom.dbgCopyState) dom.dbgCopyState.addEventListener("click", copyDebugState);
  if(dom.dbgFixSnapshot) dom.dbgFixSnapshot.addEventListener("click", ()=>{
    fixCurrentMonthSnapshotFromTemplate();
    renderAll();
    alert("この年月の入力ルール（スナップショット）を最新マスタで更新しました。");
  });
  if(dom.dbgReset) dom.dbgReset.addEventListener("click", ()=>{
    if(!confirm("全データを初期化します。元に戻せません。よろしいですか？")) return;
    localStorage.removeItem(STORAGE_KEY);
    location.reload();
  });

  // 状態参照用
  window.__WS_DEBUG__ = {
    getState: ()=>getDebugState(),
    fixSnapshot: ()=>{ fixCurrentMonthSnapshotFromTemplate(); renderAll(); },
    export: ()=>exportAllData(),
    reset: ()=>{ localStorage.removeItem(STORAGE_KEY); location.reload(); }
  };

  setInterval(()=>{ if(debugVisible) renderDebugInfo(); }, 200);
  initClickInspector();
}

function toggleDebugPanel(){
  debugVisible=!debugVisible;
  if(dom.debugSection) dom.debugSection.style.display = debugVisible ? "block" : "none";
  if(!debugVisible) clearClickHighlight();
  if(debugVisible) renderDebugInfo();
}

/* Debug: Click Inspector */
let clickInspectEnabled=true;
let __dbgClickBox=null;
let __dbgClickBox2=null;
let __dbgClickLabel=null;
function initClickInspector(){
  // UI toggle
  if(dom.dbgClickInspectToggle){
    clickInspectEnabled = dom.dbgClickInspectToggle.classList.contains("on");
    dom.dbgClickInspectToggle.addEventListener("click",(e)=>{
      e.preventDefault();
      dom.dbgClickInspectToggle.classList.toggle("on");
      clickInspectEnabled = dom.dbgClickInspectToggle.classList.contains("on");
      if(!clickInspectEnabled) clearClickHighlight();
    });
  }
  // capture click anywhere (so we can see if something is overlaying)
  document.addEventListener("click", onDebugDocumentClick, true);
}

function onDebugDocumentClick(e){
  if(!debugVisible || !clickInspectEnabled) return;

  // allow interacting with debug panel itself without noise
  const inDebug = dom.debugSection && dom.debugSection.contains(e.target);
  if(inDebug) return;

  const x=e.clientX, y=e.clientY;
  const stack = (document.elementsFromPoint ? document.elementsFromPoint(x,y) : [document.elementFromPoint(x,y)]).filter(Boolean);
  const top = stack[0] || null;
  const second = stack[1] || null;

  // Update panel text
  if(dom.debugClickInfo){
    const lines=[];
    lines.push(`<div style="margin:2px 0 6px 0;"><b>クリック解析</b>（${Math.round(x)}, ${Math.round(y)}）</div>`);
    lines.push(`<div style="margin:0 0 6px 0;">上に乗っている順（最大5件）:</div>`);
    for(let i=0;i<Math.min(stack.length,5);i++){
      const el=stack[i];
      lines.push(renderElementLine(el,i));
    }
    dom.debugClickInfo.innerHTML = lines.join("");
  }

  // Highlight top + second to see overlays
  highlightElement(top, false, x, y);
  if(second) highlightElement(second, true);

  // prevent accidental actions when debugging (optional): keep default behavior
}

function renderElementLine(el, idx){
  const cs = getComputedStyle(el);
  const id = el.id ? `#${escapeHtml(el.id)}` : "";
  const cls = (el.className && typeof el.className==="string") ? "."+el.className.trim().replace(/\s+/g,".") : "";
  const tag = (el.tagName||"").toLowerCase();
  const pe = cs.pointerEvents;
  const zi = cs.zIndex==="auto" ? "auto" : cs.zIndex;
  const pos = cs.position;
  const op = cs.opacity;
  return `<div style="margin:2px 0;"><code>${idx+1}</code> <b>${escapeHtml(tag+id+cls)}</b> <span style="color:var(--text-muted);">pos:${pos} z:${escapeHtml(String(zi))} pe:${escapeHtml(String(pe))} op:${escapeHtml(String(op))}</span></div>`;
}

function escapeHtml(s){
  return String(s).replace(/[&<>"]/g, c=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;" }[c]));
}


// 祝日の日付入力表示: "YYYY/MM/DD (曜)"（内部保存は ISO: YYYY-MM-DD）
function formatHolidayDateDisplay(isoYmd){
  const iso = normalizeIsoDate(isoYmd);
  if(!iso) return "";
  const dt = new Date(iso + "T00:00:00");
  const w = ["日","月","火","水","木","金","土"][dt.getDay()];
  return iso.replaceAll("-","/") + " (" + w + ")";
}

// 入力（"YYYY/MM/DD", "YYYY-MM-DD", 末尾に"(曜)"付きも可）→ ISO "YYYY-MM-DD" / 失敗なら ""
function parseHolidayDateDisplay(s){
  if(!s) return "";
  const t = String(s).trim()
    .replace(/\s*\([日月火水木金土]\)\s*$/,""); // 末尾の(曜)を除去
  // YYYY/MM/DD or YYYY-MM-DD
  const m = t.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if(!m) return "";
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  if(!(y>=1900 && mo>=1 && mo<=12 && d>=1 && d<=31)) return "";
  const iso = String(y).padStart(4,"0")+"-"+String(mo).padStart(2,"0")+"-"+String(d).padStart(2,"0");
  // 実在日チェック
  const dt = new Date(iso+"T00:00:00");
  if(dt.getFullYear()!==y || (dt.getMonth()+1)!==mo || dt.getDate()!==d) return "";
  return iso;
}

function normalizeIsoDate(v){
  if(!v) return "";
  // すでにISOならそれ、そうでなければ parseHolidayDateDisplay で吸収
  const s = String(v).trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return parseHolidayDateDisplay(s);
}
function ensureClickOverlays(){
  if(!__dbgClickBox){
    __dbgClickBox=document.createElement("div");
    __dbgClickBox.className="debug-click-box";
    document.body.appendChild(__dbgClickBox);
  }
  if(!__dbgClickBox2){
    __dbgClickBox2=document.createElement("div");
    __dbgClickBox2.className="debug-click-box secondary";
    document.body.appendChild(__dbgClickBox2);
  }
  if(!__dbgClickLabel){
    __dbgClickLabel=document.createElement("div");
    __dbgClickLabel.className="debug-click-label";
    document.body.appendChild(__dbgClickLabel);
  }
}

function clearClickHighlight(){
  if(__dbgClickBox) __dbgClickBox.style.width="0px";
  if(__dbgClickBox2) __dbgClickBox2.style.width="0px";
  if(__dbgClickLabel) __dbgClickLabel.style.display="none";
  if(dom.debugClickInfo) dom.debugClickInfo.innerHTML="";
}

function highlightElement(el, isSecondary, x, y){
  ensureClickOverlays();
  const box = isSecondary ? __dbgClickBox2 : __dbgClickBox;
  if(!el || !box) return;
  const r = el.getBoundingClientRect();
  box.style.left = `${Math.max(0, r.left)}px`;
  box.style.top = `${Math.max(0, r.top)}px`;
  box.style.width = `${Math.max(0, r.width)}px`;
  box.style.height = `${Math.max(0, r.height)}px`;

  if(!isSecondary && __dbgClickLabel){
    const tag = (el.tagName||"").toLowerCase();
    const id = el.id ? `#${el.id}` : "";
    const cls = (el.className && typeof el.className==="string") ? "."+el.className.trim().replace(/\s+/g,".") : "";
    __dbgClickLabel.textContent = `${tag}${id}${cls}`;
    __dbgClickLabel.style.display="block";
    const pad=12;
    const lx = (x!=null?x:r.left) + pad;
    const ly = (y!=null?y:r.top) + pad;
    // keep on-screen
    const vw = window.innerWidth, vh = window.innerHeight;
    __dbgClickLabel.style.left = `${Math.min(lx, vw-560)}px`;
    __dbgClickLabel.style.top = `${Math.min(ly, vh-40)}px`;
  }
}

function getDebugState(){
  const schedule=getCurrentSchedule();
  const snap = schedule?.masterSnapshot || null;
  const active = activeCell ? {
    rowType: activeCell.dataset.rowType,
    empId: activeCell.dataset.empId,
    day: activeCell.dataset.day,
    text: (activeCell.textContent||"").trim(),
    readonly: isCellReadonly(activeCell)
  } : null;

  let mappingCount=0;
  try{
    const mode=getInputMode();
    const workContextStoreId=getWorkContextStoreId();
    const master = snap || appState.templateMaster;
    const mapping={};
    if(mode==="work"){
      master.workCodes.forEach(w=>{
        if(!w.key) return;
        if(isSupportMode && !(w.code==="E"||w.code==="F")) return;
        const store=master.stores.find(s=>s.id===workContextStoreId);
        if(!store) return;
        const sc=(store.storeCodes||[]).find(x=>x.code===w.code);
        if(!sc || sc.type==="disabled") return;
        mapping[w.key]=w.code;
      });
    }else{
      master.leaveCodes.forEach(l=>{
        if(!l.key) return;
        mapping[l.key]=l.code;
      });
    }
    mappingCount=Object.keys(mapping).length;
  }catch(_){}

  const store = (schedule?.masterSnapshot?.stores||[]).find(s=>s.id===currentStoreId) || appState.templateMaster.stores.find(s=>s.id===currentStoreId) || null;
  const ctxStoreId = getWorkContextStoreId();
  const ctxStore = (snap?.stores||[]).find(s=>s.id===ctxStoreId) || appState.templateMaster.stores.find(s=>s.id===ctxStoreId) || null;

  return {
    yearMonth: currentYearMonth,
    currentStoreId,
    currentStoreName: store?.name || null,
    supportMode: isSupportMode,
    ctrlDown,
    inputMode: getInputMode(),
    workContextStoreId: ctxStoreId,
    workContextStoreName: ctxStore?.name || null,
    hasSchedule: !!schedule,
    mappingCount,
    selectedCount: selectedCells.length,
    drag: {
      pointerDown,
      didDrag,
      dragSelecting,
      dragJustEnded,
      additive: dragAdditive,
      rangeMode: dragRangeMode,
      rangeModeStart: dragRangeModeStart,
      previewRect: dragLastRectKeys ? dragLastRectKeys.size : 0,
      moveCount: dragMoveCount,
      start: dragStartCell ? {
        rowType: dragStartCell.dataset.rowType,
        empId: dragStartCell.dataset.empId,
        day: dragStartCell.dataset.day
      } : null,
      lastHit: dragLastHit,
      recentEvents: dragEventLog.slice(-30)
    },
    activeCell: active,
    keyDebugCount: (window.__KEY_DEBUG__?.events||[]).length,
    keyDebugLast: (window.__KEY_DEBUG__?.events||[]).slice(-12)
  };
}

function renderDebugInfo(){
  if(!dom.debugInfo) return;
  const st=getDebugState();
  dom.debugInfo.innerHTML = `
    <div><strong>年月</strong>：${escapeHtml(st.yearMonth||"")}</div>
    <div><strong>店舗</strong>：${escapeHtml(String(st.currentStoreId||""))} / ${escapeHtml(st.currentStoreName||"")}</div>
    <div><strong>入力モード</strong>：${escapeHtml(st.inputMode)}（Ctrl=${st.ctrlDown}） / 応援=${st.supportMode}</div>
    <div><strong>有効キー数</strong>：${st.mappingCount}（0だと入力できません）</div>
    <div><strong>選択セル数</strong>：${st.selectedCount}</div>
    <div><strong>ドラッグ状態</strong>：pointerDown=${pointerDown} / didDrag=${didDrag} / dragSelecting=${dragSelecting} / dragJustEnded=${dragJustEnded}</div>
    <div><strong>ドラッグ計測</strong>：moveCount=${dragMoveCount} / start=${dragStartCell ? `${dragStartCell.dataset.rowType||"-"} emp=${dragStartCell.dataset.empId||"-"} day=${dragStartCell.dataset.day||"-"} idx=${dragStartCell.dataset.rowIndex||"-"}` : "(なし)"} / lastHit=${dragLastHit ? `${dragLastHit.rowType} emp=${dragLastHit.empId||"-"} day=${dragLastHit.day||"-"}` : "(なし)"}</div>
    <div><strong>アクティブセル</strong>：${st.activeCell ? `${escapeHtml(st.activeCell.rowType)} emp=${escapeHtml(st.activeCell.empId)} day=${escapeHtml(st.activeCell.day)} text="${escapeHtml(st.activeCell.text)}" readonly=${st.activeCell.readonly}` : "(なし)"}</div>
    <div class="divider"></div>
    <div class="text-muted">Consoleで <code>__WS_DEBUG__.getState()</code> を実行すると同じ情報がJSONで取得できます。</div>
<div class="divider"></div>
<div style="font-weight:800;margin-bottom:6px;">ドラッグイベントログ（最新）</div>
<div id="debugDragLog" style="font-family:ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, \"Liberation Mono\", \"Courier New\", monospace; font-size:10px; white-space:pre-wrap; color:var(--text-muted);"></div>
  `;
  if(dom.debugDragLog){
    const lines = dragEventLog.map(ev=>{
      const t = new Date(ev.t||Date.now()).toLocaleTimeString();
      const td = ev.td ? ` td(${ev.td.rt||""} emp=${ev.td.emp||""} day=${ev.td.day||""})` : "";
      return `${t} ${ev.type||""} btn=${ev.button??""} buttons=${ev.buttons??""} target=${ev.target||""}${td}
    <div style="margin-top:10px;padding-top:8px;border-top:1px solid #e5e7eb;">
      <strong>KEY_DEBUG (last 12)</strong>
      <pre style="white-space:pre-wrap;font-size:11px;line-height:1.35;margin:6px 0 0 0;background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;padding:8px;max-height:180px;overflow:auto;">${escapeHtml((window.__KEY_DEBUG__?.events||[]).slice(-12).map(ev=>{
        const dt=new Date(ev.t); const time=dt.toLocaleTimeString();
        const extra=ev.extra?JSON.stringify(ev.extra):"";
        return `${time} [${ev.stage}] key=${ev.key} code=${ev.code} shift=${ev.shift} ctrl=${ev.ctrl} composing=${ev.composing} prevented=${ev.defaultPrevented} target=${ev.target} activeEl=${ev.activeEl} activeCell=${ev.activeCell} anchor=${ev.anchor} ${extra}`;
      }).join("\n"))}</pre>
    </div>
`;
    });
    dom.debugDragLog.textContent = lines.slice(-12).join("\n");
  }
}

function exportAllData(){
  const data=localStorage.getItem(STORAGE_KEY) || "";
  const blob=new Blob([data], {type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url;
  a.download=`workSchedule_backup_${(currentYearMonth||"data")}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function copyDebugState(){
  try{
    const st=getDebugState();
    await navigator.clipboard.writeText(JSON.stringify(st,null,2));
    alert("状態をクリップボードにコピーしました。");
  }catch(e){
    console.warn(e);
    alert("コピーに失敗しました（クリップボード権限が必要）。Consoleで __WS_DEBUG__.getState() を実行してください。");
  }
}

// “この年月のスナップショット”が古くて業務コードが無効のまま → 入力できないケースを救済
function fixCurrentMonthSnapshotFromTemplate(){
  const schedule=getCurrentSchedule();
  if(!schedule) return;
  const t=appState.templateMaster;

  // 店舗情報（storeCodes含む）をテンプレからコピーして、この年月だけ更新
  schedule.masterSnapshot.stores = deepClone(t.stores);
  schedule.masterSnapshot.workCodes = deepClone(t.workCodes);
  schedule.masterSnapshot.leaveCodes = deepClone(t.leaveCodes);

  // 既存の勤務表データは保持（byStore / assignmentsは触らない）
  persistState();
}


function openSettingsModal(afterSaveCb){
  settingsAfterSaveCallback = afterSaveCb || null;
  const schedule = getCurrentSchedule();
  settingsScope = schedule ? "current" : "template";
  // スコープ表示（簡易メモ）
  const noteEl = document.getElementById("settingsScopeNote");
  if(noteEl){
    noteEl.textContent = schedule
      ? "現在開いている年月のマスタ（この勤務表だけに反映）を編集します。"
      : "未作成の年月のため、今後の新規作成用マスタを編集します。";
  }

  // 「この変更を新規作成時にも反映」：既存の最新勤務表を開いている時のみ表示（デフォルトON）
  const propagateWrap=document.getElementById("propagateToTemplateWrap");
  const propagateCb=document.getElementById("propagateToTemplateCheckbox");
  const canPropagate=!!schedule && isCurrentYearMonthLatest();
  if(propagateWrap){
    propagateWrap.style.display = canPropagate ? "block" : "none";
  }
  if(propagateCb){
    propagateCb.checked = true;
    propagateCb.disabled = !canPropagate;
  }
  populateSettingsFromState();
  // 仕様: 設定画面を開いたら常に「店舗」タブを初期表示する（前回の表示状態は引き継がない）
  try{
    const storesTab=document.querySelector('.modal-tab[data-target="settingsStores"]');
    if(storesTab) switchSettingsTab(storesTab);
  }catch(_e){}
  dom.settingsModalBackdrop.classList.add("visible");
}
function closeSettingsModal(){ dom.settingsModalBackdrop.classList.remove("visible"); }
function switchSettingsTab(tabElem){
  document.querySelectorAll(".modal-tab").forEach(t=>t.classList.remove("active"));
  tabElem.classList.add("active");
  const targetId=tabElem.dataset.target;
  document.querySelectorAll(".settings-panel").forEach(p=>p.classList.remove("active"));
  document.getElementById(targetId).classList.add("active");

  if(targetId==="settingsEmployees"){
    refreshEmployeeStoreOptionsFromSettings();
    // 従業員タブを開いたタイミングで右側（店舗別リスト）を必ず再描画し、
    // DnDのwireも確実に初期化する。
    try{ refreshEmployeeOrderPanel(); }catch(_e){}
    try{ applyEmployeeLeftFilter(); }catch(_e){}
  }
}
function openUnsavedModal(){ dom.unsavedModalBackdrop.classList.add("visible"); }
function closeUnsavedModal(){ dom.unsavedModalBackdrop.classList.remove("visible"); }

function showNewScheduleModal(){
  const [year,month]=currentYearMonth.split("-").map(v=>parseInt(v,10));
  dom.newScheduleMessage.textContent = `${year}年${month}月の勤務表が存在しません。新規作成します。`;
  dom.newScheduleModalBackdrop.classList.add("visible");
}
function closeNewScheduleModal(){ dom.newScheduleModalBackdrop.classList.remove("visible"); }

/* ===== Initial Context ===== */
function loadInitialContext(){
  const now=new Date();
  const ymNow = `${now.getFullYear().toString().padStart(4,"0")}-${(now.getMonth()+1).toString().padStart(2,"0")}`;
  currentYearMonth = appState.settings.lastYearMonth || ymNow;
  dom.monthInput.value=currentYearMonth;

  const stores=appState.templateMaster.stores;
  if(stores.length>0){
    currentStoreId = appState.settings.lastStoreId || stores.slice().sort((a,b)=>(a.order||0)-(b.order||0))[0].id;
  }else{
    currentStoreId=null;
  }
  updateAutoSaveToggleUI();
}

/* ===== Navigation ===== */
function navigateTo(targetStoreId, targetYearMonth){
  if(!targetYearMonth) return;
  const target={storeId:targetStoreId, yearMonth:targetYearMonth};
  if(!appState.settings.autoSave && dirty){
    pendingNavigationTarget=target;
    openUnsavedModal();
    return;
  }
  performNavigation(targetStoreId, targetYearMonth);
}

function performNavigation(targetStoreId, targetYearMonth){
  currentYearMonth=targetYearMonth;
  appState.settings.lastYearMonth=targetYearMonth;

  const stores=appState.templateMaster.stores;
  if(!targetStoreId && stores.length>0) currentStoreId=stores[0].id;
  else currentStoreId=targetStoreId;

  appState.settings.lastStoreId=currentStoreId;

  dirty=false;
  updateSaveButton();

  if(!appState.schedules[currentYearMonth]){
    showNewScheduleModal();
  }
  renderAll();
  persistState();
}

/* ===== Render ===== */
function renderAll(){
  renderStoreOptions();
  renderSupportStoreOptions();
  renderKeypad();
  renderSchedule();
}

function renderStoreOptions(){
  const schedule=getCurrentSchedule();
  const baseStores=(schedule? schedule.masterSnapshot.stores : appState.templateMaster.stores);
  const stores=baseStores.slice().sort((a,b)=>(a.order||0)-(b.order||0));
  dom.storeSelect.innerHTML="";
  if(stores.length===0){
    const opt=document.createElement("option");
    opt.value="";
    opt.textContent="（店舗なし：設定から追加）";
    dom.storeSelect.appendChild(opt);
    dom.storeSelect.disabled=true;
    currentStoreId=null;
    return;
  }
  dom.storeSelect.disabled=false;
  stores.forEach(store=>{
    const opt=document.createElement("option");
    opt.value=String(store.id);
    opt.textContent = store.name + (store.enabled ? "" : "（無効）");
    dom.storeSelect.appendChild(opt);
  });
  if(!currentStoreId || !stores.some(s=>s.id===currentStoreId)) currentStoreId=stores[0].id;
  dom.storeSelect.value=String(currentStoreId);
}

function renderSupportStoreOptions(){
  const schedule=getCurrentSchedule();
  const baseStores=(schedule? schedule.masterSnapshot.stores : appState.templateMaster.stores);
  const stores=baseStores.slice().sort((a,b)=>(a.order||0)-(b.order||0));
  const prev = dom.supportStoreSelect.value || "self";
  dom.supportStoreSelect.innerHTML="";
  const optSelf=document.createElement("option");
  optSelf.value="self";
  optSelf.textContent="自店舗（通常）";
  dom.supportStoreSelect.appendChild(optSelf);
  stores.forEach(store=>{
    const opt=document.createElement("option");
    opt.value=String(store.id);
    opt.textContent=store.name;
    dom.supportStoreSelect.appendChild(opt);
  });
  dom.supportStoreSelect.value = Array.from(dom.supportStoreSelect.options).some(o=>o.value===prev) ? prev : "self";
  isSupportMode = dom.supportStoreSelect.value !== "self";
}

function renderKeypad(){
  const mode=getInputMode();
  // 設定（マスタ）は「現在開いている年月にだけ反映」されるケースがあるため、
  // パレットは必ず “現在の勤務表があればそのスナップショット” を参照する。
  // （未作成の場合のみテンプレートマスタを参照）
  const schedule=getCurrentSchedule();
  const master=(schedule? schedule.masterSnapshot : appState.templateMaster);
  const workContextStoreId=getWorkContextStoreId();
  const mapping={};
  if(mode==="work"){
    master.workCodes.forEach(w=>{
      if(!w.key) return;
      if(isSupportMode && !(w.code==="E"||w.code==="F")) return;
      const store=master.stores.find(s=>s.id===workContextStoreId);
      if(!store) return;
      const sc=(store.storeCodes||[]).find(x=>x.code===w.code);
      if(!sc || sc.type==="disabled") return;
      mapping[w.key]={type:"work",code:w.code,shiftType:sc.type};
    });
  }else{
    master.leaveCodes.forEach(l=>{
      if(!l.key) return;
      mapping[l.key]={type:"leave",code:l.code};
    });
  }

  // 入力できない原因の大半は「この店舗で有効な業務コードが0件」または「セルが未選択」です。
  const availCount = Object.keys(mapping).length;
  if(getInputMode()==="work" && availCount===0){
    dom.keypad.innerHTML = "";
    const warn=document.createElement("div");
    warn.className="alert-item";
    warn.textContent="この店舗（または応援先）で有効な業務コードがありません。設定→店舗→業務コードを「日勤/夜勤」にしてください。";
    dom.alertList.innerHTML="";
    dom.alertList.appendChild(warn);
  }
  dom.keypad.innerHTML="";
  const layout=["7","8","9","4","5","6","1","2","3"];
  layout.forEach(num=>{
    const btn=document.createElement("button");
    btn.type="button";
    btn.className="kp-key";

    // 左上にテンキー番号（コード有無で統一）
    const numTag=document.createElement("span");
    numTag.className="kp-num";
    numTag.textContent=num;
    btn.appendChild(numTag);

    const info=mapping[num];
    if(info){
      const leftWrap=document.createElement("span");
      leftWrap.className="key-left";

      if(info.type==="work"){
        const dot=document.createElement("span");
        dot.className="status-dot "+(info.shiftType==="day"?"dot-day":(info.shiftType==="night"?"dot-night":"dot-disabled"));
        dot.title = (info.shiftType==="day"?"日勤":(info.shiftType==="night"?"夜勤":"無効"));
        leftWrap.appendChild(dot);
      }

      const spanCode=document.createElement("span");
      spanCode.className="code-label";
      spanCode.textContent=info.code;
      leftWrap.appendChild(spanCode);

      btn.appendChild(leftWrap);
      btn.addEventListener("click",()=>{
        if(!activeCell) return;
        if(activeCell.dataset.rowType!=="main") return;
        handleNumericInput(num);
      });
    }else{
      btn.disabled=true;
    }
    dom.keypad.appendChild(btn);
  });

  {
    const base=getBaseInputMode();
    const isTemp = ctrlDown && (mode!==base);
    if(mode==="leave"){
      dom.modeBadge.textContent = "休暇コード" + (isTemp?"（Ctrl）":"");
      dom.modeBadge.className = "badge leave";
    }else{
      dom.modeBadge.textContent = (isSupportMode ? "業務コード（応援）" : "業務コード") + (isTemp?"（Ctrl）":"");
      dom.modeBadge.className = "badge "+(isSupportMode?"support":"work");
    }
  }
}

function renderSchedule(){
  if(!dragAdditive && !preserveSelectionOnRender){ clearSelection(); }
  dom.scheduleInner.innerHTML="";
  const stores=appState.templateMaster.stores;
  if(!currentYearMonth || stores.length===0 || !currentStoreId){
    dom.scheduleInner.appendChild(dom.schedulePlaceholder);
    dom.schedulePlaceholder.style.display="block";
    dom.scheduleTitle.textContent="";
    return;
  }

  const schedule=appState.schedules[currentYearMonth];
  if(!schedule){
    dom.scheduleInner.appendChild(dom.schedulePlaceholder);
    dom.schedulePlaceholder.style.display="block";
    const y=parseInt(currentYearMonth.slice(0,4),10);
    const m=parseInt(currentYearMonth.slice(5,7),10);
    const wk=toWareki(y,m);
    dom.scheduleTitle.textContent=`勤務予定表　　${wk.gengo}${wk.yearStr}年 ${m}月分（未作成）`;
    return;
  }

  dom.schedulePlaceholder.style.display="none";
  currentSchedule=schedule;

  const storeSnapshot=schedule.masterSnapshot.stores.find(s=>s.id===currentStoreId);
  if(!storeSnapshot){
    dom.scheduleInner.appendChild(dom.schedulePlaceholder);
    dom.schedulePlaceholder.style.display="block";
    dom.schedulePlaceholder.textContent="この年月の勤務表には、選択された店舗が含まれていません。";
    dom.scheduleTitle.textContent="";
    return;
  }

  const wk=toWareki(schedule.year, schedule.month);
  dom.scheduleTitle.textContent=`${storeSnapshot.name}　勤務予定表　　${wk.gengo}${wk.yearStr}年　${schedule.month}月分`;

  currentStoreSchedule=schedule.byStore[currentStoreId];
  const table=buildScheduleTable(schedule, currentStoreId);
  dom.scheduleInner.appendChild(table);

  validateNightShiftsForCurrent();

  if(pendingFocus){
    const {empId,rowType,day}=pendingFocus;
    pendingFocus=null;
    const tgt=table.querySelector(`td[data-row-type="${rowType}"][data-emp-id="${empId}"][data-day="${day}"]`);
    if(tgt){ tgt.focus(); setActiveCell(tgt,false); }
  }

  // 複数セル入力などで再描画しても選択を維持
  if(preserveSelectionOnRender){
    try{
      const keys = preserveSelectionKeys ? new Set(preserveSelectionKeys) : new Set();
      const prefer = preserveSelectionActiveKey ? getCellByKey(preserveSelectionActiveKey) : null;
      preserveSelectionOnRender=false;
      preserveSelectionKeys=null;
      preserveSelectionActiveKey=null;
      if(keys.size>0){
        applySelectionKeys(keys, prefer);
      }
    }catch(err){
      // fail-safe
      preserveSelectionOnRender=false;
      preserveSelectionKeys=null;
      preserveSelectionActiveKey=null;
    }
  }
}

/* ===== Table build ===== */
function buildScheduleTable(schedule, storeId){
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
  const table=document.createElement("table");
  table.className="schedule";
  table.id="scheduleTable";
  const thead=document.createElement("thead");
  const tbody=document.createElement("tbody");

  const tr1=document.createElement("tr");
  const thCorner=document.createElement("th");
  thCorner.rowSpan=2;
  thCorner.className="col-employee-header namecol";
  thCorner.textContent="従業員";
  tr1.appendChild(thCorner);
  const master = schedule.masterSnapshot || {};
  for(let d=1; d<=daysInMonth; d++){
    const date=new Date(schedule.year, schedule.month-1, d);
    const w=date.getDay();
    const hol=getHolidayInfoFromMaster(master, date);
    const isHol=!!hol;
    const th=document.createElement("th");
    th.textContent=d;
    if(isHol){ th.classList.add("day-holiday"); th.title = hol.name || ""; }
    else if(w===0) th.classList.add("day-sun");
    else if(w===6) th.classList.add("day-sat");
    th.dataset.day=String(d);
    if(isHol) th.classList.add("day-holiday");
    else if(w===0) th.classList.add("day-sun");
    else if(w===6) th.classList.add("day-sat");
    tr1.appendChild(th);
  }
  thead.appendChild(tr1);

  const tr2=document.createElement("tr");
  tr2.className="weekday-row";
  const weekDays=["日","月","火","水","木","金","土"];
  for(let d=1; d<=daysInMonth; d++){
    const date=new Date(schedule.year, schedule.month-1, d);
    const w=date.getDay();
    const hol=getHolidayInfoFromMaster(master, date);
    const isHol=!!hol;
    const th=document.createElement("th");
    th.textContent=weekDays[w];
    th.dataset.day=String(d);
    if(isHol){ th.classList.add("day-holiday"); th.title = hol.name || ""; }
    else if(w===0) th.classList.add("day-sun");
    else if(w===6) th.classList.add("day-sat");
    tr2.appendChild(th);
  }
  thead.appendChild(tr2);

  const employees=getEmployeesInStore(schedule, storeId);
  const maxNameLen=employees.reduce((mx,e)=>Math.max(mx,(e.lastName+e.firstName).length),3);
  const nameColWidth=Math.min(160, Math.max(84, 10*maxNameLen+40));
  thCorner.style.minWidth=nameColWidth+"px";
  thCorner.style.maxWidth=nameColWidth+"px";

  let rowIndexCounter=0;

  employees.forEach(emp=>{
    const empId=emp.id;
    const formatted=formatEmployeeName(emp.lastName, emp.firstName);

    const mainRow=document.createElement("tr");
    mainRow.className="employee-main-row";
    const noteRow=document.createElement("tr");
    noteRow.className="employee-note-row";

    const nameTh=document.createElement("th");
    nameTh.rowSpan=2;
    nameTh.className="cell-emp-name";
    nameTh.style.minWidth=nameColWidth+"px";
    nameTh.style.maxWidth=nameColWidth+"px";
    nameTh.textContent=formatted;
    mainRow.appendChild(nameTh);

    for(let d=1; d<=daysInMonth; d++){
      const mainCell=document.createElement("td");
      const noteCell=document.createElement("td");
      mainCell.tabIndex=0; noteCell.tabIndex=0;

      mainCell.dataset.empId=String(empId);
      mainCell.dataset.day=String(d);
      mainCell.dataset.rowType="main";
      mainCell.dataset.rowIndex=String(rowIndexCounter);

      noteCell.dataset.empId=String(empId);
      noteCell.dataset.day=String(d);
      noteCell.dataset.rowType="note";
      noteCell.dataset.rowIndex=String(rowIndexCounter+1);

      const vMain=getAssignmentValue(schedule, storeId, empId, "main", d)||"";
      const vNote=getAssignmentValue(schedule, storeId, empId, "note", d)||"";

      renderCellValue(mainCell, vMain, schedule, storeId);
      renderCellValue(noteCell, vNote, schedule, storeId);

      attachCellEvents(mainCell);
      attachCellEvents(noteCell);

      mainRow.appendChild(mainCell);
      noteRow.appendChild(noteCell);
    }

    rowIndexCounter+=2;
    tbody.appendChild(mainRow);
    tbody.appendChild(noteRow);
  });

  const storeSnapshot=schedule.masterSnapshot.stores.find(s=>s.id===storeId);
  const usedCodes=(storeSnapshot.storeCodes||[]).filter(c=>c.type!=="disabled" && (c.code==="E"||c.code==="F"));
  usedCodes.forEach((sc,idx)=>{
    const code=sc.code;
    const label=code+"謹";
    const supportRow=document.createElement("tr");
    supportRow.className = (idx===0 ? "support-row support-first" : "support-row");
    const th=document.createElement("th");
    th.className="cell-support-header";
    th.style.minWidth=nameColWidth+"px";
    th.style.maxWidth=nameColWidth+"px";
    th.textContent=label;
    supportRow.appendChild(th);

    for(let d=1; d<=daysInMonth; d++){
      const cell=document.createElement("td");
      cell.dataset.rowType="support";
      cell.dataset.day=String(d);
      cell.classList.add("cell-readonly");
      const supporter=getSupporterFor(schedule, storeId, code, d);
      cell.textContent=supporter||"";
      supportRow.appendChild(cell);
    }
    tbody.appendChild(supportRow);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  return table;
}

/* ===== New schedule create ===== */
async function createScheduleForCurrentYearMonth(){
  const [year,month]=currentYearMonth.split("-").map(v=>parseInt(v,10));
  const template=appState.templateMaster;
  await ensureHolidayMasterForYear(template, year);
  const snapshot={
    stores:deepClone(template.stores),
    employees:deepClone(template.employees),
    workCodes:deepClone(template.workCodes),
    leaveCodes:deepClone(template.leaveCodes),
    holidays:deepClone(template.holidays||{}),
  };
  const schedule={ year,month, masterSnapshot:snapshot, byStore:{} };
  const daysInMonth=getDaysInMonth(year,month);

  snapshot.stores.filter(s=>s.enabled).forEach(store=>{
    const storeId=store.id;
    const empList=snapshot.employees.filter(e=>e.storeId===storeId);
    const assignments={};
    empList.forEach(emp=>{
      assignments[emp.id]={ main:new Array(daysInMonth).fill(""), note:new Array(daysInMonth).fill("") };
    });
    schedule.byStore[storeId]={assignments};
  });

  applyPendingCarryoversToSchedule(currentYearMonth, schedule);
  appState.schedules[currentYearMonth]=schedule;
  persistState();
}

function applyPendingCarryoversToSchedule(yearMonth, schedule){
  const pending=appState.pendingNightCarryovers?.[yearMonth];
  if(!pending) return;

  Object.keys(pending).forEach(storeIdStr=>{
    const storeId=Number(storeIdStr);
    const byStore=schedule.byStore[storeId];
    if(!byStore) return;
    const empMap=pending[storeIdStr];
    Object.keys(empMap).forEach(empIdStr=>{
      const empId=Number(empIdStr);
      const dayMap=empMap[empIdStr];
      const val=dayMap["1"];
      if(!val) return;
      if(!byStore.assignments[empId]) return;
      byStore.assignments[empId].main[0]=val;
    });
  });

  delete appState.pendingNightCarryovers[yearMonth];
}

/* ===== schedule data get/set ===== */
function getCurrentSchedule(){ return appState.schedules[currentYearMonth]||null; }

function getAssignmentValue(schedule, storeId, empId, rowType, day){
  const storeData=schedule.byStore[storeId];
  if(!storeData) return "";
  const empData=storeData.assignments[empId];
  if(!empData) return "";
  const arr = rowType==="main" ? empData.main : empData.note;
  return arr[day-1] || "";
}

function setAssignmentValue(schedule, storeId, empId, rowType, day, value){
  const storeData=schedule.byStore[storeId];
  if(!storeData) return;
  if(!storeData.assignments[empId]){
    storeData.assignments[empId]={main:[],note:[]};
  }
  const arr = rowType==="main" ? storeData.assignments[empId].main : storeData.assignments[empId].note;

  const prev = arr[day-1] || "";

  // ★重要：夜勤入り(ﾃ) / 夜勤明け(ﾋ) を編集・上書き・削除したときにペアが孤立しないよう自動で掃除する
  if(rowType==="main" && prev){
    const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
    const newVal = value || "";

    // ﾃ を消した/上書きした → 翌日の ﾋ を消す（同月 or 翌月1日/保留）
    if(prev.includes("ﾃ") && !newVal.includes("ﾃ")){
      if(day < daysInMonth){
        const nextPrev = arr[day] || "";
        if(nextPrev && nextPrev.includes("ﾋ")){
          arr[day] = ""; // day+1
        }
      }else{
        const nextYm=nextMonthKey(currentYearMonth);
        const nextSchedule=appState.schedules[nextYm];
        if(nextSchedule && nextSchedule.byStore?.[storeId]?.assignments?.[empId]){
          const nextArr = nextSchedule.byStore[storeId].assignments[empId].main;
          const v = (nextArr?.[0]||"");
          if(v && v.includes("ﾋ")) nextArr[0] = "";
        }
        const p=appState.pendingNightCarryovers?.[nextYm]?.[storeId]?.[empId];
        if(p && p["1"]) delete p["1"];
      }
    }

    // ﾋ を消した/上書きした → 前日の ﾃ を消す（同月 or 前月末）
    if(prev.includes("ﾋ") && !newVal.includes("ﾋ")){
      if(day > 1){
        const prevPrev = arr[day-2] || "";
        if(prevPrev && prevPrev.includes("ﾃ")){
          arr[day-2] = ""; // day-1
        }
      }else{
        const prevYm=prevMonthKey(currentYearMonth);
        const prevSchedule=appState.schedules[prevYm];
        if(prevSchedule && prevSchedule.byStore?.[storeId]?.assignments?.[empId]){
          const prevDays=getDaysInMonth(prevSchedule.year, prevSchedule.month);
          const prevArr = prevSchedule.byStore[storeId].assignments[empId].main;
          const v = (prevArr?.[prevDays-1]||"");
          if(v && v.includes("ﾃ")) prevArr[prevDays-1] = "";
        }
      }
    }
  }

  arr[day-1] = value || "";
  markDirty();
}

function markDirty(){
  dirty=true;
  if(appState.settings.autoSave){
    persistState();
    dirty=false;
  }
  updateSaveButton();
}
function updateSaveButton(){ dom.saveButton.disabled = appState.settings.autoSave ? true : !dirty; }
function updateAutoSaveToggleUI(){
  if(appState.settings.autoSave) dom.autoSaveToggle.classList.add("on");
  else dom.autoSaveToggle.classList.remove("on");
}

/* ===== Cell events / selection ===== */
function attachCellEvents(cell){
  const rowType=cell.dataset.rowType;
  if(rowType==="support") return;

  cell.addEventListener("focus",()=>{
    if(dragSelecting) return;

    // Ctrl/Metaクリック直後は、mousedown側で選択処理を完結させる（focus側で選択を壊さない）
    if(lastMouseDownAdditive){
      // アクティブ枠だけ整える
      if(activeCell && activeCell!==cell){ activeCell.classList.remove('cell-selected'); }
      activeCell = cell;
      cell.classList.add('cell-selected');
      // 既に選択されていなければ追加（通常は既に追加済み）
      addCellToSelection(cell);
      lastMouseDownAdditive = false;
      return;
    }

    setActiveCell(cell, selectedCells.length>1);
  });

  cell.addEventListener("mousedown", (e)=>{ /* document側で処理 */ });

  cell.addEventListener("click",(e)=>{
    // ドラッグ範囲選択の直後に click が発火すると範囲選択が単一セルに潰れるため抑止
    if(dragJustEnded) return;
    // Ctrl/Metaクリックは document mousedown 側でトグル処理済み。ここで setActiveCell すると既存選択が消えるので何もしない
    if(e && (e.ctrlKey || e.metaKey || ctrlDown)) return;
    if(!dragSelecting) setActiveCell(cell,false);
  });

  cell.addEventListener("mouseover",()=>{ /* document mousemoveで処理 */ });
  cell.addEventListener("keydown",(e)=>handleCellKeyDown(e,cell));

  if(rowType==="note"){
    cell.contentEditable="true";
    cell.addEventListener("input",()=>{
      const schedule=getCurrentSchedule();
      if(!schedule) return;
      const empId=Number(cell.dataset.empId);
      const day=Number(cell.dataset.day);
      const val=(cell.textContent||"").trim();
      setAssignmentValue(schedule,currentStoreId,empId,"note",day,val);
    });
  }
}

function setActiveCell(cell, keepMulti){
  // Shift+矢印の範囲選択中はアンカーを維持（setActiveCellではリセットしない）
  if(!suppressShiftAnchorReset){ /* no-op */ }
  // クリック/移動でアクティブが変わったら Shift 範囲選択の起点をリセット

  if(!cell) return;
  if(!keepMulti){
    clearSelection();
  }else{
    // keepMultiのときは選択を維持し、段違いは無効
    if(selectedCells.length>0){
      const baseType = selectedCells[0].dataset.rowType;
      if(cell.dataset.rowType !== baseType) return;
    }
  }
  // 以前のアクティブ強調を解除
  if(activeCell && activeCell!==cell){ activeCell.classList.remove("cell-selected"); }
  activeCell = cell;
  addCellToSelection(cell);
  cell.classList.add("cell-selected");
  cell.classList.add("cell-multi");

  // キーボード移動の起点は「フォーカス」なので、常にアクティブへフォーカスを同期
  try{ cell.focus({preventScroll:true}); }catch(_){ try{ cell.focus(); }catch(__){} }

  renderDebugInfo && renderDebugInfo();
}


function clearSelection(){
  // 既存の選択表示を解除
  selectedCells.forEach(c=>c.classList.remove("cell-selected","cell-multi"));
  // 念のため残存する強調も解除
  document.querySelectorAll('#scheduleTable td.cell-selected, #scheduleTable td.cell-multi')
    .forEach(c=>c.classList.remove('cell-selected','cell-multi'));
  selectedCells=[];
  selectedKeySet=new Set();
  activeCell=null;
}

function cellKey(cell){
  return `${cell.dataset.rowType}|${cell.dataset.empId}|${cell.dataset.day}`;
}
function getCellByKey(key){
  const [rt, empId, day] = key.split("|");
  const table = document.getElementById("scheduleTable");
  if(!table) return null;
  return table.querySelector(`td[data-row-type="${rt}"][data-emp-id="${CSS.escape(empId)}"][data-day="${CSS.escape(day)}"]`);
}

function clearDragPreview(){
  const table = document.getElementById('scheduleTable');
  if(!table) return;
  table.querySelectorAll('td.cell-preview-add, td.cell-preview-remove')
    .forEach(c=>c.classList.remove('cell-preview-add','cell-preview-remove'));
  if(dom.dragHint){
    dom.dragHint.classList.add('hidden');
    dom.dragHint.classList.remove('add','remove');
    dom.dragHintText.textContent='';
    dom.dragHintPill.textContent='';
  }
}
function setDragHint(mode, count){
  if(!dom.dragHint) return;
  if(!mode || count<=0){
    dom.dragHint.classList.add('hidden');
    return;
  }
  dom.dragHint.classList.remove('hidden','add','remove');
  dom.dragHint.classList.add(mode);
  if(mode==='add'){
    dom.dragHintText.textContent = '追加選択のプレビュー';
    dom.dragHintPill.textContent = `+${count}`;
  }else if(mode==='remove'){
    dom.dragHintText.textContent = '選択解除のプレビュー';
    dom.dragHintPill.textContent = `-${count}`;
  }
}

function rebuildSelectedCellsFromKeys(){
  selectedCells = [];
  const table = document.getElementById("scheduleTable");
  if(!table) return;
  selectedKeySet.forEach(k=>{
    const c = getCellByKey(k);
    if(c){ selectedCells.push(c); c.classList.add("cell-multi"); }
  });
}
function applySelectionKeys(newKeys, activePrefer){
  const committing = !pointerDown; // mouseup等で確定反映する局面
  const table = document.getElementById('scheduleTable');

  if(committing && table){
    // 確定時は全セルから選択表示を完全に消してから再付与（取り残し防止）
    table.querySelectorAll('td.cell-selected, td.cell-multi').forEach(c=>{
      c.classList.remove('cell-selected','cell-multi');
    });
  }else{
    // ドラッグ中は最小限（ちらつき防止）
    selectedCells.forEach(c=>{
      c.classList.remove('cell-multi');
      if(!(activeCell && c===activeCell)) c.classList.remove('cell-selected');
    });
  }

  selectedKeySet = new Set(newKeys);

  // 再構築
  selectedCells = [];
  selectedKeySet.forEach(k=>{
    const c = getCellByKey(k);
    if(c){
      c.classList.add('cell-multi');
      selectedCells.push(c);
    }
  });

  // active の決定
  let nextActive = null;

  if(activePrefer){
    const kp = cellKey(activePrefer);
    if(selectedKeySet.has(kp)) nextActive = getCellByKey(kp);
  }
  if(!nextActive && activeCell){
    const ka = cellKey(activeCell);
    if(selectedKeySet.has(ka)) nextActive = getCellByKey(ka);
  }
  if(!nextActive){
    nextActive = selectedCells[selectedCells.length-1] || null;
  }

  if(activeCell && activeCell!==nextActive) activeCell.classList.remove('cell-selected');
  activeCell = nextActive;
  if(activeCell) activeCell.classList.add('cell-selected');

  // focus は endDrag 側で 1回だけ
}




function addCellToSelection(cell){
  const k = cellKey(cell);
  if(!selectedKeySet.has(k)){
    selectedKeySet.add(k);
    selectedCells.push(cell);
  }
  cell.classList.add("cell-multi");
}

function pickNextActiveAfterRemoval(removedKey){
  // できるだけ同じ従業員・同じ段で、日付が近いセルをアクティブにする
  const [rt, empId, dayStr] = removedKey.split("|");
  const removedDay = parseInt(dayStr,10);

  // 同じ行の候補を抽出
  const sameRow = Array.from(selectedKeySet).filter(k=>{
    const [rr, ee] = k.split("|");
    return rr===rt && ee===empId;
  }).map(k=>{
    const parts=k.split("|");
    return {k, day: parseInt(parts[2],10)};
  }).sort((a,b)=>a.day-b.day);

  if(sameRow.length){
    // removedDayより大きい最小、なければ小さい最大
    const right = sameRow.find(x=>x.day>removedDay);
    if(right) return right.k;
    return sameRow[sameRow.length-1].k;
  }

  // それ以外はキーをソートして最初のもの
  const all = Array.from(selectedKeySet).sort();
  return all[0] || null;
}

function toggleCellInSelection(cell){
  // 段違いは選択不可
  if(selectedCells.length>0){
    const baseType = selectedCells[0].dataset.rowType;
    if(cell.dataset.rowType !== baseType) return;
  }
  const k = cellKey(cell);
  const wasSelected = selectedKeySet.has(k);

  if(wasSelected){
    if(selectedKeySet.size===1){
      return; // 最後の1セルは解除しない
    }
    const removedActive = (activeCell && cellKey(activeCell)===k);
    selectedKeySet.delete(k);
    cell.classList.remove("cell-multi","cell-selected");

    if(removedActive){
      // 次のアクティブを決める（なるべく近いセル）
      const nextKey = pickNextActiveAfterRemoval(k);
      activeCell = nextKey ? getCellByKey(nextKey) : null;
    }
  }else{
    selectedKeySet.add(k);
    cell.classList.add("cell-multi");
    // 追加時はアクティブを動かさない（仕様）
  }

  // 表示と参照を同期
  rebuildSelectedCellsFromKeys();

  if(activeCell){
    activeCell.classList.add("cell-selected");
  }else{
    // 念のため：何か残っていれば最後をactive
    activeCell = selectedCells[selectedCells.length-1] || null;
    if(activeCell) activeCell.classList.add("cell-selected");
  }
}





function updateDragSelection(targetCell){
  if(!dragStartCell) return;
  const table = document.getElementById("scheduleTable");
  if(!table) return;

  const startRowType = dragStartCell.dataset.rowType;
  if(targetCell.dataset.rowType !== startRowType) return;

  const startIndex = parseInt(dragStartCell.dataset.rowIndex,10);
  const targetIndex = parseInt(targetCell.dataset.rowIndex,10);
  const startDay = parseInt(dragStartCell.dataset.day,10);
  const targetDay = parseInt(targetCell.dataset.day,10);

  const minRow = Math.min(startIndex, targetIndex);
  const maxRow = Math.max(startIndex, targetIndex);
  const minDay = Math.min(startDay, targetDay);
  const maxDay = Math.max(startDay, targetDay);

  // 矩形範囲のキー集合
  const rectKeys = new Set();
  table.querySelectorAll(`td[data-row-type="${startRowType}"]`).forEach(c=>{
    const rIdx = parseInt(c.dataset.rowIndex,10);
    const d = parseInt(c.dataset.day,10);
    if(rIdx>=minRow && rIdx<=maxRow && d>=minDay && d<=maxDay){
      rectKeys.add(cellKey(c));
    }
  });
  dragLastRectKeys = rectKeys;

  // Ctrlドラッグは「プレビュー」を表示し、確定は mouseup で行う
  if(dragAdditive && dragBaseSelectionKeys){
    clearDragPreview();

    const base = dragBaseSelectionKeys;
    if(dragRangeMode === "add"){
      // 追加候補（まだ選択されていないもの）
      const pending = new Set();
      rectKeys.forEach(k=>{ if(!base.has(k)) pending.add(k); });

      // 表示上は base ∪ pending を選択状態にする
      const display = new Set(base);
      pending.forEach(k=>display.add(k));
      applySelectionKeys(display, null);

      // pending をプレビュー強調
      pending.forEach(k=>{ const c=getCellByKey(k); if(c) c.classList.add("cell-preview-add"); });
      setDragHint("add", pending.size);
      return;
    }

    if(dragRangeMode === "remove"){
      // 解除候補（既に選択されているものだけ）
      const pending = new Set();
      rectKeys.forEach(k=>{ if(base.has(k)) pending.add(k); });

      // 表示上の選択は base のまま維持（消さない）
      applySelectionKeys(base, null);

      // pending をプレビュー強調（取り消し線）
      pending.forEach(k=>{ const c=getCellByKey(k); if(c) c.classList.add("cell-preview-remove"); });
      setDragHint("remove", pending.size);
      return;
    }
  }

  // 通常（置換/非Ctrl）
  // 範囲選択（置換）中は確定前プレビューとして青の破線を表示（確定は mouseup）
  if(!dragAdditive && dragRangeMode === "replace" && pointerDown){
    clearDragPreview();
    applySelectionKeys(rectKeys, null);
    // 始点セルはgetCellByKeyの揺れに関係なく必ず破線プレビューを付ける
    try{ dragStartCell.classList.add('cell-preview-add'); }catch(_){ }

    rectKeys.forEach(k=>{ const c=getCellByKey(k); if(c) c.classList.add('cell-preview-add'); });
    // ヒントは通常は出さない（必要ならここで setDragHint('add', rectKeys.size) 可能）
    return;
  }

  // 通常（置換/非Ctrl）は従来通り即時反映
  clearDragPreview();
  if(dragRangeMode === "replace"){
    applySelectionKeys(rectKeys, null);
  }else if(dragRangeMode === "add"){
    const next = new Set(selectedKeySet);
    rectKeys.forEach(k=>next.add(k));
    applySelectionKeys(next, targetCell);
  }else if(dragRangeMode === "remove"){
    const next = new Set(selectedKeySet);
    rectKeys.forEach(k=>next.delete(k));
    if(next.size===0) next.add(cellKey(dragStartCell));
    applySelectionKeys(next, targetCell);
  }
}




/* ===== Key handling ===== */

function getVisibleMaxDay(){
  try{
    const table = document.getElementById('scheduleTable');
    if(!table) return 31;
    const ths = table.querySelectorAll('thead th[data-day]');
    if(ths && ths.length){
      let max = 1;
      ths.forEach(th=>{
        const d = parseInt(th.dataset.day,10);
        if(!Number.isNaN(d) && d>max) max = d;
      });
      return max || 31;
    }
    const tds = table.querySelectorAll('td[data-day]');
    let max = 1;
    tds.forEach(td=>{
      const d = parseInt(td.dataset.day,10);
      if(!Number.isNaN(d) && d>max) max = d;
    });
    return max || 31;
  }catch(_){ return 31; }
}

function getRowIndexList(rowType){
  const table = document.getElementById('scheduleTable');
  if(!table) return [0];
  const set = new Set();
  table.querySelectorAll(`td[data-row-type="${rowType}"]`).forEach(c=>{
    const r = parseInt(c.dataset.rowIndex,10);
    if(!Number.isNaN(r)) set.add(r);
  });
  const arr = Array.from(set).sort((a,b)=>a-b);
  return arr.length ? arr : [0];
}

function getMaxRowIndexForRowType(rowType){
  const table = document.getElementById('scheduleTable');
  if(!table) return 0;
  let max = 0;
  table.querySelectorAll(`td[data-row-type="${rowType}"]`).forEach(c=>{
    const r = parseInt(c.dataset.rowIndex,10);
    if(r>max) max=r;
  });
  return max;
}
function getCellByRowDay(rowType, rowIndex, day){
  const table = document.getElementById('scheduleTable');
  if(!table) return null;
  return table.querySelector(`td[data-row-type="${rowType}"][data-row-index="${rowIndex}"][data-day="${day}"]`);
}
function buildRectKeys(anchorCell, focusCell){
  const table = document.getElementById('scheduleTable');
  if(!table || !anchorCell || !focusCell) return new Set();
  const rowType = anchorCell.dataset.rowType;
  if(focusCell.dataset.rowType !== rowType) return new Set();

  const aRow = parseInt(anchorCell.dataset.rowIndex,10);
  const fRow = parseInt(focusCell.dataset.rowIndex,10);
  const aDay = parseInt(anchorCell.dataset.day,10);
  const fDay = parseInt(focusCell.dataset.day,10);

  const minRow = Math.min(aRow, fRow);
  const maxRow = Math.max(aRow, fRow);
  const minDay = Math.min(aDay, fDay);
  const maxDay = Math.max(aDay, fDay);

  const keys = new Set();
  table.querySelectorAll(`td[data-row-type="${rowType}"]`).forEach(c=>{
    const r = parseInt(c.dataset.rowIndex,10);
    const d = parseInt(c.dataset.day,10);
    if(r>=minRow && r<=maxRow && d>=minDay && d<=maxDay){
      keys.add(cellKey(c));
    }
  });
  return keys;
}

function handleCellKeyDown(e, cell){
  const rowType=cell.dataset.rowType;
  if(rowType==="support") return;

  if(e.isComposing || e.key==="Process"){ showIMEAlert(); return; }

  
  // Shift+矢印：矩形範囲選択（アクティブを起点、循環なし）
  if(rowType!=="note" && e.shiftKey && !e.ctrlKey && !e.metaKey && (e.key==="ArrowLeft"||e.key==="ArrowRight"||e.key==="ArrowUp"||e.key==="ArrowDown")){
    // NumLock OFFテンキーは入力用途と衝突するので対象外
    if(typeof e.code==="string" && e.code.startsWith("Numpad")){
      // pass
    }else{
      e.preventDefault();
      const base = shiftAnchorCell || activeCell || cell;
      if(!base) return;
      if(!shiftAnchorCell) shiftAnchorCell = base;

      const rowType = base.dataset.rowType;
      const maxDay = getVisibleMaxDay();
      const maxRow = getMaxRowIndexForRowType(rowType);

      let r = parseInt((activeCell||cell).dataset.rowIndex,10);
      let d = parseInt((activeCell||cell).dataset.day,10);

      if(e.key==="ArrowLeft")  d = Math.max(1, d-1);
      if(e.key==="ArrowRight") d = Math.min(maxDay, d+1);
      if(e.key==="ArrowUp")    r = Math.max(0, r-1);
      if(e.key==="ArrowDown")  r = Math.min(maxRow, r+1);

      const next = getCellByRowDay(rowType, r, d);
      if(!next) return;

      const rectKeys = buildRectKeys(shiftAnchorCell, next);
      pointerDown = false; // committing扱い
      applySelectionKeys(rectKeys, next);
      try{ next.focus({preventScroll:true}); }catch(_){ try{ next.focus(); }catch(__){} }
      return;
    }
  }

// 十字キーでセル移動（Tab/Enter相当）。※循環（端から次/前の従業員へ）はしない
  // NumLock OFF のテンキーは e.key が ArrowUp 等になるため、code が Numpad の場合は移動に使わない（コード入力に回す）
  if(e.key==="ArrowLeft" || e.key==="ArrowRight" || e.key==="ArrowUp" || e.key==="ArrowDown"){
    if(typeof e.code==="string" && e.code.startsWith("Numpad")){
      // テンキー（NumLock OFF含む）は入力用途として扱う
    }else{
      e.preventDefault();
      if(selectedCells.length>1){ clearSelection(); setActiveCell(cell,false); }
      if(e.key==="ArrowLeft")  return moveFocusByDayNoWrap(cell, -1);
      if(e.key==="ArrowRight") return moveFocusByDayNoWrap(cell,  1);
      if(e.key==="ArrowUp")    return moveFocusByEmployeeNoWrap(cell, -1);
      if(e.key==="ArrowDown")  return moveFocusByEmployeeNoWrap(cell,  1);
    }
  }


  if(e.key==="Tab"){
    e.preventDefault();
    if(selectedCells.length>1){ clearSelection(); setActiveCell(cell,false); }
    moveFocusByDay(cell, e.shiftKey ? -1 : 1);
    return;
  }
  if(e.key==="Enter"){
    e.preventDefault();
    if(selectedCells.length>1){ clearSelection(); setActiveCell(cell,false); }
    moveFocusByEmployee(cell, e.shiftKey ? -1 : 1);
    return;
  }
  if(rowType!=="note" && (e.key==="Delete" || e.key==="Backspace")){
    e.preventDefault();
    handleDeleteKey();
    return;
  }
  if(rowType==="note") return;

  if(isCellReadonly(cell)){ e.preventDefault(); return; }

  // 数字入力は document 側のハンドラで一元処理（重複入力防止）

  if(/^[a-zA-Z]$/.test(e.key)){
    e.preventDefault();
    handleLetterInput(e.key.toUpperCase());
    return;
  }
}

function isCellReadonly(cell){
  const val=(cell.textContent||"").trim();
  return val.includes("ﾋ");
}

function moveFocusByDayNoWrap(cell, dir){
  const schedule=getCurrentSchedule();
  if(!schedule) return;
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
  const day=Number(cell.dataset.day);
  const rowType=cell.dataset.rowType;
  const empId=Number(cell.dataset.empId);

  const newDay=day+dir;
  if(newDay<1 || newDay>daysInMonth) return; // 循環しない

  focusCellByPosition(empId, rowType, newDay);
}

function moveFocusByEmployeeNoWrap(cell, dir){
  const schedule=getCurrentSchedule();
  if(!schedule) return;
  const day=Number(cell.dataset.day);
  const rowType=cell.dataset.rowType;
  const empId=Number(cell.dataset.empId);

  const employees=getEmployeesInStore(schedule,currentStoreId);
  const index=employees.findIndex(e=>e.id===empId);
  const newIndex=index+dir;
  if(newIndex<0 || newIndex>=employees.length) return; // 循環しない

  focusCellByPosition(employees[newIndex].id, rowType, day);
}

function moveFocusByDay(cell, dir){
  const schedule=getCurrentSchedule();
  if(!schedule) return;
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
  const day=Number(cell.dataset.day);
  const rowType=cell.dataset.rowType;
  const empId=Number(cell.dataset.empId);
  const employees=getEmployeesInStore(schedule,currentStoreId);
  const index=employees.findIndex(e=>e.id===empId);

  let newDay=day+dir;
  let newEmpIndex=index;

  if(newDay<1){
    newEmpIndex -= 1;
    if(newEmpIndex<0) newEmpIndex=employees.length-1;
    newDay=daysInMonth;
  }else if(newDay>daysInMonth){
    newEmpIndex += 1;
    if(newEmpIndex>=employees.length) newEmpIndex=0;
    newDay=1;
  }
  focusCellByPosition(employees[newEmpIndex].id, rowType, newDay);
}

function moveFocusByEmployee(cell, dir){
  const schedule=getCurrentSchedule();
  if(!schedule) return;
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
  const day=Number(cell.dataset.day);
  const rowType=cell.dataset.rowType;
  const empId=Number(cell.dataset.empId);

  const employees=getEmployeesInStore(schedule,currentStoreId);
  const index=employees.findIndex(e=>e.id===empId);

  let newEmpIndex=index+dir;
  let newDay=day;
  if(newEmpIndex<0){
    newEmpIndex=employees.length-1;
    newDay=Math.max(1, day-1);
  }else if(newEmpIndex>=employees.length){
    newEmpIndex=0;
    newDay=Math.min(daysInMonth, day+1);
  }
  focusCellByPosition(employees[newEmpIndex].id, rowType, newDay);
}

function focusCellByPosition(empId, rowType, day){
  const table=document.getElementById("scheduleTable");
  if(!table) return;
  const target=table.querySelector(`td[data-row-type="${rowType}"][data-emp-id="${empId}"][data-day="${day}"]`);
  if(target){ target.focus(); setActiveCell(target,false); }
}

function computeNextByDaySteps(schedule, storeId, empId, rowType, day, steps){
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
  const employees=getEmployeesInStore(schedule,storeId);
  let idx=employees.findIndex(e=>e.id===empId);
  let d=day;
  for(let i=0;i<steps;i++){
    d+=1;
    if(d>daysInMonth){
      d=1;
      idx+=1;
      if(idx>=employees.length) idx=0;
    }
  }
  return {empId:employees[idx].id, rowType, day:d};
}

/* ===== Input apply ===== */
function handleNumericInput(num){
  const schedule=getCurrentSchedule();
  if(!schedule||!activeCell) return;
  if(activeCell.dataset.rowType!=="main") return;

  const mode=getInputMode();
  const snapshot=schedule.masterSnapshot;

  if(mode==="leave"){
    const lc=snapshot.leaveCodes.find(l=>l.key===String(num));
    if(!lc) return;
    applyCodeToSelection({kind:"leave",code:lc.code}, schedule);
  }else{
    const wc=snapshot.workCodes.find(w=>w.key===String(num));
    if(!wc) return;
    if(isSupportMode && !(wc.code==="E"||wc.code==="F")) return;

    const ctxStoreId=getWorkContextStoreId();
    const ctxStore=snapshot.stores.find(s=>s.id===ctxStoreId);
    if(!ctxStore) return;
    const perStore=(ctxStore.storeCodes||[]).find(sc=>sc.code===wc.code);
    if(!perStore || perStore.type==="disabled") return;

    applyCodeToSelection({kind:"work",code:wc.code,shiftType:perStore.type}, schedule);
  }
}

function handleLetterInput(letter){
  const schedule=getCurrentSchedule();
  if(!schedule||!activeCell) return;
  if(activeCell.dataset.rowType!=="main") return;

  if(getInputMode()!=="work") return;
  const snapshot=schedule.masterSnapshot;
  const wc=snapshot.workCodes.find(w=>w.code===letter);
  if(!wc) return;

  const ctxStoreId=getWorkContextStoreId();
  const ctxStore=snapshot.stores.find(s=>s.id===ctxStoreId);
  if(!ctxStore) return;

  const perStore=(ctxStore.storeCodes||[]).find(sc=>sc.code===wc.code);
  if(!perStore || perStore.type==="disabled") return;

  applyCodeToSelection({kind:"work",code:wc.code,shiftType:perStore.type}, schedule);
}

function getWorkContextStoreId(){
  if(!isSupportMode) return currentStoreId;
  const id=Number(dom.supportStoreSelect.value);
  return id || currentStoreId;
}

function getSupportPrefix(snapshot){
  if(!isSupportMode) return "";
  const targetId=Number(dom.supportStoreSelect.value);
  const store=snapshot.stores.find(s=>s.id===targetId);
  return store?.shortName || "";
}

function applyCodeToSelection(codeInfo, schedule){
  if(!activeCell) return;
  const cells=(selectedCells.length>0?selectedCells.slice():[activeCell]);
  const rowType=activeCell.dataset.rowType;
  const filtered=cells.filter(c=>c.dataset.rowType===rowType);
  if(rowType!=="main") return;

  const isMulti=filtered.length>1;

  if(isMulti){
    // 複数セル入力時は選択を維持したまま再描画（選択解除＆フォーカス移動をしない）
    preserveSelectionOnRender=true;
    preserveSelectionKeys=new Set(selectedKeySet);
    preserveSelectionActiveKey=activeCell ? cellKey(activeCell) : null;
  }
  const isNight = (codeInfo.kind==="work" && codeInfo.shiftType==="night");
  const targetCells = isNight ? filtered.filter(c=>!isCellReadonly(c)) : filtered;
  if(targetCells.length===0) return;

  const snapshot=schedule.masterSnapshot;
  const prefix = (codeInfo.kind==="work") ? getSupportPrefix(snapshot) : "";

  if(codeInfo.kind==="leave"){
    targetCells.forEach(cell=>{
      const empId=Number(cell.dataset.empId);
      const day=Number(cell.dataset.day);
      // 夜勤ペアの孤立防止：上書き前に関連セルも掃除
      clearPairedNightIfNeeded(schedule,currentStoreId,empId,day);
      setAssignmentValue(schedule,currentStoreId,empId,"main",day,codeInfo.code);
    });
    if(!isMulti){
      const empId=Number(activeCell.dataset.empId);
      const day=Number(activeCell.dataset.day);
      pendingFocus=computeNextByDaySteps(schedule,currentStoreId,empId,"main",day,1);
    }
  }else if(codeInfo.kind==="work" && codeInfo.shiftType==="night"){
    applyNightCode(schedule, targetCells, prefix, codeInfo.code);
    if(!isMulti){
      const empId=Number(activeCell.dataset.empId);
      const day=Number(activeCell.dataset.day);
      pendingFocus=computeNextByDaySteps(schedule,currentStoreId,empId,"main",day,2);
    }
  }else{
    const value=prefix + codeInfo.code;
    targetCells.forEach(cell=>{
      const empId=Number(cell.dataset.empId);
      const day=Number(cell.dataset.day);
      // 夜勤ペアの孤立防止：上書き前に関連セルも掃除
      clearPairedNightIfNeeded(schedule,currentStoreId,empId,day);
      setAssignmentValue(schedule,currentStoreId,empId,"main",day,value);
    });
    if(!isMulti){
      const empId=Number(activeCell.dataset.empId);
      const day=Number(activeCell.dataset.day);
      pendingFocus=computeNextByDaySteps(schedule,currentStoreId,empId,"main",day,1);
    }
  }
  renderSchedule();
}

function applyNightCode(schedule, cells, prefix, codeLetter){
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
  const valueStart=prefix+codeLetter+"ﾃ";
  const valueEnd=prefix+codeLetter+"ﾋ";

  const grouped={};
  cells.forEach(cell=>{
    const empId=Number(cell.dataset.empId);
    const day=Number(cell.dataset.day);
    if(!grouped[empId]) grouped[empId]=[];
    grouped[empId].push(day);
  });

  Object.keys(grouped).forEach(empIdStr=>{
    const empId=Number(empIdStr);
    const days=grouped[empId].sort((a,b)=>a-b);

    for(let i=0;i<days.length;i+=2){
      const d=days[i];
      if(!d) continue;

      // 既存の夜勤ペアがあれば先に掃除（ﾋ孤立防止・上書き衝突回避）
      clearPairedNightIfNeeded(schedule,currentStoreId,empId,d);
      setAssignmentValue(schedule,currentStoreId,empId,"main",d,valueStart);

      if(d<daysInMonth){
        // 翌日セルが既存の夜勤開始(ﾃ)だった場合、そのペアも掃除してから上書き
        clearPairedNightIfNeeded(schedule,currentStoreId,empId,d+1);
        setAssignmentValue(schedule,currentStoreId,empId,"main",d+1,valueEnd);
      }else{
        const nextYm=nextMonthKey(currentYearMonth);
        const nextSchedule=appState.schedules[nextYm];

        if(nextSchedule && nextSchedule.byStore[currentStoreId] && nextSchedule.byStore[currentStoreId].assignments[empId]){
          // 次月1日が既存の夜勤明け/入りだった場合も孤立しないよう掃除してから上書き
          clearPairedNightIfNeeded(nextSchedule,currentStoreId,empId,1);
          nextSchedule.byStore[currentStoreId].assignments[empId].main[0]=valueEnd;
          markDirty();
        }else{
          if(!appState.pendingNightCarryovers[nextYm]) appState.pendingNightCarryovers[nextYm]={};
          if(!appState.pendingNightCarryovers[nextYm][currentStoreId]) appState.pendingNightCarryovers[nextYm][currentStoreId]={};
          if(!appState.pendingNightCarryovers[nextYm][currentStoreId][empId]) appState.pendingNightCarryovers[nextYm][currentStoreId][empId]={};
          appState.pendingNightCarryovers[nextYm][currentStoreId][empId]["1"]=valueEnd;
          markDirty();
        }
      }
    }
  });
}

/* ===== Delete / Backspace ===== */
function handleDeleteKey(){
  const schedule=getCurrentSchedule();
  if(!schedule) return;
  // 削除後もアクティブセルを維持するため、現在の位置を控えておく
  const focusAfterDelete = activeCell ? {
    empId: Number(activeCell.dataset.empId),
    rowType: activeCell.dataset.rowType,
    day: Number(activeCell.dataset.day)
  } : null;
  const cells=(selectedCells.length>0?selectedCells.slice():[activeCell]).filter(Boolean);
  if(cells.length===0) return;

  const isMultiSelection = (selectedKeySet && selectedKeySet.size>1) || (selectedCells && selectedCells.length>1);
  if(isMultiSelection){
    // 複数セル削除時は選択を維持したまま再描画（選択解除＆フォーカス移動をしない）
    preserveSelectionOnRender=true;
    preserveSelectionKeys=new Set(selectedKeySet);
    preserveSelectionActiveKey=activeCell ? cellKey(activeCell) : null;
  }

  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);

  cells.forEach(cell=>{
    const rowType=cell.dataset.rowType;
    if(rowType==="support") return;
    const empId=Number(cell.dataset.empId);
    const day=Number(cell.dataset.day);
    const storeId=currentStoreId;

    const val=getAssignmentValue(schedule,storeId,empId,rowType,day);

    if(rowType==="main" && val){
      const parsed=parseCellValue(val, schedule.masterSnapshot);
      if(parsed.isNight && parsed.nightPart==="start"){
        setAssignmentValue(schedule,storeId,empId,"main",day,"");
        if(day<daysInMonth){
          const nextVal=getAssignmentValue(schedule,storeId,empId,"main",day+1);
          if(nextVal.includes("ﾋ")) setAssignmentValue(schedule,storeId,empId,"main",day+1,"");
        }else{
          const nextYm=nextMonthKey(currentYearMonth);
          const nextSchedule=appState.schedules[nextYm];
          if(nextSchedule && nextSchedule.byStore[storeId] && nextSchedule.byStore[storeId].assignments[empId]){
            const nextVal=getAssignmentValue(nextSchedule,storeId,empId,"main",1);
            if(nextVal.includes("ﾋ")) setAssignmentValue(nextSchedule,storeId,empId,"main",1,"");
          }
          const p=appState.pendingNightCarryovers?.[nextYm]?.[storeId]?.[empId];
          if(p && p["1"]){ delete p["1"]; markDirty(); }
        }
      }else if(parsed.isNight && parsed.nightPart==="end"){
        setAssignmentValue(schedule,storeId,empId,"main",day,"");
        if(day>1){
          const prevVal=getAssignmentValue(schedule,storeId,empId,"main",day-1);
          if(prevVal.includes("ﾃ")) setAssignmentValue(schedule,storeId,empId,"main",day-1,"");
        }
      }else{
        setAssignmentValue(schedule,storeId,empId,"main",day,"");
      }
    }else{
      setAssignmentValue(schedule,storeId,empId,rowType,day,"");
    }
  });

  if(focusAfterDelete){ pendingFocus = focusAfterDelete; }
  renderSchedule();
}

/* ===== Cell render (color/readonly) ===== */
function renderCellValue(cell, value, schedule, currentStoreIdForTable){
  cell.textContent=value||"";
  cell.classList.remove("cell-bg-colored","cell-work-text","cell-readonly");
  cell.style.backgroundColor="";
  cell.style.color="";
  if(!value) return;
  if(cell.dataset.rowType!=="main") return;
  if(value.includes("ﾋ")) cell.classList.add("cell-readonly");

  const snapshot=schedule.masterSnapshot;
  const parsed=parseCellValue(value, snapshot);
  if(!parsed.workCode) return;

  const work=snapshot.workCodes.find(w=>w.code===parsed.workCode);
  if(!work) return;
  cell.classList.add("cell-work-text");
  cell.style.color=work.color;

  const contextStoreId=parsed.contextStoreId || currentStoreIdForTable;
  const ctxStore=snapshot.stores.find(s=>s.id===contextStoreId);
  const perStore=(ctxStore?.storeCodes||[]).find(sc=>sc.code===parsed.workCode);
  const isNight=perStore?.type==="night";

  if(!isNight || (isNight && parsed.nightPart==="start")){
    cell.style.backgroundColor=work.bgColor;
    cell.classList.add("cell-bg-colored");
  }
}

function parseCellValue(value, snapshot){
  let isNight=false, nightPart=null, workCode="", storePrefix="", contextStoreId=null;
  if(!value) return {isNight,nightPart,workCode,storePrefix,contextStoreId};

  const last=value[value.length-1];
  if(last==="ﾃ" || last==="ﾋ"){
    isNight=true;
    nightPart = (last==="ﾃ") ? "start" : "end";
    const core=value.slice(0,-1);
    if(core.length===1 && /^[A-Z]$/.test(core[0])){
      workCode=core[0];
    }else if(core.length>=2 && /^[A-Z]$/.test(core[core.length-1])){
      storePrefix=core[0];
      workCode=core[core.length-1];
    }
  }else{
    if(value.length>=2 && /^[A-Z]$/.test(value[1])){
      storePrefix=value[0];
      workCode=value[1];
    }else if(value.length===1 && /^[A-Z]$/.test(value[0])){
      workCode=value[0];
    }
  }
  if(storePrefix){
    const store=snapshot.stores.find(s=>s.shortName===storePrefix);
    if(store) contextStoreId=store.id;
  }
  return {isNight,nightPart,workCode,storePrefix,contextStoreId};
}

/* ===== Night pair cleanup helpers ===== */
// 夜勤入り(ﾃ)を消す/上書きする場合に、翌日の夜勤明け(ﾋ)が孤立しないよう必ず消す。
// 逆に夜勤明け(ﾋ)を消す/上書きする場合も、前日の夜勤入り(ﾃ)を消す。
function clearPairedNightIfNeeded(schedule, storeId, empId, day){
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);
  const val=getAssignmentValue(schedule,storeId,empId,"main",day);
  if(!val) return;
  const parsed=parseCellValue(val, schedule.masterSnapshot);
  if(!parsed.isNight) return;

  if(parsed.nightPart==="start"){
    // 同月内：翌日がﾋなら消す（コードが一致しなくても孤立回避のため消す）
    if(day < daysInMonth){
      const nextVal=getAssignmentValue(schedule,storeId,empId,"main",day+1);
      if(nextVal && nextVal.includes("ﾋ")){
        setAssignmentValue(schedule,storeId,empId,"main",day+1,"");
      }
    }else{
      // 月末：翌月1日のﾋ（または保留キャリー）を消す
      const nextYm=nextMonthKey(currentYearMonth);
      const nextSchedule=appState.schedules[nextYm];
      if(nextSchedule && nextSchedule.byStore?.[storeId]?.assignments?.[empId]){
        const nextVal=getAssignmentValue(nextSchedule,storeId,empId,"main",1);
        if(nextVal && nextVal.includes("ﾋ")){
          setAssignmentValue(nextSchedule,storeId,empId,"main",1,"");
        }
      }
      const p=appState.pendingNightCarryovers?.[nextYm]?.[storeId]?.[empId];
      if(p && p["1"]){
        delete p["1"];
        markDirty();
      }
    }
  }else if(parsed.nightPart==="end"){
    // 同月内：前日がﾃなら消す
    if(day > 1){
      const prevVal=getAssignmentValue(schedule,storeId,empId,"main",day-1);
      if(prevVal && prevVal.includes("ﾃ")){
        setAssignmentValue(schedule,storeId,empId,"main",day-1,"");
      }
    }else{
      // 月初：前月末のﾃを消す（存在する場合）
      const prevYm=prevMonthKey(currentYearMonth);
      const prevSchedule=appState.schedules[prevYm];
      if(prevSchedule && prevSchedule.byStore?.[storeId]?.assignments?.[empId]){
        const prevDays=getDaysInMonth(prevSchedule.year, prevSchedule.month);
        const prevVal=getAssignmentValue(prevSchedule,storeId,empId,"main",prevDays);
        if(prevVal && prevVal.includes("ﾃ")){
          setAssignmentValue(prevSchedule,storeId,empId,"main",prevDays,"");
        }
      }
    }
  }
}

function prevMonthKey(ym){
  const [yStr,mStr]=ym.split("-");
  let y=parseInt(yStr,10), m=parseInt(mStr,10);
  m-=1; if(m<1){ m=12; y-=1; }
  return `${y.toString().padStart(4,"0")}-${m.toString().padStart(2,"0")}`;
}


/* ===== Night consistency check (current month only) ===== */
function validateNightShiftsForCurrent(){
  dom.alertList.innerHTML="";
  if(!currentSchedule || !currentStoreSchedule) return;

  const schedule=currentSchedule;
  const storeId=currentStoreId;
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);

  const errors=[];
  const employees=getEmployeesInStore(schedule, storeId);

  const table=document.getElementById("scheduleTable");
  if(table) table.querySelectorAll("td").forEach(td=>td.classList.remove("cell-error"));

  employees.forEach(emp=>{
    for(let d=1; d<=daysInMonth; d++){
      const v=getAssignmentValue(schedule,storeId,emp.id,"main",d);
      if(!v) continue;
      const parsed=parseCellValue(v, schedule.masterSnapshot);
      if(!parsed.isNight) continue;

      const cell = table?.querySelector(`td[data-row-type="main"][data-emp-id="${emp.id}"][data-day="${d}"]`);
      if(parsed.nightPart==="start"){
        if(d===daysInMonth) continue;
        const nextVal=getAssignmentValue(schedule,storeId,emp.id,"main",d+1);
        if(!nextVal || !nextVal.includes("ﾋ")){
          if(cell) cell.classList.add("cell-error");
          errors.push(`${emp.lastName}${emp.firstName}：${d}日「夜勤入り」の翌日が「夜勤明け」になっていません。`);
        }
      }else if(parsed.nightPart==="end"){
        if(d===1) continue;
        const prevVal=getAssignmentValue(schedule,storeId,emp.id,"main",d-1);
        if(!prevVal || !prevVal.includes("ﾃ")){
          if(cell) cell.classList.add("cell-error");
          errors.push(`${emp.lastName}${emp.firstName}：${d}日「夜勤明け」の前日が「夜勤入り」になっていません。`);
        }
      }
    }
  });

  if(errors.length===0){
    const item=document.createElement("div");
    item.className="alert-item good";
    item.textContent="夜勤コードの不整合は検出されませんでした。";
    dom.alertList.appendChild(item);
  }else{
    errors.forEach(msg=>{
      const item=document.createElement("div");
      item.className="alert-item";
      item.textContent=msg;
      dom.alertList.appendChild(item);
    });
  }
}

/* ===== Support info (read only) ===== */
function getSupporterFor(schedule, targetStoreId, codeLetter, day){
  const snapshot=schedule.masterSnapshot;
  const targetStore=snapshot.stores.find(s=>s.id===targetStoreId);
  if(!targetStore || !targetStore.shortName) return "";
  const shortName=targetStore.shortName;

  for(const storeIdStr of Object.keys(schedule.byStore)){
    const storeId=Number(storeIdStr);
    const storeData=schedule.byStore[storeId];
    if(!storeData) continue;
    const assignments=storeData.assignments;
    for(const empIdStr of Object.keys(assignments)){
      const empId=Number(empIdStr);
      const val=assignments[empId].main[day-1] || "";
      if(!val) continue;
      const parsed=parseCellValue(val, snapshot);
      if(parsed.storePrefix===shortName && parsed.workCode===codeLetter){
        const emp=snapshot.employees.find(e=>e.id===empId);
        if(emp) return emp.lastName;
      }
    }
  }
  return "";
}


function getLatestYearMonthKey(){
  const keys=Object.keys(appState?.schedules || {});
  if(keys.length===0) return null;
  keys.sort(); // YYYY-MM なので辞書順でOK
  return keys[keys.length-1];
}
function isCurrentYearMonthLatest(){
  const latest=getLatestYearMonthKey();
  return !!latest && latest===currentYearMonth;
}


function updateSettingsScopeUI(){ /* v84: scope UI removed */ }


function getSettingsMasterTarget(){
  if(settingsScope==="current"){
    const schedule=getCurrentSchedule();
    return schedule ? schedule.masterSnapshot : null;
  }
  return appState.templateMaster;
}

/* ===== Settings UI ===== */
function populateSettingsFromState(){
  const master=getSettingsMasterTarget();
  if(!master){ alert('この年月の勤務表が未作成です。'); return; }

  dom.storeTableBody.innerHTML="";
  master.stores.slice().sort((a,b)=>(a.order||0)-(b.order||0)).forEach(store=>addStoreRow(store));
  updateStoreReorderButtons();

  dom.employeeTableBody.innerHTML="";
  master.employees.forEach(emp=>addEmployeeRow(emp));
  refreshEmployeeStoreOptionsFromSettings();
  refreshEmployeeOrderPanel();

  dom.workCodeTableBody.innerHTML="";
  master.workCodes.forEach(wc=>{
    const tr=document.createElement("tr");
    tr.dataset.code=wc.code;
    tr.innerHTML = `
      <td style="font-weight:800;">${wc.code}</td>
      <td><input type="number" min="0" max="9" value="${escapeHtml(wc.key||"")}"></td>
      <td><input type="color" value="${escapeHtml(wc.color||"#000000")}"></td>
      <td><input type="color" value="${escapeHtml(wc.bgColor||"#ffffff")}"></td>`;
    dom.workCodeTableBody.appendChild(tr);
  });

  dom.leaveCodeTableBody.innerHTML="";
  master.leaveCodes.forEach(lc=>addLeaveCodeRow(lc));

  // Holidays
  if(!master.holidays) master.holidays={};
  if(dom.holidayYearInput){
    const y=(getCurrentSchedule()?.year) || (currentYearMonth?Number(currentYearMonth.split("-")[0]):null) || new Date().getFullYear();
    dom.holidayYearInput.value = String(y);
  }
  renderHolidaySettingsTable();
}

function saveSettingsFromUI(){
  const master=getSettingsMasterTarget();
  if(!master){ alert('この年月の勤務表が未作成です。'); return; }

  const stores=[];
  const storeIdMap = new Map(); // 仮ID(負数) -> 実ID
  const rows=Array.from(dom.storeTableBody.querySelectorAll("tr"));
  rows.forEach((tr,idx)=>{
    let id=Number(tr.dataset.id);
    if(!Number.isFinite(id)) id=0;
    if(id<=0){
      const old=id;
      id = master.nextStoreId++;
      storeIdMap.set(String(old), String(id));
    }
    const name=tr.querySelector(".store-name").value.trim();
    const shortName=tr.querySelector(".store-short").value.trim();
    const enabled=tr.querySelector(".store-enabled").checked;

    const storeCodes=[];
    tr.querySelectorAll(".store-code-setting").forEach(sel=>{
      storeCodes.push({code:sel.dataset.code, type:sel.value});
    });
    stores.push({id,name,shortName,enabled,order:idx+1,storeCodes});
  });

  const nameSet=new Set();
  const shortSet=new Set();
  for(const s of stores){
    if(!s.name){ alert("店舗名が空の行があります。"); return; }
    if(!s.shortName){ alert("店舗略称が空の行があります。"); return; }
    if(nameSet.has(s.name)){ alert("店舗名が重複しています: "+s.name); return; }
    if(shortSet.has(s.shortName)){ alert("店舗略称が重複しています: "+s.shortName); return; }
    nameSet.add(s.name); shortSet.add(s.shortName);
  }
  master.stores=stores;
  master.nextStoreId=stores.reduce((mx,s)=>Math.max(mx,s.id),0)+1;

  const employees=[];
  Array.from(dom.employeeTableBody.querySelectorAll("tr")).forEach(tr=>{
    // data-id が欠けている場合でも、連番で補完して必ずIDを持たせる
    let idRaw = (tr.getAttribute("data-id") || "").trim();
    let idNum = Number(idRaw);
    const id = Number.isFinite(idNum) && idNum > 0 ? idNum : (master.nextEmployeeId++);

    const lastName=tr.querySelector(".emp-last").value.trim();
    const firstName=tr.querySelector(".emp-first").value.trim();

    const storeSel = tr.querySelector(".emp-store");
    const storeIdStr = ((storeSel && storeSel.value) ? String(storeSel.value) : "").trim();
    // "__none__" は未所属
    let storeId = (!storeIdStr || storeIdStr==="__none__") ? null : Number(storeIdStr);
    if(storeId!=null && storeId<=0 && storeIdMap.has(String(storeId))){
      storeId = Number(storeIdMap.get(String(storeId)));
    }
    if(!Number.isFinite(storeId)) storeId = null;

    // 店舗内順：未所属は常に0。所属ありは数値以外/空は0として後段で正規化
    const orderStr = (tr.querySelector(".emp-order")?.value ?? "").trim();
    let orderVal = Number(orderStr);
    if(!Number.isFinite(orderVal)) orderVal = 0;
    if(storeId==null) orderVal = 0;

    const orderInStore = Math.max(0, orderVal) * 10; // 表示(1刻み)→内部(10刻み)
    employees.push({id,lastName,firstName,storeId,orderInStore});
  });


  const fullSet=new Set();
  for(const e of employees){
    if(!e.lastName || !e.firstName){ alert("従業員の姓/名が空の行があります。"); return; }
    const k=e.lastName+" "+e.firstName;
    if(fullSet.has(k)){ alert("従業員の同姓同名が重複しています: "+k); return; }
    fullSet.add(k);
  }
  master.employees=employees;
  master.nextEmployeeId=employees.reduce((mx,e)=>Math.max(mx,e.id),0)+1;

  Array.from(dom.workCodeTableBody.querySelectorAll("tr")).forEach(tr=>{
    const code=tr.dataset.code;
    const inputs=tr.querySelectorAll("input");
    const key=inputs[0].value.trim();
    const color=inputs[1].value;
    const bgColor=inputs[2].value;
    const wc=master.workCodes.find(w=>w.code===code);
    if(!wc) return;
    wc.key=key; wc.color=color; wc.bgColor=bgColor;
  });

  const leaveCodes=[];
  Array.from(dom.leaveCodeTableBody.querySelectorAll("tr")).forEach(tr=>{
    const code=tr.querySelector(".leave-code").value.trim();
    const key=tr.querySelector(".leave-key").value.trim();
    if(!code) return;
    leaveCodes.push({code,key});
  });
  master.leaveCodes=leaveCodes;

  // Holidays
  const holidays={};
  Array.from(dom.holidayTableBody?.querySelectorAll("tr")||[]).forEach(tr=>{
    const dateEl = tr.querySelector(".hol-date-display");
    let date = (dateEl?.dataset?.iso || "").trim();
    if(!date){
      // fallback: try parse from visible "YYYY/MM/DD" at start
      const v = (dateEl?.value || "").trim();
      const m = v.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
      if(m) date = `${m[1]}-${m[2]}-${m[3]}`;
    }
    const name=tr.querySelector(".hol-name")?.value?.trim()||"";
    if(!date) return;
    holidays[date]={n:name};
  });
  master.holidays = holidays;


  // この月（勤務表スナップショット）を編集している場合：勤務表データ構造（byStore.assignments）をスナップショットに合わせて補正する
  const schedule=getCurrentSchedule();
  if(schedule){
    reconcileScheduleDataWithSnapshot(schedule);

    // 既存の最新勤務表を編集している場合のみ、チェックONでテンプレ（今後の新規作成用）にも反映
    const propagateCb=document.getElementById('propagateToTemplateCheckbox');
    const canPropagate=isCurrentYearMonthLatest();
    if(canPropagate && propagateCb && propagateCb.checked){
      appState.templateMaster = deepClone(schedule.masterSnapshot);
    }
  }

}


function setStoreCodeDot(dotEl, mode){
  dotEl.classList.remove("dot-day","dot-night","dot-disabled");
  if(mode==="night") dotEl.classList.add("dot-night");
  else if(mode==="disabled") dotEl.classList.add("dot-disabled");
  else dotEl.classList.add("dot-day");
}

function addStoreRow(store){
  const master=getSettingsMasterTarget() || appState.templateMaster;
  const workCodes=master.workCodes;
  const tr=document.createElement("tr");
  tr.dataset.id = store ? store.id : String(settingsTempStoreId--);

  const tdMove=document.createElement("td");
  tdMove.className="reorder-cell";
  tdMove.style.whiteSpace="nowrap";
  const up=document.createElement("button"); up.textContent="▲"; up.className="icon reorder-btn";
  up.addEventListener("click",()=>moveRow(tr,-1));
  const down=document.createElement("button"); down.textContent="▼"; down.className="icon reorder-btn";
  down.addEventListener("click",()=>moveRow(tr,1));
  tdMove.appendChild(up); tdMove.appendChild(down);

  const tdName=document.createElement("td");
  const inputName=document.createElement("input");
  inputName.type="text"; inputName.className="store-name"; inputName.value=store?store.name:"";
  tdName.appendChild(inputName);

  const tdShort=document.createElement("td");
  const inputShort=document.createElement("input");
  inputShort.type="text"; inputShort.maxLength=2; inputShort.className="store-short"; inputShort.value=store?(store.shortName||""):"";
  inputShort.addEventListener("input", ()=>refreshEmployeeStoreOptionsFromSettings());
  tdShort.appendChild(inputShort);

  const tdEnabled=document.createElement("td");
  const inputEnabled=document.createElement("input");
  inputEnabled.type="checkbox"; inputEnabled.className="store-enabled"; inputEnabled.checked=store?!!store.enabled:true;
  tdEnabled.appendChild(inputEnabled);

  const tdCodes=document.createElement("td");
  const wrap=document.createElement("div"); wrap.className="settings-inline-controls";
  workCodes.forEach(wc=>{
    const pill=document.createElement("span"); pill.className="pill";
    const codeLabel=document.createElement("span"); codeLabel.style.fontWeight="800"; codeLabel.textContent=wc.code;
    const dot=document.createElement("span"); dot.className="status-dot";
    pill.appendChild(codeLabel);
    pill.appendChild(dot);
    const sel=document.createElement("select"); sel.className="store-code-setting"; sel.dataset.code=wc.code;
    ["day","night","disabled"].forEach(mode=>{
      const opt=document.createElement("option");
      opt.value=mode;
      opt.textContent = mode==="day" ? "日勤" : mode==="night" ? "夜勤" : "無効";
      sel.appendChild(opt);
    });
    let initial="day";
    if(store){
      if(store.storeCodes && Array.isArray(store.storeCodes) && store.storeCodes.length>0){
        const sc=store.storeCodes.find(x=>x.code===wc.code);
        initial=sc ? (sc.type||"disabled") : "day";
      }else{
        initial="day";
      }
    }
    sel.value=initial;
    setStoreCodeDot(dot, initial);
    sel.addEventListener("change", ()=>setStoreCodeDot(dot, sel.value));
    pill.appendChild(sel);
    wrap.appendChild(pill);
  });
  tdCodes.appendChild(wrap);

  const tdDelete=document.createElement("td");
  const btnDel=document.createElement("button"); btnDel.className="btn-delete"; btnDel.textContent="削除";
  btnDel.addEventListener("click",()=>{
    const storeId=store?store.id:null;
    if(storeId && storeHasScheduleData(storeId)){
      alert("この店舗には勤務表データが存在するため削除できません。無効化してください。");
      return;
    }
    tr.remove();
    refreshEmployeeStoreOptionsFromSettings();
    refreshEmployeeOrderPanel();
    refreshEmployeeOrderPanel();
  });
  tdDelete.appendChild(btnDel);

  tr.appendChild(tdMove);
  tr.appendChild(tdName);
  tr.appendChild(tdShort);
  tr.appendChild(tdEnabled);
  tr.appendChild(tdCodes);
  tr.appendChild(tdDelete);

  dom.storeTableBody.appendChild(tr);
  updateStoreReorderButtons();
  // 追加直後でも従業員タブの所属店舗候補に反映する
  refreshEmployeeStoreOptionsFromSettings();
  refreshEmployeeStoreOptionsFromSettings();
}

function updateStoreReorderButtons(){
  try{
    const tbody=dom.storeTableBody;
    if(!tbody) return;
    const rows=Array.from(tbody.querySelectorAll("tr"));
    rows.forEach((tr,i)=>{
      const btns=tr.querySelectorAll("button.reorder-btn");
      if(btns.length<2) return;
      const up=btns[0], down=btns[1];
      const isFirst=(i===0);
      const isLast=(i===rows.length-1);
      up.disabled=isFirst;
      down.disabled=isLast;
      up.setAttribute("aria-disabled", isFirst ? "true":"false");
      down.setAttribute("aria-disabled", isLast ? "true":"false");
    });
  }catch(e){
    console.warn("[settings] updateStoreReorderButtons failed:", e);
  }
}

function moveRow(tr, dir){
  const tbody=tr.parentElement;
  const rows=Array.from(tbody.children);
  const idx=rows.indexOf(tr);
  const newIdx=idx+dir;
  if(newIdx<0 || newIdx>=rows.length) return;
  if(dir<0) tbody.insertBefore(tr, rows[newIdx]);
  else tbody.insertBefore(rows[newIdx], tr);
  updateStoreReorderButtons();
}

function storeHasScheduleData(storeId){
  for(const ym of Object.keys(appState.schedules)){
    const sch=appState.schedules[ym];
    if(sch.byStore && sch.byStore[storeId]) return true;
  }
  return false;
}


function getSettingsStoresSorted(master){
  return (master.stores||[]).slice().sort((a,b)=>(a.order||0)-(b.order||0));
}
function fillEmployeeStoreSelect(selectEl, master, selectedId){
  const storesSorted=getSettingsStoresSorted(master);
  const enabledStores=storesSorted.filter(s=>s && s.enabled!==false);
  const current=selectedId?String(selectedId):"";
  selectEl.innerHTML="";
  const opt0=document.createElement("option");
  opt0.value=""; opt0.textContent="未所属";
  selectEl.appendChild(opt0);

  // enabled stores
  enabledStores.forEach(s=>{
    const opt=document.createElement("option");
    opt.value=String(s.id);
    opt.textContent=s.name;
    selectEl.appendChild(opt);
  });

  // if current points to disabled store, keep it visible but not selectable
  if(current && !enabledStores.some(s=>String(s.id)===current)){
    const ds=storesSorted.find(s=>String(s.id)===current);
    if(ds){
      const opt=document.createElement("option");
      opt.value=current;
      opt.textContent=(ds.name||"") + "（無効）";
      opt.disabled=true;
      selectEl.appendChild(opt);
    }
  }
  selectEl.value=current;
}
function refreshEmployeeStoreOptionsFromSettings(){
  // 設定モーダル内では「店舗タブの編集中DOM」を単一の真実として扱う（未保存追加も即反映）
  const stores = Array.from(dom.storeTableBody.querySelectorAll("tr")).map(tr=>{
    return {
      id: tr.dataset.id || "",
      name: (tr.querySelector(".store-name")?.value || "").trim(),
      shortName: (tr.querySelector(".store-short")?.value || "").trim(),
      enabled: !!tr.querySelector(".store-enabled")?.checked
    };
  });

  dom.employeeTableBody.querySelectorAll("select.emp-store").forEach(sel=>{
    const current = sel.dataset.desired || sel.value;
    sel.innerHTML = "";

    // 未所属
    const opt0=document.createElement("option");
    opt0.value="__none__";
    opt0.textContent="（未所属）";
    sel.appendChild(opt0);

    // 有効店舗のみ選択肢にする
    stores.filter(s=>s.enabled).forEach(s=>{
      if(!s.id) return;
      const opt=document.createElement("option");
      opt.value = String(s.id);
      opt.textContent = s.name || s.shortName || `店舗${s.id}`;
      sel.appendChild(opt);
    });

    // 現在値が無効/存在しない場合：表示だけ残す（選択不可）で気づけるようにする
    if(current && !Array.from(sel.options).some(o=>o.value===current)){
      const s = stores.find(x=>String(x.id)===String(current));
      const opt=document.createElement("option");
      opt.value = String(current);
      opt.textContent = s ? `（無効）${s.name||s.shortName||current}` : `（未存在）${current}`;
      opt.disabled = true;
      sel.insertBefore(opt, sel.firstChild.nextSibling); // （未所属）の次
      sel.value = String(current);
    }else{
      sel.value = current || "";
    }
  });
}




function applyEmployeeLeftFilter(){
  // 左表の「未所属のみ表示」トグルに応じて行を表示/非表示
  if(!dom.employeeTableBody) return;
  const on = !!dom.employeeUnassignedOnlyToggle && dom.employeeUnassignedOnlyToggle.classList.contains("on");
  let visibleCount = 0;
  dom.employeeTableBody.querySelectorAll("tr").forEach(tr=>{
    const sel = tr.querySelector("select.emp-store");
    const v = sel ? String(sel.value||"").trim() : "";
    const assigned = !!(v && v!=="__none__");
    const show = !(on && assigned);
    tr.style.display = show ? "" : "none";
    if(show) visibleCount++;
  });
  if(dom.employeeUnassignedHint){
    dom.employeeUnassignedHint.style.display = (on && visibleCount===0) ? "" : "none";
  }
}


function refreshEmployeeOrderPanel(){
  if(!dom.employeeOrderStoreSelect || !dom.employeeOrderList) return;

  // 店舗一覧（設定モーダル内DOM）
  const stores = Array.from(dom.storeTableBody.querySelectorAll("tr")).map(tr=>{
    return {
      id: tr.dataset.id || "",
      name: (tr.querySelector(".store-name")?.value || "").trim(),
      shortName: (tr.querySelector(".store-short")?.value || "").trim(),
      enabled: !!tr.querySelector(".store-enabled")?.checked
    };
  }).filter(s=>s.id);

  // セレクト更新（有効店舗 + 末尾に（未所属））
  const sel = dom.employeeOrderStoreSelect;
  const prev = sel.value;
  sel.innerHTML = "";

  stores.filter(s=>s.enabled).forEach(s=>{
    const opt=document.createElement("option");
    opt.value = String(s.id);
    opt.textContent = s.name || s.shortName || `店舗${s.id}`;
    sel.appendChild(opt);
  });

  // 末尾に（未所属）を追加（未所属従業員の確認用）
  {
    const opt=document.createElement("option");
    opt.value = "__none__";
    opt.textContent = "（未所属）";
    sel.appendChild(opt);
  }

  // 選択維持（なければ先頭）
  if(prev && Array.from(sel.options).some(o=>o.value===prev)) sel.value = prev;
  else if(sel.options.length>0) sel.value = sel.options[0].value;

  const storeId = sel.value || "";

  // 従業員（設定モーダル内DOM）
  const employees = Array.from(dom.employeeTableBody.querySelectorAll("tr")).map(tr=>{
    const id = tr.dataset.id || "";
    const last = (tr.querySelector(".emp-last")?.value || "").trim();
    const first = (tr.querySelector(".emp-first")?.value || "").trim();
    const selStore = tr.querySelector("select.emp-store");
    const store = (selStore?.dataset?.desired || selStore?.value || "");
    const order = Number(tr.querySelector(".emp-order")?.value || "0") || 0;
    return {id,last,first,storeId:store,order};
  });

  const list = dom.employeeOrderList;
  list.innerHTML = "";

  if(!storeId && storeId!=="__none__"){
    const empty=document.createElement("div");
    empty.className="text-muted";
    empty.style.fontSize="12px";
    empty.textContent="店舗がありません（店舗タブで追加してください）";
    list.appendChild(empty);
    return;
  }

  const storeName = (storeId==="__none__")
    ? "（未所属）"
    : ((stores.find(s=>String(s.id)===String(storeId))?.name) || "");
  const items = employees
    .filter(e=>{
      if(String(storeId)==="__none__") return (!e.storeId || String(e.storeId)==="__none__");
      return String(e.storeId||"")===String(storeId);
    })
    .sort((a,b)=> (a.order||0)-(b.order||0) || (a.last+a.first).localeCompare(b.last+b.first,'ja'));

  // DnD用：空でもドロップできるようにゾーンを描画
  const dropzone=document.createElement("div");
  dropzone.className="employee-order-dropzone";
  dropzone.dataset.storeId = String(storeId);
  list.appendChild(dropzone);

  if(items.length===0){
    const empty=document.createElement("div");
    empty.className="text-muted";
    empty.style.fontSize="12px";
    empty.textContent = storeName ? `${storeName} に所属する従業員がいません（左からドラッグで追加）` : "所属する従業員がいません";
    dropzone.appendChild(empty);
    // ワイヤリング（空でも受け取る）
    ensureEmployeeOrderDnDInitialized();
    return;
  }

  items.forEach((e,idx)=>{
    const row=document.createElement("div");
    row.className="employee-order-item";
    row.draggable = true;
    row.dataset.empId = String(e.id);
    row.dataset.storeId = String(storeId);

    const left=document.createElement("div");
    left.style.display="flex";
    left.style.alignItems="center";
    left.style.gap="10px";

    const handle=document.createElement("div");
    handle.className="drag-handle";
    handle.title="ドラッグで並び替え / 左からドラッグで追加";
    handle.innerHTML = "⋮⋮";

    const name=document.createElement("div");
    name.innerHTML = `<div class="name">${escapeHtml(e.last)}${e.first?(" "+escapeHtml(e.first)):""}</div><div class="meta">順番: ${e.order||0}</div>`;

    left.appendChild(handle);
    left.appendChild(name);

    const right=document.createElement("div");
    right.className="text-muted";
    right.style.fontSize="11px";
    right.textContent = "ドラッグで並び替え";

    row.appendChild(left);
    row.appendChild(right);
    dropzone.appendChild(row);
  });

  // ワイヤリング（描画のたびに必要だが多重登録は防ぐ）
  ensureEmployeeOrderDnDInitialized();
}


// ===== 従業員：店舗別ドラッグ＆ドロップ並び替え（設定モーダル内） =====
let __empDnD = null;
function ensureEmployeeOrderDnDInitialized(){
  if(__empDnD) return;
  __empDnD = {
    active:false,
    from:null,        // "left" | "right"
    empId:null,
    fromStoreId:null,
    placeholder:null,
    placeholderIndex:null,
    moved:false,
    startTs:0,
    sourceEl:null,
    sourceStyle:null
  };

  // ===== drag UX helpers =====
  function setDraggingUi(isOn){
    // NOTE: ネイティブDnD中はCSS cursorが効きにくいことがあるため
    // dropEffect と併用する。
    try{
      document.documentElement.classList.toggle("is-dragging", !!isOn);
    }catch(_e){}
    try{
      // 右側のドロップ先を視覚的に示す（CSS: .drop-allowed）
      dom.employeeOrderList?.classList.toggle("drop-allowed", !!isOn);
      dom.employeeOrderList?.querySelector(".employee-order-dropzone")?.classList.toggle("drop-allowed", !!isOn);
    }catch(_e){}
  }

  function isInEmployeeOrderArea(target){
    if(!target) return false;
    try{
      return !!(target.closest && target.closest("#employeeOrderList, .employee-order-dropzone, .employee-order-item, .employee-order-placeholder"));
    }catch(_e){ return false; }
  }

  function setDropOver(isOver){
    try{
      const dz = dom.employeeOrderList?.querySelector(".employee-order-dropzone");
      if(!dz) return;
      dz.classList.toggle("drop-over", !!isOver);
      dom.employeeOrderList.classList.toggle("drop-over", !!isOver);
    }catch(_e){}
  }
  window.setDropOver = setDropOver;


  // 左テーブル：行ドラッグ開始
  // ---- left table pointer tracking (v91): remember the row the user actually grabbed (and log it) ----
  if(!window.__empDnDLeftPointerInstalled){
    window.__empDnDLeftPointerInstalled = true;
    window.__empDnDLastLeftPointer = { id:null, ts:0, x:0, y:0, tag:"", cls:"" };
    employeeTableBody.addEventListener("mousedown", (ev)=>{
      try{
        const interactive = ev.target && ev.target.closest ? ev.target.closest("input,select,textarea,button,a,label") : null;
        if(interactive) return;

        const trFromTarget = ev.target && ev.target.closest ? ev.target.closest("tr[data-id]") : null;

        let trFromPoint = null;
        try{
          const elAtPoint = document.elementFromPoint(ev.clientX||0, ev.clientY||0);
          trFromPoint = elAtPoint && elAtPoint.closest ? elAtPoint.closest("tr[data-id]") : null;
        }catch(_e){}

        // Prefer the row under the pointer (more reliable than event.target in some browsers/layouts)
        const tr = trFromPoint || trFromTarget;
        if(!tr) return;

        const id = tr.getAttribute("data-id");
        window.__empDnDLastLeftPointer = {
          id: id ? String(id) : null,
          ts: Date.now(),
          x: ev.clientX||0,
          y: ev.clientY||0,
          tag: String(ev.target && ev.target.tagName || ""),
          cls: String(ev.target && ev.target.className || "")
        };
        try{
          console.log("[settings-dnd] left mousedown", {
            rowId: window.__empDnDLastLeftPointer.id,
            // debug: what did we think was the row from each method?
            rowIdFromTarget: trFromTarget ? String(trFromTarget.getAttribute("data-id")||"") : null,
            rowIdFromPoint: trFromPoint ? String(trFromPoint.getAttribute("data-id")||"") : null,
            rowIndexInView: (function(){
              const tr = trFromPoint || trFromTarget;
              if(!tr || !dom || !dom.employeeTableBody) return null;
              const rows = Array.from(dom.employeeTableBody.querySelectorAll("tr[data-id]"));
              const i = rows.indexOf(tr);
              return i>=0 ? (i+1) : null;
            })(),
            nameAtPointer: (function(){
              const tr = trFromPoint || trFromTarget;
              if(!tr) return null;
              const last = (tr.querySelector(".emp-last")?.value || "").trim();
              const first = (tr.querySelector(".emp-first")?.value || "").trim();
              const label = (last || first) ? (last + " " + first).trim() : "(空欄)";
              return label;
            })(),
            pickedBy: trFromPoint ? "point" : "target",
            targetTag: window.__empDnDLastLeftPointer.tag,
            targetCls: window.__empDnDLastLeftPointer.cls,
            x: window.__empDnDLastLeftPointer.x,
            y: window.__empDnDLastLeftPointer.y
          });
        }catch(_e){}
      }catch(_e){}
    }, true);
  }

employeeTableBody.addEventListener("dragstart", (ev)=>{
    // v91: 行特定を安定化 + 操作ログを追加
    let tr = null;
    let chosenBy = "";
    const now = Date.now();
    const lp = window.__empDnDLastLeftPointer;

    const eventTr = (ev.target && ev.target.closest) ? ev.target.closest("tr[data-id]") : null;

    // 1) 直前mousedownで記録した行（最優先）
    try{
      if(lp && lp.id && (now - (lp.ts||0) < 1200)){
        const cand = employeeTableBody.querySelector(`tr[data-id="${lp.id}"]`);
        if(cand && cand.style && cand.style.display!=="none"){
          tr = cand;
          chosenBy = "lastPointer";
        }
      }
    }catch(_e){}

    // v92: if the stored pointer row looks wrong, re-check by current pointer position
    if(tr && chosenBy==="lastPointer"){
      try{
        const x = (lp && lp.x) ? lp.x : (ev.clientX||0);
        const y = (lp && lp.y) ? lp.y : (ev.clientY||0);
        const elAtPoint = document.elementFromPoint(x, y);
        const trP = elAtPoint && elAtPoint.closest ? elAtPoint.closest("tr[data-id]") : null;
        const idP = trP ? String(trP.getAttribute("data-id")||"") : null;
        const idT = tr ? String(tr.getAttribute("data-id")||"") : null;
        if(trP && idP && idT && idP !== idT){
          tr = trP;
          chosenBy = "lastPointer->pointFix";
        }
      }catch(_e){}
    }

    // 2) event.target から辿れる行
    if(!tr && eventTr){
      tr = eventTr;
      chosenBy = "eventTarget";
    }

    // 3) ポイント位置からも特定（環境によってtargetがズレる対策）
    if(!tr){
      try{
        const x = (lp && lp.x) ? lp.x : (ev.clientX||0);
        const y = (lp && lp.y) ? lp.y : (ev.clientY||0);
        const elAtPoint = document.elementFromPoint(x, y);
        const tr2 = elAtPoint && elAtPoint.closest ? elAtPoint.closest("tr[data-id]") : null;
        if(tr2 && tr2.style && tr2.style.display!=="none"){
          tr = tr2;
          chosenBy = "elementFromPoint";
        }
      }catch(_e){}
    }

    // ログ：直前に何を掴んだか / dragstartでどの行が選ばれたか
    try{
      console.log("[settings-dnd] left dragstart pick", {
        chosenBy,
        chosenRowId: tr ? String(tr.getAttribute("data-id")||"") : null,
        eventTargetRowId: eventTr ? String(eventTr.getAttribute("data-id")||"") : null,
        lastPointer: lp ? {id: lp.id, ts: lp.ts, ageMs: (lp.ts? (now-lp.ts): null), x: lp.x, y: lp.y, tag: lp.tag, cls: lp.cls} : null,
        eventTargetTag: String(ev.target && ev.target.tagName || ""),
        eventTargetCls: String(ev.target && ev.target.className || "")
      });
    }catch(_e){}

    // 次のdragstartに古い情報を使わない（ワンショット）
    try{ if(lp) lp.ts = 0; }catch(_e){}

// v201: 行のどこからでもドラッグ開始OK（ただし入力・ボタン等のコントロール上は除外）
    const interactive = ev.target && ev.target.closest ? ev.target.closest("input,select,textarea,button,a,label") : null;
    if(interactive){
      // コントロールの操作（選択/クリック）を邪魔しない
      ev.preventDefault();
      return;
    }

    
    // v9: つかんだ感を「行全体」で統一（1列目から掴んでも同じ）
    try{
      __empDnD.sourceEl = tr;
      __empDnD.sourceStyle = {
        opacity: tr.style.opacity,
        transform: tr.style.transform,
        boxShadow: tr.style.boxShadow,
        filter: tr.style.filter
      };
      tr.classList.add("dragging-row");
      // 左テーブルは「薄くしない」：見た目は変えず、drag image（影付き）だけを強化する

      // drag image を行全体にする（ブラウザがセル/テキストだけ掴む挙動を防ぐ）
      const rect = tr.getBoundingClientRect();
      const ox = Math.max(0, Math.min(rect.width, (ev.clientX || rect.left) - rect.left));
      const oy = Math.max(0, Math.min(rect.height, (ev.clientY || rect.top) - rect.top));

      // ghostを作ってsetDragImage（影を強くして視認性UP）
      const pad = 8; // 影の見切れ防止
      const ghostWrap = document.createElement("div");
      ghostWrap.style.position = "fixed";
      ghostWrap.style.left = "-9999px";
      ghostWrap.style.top = "-9999px";
      ghostWrap.style.padding = pad + "px";
      ghostWrap.style.background = "transparent";
      ghostWrap.style.pointerEvents = "none";
      ghostWrap.style.filter = "drop-shadow(0 14px 26px rgba(0,0,0,.40)) drop-shadow(0 4px 10px rgba(0,0,0,.25))";

      const ghostTable = document.createElement("table");
      ghostTable.className = "employee-table-fixed";
      ghostTable.style.borderCollapse = "collapse";
      ghostTable.style.width = rect.width + "px";
      ghostTable.style.pointerEvents = "none";
      ghostTable.style.background = "var(--panel-bg, #fff)";
      ghostTable.style.borderRadius = "10px";
      ghostTable.style.overflow = "hidden";

      const ghostBody = document.createElement("tbody");
      const ghostRow = tr.cloneNode(true);
      // clone内のフォーム類は見た目だけにする
      ghostRow.querySelectorAll("input,select,textarea,button,a").forEach(el=>{ el.setAttribute("disabled",""); });
      ghostBody.appendChild(ghostRow);
      ghostTable.appendChild(ghostBody);

      ghostWrap.appendChild(ghostTable);
      document.body.appendChild(ghostWrap);
      __empDnD.ghostEl = ghostWrap;

      try{ ev.dataTransfer.setDragImage(ghostWrap, ox + pad, oy + pad); }catch(_e){}
      // ghostはcleanupで撤去
    }catch(_e){}
let empId = (tr && tr.dataset) ? (tr.dataset.id || "") : "";
    if(!empId){
      try{
        const handleEl = tr ? (tr.querySelector(".emp-drag-cell") || tr.querySelector("[data-emp-id]")) : null;
        const h = handleEl && handleEl.dataset ? (handleEl.dataset.empId || handleEl.dataset.empId0 || "") : "";
        if(h) empId = h;
      }catch(_e){}
    }


    if(!empId){
      // idが取れない場合はDnDを開始しない（誤った行が動く事故防止）
      try{ ev.preventDefault(); }catch(_e){}
      cleanupEmpDnDVisual();
      return;
    }

    __empDnD.active = true;
    __empDnD.from = "left";
    __empDnD.empId = String(empId);
    __empDnD.fromStoreId = (tr.querySelector("select.emp-store")?.value || "") || "";
    if(__empDnD.fromStoreId==="__none__") __empDnD.fromStoreId="";
    __empDnD.placeholderIndex = null;

    setDraggingUi(true);
    window.setDropOver(false);

    try{
      ev.dataTransfer.effectAllowed = "move";
      ev.dataTransfer.setData("text/plain", JSON.stringify({type:"emp", id:String(empId), from:"left"}));
      try{ ev.dataTransfer.setData("text/uri-list",""); }catch(_e){}
    }catch(_e){}
    logSettingsDnD(`ドラッグ開始（左）: ${String(empId)}`);
  }, true);

  dom.employeeTableBody.addEventListener("dragend", ()=>{
    cleanupEmpDnDVisual();
  }, true);

  // 右側：captureで必ずdragover/dropを受ける
  
  // --- (v125) DOM再描画で employeeOrderList が差し替わってもDnDが動くように、document委譲を入れる ---
  if(!window.__empDnDDelegationInstalled){
    window.__empDnDDelegationInstalled = true;

    document.addEventListener("dragover", (ev)=>{
      const list = ev.target && ev.target.closest ? ev.target.closest("#employeeOrderList") : null;
      if(!list) return;
      // 二重実行防止（下位のリスナーを止める）
      ev.stopPropagation();
      if(ev.stopImmediatePropagation) ev.stopImmediatePropagation();
          if(!__empDnD.active) return;
          ev.preventDefault();
          ev.stopPropagation();
          if(ev.stopImmediatePropagation) ev.stopImmediatePropagation();
          updateEmpDnDPlaceholder(ev.clientY);
          // v117: placeholderが出ている時だけ move を表示
          if(ev.dataTransfer){
            const hasPlaceholder = !!(__empDnD.placeholder && __empDnD.placeholder.isConnected);
            ev.dataTransfer.dropEffect = "move";
          }
          window.setDropOver(true);
    }, true);

    document.addEventListener("drop", (ev)=>{
      const list = ev.target && ev.target.closest ? ev.target.closest("#employeeOrderList") : null;
      if(!list) return;
      // 二重実行防止（下位のリスナーを止める）
      ev.stopPropagation();
      if(ev.stopImmediatePropagation) ev.stopImmediatePropagation();
          // 右側ドロップ（左→右、右→右、店舗間移動）
          try{
            console.debug("[settings-dnd] drop(candidate)", {
              active: __empDnD.active,
              targetTag: ev.target?.tagName,
              targetCls: ev.target?.className,
              targetId: ev.target?.id,
              empId: __empDnD.empId,
              from: __empDnD.from,
              fromStoreId: __empDnD.fromStoreId,
              storeSelect: dom.employeeOrderStoreSelect?.value || "",
              placeholderIndex: __empDnD.placeholderIndex
            });
          }catch(_e){}
      
          if(!__empDnD.active) return;
      
          // drop を受け付ける
          ev.preventDefault();
      
          const payload = readEmpDnDPayload(ev);
          if(!payload || payload.type!=="emp"){
            try{ console.debug("[settings-dnd] drop: no payload"); }catch(_e){}
            cleanupEmpDnDVisual();
            return;
          }
      
          // ドロップ先店舗：基本はセレクト。dropzoneにdatasetがあればそれを優先（将来の拡張・保険）
          let storeId = dom.employeeOrderStoreSelect?.value || "";
          const dz = ev.target && ev.target.closest ? ev.target.closest(".employee-order-dropzone[data-store-id]") : null;
          if(dz && dz.dataset && dz.dataset.storeId) storeId = String(dz.dataset.storeId);
      
          if(!storeId){
            try{ console.debug("[settings-dnd] drop: storeId empty"); }catch(_e){}
            cleanupEmpDnDVisual();
            return;
          }
      
          const targetIndex = computeEmpDnDPlaceholderIndex();
          try{
            console.debug("[settings-dnd] drop: apply", {empId: payload.id, storeId, targetIndex});
          }catch(_e){}
      
          applyEmpDnDDrop(payload, storeId, targetIndex);
      
          cleanupEmpDnDVisual();
          try{ if(typeof refreshEmployeeStoreOptionsFromSettings==='function') refreshEmployeeStoreOptionsFromSettings(); }catch(_e){}
          try{ if(typeof refreshEmployeeOrderPanel==='function') setTimeout(()=>refreshEmployeeOrderPanel(),0); }catch(_e){}
    }, true);
  }
dom.employeeOrderList.addEventListener("dragover", (ev)=>{
    if(!__empDnD.active) return;

    // ドラッグが実際に動いたことを記録（即キャンセル時に外ドロップ扱いしないため）
    __empDnD.moved = true;
    ev.preventDefault();
    ev.stopPropagation();
    if(ev.stopImmediatePropagation) ev.stopImmediatePropagation();
    updateEmpDnDPlaceholder(ev.clientY);
    // v117: placeholderが出ている時だけ move を表示
    if(ev.dataTransfer){
      const hasPlaceholder = !!(__empDnD.placeholder && __empDnD.placeholder.isConnected);
      ev.dataTransfer.dropEffect = "move";
    }
    window.setDropOver(true);
  }, true);

  dom.employeeOrderList.addEventListener("drop", (ev)=>{
    // 右側ドロップ（左→右、右→右、店舗間移動）
    try{
      console.debug("[settings-dnd] drop(candidate)", {
        active: __empDnD.active,
        targetTag: ev.target?.tagName,
        targetCls: ev.target?.className,
        targetId: ev.target?.id,
        empId: __empDnD.empId,
        from: __empDnD.from,
        fromStoreId: __empDnD.fromStoreId,
        storeSelect: dom.employeeOrderStoreSelect?.value || "",
        placeholderIndex: __empDnD.placeholderIndex
      });
    }catch(_e){}

    if(!__empDnD.active) return;

    // drop を受け付ける
    ev.preventDefault();

    const payload = readEmpDnDPayload(ev);
    if(!payload || payload.type!=="emp"){
      try{ console.debug("[settings-dnd] drop: no payload"); }catch(_e){}
      cleanupEmpDnDVisual();
      return;
    }

    // ドロップ先店舗：基本はセレクト。dropzoneにdatasetがあればそれを優先（将来の拡張・保険）
    let storeId = dom.employeeOrderStoreSelect?.value || "";
    const dz = ev.target && ev.target.closest ? ev.target.closest(".employee-order-dropzone[data-store-id]") : null;
    if(dz && dz.dataset && dz.dataset.storeId) storeId = String(dz.dataset.storeId);

    if(!storeId){
      try{ console.debug("[settings-dnd] drop: storeId empty"); }catch(_e){}
      cleanupEmpDnDVisual();
      return;
    }

    const targetIndex = computeEmpDnDPlaceholderIndex();
    try{
      console.debug("[settings-dnd] drop: apply", {empId: payload.id, storeId, targetIndex});
    }catch(_e){}

    applyEmpDnDDrop(payload, storeId, targetIndex);

    cleanupEmpDnDVisual();
    try{ if(typeof refreshEmployeeStoreOptionsFromSettings==='function') refreshEmployeeStoreOptionsFromSettings(); }catch(_e){}
    try{ if(typeof refreshEmployeeOrderPanel==='function') refreshEmployeeOrderPanel(); }catch(_e){}
  }, true);

  // 右側：並び替えのdragstart/dragend（委譲）
  dom.employeeOrderList.addEventListener("dragstart", (ev)=>{
    const item = ev.target && ev.target.closest ? ev.target.closest(".employee-order-item[data-emp-id]") : null;
    if(!item) return;
    const empId = item.dataset.empId;
    const storeId = item.dataset.storeId || (dom.employeeOrderStoreSelect?.value||"");
    __empDnD.active = true;
    __empDnD.from = "right";
    __empDnD.empId = String(empId);
    __empDnD.fromStoreId = String(storeId);
    __empDnD.placeholderIndex = null;
      __empDnD.dropCommitted = false;
      __empDnD.moved = false;
      __empDnD.startTs = Date.now();
    setDraggingUi(true);
    window.setDropOver(false);

    // --- v8: 「つかんだ感」(見た目のみ) ＋ドラッグ画像(ghost) ---
    try{
      // 二重dragstart防止（ブラウザがdragをキャンセル→別要素でdragstart連発する対策）
      if(__empDnD && __empDnD.active && __empDnD.empId && String(__empDnD.empId)!==String(empId)){
        cleanupEmpDnDVisual();
      }
    }catch(_e){}

    try{
      __empDnD.sourceEl = item;
      __empDnD.sourceStyle = {
        opacity: item.style.opacity,
        transform: item.style.transform,
        boxShadow: item.style.boxShadow,
        filter: item.style.filter
      };
      item.classList.add("dragging-row");
      item.style.opacity = "0.35";
      item.style.transform = "scale(0.985)";
      item.style.boxShadow = "0 8px 20px rgba(0,0,0,0.14)";
      item.style.filter = "saturate(0.9)";

      // 既定のドラッグ画像が取りづらい環境向けに ghost を明示（DOM/レイアウトは壊さない）
      try{
        const g = item.cloneNode(true);
        g.classList.add("employee-order-ghost");
        g.style.position = "fixed";
        g.style.left = "-2000px";
        g.style.top = "-2000px";
        g.style.pointerEvents = "none";
        g.style.margin = "0";
        g.style.opacity = "0.85";
        document.body.appendChild(g);
        __empDnD.ghostEl = g;
        if(ev && ev.dataTransfer && ev.dataTransfer.setDragImage){
          ev.dataTransfer.setDragImage(g, 16, 16);
        }
      }catch(_e){}
    }catch(_e){}

    ensureEmpDnDPlaceholder();
    logSettingsDnD(`ドラッグ開始（右）: ${String(empId)} / store=${String(storeId)}`);
  }, true);

  dom.employeeOrderList.addEventListener("dragend", (ev)=>{
    // 右リストから「ドロップ可能な領域の外」へ落とした場合：
    // HTML5 DnD では drop イベントは“どこにも”発火しないため、dragend で検知して未所属へ戻す。
    try{
      if(__empDnD && __empDnD.active && __empDnD.from==="right" && !__empDnD.dropCommitted && __empDnD.moved && (Date.now()-(__empDnD.startTs||0) > 60)){
        const empId = String(__empDnD.empId||"");
        if(empId){
          console.debug("[settings-dnd] 右→外ドロップ（drop未発火）: 未所属へ戻す", {empId, fromStoreId: __empDnD.fromStoreId});
          try{ setEmployeeStore(empId, "", true); }catch(_e){}
          // 未所属へ移動する際は店舗内順=0（仕様）
          try{ setEmployeeOrder(empId, 0, true); }catch(_e){}
          try{ if(typeof refreshEmployeeStoreOptionsFromSettings==='function') refreshEmployeeStoreOptionsFromSettings(); }catch(_e){}
          try{ if(typeof renderEmployeeSettingsTables==='function') renderEmployeeSettingsTables(); }catch(_e){}
          try{ if(typeof refreshEmployeeOrderPanel==='function') refreshEmployeeOrderPanel(); }catch(_e){}
        }
      }
    }catch(_e){}
    cleanupEmpDnDVisual();
  }, true);

  // v8: 一部環境で dragend が source/list に届かない場合があるため、document でも後始末
  document.addEventListener("dragend", (ev)=>{
    if(!__empDnD || !__empDnD.active) return;
    cleanupEmpDnDVisual();
    try{ __setUnassignHint(false); }catch(_e){}
  }, true);

  // ページ遷移/ファイルオープン防止
  // 重要: document 全体で dragover を preventDefault すると「どこでもドロップ可能」と見なされ、
  // 禁止カーソル（not-allowed）が出なくなる。
  // → ドロップ先（右パネル）のみ preventDefault。
  document.addEventListener("dragover", (ev)=>{
    if(!__empDnD.active) return;
    // ドロップ不可の場所では何もしない（ブラウザの not-allowed 表示に任せる）
    const ok = isInEmployeeOrderArea(ev.target);
    if(!ok){
      try{ if(ev.dataTransfer) ev.dataTransfer.dropEffect = "none"; }catch(_e){}
      window.setDropOver(false);
      // v12: リスト外に出たら「ここに挿入」placeholderを消す（残像防止）
      try{
        if(__empDnD.placeholder && __empDnD.placeholder.parentNode){
          __empDnD.placeholder.parentNode.removeChild(__empDnD.placeholder);
        }
        __empDnD.placeholderIndex = null;
      }catch(_e){}
      return;
    }
    // 右パネル内でも、placeholderが出ない状況（=ドロップ不可能）は none
    try{
      const hasPlaceholder = !!(__empDnD.placeholder && __empDnD.placeholder.isConnected);
      if(ev.dataTransfer) ev.dataTransfer.dropEffect = "move";
    }catch(_e){}
  }, true);

  document.addEventListener("drop", (ev)=>{
    if(!__empDnD.active) return;
    // どこで落としてもブラウザ既定の挙動（ファイルを開く等）を抑止
    ev.preventDefault();
    ev.stopPropagation();
    window.setDropOver(false);

    // vXXX: 右（店舗内一覧）から外へドロップしたら「未所属」に戻す
    try{
      const payload = readEmpDnDPayload(ev);
      const inOrderArea = ev.target && ev.target.closest
        ? ev.target.closest("#employeeOrderList, .employee-order-head, .employee-order-dropzone, .employee-order-item, #employeeOrderStoreSelect")
        : null;

      if(payload && payload.type==="emp" && payload.from==="right" && !inOrderArea){
        // 外側に落とした＝未所属へ戻す
        try{ console.debug("[settings-dnd] drop outside -> unassign", {empId: payload.id, fromStore: payload.storeId}); }catch(_e){}
        applyEmpDnDDrop(payload, "", null);
        try{ if(typeof refreshEmployeeStoreOptionsFromSettings==='function') refreshEmployeeStoreOptionsFromSettings(); }catch(_e){}
        try{ if(typeof refreshEmployeeOrderPanel==='function') refreshEmployeeOrderPanel(); }catch(_e){}
      }
    }catch(_e){}

    cleanupEmpDnDVisual();
  }, true);

  function logSettingsDnD(msg){
    try{
      if(typeof log==='function') { log(msg); return; }
    }catch(_e){}
    try{ console.debug(`[settings-dnd] ${msg}`); }catch(_e){}
  }

  function readEmpDnDPayload(ev){
    // v103以前の互換：内部状態優先
    if(__empDnD.empId){
      return {type:"emp", id:String(__empDnD.empId), from: __empDnD.from, storeId: __empDnD.fromStoreId};
    }
    try{
      const t = ev.dataTransfer?.getData("text/plain") || "";
      if(!t) return null;
      return JSON.parse(t);
    }catch(_e){ return null; }
  }

  function ensureEmpDnDPlaceholder(){
    if(__empDnD.placeholder && __empDnD.placeholder.isConnected) return;
    const ph = document.createElement("div");
    ph.className = "employee-order-placeholder";
    ph.textContent = "ここに挿入";
    __empDnD.placeholder = ph;
  }

  function getCurrentOrderIds(storeId){
    const items = Array.from(dom.employeeOrderList.querySelectorAll('.employee-order-item[data-store-id="'+String(storeId)+'"][data-emp-id]'));
    return items.map(el=>String(el.dataset.empId));
  }

  function getEmpIdsInStoreFromTable(storeId){
    const sid = String(storeId||"__none__").trim() || "__none__";
    const rows = Array.from(dom.employeeTableBody.querySelectorAll('tr[data-id]')).filter(tr=>{
      const sel = tr.querySelector('select.emp-store');
      // option未生成/再構築の瞬間は value が空になることがあるため、dataset.desired を優先して読む
      const raw = (sel && (sel.dataset?.desired || sel.value)) ? (sel.dataset.desired || sel.value) : "__none__";
      const v = String(raw || "__none__").trim() || "__none__";
      return v === sid;
    });
    const getNameKey = (tr)=>{
      const last = (tr.querySelector(".emp-last")?.value || "").trim();
      const first = (tr.querySelector(".emp-first")?.value || "").trim();
      return (last + " " + first).trim();
    };
    rows.sort((a,b)=>{
      const oa = Number(a.querySelector(".emp-order")?.value || "0") || 0;
      const ob = Number(b.querySelector(".emp-order")?.value || "0") || 0;
      if(oa!==ob) return oa-ob;
      return getNameKey(a).localeCompare(getNameKey(b), "ja");
    });
    return rows.map(tr=>String(tr.dataset.id||"")).filter(Boolean);
  }

// v79: 店舗内順の正規化/末尾追加を DOM順や既存値の揺れに依存せず「確定的」に行う
function normalizeEmployeeOrdersForStoreSilent(storeId){
  const sid = String(storeId||"__none__");
  const ids = getEmpIdsInStoreFromTable(sid).map(String);
  ids.forEach((id,i)=>{ try{ setEmployeeOrder(String(id), (i+1), true); }catch(_e){} });
}
function appendEmployeeToStoreEnd(empId, storeId){
  const sid = String(storeId||"__none__");
  let ids = getEmpIdsInStoreFromTable(sid).map(String).filter(Boolean);
  const eid = String(empId||"");
  ids = ids.filter(id=>id!==eid);
  ids.push(eid);
  ids.forEach((id,i)=>{ try{ setEmployeeOrder(String(id), (i+1), true); }catch(_e){} });
}


  function updateEmpDnDPlaceholder(clientY){
    ensureEmpDnDPlaceholder();
    const storeId = dom.employeeOrderStoreSelect?.value || "";
    if(!storeId) return;

    const dropzone = dom.employeeOrderList.querySelector(".employee-order-dropzone");
    if(!dropzone) return;

    const rows = Array.from(dropzone.querySelectorAll(".employee-order-item[data-emp-id]"))
      .filter(el=>!el.classList.contains("dragging-row"));

    // 空リストは末尾
    if(rows.length===0){
      if(__empDnD.placeholder.parentNode!==dropzone) dropzone.appendChild(__empDnD.placeholder);
      __empDnD.placeholderIndex = 0;
      return;
    }

    let insertIndex = rows.length;
    for(let i=0;i<rows.length;i++){
      const r = rows[i];
      const rect = r.getBoundingClientRect();
      const mid = rect.top + rect.height/2;
      if(clientY < mid){
        insertIndex = i;
        break;
      }
    }
    __empDnD.placeholderIndex = insertIndex;

    // DOMへ反映
    if(insertIndex>=rows.length){
      dropzone.appendChild(__empDnD.placeholder);
    }else{
      dropzone.insertBefore(__empDnD.placeholder, rows[insertIndex]);
    }
  }

  function computeEmpDnDPlaceholderIndex(){
    const idx = __empDnD.placeholderIndex;
    if(idx==null) return null;
    return idx;
  }

  function applyEmpDnDDrop(payload, storeId, targetIndex){
    try{ __empDnD.dropCommitted = true; }catch(_e){}
    try{ __empDnD.isReordering = true; }catch(_e){}

    const empId = String(payload?.id ?? "");
    if(!empId) return;

	    const toStoreId = String(storeId ?? "");
    const from = String(__empDnD?.from ?? payload?.from ?? "");
    const fromStoreId = String(__empDnD?.fromStoreId ?? payload?.storeId ?? "");

	    // === 未所属へ移動（仕様：店舗内順=0） ===
	    // storeId が空（""）のときは未所属扱い。
	    if(!toStoreId){
	      try{ setEmployeeStore(empId, "", true); }catch(_e){}
	      try{ setEmployeeOrder(empId, 0, true); }catch(_e){}
	      // 元店舗の並び順は欠番が出るので詰め直す（未所属側は採番しない）
	      try{ if(fromStoreId) normalizeEmployeeOrdersForStoreSilent(fromStoreId); }catch(_e){}
	      try{ __empDnD.isReordering = false; }catch(_e){}
	      try{ refreshEmployeeOrderPanel(); }catch(_e){}
	      return;
	    }

    // 同一店舗内の並び替え（右→右 & 店舗一致）なら「所属変更」は不要
    const fromSameStore = (from === "right" && fromStoreId && fromStoreId === toStoreId);

    // 他店舗からの移動 / 左→右 は、まず所属店舗を更新
    if(!fromSameStore){
      try{ setEmployeeStore(empId, toStoreId, true); }catch(_e){}
      // 右→右で別店舗から移動した場合は、元店舗の並び順を詰め直す
      try{
        if(from === "right" && fromStoreId){
          const fromIds = getCurrentOrderIds(fromStoreId);
          fromIds.forEach((id,i)=>{ try{ setEmployeeOrder(String(id), (i+1)); }catch(_e){} });
        }
      }catch(_e){}
    }

    // 目的店舗の現在順を取得（ドラッグ対象は除外）
    // 同一店舗内の並び替えは右パネルの表示順を使う。
    // 新規割当/別店舗移動は左テーブル（実データ）側から作る（右パネルDOMが古い/空のケース対策）。
    let list = (fromSameStore ? getCurrentOrderIds(toStoreId) : getEmpIdsInStoreFromTable(toStoreId)).map(String).filter(id => id !== empId);

    // 挿入位置
    // - 右リスト内の並び替えはもちろん、
    // - 左→右の新規割当や店舗間移動でも、ドロップ位置（placeholder）を尊重する。
    //   targetIndex が取れない場合のみ末尾に入れる。
    let idx = Number.isFinite(targetIndex) ? Number(targetIndex) : list.length;
    if(idx < 0) idx = 0;
    if(idx > list.length) idx = list.length;

    list.splice(idx, 0, empId);

    // 1..n を採番（重複を作らない）
    // 途中でinput/changeが走ると並びが揺れるので、まずはsilentで一括更新し、最後に一度だけ整形＆再描画する
    list.forEach((id,i)=>{ try{ setEmployeeOrder(String(id), (i+1), true); }catch(_e){} });

    // 最終的に対象店舗の店舗内順を 1..n に正規化（保険：重複や欠番を潰す）
    try{ __empDnD.isReordering = false; }catch(_e){}
    try{ recalcEmployeeOrdersForStore(toStoreId); }catch(_e){}

    // 右パネル表示を更新（ここで一度だけ）
    try{ refreshEmployeeOrderPanel(); }catch(_e){}
    // 左テーブルのフィルタ（未所属のみ表示）を即時反映させる
    try{ if(typeof renderEmployeeSettingsTables==='function') renderEmployeeSettingsTables(); }catch(_e){}
}

  function setEmployeeStore(empId, storeId, silent){
    const tr = dom.employeeTableBody.querySelector('tr[data-id="'+String(empId)+'"]');
    if(!tr) return;
    const sel = tr.querySelector("select.emp-store");
    if(!sel) return;

    // 重要: option再構築の瞬間など、valueにセットしても空になる場合がある。
    // その場合でも所属が失われないよう、意図した値を dataset.desired に必ず保持する。
    const intended = (storeId!==undefined && storeId!==null && String(storeId).trim()!=="") ? String(storeId) : "__none__";
    sel.dataset.desired = intended;

    // 可能なら表示上のvalueも合わせる（optionが無い場合は空のままでもOK。desiredが真実）
    try{ sel.value = intended; }catch(_e){}

    if(!silent){
      // changeイベントを起こして未保存判定を正しくする（既存実装に合わせる）
      sel.dispatchEvent(new Event("change", {bubbles:true}));
    }
  }
  }

  function setEmployeeOrder(empId, order, silent){
    const tr = dom.employeeTableBody.querySelector('tr[data-id="'+String(empId)+'"]');
    if(!tr) return;
    const input = tr.querySelector(".emp-order");
    if(!input) return;
    // order は表示値（1..n）。readonlyだがプログラムからは更新できる。
    input.value = String(Number(order)||0);
    if(!silent){
      input.dispatchEvent(new Event("input", {bubbles:true}));
      input.dispatchEvent(new Event("change", {bubbles:true}));
    }
  }

  // ===== GLOBAL DnD UI helpers (v84) =====
  // 一部の関数がスコープ外から参照されるケースがあり、ReferenceError になることがあったため
  // グローバルに同名のフォールバック実装を用意する。
  function setDraggingUi(isOn){
    try{ document.documentElement.classList.toggle("is-dragging", !!isOn); }catch(_e){}
    try{
      const list = document.getElementById("employeeOrderList");
      if(list){
        list.classList.toggle("drop-allowed", !!isOn);
        const dz = list.querySelector(".employee-order-dropzone");
        if(dz) dz.classList.toggle("drop-allowed", !!isOn);
      }
    }catch(_e){}
  }

  function cleanupEmpDnDVisual(){
    try{ __empDnD.isReordering = false; }catch(_e){}
    __empDnD.active = false;
    __empDnD.from = null;
    __empDnD.empId = null;
    __empDnD.fromStoreId = null;
    __empDnD.placeholderIndex = null;

        // v6: 元要素の復帰とスペーサの撤去
    try{
      const el = __empDnD.sourceEl || __empDnD.dragItem;
      const st = __empDnD.sourceStyle || null;
      if(el){
        el.classList.remove("dragging-row");
        // つかんだ感のために一時的に入れたインラインスタイルを復元
        if(st){
          el.style.opacity = st.opacity || "";
          el.style.transform = st.transform || "";
          el.style.boxShadow = st.boxShadow || "";
          el.style.filter = st.filter || "";
        }else{
          // 保険
          el.style.opacity = "";
          el.style.transform = "";
          el.style.boxShadow = "";
          el.style.filter = "";
        }
      }
      if(__empDnD.dragSpacer && __empDnD.dragSpacer.parentNode){
        __empDnD.dragSpacer.parentNode.removeChild(__empDnD.dragSpacer);
      }
      // ghost drag image の撤去
      if(__empDnD.ghostEl && __empDnD.ghostEl.parentNode){ __empDnD.ghostEl.parentNode.removeChild(__empDnD.ghostEl); }
    }catch(_e){}
    __empDnD.sourceEl = null;
    __empDnD.sourceStyle = null;
    __empDnD.dragItem = null;
    __empDnD.dragSpacer = null;
    __empDnD.ghostEl = null;

document.querySelectorAll(".dragging-row").forEach(el=>el.classList.remove("dragging-row"));
    if(__empDnD.placeholder && __empDnD.placeholder.parentNode){
      __empDnD.placeholder.parentNode.removeChild(__empDnD.placeholder);
    }
    window.setDropOver(false);
    setDraggingUi(false);
  }

  // === v4: 右リストから外へドラッグしたときも「ドロップ可能」にして禁止カーソルを消し、ヒントを出す ===
  const __unassignHintEl = document.getElementById("unassignHintOverlay");
  function __setUnassignHint(on, x, y){
    if(!__unassignHintEl) return;
    if(on){
      __unassignHintEl.classList.add("show");
      if(typeof x==="number" && typeof y==="number"){
        __unassignHintEl.style.setProperty("--x", x + "px");
        __unassignHintEl.style.setProperty("--y", y + "px");
      }
    }else{
      __unassignHintEl.classList.remove("show");
    }
  }

  // 右リスト外は「未所属へ戻す」扱い：document をドロップ可能にする（cursor: not-allowed を回避）
  document.addEventListener("dragover", (ev)=>{
    if(!__empDnD || !__empDnD.active) return;
    if(__empDnD.from !== "right") return;

    __empDnD.moved = true;

    const overOrderList = ev.target && ev.target.closest ? ev.target.closest("#employeeOrderList") : null;
    if(overOrderList){
      __setUnassignHint(false);
      return; // 通常の並び替え/挿入
    }

    // ここをドロップ可能にして、禁止カーソルを「移動」に変える
    ev.preventDefault();
    try{ ev.dataTransfer.dropEffect = "move"; }catch(_e){}
    __setUnassignHint(true, ev.clientX, ev.clientY);
  }, true);

  document.addEventListener("drop", (ev)=>{
    if(!__empDnD || !__empDnD.active) return;
    if(__empDnD.from !== "right") return;

    const overOrderList = ev.target && ev.target.closest ? ev.target.closest("#employeeOrderList") : null;
    if(overOrderList){
      __setUnassignHint(false);
      return; // 通常の drop ハンドラに任せる
    }

    // 外に落とした＝未所属へ戻す
    ev.preventDefault();
    ev.stopPropagation();
    if(ev.stopImmediatePropagation) ev.stopImmediatePropagation();

    __empDnD.dropCommitted = true;
    __setUnassignHint(false);
    const payload = { type:"emp", from:"right", id: __empDnD.empId, storeId: __empDnD.fromStoreId };
    try{ console.debug("[settings-dnd] drop(outside) -> unassign", payload); }catch(_e){}
    try{
      applyEmpDnDDrop(payload, "", null);
      if(typeof refreshEmployeeStoreOptionsFromSettings === "function") refreshEmployeeStoreOptionsFromSettings();
      if(typeof refreshEmployeeOrderPanel === "function") refreshEmployeeOrderPanel();
      if(typeof renderEmployeeTablesInSettings === "function") renderEmployeeTablesInSettings();
    }catch(_e){
      console.error("[settings-dnd] unassign failed:", _e);
    }
  }, true);

  document.addEventListener("dragend", ()=>{
    __setUnassignHint(false);
  }, true);









// ===== 従業員タブ：店舗内順の採番/割当（グローバル補助）=====
// 右パネルDnD実装が内部スコープ化されている場合でも、左の所属店舗変更や「外ドロップで未所属」など
// 仕様上必須の挙動が常に動くように、最低限の処理をグローバル関数として提供する。
function getEmpIdsInStoreFromTable(storeId){
  if(!dom || !dom.employeeTableBody) return [];
  const sid = String(storeId||"__none__").trim() || "__none__";
  const rows = Array.from(dom.employeeTableBody.querySelectorAll('tr[data-id]')).filter(tr=>{
    const sel = tr.querySelector('select.emp-store');
    const raw = (sel && (sel.dataset?.desired || sel.value)) ? (sel.dataset.desired || sel.value) : "__none__";
    const v = String(raw || "__none__").trim() || "__none__";
    return v === sid;
  });
  const nameKey = (tr)=>{
    const last = (tr.querySelector(".emp-last")?.value || "").trim();
    const first = (tr.querySelector(".emp-first")?.value || "").trim();
    return (last + " " + first).trim();
  };
  rows.sort((a,b)=>{
    const oa = Number(a.querySelector(".emp-order")?.value || "0") || 0;
    const ob = Number(b.querySelector(".emp-order")?.value || "0") || 0;
    if(oa!==ob) return oa-ob;
    return nameKey(a).localeCompare(nameKey(b), "ja");
  });
  return rows.map(tr=>String(tr.dataset.id||"")).filter(Boolean);
}

function setEmployeeOrder(empId, order, silent){
  if(!dom || !dom.employeeTableBody) return;
  const tr = dom.employeeTableBody.querySelector('tr[data-id="'+String(empId)+'"]');
  if(!tr) return;
  const input = tr.querySelector(".emp-order");
  if(!input) return;
  input.value = String(Number(order)||0);
  if(!silent){
    input.dispatchEvent(new Event("input", {bubbles:true}));
    input.dispatchEvent(new Event("change", {bubbles:true}));
  }
}
function setEmployeeStore(empId, storeId, silent){
  if(!dom || !dom.employeeTableBody) return;
  const tr = dom.employeeTableBody.querySelector('tr[data-id="'+String(empId)+'"]');
  if(!tr) return;
  const sel = tr.querySelector("select.emp-store");
  if(!sel) return;
  const intended = (storeId!==undefined && storeId!==null && String(storeId).trim()!=="") ? String(storeId) : "__none__";
  sel.dataset.desired = intended;
  try{ sel.value = intended; }catch(_e){}
  if(!silent){
    sel.dispatchEvent(new Event("change", {bubbles:true}));
  }
}

// 欠番/重複を潰して 1..n を振り直す（静かに）
function normalizeEmployeeOrdersForStoreSilent(storeId){
  const sid = String(storeId||"__none__");
  const ids = getEmpIdsInStoreFromTable(sid);
  ids.forEach((id,i)=>{ try{ setEmployeeOrder(id, i+1, true); }catch(_e){} });
}
// 指定従業員を末席に移動してから 1..n 採番（静かに）
function appendEmployeeToStoreEnd(empId, storeId){
  const sid = String(storeId||"__none__");
  const eid = String(empId||"");
  let ids = getEmpIdsInStoreFromTable(sid).filter(id=>id!==eid);
  ids.push(eid);
  ids.forEach((id,i)=>{ try{ setEmployeeOrder(id, i+1, true); }catch(_e){} });
}

// drop適用（左→右の割当 / 右→右の並び替え / 右→外の未所属化）
// toStoreId: "" または "__none__" を未所属扱い
function applyEmpDnDDrop(payload, toStoreId, targetIndex){
  if(!payload || payload.type!=="emp") return;
  const empId = String(payload.id||"");
  if(!empId) return;

  // mark committed (for right-dragend fallback)
  try{ if(window.__empDnD){ __empDnD.dropCommitted = true; } }catch(_e){}

  const toSid = (toStoreId && String(toStoreId).trim() && String(toStoreId)!=="__none__") ? String(toStoreId) : "";
  const from = String(payload.from||"");
  const fromSid = (payload.storeId && String(payload.storeId)!=="__none__") ? String(payload.storeId) : "";

  // === 未所属へ（仕様：店舗内順=0） ===
  if(!toSid){
    try{ setEmployeeStore(empId, "", true); }catch(_e){}
    try{ setEmployeeOrder(empId, 0, true); }catch(_e){}
    try{ if(fromSid) normalizeEmployeeOrdersForStoreSilent(fromSid); }catch(_e){}
    try{ refreshEmployeeOrderPanel(); }catch(_e){}
    try{ applyEmployeeLeftFilter(); }catch(_e){}
    return;
  }

  // 同一店舗内の並び替え？
  const sameStoreReorder = (from==="right" && fromSid && fromSid===toSid);

  // 所属を更新（左→右 / 別店舗移動）
  if(!sameStoreReorder){
    try{ setEmployeeStore(empId, toSid, true); }catch(_e){}
    // 元店舗の欠番を詰める（右→右で別店舗移動したケース）
    try{ if(from==="right" && fromSid) normalizeEmployeeOrdersForStoreSilent(fromSid); }catch(_e){}
  }

  // 対象店舗の現状（実データ=左テーブルDOM）を取得し、ドラッグ対象は除外
  let ids = getEmpIdsInStoreFromTable(toSid).map(String).filter(id=>id!==empId);

  // 挿入位置：左→右の新規割当でも targetIndex（=placeholder）を尊重。取れなければ末尾。
  let idx = Number.isFinite(targetIndex) ? Number(targetIndex) : ids.length;
  if(idx < 0) idx = 0;
  if(idx > ids.length) idx = ids.length;

  ids.splice(idx, 0, empId);
  ids.forEach((id,i)=>{ try{ setEmployeeOrder(id, i+1, true); }catch(_e){} });

  try{ refreshEmployeeOrderPanel(); }catch(_e){}
  try{ applyEmployeeLeftFilter(); }catch(_e){}
}


/* v71: 従業員タブ（左表）の所属店舗変更時も店舗内順を再計算 */
function recalcEmployeeOrdersForStore(storeId){
  try{ if(window.__empDnD && __empDnD.isReordering) return; }catch(_e){}
  if(!dom || !dom.employeeTableBody) return;
  const sid = String(storeId||"__none__");

  // 対象店舗の行を集める
  const rows = Array.from(dom.employeeTableBody.querySelectorAll("tr[data-id]")).filter(tr=>{
    const sel = tr.querySelector("select.emp-store");
    if(!sel) return false;
    const v = String(sel.value||"__none__");
    return v===sid;
  });

  if(rows.length===0) return;

  // 現在の表示順(=店舗内順)を尊重しつつ、欠番を詰めて 1..n にする
  // 同順の場合は姓+名で安定ソート
  const getNameKey = (tr)=>{
    const last = (tr.querySelector(".emp-last")?.value || "").trim();
    const first = (tr.querySelector(".emp-first")?.value || "").trim();
    return (last + " " + first).trim();
  };
  rows.sort((a,b)=>{
    const oa = Number(a.querySelector(".emp-order")?.value || "0") || 0;
    const ob = Number(b.querySelector(".emp-order")?.value || "0") || 0;
    if(oa!==ob) return oa-ob;
    return getNameKey(a).localeCompare(getNameKey(b), "ja");
  });

  rows.forEach((tr, i)=>{
    const input = tr.querySelector(".emp-order");
    if(!input) return;
    const next = String(i+1);
    if(input.value !== next){
      input.value = next;
      // 既存の「右パネルへ即反映」ハンドラを活かす
      input.dispatchEvent(new Event("input", {bubbles:true}));
      input.dispatchEvent(new Event("change", {bubbles:true}));
    }
  });

  try{ if(typeof markDirty==="function") markDirty(); }catch(_e){}
}

function addEmployeeRow(emp){
  const master=getSettingsMasterTarget() || appState.templateMaster;
  // 追加ボタンからの新規作成時は、この場でIDを確定させる（DnD等で即時に参照できるように）
  // ※保存時に採番する方式だと dataset.id が空のままになり、左→右DnDが成立しない。
  if(!emp){
    if(!master.nextEmployeeId){
      master.nextEmployeeId = (master.employees||[]).reduce((mx,e)=>Math.max(mx, Number(e.id)||0),0) + 1;
    }
    emp = {
      id: master.nextEmployeeId++,
      lastName: "",
      firstName: "",
      storeId: null,
      orderInStore: 0
    };
  }
  if(emp.id==null || String(emp.id).trim()===""){
    emp.id = master.nextEmployeeId++;
  }
  const tr=document.createElement("tr");
  tr.setAttribute("data-id", String(emp.id));
  // Safari/Chromium系で td.draggable だけだと dragstart が発火しないケースがあるため、
  // 行(tr)自体も draggable にしておく（ただし開始判定はハンドル列で絞る）。
  tr.draggable = true;

  // DnD handle (left -> right: assign to store / move)
  const tdHandle=document.createElement("td");
  tdHandle.className="emp-drag-cell";
  const handle=document.createElement("span");
  handle.className="emp-drag-handle";
  handle.textContent="⋮⋮";
  handle.title="ドラッグで店舗に追加 / 移動";
  tdHandle.title="ドラッグで店舗に追加 / 移動";
  // 実際のドラッグ対象は tr だが、UX的にハンドル列も draggable にしておく
  tdHandle.draggable = true;
  // 新規作成時(empが未指定)は id が無いので空文字にする（ここで例外が出ると追加できなくなる）
  tdHandle.dataset.empId = emp.id;
  // ハンドル自体でもドラッグ開始できるようにする
  handle.draggable = true;
  tdHandle.appendChild(handle);

  const tdLast=document.createElement("td");
  const inputLast=document.createElement("input");
  inputLast.type="text"; inputLast.className="emp-last"; inputLast.value=emp?emp.lastName:"";
  tdLast.appendChild(inputLast);
  // 左の編集（姓）を右パネルへ即反映
  inputLast.addEventListener("input", ()=>{ refreshEmployeeOrderPanel(); });
  inputLast.addEventListener("change", ()=>{ refreshEmployeeOrderPanel(); });

  tr.appendChild(tdHandle);

  const tdFirst=document.createElement("td");
  const inputFirst=document.createElement("input");
  inputFirst.type="text"; inputFirst.className="emp-first"; inputFirst.value=emp?emp.firstName:"";
  tdFirst.appendChild(inputFirst);
  // 左の編集（名）を右パネルへ即反映
  inputFirst.addEventListener("input", ()=>{ refreshEmployeeOrderPanel(); });
  inputFirst.addEventListener("change", ()=>{ refreshEmployeeOrderPanel(); });

  const tdStore=document.createElement("td");
  const selectStore=document.createElement("select"); selectStore.className="emp-store";
  // 既存データの所属店舗を保持（refreshEmployeeStoreOptionsFromSettings で確定反映）
  if(emp){
    selectStore.dataset.desired = (emp.storeId!==undefined && emp.storeId!==null && emp.storeId!=="" ) ? String(emp.storeId) : "__none__";
  }
  selectStore.addEventListener("focus", ()=>{ 
    // 変更前の所属を保持（change時の再採番に使う）
    selectStore.dataset.prevStore = String(selectStore.value||"__none__");
  });
  selectStore.addEventListener("mousedown", ()=>{ 
    // focusが飛ばないブラウザ対策
    selectStore.dataset.prevStore = String(selectStore.value||"__none__");
  });
  selectStore.addEventListener("change", ()=>{
    const tr0 = selectStore.closest("tr");
    const empId0 = tr0 ? String(tr0.dataset.id||"") : "";

    const prev = String(selectStore.dataset.prevStore || "__none__");
    const next = String(selectStore.value||"__none__");
    // keep desired in sync so rebuild won't revert
    selectStore.dataset.desired = next;


// ★要望: 既存店舗へ「新規割当」したときは「必ず末席」に入れる（パターン依存を排除）
// DnD中は applyEmpDnDDrop 側で順番を決めるので、ここでは一切触らない（干渉防止）
const isDnD = !!(window.__empDnD && (__empDnD.isReordering || __empDnD.active));
if(!isDnD && empId0){
  try{
    // 元店舗（prev）は欠番が出るので詰め直し
    if(prev && prev!=="__none__") normalizeEmployeeOrdersForStoreSilent(prev);

    if(next && next!=="__none__"){
      // 新店舗（next）は「末席に追加」を確定させてから 1..n を振り直す
      appendEmployeeToStoreEnd(empId0, next);
    }else{
      // 未所属へ移動する際は店舗内順=0（仕様）
      setEmployeeOrder(empId0, 0, true);
    }
  }catch(_e){}
}else{
  // DnD中（またはempIdなし）は従来通り、必要なら別処理側で整形される
}// 次回の差分計算用
    selectStore.dataset.prevStore = next;

    refreshEmployeeOrderPanel();
    applyEmployeeLeftFilter();
  });
  const optNone=document.createElement("option"); optNone.value="__none__"; optNone.textContent="（未所属）";
  selectStore.appendChild(optNone);
    // 店舗候補は設定モーダル内の店舗タブDOM（未保存追加を含む）から都度反映する
  // （ここでは最低限の選択肢だけ作り、後で refreshEmployeeStoreOptionsFromSettings() で埋める）
  if(emp && emp.storeId) selectStore.value=String(emp.storeId);
  // option再構築で value が空になる瞬間があるため、desired を常に持たせておく
  try{ selectStore.dataset.desired = String(selectStore.value||"__none__"); }catch(_e){}
  try{ selectStore.dataset.prevStore = String(selectStore.value||"__none__"); }catch(_e){}
  tdStore.appendChild(selectStore);

  const tdOrder=document.createElement("td");
  const inputOrder=document.createElement("input");
  inputOrder.type="number";
  inputOrder.className="emp-order";
  inputOrder.readOnly = true;
  inputOrder.tabIndex = -1;
  // 内部は10刻みだが、ユーザー表示は1刻み（表示=内部/10）。入力は不要なのでreadonly。
  // 新規追加（empなし）の初期値は 0（仕様）
  inputOrder.value = emp && Number.isFinite(emp.orderInStore)
    ? String(Math.round((emp.orderInStore||0)/10))
    : "0";
  tdOrder.appendChild(inputOrder);
  // 左の編集（店舗内順）を右パネルへ即反映
  inputOrder.addEventListener("input", ()=>{ refreshEmployeeOrderPanel(); });
  inputOrder.addEventListener("change", ()=>{ refreshEmployeeOrderPanel(); });

  const tdDelete=document.createElement("td");
  const btnDel=document.createElement("button"); btnDel.className="btn-delete"; btnDel.textContent="削除";
  btnDel.addEventListener("click",()=>{
    const empId=emp?emp.id:null;
    if(empId && employeeHasScheduleData(empId)){
      alert("この従業員には勤務表データが存在するため削除できません。所属店舗を未所属にしてください。");
      return;
    }
    tr.remove();
    refreshEmployeeStoreOptionsFromSettings();
    refreshEmployeeOrderPanel();
  });
  tdDelete.appendChild(btnDel);

  tr.appendChild(tdLast); tr.appendChild(tdFirst); tr.appendChild(tdStore); tr.appendChild(tdOrder); tr.appendChild(tdDelete);
  dom.employeeTableBody.appendChild(tr);
  // 店舗タブの編集内容（未保存店舗追加/無効化など）を即反映
  refreshEmployeeStoreOptionsFromSettings();
  refreshEmployeeOrderPanel();
}

function employeeHasScheduleData(empId){
  for(const ym of Object.keys(appState.schedules)){
    const sch=appState.schedules[ym];
    for(const storeIdStr of Object.keys(sch.byStore||{})){
      const storeData=sch.byStore[Number(storeIdStr)];
      const rec=storeData.assignments?.[empId];
      if(!rec) continue;
      if((rec.main||[]).some(Boolean) || (rec.note||[]).some(Boolean)) return true;
    }
  }
  return false;
}

function addLeaveCodeRow(lc){
  const tr=document.createElement("tr");
  const tdCode=document.createElement("td");
  const inputCode=document.createElement("input");
  inputCode.type="text"; inputCode.className="leave-code"; inputCode.value=lc?lc.code:"";
  tdCode.appendChild(inputCode);

  const tdKey=document.createElement("td");
  const inputKey=document.createElement("input");
  inputKey.type="number"; inputKey.min="0"; inputKey.max="9"; inputKey.className="leave-key"; inputKey.value=lc?(lc.key||""):"";
  tdKey.appendChild(inputKey);

  const tdDel=document.createElement("td");
  const btnDel=document.createElement("button");
  btnDel.className = "btn-delete"; // 休暇コードタブも従業員タブと同じ削除ボタンスタイルに揃える
  btnDel.textContent="削除";
  btnDel.addEventListener("click",()=>tr.remove());
  tdDel.appendChild(btnDel);

  tr.appendChild(tdCode); tr.appendChild(tdKey); tr.appendChild(tdDel);
  dom.leaveCodeTableBody.appendChild(tr);
}

function refreshAfterMasterChange(){
  renderStoreOptions();
  renderSupportStoreOptions();
  renderKeypad();
  renderSchedule();
}

function reconcileScheduleDataWithSnapshot(schedule){
  const snapshot=schedule.masterSnapshot;
  const daysInMonth=getDaysInMonth(schedule.year, schedule.month);

  // 1) byStore の存在を保証（有効店舗 + 従業員が所属している店舗）
  const storeIds = new Set();
  snapshot.stores.forEach(s=>storeIds.add(s.id));
  snapshot.employees.forEach(e=>{ if(e.storeId) storeIds.add(e.storeId); });

  for(const storeId of storeIds){
    if(!schedule.byStore[storeId]){
      schedule.byStore[storeId]={assignments:{}};
    }
  }

  // 2) 従業員の所属変更・追加：assignments を移動/作成
  // まず現状の配置を収集
  const currentEmpStore = new Map(); // empId -> storeId
  for(const [storeIdStr, storeData] of Object.entries(schedule.byStore)){
    const storeId=Number(storeIdStr);
    for(const empIdStr of Object.keys(storeData.assignments||{})){
      currentEmpStore.set(Number(empIdStr), storeId);
    }
  }

  snapshot.employees.forEach(emp=>{
    const empId=emp.id;
    const targetStoreId=emp.storeId || null;

    const existingStoreId = currentEmpStore.has(empId) ? currentEmpStore.get(empId) : null;

    if(targetStoreId){
      // 移動 or 新規
      if(existingStoreId && existingStoreId!==targetStoreId){
        const from = schedule.byStore[existingStoreId];
        const to = schedule.byStore[targetStoreId] || (schedule.byStore[targetStoreId]={assignments:{}});
        to.assignments[empId] = from.assignments[empId];
        delete from.assignments[empId];
      }else if(!existingStoreId){
        const to = schedule.byStore[targetStoreId] || (schedule.byStore[targetStoreId]={assignments:{}});
        to.assignments[empId] = { main:new Array(daysInMonth).fill(""), note:new Array(daysInMonth).fill("") };
      }
      // 配列長の補正
      const rec = schedule.byStore[targetStoreId].assignments[empId];
      if(rec){
        if(!Array.isArray(rec.main)) rec.main=[];
        if(!Array.isArray(rec.note)) rec.note=[];
        rec.main.length = daysInMonth; rec.note.length = daysInMonth;
        for(let i=0;i<daysInMonth;i++){
          if(rec.main[i]===undefined) rec.main[i] = "";
          if(rec.note[i]===undefined) rec.note[i] = "";
        }
      }
    }else{
      // 未所属にした：assignments は残しても良いが、表示対象外になる
      // ただし、既存の assignments が別店舗に残っていればそのまま保持（データ保護）
    }
  });

  // 3) snapshotから削除された従業員：データが空なら assignments から消す、空でなければ保持（表示されないなら storeId=null を推奨）
  const snapshotEmpIds = new Set(snapshot.employees.map(e=>e.id));
  for(const [storeIdStr, storeData] of Object.entries(schedule.byStore)){
    for(const empIdStr of Object.keys(storeData.assignments||{})){
      const empId=Number(empIdStr);
      if(snapshotEmpIds.has(empId)) continue;
      const rec=storeData.assignments[empId];
      const hasData = (rec?.main||[]).some(v=>v) || (rec?.note||[]).some(v=>v);
      if(!hasData){
        delete storeData.assignments[empId];
      }
    }
  }

  persistState();
}

/* ===== Printing ===== */
function getStoreIdsInSchedule(schedule){
  const stores=schedule.masterSnapshot.stores.slice().sort((a,b)=>(a.order||0)-(b.order||0));
  const ids=[];
  stores.forEach(s=>{ if(schedule.byStore[s.id]) ids.push(s.id); });
  return ids;
}

function openPrintWindowForMonth(schedule, storeIds){
  const html=buildPrintHtml(schedule, storeIds);
  const w=window.open("", "_blank");
  if(!w){ alert("ポップアップがブロックされています。Edgeの設定で許可してください。"); return; }
  w.document.open();
  w.document.write(html);
  w.document.close();
  w.focus();
  setTimeout(()=>w.print(), 250);
}

function buildPrintHtml(schedule, storeIds){
  const {year,month}=schedule;
  const wk=toWareki(year,month);
  const daysInMonth=getDaysInMonth(year,month);
  const weekDays=["日","月","火","水","木","金","土"];

  const style = `
    <style>
      @page{size:A4 landscape;margin:8mm}
      body{font-family:system-ui,-apple-system,"Segoe UI","Yu Gothic UI","Meiryo",sans-serif;color:#111827}
      .page{page-break-after:always}
      .page:last-child{page-break-after:auto}
      .title{text-align:center;font-size:12.5pt;font-weight:800;letter-spacing:.06em;margin:0 0 6mm}
      .subtitle{text-align:center;font-size:11pt;font-weight:700;margin:-4mm 0 6mm;color:#374151}
      table{border-collapse:collapse;width:100%;table-layout:fixed}
      th,td{border:1px solid #d1d5db;text-align:center;padding:2mm 1mm;font-size:8.6pt}
      thead th{background:#f3f4f6;font-weight:800}
      .weekday{font-size:8pt;color:#6b7280;font-weight:800}
      .sun{color:#d32f2f}.sat{color:#1976d2}.hol{color:#d32f2f}
      .namecol{width:24mm;font-weight:800}
      .noteRow td{border-top-style:dashed;border-top-color:#9ca3af}
      .supportHead{background:#f9fafb;font-weight:800}
      .supportRow td{background:#fafafa}
      .supportFirst th,.supportFirst td{border-top:3px double #6b7280}
      .small{font-size:8pt;color:#6b7280;margin-top:2mm}
    
/* v107 layout fix */
.employee-settings-container{
  display: grid;
  grid-template-columns: minmax(420px, 520px) 1fr;
  gap: 16px;
  align-items: stretch;
}
.employee-left{
  max-width: 520px;
  overflow-x: auto;
}
.employee-right{
  min-width: 0;
}


/* v128+: left drag handle = 6-dot grip (no icon in column header) */
.employee-table-fixed td.emp-drag-cell{
  position:relative;
  user-select:none;
  text-align:center;
}
.employee-table-fixed .employee-table-fixed td.emp-drag-cell::before{
  content:"";
  position:absolute;
  left:50%;
  top:50%;
  width:14px;
  height:18px;
  transform:translate(-50%,-50%);
  background:
    radial-gradient(circle, rgba(107,114,128,.95) 1.2px, transparent 1.3px) 0 0 / 6px 6px,
    radial-gradient(circle, rgba(107,114,128,.95) 1.2px, transparent 1.3px) 3px 3px / 6px 6px;
  pointer-events:none;
  opacity:1;
}
/* span is kept for accessibility / grab cursor; visual is on the cell */
.employee-table-fixed td.emp-drag-cell .emp-drag-handle{
  /* v160: match row control height and keep icon centered */
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 100%;
  min-height: 32px;  /* matches input/select/button height (v158) */
  line-height: 1;
  cursor: grab;
}
.employee-table-fixed th:nth-child(1),
.employee-table-fixed td:nth-child(1){
  border-right-color: transparent;
}
.employee-table-fixed th:nth-child(2),
.employee-table-fixed td:nth-child(2){
  border-left-color: transparent;
}
}



/* v150: hard lock drag column to 28px */

#settingsEmployees table.settings-table.employee-table-fixed{
  table-layout: fixed !important;
}

/* 1列目 col を直接ロック（class を付与） */
#settingsEmployees table.settings-table.employee-table-fixed col.drag-col{
  width: 28px !important;
  min-width: 28px !important;
  max-width: 28px !important;
}

/* ヘッダ/セルも class でロック（列がズレても効く） */
#settingsEmployees table.settings-table.employee-table-fixed th.employee-drag-handle,
#settingsEmployees table.settings-table.employee-table-fixed td.emp-drag-cell{
  width: 28px !important;
  min-width: 28px !important;
  max-width: 28px !important;

  padding-left: 0 !important;
  padding-right: 0 !important;
  box-sizing: border-box !important;
  overflow: hidden !important;

  /* 何かが flex に変えても列幅計算を壊さない */
  display: table-cell !important;
}

/* flex レイアウトに巻き込まれた場合の保険（table外で使われた時も） */
#settingsEmployees table.settings-table.employee-table-fixed th.employee-drag-handle,
#settingsEmployees table.settings-table.employee-table-fixed td.emp-drag-cell{
  flex: 0 0 28px !important;
}

/* 見た目：セル内のハンドルを中央寄せ */
#settingsEmployees table.settings-table.employee-table-fixed td.emp-drag-cell .emp-drag-handle,
#settingsEmployees table.settings-table.employee-table-fixed th.employee-drag-handle .emp-drag-handle{
  /* v160: match row control height and keep icon centered */
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 100%;
  min-height: 32px;  /* matches input/select/button height (v158) */
  line-height: 1;
  cursor: grab;
}




/* v162: force thead row/cell height to match body row controls */
#settingsEmployees table.employee-table-fixed thead tr{
  height: 32px !important;      /* table row height works more reliably than cell pseudo content */
}
#settingsEmployees table.employee-table-fixed thead th{
  height: 32px !important;
  vertical-align: middle !important;
}
#settingsEmployees table.employee-table-fixed thead th.employee-drag-handle{
  padding: 0 !important;
}
/* keep the header drag cell from collapsing when empty */
#settingsEmployees table.employee-table-fixed thead th.employee-drag-handle{
  min-width: 28px !important;
}


/* v164: make header drag-handle cell match other header cell height (padding-driven) */
#settingsEmployees table.employee-table-fixed thead th.employee-drag-handle{
  /* other header cells use padding: 6px 10px; match vertical padding but keep horizontal tight */
  padding-top: 6px !important;
  padding-bottom: 6px !important;
  padding-left: 0 !important;
  padding-right: 0 !important;
  line-height: normal !important;
  vertical-align: middle !important;
}


/* v165: header drag cell should behave like other header cells (height/padding) */
#settingsEmployees table.employee-table-fixed thead th.employee-drag-handle{
  display: table-cell !important;
  padding: 6px 8px !important;   /* match .employee-table-fixed th,td default */
  cursor: default !important;
  user-select: none !important;
  vertical-align: middle !important;
}


/* v166: round table outer corners (employee tab left table) */
#settingsEmployees .employee-col-left{
  border-radius: 12px !important;
  overflow: hidden !important; /* clip inner borders to rounded corners */
}

/* border-collapse: separate is needed for border-radius to render reliably */
#settingsEmployees table.employee-table-fixed{
  border-collapse: separate !important;
  border-spacing: 0 !important;
  border-radius: 12px !important;
}

/* top corners */
#settingsEmployees table.employee-table-fixed thead tr:first-child th:first-child{
  border-top-left-radius: 12px !important;
}
#settingsEmployees table.employee-table-fixed thead tr:first-child th:last-child{
  border-top-right-radius: 12px !important;
}

/* bottom corners */
#settingsEmployees table.employee-table-fixed tbody tr:last-child td:first-child{
  border-bottom-left-radius: 12px !important;
}
#settingsEmployees table.employee-table-fixed tbody tr:last-child td:last-child{
  border-bottom-right-radius: 12px !important;
}

/* avoid background bleeding through rounded corners */
#settingsEmployees table.employee-table-fixed th,
#settingsEmployees table.employee-table-fixed td{
  background-clip: padding-box;
}


/* v167: rounded OUTER border via wrapper, avoid relying on cell borders at corners */
/* frame: table only (not the add button) */
#settingsEmployees .employee-table-frame{
  border: 1px solid var(--grid-border, #d7dde6) !important;
  border-radius: 12px !important;
  overflow: hidden !important;
  background: #fff !important;
}

/* remove the table's own outer border (use wrapper) */
#settingsEmployees .employee-table-frame table.employee-table-fixed{
  border-collapse: separate !important;
  border-spacing: 0 !important;
  width: 100% !important;
}

/* ensure cell borders remain, but drop borders on the outer edge so wrapper border is the outline */
#settingsEmployees .employee-table-frame table.employee-table-fixed thead tr:first-child th{
  border-top: 0 !important;
}
#settingsEmployees .employee-table-frame table.employee-table-fixed tbody tr:last-child td{
  border-bottom: 0 !important;
}
#settingsEmployees .employee-table-frame table.employee-table-fixed tr > :first-child{
  border-left: 0 !important;
}
#settingsEmployees .employee-table-frame table.employee-table-fixed tr > :last-child{
  border-right: 0 !important;
}


/* v168: draw rounded outer border as overlay to avoid "missing corner" from cell borders/sticky layers */
/* frame: table only (not the add button) */
#settingsEmployees .employee-table-frame{
  border: 0 !important;                 /* use overlay instead */
  border-radius: 12px !important;
  overflow: hidden !important;
  position: relative !important;
  background: #fff !important;
}
/* overlay border sits above table/cells so corners look rounded even with sticky headers */
#settingsEmployees .employee-table-frame::after{
  content: "";
  position: absolute;
  inset: 0;
  /* draw border fully INSIDE to avoid overflow clipping at corners */
  border-radius: 12px;
  box-shadow: inset 0 0 0 1px var(--grid-border, #d7dde6);
  pointer-events: none;
  z-index: 10;
}
/* keep table from painting outside */
#settingsEmployees .employee-table-frame table.employee-table-fixed{
  background: transparent !important;
}
/* ensure sticky header doesn't cover the border overlay */
#settingsEmployees .employee-table-frame thead th{
  z-index: 1 !important;
}

</style>`;

  const pages=storeIds.map(storeId=>{
    const store=schedule.masterSnapshot.stores.find(s=>s.id===storeId);
    if(!store) return "";
    const employees=getEmployeesInStore(schedule,storeId);
    const storeCodes=(store.storeCodes||[]).filter(c=>c.type!=="disabled");

    const thead = `
      <thead>
        <tr>
          <th class="namecol" rowspan="2">従業員</th>
          ${Array.from({length:daysInMonth},(_,i)=>{
            const d=i+1;
            const date=new Date(year,month-1,d);
            const w=date.getDay();
            const isHol=isJapaneseHoliday(date);
            const cls = isHol ? "hol" : (w===0 ? "sun" : w===6 ? "sat" : "");
            return `<th class="${cls}">${d}</th>`;
          }).join("")}
        </tr>
        <tr>
          ${Array.from({length:daysInMonth},(_,i)=>{
            const d=i+1;
            const date=new Date(year,month-1,d);
            const w=date.getDay();
            const isHol=isJapaneseHoliday(date);
            const cls = isHol ? "hol" : (w===0 ? "sun" : w===6 ? "sat" : "");
            return `<th class="weekday ${cls}">${weekDays[w]}</th>`;
          }).join("")}
        </tr>
      </thead>`;

    const bodyRows=employees.map(emp=>{
      const name=formatEmployeeName(emp.lastName, emp.firstName);
      const mainTds=Array.from({length:daysInMonth},(_,i)=>{
        const d=i+1;
        const v=getAssignmentValue(schedule,storeId,emp.id,"main",d)||"";
        return `<td>${escapeHtml(v)}</td>`;
      }).join("");
      const noteTds=Array.from({length:daysInMonth},(_,i)=>{
        const d=i+1;
        const v=getAssignmentValue(schedule,storeId,emp.id,"note",d)||"";
        return `<td>${escapeHtml(v)}</td>`;
      }).join("");
      return `
        <tr>
          <th class="namecol" rowspan="2">${escapeHtml(name)}</th>
          ${mainTds}
        </tr>
        <tr class="noteRow">
          ${noteTds}
        </tr>`;
    }).join("");

    const supportRows=storeCodes.filter(sc=>sc.type!=="disabled" && (sc.code==="E"||sc.code==="F")).map((sc,idx)=>{
      const code=sc.code;
      const label=code+"謹";
      const tds=Array.from({length:daysInMonth},(_,i)=>{
        const d=i+1;
        const supporter=getSupporterFor(schedule,storeId,code,d)||"";
        return `<td>${escapeHtml(supporter)}</td>`;
      }).join("");
      return `
        <tr class="supportRow ${idx===0 ? "supportFirst" : ""}">
          <th class="namecol supportHead">${escapeHtml(label)}</th>
          ${tds}
        </tr>`;
    }).join("");

    return `
      <div class="page">
        <div class="title">${escapeHtml(store.name)}　${escapeHtml(wk.gengo)}${escapeHtml(wk.yearStr)}年${month}月</div>
        <div class="subtitle">勤務予定表</div>
        <table>${thead}<tbody>${bodyRows}${supportRows}</tbody></table>
        <div class="small">※ 印刷ダイアログで「A4 横」「余白：標準（または最小）」を推奨します。</div>
      </div>`;
  }).join("");

  return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"/><title>${escapeHtml(year)}-${String(month).padStart(2,"0")} 勤務表</title>${style}
<style id="v186-employee-compact-28-and-delete-radius">
/* v186: 従業員タブも 28px 高さに統一（休暇コードタブと揃える） */
#settingsEmployees table.settings-table.employee-table-fixed input,
#settingsEmployees table.settings-table.employee-table-fixed select{
  height: 28px !important;
}

/* 従業員タブのテーブル内ボタンも 28px に（「削除」「追加」等を含む） */
#settingsEmployees table.settings-table.employee-table-fixed button{
  height: 28px !important;
  min-height: 28px !important;
  padding: 0 10px !important;  /* 上下paddingをゼロにして高さを抑える */
  line-height: 1 !important;
  box-sizing: border-box !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
}

/* 従業員タブ：ドラッグハンドルの高さも合わせる */
#settingsEmployees td .emp-drag-handle{
  min-height: 28px !important;
}

/* 従業員タブ：削除ボタンの角丸を休暇コードタブ（radius:10px）に合わせる */
#settingsEmployees .employee-table-frame .btn-delete,
#settingsEmployees .employee-table-frame button[data-action="delete-employee"]{
  border-radius: 10px !important;
  font-size: 12px !important;
}
</style>
</head><body>${pages}
  <div id="unassignHintOverlay" class="unassign-hint" style="--x:0px;--y:0px" aria-hidden="true"><span class="unassign-hint-dot"></span><span><b>外に放す</b>と未所属へ戻します</span></div>

</body></html>`;
}

/* ===== IME alert ===== */
function showIMEAlert(){
  dom.imeAlert.classList.add("visible");
  clearTimeout(imeAlertTimer);
  imeAlertTimer=setTimeout(()=>dom.imeAlert.classList.remove("visible"), 2500);
}

/* ===== Utilities ===== */
function deepClone(obj){ return JSON.parse(JSON.stringify(obj)); }
function lightenColor(hex, amount){
  const num=parseInt(hex.slice(1),16);
  let r=(num>>16)&0xff, g=(num>>8)&0xff, b=num&0xff;
  r=Math.round(r+(255-r)*amount);
  g=Math.round(g+(255-g)*amount);
  b=Math.round(b+(255-b)*amount);
  return "#"+((1<<24)+(r<<16)+(g<<8)+b).toString(16).slice(1);
}
function formatEmployeeName(lastName, firstName){
  const s=(lastName||"")+(firstName||"");
  if(s.length<=4) return (lastName||"")+"　"+(firstName||"");
  return (lastName||"")+(firstName||"");
}
function getDaysInMonth(year, month){ return new Date(year, month, 0).getDate(); }

/* ===== Japanese holiday (簡易) =====
   対応範囲: 2000年以降を主対象（令和中心）。2020/2021の五輪特例も考慮。
   ※ 法改正が入る年は都度アップデートが必要。
*/
const __JP_HOLIDAY_CACHE__ = new Map();
function pad2(n){ return String(n).padStart(2,"0"); }
function ymdKey(y,m,d){ return `${y}-${pad2(m)}-${pad2(d)}`; }
function nthMonday(year, month, nth){
  // month: 1-12
  const first = new Date(year, month-1, 1);
  const firstDow = first.getDay(); // 0 Sun
  const offset = (8 - firstDow) % 7; // days to first Monday
  const day = 1 + offset + (nth-1)*7;
  return day;
}
function vernalEquinoxDay(year){
  // 1980-2099 近似式
  return Math.floor(20.8431 + 0.242194*(year-1980) - Math.floor((year-1980)/4));
}
function autumnEquinoxDay(year){
  return Math.floor(23.2488 + 0.242194*(year-1980) - Math.floor((year-1980)/4));
}
function getJapaneseHolidays(year){
  if(__JP_HOLIDAY_CACHE__.has(year)) return __JP_HOLIDAY_CACHE__.get(year);

  const set = new Set();
  const add = (m,d)=>set.add(ymdKey(year,m,d));

  // Fixed
  add(1,1);   // 元日
  add(2,11);  // 建国記念の日
  if(year>=2020) add(2,23); // 天皇誕生日（令和）
  else if(year>=1989 && year<=2018) add(12,23); // 参考（平成）
  add(4,29);  // 昭和の日
  add(5,3); add(5,4); add(5,5); // 憲法/みどり/こども
  if(year>=2016) add(8,11); // 山の日（通常）
  add(11,3); add(11,23); // 文化/勤労感謝

  // Variable (Happy Monday)
  if(year>=2000) add(1, nthMonday(year,1,2)); // 成人の日
  if(year>=2003) add(7, nthMonday(year,7,3)); // 海の日
  if(year>=2003) add(9, nthMonday(year,9,3)); // 敬老の日
  if(year>=2000) add(10, nthMonday(year,10,2)); // スポーツの日

  // Equinox
  add(3, vernalEquinoxDay(year));
  add(9, autumnEquinoxDay(year));

  // 2020/2021 五輪特例（祝日移動）
  if(year===2020){
    set.delete(ymdKey(2020,7,nthMonday(2020,7,3)));
    set.delete(ymdKey(2020,10,nthMonday(2020,10,2)));
    set.delete(ymdKey(2020,8,11));
    add(7,23); add(7,24); add(8,10);
  }else if(year===2021){
    set.delete(ymdKey(2021,7,nthMonday(2021,7,3)));
    set.delete(ymdKey(2021,10,nthMonday(2021,10,2)));
    set.delete(ymdKey(2021,8,11));
    add(7,22); add(7,23); add(8,8);
  }

  // Substitute holiday (振替休日)
  const keys = Array.from(set).sort();
  for(const k of keys){
    const [y,m,d] = k.split("-").map(n=>parseInt(n,10));
    if(new Date(y, m-1, d).getDay()!==0) continue; // Sunday only
    let dt = new Date(y, m-1, d+1);
    while(true){
      const kk = ymdKey(dt.getFullYear(), dt.getMonth()+1, dt.getDate());
      if(!set.has(kk)){ set.add(kk); break; }
      dt = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()+1);
    }
  }

  // Citizens holiday (国民の休日)
  for(let month=1; month<=12; month++){
    const dim = new Date(year, month, 0).getDate();
    for(let day=1; day<=dim; day++){
      const k=ymdKey(year,month,day);
      if(set.has(k)) continue;
      const dow=new Date(year,month-1,day).getDay();
      if(dow===0 || dow===6) continue;
      const prev=new Date(year,month-1,day-1);
      const next=new Date(year,month-1,day+1);
      const pk=ymdKey(prev.getFullYear(), prev.getMonth()+1, prev.getDate());
      const nk=ymdKey(next.getFullYear(), next.getMonth()+1, next.getDate());
      if(set.has(pk) && set.has(nk)) set.add(k);
    }
  }

  __JP_HOLIDAY_CACHE__.set(year,set);
  return set;
}

/* ===== Holidays (master + auto-fetch) ===== */
function toYmd(date){
  const y=date.getFullYear();
  const m=String(date.getMonth()+1).padStart(2,"0");
  const d=String(date.getDate()).padStart(2,"0");
  return `${y}-${m}-${d}`;
}
function getHolidayInfoFromMaster(master, date){
  const key = (date instanceof Date) ? toYmd(date) : String(date||"");
  const holMap = master?.holidays || null;
  if(!holMap) return null;
  const ent = holMap[key];
  if(!ent) return null;
  return { date:key, name: ent.name||"", kind: ent.kind||"custom" };
}
async function fetchHolidaysFromHolidaysJP(year){
  // Prefer year-specific endpoint if available.
  const urlYear = `https://holidays-jp.github.io/api/v1/${year}/date.json`;
  const urlDefault = `https://holidays-jp.github.io/api/v1/date.json`;
  const controller = new AbortController();
  const t=setTimeout(()=>controller.abort(), 8000);
  try{
    let res = await fetch(urlYear, {signal: controller.signal, cache:"no-cache"});
    if(!res.ok) res = await fetch(urlDefault, {signal: controller.signal, cache:"no-cache"});
    if(!res.ok) throw new Error("holiday fetch failed");
    const data = await res.json();
    // data: { "YYYY-MM-DD": "name", ... }
    const map = {};
    Object.keys(data||{}).forEach(k=>{
      if(!k || !String(k).startsWith(String(year)+"-")) return;
      map[k]=String(data[k]||"");
    });
    return map;
  }finally{
    clearTimeout(t);
  }
}
function mergeHolidaysIntoMaster(master, year, fetchedMap){
  if(!master.holidays) master.holidays={};
  const hol = master.holidays;
  Object.keys(fetchedMap||{}).forEach(date=>{
    if(!String(date).startsWith(String(year)+"-")) return;
    const name=fetchedMap[date]||"";
    const existing=hol[date];
    if(existing && (existing.kind==="custom")) return; // user's custom wins
    if(existing && existing.name && existing.name.trim()) return; // don't overwrite edited name
    hol[date]={name, kind:"official"};
  });
}
async function ensureHolidayMasterForYear(master, year, opt){
  const forceFetch=!!(opt&&opt.forceFetch);
  if(!master.holidays) master.holidays={};
  const already = Object.keys(master.holidays).some(k=>String(k).startsWith(String(year)+"-"));
  if(already && !forceFetch) return;
  try{
    const fetched = await fetchHolidaysFromHolidaysJP(year);
    mergeHolidaysIntoMaster(master, year, fetched);
  }catch(e){
    console.warn("[holiday] fetch failed; fallback to local calc", e);
    // Fallback: use local calc for the year (existing isJapaneseHoliday)
    // We'll add only if there is no data for that year at all.
    const still = Object.keys(master.holidays).some(k=>String(k).startsWith(String(year)+"-"));
    if(!still){
      for(let m=0;m<12;m++){
        for(let d=1;d<=31;d++){
          const dt=new Date(year,m,d);
          if(dt.getMonth()!==m) continue;
          if(isJapaneseHoliday(dt)){
            const k=toYmd(dt);
            if(!master.holidays[k]) master.holidays[k]={name:"(祝日)", kind:"official"};
          }
        }
      }
    }
  }
}

/* ===== Holidays Settings UI ===== */

/* v55: Custom Modern Date Picker (Holiday tab)
   - display input holds ISO in data-iso for saving
   - visible .hol-date-display shows "YYYY/MM/DD (曜)"
*/
const __mdp = (() => {
  const weekdays = ["日","月","火","水","木","金","土"];
  const state = { pop:null, active:null, viewY:null, viewM:null };

  const pad2 = (n)=>String(n).padStart(2,"0");
  const toIso = (y,m,d)=>`${y}-${pad2(m)}-${pad2(d)}`;
  const toDisp = (iso)=>{
    if(!iso) return "";
    const [y,m,d]=iso.split("-").map(Number);
    const dt=new Date(y, (m||1)-1, d||1);
    const wd=weekdays[dt.getDay()]||"";
    return `${y}/${pad2(m)}/${pad2(d)} (${wd})`;
  };
  const clampPos = (x,y,w,h)=>{
    const vw = Math.max(320, window.innerWidth||320);
    const vh = Math.max(320, window.innerHeight||320);
    const nx = Math.min(Math.max(8, x), vw - w - 8);
    const ny = Math.min(Math.max(8, y), vh - h - 8);
    return {x:nx,y:ny};
  };

  function ensurePop(){
    if(state.pop) return state.pop;
    const pop=document.createElement("div");
    pop.className="mdp-pop";
    pop.style.display="none";
    pop.innerHTML = `
      <div class="mdp-head">
        <button type="button" class="mdp-nav" data-nav="prev" aria-label="前の月">‹</button>
        <div class="mdp-month" data-role="monthLabel"></div>
        <button type="button" class="mdp-nav" data-nav="next" aria-label="次の月">›</button>
      </div>
      <div class="mdp-grid" data-role="grid"></div>
      <div class="mdp-foot">
        <span class="mdp-mini" data-role="mini"></span>
        <button type="button" class="mdp-today" data-action="today">今日</button>
      </div>
    `;
    document.body.appendChild(pop);

    pop.addEventListener("click", (e)=>{
      const t=e.target;
      if(!(t instanceof HTMLElement)) return;
      const nav=t.closest("[data-nav]")?.getAttribute("data-nav");
      if(nav){
        if(nav==="prev"){
          state.viewM--; if(state.viewM<1){ state.viewM=12; state.viewY--; }
        }else{
          state.viewM++; if(state.viewM>12){ state.viewM=1; state.viewY++; }
        }
        render();
        return;
      }
      const todayBtn=t.closest('[data-action="today"]');
      if(todayBtn){
        const now=new Date();
        state.viewY = now.getFullYear();
        state.viewM = now.getMonth()+1;
        render();
        return;
      }
      const dayBtn=t.closest("[data-iso]");
      if(dayBtn && state.active){
        const iso=dayBtn.getAttribute("data-iso")||"";
        state.active.display.dataset.iso = iso;
        state.active.display.value = toDisp(iso);
        close();
        markDirtyIfNeeded();
      }
    });

    // close on outside click
    document.addEventListener("pointerdown",(e)=>{
      if(!state.active) return;
      const target=e.target;
      if(!(target instanceof Node)) return;
      const inPop = pop.contains(target);
      const inAnchor = state.active.wrap.contains(target);
      if(!inPop && !inAnchor) close();
    }, true);

    // close on esc
    document.addEventListener("keydown",(e)=>{
      if(e.key==="Escape") close();
    });

    state.pop=pop;
    return pop;
  }

  function markDirtyIfNeeded(){
    try{
      // keep existing "dirty" flow without new dependencies
      if(typeof markDirty==="function") markDirty();
    }catch(_){}
  }

  function render(){
    const pop=ensurePop();
    if(!state.active) return;

    const y=state.viewY, m=state.viewM;
    pop.querySelector('[data-role="monthLabel"]').textContent = `${y}年 ${m}月`;

    const grid=pop.querySelector('[data-role="grid"]');
    grid.innerHTML="";
    // weekday header
    for(const wd of weekdays){
      const el=document.createElement("div");
      el.className="mdp-wd";
      el.textContent=wd;
      grid.appendChild(el);
    }
    const first=new Date(y,m-1,1);
    const firstW=first.getDay(); // 0..6
    const daysInMonth=new Date(y,m,0).getDate();
    const daysPrevMonth=new Date(y,m-1,0).getDate();
    const selIso=(state.active.display.dataset.iso||"");
    const today=new Date();
    const todayIso=toIso(today.getFullYear(), today.getMonth()+1, today.getDate());

    // 6 weeks * 7 = 42 cells
    for(let i=0;i<42;i++){
      const cell=document.createElement("div");
      cell.className="mdp-day";
      const col=i%7;
      if(col===0) cell.classList.add("is-sun");
      if(col===6) cell.classList.add("is-sat");

      let dNum, iso, out=false;
      const dayIndex=i-firstW+1; // 1..daysInMonth within month
      if(dayIndex<=0){
        // prev month
        dNum = daysPrevMonth + dayIndex;
        const pm = m-1<=0 ? 12 : m-1;
        const py = m-1<=0 ? y-1 : y;
        iso = toIso(py, pm, dNum);
        out=true;
      }else if(dayIndex>daysInMonth){
        // next month
        dNum = dayIndex - daysInMonth;
        const nm = m+1>12 ? 1 : m+1;
        const ny = m+1>12 ? y+1 : y;
        iso = toIso(ny, nm, dNum);
        out=true;
      }else{
        dNum = dayIndex;
        iso = toIso(y,m,dNum);
      }

      cell.textContent=String(dNum);
      cell.dataset.iso=iso;
      if(out) cell.classList.add("is-out");
      if(iso===todayIso) cell.classList.add("is-today");
      if(selIso && iso===selIso) cell.classList.add("is-selected");

      grid.appendChild(cell);
    }

    const mini=pop.querySelector('[data-role="mini"]');
    mini.textContent = state.active.display.dataset.iso ? toDisp(state.active.display.dataset.iso) : "未選択";
  }

  function openFor(wrap, display){
    const pop = ensurePop();
    state.active = { wrap, display };

    const selected = (display.dataset.iso || "").trim();
    const base = selected ? new Date(selected + "T00:00:00") : new Date();
    // manual parse for safety
    let y, m;
    if(selected && /^\d{4}-\d{2}-\d{2}$/.test(selected)){
      const a = selected.split("-").map(Number);
      y = a[0]; m = a[1];
    }else{
      y = base.getFullYear(); m = base.getMonth() + 1;
    }
    state.viewY = y; state.viewM = m;

    // show (hidden) first, render to get correct size, then position on next frame
    const r = display.getBoundingClientRect();
    pop.style.display = "block";
    pop.style.visibility = "hidden";
    pop.style.left = "0px"; pop.style.top = "0px";

    render();

    const gap = 6;

    // IMPORTANT: First open can mis-measure height unless we wait a frame (layout not committed yet).
    requestAnimationFrame(() => {
      const pw = pop.offsetWidth || 0;
      const ph = pop.offsetHeight || 0;

      let left = Math.round(r.left);
      let topBelow = Math.round(r.bottom + gap);
      let topAbove = Math.round(r.top - ph - gap);
      let top = topBelow;

      // flip above if overflowing below
      if(ph && (topBelow + ph > window.innerHeight - 4) && topAbove >= 4) top = topAbove;

      // clamp to viewport
      if(pw) left = Math.min(Math.max(left, 4), window.innerWidth - pw - 4);
      if(ph) top  = Math.min(Math.max(top, 4), window.innerHeight - ph - 4);

      pop.style.left = left + "px";
      pop.style.top  = top  + "px";
      pop.style.visibility = "visible";
      pop.style.display = "block";
    });
  }


  function close(){
    if(!state.pop) return;
    state.pop.style.display="none";
    state.active=null;
  }

  function attach(wrap){
    const display=wrap.querySelector(".hol-date-display");
    if(!display) return;

    // Ensure initial display from data-iso (preferred) or current value
    if(!display.dataset.iso){
      const v=(display.value||"").trim();
      const m=v.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if(m){
        const y=Number(m[1]), mo=Number(m[2]), d=Number(m[3]);
        const dt=new Date(y, mo-1, d);
        if(dt.getFullYear()===y && (dt.getMonth()+1)===mo && dt.getDate()===d){
          display.dataset.iso = toIso(y,mo,d);
        }
      }
    }
    display.value = toDisp(display.dataset.iso || "");

    const toggleOpen = (e)=>{
      e?.preventDefault?.();
      if(state.active && state.active.wrap===wrap){
        close();
        return;
      }
      openFor(wrap, display);
    };

    // click toggles open/close
    display.addEventListener("click", toggleOpen);

    // keyboard support
    display.addEventListener("keydown", (e)=>{
      if(e.key==="Enter" || e.key===" "){
        toggleOpen(e);
      }else if(e.key==="Escape"){
        if(state.active && state.active.wrap===wrap){
          e.preventDefault();
          close();
        }
      }
    });
  }

  return { attach };
})();


function addHolidayRow(entry){
  const tr=document.createElement("tr");
  const dateVal = entry?.date || "";
  const nameVal = entry?.name || "";

  tr.innerHTML = `
    <td>
      <div class="hol-date-field mdp-field">
        <input class="hol-date-display mdp-display" type="text" value="${escapeHtml(formatHolidayDateDisplay(dateVal))}" readonly aria-label="日付" data-iso="${escapeHtml(dateVal)}">
        <svg class="calendar-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24" aria-hidden="true">
          <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
          <line x1="16" y1="2" x2="16" y2="6"></line>
          <line x1="8" y1="2" x2="8" y2="6"></line>
          <line x1="3" y1="10" x2="21" y2="10"></line>
        </svg>
      </div>
    </td>
    <td><input class="hol-name" type="text" value="${escapeHtml(nameVal)}" placeholder="例：臨時休業"></td>
    <td class="hol-del-cell"><button class="btn-delete" type="button" data-action="delete-holiday">削除</button></td>
  `;

  tr.querySelector('button[data-action="delete-holiday"]').addEventListener("click", ()=>{
    tr.remove();
    try{ if(typeof markDirty==="function") markDirty(); }catch(_){}
  });

  dom.holidayTableBody.appendChild(tr);

  // attach custom picker
  const wrap = tr.querySelector(".mdp-field");
  if(wrap) __mdp.attach(wrap);
}
function renderHolidaySettingsTable(){
  const master=getSettingsMasterTarget();
  if(!master || !dom.holidayTableBody) return;
  if(!master.holidays) master.holidays={};
  const year = Number(dom.holidayYearInput?.value)||null;
  dom.holidayTableBody.innerHTML="";
  if(!year) return;
  const rows = Object.keys(master.holidays)
    .filter(k=>String(k).startsWith(String(year)+"-"))
    .sort()
    .map(k=>({date:k, name:(master.holidays[k]?.n ?? master.holidays[k]?.name ?? master.holidays[k] ?? "")}));
  rows.forEach(r=>addHolidayRow(r));
}

function isJapaneseHoliday(date){
  const y=date.getFullYear();
  const m=date.getMonth()+1;
  const d=date.getDate();
  return getJapaneseHolidays(y).has(ymdKey(y,m,d));
}

function toWareki(year, month){
  if(year>2019 || (year===2019 && month>=5)){
    const n=year-2019+1;
    return {gengo:"令和", yearStr: n===1 ? "元" : String(n)};
  }
  const n=year-1989+1;
  return {gengo:"平成", yearStr:String(n)};
}
function nextMonthKey(ym){
  const [yStr,mStr]=ym.split("-");
  let y=parseInt(yStr,10), m=parseInt(mStr,10);
  m+=1; if(m>12){ m=1; y+=1; }
  return `${y.toString().padStart(4,"0")}-${m.toString().padStart(2,"0")}`;
}
function getEmployeesInStore(schedule, storeId){
  return schedule.masterSnapshot.employees
    .filter(e=>e.storeId===storeId)
    .sort((a,b)=>(a.orderInStore||9999)-(b.orderInStore||9999));
}
function escapeHtml(str){
  return String(str)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}