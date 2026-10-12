/* app.js — управление UI: заказы (localStorage), редактор листов, реестры, печать */
(function(){
const C=window.Calc, $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const LS='mest_webapp2';

/* ================= ХРАНИЛИЩЕ ================= */
let store={orders:[],r2:[],our:[],customers:[],dictMat:[],dictWork:[],dictMatSections:[]};

const DEFAULT_MAT_SECTIONS = [
  'Кабельная продукция',
  'Арматура и крепеж',
  'ВРУ и щитовое оборудование',
  'Светотехническая продукция',
  'Электроустановочные изделия',
  'Трубы и кабеленесущие системы',
  'Прочие материалы'
];

function classifyMatSection(name, fallback = 'Прочие материалы'){
  const k = (name || '').toLowerCase();
  if(/(кабель|провод|ввг|аввг|пув|пугв|шввп|витая пара|utp|ftp)/i.test(k)) return 'Кабельная продукция';
  if(/(вру|авр|щап|шкаф|щит|ящик гзш|ятп|щиток|панель вводная|панель распределительная)/i.test(k)) return 'ВРУ и щитовое оборудование';
  if(/(светильник|лампа|ламп|бра |лента светодиодная|патрон подвесной|дпо|дку|прожектор|led)/i.test(k)) return 'Светотехническая продукция';
  if(/(лоток|консоль|кронштейн потолочный|шпилька|анкер|гайка с насечкой|поворот на 45|поворот на 90)/i.test(k)) return 'Арматура и крепеж';
  if(/(труба |муфта труба|держатель для труб|винипластовая)/i.test(k)) return 'Трубы и кабеленесущие системы';
  if(/(розетка|выключатель|переключатель|рамка |подрозетник|коробка установочная|коробка распределительная|кнопка для звонков|звонок зуммер|разъем рш-вш)/i.test(k)) return 'Электроустановочные изделия';
  return fallback;
}

/* Справочники материалов и работ (аналог листов-баз Excel «Материалы»/«Работы»).
   Заполняются один раз из SEED.rows (1178 строк исходной книги), далее редактируются
   пользователем на вкладке «Справочник». */
function buildDictsFromSeed(){
  const m={},w={};
  SEED.rows.forEach(r=>{
    if(r.t==='Материалы'){
      const k=(r.n||'').trim().toLowerCase();
      if(k&&!m[k]) m[k]={n:r.n,u:r.u||'',p:r.p||0,sec:classifyMatSection(r.n)};
    }
    else if(r.t==='Работа'){
      const k=(r.n||'').trim().toLowerCase();
      if(k&&!w[k]) w[k]={n:r.n,u:r.u||'',p:r.p||0};
    }
  });
  return {mat:Object.values(m).sort((a,b)=>a.n.localeCompare(b.n,'ru')),
          work:Object.values(w).sort((a,b)=>a.n.localeCompare(b.n,'ru'))};
}
function r2ToOrder(r, idx){
  return {
    id: 'o-r2-' + idx,
    created: (r.dateRec && r.dateRec !== '00:00:00') ? r.dateRec : new Date().toISOString(),
    meta: {
      objectName: r.object || ('Заказ ' + (idx + 1)),
      org: r.form || 'МЭС',
      customerList: r.customer || '',
      dogovorDate: (r.ddate && r.ddate !== '00:00:00') ? r.ddate : '',
      stroyka: r.addrObj || '',
      obekt: r.object || '',
      dogovorNum: r.source || '',
      coefMat: 1,
      coefWork: 1,
      ks2date: r.dateRec || '',
      dates: [],
      status: Array(9).fill('----'),
      ndsMode: 'НДС 22%',
      perFrom: '',
      perTo: ''
    },
    rows: [
      { t: null, n: '1. ' + (r.object || 'Раздел'), u: null, q: null, p: null },
      { t: 'Материалы', n: 'Материалы по калькуляции', u: 'компл', q: 1, p: r.matPlan || 0 },
      { t: 'Работа', n: 'Монтажные и наладочные работы', u: 'компл', q: 1, p: r.uslPlan || 0 }
    ],
    vols: {},
    totals: {
      mat: r.matPlan || 0,
      work: r.uslPlan || 0,
      all: (r.planCost || (r.matPlan || 0) + (r.uslPlan || 0))
    },
    ks2: {
      total: r.closeSum || 0
    }
  };
}

function loadStore(){
  try{
    const s=JSON.parse(localStorage.getItem(LS));
    if(s&&typeof s==='object'){
      if(Array.isArray(s.orders)) store.orders=s.orders;
      if(Array.isArray(s.r2)) store.r2=s.r2;
      if(Array.isArray(s.our)&&s.our.length) store.our=s.our;
      if(Array.isArray(s.customers)&&s.customers.length) store.customers=s.customers;
      if(Array.isArray(s.dictMat)&&s.dictMat.length) store.dictMat=s.dictMat;
      if(Array.isArray(s.dictWork)&&s.dictWork.length) store.dictWork=s.dictWork;
      if(Array.isArray(s.dictMatSections)&&s.dictMatSections.length) store.dictMatSections=s.dictMatSections;
    }
  }catch(e){}
  if(!store.our || !store.our.length){
    store.our=(window.SEED?.our||[]).map(x=>({...x}));
  }
  if(!store.customers || !store.customers.length){
    store.customers=(window.SEED?.customers||[]).map(x=>({...x}));
  }
  if(!store.r2 || !store.r2.length){
    store.r2=(window.SEED?.r2||[]).map(x=>({...x}));
  }
  if(!store.orders || !store.orders.length){
    const mainOrder=seedOrder();
    const archiveOrders=(window.SEED?.r2||[]).map(r2ToOrder);
    store.orders=[mainOrder, ...archiveOrders];
  }
  if(!store.dictMat || !store.dictMat.length || !store.dictWork || !store.dictWork.length){
    const d=buildDictsFromSeed();
    if(!store.dictMat || !store.dictMat.length) store.dictMat=d.mat;
    if(!store.dictWork || !store.dictWork.length) store.dictWork=d.work;
  }
  if(!store.dictMatSections || !store.dictMatSections.length){
    store.dictMatSections = [...DEFAULT_MAT_SECTIONS];
  }
  // Убеждаемся, что у всех существующих материалов проставлен раздел
  if(store.dictMat && store.dictMat.length){
    store.dictMat.forEach(item => {
      if(!item.sec) item.sec = classifyMatSection(item.n);
      if(!store.dictMatSections.includes(item.sec)) store.dictMatSections.push(item.sec);
    });
  }
  save();
}
function seedOrder(){
  return {id:'o-seed',created:new Date().toISOString(),
    meta:{objectName:SEED.meta.objectName,org:SEED.meta.org,customerList:SEED.meta.customerList,
      source:'',dogovorDate:SEED.meta.dogovorDate,stroyka:SEED.meta.stroyka,obekt:SEED.meta.obekt,
      dogovorNum:SEED.meta.dogovorNum,coefMat:1,coefWork:1,ks2date:'2026-10-09',
      dates:SEED.meta.dates,status:SEED.meta.status,ndsMode:'НДС 22%',perFrom:'2026-09-08',perTo:'2026-10-09'},
    rows:JSON.parse(JSON.stringify(SEED.rows)),vols:JSON.parse(JSON.stringify(SEED.vols))};
}
function newBlankOrder(){
  return {id:'o-'+Date.now(),created:new Date().toISOString(),
    meta:{objectName:'',org:store.our[0]?.list||'МЭС',customerList:'',source:'',dogovorDate:'',
      stroyka:'',obekt:'',dogovorNum:'',coefMat:1,coefWork:1,
      ks2date:new Date().toISOString().slice(0,10),dates:[],status:Array(9).fill('----'),
      ndsMode:'НДС 22%',perFrom:'',perTo:''},
    rows:[{t:null,n:'1.Новый раздел',u:null,q:null,p:null},{t:'Материалы',n:'',u:'шт',q:0,p:0}],
    vols:{}};
}
let cur=null; // открытый заказ

/* ================= UNDO / REDO СИСТЕМА ================= */
const undoStack = [];
const redoStack = [];
const MAX_UNDO = 60;

function saveSnapshot(action = ''){
  if(!cur) return;
  try {
    undoStack.push({
      action,
      cur: JSON.parse(JSON.stringify(cur)),
      time: Date.now()
    });
    if(undoStack.length > MAX_UNDO) undoStack.shift();
    redoStack.length = 0; // Сброс повтора при новом действии
    updateUndoRedoUI();
  } catch(e){}
}

function doUndo(){
  if(!cur || !undoStack.length) {
    return toast('Нет действий для возврата');
  }
  try {
    redoStack.push({
      action: 'повтор',
      cur: JSON.parse(JSON.stringify(cur)),
      time: Date.now()
    });
    const item = undoStack.pop();
    cur = item.cur;
    syncCurToStore();
    save();
    renderCurrent();
    fillHeader();
    updateUndoRedoUI();
    toast(`Возврат: ${item.action || 'действие отменено'}`);
  } catch(e){
    toast('Ошибка возврата');
  }
}

function doRedo(){
  if(!cur || !redoStack.length) {
    return toast('Нет действий для повтора');
  }
  try {
    undoStack.push({
      action: 'возврат',
      cur: JSON.parse(JSON.stringify(cur)),
      time: Date.now()
    });
    const item = redoStack.pop();
    cur = item.cur;
    syncCurToStore();
    save();
    renderCurrent();
    fillHeader();
    updateUndoRedoUI();
    toast('Действие повторено');
  } catch(e){
    toast('Ошибка повтора');
  }
}

function syncCurToStore(){
  if(!cur) return;
  const idx = store.orders.findIndex(o => o.id === cur.id);
  if(idx >= 0) store.orders[idx] = cur;
}

function updateUndoRedoUI(){
  const canUndo = undoStack.length > 0;
  const canRedo = redoStack.length > 0;
  $$('.btn-sheet-undo').forEach(btn => {
    btn.disabled = !canUndo;
    btn.classList.toggle('disabled', !canUndo);
  });
  $$('.btn-sheet-redo').forEach(btn => {
    btn.disabled = !canRedo;
    btn.classList.toggle('disabled', !canRedo);
  });
}

// Горячие клавиши Ctrl+Z / Ctrl+Y (а также Ctrl+Shift+Z и русская раскладка)
// Модальное окно подтверждения (гарантированно работает в любых iframe без блокировок)
let confirmCallback = null;
function askConfirm(msg, onOk, title = 'Подтверждение удаления'){
  const modal = document.getElementById('modal-confirm');
  const msgEl = document.getElementById('modal-confirm-msg');
  const titleEl = document.getElementById('modal-confirm-title');
  const cancelBtn = document.getElementById('modal-confirm-cancel');
  if(!modal || !msgEl){
    if(window.confirm ? window.confirm(msg) : true) onOk();
    return;
  }
  if(titleEl) titleEl.textContent = title;
  msgEl.textContent = msg;
  confirmCallback = onOk;
  modal.classList.remove('hidden');
  if(cancelBtn) cancelBtn.focus();
}

function closeConfirmModal(){
  const modal = document.getElementById('modal-confirm');
  if(modal) modal.classList.add('hidden');
  confirmCallback = null;
}

// Модальное окно ввода текста (создание и переименование разделов)
let promptCallback = null;
function showPromptDialog(title, msg, defaultValue, onOk){
  const modal = document.getElementById('modal-prompt');
  const titleEl = document.getElementById('modal-prompt-title');
  const msgEl = document.getElementById('modal-prompt-msg');
  const inputEl = document.getElementById('modal-prompt-input');
  if(!modal || !inputEl){
    const res = window.prompt(msg, defaultValue);
    if(res !== null && res.trim()) onOk(res.trim());
    return;
  }
  if(titleEl) titleEl.textContent = title;
  if(msgEl) msgEl.textContent = msg;
  inputEl.value = defaultValue || '';
  promptCallback = onOk;
  modal.classList.remove('hidden');
  setTimeout(() => {
    inputEl.focus();
    inputEl.select();
  }, 50);
}

function closePromptModal(){
  const modal = document.getElementById('modal-prompt');
  if(modal) modal.classList.add('hidden');
  promptCallback = null;
}

document.addEventListener('click', e => {
  const cancelBtn = e.target.closest('#modal-confirm-cancel');
  if(cancelBtn){
    e.preventDefault();
    closeConfirmModal();
    return;
  }
  const okBtn = e.target.closest('#modal-confirm-ok');
  if(okBtn){
    e.preventDefault();
    if(confirmCallback){
      const cb = confirmCallback;
      closeConfirmModal();
      cb();
    }
    return;
  }
  const modal = document.getElementById('modal-confirm');
  if(e.target === modal){
    closeConfirmModal();
    return;
  }

  // Обработчики модального окна ввода (prompt)
  const pCancelBtn = e.target.closest('#modal-prompt-cancel');
  if(pCancelBtn){
    e.preventDefault();
    closePromptModal();
    return;
  }
  const pOkBtn = e.target.closest('#modal-prompt-ok');
  if(pOkBtn){
    e.preventDefault();
    const inputEl = document.getElementById('modal-prompt-input');
    const val = (inputEl ? inputEl.value : '').trim();
    if(promptCallback && val){
      const cb = promptCallback;
      closePromptModal();
      cb(val);
    } else {
      closePromptModal();
    }
    return;
  }
  const pModal = document.getElementById('modal-prompt');
  if(e.target === pModal){
    closePromptModal();
  }
});

document.addEventListener('keydown', e => {
  if(e.key === 'Escape'){
    const pModal = document.getElementById('modal-prompt');
    if(pModal && !pModal.classList.contains('hidden')){
      closePromptModal();
      return;
    }
    const modal = document.getElementById('modal-confirm');
    if(modal && !modal.classList.contains('hidden')){
      closeConfirmModal();
      return;
    }
    if(rowClipboard.length > 0){
      rowClipboard = [];
      const ci = document.getElementById('clip-info');
      if(ci) ci.textContent = '';
      renderMaterials();
      toast('Вставка отменена');
      return;
    }
  }

  if(e.key === 'Enter'){
    const pModal = document.getElementById('modal-prompt');
    if(pModal && !pModal.classList.contains('hidden')){
      const inputEl = document.getElementById('modal-prompt-input');
      const val = (inputEl ? inputEl.value : '').trim();
      if(promptCallback && val){
        e.preventDefault();
        const cb = promptCallback;
        closePromptModal();
        cb(val);
        return;
      }
    }
  }

  const isCtrl = e.ctrlKey || e.metaKey;
  if(!isCtrl) return;
  const key = e.key.toLowerCase();
  const code = e.code;
  if(!e.shiftKey && (key === 'z' || key === 'я' || code === 'KeyZ')){
    e.preventDefault();
    doUndo();
  } else if((key === 'y' || key === 'н' || code === 'KeyY') || (e.shiftKey && (key === 'z' || key === 'я' || code === 'KeyZ'))){
    e.preventDefault();
    doRedo();
  } else if(key === 'c' || key === 'с' || code === 'KeyC'){
    if(window.getSelection && window.getSelection().toString()) return;
    if(cur && selRows.size > 0){
      e.preventDefault();
      const idxs = [...selRows].sort((a,b)=>a-b);
      rowClipboard = idxs.map(j => JSON.parse(JSON.stringify(cur.rows[j])));
      selRows.clear();
      renderMaterials();
      toast(`Скопировано строк: ${rowClipboard.length}. Нажмите «Вставить» слева от нужной строки или Esc для отмены`);
    }
  }
});

// Клик по кнопкам отмены / повтора на любом листе
document.addEventListener('click', e => {
  const uBtn = e.target.closest('.btn-sheet-undo');
  if(uBtn){
    e.preventDefault();
    e.stopPropagation();
    doUndo();
    return;
  }
  const rBtn = e.target.closest('.btn-sheet-redo');
  if(rBtn){
    e.preventDefault();
    e.stopPropagation();
    doRedo();
    return;
  }
});

let saveTimer = null;
function debouncedSave(){
  if(saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    save();
  }, 600);
}

function save(){
  try{
    if(saveTimer){ clearTimeout(saveTimer); saveTimer = null; }
    localStorage.setItem(LS, JSON.stringify(store));
  }catch(e){
    console.error('Ошибка сохранения:', e);
  }
}

/* ================= НАВИГАЦИЯ ================= */
$$('.tab').forEach(b=>b.onclick=()=>{
  $$('.tab').forEach(x=>x.classList.remove('active')); b.classList.add('active');
  $$('.view').forEach(v=>v.classList.add('hidden'));
  const target = $('#view-'+b.dataset.view);
  if(target) target.classList.remove('hidden');
  if(b.dataset.view==='kontr') renderKontr();
  if(b.dataset.view==='orders') renderOrders();
  if(b.dataset.view==='sprav') renderSprav();
  if(b.dataset.view==='editor'){
    if(!cur && store.orders.length) cur = store.orders[0];
    if(cur){ fillHeader(); renderCurrent(); }
  }
});
$('#ed-tabs') && $$('.sheet-tab').forEach(b=>b.onclick=()=>{
  $$('.sheet-tab').forEach(x=>x.classList.remove('active')); b.classList.add('active');
  const keep=['materials','smeta','ks6','ks2','ks3'];
  keep.forEach(k=>$('#panel-'+k).classList.toggle('hidden',k!==b.dataset.sheet));
  renderSheet(b.dataset.sheet);
});
function openEditor(order){
  cur=order;
  undoStack.length = 0;
  redoStack.length = 0;
  updateUndoRedoUI();
  $('#view-orders').classList.add('hidden');
  $('#view-editor').classList.remove('hidden');
  $$('.tab').forEach(x=>x.classList.toggle('active',x.dataset.view==='editor'));
  const cloneBtn = $('#btn-clone-order');
  if(cloneBtn){
    cloneBtn.onclick = () => {
      const copy = cloneOrder(cur);
      if(copy) openEditor(copy);
    };
  }
  fillHeader();
  renderSheet($('.sheet-tab.active')?.dataset.sheet||'materials');
}

/* ================= ДУБЛИРОВАНИЕ ЗАКАЗА ================= */
function cloneOrder(sourceOrder){
  if(!sourceOrder) return;
  const cloned = JSON.parse(JSON.stringify(sourceOrder));
  cloned.id = 'o-' + Date.now();
  cloned.created = new Date().toISOString();
  cloned.meta.objectName = (cloned.meta.objectName || 'Заказ') + ' (копия)';
  store.orders.unshift(cloned);
  save();
  renderOrders();
  toast(`Заказ продублирован: ${cloned.meta.objectName}`);
  return cloned;
}

/* ================= СПИСОК ЗАКАЗОВ ================= */
function renderOrders(){
  const q=($('#order-search')?.value||'').toLowerCase();
  const tb=$('#orders-table tbody'); if(!tb) return; tb.innerHTML='';
  store.orders.filter(o=>o.meta.objectName.toLowerCase().includes(q)).forEach((o,i)=>{
    C.calcMaterials(o); C.calcKS2(o);
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${i+1}</td><td>${esc(o.meta.objectName)}</td><td>${esc(o.meta.org)}</td>
      <td>${esc(o.meta.customerList)}</td><td>${o.meta.dogovorDate||''}</td>
      <td class="num">${C.money(o.totals.all)}</td><td class="num">${C.money(o.ks2.total)}</td>
      <td>${o.created.slice(0,10)}</td>
      <td><button class="btn mini" data-open="${o.id}">Открыть</button>
          <button class="btn mini" data-clone="${o.id}" title="Создать копию заказа">⧉ Дублировать</button>
          <button class="btn mini danger" data-del="${o.id}">✕</button></td>`;
    tb.appendChild(tr);
  });
  tb.querySelectorAll('[data-open]').forEach(b=>b.onclick=()=>openEditor(store.orders.find(o=>o.id===b.dataset.open)));
  tb.querySelectorAll('[data-clone]').forEach(b=>b.onclick=()=>cloneOrder(store.orders.find(o=>o.id===b.dataset.clone)));
  tb.querySelectorAll('[data-del]').forEach(b=>b.onclick=()=>{
    const ord = store.orders.find(o=>o.id===b.dataset.del);
    const title = ord?.meta?.objectName || 'заказ';
    askConfirm(`Вы действительно хотите удалить заказ «${title}»?`, () => {
      store.orders=store.orders.filter(o=>o.id!==b.dataset.del);
      save();
      renderOrders();
      toast(`Заказ «${title}» удален`);
    }, 'Удаление заказа');
  });
}
$('#order-search').oninput=renderOrders;
$('#btn-new-order').onclick=()=>{const o=newBlankOrder();store.orders.push(o);save();openEditor(o);};

/* ================= РЕДАКТОР: шапка ================= */
function fillHeader(){
  $('#ed-title').textContent='Заказ: '+(cur.meta.objectName||'(без названия)');
  $('#f-object').value=cur.meta.objectName||'';
  $('#f-org').innerHTML=store.our.map(o=>`<option ${o.list===cur.meta.org?'selected':''}>${esc(o.list)}</option>`).join('');
  $('#f-customer').innerHTML='<option value="">—</option>'+store.customers.map(c=>`<option ${c.list===cur.meta.customerList?'selected':''}>${esc(c.list)}</option>`).join('');
  $('#f-stroyka').value=cur.meta.stroyka||''; $('#f-obekt').value=cur.meta.obekt||'';
  $('#f-coefm').value=cur.meta.coefMat??1; $('#f-coefw').value=cur.meta.coefWork??1;
  $('#hdr-coefm').value=cur.meta.coefMat??1; $('#hdr-coefw').value=cur.meta.coefWork??1;
  $('#f-ks2date').value=cur.meta.ks2date||''; $('#f-per-from').value=cur.meta.perFrom||''; $('#f-per-to').value=cur.meta.perTo||'';
  bindH('#f-object','objectName');bindH('#f-org','org');bindH('#f-customer','customerList');
  bindH('#f-stroyka','stroyka');bindH('#f-obekt','obekt');
  bindH('#f-ks2date','ks2date');bindH('#f-per-from','perFrom');bindH('#f-per-to','perTo');
  $('#f-coefm').onchange=e=>{cur.meta.coefMat=C.num(e.target.value);$('#hdr-coefm').value=e.target.value;recalcAll();};
  $('#f-coefw').onchange=e=>{cur.meta.coefWork=C.num(e.target.value);$('#hdr-coefw').value=e.target.value;recalcAll();};
  $('#hdr-coefm').onchange=e=>{cur.meta.coefMat=C.num(e.target.value);$('#f-coefm').value=e.target.value;recalcAll();};
  $('#hdr-coefw').onchange=e=>{cur.meta.coefWork=C.num(e.target.value);$('#f-coefw').value=e.target.value;recalcAll();};
}
function bindH(sel, key){
  const el = $(sel);
  if(el){
    el.onchange = e => {
      saveSnapshot('Изменение реквизитов заказа: ' + key);
      cur.meta[key] = e.target.value;
      recalcAllFast();
    };
  }
}

function recalcAll(){
  if(!cur) return;
  C.calcKS3(cur);
  ["materials","smeta","ks6"].forEach(n => {
    const el = document.getElementById("panel-" + n);
    if(el && !el.classList.contains("hidden")) renderSheet(n);
  });
  debouncedSave();
}

function recalcAllFast(){
  if(!cur) return;
  C.calcKS3(cur);
  ["smeta","ks6"].forEach(n => {
    const el = document.getElementById("panel-" + n);
    if(el && !el.classList.contains("hidden")) renderSheet(n);
  });
  debouncedSave();
}

function renderCurrent(){
  renderSheet($('.sheet-tab.active')?.dataset.sheet || 'materials');
}

/* ================= ЛИСТЫ ================= */
function renderSheet(name){
  if(!cur) return;
  if(name === 'materials') renderMaterials();
  if(name === 'smeta') renderSmeta();
  if(name === 'ks6'){ renderKs6(); renderObems(); }
  if(name === 'ks2') renderKs2Doc();
  if(name === 'ks3') renderKs3Doc();
}

/* ---- Материалы (редактируемая таблица заказа) ---- */
let rowClipboard = [];   // буфер копирования строк заказа
let selRows = new Set(); // выделенные строки (для копирования/удаления группы)

function rowsInDom(){
  return $$('#mat-table tbody tr').map(tr => {
    const el = tr.querySelector('[data-i]');
    return el ? +el.dataset.i : -1;
  }).filter(i => i >= 0);
}

function dictLookupAll(name){
  const k = (name || '').trim().toLowerCase();
  if(!k) return null;
  const workExact = (store.dictWork || []).find(d => (d.n || '').trim().toLowerCase() === k);
  const matExact = (store.dictMat || []).find(d => (d.n || '').trim().toLowerCase() === k);
  if(workExact && !matExact) return { type: 'Работа', item: workExact };
  if(matExact && !workExact) return { type: 'Материалы', item: matExact };
  if(workExact && matExact) return { type: 'Работа', item: workExact };
  if(k.length >= 4){
    const workPrefix = (store.dictWork || []).find(d => (d.n || '').trim().toLowerCase().startsWith(k));
    const matPrefix = (store.dictMat || []).find(d => (d.n || '').trim().toLowerCase().startsWith(k));
    if(workPrefix && !matPrefix) return { type: 'Работа', item: workPrefix };
    if(matPrefix && !workPrefix) return { type: 'Материалы', item: matPrefix };
  }
  return null;
}

function classifyTypeByName(name){
  const k = (name || '').trim().toLowerCase();
  if(!k) return 'Материалы';
  const inWork = (store.dictWork || []).some(d => (d.n || '').trim().toLowerCase() === k);
  if(inWork) return 'Работа';
  const inMat = (store.dictMat || []).some(d => (d.n || '').trim().toLowerCase() === k);
  if(inMat) return 'Материалы';
  const WORK_REGEX = /^(монтаж|демонтаж|установка|прокладка|подключение|сборка|устройство|пусконаладка|изготовление|затягивание|расключение|окраска|вмазывание|высверливание|штробление|сварка|пайка|испытани|прозвонк|наладк|ремонт|замена|пробивка|бурение)/i;
  if(WORK_REGEX.test(k)) return 'Работа';
  if(k.length >= 4){
    const pw = (store.dictWork || []).some(d => (d.n || '').trim().toLowerCase().startsWith(k));
    const pm = (store.dictMat || []).some(d => (d.n || '').trim().toLowerCase().startsWith(k));
    if(pw && !pm) return 'Работа';
    if(pm && !pw) return 'Материалы';
  }
  return 'Материалы';
}

function doInsertRow(at, type){
  if(!cur) return;
  if(at < 0) at = 0;
  if(at > cur.rows.length) at = cur.rows.length;
  let newRows = [];
  if(type === 'header'){
    saveSnapshot('Добавление строки: Заголовок');
    newRows = [{ t: null, sub: false, n: 'Новый раздел', u: null, q: null, p: null }];
  } else if(type === 'sub'){
    saveSnapshot('Добавление строки: Подзаголовок');
    newRows = [{ t: null, sub: true, n: 'Новый подзаголовок', u: null, q: null, p: null }];
  } else if(type === 'item'){
    saveSnapshot('Добавление строки: Позиция');
    newRows = [{ t: 'Материалы', n: '', u: 'шт', q: 1, p: 0 }];
  } else if(type === 'paste'){
    if(!rowClipboard.length) return toast('Буфер пуст');
    saveSnapshot('Вставка строк из буфера (' + rowClipboard.length + ' шт.)');
    newRows = rowClipboard.map(r => JSON.parse(JSON.stringify(r)));
    rowClipboard = []; // после вставки возвращаются кнопки Заг., под Заг., Поз.
    toast(`Вставлено строк: ${newRows.length}`);
  }
  cur.rows.splice(at, 0, ...newRows);
  renumberVols();
  recalcAllFast();
  renderMaterials();
  setTimeout(() => {
    const targetInp = $(`#mat-table [data-i="${at}"]`);
    if(targetInp){
      targetInp.focus();
      if(targetInp.select) targetInp.select();
    }
  }, 40);
}

function updateRowSumsInDom(tr, r){
  const km = cur.meta.coefMat || 1;
  const kw = cur.meta.coefWork || 1;
  r.sum = (r.q || 0) * (r.p || 0);
  r.I = (r.t === 'Материалы') ? r.sum * km : 0;
  r.J = (r.t === 'Работа') ? r.sum * kw : 0;
  if(tr && tr.cells && tr.cells.length >= 10){
    tr.cells[6].textContent = C.money(r.sum);
    tr.cells[7].textContent = C.money(r.I);
    tr.cells[8].textContent = C.money(r.J);
  }
}

function updateTotalsInDom(){
  if(!cur) return;
  C.calcMaterials(cur);
  const tm = $('#tot-mat'), tw = $('#tot-work'), ta = $('#tot-all'), ht = $('#hdr-total');
  if(tm) tm.textContent = C.money(cur.totals.mat);
  if(tw) tw.textContent = C.money(cur.totals.work);
  if(ta) ta.textContent = C.money(cur.totals.all);
  if(ht) ht.textContent = C.money(cur.totals.all);
}

function handleNameInput(inp, isCommit = false){
  const i = +inp.dataset.i;
  const r = cur.rows[i];
  if(!r) return;
  const val = inp.value;
  r.n = val;
  r.t = classifyTypeByName(val);
  const match = dictLookupAll(val);
  const tr = inp.closest('tr');
  if(match){
    r.t = match.type;
    if(match.item.u) r.u = match.item.u;
    if(match.item.p !== undefined && match.item.p !== null) r.p = match.item.p;
  }
  if(tr){
    const badge = tr.querySelector('.badge-type');
    if(badge){
      badge.className = `badge-type ${r.t === 'Работа' ? 'work' : 'mat'}`;
      badge.textContent = r.t;
    }
    if(match){
      const uSpan = tr.querySelector('[data-unit-i]');
      if(uSpan) uSpan.textContent = r.u || '—';
      const pInp = tr.querySelector('[data-f="p"]');
      if(pInp) pInp.value = r.p ?? '';
    }
    updateRowSumsInDom(tr, r);
  }
  if(match && isCommit){
    toast(`Из справочника (${r.t}): ${match.item.n} (${r.u}, ${C.money(r.p)} руб.)`);
  }
  updateTotalsInDom();
  debouncedSave();
}

function renderMaterials(){
  C.calcMaterials(cur);
  const tb = $('#mat-table tbody');
  if(!tb) return;

  const isPasteMode = rowClipboard.length > 0;
  const rowsHtml = [];

  if(!cur.rows.length){
    rowsHtml.push(`<tr>
      <td colspan="10" class="center" style="padding:24px;">
        <div style="margin-bottom:10px; color:var(--hint);">В заказе пока нет строк. Добавьте первую строку:</div>
        <div style="display:flex; gap:6px; justify-content:center;">
          ${isPasteMode ? `
            <button type="button" class="btn-seam-paste" data-paste-at="0">📋 Вставить скопированные строки (${rowClipboard.length} шт.)</button>
          ` : `
            <button type="button" class="btn btn-divider-add btn-add-hdr" data-add-at="0" data-type="header">📑 Заг.</button>
            <button type="button" class="btn btn-divider-add btn-add-sub" data-add-at="0" data-type="sub">📄 под Заг.</button>
            <button type="button" class="btn primary btn-divider-add btn-add-pos" data-add-at="0" data-type="item">＋ Поз.</button>
          `}
        </div>
      </td>
    </tr>`);
  } else {
    const totalRows = cur.rows.length;
    cur.rows.forEach((row, i) => {
      const isLast = (i === totalRows - 1);

      // Стык перед строкой i
      const seamBefore = `
        <div class="seam-controls" data-seam-at="${i}">
          ${isPasteMode ? `
            <button type="button" class="btn-seam-paste" data-paste-at="${i}" title="Вставить скопированные строки сюда">📋 Вставить</button>
          ` : `
            <div class="seam-btns-group">
              <button type="button" class="btn-divider-add btn-add-hdr" data-add-at="${i}" data-type="header" title="Вставить строку: Заголовок">Заг.</button>
              <button type="button" class="btn-divider-add btn-add-sub" data-add-at="${i}" data-type="sub" title="Вставить строку: под Заголовок">под Заг.</button>
              <button type="button" class="btn-divider-add btn-add-pos" data-add-at="${i}" data-type="item" title="Вставить строку: Позиция">Поз.</button>
            </div>
          `}
          <span class="seam-guide-line"></span>
        </div>`;

      // Стык после последней строки
      const seamAfter = isLast ? `
        <div class="seam-controls at-bottom" data-seam-at="${totalRows}">
          ${isPasteMode ? `
            <button type="button" class="btn-seam-paste" data-paste-at="${totalRows}" title="Вставить скопированные строки в самый конец">📋 Вставить</button>
          ` : `
            <div class="seam-btns-group">
              <button type="button" class="btn-divider-add btn-add-hdr" data-add-at="${totalRows}" data-type="header" title="Вставить Заголовок в конец">Заг.</button>
              <button type="button" class="btn-divider-add btn-add-sub" data-add-at="${totalRows}" data-type="sub" title="Вставить под Заголовок в конец">под Заг.</button>
              <button type="button" class="btn-divider-add btn-add-pos" data-add-at="${totalRows}" data-type="item" title="Вставить Позицию в конец">Поз.</button>
            </div>
          `}
          <span class="seam-guide-line"></span>
        </div>` : '';

      const gutterCell = `
        <td class="col-gutter">
          <div class="gutter-anchor">
            ${seamBefore}
            ${seamAfter}
            <input type="checkbox" class="pick" data-pick="${i}" ${selRows.has(i) ? 'checked' : ''}>
          </div>
        </td>`;

      if(!row.t){ // Заголовок раздела / подзаголовок
        rowsHtml.push(`<tr class="section-row${row.sub ? ' is-sub' : ' is-hdr'}${selRows.has(i) ? ' sel' : ''}" data-row-idx="${i}">
          ${gutterCell}
          <td class="center">
            <button type="button" class="btn-sec-toggle ${row.sub ? 'sub' : 'hdr'}" data-toggle-sec="${i}" title="Нажмите, чтобы переключить тип: Заголовок ↔ Подзаголовок">
              ${row.sub ? 'под Заг.' : 'Заг.'}
            </button>
          </td>
          <td colspan="7">
            <input class="sec-name${row.sub ? ' sub' : ''}" data-i="${i}" value="${esc(row.n || '')}" placeholder="${row.sub ? 'Название подзаголовка…' : 'Название раздела (заголовка)…'}">
          </td>
          <td class="center">
            <button type="button" class="btn mini danger" data-rm="${i}" title="Удалить строку">✕</button>
          </td>
        </tr>`);
      } else { // Позиция (Материалы или Работа)
        rowsHtml.push(`<tr class="${selRows.has(i) ? 'sel' : ''}" data-row-idx="${i}">
          ${gutterCell}
          <td>
            <div class="type-cell">
              <span class="badge-type ${row.t === 'Работа' ? 'work' : 'mat'}" title="Тип определен автоматически по наименованию">
                ${row.t === 'Работа' ? 'Работа' : 'Материалы'}
              </span>
            </div>
          </td>
          <td><input data-f="n" data-i="${i}" value="${esc(row.n || '')}" placeholder="Начните вводить материал или работу…"></td>
          <td class="center"><span class="unit-badge" data-unit-i="${i}" title="Ед.изм. выбирается из справочника">${esc(row.u || '—')}</span></td>
          <td><input data-f="q" data-i="${i}" class="num" value="${row.q ?? ''}"></td>
          <td><input data-f="p" data-i="${i}" class="num" value="${row.p ?? ''}"></td>
          <td class="num">${C.money(row.sum)}</td>
          <td class="num">${C.money(row.I)}</td>
          <td class="num">${C.money(row.J)}</td>
          <td class="center">
            <button type="button" class="btn mini danger" data-rm="${i}" title="Удалить строку">✕</button>
          </td>
        </tr>`);
      }
    });
  }

  tb.innerHTML = rowsHtml.join('');

  updateTotalsInDom();
  const ci = $('#clip-info');
  if(ci){
    ci.textContent = isPasteMode
      ? `Скопировано: ${rowClipboard.length} стр. (Нажмите «Вставить» слева от нужной строки или Esc для отмены)`
      : '';
  }
  const botNorm = $('#bottom-normal-actions');
  const botPaste = $('#bottom-paste-actions');
  if(botNorm && botPaste){
    botNorm.classList.toggle('hidden', isPasteMode);
    botPaste.classList.toggle('hidden', !isPasteMode);
  }
}

// Делегированные обработчики на mat-table tbody
const matTableTbody = $('#mat-table tbody');
if(matTableTbody){
  // Не открывать список datalist при простом клике или фокусе (только при вводе символов)
  const clearNameList = e => {
    const el = e.target;
    if(el && el.dataset && el.dataset.f === 'n'){
      el.removeAttribute('list');
    }
  };
  matTableTbody.addEventListener('pointerdown', clearNameList);
  matTableTbody.addEventListener('mousedown', clearNameList);
  matTableTbody.addEventListener('click', clearNameList);
  matTableTbody.addEventListener('focusin', clearNameList);
  matTableTbody.addEventListener('focusout', clearNameList);

  matTableTbody.onclick = e => {
    // 1. Добавление через кнопки (Заг., под Заг., Поз.)
    const divBtn = e.target.closest('.btn-divider-add');
    if(divBtn){
      e.stopPropagation();
      const at = divBtn.dataset.addAt === 'end' ? cur.rows.length : +divBtn.dataset.addAt;
      const type = divBtn.dataset.type;
      doInsertRow(at, type);
      return;
    }

    // 2. Вставка скопированных строк по кнопке "Вставить"
    const pasteBtn = e.target.closest('.btn-seam-paste, [data-paste-at]');
    if(pasteBtn){
      e.stopPropagation();
      const at = pasteBtn.dataset.pasteAt === 'end' ? cur.rows.length : +pasteBtn.dataset.pasteAt;
      doInsertRow(isNaN(at) ? cur.rows.length : at, 'paste');
      return;
    }

    // 3. Смена Заголовок ↔ Подзаголовок одним нажатием
    const secBtn = e.target.closest('[data-toggle-sec]');
    if(secBtn){
      e.stopPropagation();
      const i = +secBtn.dataset.toggleSec;
      if(cur && cur.rows[i]){
        saveSnapshot('Смена типа: ' + (cur.rows[i].sub ? 'Заголовок' : 'Подзаголовок'));
        cur.rows[i].sub = !cur.rows[i].sub;
        recalcAllFast();
        renderMaterials();
        toast(cur.rows[i].sub ? 'Переключено: под Заголовок' : 'Переключено: Заголовок');
      }
      return;
    }

    // 4. Удаление строки (✕) с модальным подтверждением (работает в любом iframe)
    const rmBtn = e.target.closest('[data-rm]');
    if(rmBtn){
      e.stopPropagation();
      const i = +rmBtn.dataset.rm;
      const isMulti = selRows.size > 1 && selRows.has(i);
      const rowName = cur.rows[i]?.n || (cur.rows[i]?.sub ? 'подзаголовок' : (cur.rows[i]?.t ? 'позицию' : 'раздел'));
      const msg = isMulti
        ? `Вы действительно хотите удалить выбранные строки (${selRows.size} шт.) из заказа?`
        : `Вы действительно хотите удалить строку «${rowName}» из заказа?`;
      askConfirm(msg, () => {
        saveSnapshot('Удаление строки');
        if(isMulti){
          cur.rows = cur.rows.filter((_, j) => !selRows.has(j));
          selRows.clear();
        } else {
          cur.rows.splice(i, 1);
          selRows.clear();
        }
        renumberVols();
        recalcAllFast();
        renderMaterials();
        toast('Строка удалена');
      }, 'Удаление строки заказа');
      return;
    }
  };

  matTableTbody.oninput = e => {
    const el = e.target;
    if(!el) return;
    if(el.dataset && el.dataset.f === 'n'){
      // Открывать список только когда пользователь начинает вводить символы:
      if(el.value.trim().length >= 1){
        if(el.getAttribute('list') !== 'dict-list-all'){
          el.setAttribute('list', 'dict-list-all');
        }
      } else {
        el.removeAttribute('list');
      }
      handleNameInput(el, false);
      return;
    }
    if(el.dataset && (el.dataset.f === 'q' || el.dataset.f === 'p')){
      const i = +el.dataset.i;
      const r = cur.rows[i];
      if(r){
        r[el.dataset.f] = C.num(el.value);
        const tr = el.closest('tr');
        if(tr) updateRowSumsInDom(tr, r);
        updateTotalsInDom();
        debouncedSave();
      }
      return;
    }
    if(el.classList.contains('sec-name')){
      const i = +el.dataset.i;
      if(cur.rows[i]){
        cur.rows[i].n = el.value;
        debouncedSave();
      }
      return;
    }
  };

  matTableTbody.onchange = e => {
    const el = e.target;
    if(!el) return;
    if(el.classList.contains('pick')){
      const i = +el.dataset.pick;
      el.checked ? selRows.add(i) : selRows.delete(i);
      el.closest('tr')?.classList.toggle('sel', el.checked);
      $('#sel-count').textContent = selRows.size ? `Выделено: ${selRows.size}` : '';
      return;
    }
    if(el.classList.contains('sec-name')){
      saveSnapshot('Название раздела');
      cur.rows[+el.dataset.i].n = el.value;
      recalcAllFast();
      return;
    }
    if(el.dataset && el.dataset.f === 'n'){
      saveSnapshot('Наименование позиции');
      handleNameInput(el, true);
      el.removeAttribute('list');
      recalcAllFast();
      return;
    }
    if(el.dataset && el.dataset.f !== undefined){
      saveSnapshot('Изменение поля ' + el.dataset.f);
      const i = +el.dataset.i, r = cur.rows[i], f = el.dataset.f;
      if(r){
        r[f] = (f === 'q' || f === 'p') ? C.num(el.value) : el.value;
        const tr = el.closest('tr');
        if(tr) updateRowSumsInDom(tr, r);
        updateTotalsInDom();
        recalcAllFast();
      }
      return;
    }
  };
}

// Кнопки добавления внизу таблицы материалов
const bottomActions = $('.table-bottom-actions');
if(bottomActions){
  bottomActions.onclick = e => {
    const btn = e.target.closest('.btn-divider-add');
    if(btn){
      e.stopPropagation();
      const type = btn.dataset.type;
      doInsertRow(cur ? cur.rows.length : 0, type);
      return;
    }
    const pasteBtn = e.target.closest('#btn-bottom-paste');
    if(pasteBtn){
      e.stopPropagation();
      doInsertRow(cur ? cur.rows.length : 0, 'paste');
      return;
    }
    const cancelBtn = e.target.closest('#btn-bottom-cancel-paste');
    if(cancelBtn){
      e.stopPropagation();
      rowClipboard = [];
      renderMaterials();
      toast('Вставка отменена');
      return;
    }
  };
}

const btnCutRows = document.getElementById('btn-cut-rows');
if(btnCutRows){
  btnCutRows.addEventListener('click', () => {
    if(!selRows.size) return toast('Выделите строки чекбоксами');
    askConfirm(`Вырезать выбранные строки (${selRows.size} шт.) в буфер?`, () => {
      saveSnapshot('Вырезание строк');
      rowClipboard = [...selRows].sort((a,b)=>a-b).map(j => JSON.parse(JSON.stringify(cur.rows[j])));
      cur.rows = cur.rows.filter((_, j) => !selRows.has(j));
      selRows.clear();
      renumberVols();
      recalcAllFast();
      renderMaterials();
      toast(`Вырезано строк: ${rowClipboard.length}. Нажмите «Вставить» слева от нужной строки или Esc`);
    }, 'Вырезание строк');
  });
}

const btnCopyRows = $('#btn-copy-rows');
if(btnCopyRows){
  btnCopyRows.onclick = () => {
    if(!selRows.size) return toast('Отметьте галочками строки для копирования');
    const idxs = [...selRows].sort((a,b)=>a-b);
    rowClipboard = idxs.map(j => JSON.parse(JSON.stringify(cur.rows[j])));
    selRows.clear();
    renderMaterials();
    toast(`Скопировано строк: ${rowClipboard.length}. Нажмите «Вставить» слева от нужной строки или Esc для отмены`);
  };
}

const btnPasteRows = document.getElementById('btn-paste-rows');
if(btnPasteRows){
  btnPasteRows.addEventListener('click', () => {
    if(!rowClipboard.length) return toast('Буфер пуст. Сначала скопируйте строки галочками');
    doInsertRow(cur ? cur.rows.length : 0, 'paste');
  });
}

function renumberVols(){ /* позиции после удаления/добавления пересчитываются при рендере */ }

/* ---- Смета ---- */
function renderSmeta(){
  C.calcSmeta(cur);
  const t=cur.smetaTotals;
  $('#sm-total').textContent=C.money(t.total);
  $('#sm-done').textContent=C.money(t.doneNds);
  $('#sm-ytd').textContent=C.money(t.ytd);
  const tb=$('#smeta-table tbody'); tb.innerHTML='';
  cur.smeta.forEach(x=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${x.pos}</td><td>${esc(x.name)}</td><td>${esc(x.u||'')}</td>
      <td class="num">${C.money(x.price)}</td><td class="num">${x.qty}</td>
      <td class="num">${C.money(x.costMat)}</td><td class="num">${C.money(x.costUsl)}</td>
      <td class="num">${x.doneAll}</td><td class="num">${C.money(x.doneSum)}</td>
      <td class="num">${x.rest}</td><td class="num">${C.money(x.restSum)}</td>`;
    tb.appendChild(tr);
  });
}

/* ---- Объемы закрытия (блоки под журналом КС-6) ---- */
function renderObems(){
  C.calcObems(cur);
  const box=$('#obems-blocks'); if(!box) return; box.innerHTML='';
  cur.obems.forEach(m=>{
    if(!m.rows.length && !m.date) return;
    const d=document.createElement('div'); d.className='obem-block';
    d.innerHTML=`<h4>${C.fmtMY(m.date)} <span class="tag ${m.status==='Закрытие'?'ok':''}">${m.status}</span> (${m.rows.length} поз.)</h4>
      <table class="grid compact"><thead><tr><th>№ позиции</th><th>Наименование</th><th>Объём</th><th>Остаток в смете</th></tr></thead>
      <tbody>${m.rows.map(r=>`<tr><td>${r.pos}</td><td>${esc(r.name)}</td><td class="num">${r.vol}</td><td class="num">${r.restQty}</td></tr>`).join('')}</tbody></table>`;
    box.appendChild(d);
  });
}

/* ---- КС-6 ----
   Выбор столбца редактирования и месяца/года — НЕ в отдельном поле, а прямо
   в шапке таблицы КС-6 (как строка дат F32:N32 в оригинальном листе Excel).
   Ровно ОДИН столбец может иметь статус «Закрытие», остальные — «----».
   В выбранном столбце дата отображается двумя выпадающими списками прямо
   в <th>: «Август» и «2026» (формат mmmm yyyy). Дублирующего блока настроек нет. */
function renderKs6(){
  C.calcSmeta(cur);
  const dogInp=$('#ks6-dogovor'), ddateInp=$('#ks6-ddate');
  if(dogInp){
    dogInp.value=cur.meta.dogovorNum||'';
    dogInp.onchange=e=>{cur.meta.dogovorNum=e.target.value;recalcAll();};
  }
  if(ddateInp){
    ddateInp.value=cur.meta.dogovorDate||'';
    ddateInp.onchange=e=>{cur.meta.dogovorDate=e.target.value;recalcAll();};
  }
  if(!cur.meta.mmyy) cur.meta.mmyy=[];      // миграция старых заказов: даты -> {m,y}
  C.months(cur).forEach(m=>{ if(!cur.meta.mmyy[m.i]) cur.meta.mmyy[m.i]={m:m.date.getMonth(),y:m.date.getFullYear()}; });
  const months=C.months(cur);

  const onHeaderChange=e=>{
    const t=e.target, i=+t.dataset.hi; if(!i&&i!==0) return;
    if(t.hasAttribute('data-st')){                       // переключение статуса столбца
      if(!cur.meta.status||cur.meta.status.length!==9) cur.meta.status=Array(9).fill('----');
      if(t.value==='Закрытие') cur.meta.status.fill('----');   // ровно один «Закрытие»
      cur.meta.status[i]=t.value;
    } else {                       // смена месяца / года в шапке (только в «Закрытии»)
      const mm=cur.meta.mmyy[i];
      const m=t.hasAttribute('data-mm')?+t.value:mm.m;
      const y=t.hasAttribute('data-yy')?+t.value:mm.y;
      C.setMonthYear(cur,i,m,y);   // при смене МЕСЯЦА последующие столбцы сдвигаются авто по порядку
      renderKs6();                 // перерисовать шапку: каскад виден сразу, без потери фокуса на году
      C.calcKS3(cur); renderSmeta(); renderObems(); save();
      return;
    }
    recalcAll();
  };
  $('#ks6-table thead').onchange=onHeaderChange;

  const P=C.positions(cur);
  const thead=$('#ks6-table thead');
  thead.innerHTML=`<tr><th>№ п/п</th><th>№ поз.</th><th>Наименование</th><th>Ед.изм.</th><th>Кол-во</th>`+
    months.map(m=>{
      const act=m.status==='Закрытие';
      const inner = act
        ? `<select data-hi="${m.i}" data-mm title="Месяц закрытия">${C.MONTH_NAMES.map((n,mi)=>`<option value="${mi}"${mi===m.date.getMonth()?' selected':''}>${n}</option>`).join('')}</select>
           <select data-hi="${m.i}" data-yy title="Год закрытия">${Array.from({length:8},(_,k)=>2023+k).map(y=>`<option${y===m.date.getFullYear()?' selected':''}>${y}</option>`).join('')}</select>`
        : `${C.fmtMY(m.date)}`;
      return `<th class="hmonth${act?' unlock':''}"><div class="hdate">${inner}</div>
        <select class="hstatus" data-hi="${m.i}" data-st title="Редактировать можно только один столбец">
          <option value="----"${!act?' selected':''}>----</option>
          <option value="Закрытие"${act?' selected':''}>Закрытие</option>
        </select></th>`;
    }).join('')+`<th>Остаток</th></tr>`;
  const tb=$('#ks6-table tbody'); tb.innerHTML='';
  cur.rows.forEach((row,ri)=>{
    if(!row.t) return;
    const p=P.find(x=>x.idx===ri); if(!p) return;
    const v=cur.vols[p.pos]||{};
    let sum=0; months.forEach(m=>sum+=C.num(v['m'+m.i]));
    const tr=document.createElement('tr');
    tr.innerHTML=`<td>${p.pos}</td><td>${p.pos}</td><td>${esc(row.n)}</td><td>${esc(row.u||'')}</td><td class="num">${C.num(row.q)}</td>`+
      months.map(m=>`<td><input class="num vol ${m.status==='Закрытие'?'':'locked'}" data-pos="${p.pos}" data-m="${m.i}" value="${v['m'+m.i]??''}" ${m.status==='Закрытие'?'':'disabled'}></td>`).join('')+
      `<td class="num ${sum>C.num(row.q)?'over':''}">${sum>C.num(row.q)?'Превышение':C.num(row.q)-sum}</td>`;
    tb.appendChild(tr);
  });
  tb.querySelectorAll('.vol').forEach(inp=>inp.onchange=e=>{
    const pos=e.target.dataset.pos,m=e.target.dataset.m,val=C.num(e.target.value);
    cur.vols=pos in cur.vols?cur.vols:{...cur.vols};
    if(!cur.vols[pos]) cur.vols[pos]={};
    if(val>0) cur.vols[pos]['m'+m]=val; else delete cur.vols[pos]['m'+m];
    recalcAll();
  });
}

/* ---- КС-2 документ ---- */
function renderKs2Doc(){
  C.calcKS2(cur);
  const k=cur.ks2, cust=(store.customers.find(c=>c.list===cur.meta.customerList)||{});
  const our=store.our.find(o=>o.list===cur.meta.org)||{};
  let rowsHtml='';
  k.lines.forEach(l=>{
    if(l.type==='head') rowsHtml+=`<tr class="doc-head"><td colspan="5"><b>${esc(l.name)}</b></td><td></td><td></td></tr>`;
    else rowsHtml+=`<tr><td>${l.pp}</td><td>${l.pos}</td><td>${esc(l.name)}</td><td>${esc(l.u||'')}</td>
      <td class="num">${l.vol}</td><td class="num">${C.money(l.price)}</td><td class="num">${C.money(l.sum)}</td></tr>`;
  });
  $('#ks2-doc').innerHTML=`
  <div class="doc-head-line">${C.fmtD(C.parseD(cur.meta.ks2date))} &nbsp;&nbsp; Унифицированная форма № КС-2, ОКУД 0322005</div>
  <div>Заказчик (Генподрядчик) — ${esc(cust.name||cur.meta.customerList)}</div>
  <div>Подрядчик (Субподрядчик) — ${esc(our.name||cur.meta.org)}</div>
  <div>Стройка — ${esc(cur.meta.stroyka)}</div><div>Объект — ${esc(cur.meta.obekt)}</div>
  <div>Договор подряда: ${esc(String(cur.meta.dogovorNum).replace(/\n/g,' '))} от ${cur.meta.dogovorDate||''}</div>
  <h3 class="center">АКТ № ${k.docNum} О ПРИЕМКЕ ВЫПОЛНЕННЫХ РАБОТ</h3>
  <div>Отчётный период: с ${C.fmtD(k.perFrom)} по ${C.fmtD(k.perTo)} &nbsp;|&nbsp; ${esc(k.restFormula)}</div>
  <table class="grid doc-table"><thead><tr><th>п.п</th><th>поз.по калькуляции</th><th>Наименование работ</th><th>Ед.изм</th>
   <th>Объём</th><th>Цена руб.</th><th>Стоимость руб.</th></tr></thead><tbody>${rowsHtml}</tbody>
   <tfoot><tr><td colspan="6">Итого, в том числе:</td><td class="num">${C.money(k.total)}</td></tr>
   <tr><td colspan="6">Работы</td><td class="num">${C.money(k.work)}</td></tr>
   <tr><td colspan="6">Материалы</td><td class="num">${C.money(k.mat)}</td></tr>
   <tr><td colspan="6"><b>Всего по акту</b></td><td class="num"><b>${C.money(k.total)}</b></td></tr>
   <tr><td colspan="6">${k.ndsLabel}</td><td class="num">${C.money(k.nds)}</td></tr></tfoot></table>
  <div class="signs"><div>Сдал: ${esc(our.sign||'')} — ${esc(our.post||'')} ${esc(our.fio||'')}</div>
  <div>Принял: ${esc(cust.sign||cust.name||'')} — ${esc(cust.post||'')} ${esc(cust.fio||'')}</div></div>`;
}

/* ---- КС-3 документ ---- */
function renderKs3Doc(){
  C.calcKS3(cur);
  const x=cur.ks3, cust=(store.customers.find(c=>c.list===cur.meta.customerList)||{});
  const our=store.our.find(o=>o.list===cur.meta.org)||{};
  $('#ks3-doc').innerHTML=`
  <div class="doc-head-line">Форма № КС-3, ОКУД 0322001 &nbsp; ${C.fmtD(C.parseD(cur.meta.ks2date))}</div>
  <div>Инвестор: ${esc(cust.name||'')}</div><div>Подрядчик: ${esc(our.name||'')}</div>
  <div>Контракт: ${esc(String(cur.meta.dogovorNum).replace(/\n/g,' '))} от ${cur.meta.dogovorDate||''}</div>
  <h3 class="center">СПРАВКА О СТОИМОСТИ ВЫПОЛНЕННЫХ РАБОТ И ЗАТРАТ № ${x.docNum}</h3>
  <div>Отчётный период: с ${x.perFrom} по ${x.perTo}</div>
  <table class="grid doc-table"><thead><tr><th>№ п/п</th><th>Наименование</th><th>с начала проведения работ</th>
   <th>с начала года</th><th>в т.ч. за отчётный период</th></tr></thead><tbody>
   <tr><td>1</td><td>Всего работ и затрат</td><td class="num">${C.money(x.fromStart)}</td><td class="num">${C.money(x.ytd)}</td><td class="num">${C.money(x.forPeriod-x.nds)}</td></tr>
   <tr><td>1.1</td><td>${esc(String(cur.meta.obekt).slice(0,120))}</td><td class="num">${C.money(x.fromStart)}</td><td class="num">${C.money(x.ytd)}</td><td class="num">${C.money(x.forPeriod-x.nds)}</td></tr>
   <tr><td></td><td><b>Итого</b></td><td></td><td></td><td class="num"><b>${C.money(x.forPeriod-x.nds)}</b></td></tr>
   <tr><td></td><td>${x.labelNds}</td><td></td><td></td><td class="num">${C.money(x.nds)}</td></tr>
   <tr><td></td><td><b>Всего с учетом НДС</b></td><td></td><td></td><td class="num"><b>${C.money(x.totalWithNds)}</b></td></tr>
   </tbody></table>
  <div class="signs"><div>Подрядчик: ${esc(our.sign||'')} ${esc(our.fio||'')}</div>
  <div>Заказчик (генподрядчик): ${esc(cust.sign||cust.name||'')}</div></div>`;
}

/* ---- Экспорт ---- */
function renderExport(){
  C.calcKS3(cur);
  const row=C.exportRow(cur);
  const map=[['Источник Заказа','source'],['Дата договора','ddate'],['Имя ОБЪЕКТА в РЕЕСТР','object'],
   ['Название Заказчика','customer'],['Предмет договора','predmet'],['Адрес объекта','addrObj'],
   ['Закрытие на Сумму','closeSum'],['ИНН/КПП','inn'],['Стоимость контракта (ПЛАН)','planCost'],
   ['Стоимость контракта (ФАКТ)','factCost'],['Материалы (ПЛАН)','matPlan'],['Услуги (ПЛАН)','uslPlan'],
   ['Версия программы','version'],['Путь к файлу','path'],['ДатаЗаписи','dateRec'],['Пользователь','user']];
  const etb = $('#export-table tbody');
  if(etb) etb.innerHTML=map.map(([h,k])=>`<tr><td><b>${h}</b></td><td>${esc(String(row[k]??''))}</td></tr>`).join('');
}

/* ---- кнопки тулбара редактора ---- */
const btnSaveOrder = $('#btn-save-order');
if(btnSaveOrder) btnSaveOrder.onclick = () => { save(); toast('Заказ сохранён'); };

const btnBackOrders = $('#btn-back-orders');
if(btnBackOrders) btnBackOrders.onclick = () => {
  $('#view-editor').classList.add('hidden');
  $('#view-orders').classList.remove('hidden');
  $$('.tab').forEach(x => x.classList.toggle('active', x.dataset.view === 'orders'));
  renderOrders();
};

const btnExportReestr = $('#btn-export-reestr');
if(btnExportReestr) btnExportReestr.onclick = () => {
  if(!cur) return;
  C.calcKS3(cur);
  const row = C.exportRow(cur);
  const ex = store.r2.find(r => r.object === row.object);
  if(ex) { Object.assign(ex, row); toast('Запись обновлена в Реестре'); }
  else { store.r2.push(row); toast('Новая запись добавлена в Реестр'); }
  save(); renderR2();
};

const btnPrintKs2 = $('#btn-print-ks2');
if(btnPrintKs2) btnPrintKs2.onclick = () => printDoc('ks2-doc');

const btnPrintKs3 = $('#btn-print-ks3');
if(btnPrintKs3) btnPrintKs3.onclick = () => printDoc('ks3-doc');

const btnPrintKs6 = $('#btn-print-ks6');
if(btnPrintKs6) btnPrintKs6.onclick = () => { renderKs6(); renderObems(); printArea('#panel-ks6'); };

const btnDownloadZip = $('#btn-download-zip');
if(btnDownloadZip) btnDownloadZip.onclick = () => {
  if(!cur) return;
  const blob = new Blob([JSON.stringify(cur, null, 1)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'Заказ_' + (cur.meta.objectName || 'new') + '.json';
  a.click();
};

const btnCloseOrder = $('#btn-close-order');
if(btnCloseOrder) btnCloseOrder.onclick = () => {
  $('#view-editor').classList.add('hidden');
  $('#view-orders').classList.remove('hidden');
  $$('.tab').forEach(x => x.classList.toggle('active', x.dataset.view === 'orders'));
  renderOrders();
};

function switchSheet(k){
  $$('.sheet-tab').forEach(b => b.classList.toggle('active', b.dataset.sheet === k));
  const keep = ['materials', 'smeta', 'ks6', 'ks2', 'ks3'];
  keep.forEach(p => {
    const el = $('#panel-' + p);
    if(el) el.classList.toggle('hidden', p !== k);
  });
  renderSheet(k);
}

function printDoc(id){
  window.print();
}
function printArea(sel){
  window.print();
}

/* ================= РЕЕСТР 2 (данные для совместимости) ================= */
const R2COLS=[['object','Имя ОБЪЕКТА'],['source','Источник'],['ddate','Дата договора'],['customer','Заказчик'],
 ['closeSum','Закрытие на сумму'],['planCost','Контракт ПЛАН'],['factCost','Контракт ФАКТ'],
 ['matPlan','Матер. ПЛАН'],['uslPlan','Усл. ПЛАН'],['inn','ИНН/КПП'],['kurator','Куратор'],
 ['version','Версия'],['dateRec','Дата записи'],['user','Польз.'],['path','Путь к файлу']];
function renderR2(){
  const searchEl = $('#r2-search');
  const countEl = $('#r2-count');
  const tableHead = $('#r2-table thead');
  const tb = $('#r2-table tbody');
  if(!tb) return;
  const q = (searchEl?.value || '').toLowerCase();
  if(countEl) countEl.textContent = store.r2.length;
  if(tableHead) tableHead.innerHTML = '<tr>' + R2COLS.map(c => `<th>${c[1]}</th>`).join('') + '</tr>';
  tb.innerHTML = '';
  store.r2.filter(r => !q || JSON.stringify(r).toLowerCase().includes(q)).slice(0, 500).forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = R2COLS.map(([k]) => {
      const v = r[k] ?? '';
      return `<td class="${typeof v === 'number' ? 'num' : ''}" title="${esc(String(v))}">${typeof v === 'number' ? C.money(v) : esc(String(v).slice(0, 60))}</td>`;
    }).join('');
    tb.appendChild(tr);
  });
}
const r2Search = $('#r2-search');
if(r2Search) r2Search.oninput = renderR2;
const btnR2Add = $('#btn-r2-add');
if(btnR2Add) btnR2Add.onclick = () => {
  store.r2.push({ object: 'Новый заказ', source: '', ddate: '', customer: '', closeSum: 0, planCost: 0, factCost: 0, matPlan: 0, uslPlan: 0, inn: '', kurator: '', version: 'web', dateRec: new Date().toISOString().slice(0, 10), user: 'web', path: '' });
  save();
  renderR2();
};

/* ================= КОНТРАГЕНТЫ ================= */
let kontrFilter = '';

function renderKontr(){
  const q = (kontrFilter || '').trim().toLowerCase();

  // 1. Наши организации (Подрядчик)
  const otb = $('#our-table tbody');
  if(otb){
    otb.innerHTML = '';
    const ourList = store.our || [];
    let shownOur = 0;
    ourList.forEach((o, i) => {
      const match = !q || [o.list, o.name, o.inn, o.kpp, o.addr, o.fio].some(v => String(v || '').toLowerCase().includes(q));
      if(!match) return;
      shownOur++;
      const tr = document.createElement('tr');
      const keys = ['list','name','inn','kpp','okpo','addr','sign','post','fio','nds'];
      tr.innerHTML = keys.map(k => {
        const val = esc(o[k] || '');
        const ph = k === 'list' ? 'Кратко' : k === 'name' ? 'Полное назв.' : k === 'inn' ? 'ИНН' : '';
        return `<td><input data-k="${k}" data-i="${i}" data-src="our" value="${val}" placeholder="${ph}"></td>`;
      }).join('') + `<td><button type="button" class="btn mini danger" data-del-o="${i}" title="Удалить">✕</button></td>`;
      otb.appendChild(tr);
    });
    const ourCount = $('#our-count');
    if(ourCount) ourCount.textContent = `${shownOur} / ${ourList.length}`;
  }

  // 2. Заказчики
  const ctb = $('#cust-table tbody');
  if(ctb){
    ctb.innerHTML = '';
    const custList = store.customers || [];
    let shownCust = 0;
    custList.forEach((o, i) => {
      const match = !q || [o.list, o.name, o.inn, o.kpp, o.addr, o.sign].some(v => String(v || '').toLowerCase().includes(q));
      if(!match) return;
      shownCust++;
      const tr = document.createElement('tr');
      const keys = ['list','name','inn','kpp','okpo','addr','sign'];
      tr.innerHTML = keys.map(k => {
        const val = esc(o[k] || '');
        const ph = k === 'list' ? 'Кратко' : k === 'name' ? 'Полное назв.' : k === 'inn' ? 'ИНН' : '';
        return `<td><input data-k="${k}" data-i="${i}" data-src="cust" value="${val}" placeholder="${ph}"></td>`;
      }).join('') + `<td><button type="button" class="btn mini danger" data-del-c="${i}" title="Удалить">✕</button></td>`;
      ctb.appendChild(tr);
    });
    const custCount = $('#cust-count');
    if(custCount) custCount.textContent = `${shownCust} / ${custList.length}`;
  }

  // Делегирование ввода контрагентов: моментальное сохранение и обновление списков
  $$('#our-table input, #cust-table input').forEach(inp => {
    inp.oninput = e => {
      const src = e.target.dataset.src;
      const i = +e.target.dataset.i;
      const k = e.target.dataset.k;
      const targetList = src === 'our' ? store.our : store.customers;
      if(targetList && targetList[i]){
        targetList[i][k] = e.target.value;
        save();
      }
    };
    inp.onchange = () => {
      save();
      if(cur) fillHeader();
    };
  });

  // Удаление
  $$('[data-del-o]').forEach(b => {
    b.onclick = () => {
      const i = +b.dataset.delO;
      const name = store.our[i]?.list || store.our[i]?.name || 'организацию';
      askConfirm(`Вы действительно хотите удалить организацию «${name}»?`, () => {
        store.our.splice(i, 1);
        save();
        renderKontr();
        if(cur) fillHeader();
        toast(`Удалена организация «${name}»`);
      }, 'Удаление организации');
    };
  });
  $$('[data-del-c]').forEach(b => {
    b.onclick = () => {
      const i = +b.dataset.delC;
      const name = store.customers[i]?.list || store.customers[i]?.name || 'заказчика';
      askConfirm(`Вы действительно хотите удалить заказчика «${name}»?`, () => {
        store.customers.splice(i, 1);
        save();
        renderKontr();
        if(cur) fillHeader();
        toast(`Удален заказчик «${name}»`);
      }, 'Удаление заказчика');
    };
  });
}

const btnOurAdd = $('#btn-our-add');
if(btnOurAdd){
  btnOurAdd.onclick = () => {
    store.our.unshift({ list: 'Новая орг.', name: '', inn: '', kpp: '', okpo: '', addr: '', sign: '', post: '', fio: '', nds: 'НДС 22%' });
    save();
    renderKontr();
    if(cur) fillHeader();
    toast('Добавлена новая организация');
    const firstInput = $('#our-table tbody tr input');
    if(firstInput) firstInput.focus();
  };
}

const btnCustAdd = $('#btn-cust-add');
if(btnCustAdd){
  btnCustAdd.onclick = () => {
    store.customers.unshift({ list: 'Новый заказчик', name: '', inn: '', kpp: '', okpo: '', addr: '', sign: '', post: '', fio: '', nds: '', header: '', ks6str: '' });
    save();
    renderKontr();
    if(cur) fillHeader();
    toast('Добавлен новый заказчик');
    const firstInput = $('#cust-table tbody tr input');
    if(firstInput) firstInput.focus();
  };
}

const kontrSearch = $('#kontr-search');
if(kontrSearch){
  kontrSearch.oninput = e => {
    kontrFilter = e.target.value;
    renderKontr();
  };
}

const btnKontrReset = $('#btn-kontr-reset');
if(btnKontrReset){
  btnKontrReset.onclick = () => {
    askConfirm('Восстановить организации и заказчиков из оригинальных данных книги?', () => {
      store.our = (window.SEED?.our || []).map(x => ({ ...x }));
      store.customers = (window.SEED?.customers || []).map(x => ({ ...x }));
      save();
      renderKontr();
      if(cur) fillHeader();
      toast('Контрагенты восстановлены из оригинала');
    }, 'Восстановление контрагентов');
  };
}

/* ================= СПРАВОЧНИК МАТЕРИАЛОВ И РАБОТ =================
   Два раздельных окна: Окно 1 — Материалы, Окно 2 — Работы.
   Переключаются вкладками вверху справочника. Каждое окно имеет собственный
   поиск, счётчик позиций, кнопку добавления и таблицу.
   При выборе любого наименования в заказе признак работы/материала определяется
   автоматически из этих справочников. */
let spravFilter = { m: '', w: '' };
let activeSpravWindow = 'mat'; // 'mat' | 'work'
let expandedMatSections = new Set(DEFAULT_MAT_SECTIONS);
let selSpravMat = new Set(); // индексы выбранных материалов в store.dictMat

function renderSprav(){
  const tabMCount = $('#sprav-tab-m-count');
  const tabWCount = $('#sprav-tab-w-count');
  if(tabMCount) tabMCount.textContent = store.dictMat.length;
  if(tabWCount) tabWCount.textContent = store.dictWork.length;

  if(!store.dictMatSections || !store.dictMatSections.length){
    store.dictMatSections = [...DEFAULT_MAT_SECTIONS];
  }
  // Убеждаемся, что у всех материалов есть раздел
  store.dictMat.forEach(item => {
    if(!item.sec) item.sec = classifyMatSection(item.n);
    if(!store.dictMatSections.includes(item.sec)) store.dictMatSections.push(item.sec);
  });

  // 1. Окно Материалов с разделами
  renderMatSections();

  // 2. Окно Работ
  renderWorkTable();
}

function renderMatSections(){
  const container = $('#sprav-mat-sections-container');
  if(!container) return;

  const qM = (spravFilter.m || '').trim().toLowerCase();
  const allSections = [...store.dictMatSections];
  store.dictMat.forEach(m => {
    if(m.sec && !allSections.includes(m.sec)) allSections.push(m.sec);
  });

  // Группируем материалы по разделам с сохранением исходного индекса
  const secMap = new Map();
  allSections.forEach(s => secMap.set(s, []));
  store.dictMat.forEach((item, origIdx) => {
    const sec = item.sec || 'Прочие материалы';
    if(!secMap.has(sec)) secMap.set(sec, []);
    secMap.get(sec).push({ item, origIdx });
  });

  let totalShownM = 0;
  const sectionsHtml = [];

  allSections.forEach(secName => {
    const allItemsInSec = secMap.get(secName) || [];
    const visibleItems = allItemsInSec.filter(({ item }) => {
      if(!qM) return true;
      return (item.n || '').toLowerCase().includes(qM) || secName.toLowerCase().includes(qM);
    });

    if(qM && visibleItems.length > 0){
      expandedMatSections.add(secName);
    }

    const isExpanded = expandedMatSections.has(secName);
    totalShownM += visibleItems.length;

    const allSelected = visibleItems.length > 0 && visibleItems.every(({ origIdx }) => selSpravMat.has(origIdx));

    let tableHtml = '';
    if(visibleItems.length === 0){
      tableHtml = `<div class="empty-sec-hint">${qM ? 'По вашему запросу в этом разделе ничего не найдено' : 'В этом разделе пока нет материалов. Нажмите «＋ Материал» для добавления.'}</div>`;
    } else {
      const rows = visibleItems.map(({ item, origIdx }, rowNum) => {
        const isChecked = selSpravMat.has(origIdx);
        const secOptions = allSections.map(s =>
          `<option value="${esc(s)}" ${s === secName ? 'selected' : ''}>${esc(s)}</option>`
        ).join('');

        return `<tr class="${isChecked ? 'sel' : ''}">
          <td style="text-align:center; width:34px;">
            <input type="checkbox" class="pick-sprav-mat" data-pick-mat="${origIdx}" ${isChecked ? 'checked' : ''}>
          </td>
          <td style="text-align:center; width:36px; color:var(--hint);">${rowNum + 1}</td>
          <td><input data-k="n" data-i="${origIdx}" data-src="dictMat" value="${esc(item.n)}" placeholder="Наименование материала…"></td>
          <td style="width:75px;"><input data-k="u" data-i="${origIdx}" data-src="dictMat" value="${esc(item.u || '')}" size="6" style="text-align:center;"></td>
          <td style="width:115px;"><input data-k="p" data-i="${origIdx}" data-src="dictMat" class="num" value="${item.p ?? ''}"></td>
          <td style="width:140px;">
            <select data-move-single="${origIdx}" class="mat-row-move-sec" title="Переместить материал в другой раздел">
              ${secOptions}
            </select>
          </td>
          <td style="width:110px; white-space:nowrap; text-align:center;">
            <button type="button" class="btn mini" data-addorder="m" data-i="${origIdx}" title="Добавить материал в открытый заказ">→ Заказ</button>
            <button type="button" class="btn mini danger" data-del="m" data-i="${origIdx}" title="Удалить материал">✕</button>
          </td>
        </tr>`;
      }).join('');

      tableHtml = `<div class="scroll-x">
        <table class="grid compact editable">
          <thead>
            <tr>
              <th style="width:34px; text-align:center;">✓</th>
              <th style="width:36px;">№</th>
              <th>Наименование материала</th>
              <th style="width:75px;">Ед.изм.</th>
              <th style="width:115px;">Цена руб. с НДС</th>
              <th style="width:140px;">Раздел</th>
              <th style="width:110px;">Действия</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
    }

    sectionsHtml.push(`
      <div class="sprav-sec-card" data-sec-card="${esc(secName)}">
        <div class="sprav-sec-header" data-sec-header="${esc(secName)}">
          <div class="sprav-sec-left">
            <button type="button" class="btn-sec-toggle-expand" data-toggle-sec-expand="${esc(secName)}" title="${isExpanded ? 'Свернуть раздел' : 'Развернуть раздел (+)'}">
              ${isExpanded ? '−' : '＋'}
            </button>
            <span class="sprav-sec-title">
              📁 ${esc(secName)}
              <span class="sec-badge-count">${allItemsInSec.length} поз.</span>
            </span>
          </div>
          <div class="sprav-sec-actions">
            ${visibleItems.length > 0 ? `
              <label class="sec-select-all" title="Выделить все материалы в этом разделе">
                <input type="checkbox" data-sec-all="${esc(secName)}" ${allSelected ? 'checked' : ''}> Выбрать все
              </label>
            ` : ''}
            <button type="button" class="btn mini btn-sec-act" data-sec-add-mat="${esc(secName)}" title="Добавить материал в этот раздел">＋ Материал</button>
            <button type="button" class="btn mini btn-sec-act" data-sec-rename="${esc(secName)}" title="Переименовать раздел">✎ Переименовать</button>
            <button type="button" class="btn mini danger btn-sec-act" data-sec-del="${esc(secName)}" title="${allItemsInSec.length > 0 ? 'Раздел содержит материалы и не может быть удален' : 'Удалить пустой раздел'}">✕ Удалить</button>
          </div>
        </div>
        <div class="sprav-sec-body ${isExpanded ? '' : 'collapsed'}" data-sec-body="${esc(secName)}">
          ${tableHtml}
        </div>
      </div>
    `);
  });

  container.innerHTML = sectionsHtml.join('');
  const countEl = $('#sprav-m-count');
  if(countEl) countEl.textContent = `${totalShownM} / ${store.dictMat.length}`;

  updateBulkBar(allSections);
}

function updateBulkBar(allSections){
  const bar = $('#sprav-mat-bulk-bar');
  if(!bar) return;
  const countEl = $('#sprav-mat-sel-count');
  const selSec = $('#sprav-mat-bulk-target-sec');

  if(selSpravMat.size > 0){
    bar.classList.remove('hidden');
    if(countEl) countEl.textContent = selSpravMat.size;
    if(selSec){
      selSec.innerHTML = allSections.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
    }
  } else {
    bar.classList.add('hidden');
  }
}

function renderWorkTable(){
  const qW = (spravFilter.w || '').toLowerCase();
  const tbW = $('#sprav-w-tbody');
  if(!tbW) return;
  tbW.innerHTML = '';
  let shownW = 0;
  store.dictWork.forEach((d, i) => {
    if(qW && !d.n.toLowerCase().includes(qW)) return;
    shownW++;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${i+1}</td>
      <td><input data-k="n" data-i="${i}" data-src="dictWork" value="${esc(d.n)}"></td>
      <td><input data-k="u" data-i="${i}" data-src="dictWork" value="${esc(d.u||'')}" size="6"></td>
      <td><input data-k="p" data-i="${i}" data-src="dictWork" class="num" value="${d.p??''}"></td>
      <td><button class="btn mini" data-addorder="w" data-i="${i}" title="Добавить работу в открытый заказ">→ Заказ</button>
          <button class="btn mini danger" data-del="w" data-i="${i}">✕</button></td>`;
    tbW.appendChild(tr);
  });
  const countWEl = $('#sprav-w-count');
  if(countWEl) countWEl.textContent = `${shownW} / ${store.dictWork.length}`;
}

// Делегированные обработчики событий для справочника материалов
const spravWinMat = $('#sprav-win-mat');
if(spravWinMat){
  // 1. Клики по кнопкам в окне материалов
  spravWinMat.addEventListener('click', e => {
    // Развернуть/свернуть раздел (+)
    const toggleBtn = e.target.closest('[data-toggle-sec-expand]');
    if(toggleBtn){
      e.stopPropagation();
      const secName = toggleBtn.dataset.toggleSecExpand;
      if(expandedMatSections.has(secName)){
        expandedMatSections.delete(secName);
      } else {
        expandedMatSections.add(secName);
      }
      renderMatSections();
      return;
    }

    // Добавить материал в конкретный раздел
    const addMatBtn = e.target.closest('[data-sec-add-mat]');
    if(addMatBtn){
      e.stopPropagation();
      const secName = addMatBtn.dataset.secAddMat;
      store.dictMat.unshift({ n: 'Новый материал', u: 'шт', p: 0, sec: secName });
      expandedMatSections.add(secName);
      save();
      fillDatalists();
      renderSprav();
      toast(`Материал добавлен в раздел «${secName}»`);
      return;
    }

    // Переименовать раздел
    const renameBtn = e.target.closest('[data-sec-rename]');
    if(renameBtn){
      e.stopPropagation();
      const oldName = renameBtn.dataset.secRename;
      showPromptDialog('Переименование раздела', `Введите новое название для раздела «${oldName}»:`, oldName, newName => {
        const trimmed = newName.trim();
        if(!trimmed || trimmed === oldName) return;
        const idx = store.dictMatSections.indexOf(oldName);
        if(idx >= 0) store.dictMatSections[idx] = trimmed;
        else store.dictMatSections.push(trimmed);
        store.dictMat.forEach(item => {
          if((item.sec || 'Прочие материалы') === oldName) item.sec = trimmed;
        });
        if(expandedMatSections.has(oldName)){
          expandedMatSections.delete(oldName);
          expandedMatSections.add(trimmed);
        }
        save();
        fillDatalists();
        renderSprav();
        toast(`Раздел переименован в «${trimmed}»`);
      });
      return;
    }

    // Удалить раздел (только если пустой!)
    const delSecBtn = e.target.closest('[data-sec-del]');
    if(delSecBtn){
      e.stopPropagation();
      const secName = delSecBtn.dataset.secDel;
      const inSec = store.dictMat.filter(m => (m.sec || 'Прочие материалы') === secName);
      if(inSec.length > 0){
        askConfirm(`Раздел «${secName}» содержит ${inSec.length} позиций и не может быть удален! Вы можете переименовать его (кнопка ✎), либо сначала переместить или удалить материалы из этого раздела.`, () => {}, 'Удаление невозможно');
        return;
      }
      askConfirm(`Вы действительно хотите удалить пустой раздел «${secName}»?`, () => {
        store.dictMatSections = store.dictMatSections.filter(s => s !== secName);
        expandedMatSections.delete(secName);
        save();
        renderSprav();
        toast(`Раздел «${secName}» удален`);
      }, 'Удаление пустого раздела');
      return;
    }

    // Удалить один материал (✕)
    const delMatBtn = e.target.closest('[data-del="m"]');
    if(delMatBtn){
      e.stopPropagation();
      const i = +delMatBtn.dataset.i;
      const name = store.dictMat[i]?.n || 'материал';
      askConfirm(`Вы действительно хотите удалить материал «${name}» из справочника?`, () => {
        store.dictMat.splice(i, 1);
        selSpravMat.delete(i);
        save();
        fillDatalists();
        renderSprav();
        toast('Материал удален');
      }, 'Удаление материала');
      return;
    }

    // Добавить один материал в открытый заказ (→ Заказ)
    const addOrderBtn = e.target.closest('[data-addorder="m"]');
    if(addOrderBtn){
      e.stopPropagation();
      if(!cur) return toast('Сначала откройте заказ (вкладка Заказы → Открыть)');
      const d = store.dictMat[+addOrderBtn.dataset.i];
      if(d){
        cur.rows.push({ t: 'Материалы', n: d.n, u: d.u, q: 1, p: d.p });
        recalcAll();
        switchSheet('materials');
        toast(`Добавлен в заказ: ${String(d.n).slice(0, 45)}`);
      }
      return;
    }
  });

  // 2. Изменения checkbox и select в окне материалов
  spravWinMat.addEventListener('change', e => {
    // Выбрать все в разделе
    if(e.target.dataset && e.target.dataset.secAll){
      const secName = e.target.dataset.secAll;
      const isChecked = e.target.checked;
      const qM = (spravFilter.m || '').trim().toLowerCase();
      store.dictMat.forEach((item, idx) => {
        if((item.sec || 'Прочие материалы') === secName){
          if(!qM || (item.n || '').toLowerCase().includes(qM)){
            isChecked ? selSpravMat.add(idx) : selSpravMat.delete(idx);
          }
        }
      });
      renderMatSections();
      return;
    }

    // Выбрать одну позицию
    if(e.target.dataset && e.target.dataset.pickMat !== undefined){
      const idx = +e.target.dataset.pickMat;
      e.target.checked ? selSpravMat.add(idx) : selSpravMat.delete(idx);
      renderMatSections();
      return;
    }

    // Перемещение одной позиции в другой раздел
    if(e.target.dataset && e.target.dataset.moveSingle !== undefined){
      const idx = +e.target.dataset.moveSingle;
      const targetSec = e.target.value;
      if(store.dictMat[idx] && targetSec){
        store.dictMat[idx].sec = targetSec;
        expandedMatSections.add(targetSec);
        save();
        renderSprav();
        toast(`Материал перемещен в раздел «${targetSec}»`);
      }
      return;
    }

    // Редактирование полей материала (наименование, ед.изм, цена)
    if(e.target.dataset && e.target.dataset.src === 'dictMat'){
      const i = +e.target.dataset.i;
      const k = e.target.dataset.k;
      const d = store.dictMat[i];
      if(d){
        d[k] = (k === 'p') ? C.num(e.target.value) : e.target.value;
        save();
        fillDatalists();
      }
      return;
    }
  });
}

// Обработчики тулбара материалов
const btnAddSec = $('#btn-sprav-add-sec');
if(btnAddSec){
  btnAddSec.onclick = () => {
    showPromptDialog('Создать раздел', 'Введите название нового раздела материалов (например: Кабельная продукция, Арматура, ВРУ и др.):', '', name => {
      const trimmed = name.trim();
      if(!trimmed) return;
      if(store.dictMatSections.includes(trimmed)){
        return toast(`Раздел «${trimmed}» уже существует`);
      }
      store.dictMatSections.push(trimmed);
      expandedMatSections.add(trimmed);
      save();
      renderSprav();
      toast(`Раздел «${trimmed}» успешно создан`);
    });
  };
}

const btnExpandAll = $('#btn-sprav-expand-all');
if(btnExpandAll){
  btnExpandAll.onclick = () => {
    (store.dictMatSections || []).forEach(s => expandedMatSections.add(s));
    renderMatSections();
  };
}

const btnCollapseAll = $('#btn-sprav-collapse-all');
if(btnCollapseAll){
  btnCollapseAll.onclick = () => {
    expandedMatSections.clear();
    renderMatSections();
  };
}

// Групповые действия в тулбаре материалов
const btnBulkMove = $('#btn-sprav-mat-bulk-move');
if(btnBulkMove){
  btnBulkMove.onclick = () => {
    if(!selSpravMat.size) return;
    const targetSec = $('#sprav-mat-bulk-target-sec')?.value;
    if(!targetSec) return;
    selSpravMat.forEach(idx => {
      if(store.dictMat[idx]) store.dictMat[idx].sec = targetSec;
    });
    expandedMatSections.add(targetSec);
    const count = selSpravMat.size;
    selSpravMat.clear();
    save();
    renderSprav();
    toast(`Перемещено ${count} поз. в раздел «${targetSec}»`);
  };
}

const btnBulkDel = $('#btn-sprav-mat-bulk-del');
if(btnBulkDel){
  btnBulkDel.onclick = () => {
    if(!selSpravMat.size) return;
    const count = selSpravMat.size;
    askConfirm(`Вы действительно хотите удалить выбранные материалы (${count} шт.) из справочника?`, () => {
      store.dictMat = store.dictMat.filter((_, idx) => !selSpravMat.has(idx));
      selSpravMat.clear();
      save();
      fillDatalists();
      renderSprav();
      toast(`Удалено ${count} материалов`);
    }, 'Удаление материалов');
  };
}

const btnBulkUncheck = $('#btn-sprav-mat-bulk-uncheck');
if(btnBulkUncheck){
  btnBulkUncheck.onclick = () => {
    selSpravMat.clear();
    renderMatSections();
  };
}

// Обработчики окна Работ
const spravWinWork = $('#sprav-win-work');
if(spravWinWork){
  spravWinWork.addEventListener('change', e => {
    if(e.target.dataset && e.target.dataset.src === 'dictWork'){
      const i = +e.target.dataset.i;
      const k = e.target.dataset.k;
      const d = store.dictWork[i];
      if(d){
        d[k] = (k === 'p') ? C.num(e.target.value) : e.target.value;
        save();
        fillDatalists();
      }
    }
  });

  spravWinWork.addEventListener('click', e => {
    const delWorkBtn = e.target.closest('[data-del="w"]');
    if(delWorkBtn){
      e.stopPropagation();
      const i = +delWorkBtn.dataset.i;
      const name = store.dictWork[i]?.n || 'работу';
      askConfirm(`Вы действительно хотите удалить работу «${name}» из справочника?`, () => {
        store.dictWork.splice(i, 1);
        save();
        fillDatalists();
        renderSprav();
        toast('Работа удалена');
      }, 'Удаление работы');
      return;
    }

    const addOrderBtn = e.target.closest('[data-addorder="w"]');
    if(addOrderBtn){
      e.stopPropagation();
      if(!cur) return toast('Сначала откройте заказ (вкладка Заказы → Открыть)');
      const d = store.dictWork[+addOrderBtn.dataset.i];
      if(d){
        cur.rows.push({ t: 'Работа', n: d.n, u: d.u, q: 1, p: d.p });
        recalcAll();
        switchSheet('materials');
        toast(`Добавлена в заказ: ${String(d.n).slice(0, 45)}`);
      }
      return;
    }
  });
}

// Переключение окон справочника (Материалы / Работы)
$$('.sprav-tab-btn').forEach(btn => {
  btn.onclick = () => {
    $$('.sprav-tab-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    activeSpravWindow = btn.dataset.spravView;
    const isMat = activeSpravWindow === 'mat';
    $('#sprav-win-mat').classList.toggle('hidden', !isMat);
    $('#sprav-win-work').classList.toggle('hidden', isMat);
  };
});

$('#btn-sprav-add-m').onclick=()=>{
  const targetSec = store.dictMatSections[0] || 'Кабельная продукция';
  store.dictMat.unshift({n:'Новый материал',u:'шт',p:0,sec:targetSec});
  expandedMatSections.add(targetSec);
  save(); fillDatalists(); renderSprav();
  toast(`Материал добавлен в раздел «${targetSec}»`);
};
$('#btn-sprav-add-w').onclick=()=>{
  store.dictWork.unshift({n:'Новая работа',u:'шт',p:0});
  save(); fillDatalists(); renderSprav();
  toast('Работа добавлена в справочник');
};
$('#sprav-search-m').oninput=e=>{spravFilter.m=e.target.value.toLowerCase();renderSprav();};
$('#sprav-search-w').oninput=e=>{spravFilter.w=e.target.value.toLowerCase();renderSprav();};
$('#btn-sprav-reset').onclick=()=>{
  askConfirm('Перезаполнить справочники из данных исходной Excel-книги? Ваши правки будут потеряны.', () => {
    const d=buildDictsFromSeed(); store.dictMat=d.mat; store.dictWork=d.work;
    store.dictMatSections=[...DEFAULT_MAT_SECTIONS];
    save(); fillDatalists(); renderSprav();
    toast('Справочники восстановлены из оригинала');
  }, 'Восстановление справочников');
};

/* ================= utils ================= */
function esc(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

/* Единый datalist со всеми материалами и работами для автоподстановки в заказе:
   при выборе из списка тип (Материал или Работа), ед.изм. и цена проставляются автоматически */
function fillDatalists(){
  const allMap = new Map();
  (store.dictWork || []).forEach(d => {
    const k = (d.n || '').trim();
    if(k && !allMap.has(k)) allMap.set(k, { type: 'Работа', u: d.u || '', p: d.p || 0 });
  });
  (store.dictMat || []).forEach(d => {
    const k = (d.n || '').trim();
    if(k && !allMap.has(k)) allMap.set(k, { type: 'Материалы', u: d.u || '', p: d.p || 0 });
  });

  const allSorted = Array.from(allMap.keys()).sort((a,b) => a.localeCompare(b, 'ru'));
  const allHtml = allSorted.slice(0, 1500).map(n => {
    const item = allMap.get(n);
    return `<option value="${esc(n)}">${esc(item.type)}${item.u ? ' · ' + esc(item.u) : ''}${item.p ? ' · ' + C.money(item.p) + ' ₽' : ''}</option>`;
  }).join('');
  const dlAll = $('#dict-list-all');
  if(dlAll) dlAll.innerHTML = allHtml;

  const dlm = $('#dict-list-m'); if(dlm) dlm.innerHTML = (store.dictMat || []).slice(0, 500).map(d=>`<option value="${esc(d.n)}">`).join('');
  const dlw = $('#dict-list-w'); if(dlw) dlw.innerHTML = (store.dictWork || []).slice(0, 500).map(d=>`<option value="${esc(d.n)}">`).join('');
}
function toast(t){const el=$('#toast');el.textContent=t;el.classList.remove('hidden');setTimeout(()=>el.classList.add('hidden'),2500);}

/* ================= ТЕМЫ (светлая/тёмная/синяя/сепия) ================= */
const THEMES={light:'Светлая',dark:'Тёмная',blue:'Синяя',sepia:'Сепия'};
function applyTheme(t){
  if(!THEMES[t]) t='light';
  document.body.dataset.theme=t;
  document.documentElement.setAttribute('data-theme', t);
  const sel=$('#theme-select'); if(sel) sel.value=t;
  try{localStorage.setItem('mes_theme',t);}catch(e){}
}
const themeSelect=$('#theme-select');
if(themeSelect) themeSelect.onchange=e=>applyTheme(e.target.value);
let savedTheme='light'; try{savedTheme=localStorage.getItem('mes_theme')||'light';}catch(e){}
if(!THEMES[savedTheme]) savedTheme='light';
applyTheme(savedTheme);

loadStore(); renderOrders(); fillDatalists(); renderKontr();
/* отладка/тесты: доступ к внутреннему состоянию */
window.__APP={get cur(){return cur},set cur(v){cur=v},store,rowClipboard:()=>rowClipboard,selRows,applyTheme,renderKontr};
})();
