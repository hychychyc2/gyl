/**
 * 芯片齐套管理系统 - 纯前端版 (IndexedDB)
 * 无需后端，通过 OpenClaw canvas 直接运行
 */
// API模式检测：自动检测后端是否可达
let API_MODE = 'indexeddb';
let API_BASE = '';

async function detectApiMode() {
  // 尝试同源下的 /api/ 路径
  const candidates = [
    '',           // 同源
    'http://localhost:8765',  // 本地部署
    'http://127.0.0.1:8765',
  ];
  for (const base of candidates) {
    try {
      const resp = await fetch(base + '/api/dashboard');
      const data = await resp.json();
      if (data.ok) {
        API_MODE = 'api';
        API_BASE = base;
        console.log('✅ API模式：已连接 ' + (base || '同源') + ' 后端服务器');
        return true;
      }
    } catch(e) {}
  }
  console.log('📦 IndexedDB模式：使用本地存储');
  return false;
}

// ============ IndexedDB 数据库层 ============
const DB = {
  _db: null,
  _ready: null,

  async init() {
    if (this._ready) return this._ready;
    this._ready = new Promise((resolve, reject) => {
      const req = indexedDB.open('ChipKitDB', 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        const tables = {
          inventory: { keyPath: 'id', autoIncrement: true, indexes: ['device','warehouse_type','device_prog_bin'] },
          shipping_detail: { keyPath: 'id', autoIncrement: true, indexes: ['device_pn','osat','ship_date'] },
          model_mapping: { keyPath: 'id', autoIncrement: true, indexes: ['device','device_prog_bin'] },
          usage_mapping: { keyPath: 'id', autoIncrement: true, indexes: ['device'] },
          mix_bin: { keyPath: 'id', autoIncrement: true },
          subcontractor_mapping: { keyPath: 'id', autoIncrement: true },
          logistics_time: { keyPath: 'id', autoIncrement: true },
          material_device: { keyPath: 'id', autoIncrement: true },
          kit_completion: { keyPath: 'id', autoIncrement: true },
          shipping_plan: { keyPath: 'id', autoIncrement: true },
          erp_inventory: { keyPath: 'id', autoIncrement: true },
          email_config: { keyPath: 'id', autoIncrement: true },
          users: { keyPath: 'id', autoIncrement: true },
        };
        for (const [name, opts] of Object.entries(tables)) {
          if (!db.objectStoreNames.contains(name)) {
            const store = db.createObjectStore(name, opts);
            (opts.indexes||[]).forEach(idx => store.createIndex(idx, idx));
          }
        }
      };
      req.onsuccess = (e) => {
        this._db = e.target.result;
        console.log('✅ IndexedDB 初始化完成');
        resolve();
      };
      req.onerror = () => reject(req.error);
    });
    return this._ready;
  },

  async getAll(storeName, indexName, value) {
    await this.init();
    return new Promise((resolve) => {
      const tx = this._db.transaction(storeName, 'readonly');
      const store = tx.objectStore(storeName);
      const req = indexName && value ? store.index(indexName).getAll(value) : store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => resolve([]);
    });
  },

  async getSome(storeName, limit = 500) {
    await this.init();
    return new Promise((resolve) => {
      const tx = this._db.transaction(storeName, 'readonly');
      const store = tx.objectStore(storeName);
      const req = store.getAll(null, limit);
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => resolve([]);
    });
  },

  async put(storeName, data) {
    await this.init();
    return new Promise((resolve) => {
      const tx = this._db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      const req = data.id ? store.put(data) : store.add(data);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    });
  },

  async putMany(storeName, items) {
    await this.init();
    return new Promise((resolve) => {
      const tx = this._db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      let count = 0;
      items.forEach(item => {
        const req = item.id ? store.put(item) : store.add(item);
        req.onsuccess = () => { count++; if (count === items.length) resolve(count); };
        req.onerror = () => { count++; if (count === items.length) resolve(count); };
      });
      if (items.length === 0) resolve(0);
    });
  },

  async delete(storeName, id) {
    await this.init();
    return new Promise((resolve) => {
      const tx = this._db.transaction(storeName, 'readwrite');
      store.delete(id);
      tx.oncomplete = () => resolve(true);
    });
  },

  async count(storeName) {
    await this.init();
    return new Promise((resolve) => {
      const tx = this._db.transaction(storeName, 'readonly');
      const req = tx.objectStore(storeName).count();
      req.onsuccess = () => resolve(req.result);
    });
  },

  async clear(storeName) {
    await this.init();
    return new Promise((resolve) => {
      const tx = this._db.transaction(storeName, 'readwrite');
      tx.objectStore(storeName).clear();
      tx.oncomplete = () => resolve();
    });
  }
};

// ============ 全局状态 ============
const App = {
  currentModule: 'dashboard',
  modules: [
    { id: 'dashboard', name: '📊 仪表盘', icon: '📊' },
    { id: 'inventory', name: '📦 库存总览', icon: '📦' },
    { id: 'shipping', name: '🚚 出货明细', icon: '🚚' },
    { id: 'model', name: '🔗 机型对照', icon: '🔗' },
    { id: 'mixbin', name: '🔀 混BIN分配', icon: '🔀' },
    { id: 'kit', name: '✅ 齐套达成', icon: '✅' },
    { id: 'plan', name: '📋 出货计划', icon: '📋' },
    { id: 'erp', name: '🏭 ERP库存', icon: '🏭' },
    { id: 'mapping', name: '🗺️ 映射管理', icon: '🗺️' },
    { id: 'import', name: '📥 数据导入', icon: '📥' },
  ],
};

// ============ UI 工具 ============
function $(sel) { return document.querySelector(sel); }
function $$(sel) { return document.querySelectorAll(sel); }

function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k === 'class') e.className = v;
    else if (k === 'html') e.innerHTML = v;
    else e.setAttribute(k, v);
  }
  for (const c of children) {
    if (typeof c === 'string') e.appendChild(document.createTextNode(c));
    else if (c) e.appendChild(c);
  }
  return e;
}

function showToast(msg, type = 'info') {
  const t = el('div', { class: `toast toast-${type}` }, msg);
  document.body.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; setTimeout(() => t.remove(), 300); }, 2500);
}

async function confirmDialog(msg) {
  return new Promise(resolve => {
    const overlay = el('div', { class: 'modal-overlay', onclick: () => { overlay.remove(); resolve(false); } });
    const box = el('div', { class: 'modal-box' },
      el('p', {}, msg),
      el('div', { class: 'modal-btns' },
        el('button', { class: 'btn btn-secondary', onclick: () => { overlay.remove(); resolve(false); } }, '取消'),
        el('button', { class: 'btn btn-primary', onclick: () => { overlay.remove(); resolve(true); } }, '确认')
      )
    );
    overlay.appendChild(box);
    document.body.appendChild(overlay);
  });
}

function createTable(columns, rows, options = {}) {
  const { editable, onEdit, onDelete, height = 'calc(100vh - 280px)' } = options;
  const container = el('div', { class: 'data-table-container', style: { maxHeight: height } });
  const table = el('table', { class: 'data-table' });
  const thead = el('thead');
  const headerRow = el('tr');
  columns.forEach(col => headerRow.appendChild(el('th', {}, col.label || col)));
  if (editable || onDelete) headerRow.appendChild(el('th', { style: { width: '100px' } }, '操作'));
  thead.appendChild(headerRow);
  table.appendChild(thead);
  const tbody = el('tbody');
  if (!rows || rows.length === 0) {
    tbody.appendChild(el('tr', {}, el('td', { colspan: columns.length + (editable||onDelete?1:0), style: { textAlign: 'center', padding: '40px', color: '#999' } }, '暂无数据')));
  } else {
    rows.forEach(row => {
      const tr = el('tr');
      columns.forEach(col => {
        const key = typeof col === 'string' ? col : col.key;
        let val = row[key] ?? '';
        if (typeof val === 'number') val = val.toLocaleString();
        if (typeof val === 'object') val = JSON.stringify(val);
        tr.appendChild(el('td', {}, String(val).substring(0, 200)));
      });
      if (editable || onDelete) {
        const td = el('td', { class: 'action-cell' });
        if (editable) td.appendChild(el('button', { class: 'btn btn-sm btn-outline', onclick: (e) => { e.stopPropagation(); onEdit && onEdit(row); } }, '✏️'));
        if (onDelete) td.appendChild(el('button', { class: 'btn btn-sm btn-danger', onclick: async (e) => { e.stopPropagation(); if (await confirmDialog('确定删除？')) onDelete(row); } }, '🗑️'));
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
  }
  table.appendChild(tbody);
  container.appendChild(table);
  return container;
}

// ============ 路由 ============
async function navigate(moduleId) {
  App.currentModule = moduleId;
  $$('.nav-item').forEach(el => el.classList.remove('active'));
  $(`.nav-item[data-module="${moduleId}"]`)?.classList.add('active');
  const main = $('#main-content');
  main.innerHTML = '<div class="loading">加载中...</div>';
  try {
    switch (moduleId) {
      case 'dashboard': await loadDashboard(); break;
      case 'inventory': await loadInventory(); break;
      case 'shipping': await loadShipping(); break;
      case 'model': await loadModel(); break;
      case 'mixbin': await loadMixBin(); break;
      case 'kit': await loadKit(); break;
      case 'plan': await loadPlan(); break;
      case 'erp': await loadERP(); break;
      case 'mapping': await loadMapping(); break;
      case 'import': await loadImport(); break;
    }
  } catch (err) {
    main.innerHTML = `<div class="error">加载失败: ${err.message}</div>`;
  }
}

// ============ 仪表盘 ============
async function loadDashboard() {
  const main = $('#main-content');
  main.innerHTML = '';
  main.appendChild(el('h2', {}, '📊 仪表盘'));

  let shipCount, invCount, modelCount;
  if (API_MODE === 'api') {
    const resp = await (await fetch(API_BASE + '/api/dashboard')).json();
    if (resp.ok) {
      shipCount = resp.data.total_shipping;
      invCount = resp.data.total_inventory;
      modelCount = resp.data.total_models;
    }
  } else {
    [shipCount, invCount, modelCount] = await Promise.all([
      DB.count('shipping_detail'), DB.count('inventory'), DB.count('model_mapping')
    ]);
  }

  const cards = el('div', { class: 'cards-grid' });
  [
    { label: '出货明细', value: shipCount, color: '#3b82f6' },
    { label: '库存记录', value: invCount, color: '#10b981' },
    { label: '机型对照', value: modelCount, color: '#f59e0b' },
  ].forEach(item => {
    cards.appendChild(el('div', { class: 'card' },
      el('div', { class: 'card-value', style: { color: item.color } }, String(item.value)),
      el('div', { class: 'card-label' }, item.label)
    ));
  });
  main.appendChild(cards);

  const modeText = API_MODE === 'api' ? '🟢 API模式（连接后端服务器）' : '🟡 离线模式（IndexedDB本地存储）';
  main.appendChild(el('p', { style: { color: '#6b7280', fontSize: '14px' } }, modeText + ' | 💡 离线模式下请到 📥 数据导入 页面加载JSON数据'));
}

// ============ 库存总览 ============
async function loadInventory() {
  const main = $('#main-content');
  main.innerHTML = '';
  main.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '📦 库存总览'),
    el('div', { class: 'toolbar-actions' },
      el('input', { id: 'inv-search', placeholder: '搜索芯片型号...', oninput: debounce(loadInventoryData, 300) }),
    )
  ));
  const container = el('div', { id: 'inv-table-container' });
  main.appendChild(container);
  await loadInventoryData();
}

async function loadInventoryData() {
  const search = ($('#inv-search')?.value || '').toLowerCase();
  const container = $('#inv-table-container');
  if (!container) return;

  let data = await DB.getSome('inventory', 500);
  if (search) data = data.filter(r => (r.device||'').toLowerCase().includes(search));

  const cols = [
    { key: 'device', label: '芯片型号' }, { key: 'bin', label: 'BIN' },
    { key: 'test_program', label: '测试程序' }, { key: 'qty', label: '数量' },
    { key: 'warehouse_type', label: '仓库类型' }, { key: 'warehouse_name', label: '仓库名称' },
    { key: 'status', label: '状态' },
  ];
  container.innerHTML = '';
  container.appendChild(createTable(cols, data, { editable: true, onEdit: editInventory, onDelete: deleteInventory }));
}

async function editInventory(row) {
  const qty = prompt('修改数量:', row.qty);
  if (qty === null) return;
  row.qty = parseInt(qty) || 0;
  await DB.put('inventory', row);
  showToast('更新成功', 'success');
  loadInventoryData();
}

async function deleteInventory(row) {
  await DB.delete('inventory', row.id);
  showToast('删除成功');
  loadInventoryData();
}

function debounce(fn, ms) {
  let timer;
  return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), ms); };
}

// ============ 出货明细 ============
async function loadShipping() {
  const main = $('#main-content');
  main.innerHTML = '';
  main.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '🚚 出货明细'),
    el('input', { id: 'ship-search', placeholder: '搜索...', oninput: debounce(loadShippingData, 300) }),
  ));
  const container = el('div', { id: 'ship-table-container' });
  main.appendChild(container);
  await loadShippingData();
}

async function loadShippingData() {
  const search = ($('#ship-search')?.value || '').toLowerCase();
  const container = $('#ship-table-container');
  if (!container) return;
  let data = await DB.getSome('shipping_detail', 200);
  if (search) data = data.filter(r => (r.device_pn||'').toLowerCase().includes(search) || (r.osat||'').toLowerCase().includes(search));
  const cols = ['ship_date','osat','device_pn','bin','good_qty','invoice_no','ship_to','po'].map(k => ({ key: k, label: k }));
  container.innerHTML = '';
  container.appendChild(createTable(cols, data));
}

// ============ 机型对照 ============
async function loadModel() {
  const main = $('#main-content');
  main.innerHTML = '';
  main.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '🔗 机型对照表'),
    el('input', { id: 'model-search', placeholder: '搜索...', oninput: debounce(loadModelData, 300) }),
  ));
  const container = el('div', { id: 'model-table-container' });
  main.appendChild(container);
  await loadModelData();
}

async function loadModelData() {
  const search = ($('#model-search')?.value || '').toLowerCase();
  const container = $('#model-table-container');
  if (!container) return;
  let data = await DB.getSome('model_mapping', 200);
  if (search) data = data.filter(r => (r.device||'').toLowerCase().includes(search));
  const cols = ['device','test_program','bin','model1','model2','model3','exclusive_bin','project'].map(k => ({ key: k, label: k }));
  container.innerHTML = '';
  container.appendChild(createTable(cols, data, { editable: true, onEdit: async (row) => {
    const v = prompt('修改:', row.device);
    if (v === null) return;
    row.device = v;
    await DB.put('model_mapping', row);
    loadModelData();
  }, onDelete: async (row) => {
    await DB.delete('model_mapping', row.id);
    loadModelData();
  }}));
}

// ============ 混BIN ============
async function loadMixBin() {
  const main = $('#main-content');
  main.innerHTML = el('h2', {}, '🔀 混BIN分配');
  const data = await DB.getSome('mix_bin', 100);
  const cols = ['device_prog_bin','device','bin','model_name','mix_group','stock_qty','chips_per_unit','convertible_qty','summary_actual','is_exclusive'].map(k => ({ key: k, label: k }));
  main.appendChild(createTable(cols, data, { editable: true, onEdit: async (row) => {
    const v = prompt('库存:', row.stock_qty);
    if (v === null) return;
    row.stock_qty = parseInt(v) || 0;
    await DB.put('mix_bin', row);
    loadMixBin();
  }, onDelete: async (row) => {
    await DB.delete('mix_bin', row.id);
    loadMixBin();
  }}));
}

// ============ 齐套达成 ============
async function loadKit() {
  const main = $('#main-content');
  main.innerHTML = el('h2', {}, '✅ 齐套达成');
  const data = await DB.getSome('kit_completion', 200);
  const cols = ['region','device','model_name','project','usage_per_unit','subcontractor','sub_code','initial_stock','current_stock','remark'].map(k => ({ key: k, label: k }));
  main.appendChild(createTable(cols, data, { editable: true, onEdit: async (row) => {
    const v = prompt('当前库存:', row.current_stock);
    if (v === null) return;
    row.current_stock = parseInt(v) || 0;
    await DB.put('kit_completion', row);
    loadKit();
  }}));
}

// ============ 出货计划 ============
async function loadPlan() {
  const main = $('#main-content');
  main.innerHTML = el('h2', {}, '📋 出货计划');
  const data = await DB.getSome('shipping_plan', 200);
  const cols = ['plan_date','osat','device','bin','qty','warehouse_type','from_warehouse','ship_to','model_name','status'].map(k => ({ key: k, label: k }));
  main.appendChild(createTable(cols, data, { editable: true, onEdit: async (row) => {
    const v = prompt('状态 (待确认/已确认/已出货):', row.status);
    if (!v) return;
    row.status = v;
    await DB.put('shipping_plan', row);
    loadPlan();
  }}));
}

// ============ ERP库存 ============
async function loadERP() {
  const main = $('#main-content');
  main.innerHTML = el('h2', {}, '🏭 ERP库存');
  const data = await DB.getSome('erp_inventory', 200);
  const cols = ['org','material_code','device','bin','qty','sub_inventory','batch'].map(k => ({ key: k, label: k }));
  main.appendChild(createTable(cols, data));
}

// ============ 映射管理 ============
async function loadMapping() {
  const main = $('#main-content');
  main.innerHTML = '';
  const tabs = ['subcontractor_mapping', 'logistics_time', 'material_device'];
  const tabNames = { subcontractor_mapping: '外协代码', logistics_time: '物流时间', material_device: '料号Device' };
  const tabBar = el('div', { class: 'tabs' });
  tabs.forEach(t => tabBar.appendChild(el('button', { class: `tab ${t === tabs[0] ? 'active' : ''}`, onclick: () => loadMappingTable(t) }, tabNames[t])));
  main.appendChild(tabBar);
  main.appendChild(el('div', { id: 'mapping-table-container' }));
  await loadMappingTable(tabs[0]);
}

async function loadMappingTable(table) {
  const container = $('#mapping-table-container');
  if (!container) return;
  $$('.tab').forEach(t => t.classList.remove('active'));
  event?.target?.classList?.add('active');
  const data = await DB.getSome(table, 200);
  let cols = table === 'subcontractor_mapping' ? ['type','short_name','internal_code','external_name','ship_to_code','contact']
    : table === 'logistics_time' ? ['destination','transit_days','latest_ship_day']
    : ['erp_code','device','wafer_pn','description'];
  cols = cols.map(k => ({ key: k, label: k }));
  container.innerHTML = '';
  container.appendChild(el('div', { class: 'toolbar' }, el('button', { class: 'btn btn-primary', onclick: () => addMappingRow(table) }, '➕ 新增')));
  container.appendChild(createTable(cols, data, {
    editable: true,
    onEdit: async (row) => {
      const k = Object.keys(row).find(k => k !== 'id' && k !== 'version' && k !== 'created_at');
      const v = prompt('修改:', row[k]);
      if (v === null) return;
      row[k] = v;
      await DB.put(table, row);
      loadMappingTable(table);
    },
    onDelete: async (row) => {
      await DB.delete(table, row.id);
      loadMappingTable(table);
    }
  }));
}

async function addMappingRow(table) {
  const data = {};
  if (table === 'subcontractor_mapping') {
    data.type = prompt('类型:') || ''; data.short_name = prompt('简称:') || '';
    data.internal_code = prompt('内部代码:') || ''; data.external_name = prompt('外部名称:') || '';
  } else if (table === 'logistics_time') {
    data.destination = prompt('目的地:') || ''; data.transit_days = parseInt(prompt('物流天数:') || '0');
  } else {
    data.erp_code = prompt('ERP Code:') || ''; data.device = prompt('Device:') || '';
  }
  if (!Object.values(data).some(v => v)) return;
  await DB.put(table, data);
  showToast('添加成功');
  loadMappingTable(table);
}

// ============ 数据导入 ============
async function loadImport() {
  const main = $('#main-content');
  main.innerHTML = '';
  main.appendChild(el('h2', {}, '📥 数据导入'));

  main.appendChild(el('h3', {}, '📤 从 JSON 文件导入'));

  const tables = [
    { name: 'inventory', label: '库存' },
    { name: 'shipping_detail', label: '出货明细' },
    { name: 'model_mapping', label: '机型对照' },
    { name: 'usage_mapping', label: '用量对照' },
    { name: 'mix_bin', label: '混BIN' },
    { name: 'kit_completion', label: '齐套达成' },
    { name: 'subcontractor_mapping', label: '外协代码' },
    { name: 'logistics_time', label: '物流时间' },
    { name: 'material_device', label: '料号Device' },
  ];

  tables.forEach(t => {
    const row = el('div', { class: 'toolbar', style: { marginBottom: '8px' } },
      el('span', {}, t.label),
      el('input', { type: 'file', accept: '.json', id: `import-${t.name}`, style: { display: 'none' },
        onchange: async (e) => {
          const file = e.target.files[0];
          if (!file) return;
          const text = await file.text();
          try {
            const data = JSON.parse(text);
            const items = Array.isArray(data) ? data : [data];
            const count = await DB.putMany(t.name, items);
            showToast(`${t.label}: 导入 ${count} 条`, 'success');
          } catch (err) {
            showToast(`导入失败: ${err.message}`, 'error');
          }
        }
      }),
      el('button', { class: 'btn btn-outline btn-sm', onclick: () => $(`#import-${t.name}`).click() }, '选择文件'),
      el('button', { class: 'btn btn-sm', style: { background: '#ef4444', color: '#fff' }, onclick: async () => {
        if (await confirmDialog(`确定清空 ${t.label} 数据？`)) {
          await DB.clear(t.name);
          showToast(`${t.label} 已清空`);
        }
      } }, '清空'),
    );
    main.appendChild(row);
  });

  main.appendChild(el('hr', { style: { margin: '20px 0' } }));
  main.appendChild(el('p', { style: { color: '#6b7280', fontSize: '14px' } },
    '💡 提示：使用后端API导出JSON文件，然后在此导入。也可以直接使用后端API服务 (http://localhost:8765) 来管理数据。'));
}

// ============ 初始化 ============
async function init() {
  await detectApiMode();
  if (API_MODE === 'indexeddb') {
    await DB.init();
  }

  const nav = $('#nav-menu');
  App.modules.forEach(mod => {
    nav.appendChild(el('div', {
      class: `nav-item ${mod.id === 'dashboard' ? 'active' : ''}`,
      'data-module': mod.id,
      onclick: () => navigate(mod.id)
    }, `${mod.icon} ${mod.name}`));
  });

  await navigate('dashboard');
}

document.addEventListener('DOMContentLoaded', init);