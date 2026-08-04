/**
 * 芯片齐套管理系统 - 前端
 */
const API = '';

async function api(path, opts = {}) {
  const url = API + path;
  const o = { headers: { 'Content-Type': 'application/json' }, ...opts };
  if (o.body && typeof o.body === 'object') o.body = JSON.stringify(o.body);
  try {
    const r = await fetch(url, o);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.json();
  } catch (err) {
    console.error('API error:', url, err);
    return { ok: false, error: err.message };
  }
}

function $(s) { return document.querySelector(s); }
function $$(s) { return document.querySelectorAll(s); }
function el(t, a = {}, ...c) {
  const e = document.createElement(t);
  for (const [k, v] of Object.entries(a)) {
    if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (k === 'class') e.className = v;
    else if (k === 'html') e.innerHTML = v;
    else e.setAttribute(k, v);
  }
  c.forEach(x => { if (typeof x === 'string') e.appendChild(document.createTextNode(x)); else if (x) e.appendChild(x); });
  return e;
}

function toast(msg, type = 'info') {
  const t = el('div', { class: `toast toast-${type}` }, msg);
  document.body.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; setTimeout(() => t.remove(), 300); }, 2500);
}

function confirm(msg) {
  return new Promise(r => {
    const o = el('div', { class: 'modal-overlay', onclick: () => { o.remove(); r(false); } });
    o.appendChild(el('div', { class: 'modal-box' },
      el('p', {}, msg),
      el('div', { class: 'modal-btns' },
        el('button', { class: 'btn btn-s', onclick: () => { o.remove(); r(false); } }, '取消'),
        el('button', { class: 'btn btn-p', onclick: () => { o.remove(); r(true); } }, '确认')
      )
    ));
    document.body.appendChild(o);
  });
}

function table(cols, rows, opt = {}) {
  const { edit, del, editFn, delFn, h = 'calc(100vh - 260px)' } = opt;
  const c = el('div', { class: 'dt-container', style: { maxHeight: h } });
  const t = el('table', { class: 'dt' });
  const hd = el('thead');
  const hr = el('tr');
  cols.forEach(cl => hr.appendChild(el('th', {}, cl.label || cl)));
  if (edit || del) hr.appendChild(el('th', { style: { width: '80px' } }, '操作'));
  hd.appendChild(hr); t.appendChild(hd);
  const tb = el('tbody');
  if (!rows || !rows.length) {
    tb.appendChild(el('tr', {}, el('td', { colspan: cols.length + (edit || del ? 1 : 0), style: { textAlign: 'center', padding: '30px', color: '#999' } }, '暂无数据')));
  } else {
    rows.forEach(r => {
      const tr = el('tr');
      cols.forEach(cl => {
        const k = typeof cl === 'string' ? cl : cl.key;
        let v = r[k] ?? '';
        if (typeof v === 'number') v = v.toLocaleString();
        if (typeof v === 'object') v = JSON.stringify(v);
        tr.appendChild(el('td', {}, String(v).substring(0, 200)));
      });
      if (edit || del) {
        const td = el('td', { class: 'action-cell' });
        if (edit) td.appendChild(el('button', { class: 'btn btn-o btn-sm', onclick: e => { e.stopPropagation(); editFn(r); } }, '✏️'));
        if (del) td.appendChild(el('button', { class: 'btn btn-d btn-sm', onclick: async e => { e.stopPropagation(); if (await confirm('确定删除？')) delFn(r); } }, '🗑️'));
        tr.appendChild(td);
      }
      tb.appendChild(tr);
    });
  }
  t.appendChild(tb); c.appendChild(t);
  return c;
}

const App = {
  mod: 'dashboard',
  modules: [
    { id: 'dashboard', name: '📊 仪表盘' },
    { id: 'inventory', name: '📦 库存总览' },
    { id: 'pivot', name: '📊 库存透视' },
    { id: 'shipping', name: '🚚 出货明细' },
    { id: 'model', name: '🔗 机型对照' },
    { id: 'mixbin', name: '🔀 混BIN' },
    { id: 'kit', name: '✅ 齐套达成' },
    { id: 'plan', name: '📋 出货计划' },
    { id: 'erp', name: '🏭 ERP库存' },
    { id: 'mapping', name: '🗺️ 映射管理' },
    { id: 'email', name: '📧 邮件配置' },
    { id: 'upload', name: '📤 数据导入' },
    { id: 'settings', name: '⚙️ 设置' },
  ],
};

async function nav(id) {
  App.mod = id;
  $$('.nav-item').forEach(e => e.classList.remove('active'));
  $(`.nav-item[data-mod="${id}"]`)?.classList.add('active');
  const m = $('#main');
  m.innerHTML = '<div class="loading">加载中...</div>';
  try {
    const fns = { dashboard, inventory, pivot, shipping, model, mixbin, kit, plan, erp, mapping, email, upload, settings };
    await (fns[id] || dashboard)();
  } catch (e) {
    m.innerHTML = `<div class="error">加载失败: ${e.message}</div>`;
  }
}

// ============ 仪表盘 ============
async function dashboard() {
  const m = $('#main');
  const r = await api('/api/dashboard');
  if (!r.ok) { m.innerHTML = `<div class="error">连接服务器失败: ${r.error || '请确认后端已启动 (python server.py)'}</div>`; return; }

  const d = r.data;
  m.innerHTML = '';
  m.appendChild(el('h2', {}, '📊 仪表盘'));

  const cards = el('div', { class: 'cards' });
  [
    ['出货明细', d.total_shipping, '#3b82f6'],
    ['库存记录', d.total_inventory, '#10b981'],
    ['机型对照', d.total_models, '#f59e0b'],
    ['齐套达成', d.total_kit, '#8b5cf6'],
  ].forEach(([l, v, c]) => cards.appendChild(el('div', { class: 'card' },
    el('div', { class: 'card-val', style: { color: c } }, String(v)),
    el('div', { class: 'card-lbl' }, l)
  )));
  m.appendChild(cards);

  m.appendChild(el('h3', {}, '库存分布'));
  m.appendChild(table(
    [{ key: 'warehouse_type', label: '仓库类型' }, { key: 'total_qty', label: '总数量' }],
    (d.inventory_by_type || []).map(r => ({ ...r, total_qty: Number(r.total_qty).toLocaleString() }))
  ));

  const st = el('div', { class: 'stat-row' });
  st.appendChild(el('div', { class: 'stat-item', html: `⏰ 定时采集: <b>每天 9:00 & 21:00</b>` }));
  st.appendChild(el('div', { class: 'stat-item', html: `🔄 <button class="btn btn-o btn-xs" onclick="fetchAllEmails()">手动采集全部</button>` }));
  st.appendChild(el('div', { class: 'stat-item', html: `📥 <a href="/api/export/excel" class="btn btn-p btn-xs" style="text-decoration:none;color:#fff">导出Excel</a>` }));
  m.appendChild(st);
}

async function fetchAllEmails() {
  toast('开始采集...', 'info');
  const r = await api('/api/email/fetch_all');
  if (r.ok) toast('采集完成', 'success');
  else toast('采集失败: ' + r.error, 'error');
}

// ============ 库存总览 ============
async function inventory() {
  const m = $('#main');
  m.innerHTML = '';
  m.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '📦 库存总览'),
    el('div', { class: 'toolbar-act' },
      el('select', { id: 'inv-wh', onchange: loadInv },
        el('option', { value: '' }, '全部'),
        el('option', { value: 'osat' }, 'OSAT'),
        el('option', { value: 'bonded' }, '保税仓'),
        el('option', { value: 'other' }, '其他仓'),
        el('option', { value: 'hold' }, 'Hold'),
        el('option', { value: 'ems' }, 'EMS'),
      ),
      el('input', { id: 'inv-search', placeholder: '搜索芯片型号...', oninput: loadInv }),
    )
  ));
  m.appendChild(el('div', { id: 'inv-tbl' }));
  await loadInv();
}

async function loadInv() {
  const s = ($('#inv-search')?.value || '').toLowerCase();
  const wh = $('#inv-wh')?.value || '';
  
  // 调用关联机型的API
  let url = '/api/inventory/with_model';
  if (s || wh) {
    let params = [];
    if (s) params.push(`device=${encodeURIComponent(s)}`);
    if (wh) params.push(`warehouse_type=${encodeURIComponent(wh)}`);
    url += '?' + params.join('&');
  }
  
  const r = await api(url);
  const c = $('#inv-tbl');
  if (!c) return;
  const cols = [
    { key: 'device', label: '芯片' }, { key: 'bin', label: 'BIN' }, { key: 'test_program', label: '程序' },
    { key: 'total_qty', label: '数量' }, { key: 'warehouse_type', label: '仓库类型' }, { key: 'warehouse_name', label: '仓库名称' },
    { key: 'model1', label: '机型1' }, { key: 'model2', label: '机型2' },
    { key: 'usage_qty', label: '单机用量' }, { key: 'machine_count', label: '可做台数' },
    { key: 'status', label: '状态' },
  ];
  c.innerHTML = '';
  c.appendChild(table(cols, r.data || [], { edit: true, editFn: async row => {
    const v = prompt('数量:', row.total_qty);
    if (v == null) return;
    await api(`/api/update/inventory/${row.id}`, { method: 'PUT', body: { qty: parseInt(v) || 0, version: row.version } });
    loadInv();
  }, del: true, delFn: async row => {
    await api(`/api/delete/inventory/${row.id}`, { method: 'DELETE' });
    loadInv();
  }}));
}

// ============ 库存透视 ============
async function pivot() {
  const m = $('#main');
  m.innerHTML = '';
  m.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '📊 库存透视'),
    el('div', { class: 'toolbar-act' },
      el('select', { id: 'pvt-wh', onchange: loadPivot },
        el('option', { value: '' }, '全部仓库'),
        el('option', { value: 'osat' }, 'OSAT'),
        el('option', { value: 'bonded' }, '保税仓'),
        el('option', { value: 'other' }, '其他仓'),
        el('option', { value: 'hold' }, 'Hold'),
        el('option', { value: 'ems' }, 'EMS'),
      ),
      el('input', { id: 'pvt-model', placeholder: '筛选机型...', oninput: loadPivot }),
    )
  ));
  m.appendChild(el('div', { id: 'pvt-tbl' }));
  await loadPivot();
}

async function loadPivot() {
  const model = $('#pvt-model')?.value || '';
  const wh = $('#pvt-wh')?.value || '';
  const r = await api(`/api/inventory/pivot?model=${encodeURIComponent(model)}&warehouse_type=${encodeURIComponent(wh)}`);
  const c = $('#pvt-tbl');
  if (!c) return;
  c.innerHTML = '';

  // 汇总数据
  const data = r.data || [];
  let totalQty = 0, totalMachines = 0;
  const byType = {};
  data.forEach(d => {
    totalQty += d.total_qty || 0;
    totalMachines += d.machine_count || 0;
    const t = d.warehouse_type || '未知';
    byType[t] = (byType[t] || 0) + (d.total_qty || 0);
  });

  // 汇总卡片
  const cards = el('div', { class: 'cards', style: { marginBottom: '12px' } });
  cards.appendChild(el('div', { class: 'card' }, el('div', { class: 'card-val', style: { color: '#3b82f6' } }, totalQty.toLocaleString()), el('div', { class: 'card-lbl' }, '总芯片数')));
  cards.appendChild(el('div', { class: 'card' }, el('div', { class: 'card-val', style: { color: '#10b981' } }, Math.floor(totalMachines).toLocaleString()), el('div', { class: 'card-lbl' }, '可做台数')));
  cards.appendChild(el('div', { class: 'card' }, el('div', { class: 'card-val', style: { color: '#f59e0b' } }, (byType['osat'] || 0).toLocaleString()), el('div', { class: 'card-lbl' }, 'OSAT库存')));
  cards.appendChild(el('div', { class: 'card' }, el('div', { class: 'card-val', style: { color: '#8b5cf6' } }, (byType['bonded'] || 0).toLocaleString()), el('div', { class: 'card-lbl' }, '保税仓库存')));
  c.appendChild(cards);

  const cols = [
    { key: 'model1', label: '机型' }, { key: 'device', label: '芯片' },
    { key: 'warehouse_type', label: '库存类型' }, { key: 'warehouse_name', label: '仓库名称' },
    { key: 'total_qty', label: '芯片数量' }, { key: 'usage_qty', label: '单机用量' },
    { key: 'machine_count', label: '可做台数' },
  ];
  c.appendChild(table(cols, data));
}

// ============ 出货明细 ============
async function shipping() {
  const m = $('#main');
  m.innerHTML = '';
  m.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '🚚 出货明细'),
    el('div', { class: 'toolbar-act' },
      el('input', { id: 'ship-s', placeholder: '搜索...', oninput: debounce(loadShip, 300) }),
      el('button', { class: 'btn btn-o', onclick: loadExpired }, '⚠️ 过期明细'),
    )
  ));
  m.appendChild(el('div', { id: 'ship-tbl' }));
  await loadShip();
}

async function loadShip() {
  const s = ($('#ship-s')?.value || '').toLowerCase();
  let where = ''; let params = [];
  if (s) { where = 'device_pn LIKE ? OR osat LIKE ? OR invoice_no LIKE ?'; params = [`%${s}%`, `%${s}%`, `%${s}%`]; }
  const r = await api('/api/query/', { method: 'POST', body: { table: 'shipping_detail', where, params, order_by: 'ship_date DESC', limit: 200 } });
  const c = $('#ship-tbl'); if (!c) return;
  c.innerHTML = '';
  c.appendChild(table(
    ['ship_date','osat','device_pn','bin','good_qty','invoice_no','ship_to','po','source'].map(k => ({ key: k, label: k })),
    r.data || []
  ));
}

async function loadExpired() {
  const r = await api('/api/shipping/expired');
  const c = $('#ship-tbl'); if (!c) return;
  c.innerHTML = el('h3', { style: { color: 'red', marginBottom: '10px' } }, '⚠️ 过期出货明细（>180天）');
  c.appendChild(table(['ship_date','osat','device_pn','bin','good_qty','invoice_no'].map(k => ({ key: k, label: k })), r.data || []));
}

function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }

// ============ 机型对照 ============
async function model() {
  const m = $('#main');
  m.innerHTML = '';
  m.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '🔗 机型对照'),
    el('div', { class: 'toolbar-act' },
      el('input', { id: 'mdl-s', placeholder: '搜索...', oninput: debounce(loadModel, 300) }),
      el('button', { class: 'btn btn-o', onclick: loadModelStock }, '📊 关联库存'),
    )
  ));
  m.appendChild(el('div', { id: 'mdl-tbl' }));
  await loadModel();
}

async function loadModel() {
  const s = $('#mdl-s')?.value || '';
  const r = await api(`/api/model/mapping?device=${s}`);
  const c = $('#mdl-tbl'); if (!c) return;
  c.innerHTML = '';
  c.appendChild(table(
    ['device','test_program','bin','model1','model2','exclusive_bin','project'].map(k => ({ key: k, label: k })),
    r.data || [], { edit: true, editFn: async row => {
      const v = prompt('device:', row.device); if (v == null) return;
      await api(`/api/update/model_mapping/${row.id}`, { method: 'PUT', body: { device: v, version: row.version } });
      loadModel();
    }, del: true, delFn: async row => {
      await api(`/api/delete/model_mapping/${row.id}`, { method: 'DELETE' });
      loadModel();
    }}
  ));
}

async function loadModelStock() {
  const r = await api('/api/model/with_stock');
  const c = $('#mdl-tbl'); if (!c) return;
  c.innerHTML = el('h3', { style: { marginBottom: '8px' } }, '📊 机型关联库存');
  c.appendChild(table(
    ['device','bin','model1','stock_qty','usage_qty','machine_count'].map(k => ({ key: k, label: k })),
    r.data || []
  ));
}

// ============ 混BIN ============
async function mixbin() {
  const m = $('#main');
  const r = await api('/api/mixbin/list');
  m.innerHTML = el('h2', {}, '🔀 混BIN分配');
  m.appendChild(table(
    ['device_prog_bin','device','bin','model_name','mix_group','stock_qty','chips_per_unit','convertible_qty','summary_actual'].map(k => ({ key: k, label: k })),
    r.data || [], { edit: true, editFn: async row => {
      const v = prompt('库存:', row.stock_qty); if (v == null) return;
      await api(`/api/update/mix_bin/${row.id}`, { method: 'PUT', body: { stock_qty: parseInt(v) || 0, version: row.version } });
      mixbin();
    }, del: true, delFn: async row => {
      await api(`/api/delete/mix_bin/${row.id}`, { method: 'DELETE' });
      mixbin();
    }}
  ));
}

// ============ 齐套达成 ============
async function kit() {
  const m = $('#main');
  m.innerHTML = '';
  m.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '✅ 齐套达成'),
    el('div', { class: 'toolbar-act' },
      el('select', { id: 'kit-r', onchange: loadKit },
        el('option', { value: '' }, '全部'),
        el('option', { value: '国内' }, '国内'),
        el('option', { value: '海外' }, '海外'),
      ),
      el('button', { class: 'btn btn-o', onclick: calcShortage }, '🔄 计算欠料'),
      el('button', { class: 'btn btn-p', onclick: autoPlan }, '📋 生成出货计划'),
    )
  ));
  m.appendChild(el('div', { id: 'kit-tbl' }));
  await loadKit();
}

async function loadKit() {
  const r = await api('/api/kit/completion');
  const reg = $('#kit-r')?.value || '';
  let data = r.data || [];
  if (reg) data = data.filter(d => d.region === reg);
  const c = $('#kit-tbl'); if (!c) return;
  c.innerHTML = '';
  c.appendChild(table(
    ['region','device','model_name','project','usage_per_unit','subcontractor','sub_code','initial_stock','current_stock','remark'].map(k => ({ key: k, label: k })),
    data, { edit: true, editFn: async row => {
      const v = prompt('库存:', row.current_stock); if (v == null) return;
      await api(`/api/update/kit_completion/${row.id}`, { method: 'PUT', body: { current_stock: parseInt(v) || 0, version: row.version } });
      loadKit();
    }}
  ));
}

async function calcShortage() {
  const r = await api('/api/kit/calculate_shortage', { method: 'POST', body: { region: $('#kit-r')?.value || '' } });
  if (r.ok) { toast('欠料计算完成', 'success'); loadKit(); }
}

async function autoPlan() {
  const r = await api('/api/shipping/auto_plan', { method: 'POST' });
  if (r.ok) toast(`生成 ${r.total} 条出货计划`, 'success');
}

// ============ 出货计划 ============
async function plan() {
  const m = $('#main');
  const r = await api('/api/query/', { method: 'POST', body: { table: 'shipping_plan', order_by: 'warehouse_type, device', limit: 500 } });
  m.innerHTML = el('h2', {}, '📋 出货计划');
  m.appendChild(table(
    ['plan_date','osat','device','bin','qty','warehouse_type','from_warehouse','ship_to','model_name','status'].map(k => ({ key: k, label: k })),
    r.data || [], { edit: true, editFn: async row => {
      const v = prompt('状态 (待确认/已确认/已出货):', row.status); if (!v) return;
      await api(`/api/update/shipping_plan/${row.id}`, { method: 'PUT', body: { status: v, version: row.version } });
      plan();
    }}
  ));
}

// ============ ERP库存 ============
async function erp() {
  const m = $('#main');
  const r = await api('/api/query/', { method: 'POST', body: { table: 'erp_inventory', limit: 500, order_by: 'created_at DESC' } });
  m.innerHTML = el('h2', {}, '🏭 ERP库存');
  m.appendChild(table(
    ['org','material_code','device','bin','qty','sub_inventory','batch'].map(k => ({ key: k, label: k })),
    r.data || []
  ));
}

// ============ 映射管理 ============
async function mapping() {
  const m = $('#main');
  m.innerHTML = '';
  const tabs = ['subcontractor_mapping', 'logistics_time', 'material_device'];
  const names = { subcontractor_mapping: '外协代码', logistics_time: '物流时间', material_device: '料号Device' };
  const tb = el('div', { class: 'tabs' });
  tabs.forEach(t => tb.appendChild(el('button', { class: `tab ${t === tabs[0] ? 'active' : ''}`, onclick: () => loadMapping(t) }, names[t])));
  m.appendChild(tb);
  m.appendChild(el('div', { id: 'map-tbl' }));
  await loadMapping(tabs[0]);
}

async function loadMapping(table) {
  $$('.tab').forEach(t => t.classList.remove('active'));
  event?.target?.classList?.add('active');
  const r = await api(`/api/mapping/${table}`);
  const c = $('#map-tbl'); if (!c) return;

  let cols = [];
  if (table === 'subcontractor_mapping') cols = ['type','short_name','internal_code','external_name','ship_to_code','contact'];
  else if (table === 'logistics_time') cols = ['destination','transit_days','latest_ship_day'];
  else cols = ['erp_code','device','wafer_pn','description'];
  cols = cols.map(k => ({ key: k, label: k }));

  c.innerHTML = '';
  c.appendChild(el('div', { style: { marginBottom: '10px' } }, el('button', { class: 'btn btn-p', onclick: () => addMapping(table) }, '➕ 新增')));
  c.appendChild(table(cols, r.data || [], {
    edit: true, editFn: async row => {
      const k = Object.keys(row).find(k => !['id','version','created_at','updated_at'].includes(k));
      const v = prompt('修改:', row[k]); if (v == null) return;
      row[k] = v;
      await api(`/api/mapping/${table}/${row.id}`, { method: 'PUT', body: { ...row } });
      loadMapping(table);
    },
    del: true, delFn: async row => {
      await api(`/api/mapping/${table}/${row.id}`, { method: 'DELETE' });
      loadMapping(table);
    }
  }));
}

async function addMapping(table) {
  const d = {};
  if (table === 'subcontractor_mapping') {
    d.type = prompt('类型(EMS/原材料仓/保税仓/HUB仓):') || '';
    d.short_name = prompt('公司简称:') || '';
    d.internal_code = prompt('内部代码:') || '';
    d.external_name = prompt('外部名称:') || '';
  } else if (table === 'logistics_time') {
    d.destination = prompt('目的地:') || '';
    d.transit_days = parseInt(prompt('物流天数:') || '0');
    d.latest_ship_day = prompt('最晚发料日:') || '';
  } else {
    d.erp_code = prompt('ERP Code:') || '';
    d.device = prompt('Device:') || '';
  }
  if (!Object.values(d).some(v => v)) return;
  await api(`/api/mapping/${table}`, { method: 'POST', body: d });
  toast('添加成功', 'success');
  loadMapping(table);
}

// ============ 邮件配置 ============
async function email() {
  const m = $('#main');
  m.innerHTML = '';
  m.appendChild(el('div', { class: 'toolbar' },
    el('h2', {}, '📧 邮件配置'),
    el('button', { class: 'btn btn-p', onclick: showEmailConfig }, '➕ 新增'),
  ));

  const r = await api('/api/email_configs');
  if (!r.ok) return;

  r.data.forEach(row => {
    const card = el('div', { class: 'card', style: { marginBottom: '10px' } },
      el('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
        el('div', {},
          el('strong', {}, row.purpose), ' - ', row.description || '',
          el('br'), el('small', { style: { color: '#999' } }, `${row.account} | ${row.match_key || '无匹配'} | 最后采集: ${row.last_fetch || '无'}`)
        ),
        el('div', { style: { display: 'flex', gap: '6px' } },
          el('button', { class: 'btn btn-o btn-sm', onclick: async () => {
            toast('开始采集...', 'info');
            const r = await api(`/api/email_configs/${row.id}/fetch`, { method: 'POST' });
            if (r.ok) toast(`采集: ${r.count} 条`, 'success');
            else toast('失败: ' + r.error, 'error');
          } }, '🔄 采集'),
          el('button', { class: 'btn btn-o btn-sm', onclick: () => editEmailConfig(row) }, '✏️ 编辑'),
          el('button', { class: 'btn btn-d btn-sm', onclick: async () => {
            if (await confirm('删除配置？')) { await api(`/api/email_configs/${row.id}`, { method: 'DELETE' }); email(); }
          } }, '🗑️'),
        )
      )
    );
    m.appendChild(card);
  });
}

async function showEmailConfig() {
  const o = el('div', { class: 'modal-overlay', onclick: e => { if (e.target === o) o.remove(); } });
  o.appendChild(el('div', { class: 'modal-box modal-lg' },
    el('h3', {}, '📧 新增邮件配置'),
    el('div', { class: 'form-g' }, el('label', {}, '用途'), el('select', { id: 'ec-p' },
      el('option', { value: 'shipping_detail' }, '出货明细'),
      el('option', { value: 'osat_inventory' }, 'OSAT库存'),
      el('option', { value: 'hold_inventory' }, 'Hold库存'),
      el('option', { value: 'model_mapping' }, '机型对照表'),
      el('option', { value: 'mix_bin' }, '混BIN关系'),
      el('option', { value: 'order_allocation' }, '订单分配'),
    )),
    el('div', { class: 'form-g' }, el('label', {}, '描述'), el('input', { id: 'ec-desc' })),
    el('div', { class: 'form-g' }, el('label', {}, 'IMAP服务器'), el('input', { id: 'ec-imap', value: 'imap.appia.vip' })),
    el('div', { class: 'form-g' }, el('label', {}, '邮箱'), el('input', { id: 'ec-acc' })),
    el('div', { class: 'form-g' }, el('label', {}, '密码'), el('input', { id: 'ec-pw', type: 'password' })),
    el('div', { class: 'form-g' }, el('label', {}, '文件夹'), el('input', { id: 'ec-folder', value: 'INBOX' })),
    el('div', { class: 'form-g' }, el('label', {}, '匹配关键词'), el('input', { id: 'ec-key' })),
    el('div', { class: 'form-g' }, el('label', {}, '文件后缀'), el('input', { id: 'ec-suffix', value: '.xlsx' })),
    el('div', { class: 'modal-btns' },
      el('button', { class: 'btn btn-s', onclick: () => o.remove() }, '取消'),
      el('button', { class: 'btn btn-p', onclick: async () => {
        const d = {
          purpose: $('#ec-p').value, description: $('#ec-desc').value,
          email_address: $('#ec-acc').value, imap_server: $('#ec-imap').value,
          account: $('#ec-acc').value, password_encrypted: $('#ec-pw').value,
          root_folder: $('#ec-folder').value,
          match_key: $('#ec-key').value, suffix: $('#ec-suffix').value,
          mapping_config: '{}',
        };
        const r = await api('/api/email_configs', { method: 'POST', body: d });
        if (r.ok) { toast('配置已添加', 'success'); o.remove(); email(); }
        else toast('失败: ' + r.error, 'error');
      } }, '保存')
    )
  ));
  document.body.appendChild(o);
}

async function editEmailConfig(row) {
  const o = el('div', { class: 'modal-overlay', onclick: e => { if (e.target === o) o.remove(); } });
  o.appendChild(el('div', { class: 'modal-box modal-lg' },
    el('h3', {}, '✏️ 编辑邮件配置'),
    el('div', { class: 'form-g' }, el('label', {}, '用途'), el('input', { id: 'ec-p', value: row.purpose, disabled: true })),
    el('div', { class: 'form-g' }, el('label', {}, '描述'), el('input', { id: 'ec-desc', value: row.description || '' })),
    el('div', { class: 'form-g' }, el('label', {}, 'IMAP服务器'), el('input', { id: 'ec-imap', value: row.imap_server || '' })),
    el('div', { class: 'form-g' }, el('label', {}, '邮箱'), el('input', { id: 'ec-acc', value: row.account || '' })),
    el('div', { class: 'form-g' }, el('label', {}, '密码（留空不修改）'), el('input', { id: 'ec-pw', type: 'password' })),
    el('div', { class: 'form-g' }, el('label', {}, '文件夹'), el('input', { id: 'ec-folder', value: row.root_folder || 'INBOX' })),
    el('div', { class: 'form-g' }, el('label', {}, '匹配关键词'), el('input', { id: 'ec-key', value: row.match_key || '' })),
    el('div', { class: 'form-g' }, el('label', {}, '文件后缀'), el('input', { id: 'ec-suffix', value: row.suffix || '.xlsx' })),
    el('div', { class: 'modal-btns' },
      el('button', { class: 'btn btn-s', onclick: () => o.remove() }, '取消'),
      el('button', { class: 'btn btn-p', onclick: async () => {
        const d = {
          description: $('#ec-desc').value,
          imap_server: $('#ec-imap').value,
          account: $('#ec-acc').value,
          email_address: $('#ec-acc').value,
          root_folder: $('#ec-folder').value,
          match_key: $('#ec-key').value,
          suffix: $('#ec-suffix').value,
          version: row.version,
        };
        const pw = $('#ec-pw').value;
        if (pw) d.password_encrypted = pw;
        const r = await api(`/api/email_configs/${row.id}`, { method: 'PUT', body: d });
        if (r.ok) { toast('更新成功', 'success'); o.remove(); email(); }
        else toast('更新失败: ' + r.error, 'error');
      } }, '保存')
    )
  ));
  document.body.appendChild(o);
}

// ============ 数据导入 ============
async function upload() {
  const m = $('#main');
  m.innerHTML = '';

  m.appendChild(el('h2', {}, '📤 数据导入'));
  m.appendChild(el('p', { style: { color: '#6b7280', marginBottom: '16px' } }, '导入Excel格式文件（MES/ERP/EMS格式），自动识别并覆盖对应仓库数据'));

  // 库存导入
  const card1 = el('div', { class: 'card', style: { marginBottom: '14px' } },
    el('h3', { style: { marginBottom: '8px' } }, '📦 库存导入'),
    el('div', { class: 'form-g' }, el('label', {}, '格式类型'), el('select', { id: 'up-fmt' },
      el('option', { value: 'mes' }, 'MES格式（SZKXYCL/HSJXYCL等）'),
      el('option', { value: 'erp' }, 'ERP格式（QHBS等）'),
      el('option', { value: 'ems' }, 'EMS格式（芯片结存统计）'),
    )),
    el('div', { class: 'form-g' }, el('label', {}, '仓库名称'), el('input', { id: 'up-wh', placeholder: '如: SZKXYCL, QHBS, DPTMOLTYCL...' })),
    el('div', { class: 'form-g' }, el('label', {}, '仓库类型'), el('select', { id: 'up-wht' },
      el('option', { value: 'other' }, '其他仓'),
      el('option', { value: 'bonded' }, '保税仓'),
      el('option', { value: 'ems' }, 'EMS外协'),
    )),
    el('div', { class: 'form-g' }, el('label', {}, '文件'), el('input', { type: 'file', id: 'up-file', accept: '.xlsx,.xls' })),
    el('button', { class: 'btn btn-p', onclick: async () => {
      const f = $('#up-file')?.files?.[0]; if (!f) { toast('请选择文件', 'error'); return; }
      const fd = new FormData();
      fd.append('file', f);
      fd.append('format_type', $('#up-fmt').value);
      fd.append('warehouse_name', $('#up-wh').value);
      fd.append('warehouse_type', $('#up-wht').value);
      const r = await fetch(API + '/api/upload/inventory', { method: 'POST', body: fd });
      const d = await r.json();
      if (d.ok) toast(`导入成功: ${d.count} 条`, 'success');
      else toast('导入失败: ' + d.error, 'error');
    } }, '上传')
  );
  m.appendChild(card1);

  // 出货明细导入
  const card2 = el('div', { class: 'card', style: { marginBottom: '14px' } },
    el('h3', { style: { marginBottom: '8px' } }, '🚚 出货明细导入'),
    el('div', { class: 'form-g' }, el('label', {}, '文件'), el('input', { type: 'file', id: 'up-ship', accept: '.xlsx,.xls' })),
    el('button', { class: 'btn btn-p', onclick: async () => {
      const f = $('#up-ship')?.files?.[0]; if (!f) { toast('请选择文件', 'error'); return; }
      const fd = new FormData(); fd.append('file', f);
      const r = await fetch(API + '/api/upload/shipping', { method: 'POST', body: fd });
      const d = await r.json();
      if (d.ok) toast(`导入成功: ${d.count} 条`, 'success');
      else toast('导入失败: ' + d.error, 'error');
    } }, '上传')
  );
  m.appendChild(card2);

  // ERP库存导入
  const card3 = el('div', { class: 'card' },
    el('h3', { style: { marginBottom: '8px' } }, '🏭 ERP库存导入'),
    el('div', { class: 'form-g' }, el('label', {}, '仓库名称'), el('input', { id: 'up-erp-wh', placeholder: '如: QHBS' })),
    el('div', { class: 'form-g' }, el('label', {}, '文件'), el('input', { type: 'file', id: 'up-erp', accept: '.xlsx,.xls' })),
    el('button', { class: 'btn btn-p', onclick: async () => {
      const f = $('#up-erp')?.files?.[0]; if (!f) { toast('请选择文件', 'error'); return; }
      const fd = new FormData(); fd.append('file', f); fd.append('warehouse_name', $('#up-erp-wh').value);
      const r = await fetch(API + '/api/upload/erp_inventory', { method: 'POST', body: fd });
      const d = await r.json();
      if (d.ok) toast(`导入成功: ${d.count} 条`, 'success');
      else toast('导入失败: ' + d.error, 'error');
    } }, '上传')
  );
  m.appendChild(card3);
}

// ============ 设置 ============
async function settings() {
  const m = $('#main');
  m.innerHTML = '';

  m.appendChild(el('h2', {}, '⚙️ 系统设置'));

  m.appendChild(el('h3', {}, '👥 用户'));
  const ur = await api('/api/users');
  m.appendChild(table(['email','name','role','active'].map(k => ({ key: k, label: k })), ur.data || []));
  m.appendChild(el('button', { class: 'btn btn-p', style: { marginTop: '10px' }, onclick: async () => {
    const email = prompt('邮箱:'); if (!email) return;
    await api('/api/users', { method: 'POST', body: { email, name: prompt('姓名:') || '', role: prompt('角色(admin/editor/viewer):','viewer') || 'viewer' } });
    settings();
  } }, '➕ 添加用户'));

  m.appendChild(el('h3', { style: { marginTop: '20px' } }, '📝 日志'));
  const lr = await api('/api/logs?limit=50');
  m.appendChild(table(['action','table_name','record_id','detail','operator','created_at'].map(k => ({ key: k, label: k })), lr.data || []));
}

// ============ 初始化 ============
function init() {
  const navEl = $('#nav-menu');
  App.modules.forEach(mod => {
    navEl.appendChild(el('div', { class: `nav-item ${mod.id === 'dashboard' ? 'active' : ''}`, 'data-mod': mod.id, onclick: () => nav(mod.id) }, `${mod.name}`));
  });
  nav('dashboard');
}

document.addEventListener('DOMContentLoaded', () => {
  // 检查登录状态
  const user = localStorage.getItem('chipkit_user');
  if (!user) {
    location.href = '/login.html';
    return;
  }
  // 显示用户信息
  const u = JSON.parse(user);
  const logo = document.getElementById('sb-logo');
  if (logo) logo.innerHTML = `🦞 齐套管理<br><small style="font-size:11px;color:#888">${u.name || u.email}</small>`;
  init();
});