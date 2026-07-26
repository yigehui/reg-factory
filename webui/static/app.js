// reg-factory WebUI 前端逻辑（原生 JS，无构建）
let SCRIPTS = [];
let EMBEDS = [];
let curRun = null;     // 当前运行 run_id
let curSrc = null;     // 当前选中脚本
let evtSrc = null;     // EventSource
let smsTimer = null;   // 接码助手倒计时刷新
let curSavedArgs = {}; // 当前脚本已保存的参数
let selectedAccountIds = new Set();
let taskState = {
  page: 1,
  pageSize: 20,
  total: 0,
  loaded: [],
};
let accountState = {
  page: 1,
  pageSize: 20,
  total: 0,
  filters: {taskId:'', email:'', status:''},
  loaded: [],
};
let taskDetailState = {
  taskId: '',
  page: 1,
  pageSize: 20,
  total: 0,
  loaded: [],
};

const $ = (s, r=document) => r.querySelector(s);
const $$ = (s, r=document) => [...r.querySelectorAll(s)];
const LogAutoScroll = globalThis.LogAutoScroll || {
  appendLogLineWithAutoScroll(log, line){
    log.textContent += line + '\n';
    log.scrollTop = log.scrollHeight;
  }
};

function escHtml(v){
  return String(v ?? '').replace(/[&<>"']/g, ch => ({
    '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'
  }[ch]));
}

function escAttr(v){
  return escHtml(v);
}

function choiceValue(c){
  return (c && typeof c === 'object') ? c.value : c;
}

function choiceLabel(c){
  return (c && typeof c === 'object') ? (c.label ?? c.value ?? '') : c;
}

function renderChoiceOptions(choices, selected){
  return (choices || []).map(c=>{
    const value = choiceValue(c);
    const text = choiceLabel(c);
    return `<option value="${escAttr(value)}" ${String(value)==String(selected)?'selected':''}>${escHtml(text)}</option>`;
  }).join('');
}

// ---------------------------------------------------------------- 状态灯轮询
async function pollStatus(){
  try{
    const s = await (await fetch('/api/status')).json();
    $('#dot-bb').classList.toggle('on', s.bitbrowser);
    const label = s.browser_provider === 'adspower' ? 'AdsPower' : 'BitBrowser';
    $('#browser-label').textContent = label;
    $('#dot-clash').classList.toggle('on', s.clash);
    $('#node').textContent = '节点 ' + (s.node || '--');
    $('#running').textContent = s.running ? `● ${s.running} 个任务运行中` : '';
  }catch(e){}
}
setInterval(pollStatus, 5000);

// ---------------------------------------------------------------- 视图切换
function showView(v){
  const setDisplay = (selector, displayWhenActive, active) => {
    const el = $(selector);
    if(el) el.style.display = active ? displayWhenActive : 'none';
  };
  setDisplay('#view-run', 'flex', v==='run');
  setDisplay('#view-tasks', 'block', v==='tasks');
  setDisplay('#view-accounts', 'block', v==='accounts');
  setDisplay('#view-env', 'block', v==='env');
  setDisplay('#view-embed', 'block', v==='embed');
  setDisplay('#view-mailpool', 'block', v==='mailpool');
  $$('.navbtn').forEach(b=>b.classList.toggle('active', b.dataset.view===v));
  if(v==='env') loadEnv();
  if(v==='mailpool') loadMailpool();
  if(v==='tasks') loadTaskDashboard();
  if(v==='accounts') loadAccounts();
}
$$('.navbtn').forEach(b=> b.onclick = ()=> showView(b.dataset.view));

// ---------------------------------------------------------------- 脚本导航
async function loadScripts(){
  SCRIPTS = (await (await fetch('/api/scripts')).json()).scripts;
  const nav = $('#script-nav');
  const cats = {};
  SCRIPTS.forEach(s => (cats[s.category]=cats[s.category]||[]).push(s));
  nav.innerHTML = '';

  // 内嵌功能页(Gmail 等) —— 放最上面
  try{
    EMBEDS = (await (await fetch('/api/embeds')).json()).embeds || [];
    if(EMBEDS.length){
      const t=document.createElement('div'); t.className='cat-title'; t.textContent='功能'; nav.appendChild(t);
      EMBEDS.forEach(e=>{
        const b=document.createElement('button');
        b.className='scriptbtn'; b.textContent='🌐 '+e.title; b.dataset.embed=e.id;
        b.onclick=()=>openEmbed(e.id);
        nav.appendChild(b);
      });
    }
  }catch(err){}

  for(const cat of Object.keys(cats)){
    const t = document.createElement('div');
    t.className='cat-title'; t.textContent=cat; nav.appendChild(t);
    cats[cat].forEach(s=>{
      const b=document.createElement('button');
      b.className='scriptbtn'; b.textContent=s.title; b.dataset.id=s.id;
      b.onclick=()=>{ showView('run'); selectScript(s.id); };
      nav.appendChild(b);
    });
  }
  // 外部工具链接(新标签打开)
  try{
    const links = (await (await fetch('/api/links')).json()).links || [];
    if(links.length){
      const t = document.createElement('div');
      t.className='cat-title'; t.textContent='外部工具'; nav.appendChild(t);
      links.forEach(l=>{
        const a=document.createElement('a');
        a.className='scriptbtn linkbtn'; a.href=l.url; a.target='_blank'; a.rel='noopener';
        a.title=l.desc||l.url; a.innerHTML=`🔗 ${l.title}`;
        nav.appendChild(a);
      });
    }
  }catch(e){}
}

async function selectScript(id){
  curSrc = SCRIPTS.find(s=>s.id===id);
  $$('.scriptbtn').forEach(b=>b.classList.toggle('active', b.dataset.id===id));
  curSavedArgs = await loadScriptConfig(id);
  renderForm(curSrc, curSavedArgs);
}

// ---------------------------------------------------------------- 渲染表单
async function loadScriptConfig(id){
  try{
    const r = await (await fetch(`/api/script-config/${encodeURIComponent(id)}`)).json();
    return r.args || {};
  }catch(e){
    return {};
  }
}

function argValue(a, saved){
  if(saved && Object.prototype.hasOwnProperty.call(saved, a.flag)) return saved[a.flag];
  return a.default;
}

function renderForm(s, saved={}){
  const p = $('#form-panel');
  p.innerHTML = '';
  p.classList.toggle('two-col-form', ['register_outlook_ruoyi'].includes(s.id));
  const h = document.createElement('div');
  h.className = 'form-head';
  h.innerHTML = `<h2 class="form-title">${s.title}</h2><p class="form-desc">${s.desc||''}</p>`;
  p.appendChild(h);

  s.args.forEach(a=>{
    const f = document.createElement('div'); f.className='field';
    const label = a.flag.replace(/^--/,'');
    const val = argValue(a, saved);
    if(a.type==='bool'){
      f.className='field checkbox';
      f.innerHTML = `<input type="checkbox" id="f_${label}" ${val?'checked':''}>
        <label for="f_${label}">${label}</label>`;
      if(a.help){ const hh=document.createElement('div'); hh.className='fhelp'; hh.textContent=a.help; f.appendChild(hh); }
    }else if(a.type==='choice'){
      const options = renderChoiceOptions(a.choices, val);
      f.innerHTML = `<label>${escHtml(label)}</label>
        <select id="f_${label}">${options}</select>
        ${a.help?`<div class="fhelp">${escHtml(a.help)}</div>`:''}`;
    }else if(a.type==='multi'){
      const def = Array.isArray(val) ? val.map(String) : (val ? [String(val)] : []);
      f.innerHTML = `<label>${escHtml(label)}</label>
        <div class="multi">${(a.choices||[]).map(c=>{
          const value = choiceValue(c);
          const text = choiceLabel(c);
          return `<label><input type="checkbox" value="${escAttr(value)}" ${def.includes(String(value))?'checked':''} data-multi="${escAttr(label)}">${escHtml(text)}</label>`;
        }).join('')}</div>
        ${a.help?`<div class="fhelp">${escHtml(a.help)}</div>`:''}`;
    }else{
      const t = a.type==='int' ? 'number' : 'text';
      f.innerHTML = `<label>${escHtml(label)}</label>
        <input type="${t}" id="f_${label}" value="${escAttr(val!==undefined&&val!==null?val:'')}" placeholder="${escAttr(a.help||'')}">
        ${a.help?`<div class="fhelp">${escHtml(a.help)}</div>`:''}`;
    }
    p.appendChild(f);
  });

  const actions = document.createElement('div');
  actions.className = 'form-actions';
  const btn = document.createElement('button');
  btn.className='btn-run'; btn.textContent='▶ 运行';
  btn.onclick = runScript;
  actions.appendChild(btn);
  if(['register_outlook_ruoyi'].includes(s.id)){
    const saveBtn = document.createElement('button');
    saveBtn.className = 'btn-run';
    saveBtn.textContent = '保存当前配置';
    saveBtn.onclick = saveScriptConfig;
    actions.appendChild(saveBtn);
    const resetBtn = document.createElement('button');
    resetBtn.className = 'btn-run';
    resetBtn.textContent = '清除保存';
    resetBtn.onclick = clearScriptConfig;
    actions.appendChild(resetBtn);
    const msg = document.createElement('span');
    msg.id = 'script-config-msg';
    msg.className = 'env-msg';
    actions.appendChild(msg);
  }
  p.appendChild(actions);
  const cmd = document.createElement('div'); cmd.className='cmd-line'; cmd.id='cmd-preview';
  p.appendChild(cmd);
}

function collectFormValues(s){
  const args = {};
  s.args.forEach(a=>{
    const label = a.flag.replace(/^--/,'');
    if(a.type==='bool'){
      args[a.flag] = $(`#f_${label}`).checked;
    }else if(a.type==='multi'){
      args[a.flag] = $$(`input[data-multi="${label}"]:checked`).map(x=>x.value);
    }else{
      args[a.flag] = $(`#f_${label}`).value.trim();
    }
  });
  return args;
}

function collectArgs(s){
  const args = {};
  s.args.forEach(a=>{
    const label = a.flag.replace(/^--/,'');
    if(a.type==='bool'){
      args[a.flag] = $(`#f_${label}`).checked;
    }else if(a.type==='multi'){
      args[a.flag] = $$(`input[data-multi="${label}"]:checked`).map(x=>x.value);
    }else{
      const v = $(`#f_${label}`).value.trim();
      if(v!=='') args[a.flag] = a.type==='int' ? parseInt(v,10) : v;
    }
  });
  return args;
}

async function saveScriptConfig(){
  if(!curSrc) return;
  const msg = $('#script-config-msg');
  const args = collectFormValues(curSrc);
  try{
    const r = await (await fetch(`/api/script-config/${encodeURIComponent(curSrc.id)}`,{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({args})
    })).json();
    if(r.ok){
      curSavedArgs = args;
      msg.textContent = `✓ 已保存 ${r.saved} 项`;
    }else{
      msg.textContent = '保存失败: '+(r.error||'');
    }
  }catch(e){
    msg.textContent = '保存失败: '+e;
  }
  setTimeout(()=>{ if(msg) msg.textContent=''; }, 3000);
}

async function clearScriptConfig(){
  if(!curSrc) return;
  const msg = $('#script-config-msg');
  try{
    const r = await (await fetch(`/api/script-config/${encodeURIComponent(curSrc.id)}`,{method:'DELETE'})).json();
    if(r.ok){
      curSavedArgs = {};
      renderForm(curSrc, {});
      const nextMsg = $('#script-config-msg');
      if(nextMsg) nextMsg.textContent = '✓ 已清除，已恢复默认值';
      setTimeout(()=>{ const m=$('#script-config-msg'); if(m) m.textContent=''; }, 3000);
    }else{
      msg.textContent = '清除失败: '+(r.error||'');
    }
  }catch(e){
    msg.textContent = '清除失败: '+e;
  }
}

// ---------------------------------------------------------------- 运行 + SSE 日志
async function runScript(){
  if(curRun && evtSrc){ evtSrc.close(); }
  const args = collectArgs(curSrc);
  const log = $('#log'); log.textContent='';
  $('#log-title').textContent = `运行日志 — ${curSrc.title}`;
  const r = await (await fetch('/api/run',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({script:curSrc.id, args})})).json();
  if(r.error){ log.textContent='错误: '+r.error; return; }
  curRun = r.run_id;
  $('#cmd-preview').textContent = '$ '+r.cmd;
  $('#btn-stop').disabled = false;
  evtSrc = new EventSource(`/api/logs/${curRun}`);
  evtSrc.onmessage = e=>{ LogAutoScroll.appendLogLineWithAutoScroll(log, e.data); };
  evtSrc.addEventListener('done', ()=>{ evtSrc.close(); $('#btn-stop').disabled = true; pollStatus(); loadTaskDashboard(); loadAccounts(); });
  evtSrc.onerror = ()=>{ evtSrc.close(); $('#btn-stop').disabled = true; };
}

$('#btn-stop').onclick = async ()=>{
  if(!curRun) return;
  await fetch(`/api/stop/${curRun}`,{method:'POST'});
  $('#btn-stop').disabled = true;
};

// ---------------------------------------------------------------- 配置页
async function loadEnv(){
  const data = await (await fetch('/api/env')).json();
  const wrap = $('#env-groups'); wrap.innerHTML='';
  data.groups.forEach(g=>{
    const box = document.createElement('div'); box.className='env-group';
    const tests = (g.tests||[]).map(t=>
      `<button class="btn-test" data-test="${t.target}">${t.label}</button>`).join('');
    box.innerHTML = `<div class="env-group-title">
        <span>${g.group}</span>
        <span class="test-area">${tests}<span class="test-result" data-result-for="${g.group}"></span></span>
      </div>`;
    g.items.forEach(it=>{
      const row = document.createElement('div'); row.className='env-item';
      const type = it.secret ? 'password':'text';
      const value = it.value || it.default || '';
      const control = it.type === 'choice'
        ? `<select data-env="${escAttr(it.key)}">${renderChoiceOptions(it.choices, value)}</select>`
        : `<input type="${type}" data-env="${escAttr(it.key)}" value="${escAttr(it.value||'')}"
                 placeholder="${escAttr(it.default? '默认 '+it.default : '')}">`;
      row.innerHTML = `
        <div class="k">${escHtml(it.key)}${it.required?'<span class="req">*</span>':''}</div>
        <div class="v">
          ${control}
          ${it.help?`<div class="ehelp">${escHtml(it.help)}</div>`:''}
        </div>`;
      box.appendChild(row);
    });
    // 绑定该组的测试按钮
    box.querySelectorAll('.btn-test').forEach(btn=>{
      btn.onclick = ()=> runTest(btn.dataset.test, btn);
    });
    wrap.appendChild(box);
  });
}

// 连通测试：把当前页面所有 .env 输入(含未保存的)一起发过去，用最新值测
async function runTest(target, btn){
  const env = {};
  $$('input[data-env],select[data-env]').forEach(i=>{ if(i.value!=='') env[i.dataset.env]=i.value; });
  const old = btn.textContent;
  btn.disabled = true; btn.textContent = '测试中…';
  const res = btn.closest('.env-group').querySelector('.test-result');
  res.textContent=''; res.className='test-result';
  try{
    const r = await (await fetch(`/api/test/${target}`,{method:'POST',
      headers:{'Content-Type':'application/json'}, body:JSON.stringify({env})})).json();
    res.textContent = (r.ok?'✓ ':'✗ ') + r.msg;
    res.classList.add(r.ok?'ok':'bad');
  }catch(e){
    res.textContent = '✗ 测试请求失败: '+e; res.classList.add('bad');
  }finally{
    btn.disabled=false; btn.textContent=old;
  }
}

$('#btn-save-env').onclick = async ()=>{
  const env = {};
  $$('input[data-env],select[data-env]').forEach(i=>{ env[i.dataset.env] = i.value; });
  const r = await (await fetch('/api/env',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({env})})).json();
  const msg = $('#env-msg');
  msg.textContent = r.ok ? `✓ 已保存 ${r.saved} 项` : ('保存失败: '+(r.error||''));
  setTimeout(()=>msg.textContent='', 3000);
};

// ---------------------------------------------------------------- 内嵌页 + 接码助手
function openEmbed(id){
  const e = EMBEDS.find(x=>x.id===id);
  if(!e) return;
  showView('embed');
  $$('.scriptbtn').forEach(b=>b.classList.toggle('active', b.dataset.embed===id));
  $('#embed-title').textContent = e.title;
  $('#embed-open').href = e.url;
  $('#embed-frame').src = e.url;
  const helper = $('#sms-helper');
  if(e.sms_helper){
    helper.style.display='block';
    if(e.sms_service_default) $('#sms-service').placeholder = e.sms_service_default;
    refreshRents();
  }else{
    helper.style.display='none';
  }
}

async function copyText(txt, btn){
  try{ await navigator.clipboard.writeText(txt); if(btn){const o=btn.textContent;btn.textContent='已复制';setTimeout(()=>btn.textContent=o,1200);} }
  catch(e){ alert('复制失败,请手动选择: '+txt); }
}

function fmtRemain(sec){
  sec=Math.max(0,sec); const m=Math.floor(sec/60), s=sec%60;
  return `${m}:${String(s).padStart(2,'0')}`;
}

async function refreshRents(){
  let data;
  try{ data = await (await fetch('/api/sms/rents')).json(); }catch(e){ return; }
  const wrap = $('#sms-rents'); wrap.innerHTML='';
  (data.rents||[]).forEach(r=>{
    const card = document.createElement('div'); card.className='rent-card';
    const codesHtml = r.codes.length
      ? r.codes.map(c=>`<span class="code-chip">${c}<button class="mini" data-copy="${c}">复制</button></span>`).join('')
      : '<span class="dim">暂无验证码</span>';
    card.innerHTML = `
      <div class="rent-phone">
        <b>+${r.phone}</b>
        <button class="mini" data-copy="${r.phone}">复制号码</button>
        <span class="multi-badge ${r.can_multi?'ok':'no'}">${r.can_multi?'多次接码':'单次'}</span>
        <span class="remain" data-remain="${r.remain}">剩 ${fmtRemain(r.remain)}</span>
      </div>
      <div class="rent-actions">
        <button class="btn-sm getcode" data-pkey="${r.pkey}">获取验证码</button>
        <button class="btn-sm release" data-pkey="${r.pkey}">完成/释放</button>
      </div>
      <div class="codes">${codesHtml}</div>
      <div class="sms-msg" data-msg="${r.pkey}"></div>`;
    wrap.appendChild(card);
  });
  // 绑定
  wrap.querySelectorAll('[data-copy]').forEach(b=> b.onclick=()=>copyText(b.dataset.copy,b));
  wrap.querySelectorAll('.getcode').forEach(b=> b.onclick=()=>getCode(b.dataset.pkey,b));
  wrap.querySelectorAll('.release').forEach(b=> b.onclick=()=>releaseNum(b.dataset.pkey));
  // 倒计时滴答
  if(smsTimer) clearInterval(smsTimer);
  if((data.rents||[]).length){
    smsTimer = setInterval(()=>{
      $$('.remain').forEach(el=>{
        let r=parseInt(el.dataset.remain,10)-1; el.dataset.remain=r;
        el.textContent = r>0 ? '剩 '+fmtRemain(r) : '已过期';
      });
    },1000);
  }
}

$('#btn-rent').onclick = async ()=>{
  const btn=$('#btn-rent'); const o=btn.textContent; btn.disabled=true; btn.textContent='租号中…';
  const body={ service: $('#sms-service').value.trim()||undefined, country: $('#sms-country').value.trim()||undefined,
    prefer_multi: $('#sms-prefer-multi') ? $('#sms-prefer-multi').checked : true };
  try{
    const r = await (await fetch('/api/sms/rent',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)})).json();
    if(!r.ok){ alert('获取号码失败: '+r.msg); }
    await refreshRents();
  }finally{ btn.disabled=false; btn.textContent=o; }
};

async function getCode(pkey, btn){
  const o=btn.textContent; btn.disabled=true; btn.textContent='等待验证码…';
  const msg = document.querySelector(`[data-msg="${pkey}"]`);
  if(msg) msg.textContent='';
  try{
    const r = await (await fetch('/api/sms/code',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pkey})})).json();
    if(r.ok){ await refreshRents(); }
    else if(msg){ msg.textContent = (r.expired?'⏰ ':'') + r.msg; }
  }finally{ btn.disabled=false; btn.textContent=o; }
}

async function releaseNum(pkey){
  await fetch('/api/sms/release',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pkey})});
  await refreshRents();
}

// ---------------------------------------------------------------- 邮箱池
async function loadMailpool(){
  try{
    const d = await (await fetch('/api/mailpool')).json();
    $('#mailpool-total').textContent = `当前池中 ${d.total} 个邮箱`;
  }catch(e){}
}

$('#btn-import-mail').onclick = async ()=>{
  const text = $('#mailpool-input').value;
  if(!text.trim()){ $('#mailpool-msg').textContent='请先粘贴邮箱'; return; }
  const btn=$('#btn-import-mail'); const o=btn.textContent; btn.disabled=true; btn.textContent='导入中…';
  const msg=$('#mailpool-msg'); msg.textContent='';
  try{
    const r = await (await fetch('/api/mailpool',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({text})})).json();
    if(r.ok){
      let m = `✓ 导入 ${r.added}，跳过重复 ${r.skipped}`;
      if(r.bad) m += `，格式错误 ${r.bad}`;
      m += `，池中共 ${r.total}`;
      msg.textContent = m;
      if(r.bad && r.bad_samples.length) msg.textContent += `（错误样例：${r.bad_samples[0]}…）`;
      $('#mailpool-total').textContent = `当前池中 ${r.total} 个邮箱`;
      if(r.added) $('#mailpool-input').value='';
    }else{ msg.textContent='导入失败: '+(r.msg||''); }
  }catch(e){ msg.textContent='导入请求失败: '+e; }
  finally{ btn.disabled=false; btn.textContent=o; }
};

// ---------------------------------------------------------------- 任务记录 / 统计
function fmtNum(v, digits=2){
  const n = Number(v || 0);
  return Number.isFinite(n) ? n.toFixed(digits) : '0.00';
}

function formatDisplayTime(v){
  const s = String(v || '').trim();
  if(!s) return '';
  return s
    .replace('T', ' ')
    .replace(/\.\d+$/, '')
    .replace(/Z$/, '');
}

function renderBars(target, items, cls=''){
  const el = $(target);
  if(!el) return;
  const data = items || [];
  const max = Math.max(1, ...data.map(x=>Number(x.value||0)));
  el.innerHTML = data.length ? data.map(item=>{
    const value = Number(item.value||0);
    const h = Math.max(2, Math.round((value/max)*80));
    return `<div class="mini-bar">
      <div class="top">${escHtml(value)}</div>
      <div class="bar ${cls}" style="height:${h}px"></div>
      <div class="bottom">${escHtml(item.day || item.label || '')}</div>
    </div>`;
  }).join('') : '<div class="dim">暂无数据</div>';
}

function renderLineChart(target, items, cls=''){
  const el = $(target);
  if(!el) return;
  const data = items || [];
  if(!data.length){
    el.innerHTML = '<div class="dim">暂无数据</div>';
    return;
  }
  const width = 640;
  const height = 180;
  const left = 28;
  const right = 20;
  const top = 20;
  const bottom = 34;
  const innerWidth = width - left - right;
  const innerHeight = height - top - bottom;
  const max = Math.max(1, ...data.map(x=>Number(x.value || 0)));
  const step = data.length > 1 ? innerWidth / (data.length - 1) : 0;
  const points = data.map((item, idx)=>{
    const value = Number(item.value || 0);
    const x = left + (step * idx);
    const y = top + innerHeight - ((value / max) * innerHeight);
    return {x, y, value, label: String(item.day || item.label || '')};
  });
  const color = cls === 'green' ? '#16a34a' : '#3b82f6';
  const polyline = points.map(p=>`${p.x},${p.y}`).join(' ');
  const circles = points.map(p=>`<circle cx="${p.x}" cy="${p.y}" r="3" fill="${color}"></circle>`).join('');
  const labels = points.map(p=>`<text x="${p.x}" y="${height - 6}" text-anchor="middle" font-size="12" fill="#94a3b8">${escHtml(p.label)}</text>`).join('');
  const values = points.map(p=>`<text x="${p.x}" y="${Math.max(10, p.y - 6)}" text-anchor="middle" font-size="12" fill="#e2e8f0">${escHtml(p.value)}</text>`).join('');
  el.innerHTML = `<svg viewBox="0 0 ${width} ${height}" width="100%" height="180" aria-label="trend line chart" role="img">
    <line x1="${left}" y1="${top + innerHeight}" x2="${width - right}" y2="${top + innerHeight}" stroke="#334155" stroke-width="1"></line>
    <polyline fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" points="${polyline}"></polyline>
    ${circles}
    ${values}
    ${labels}
  </svg>`;
}

function renderTaskCards(cards){
  const wrap = $('#task-cards');
  if(!wrap) return;
  const items = [
    ['总任务数', cards.total_tasks || 0],
    ['运行中', cards.running_tasks || 0],
    ['成功账号数', cards.total_success_accounts || 0],
    ['平均任务耗时(s)', fmtNum(cards.avg_task_duration_seconds || 0)],
    ['平均成功耗时(s)', fmtNum(cards.avg_success_duration_seconds || 0)],
  ];
  wrap.innerHTML = items.map(([label, value])=>`<div class="stat-card"><div class="label">${escHtml(label)}</div><div class="value">${escHtml(value)}</div></div>`).join('');
}

function renderTaskTable(items){
  const wrap = $('#task-table-wrap');
  if(!wrap) return;
  wrap.innerHTML = `<div class="table-wrap"><table class="data-table">
    <thead><tr>
      <th>脚本</th><th>开始</th><th>结束</th><th>耗时</th>
      <th>成功</th><th>失败</th><th>no_graph</th><th>平均成功耗时</th><th>状态</th><th>操作</th>
    </tr></thead>
    <tbody>
      ${(items||[]).map(it=>`<tr>
        <td>${escHtml(it.script_title || it.script_id || '')}</td>
        <td>${escHtml(formatDisplayTime(it.started_at || ''))}</td>
        <td>${escHtml(formatDisplayTime(it.ended_at || ''))}</td>
        <td>${escHtml(it.duration_seconds ?? '')}</td>
        <td>${escHtml(it.success_count ?? 0)}</td>
        <td>${escHtml(it.fail_count ?? 0)}</td>
        <td>${escHtml(it.no_graph_count ?? 0)}</td>
        <td>${escHtml(it.avg_success_duration_seconds ?? '')}</td>
        <td>${escHtml(it.status || '')}</td>
        <td><div class="table-actions">
          <button class="btn-lite" data-task-view="${escAttr(it.id ?? '')}">查看</button>
          <button class="btn-lite" data-task-log="${escAttr(it.id ?? '')}" data-task-log-title="${escAttr(it.script_title || it.script_id || '')}">日志</button>
          ${String(it.status || '').toLowerCase() === 'running' ? `<button class="btn-stop" data-task-stop="${escAttr(it.run_id ?? '')}">停止</button>` : ''}
          ${String(it.status || '').toLowerCase() === 'running' ? '' : `<button class="btn-lite" data-task-delete="${escAttr(it.id ?? '')}">删除</button>`}
        </div></td>
      </tr>`).join('') || '<tr><td colspan="10" class="dim">暂无任务记录</td></tr>'}
    </tbody>
  </table></div>`;
  $$('[data-task-view]', wrap).forEach(btn=> btn.onclick = ()=> loadTaskDetail(btn.dataset.taskView));
  $$('[data-task-log]', wrap).forEach(btn=> btn.onclick = ()=> openTaskLogModal(btn.dataset.taskLog, btn.dataset.taskLogTitle || '任务日志'));
  $$('[data-task-stop]', wrap).forEach(btn=> btn.onclick = ()=> stopTaskRun(btn.dataset.taskStop, btn));
  $$('[data-task-delete]', wrap).forEach(btn=> btn.onclick = ()=> deleteTaskRun(btn.dataset.taskDelete, btn));
}

function renderAccountTable(target, items, optsOrSelectable=false){
  const wrap = $(target);
  if(!wrap) return;
  const opts = typeof optsOrSelectable === 'object'
    ? {selectable: !!optsOrSelectable.selectable, actions: !!optsOrSelectable.actions}
    : {selectable: !!optsOrSelectable, actions: false};
  const rows = items || [];
  const selectable = !!opts.selectable;
  const actions = !!opts.actions;
  wrap.innerHTML = `<div class="table-wrap"><table class="data-table">
    <thead><tr>
      ${selectable ? '<th><input type="checkbox" id="accounts-check-all"></th>' : ''}
      <th>任务ID</th><th>邮箱</th><th>密码</th><th>创建时间</th><th>注册IP</th><th>注册地区</th><th>状态</th>
      ${actions ? '<th>操作</th>' : ''}
    </tr></thead>
    <tbody>
      ${rows.map(it=>`<tr>
        ${selectable ? `<td><input type="checkbox" data-account-id="${escAttr(it.id || '')}" ${selectedAccountIds.has(String(it.id || '')) ? 'checked' : ''}></td>` : ''}
        <td>${escHtml(it.task_run_id ?? '')}</td>
        <td>${escHtml(it.email || '')}</td>
        <td>${escHtml(it.password || '')}</td>
        <td>${escHtml(formatDisplayTime(it.created_at || it.generated_at || ''))}</td>
        <td>${escHtml(it.register_ip || '')}</td>
        <td>${escHtml(it.register_region || '')}</td>
        <td>${escHtml(it.status || '')}</td>
        ${actions ? `<td><div class="table-actions"><button class="btn-lite" data-edit-account="${escAttr(it.id || '')}">编辑</button><button class="btn-lite" data-copy-account="${escAttr(it.id || '')}">复制</button></div></td>` : ''}
      </tr>`).join('') || `<tr><td colspan="${(selectable ? 1 : 0) + 7 + (actions ? 1 : 0)}" class="dim">暂无账号数据</td></tr>`}
    </tbody>
  </table></div>`;
  if(selectable){
    $$('input[data-account-id]', wrap).forEach(box=>{
      box.onchange = ()=>{
        const id = String(box.dataset.accountId || '');
        if(!id) return;
        if(box.checked) selectedAccountIds.add(id);
        else selectedAccountIds.delete(id);
      };
    });
    const all = $('#accounts-check-all', wrap);
    if(all){
      all.checked = rows.length > 0 && rows.every(it=> selectedAccountIds.has(String(it.id || '')));
      all.onchange = ()=>{
        $$('input[data-account-id]', wrap).forEach(box=>{
          box.checked = all.checked;
          const id = String(box.dataset.accountId || '');
          if(!id) return;
          if(all.checked) selectedAccountIds.add(id);
          else selectedAccountIds.delete(id);
        });
      };
    }
  }
  if(actions){
    $$('[data-copy-account]', wrap).forEach(btn=> btn.onclick = ()=>{
      const item = rows.find(it=> String(it.id || '') === String(btn.dataset.copyAccount || ''));
      if(item) copyText(formatAccountTxtLine(item), btn);
    });
    $$('[data-edit-account]', wrap).forEach(btn=> btn.onclick = ()=>{
      const item = rows.find(it=> String(it.id || '') === String(btn.dataset.editAccount || ''));
      if(item) openAccountEditModal(item);
    });
  }
}

function formatAccountTxtLine(item){
  return [item?.email || '', item?.password || '', item?.refresh_token || '', item?.client_id || '']
    .map(v=> String(v || '').trim())
    .join('----')
    .replace(/-+$/, '');
}

function buildPageItems(page, totalPages){
  if(totalPages <= 5) return Array.from({length: totalPages}, (_, i)=> i + 1);
  if(page <= 3) return [1, 2, 3, 4, '...', totalPages];
  if(page >= totalPages - 2) return [1, '...', totalPages - 3, totalPages - 2, totalPages - 1, totalPages];
  return [1, '...', page - 1, page, page + 1, '...', totalPages];
}

function renderPagination(target, state, onChange){
  const wrap = $(target);
  if(!wrap) return;
  const total = Number(state?.total || 0);
  const pageSize = Number(state?.pageSize || 20);
  const totalPages = Math.max(1, Math.ceil(total / pageSize) || 1);
  const page = Math.min(Math.max(1, Number(state?.page || 1)), totalPages);
  state.page = page;
  const pageItems = buildPageItems(page, totalPages);
  wrap.innerHTML = `<div class="pagination-meta">共 ${escHtml(total)} 条，第 ${escHtml(page)} / ${escHtml(totalPages)} 页</div>
    <div class="pagination-actions">
      <button class="btn-lite" data-page-nav="prev" ${page <= 1 ? 'disabled' : ''}>上一页</button>
      <div class="pagination-pages">
        ${pageItems.map(item=> item === '...'
          ? '<span class="pagination-ellipsis">...</span>'
          : `<button class="btn-lite ${Number(item) === page ? 'is-active' : ''}" data-page-num="${escAttr(item)}">${escHtml(item)}</button>`
        ).join('')}
      </div>
      <button class="btn-lite" data-page-nav="next" ${page >= totalPages ? 'disabled' : ''}>下一页</button>
    </div>`;
  $$('[data-page-num]', wrap).forEach(btn=>{
    btn.onclick = ()=>{
      const nextPage = Number(btn.dataset.pageNum || page);
      if(nextPage !== state.page) onChange(nextPage);
    };
  });
  const prev = $('[data-page-nav="prev"]', wrap);
  const next = $('[data-page-nav="next"]', wrap);
  if(prev) prev.onclick = ()=> state.page > 1 && onChange(state.page - 1);
  if(next) next.onclick = ()=> state.page < totalPages && onChange(state.page + 1);
}

function renderAccountPagination(){
  renderPagination('#accounts-pagination', accountState, nextPage=>{
    accountState.page = nextPage;
    loadAccounts();
  });
}

function renderTaskPagination(){
  renderPagination('#tasks-pagination', taskState, nextPage=>{
    taskState.page = nextPage;
    loadTaskDashboard();
  });
}

function renderTaskDetailPagination(){
  renderPagination('#task-detail-accounts-pagination', taskDetailState, nextPage=>{
    taskDetailState.page = nextPage;
    loadTaskDetailAccounts();
  });
}

function clearTaskDetail(message='点击上方查看任务详情'){
  const title = $('#task-detail-title');
  const meta = $('#task-detail-meta');
  const accounts = $('#task-detail-accounts');
  const pager = $('#task-detail-accounts-pagination');
  taskDetailState.taskId = '';
  taskDetailState.page = 1;
  taskDetailState.total = 0;
  taskDetailState.loaded = [];
  if(title) title.textContent = '';
  if(meta) meta.innerHTML = '';
  if(accounts) accounts.innerHTML = `<div class="dim">${escHtml(message)}</div>`;
  if(pager) pager.innerHTML = '';
}

async function loadTaskDashboard(){
  const msg = $('#task-msg');
  try{
    const qs = new URLSearchParams({
      page: String(taskState.page || 1),
      page_size: String(taskState.pageSize || 20),
    });
    const [statsRes, tasksRes] = await Promise.all([
      fetch('/api/stats/overview?days=7'),
      fetch('/api/task-runs?' + qs.toString()),
    ]);
    const stats = await statsRes.json();
    const tasks = await tasksRes.json();
    taskState.page = Number(tasks.page || taskState.page || 1);
    taskState.pageSize = Number(tasks.page_size || taskState.pageSize || 20);
    taskState.total = Number(tasks.total || 0);
    taskState.loaded = tasks.items || [];
    renderTaskCards(stats.cards || {});
    renderLineChart('#chart-task-trend', stats.task_trend || []);
    renderLineChart('#chart-success-trend', stats.success_trend || [], 'green');
    renderBars('#chart-regions', stats.top_regions || [], 'dim');
    renderTaskTable(taskState.loaded);
    renderTaskPagination();
    msg.textContent = stats.error || tasks.error || '';
  }catch(e){
    taskState.loaded = [];
    taskState.total = 0;
    renderTaskTable([]);
    renderTaskPagination();
    if(msg) msg.textContent = '加载任务记录失败: ' + e;
  }
}

async function loadTaskDetail(taskId){
  const meta = $('#task-detail-meta');
  const title = $('#task-detail-title');
  const accounts = $('#task-detail-accounts');
  try{
    taskDetailState.taskId = String(taskId || '');
    taskDetailState.page = 1;
    if(title) title.textContent = `任务详情：${taskId}`;
    if(meta) meta.innerHTML = '<div class="dim">加载任务详情中...</div>';
    if(accounts) accounts.innerHTML = '<div class="dim">加载账号列表中...</div>';
    const detailPromise = fetch(`/api/task-runs/${taskId}`).then(r=>r.json());
    const accountsPromise = loadTaskDetailAccounts();
    const detail = await detailPromise;
    title.textContent = `任务详情：${detail.script_title || detail.script_id || taskId}`;
    meta.innerHTML = [
      ['脚本', detail.script_title || detail.script_id || ''],
      ['状态', detail.status || ''],
      ['开始', formatDisplayTime(detail.started_at || '')],
      ['结束', formatDisplayTime(detail.ended_at || '')],
      ['耗时', detail.duration_seconds ?? ''],
      ['成功', detail.success_count ?? 0],
      ['失败', detail.fail_count ?? 0],
      ['no_graph', detail.no_graph_count ?? 0],
      ['成功账号平均耗时', detail.avg_success_duration_seconds ?? ''],
    ].map(([k,v])=>`<div class="detail-item"><div class="k">${escHtml(k)}</div><div class="v">${escHtml(v)}</div></div>`).join('');
    await accountsPromise;
  }catch(e){
    clearTaskDetail('任务详情加载失败');
    if(title) title.textContent = '任务详情加载失败';
    if(meta) meta.innerHTML = `<div class="dim">${escHtml(String(e || ''))}</div>`;
  }
}

async function loadTaskDetailAccounts(){
  if(!taskDetailState.taskId) return;
  const accounts = $('#task-detail-accounts');
  try{
    if(accounts) accounts.innerHTML = '<div class="dim">加载账号列表中...</div>';
    const qs = new URLSearchParams({
      page: String(taskDetailState.page || 1),
      page_size: String(taskDetailState.pageSize || 20),
    });
    const data = await (await fetch(`/api/task-runs/${encodeURIComponent(taskDetailState.taskId)}/accounts?` + qs.toString())).json();
    taskDetailState.page = Number(data.page || taskDetailState.page || 1);
    taskDetailState.pageSize = Number(data.page_size || taskDetailState.pageSize || 20);
    taskDetailState.total = Number(data.total || 0);
    taskDetailState.loaded = data.items || [];
    renderAccountTable('#task-detail-accounts', taskDetailState.loaded, {selectable:false, actions:true});
    renderTaskDetailPagination();
  }catch(e){
    taskDetailState.loaded = [];
    taskDetailState.total = 0;
    renderAccountTable('#task-detail-accounts', [], {selectable:false, actions:true});
    renderTaskDetailPagination();
    if(accounts) accounts.insertAdjacentHTML('beforeend', `<div class="dim">加载账号失败：${escHtml(String(e || ''))}</div>`);
  }
}

async function stopTaskRun(runId, btn){
  if(!runId) return;
  const old = btn ? btn.textContent : '';
  if(btn){ btn.disabled = true; btn.textContent = '停止中...'; }
  try{
    await fetch(`/api/stop/${encodeURIComponent(runId)}`, {method:'POST'});
  }finally{
    await Promise.all([loadTaskDashboard(), loadAccounts()]);
    if(btn){ btn.disabled = false; btn.textContent = old || '停止'; }
  }
}

async function deleteTaskRun(taskId, btn){
  if(!taskId) return;
  if(globalThis.confirm && !globalThis.confirm(`确认删除任务 ${taskId} 吗？`)) return;
  const old = btn ? btn.textContent : '';
  if(btn){ btn.disabled = true; btn.textContent = '删除中...'; }
  try{
    const res = await fetch(`/api/task-runs/${encodeURIComponent(taskId)}`, {method:'DELETE'});
    const data = await res.json().catch(()=>({}));
    if(!res.ok || !data.ok){
      throw new Error(data.error || res.statusText || '删除失败');
    }
    if(String(taskDetailState.taskId || '') === String(taskId || '')){
      clearTaskDetail('任务已删除');
    }
    if(taskState.page > 1 && taskState.loaded.length <= 1){
      taskState.page -= 1;
    }
    await loadTaskDashboard();
  }catch(e){
    const msg = $('#task-msg');
    if(msg) msg.textContent = '删除任务失败: ' + e;
  }finally{
    if(btn){ btn.disabled = false; btn.textContent = old || '删除'; }
  }
}

async function openTaskLogModal(taskId, title='任务日志'){
  const modal = $('#task-log-modal');
  const titleEl = $('#task-log-title');
  const logEl = $('#task-log-content');
  if(titleEl) titleEl.textContent = title;
  if(logEl) logEl.textContent = '加载中...';
  if(modal) modal.style.display = 'flex';
  try{
    const text = await (await fetch(`/api/task-runs/${taskId}/log`)).text();
    if(logEl) logEl.textContent = text || '暂无日志';
  }catch(e){
    if(logEl) logEl.textContent = '日志加载失败: ' + e;
  }
}

function closeTaskLogModal(){
  const modal = $('#task-log-modal');
  if(modal) modal.style.display = 'none';
}

function syncAccountFilters(resetPage=false){
  accountState.filters = {
    taskId: ($('#accounts-task-filter')?.value || '').trim(),
    email: ($('#accounts-email-filter')?.value || '').trim(),
    status: ($('#accounts-status-filter')?.value || '').trim(),
  };
  if(resetPage) accountState.page = 1;
}

async function loadAccounts(opts={}){
  const msg = $('#accounts-msg');
  if(opts.syncFilters) syncAccountFilters(!!opts.resetPage);
  const filters = accountState.filters || {taskId:'', email:'', status:''};
  const qs = new URLSearchParams({
    page: String(accountState.page || 1),
    page_size: String(accountState.pageSize || 20),
    sort_by: 'created_at',
    sort_dir: 'desc',
  });
  if(filters.taskId) qs.set('task_run_id', filters.taskId);
  if(filters.email) qs.set('email', filters.email);
  if(filters.status) qs.set('status', filters.status);
  try{
    const data = await (await fetch('/api/accounts?' + qs.toString())).json();
    accountState.page = Number(data.page || accountState.page || 1);
    accountState.pageSize = Number(data.page_size || accountState.pageSize || 20);
    accountState.total = Number(data.total || 0);
    accountState.loaded = data.items || [];
    renderAccountTable('#accounts-table-wrap', accountState.loaded, {selectable:true, actions:true});
    renderAccountPagination();
    if(msg) msg.textContent = data.error || '';
  }catch(e){
    accountState.loaded = [];
    renderAccountTable('#accounts-table-wrap', [], {selectable:true, actions:true});
    renderAccountPagination();
    if(msg) msg.textContent = '加载账号失败: ' + e;
  }
}

async function importAccounts(){
  const text = $('#accounts-import-input').value;
  const msg = $('#accounts-msg');
  if(!text.trim()){
    if(msg) msg.textContent = '请先粘贴账号';
    return;
  }
  const r = await (await fetch('/api/accounts/import', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({text}),
  })).json();
  if(msg){
    msg.textContent = r.ok ? `导入 ${r.added}，跳过 ${r.skipped}，坏行 ${r.bad}` : ('导入失败: ' + (r.error || ''));
  }
  if(r.ok && r.added){
    selectedAccountIds.clear();
    accountState.page = 1;
    $('#accounts-import-input').value = '';
    const file = $('#accounts-import-file');
    if(file) file.value = '';
    closeAccountsImportModal();
  }
  loadAccounts();
  loadTaskDashboard();
}

async function exportAccounts(){
  const filters = accountState.filters || {taskId:'', email:'', status:''};
  const qs = new URLSearchParams();
  if(filters.taskId) qs.set('task_run_id', filters.taskId);
  if(filters.email) qs.set('email', filters.email);
  if(filters.status) qs.set('status', filters.status);
  if(selectedAccountIds.size) qs.set('ids', [...selectedAccountIds].join(','));
  const res = await fetch('/api/accounts/export?' + qs.toString());
  const text = await res.text();
  const blob = new Blob([text], {type:'text/plain;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filters.taskId ? `accounts_task_${filters.taskId}.txt` : 'accounts_export.txt';
  a.click();
  URL.revokeObjectURL(a.href);
}

function openAccountsImportModal(){
  const modal = $('#accounts-import-modal');
  if(modal) modal.style.display = 'flex';
}

function closeAccountsImportModal(){
  const modal = $('#accounts-import-modal');
  if(modal) modal.style.display = 'none';
}

function openAccountEditModal(item){
  $('#accounts-edit-id').value = item?.id || '';
  $('#accounts-edit-email').value = item?.email || '';
  $('#accounts-edit-password').value = item?.password || '';
  $('#accounts-edit-client-id').value = item?.client_id || '';
  $('#accounts-edit-refresh-token').value = item?.refresh_token || '';
  const modal = $('#accounts-edit-modal');
  if(modal) modal.style.display = 'flex';
}

function closeAccountEditModal(){
  const modal = $('#accounts-edit-modal');
  if(modal) modal.style.display = 'none';
}

async function saveAccountEdit(){
  const id = ($('#accounts-edit-id')?.value || '').trim();
  if(!id) return;
  const msg = $('#accounts-msg');
  const body = {
    email: ($('#accounts-edit-email')?.value || '').trim(),
    password: ($('#accounts-edit-password')?.value || '').trim(),
    client_id: ($('#accounts-edit-client-id')?.value || '').trim(),
    refresh_token: ($('#accounts-edit-refresh-token')?.value || '').trim(),
  };
  const res = await fetch(`/api/accounts/${encodeURIComponent(id)}`, {
    method:'PUT',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify(body),
  });
  const data = await res.json();
  if(!res.ok || !data.ok){
    if(msg) msg.textContent = '保存失败: ' + (data.error || res.statusText || '');
    return;
  }
  closeAccountEditModal();
  if(msg) msg.textContent = '账号已更新';
  loadAccounts();
}

async function readAccountsImportFile(){
  const file = $('#accounts-import-file')?.files?.[0];
  if(!file) return;
  $('#accounts-import-input').value = await file.text();
}

function runAccountQuery(){
  selectedAccountIds.clear();
  loadAccounts({syncFilters:true, resetPage:true});
}

$('#btn-refresh-tasks') && ($('#btn-refresh-tasks').onclick = loadTaskDashboard);
$('#btn-refresh-accounts') && ($('#btn-refresh-accounts').onclick = loadAccounts);
$('#btn-import-accounts') && ($('#btn-import-accounts').onclick = importAccounts);
$('#btn-export-accounts') && ($('#btn-export-accounts').onclick = exportAccounts);
$('#btn-query-accounts') && ($('#btn-query-accounts').onclick = runAccountQuery);
$('#tasks-page-size') && ($('#tasks-page-size').onchange = e=>{ taskState.pageSize = Number(e.target.value || 20); taskState.page = 1; loadTaskDashboard(); });
$('#accounts-page-size') && ($('#accounts-page-size').onchange = e=>{ accountState.pageSize = Number(e.target.value || 20); accountState.page = 1; loadAccounts(); });
$('#btn-open-import-accounts') && ($('#btn-open-import-accounts').onclick = openAccountsImportModal);
$('#btn-close-import-accounts') && ($('#btn-close-import-accounts').onclick = closeAccountsImportModal);
$('#accounts-import-file') && ($('#accounts-import-file').onchange = readAccountsImportFile);
$('#btn-close-edit-accounts') && ($('#btn-close-edit-accounts').onclick = closeAccountEditModal);
$('#btn-save-edit-accounts') && ($('#btn-save-edit-accounts').onclick = saveAccountEdit);
$('#btn-close-task-log') && ($('#btn-close-task-log').onclick = closeTaskLogModal);
taskState.pageSize = Number($('#tasks-page-size')?.value || 20);
accountState.pageSize = Number($('#accounts-page-size')?.value || 20);
syncAccountFilters(false);
clearTaskDetail();

// ---------------------------------------------------------------- 启动
loadScripts();
pollStatus();
