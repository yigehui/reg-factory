// reg-factory WebUI 前端逻辑（原生 JS，无构建）
let SCRIPTS = [];
let EMBEDS = [];
let curRun = null;     // 当前运行 run_id
let curSrc = null;     // 当前选中脚本
let evtSrc = null;     // EventSource
let smsTimer = null;   // 接码助手倒计时刷新
let curSavedArgs = {}; // 当前脚本已保存的参数

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

// ---------------------------------------------------------------- 条件字段 visible_if
// schema 可在参数上加 visible_if: {flag, equals|in}
// 依赖字段值满足时显示该行，否则隐藏；隐藏字段不进 collectArgs/collectFormValues
function fieldVisible(a, values){
  const cond = a.visible_if;
  if(!cond || !cond.flag) return true;
  const cur = values[cond.flag];
  const curStr = (cur===undefined || cur===null) ? '' : String(cur);
  if(cond.in) return cond.in.map(String).includes(curStr);
  return cond.equals !== undefined ? String(cond.equals)===curStr : true;
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
  $('#view-run').style.display  = v==='run' ? 'flex' : 'none';
  $('#view-env').style.display  = v==='env' ? 'block' : 'none';
  $('#view-embed').style.display = v==='embed' ? 'block' : 'none';
  $('#view-mailpool').style.display = v==='mailpool' ? 'block' : 'none';
  $$('.navbtn').forEach(b=>b.classList.toggle('active', b.dataset.view===v));
  if(v==='env') loadEnv();
  if(v==='mailpool') loadMailpool();
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

  // 依赖字段当前值快照(供 visible_if 判定)
  const values = {};
  s.args.forEach(a=>{
    const label = a.flag.replace(/^--/,'');
    const val = argValue(a, saved);
    if(a.type==='multi'){
      values[a.flag] = Array.isArray(val) ? val.map(String) : (val ? [String(val)] : []);
    }else{
      values[a.flag] = (val===undefined||val===null) ? '' : val;
    }
  });

  const applyVisibility = ()=>{
    s.args.forEach(a=>{
      const label = a.flag.replace(/^--/,'');
      // 依赖字段值实时取最新控件值
      const depEl = $(`#f_${label}`);
      if(depEl && a.type!=='multi'){
        if(a.type==='bool'){ values[a.flag] = depEl.checked; }
        else if(a.type==='int'){ const raw=depEl.value.trim(); values[a.flag]= raw===''? '': parseInt(raw,10); }
        else { values[a.flag] = depEl.value.trim(); }
      }
      if(a.type==='multi'){
        values[a.flag] = $$(`input[data-multi="${label}"]:checked`).map(x=>x.value);
      }
    });
    s.args.forEach(a=>{
      const row = document.getElementById(`row_${a.flag}`);
      if(!row) return;
      row.style.display = fieldVisible(a, values) ? '' : 'none';
    });
  };

  s.args.forEach(a=>{
    const f = document.createElement('div'); f.className='field';
    f.id = `row_${a.flag}`;
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
      const t = a.type==='int' ? 'number' : (a.secret ? 'password' : 'text');
      f.innerHTML = `<label>${escHtml(label)}</label>
        <input type="${t}" id="f_${label}" value="${escAttr(val!==undefined&&val!==null?val:'')}" placeholder="${escAttr(a.help||'')}" ${a.secret?'autocomplete="new-password"':''}>
        ${a.help?`<div class="fhelp">${escHtml(a.help)}</div>`:''}`;
    }
    p.appendChild(f);
    // 任何字段变化都触发显隐重算(被依赖的字段如 --proxy-source 自身没有 visible_if,
    // 它的 change 也要触发,否则依赖它的 --proxy-url 等不会显隐更新)
    const el = $(`#f_${label}`);
    if(el){ el.addEventListener('change', applyVisibility); }
  });

  applyVisibility();

  const actions = document.createElement('div');
  actions.className = 'form-actions';
  const btn = document.createElement('button');
  btn.className='btn-run'; btn.textContent='▶ 运行';
  btn.onclick = runScript;
  actions.appendChild(btn);
  if(['register_outlook_ruoyi', 'unlock_outlook'].includes(s.id)){
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
  const values = {};
  s.args.forEach(a=>{
    const label = a.flag.replace(/^--/,'');
    const depEl = $(`#f_${label}`);
    if(a.type==='bool'){
      values[a.flag] = depEl ? depEl.checked : a.default;
    }else if(a.type==='multi'){
      values[a.flag] = $$(`input[data-multi="${label}"]:checked`).map(x=>x.value);
    }else{
      values[a.flag] = depEl ? depEl.value.trim() : (a.default||'');
    }
  });
  s.args.forEach(a=>{
    if(!fieldVisible(a, values)) return;
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
  const values = {};
  s.args.forEach(a=>{
    const label = a.flag.replace(/^--/,'');
    const depEl = $(`#f_${label}`);
    if(a.type==='bool'){
      values[a.flag] = depEl ? depEl.checked : a.default;
    }else if(a.type==='multi'){
      values[a.flag] = $$(`input[data-multi="${label}"]:checked`).map(x=>x.value);
    }else{
      values[a.flag] = depEl ? depEl.value.trim() : (a.default||'');
    }
  });
  s.args.forEach(a=>{
    if(!fieldVisible(a, values)) return;
    const label = a.flag.replace(/^--/,'');
    if(a.type==='bool'){
      if($(`#f_${label}`).checked) args[a.flag] = true;
    }else if(a.type==='multi'){
      const v = $$(`input[data-multi="${label}"]:checked`).map(x=>x.value);
      if(v.length) args[a.flag] = v;
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
  evtSrc.addEventListener('done', ()=>{ evtSrc.close(); $('#btn-stop').disabled = true; pollStatus(); });
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

// ---------------------------------------------------------------- 启动
loadScripts();
pollStatus();
