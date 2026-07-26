const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const source = fs.readFileSync(path.join(__dirname, 'webui/static/app.js'), 'utf8');

function extractFunction(name){
  const start = source.indexOf(`function ${name}(`);
  assert.notStrictEqual(start, -1, `missing function ${name}`);
  let i = source.indexOf('{', start);
  assert.notStrictEqual(i, -1, `missing body for ${name}`);
  let depth = 0;
  for(let j=i; j<source.length; j++){
    const ch = source[j];
    if(ch === '{') depth += 1;
    if(ch === '}'){
      depth -= 1;
      if(depth === 0) return source.slice(start, j + 1);
    }
  }
  throw new Error(`unterminated function ${name}`);
}

function buildRuntime(names){
  const els = {};
  const sandbox = {
    console,
    Math,
    Number,
    String,
    Object,
    Array,
    els,
    $: (selector) => els[selector] || null,
    $$: () => [],
    fmtNum: (v) => Number(v).toFixed(2),
  };
  vm.createContext(sandbox);
  for(const name of names){
    vm.runInContext(extractFunction(name), sandbox, {filename: `app.js:${name}`});
  }
  return sandbox;
}

// RED: ISO 时间应显示成空格分隔格式
{
  const rt = buildRuntime(['formatDisplayTime']);
  assert.strictEqual(rt.formatDisplayTime('2026-07-26T16:26:37'), '2026-07-26 16:26:37');
  assert.strictEqual(rt.formatDisplayTime('2026-07-26 16:26:37'), '2026-07-26 16:26:37');
  assert.strictEqual(rt.formatDisplayTime(''), '');
}

// RED: 趋势图应输出 SVG 折线，不是柱条
{
  const rt = buildRuntime(['escHtml', 'renderLineChart']);
  rt.els['#chart-task-trend'] = { innerHTML: '' };
  rt.renderLineChart('#chart-task-trend', [
    {day: '07-20', value: 2},
    {day: '07-21', value: 5},
    {day: '07-22', value: 3},
  ], 'blue');
  const html = rt.els['#chart-task-trend'].innerHTML;
  assert.ok(html.includes('<svg'), 'line chart should render svg');
  assert.ok(html.includes('<polyline'), 'line chart should render polyline');
  assert.ok(html.includes('07-20'), 'line chart should render labels');
}

// RED: 任务表时间列应统一格式化
{
  const rt = buildRuntime(['escHtml', 'escAttr', 'formatDisplayTime', 'renderTaskTable']);
  rt.els['#task-table-wrap'] = { innerHTML: '' };
  rt.renderTaskTable([
    {
      id: 1,
      run_id: 'r1',
      script_title: 'ruoyi',
      started_at: '2026-07-26T16:26:37',
      ended_at: '2026-07-26T16:30:00',
      duration_seconds: 12,
      success_count: 1,
      fail_count: 0,
      no_graph_count: 0,
      avg_success_duration_seconds: 12,
      status: 'ok',
    }
  ]);
  const html = rt.els['#task-table-wrap'].innerHTML;
  assert.ok(html.includes('2026-07-26 16:26:37'));
  assert.ok(html.includes('2026-07-26 16:30:00'));
}

console.log('test_webui_task_trends OK');
