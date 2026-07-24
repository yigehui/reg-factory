const assert = require('assert');

const {
  isLogNearBottom,
  appendLogLineWithAutoScroll,
} = require('./webui/static/log_autoscroll.js');

function makeLogBox({ text = '', scrollTop = 0, clientHeight = 100, scrollHeight = 100 }) {
  return { textContent: text, scrollTop, clientHeight, scrollHeight };
}

function recalcScrollHeight(log) {
  log.scrollHeight = Math.max(log.clientHeight, log.textContent.length);
  return log;
}

// RED: user reading old logs should not be forced to bottom
{
  const log = recalcScrollHeight(
    makeLogBox({ text: 'x'.repeat(200), scrollTop: 20, clientHeight: 100, scrollHeight: 200 })
  );
  appendLogLineWithAutoScroll(log, 'new line');
  assert.strictEqual(log.scrollTop, 20, 'manual scroll position should stay unchanged');
}

// RED: if already near bottom, new logs should keep following bottom
{
  const log = recalcScrollHeight(
    makeLogBox({ text: 'x'.repeat(200), scrollTop: 100, clientHeight: 100, scrollHeight: 200 })
  );
  assert.strictEqual(isLogNearBottom(log), true);
  appendLogLineWithAutoScroll(log, 'new line');
  assert.strictEqual(log.scrollTop, log.scrollHeight, 'near-bottom log should keep following tail');
}

console.log('test_webui_log_autoscroll OK');
