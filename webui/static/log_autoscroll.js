(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
    return;
  }
  root.LogAutoScroll = factory();
})(typeof globalThis !== 'undefined' ? globalThis : window, function () {
  function isLogNearBottom(log, threshold = 24) {
    if (!log) return true;
    const scrollTop = Number(log.scrollTop || 0);
    const clientHeight = Number(log.clientHeight || 0);
    const scrollHeight = Number(log.scrollHeight || 0);
    return scrollTop + clientHeight >= scrollHeight - threshold;
  }

  function appendLogLineWithAutoScroll(log, line) {
    if (!log) return;
    const followTail = isLogNearBottom(log);
    log.textContent += `${line}\n`;
    if (followTail) {
      log.scrollTop = log.scrollHeight;
    }
  }

  return {
    isLogNearBottom,
    appendLogLineWithAutoScroll,
  };
});
