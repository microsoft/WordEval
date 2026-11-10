
(function () {
  'use strict';
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, attrs, parent, text) {
    var n = document.createElementNS(NS, tag);
    for (var k in attrs) n.setAttribute(k, attrs[k]);
    if (text != null) n.textContent = text;
    if (parent) parent.appendChild(n);
    return n;
  }

  /* ---------- Atomic and Compositional results: dumbbells (shape is a second channel next to colour) ---------- */
  function marker(g, shape, x, y, fill, r) {
    if (shape === 'circle') return el('circle', { cx: x, cy: y, r: r, fill: fill, 'class': 'd-mk' }, g);
    if (shape === 'diamond') return el('path', { d: 'M' + x + ' ' + (y - r - 1.5) + ' L' + (x + r + 1.5) + ' ' + y + ' L' + x + ' ' + (y + r + 1.5) + ' L' + (x - r - 1.5) + ' ' + y + ' Z', fill: fill, 'class': 'd-mk' }, g);
    return el('rect', { x: x - r + .5, y: y - r + .5, width: 2 * r - 1, height: 2 * r - 1, rx: 2, fill: fill, 'class': 'd-mk' }, g);
  }
  function legend(id, series) {
    var ul = document.getElementById(id);
    series.forEach(function (s) {
      var li = document.createElement('li'), sv = document.createElementNS(NS, 'svg');
      sv.setAttribute('viewBox', '0 0 13 13'); marker(sv, s.shape, 6.5, 6.5, s.color, 5);
      li.appendChild(sv); li.appendChild(document.createTextNode(s.name)); ul.appendChild(li);
    });
  }
  function dumbbell(hostId, rows, series, labelIdx, labelSide) {
    var host = document.getElementById(hostId), W = host.clientWidth;
    if (!W) return;
    var narrow = W < 560, nameW = narrow ? 112 : 210, padR = 26, rowH = narrow ? 34 : 30, top = 6;
    var x0 = nameW + 14, x1 = W - padR, sx = function (v) { return x0 + v / 100 * (x1 - x0); };
    // Keep the human reference separate from the original agent ordering.
    var cursor = top;
    var layout = rows.map(function (r) {
      var human = r[0] === 'Human', h = human ? (narrow ? 48 : 42) : rowH;
      var item = { row: r, human: human, height: h, y: cursor + h / 2 };
      cursor += h + (human ? 16 : 0);
      return item;
    });
    var plotBottom = cursor, H = plotBottom + 26;
    host.innerHTML = '';
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, width: W, height: H }, host);
    [0, 25, 50, 75, 100].forEach(function (t) {
      el('line', { x1: sx(t), x2: sx(t), y1: top - 4, y2: plotBottom, 'class': t === 0 ? 'd-base' : 'd-grid' }, svg);
      el('text', { x: sx(t), y: plotBottom + 18, 'text-anchor': 'middle', 'class': 'd-tick' }, svg, String(t));
    });
    layout.forEach(function (item, i) {
      var r = item.row, y = item.y, g = el('g', {role:'img', 'aria-label':r[0] + ' ' + r[1] + ': ' + r[2].map(function(v,k){return series[k].name + ' ' + v.toFixed(1) + '%';}).join(', ')}, svg);
      if (item.human) {
        var dividerY = y + item.height / 2 + 8;
        el('line',{x1:0,x2:W,y1:dividerY,y2:dividerY,'class':'d-grid'},g);
      }
      var details = item.human && narrow && r[1] === 'mean of 3 participants' ? ['mean of 3', 'participants'] : [r[1]];
      var nameY = narrow || item.human ? y - (details.length > 1 ? 10 : 2) : y + 4;
      var t = el('text', { x: nameW, y: nameY, 'text-anchor': 'end', 'class': 'd-name' }, g, item.human ? 'Human baseline' : r[0]);
      if (narrow || item.human) details.forEach(function (detail, j) {
        el('text', { x: nameW, y: nameY + (j + 1) * 13, 'text-anchor': 'end', 'class': 'd-model' }, g, detail);
      });
      else el('tspan', { dx: 6, 'class': 'd-model' }, t, r[1]);
      var vals = r[2], lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
      var lg = el('linearGradient', { id: hostId + '-g' + i, gradientUnits: 'userSpaceOnUse', x1: sx(hi), x2: sx(lo), y1: 0, y2: 0 }, el('defs', {}, g));
      el('stop', { offset: 0, 'stop-color': series[0].color, 'stop-opacity': .45 }, lg);
      el('stop', { offset: 1, 'stop-color': series[series.length - 1].color, 'stop-opacity': .45 }, lg);
      el('line', { x1: sx(lo), x2: sx(hi), y1: y, y2: y, 'class': 'd-span', stroke: 'url(#' + hostId + '-g' + i + ')' }, g);
      vals.forEach(function (v, k) {
        var m = marker(g, series[k].shape, sx(v), y, series[k].color, 6);
        el('title', {}, m, r[0] + ' ' + r[1] + ' · ' + series[k].name + ' ' + v.toFixed(1) + '%');
      });
      labelIdx.forEach(function (k) {
        var v = vals[k], left = labelSide[k] === 'left';
        el('text', { x: sx(v) + (left ? -11 : 11), y: y + 4, 'text-anchor': left ? 'end' : 'start', 'class': 'd-val', fill: series[k].ink }, g, v.toFixed(1));
      });
    });
  }
  var SA = [{ name: 'Atomic', color: '#4f46e5', ink: '#4338ca', shape: 'circle' }, { name: 'Compositional', color: '#a855f7', ink: '#7e22ce', shape: 'diamond' }];
  // build/paper-r3/csv/main_results.csv (case rates)
  var RA = [['Human', 'mean of 3 participants', [85.2, 69.4]], ['Claude Code', 'Opus 5', [87.5, 84.4]], ['Codex', 'GPT-5.6', [67.6, 61.5]], ['Kimi Code', 'K3', [63.4, 53.1]], ['OpenCode', 'Gemini 3.8 Flash', [57.9, 49.0]],
            ['ZCode', 'GLM-5.2', [58.2, 39.6]], ['OpenCode', 'DeepSeek V4 Pro', [54.9, 40.6]], ['Qwen Code', 'Qwen3.8-27B', [50.4, 31.2]], ['Mini Agent', 'MiniMax M3', [47.3, 29.2]]];
  legend('leg-a', SA);
  function drawAll() { dumbbell('dumb-a', RA, SA, [0, 1], ['right', 'left']); }

  /* ---------- Failure analysis ---------- */
  // build/paper-r3/csv/failure_summary.csv and failure_forms.csv
  (function () {
    var FAIL = [[['Claude Code', 'Opus 5'], 59, 0, 24], [['Codex', 'GPT-5.6'], 96, 94, 24], [['Kimi Code', 'K3'], 143, 13, 89], [['OpenCode', 'Gemini 3.8 Flash'], 11, 0, 268], [['ZCode', 'GLM-5.2'], 179, 49, 58], [['OpenCode', 'DeepSeek V4 Pro'], 250, 44, 9], [['Qwen Code', 'Qwen3.8-27B'], 214, 49, 74], [['Mini Agent', 'MiniMax M3'], 245, 44, 67]];
    var FORMS = [['Property value mismatch', 23], ['Structure mismatch', 20], ['Missing element', 7], ['Off-target change', 4], ['Content mismatch', 4], ['Unrequested change', 1]];
    var N = 642, cats = ['incorrect', 'invalid', 'missing'], names = ['Incorrect output', 'Invalid output', 'Missing output'];
    function pct(v) { return (v * 100).toFixed(1); }
    function systemName(system) { return system[0] + ' (' + system[1] + ')'; }
    var a = '<div class="fa-grid"><div class="fa-panel fa-a"><p class="plabel" style="margin:0"><b>A</b>Output status of non-passing cases</p><ul class="fa-legend">' +
      names.map(function (n, i) { return '<li><i class="' + cats[i] + '"></i>' + n + '</li>'; }).join('') + '</ul><ul class="fa-bars">';
    FAIL.forEach(function (r, i) {
      var tot = r[1] + r[2] + r[3];
      a += '<li' + (i === 0 ? ' class="lead"' : '') + '><span class="fa-name"><span class="fa-system">' + r[0][0] + '</span> <span class="fa-model">' + r[0][1] + '</span></span><span class="fa-track"><span class="fa-stack" style="width:' + (tot / N * 100 / 60 * 100) + '%">';
      [1, 2, 3].forEach(function (k) { if (r[k]) a += '<i class="fa-seg ' + cats[k - 1] + '" style="flex-grow:' + r[k] + '" title="' + systemName(r[0]) + ' · ' + names[k - 1] + ' ' + pct(r[k] / N) + '% (' + r[k] + ' of 642)"></i>'; });
      a += '</span></span><span class="fa-val">' + pct(tot / N) + '%</span></li>';
    });
    a += '</ul><div class="fa-axis"><span></span><div class="fa-ticks">' + [0, 20, 40, 60].map(function (t) { return '<i style="left:' + (t / 60 * 100) + '%">' + t + '</i>'; }).join('') + '</div><span></span></div><p class="fa-axis-title">Share of all cases (%)</p></div>';
    a += '<div class="fa-panel fa-b"><p class="plabel" style="margin:0"><b>B</b>Mismatch types: ' + systemName(FAIL[0][0]) + '</p><ul class="fa-legend"><li><i class="incorrect"></i>Incorrect outputs, one type per case</li></ul><ul class="fa-bars">';
    FORMS.forEach(function (f) {
      a += '<li><span class="fa-name">' + f[0] + '</span><span class="fa-track"><span class="fa-stack" style="width:' + (f[1] / 30 * 100) + '%"><i class="fa-seg incorrect" title="' + f[0] + ' · ' + f[1] + ' of 59"></i></span></span><span class="fa-val">' + f[1] + '</span></li>';
    });
    a += '</ul><div class="fa-axis"><span></span><div class="fa-ticks">' + [0, 10, 20, 30].map(function (t) { return '<i style="left:' + (t / 30 * 100) + '%">' + t + '</i>'; }).join('') + '</div><span></span></div><p class="fa-axis-title">Cases (n = 59)</p></div></div>';
    document.getElementById('outcomes').innerHTML = a;
  })();

  var lastW = 0;
  function onResize() { var w = document.querySelector('.post').clientWidth; if (w !== lastW) { lastW = w; drawAll(); } }
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(drawAll);
  drawAll();
  addEventListener('resize', function () { cancelAnimationFrame(onResize.f); onResize.f = requestAnimationFrame(onResize); });
})();
