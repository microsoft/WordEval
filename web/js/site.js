/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* WordEval homepage scripts. Every part is a self-contained IIFE.
   One file, in the order index.html used to load its parts; each part keeps its own header below. */

/* ==================== v2.js ==================== */
/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
(() => {
  'use strict';
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const press = (buttons, active) => buttons.forEach(b => b.setAttribute('aria-pressed', active(b)));

  $('#copy-citation')?.addEventListener('click', async e => {
    const button = e.currentTarget;
    try { await navigator.clipboard.writeText($('#bibtex').textContent); button.textContent = 'Copied ✓'; $('#copy-status').textContent = 'Citation copied.'; }
    catch { getSelection().selectAllChildren($('#bibtex')); button.textContent = 'Selected'; $('#copy-status').textContent = 'Citation selected. Press Control+C or Command+C to copy.'; }
    setTimeout(() => { button.textContent = 'Copy'; }, 2200);
  });

  const D = window.WORDEVAL_RESEARCH;
  if (!D) {
    $$('#sankey, #heatmap, #comp-paper, #outcomes').forEach(el => { el.innerHTML = '<p class="notice">The interactive figure could not be loaded. Reload the page to try again.</p>'; });
    return;
  }
  const short = m => m.name.replace(', high', '');
  // Class definitions follow the paper (Sec. 3, composition classes).
  const classes = {
    additive: { name: 'Additive', mark: 'A + B', description: 'Independent edits. Every precondition already holds, no edit undoes another, and order does not matter.', recipe: 'Bold + italic + highlight' },
    prerequisite: { name: 'Prerequisite-chain', mark: 'A → B', description: 'A later edit needs an object or state that an earlier edit creates.', recipe: 'Insert a chart → add data labels' },
    scope_changing: { name: 'Scope-changing', mark: 'A → {B}', description: 'An earlier edit changes which targets, or what structure, a later edit applies to.', recipe: 'Insert a row → split the table → distribute rows' },
    layout_coupled: { name: 'Layout-coupled', mark: 'A ↔ B', description: 'Later geometry or layout checks depend on wrapping, alignment, or positions set by earlier edits.', recipe: 'Rotate a shape → align top edges' }
  };

  /* ---------- Figure 3: the evaluation loop ---------- */
  // Each stage pairs the general rule with what happened to the Figure 1 task.
  const stages = [
    { title: 'An input document and a request', text: 'Every task gives the agent an input document and a natural-language request. Both come from one editing-unit specification, together with a reference document and the checks, so all four stay aligned.', example: 'Three plain paragraphs, and “Add solid circle bullet points to all paragraphs and set the line spacing to 1.5 lines.”' },
    { title: 'The agent edits with its own tools', text: 'WordEval does not prescribe how to edit. The agent may use any tools it has, and only the document it saves is judged.', example: 'The agent typed a “•” character in front of each paragraph and set 1.5-line spacing.' },
    { title: 'A saved .docx file', text: 'The agent saves the edited document. A missing output, or a document that Word cannot open or that fails OOXML validation, counts as a failure.', example: 'output.docx, which looks like a bulleted list.' },
    { title: 'Word opens the file and reads object properties', text: 'Microsoft Word opens the output, and the oracle reads task-specific properties. It checks the requested change and what must stay the same against the reference document, which the agent never sees.', example: 'ListFormat.ListString: reference U+F0B7, output empty ✕. Line spacing: 1.5 lines in both ✓.' },
    { title: 'Pass only if every check passes', text: 'One failed check fails the task. When a request combines several edits, every edit is checked, and the single-edit tasks for the same edits show whether each passes on its own.', example: 'Spacing ✓, bullets ✕, so the task fails.' }
  ];
  let stage = 3, flowTimer = null;
  function showStage(k) {
    stage = k;
    $$('.flow .stage').forEach(b => b.setAttribute('aria-pressed', Number(b.dataset.stage) === k));
    $$('.flow-arrow').forEach((el, i) => el.classList.toggle('lit', i < k));
    const st = stages[k];
    $('#stage-detail').innerHTML = `<p class="stage-title"><span class="stage-n">${k + 1}</span>${esc(st.title)}</p><p>${esc(st.text)}</p><p class="stage-example"><span class="kicker">In Figure 1</span>${esc(st.example)}</p>`;
  }
  const stopFlow = () => { clearInterval(flowTimer); flowTimer = null; $('#flow-play').textContent = '▶ Play'; };
  $('.flow').addEventListener('click', e => { const b = e.target.closest('.stage'); if (b) { stopFlow(); showStage(Number(b.dataset.stage)); } });
  $('#flow-play').addEventListener('click', () => {
    if (flowTimer) { stopFlow(); return; }
    $('#flow-play').textContent = '❚❚ Pause';
    showStage(0);
    flowTimer = setInterval(() => { if (stage >= stages.length - 1) stopFlow(); else showStage(stage + 1); }, 2200);
  });

  /* ---------- Shared labels and area order ---------- */
  // Word control IDs carry UI-widget words (Gallery, Picker, ...); drop them for display.
  const commandLabel = id => id.replace(/_?3D/g, ' 3-D ').replace(/Ratio(\d+)To(\d+)/, ' Ratio $1:$2').replace(/Nambers/, 'Numbers').replace(/Tempertature/, 'Temperature')
    .replace(/SmartArt/g, '§ ').replace(/([a-z0-9])([A-Z])/g, '$1 $2').replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2').replace(/§/g, 'SmartArt')
    .split(/\s+/).filter(w => w && !['Gallery', 'Picker', 'Word', 'Dialog', 'Alternate', 'Smart', 'Classic'].includes(w)).join(' ');
  const templateLabel = id => commandLabel(id.replace(/^(Additive|Prerequisite|ScopeChanging|LayoutCoupled)/, '').replace(/Composition$/, ''));
  // Shared through window.WORDEVAL_UI with the Sankey (Figure 2), coverage (Figure 4) and explorer (Figure 5) parts.
  // Feature areas in the order of the capability heatmap (mean pass rate across systems, highest first). Every
  // figure that lists the areas uses this order: the Sankey, Figures 4 and 5, and the heatmap itself.
  const areaMean = a => D.capability.filter(c => c.area === a).reduce((sum, c) => sum + c.macro, 0);
  const areasByMean = [...D.areas].sort().sort((a, b) => areaMean(b) - areaMean(a));
  window.WORDEVAL_UI = { classes, commandLabel, templateLabel, esc, areaOrder: areasByMean };


  // Table 1 sorts by any score column (descending first, click again for ascending).
  const tableKeys = ['atomic_case', 'atomic_macro', 'composition_case', 'composition_macro', 'overall_case', 'overall_macro'];
  const tableSort = { key: 'overall_case', dir: -1 };
  function renderTable() {
    const keys = tableKeys;
    const best = Object.fromEntries(keys.map(k => [k, Math.max(...D.models.map(m => m.metrics[k]))]));
    const rows = [...D.models].sort((a, b) => tableSort.dir * (a.metrics[tableSort.key] - b.metrics[tableSort.key]) || D.models.indexOf(a) - D.models.indexOf(b));
    // As in the paper's Table 1, the human mean sits below the agents, outside the sort and the best marks.
    const human = D.human_display ? `<tr class="human-row"><th scope="row">Human<small>mean of 3</small></th>${keys.map(k => `<td>${esc(D.human_display[k])}</td>`).join('')}</tr>` : '';
    $('#main-results').innerHTML = rows.map(m => `<tr><th scope="row">${esc(m.agent)}<small>${esc(m.name)}</small></th>${keys.map(k => `<td class="${[m.metrics[k] === best[k] ? 'best' : '', k === tableSort.key ? 'sorted' : ''].join(' ').trim()}">${m.display_rates[k]}</td>`).join('')}</tr>`).join('') + human;
    const heads = $$('.results-table thead tr:last-child th');
    heads.forEach((th, i) => {
      const k = keys[i];
      if (!th.querySelector('button')) th.innerHTML = `<button type="button" class="sort-btn" data-sort="${k}">${th.textContent}<span class="sort-ind" aria-hidden="true"></span></button>`;
      th.setAttribute('aria-sort', k === tableSort.key ? (tableSort.dir < 0 ? 'descending' : 'ascending') : 'none');
    });
  }
  document.addEventListener('click', e => {
    const b = e.target.closest('.results-table .sort-btn');
    if (!b) return;
    tableSort.dir = tableSort.key === b.dataset.sort ? -tableSort.dir : -1;
    tableSort.key = b.dataset.sort;
    renderTable();
  });

  /* ---------- Figure 6: capability heatmap ---------- */
  // Heatmap columns use the same area names as every other figure (the paper's coverage table names).
  const areaShort = {};
  // Same bins and colors as the paper figure (paper/scripts/style.py): darker red marks a lower pass rate.
  // Colors come from the --heat-k tokens in site.css.
  const heatBin = pctValue => [20, 40, 60, 80, 95, Infinity].findIndex(top => pctValue < top);
  function renderHeatmap() {
    $('#heatmap').innerHTML = `<table><thead><tr><th></th>${areasByMean.map(a => `<th scope="col"><span>${esc(areaShort[a] || a)}</span></th>`).join('')}</tr></thead><tbody>${D.models.map(m => `<tr><th scope="row">${esc(short(m))}</th>${areasByMean.map(a => {
      const cell = D.capability.find(c => c.model === m.id && c.area === a), t = cell.macro;
      return `<td><button type="button" class="heat-cell" data-model="${m.id}" data-heat="${esc(a)}" aria-pressed="false" style="background:var(--heat-${heatBin(t * 100)});color:var(--heat-ink-${heatBin(t * 100)})" aria-label="${esc(`${m.agent} / ${m.name}, ${a}: ${cell.display_macro} percent`)}">${cell.display_macro}</button></td>`;
    }).join('')}</tr>`).join('')}</tbody></table>`;
  }
  $('#heatmap').addEventListener('click', e => {
    const b = e.target.closest('[data-heat]');
    if (!b) return;
    press($$('[data-heat]'), x => x === b);
    const m = D.models.find(x => x.id === b.dataset.model), cell = D.capability.find(c => c.model === m.id && c.area === b.dataset.heat);
    $('#heat-detail').innerHTML = `<strong>${esc(m.agent)} / ${esc(m.name)} · ${esc(cell.area)}</strong>: ${cell.display_macro}% command-macro pass rate · ${cell.passed} of ${cell.n} tasks passed across ${cell.commands} commands.`;
  });


  /* ---------- Figure 7: composition pass rates by class ---------- */
  // RQ2 in the paper's format (composition_pass_rate_by_class): five selected systems, the paper's palette,
  // (A) pass rate over all cases in each class, (B) conditional pass rate where all linked Atomic tasks pass.
  // site.css can retheme the bars with --cpf-0 … --cpf-4; without them the paper colors show.
  const paperSystems = ['claude-code-opus-5', 'codex-gpt-56-sol-high', 'kimi-code-k3-high', 'opencode-gemini-38-flash-high', 'zcode-glm-52-high'];
  const paperPalette = ['#513565', '#78608F', '#9B83B0', '#BCAACB', '#DED5E7'].map(hex => {
    const v = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16));
    return `rgb(${v.map(x => Math.round(.92 * x + .08 * 255)).join(',')})`;
  });
  function renderCompositionPaper() {
    const host = $('#comp-paper');
    if (!host) return;
    const systems = paperSystems.map(id => D.models.find(m => m.id === id));
    const W = 240, H = 156, x0 = 40, x1 = 234, yTop = 12, yBot = 140, y = v => yBot - v / 100 * (yBot - yTop);
    const band = (x1 - x0) / systems.length, bw = band * .65;
    const panel = (cls, row) => {
      let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(`${classes[cls].name}, ${row === 'A' ? 'pass rate' : 'conditional pass rate'} for five systems`)}">`;
      [0, 25, 50, 75, 100].forEach(t => { svg += `<line class="cpf-tick" x1="${x0 - 4}" x2="${x0}" y1="${y(t)}" y2="${y(t)}"/><text class="cpf-ylab" x="${x0 - 7}" y="${y(t) + 3.5}" text-anchor="end">${t}</text>`; });
      systems.forEach((m, i) => {
        const r = D.composition.find(c => c.model === m.id && c.class === cls);
        const n = row === 'A' ? r.n : r.eligible, k = row === 'A' ? r.passed : r.conditional_passed;
        const label = n ? `${row === 'A' ? r.display_rate : r.display_conditional}%` : 'undefined';
        const bx = x0 + band * i + (band - bw) / 2;
        const v = n ? k / n * 100 : 0;
        // Paper rule (Figure 5B): conditional rates over fewer than five eligible cases are omitted; the count shows instead.
        const small = row === 'B' && n > 0 && n < 5;
        svg += `<g class="cpf-bar${small ? ' small' : ''}"><title>${esc(`${short(m)} · ${classes[cls].name}: ${small ? `${k} of ${n} eligible, fewer than five, so no rate is shown` : `${label} (${k} of ${n}${row === 'B' ? ' eligible' : ''})`}`)}</title>`;
        if (small) svg += `<text class="cpf-small" x="${bx + bw / 2}" y="${yBot - 4}" text-anchor="middle">${k}/${n}</text>`;
        else if (n) svg += `<rect x="${bx}" y="${y(v)}" width="${bw}" height="${Math.max(0, yBot - y(v))}" style="fill:var(--cpf-${i}, ${paperPalette[i]})"/>`;
        svg += `<rect class="cpf-hit" x="${x0 + band * i}" y="${yTop}" width="${band}" height="${yBot - yTop}"/></g>`;
      });
      svg += `<path class="cpf-axis" d="M${x0} ${yTop - 4} V${yBot} H${x1}"/></svg>`;
      return `<figure class="cpf-panel"><figcaption>${esc(classes[cls].name)}</figcaption>${svg}</figure>`;
    };
    const rows = [['A', 'Pass rate (%)'], ['B', 'Conditional pass rate (%)']];
    host.innerHTML = rows.map(([row, label]) => `<div class="cpf-row" data-row="${row}"><p class="cpf-rowlabel"><b>${row}</b>${label}</p><div class="cpf-panels">${Object.keys(classes).map(cls => panel(cls, row)).join('')}</div></div>`).join('') +
      `<ul class="cpf-legend">${systems.map((m, i) => `<li><i style="background:var(--cpf-${i}, ${paperPalette[i]})"></i>${esc(short(m))}</li>`).join('')}</ul>`;
  }


  showStage(stage); renderTable(); renderHeatmap(); renderCompositionPaper();
  document.documentElement.dataset.ready = 'true';
})();

/* ==================== v3.js ==================== */
/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* v3 page shell: sticky nav state, scroll reveal, and counters. Figures come from v2.js. */
(() => {
  'use strict';
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;

  // Site navigation is shared with the blog in assets/site-shell.js.
  if (!('IntersectionObserver' in window)) return;

  if (reduce) return;

  document.documentElement.classList.add('reveal-ready');
  const reveal = new IntersectionObserver(entries => entries.forEach(e => {
    if (e.isIntersecting) { e.target.classList.add('in'); reveal.unobserve(e.target); }
  }), { threshold: .12, rootMargin: '0px 0px -5% 0px' });
  $$('[data-reveal]').forEach(el => reveal.observe(el));

  // Count up to the value already written in the page (so the number is correct without JS).
  const count = el => {
    const target = parseFloat(el.dataset.count), decimals = (el.dataset.count.split('.')[1] || '').length, suffix = el.dataset.suffix || '', start = performance.now();
    const tick = now => {
      const p = Math.min(1, (now - start) / 1400);
      el.textContent = (target * (1 - Math.pow(1 - p, 3))).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) + suffix;
      if (p < 1) requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  };
  const counter = new IntersectionObserver(entries => entries.forEach(e => {
    if (e.isIntersecting) { counter.unobserve(e.target); count(e.target); }
  }), { threshold: .4 });
  $$('[data-count]').forEach(el => counter.observe(el));
})();

/* ==================== v4.js ==================== */
/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* v4: the task landscape as a Sankey of feature areas flowing into composition classes.
   The Sankey starts full width; selecting a node narrows it and opens the command list beside it. */
(() => {
  'use strict';
  const D = window.WORDEVAL_RESEARCH, UI = window.WORDEVAL_UI, host = document.getElementById('sankey');
  if (!D || !UI || !host) return;
  const { classes, commandLabel, templateLabel, esc } = UI;
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const NS = 'http://www.w3.org/2000/svg';
  const panel = $('#sankey-panel'), box = $('#dist-detail');

  const classKeys = ['additive', 'prerequisite', 'scope_changing', 'layout_coupled'];
  const atomicIn = a => D.cases.filter(c => c.track === 'atomic' && c.areas[0] === a);
  const areaOf = new Map(D.cases.filter(c => c.track === 'atomic').map(c => [c.id, c.areas[0]]));
  // A link counts the edits (atomic intents) of a class's tasks whose Atomic counterpart lies in an area.
  const weight = new Map();
  D.cases.filter(c => c.track === 'compositional').forEach(c => c.linked.forEach(id => {
    const a = areaOf.get(id);
    if (a) weight.set(`${a}|${c.class}`, (weight.get(`${a}|${c.class}`) || 0) + 1);
  }));
  const links = [...weight].map(([key, v]) => { const [area, cls] = key.split('|'); return { area, cls, v }; });
  const outOf = a => links.filter(l => l.area === a).reduce((s, l) => s + l.v, 0);
  const into = k => links.filter(l => l.cls === k).reduce((s, l) => s + l.v, 0);
  const total = links.reduce((s, l) => s + l.v, 0);
  // Areas in the heatmap's order (UI.areaOrder: mean pass rate across systems), as in every figure that lists them.
  const areas = UI.areaOrder ? [...UI.areaOrder] : [...D.areas];
  const nAtomic = D.cases.filter(c => c.track === 'atomic').length, nComp = D.cases.length - nAtomic;

  let pinned = null, hovered = null, drawn = false;

  function el(tag, attrs, parent, text) {
    const node = document.createElementNS(NS, tag);
    Object.entries(attrs).forEach(([k, v]) => node.setAttribute(k, v));
    if (text !== undefined) node.textContent = text;
    if (parent) parent.appendChild(node);
    return node;
  }

  function draw() {
    const W = host.clientWidth;
    if (!W) return;
    // 'narrow' is a phone-width figure; the opened desktop view (~600px) keeps full labels.
    const narrow = W < 520, top = 46;
    const scrollable = !!host.closest('.sankey-viewport');
    const labelL = narrow ? (scrollable ? 132 : 108) : 150, labelR = narrow ? (scrollable ? 144 : 124) : W < 700 ? 168 : 150, nodeW = 8, k = narrow ? .95 : .88;
    const minBand = narrow ? 19 : 18, gapL = narrow ? 6 : 4, gapR = narrow ? 30 : 22, minRight = narrow ? 19 : 32;
    const x0 = labelL, x1 = W - labelR;
    const L = {}, R = {};
    let y = top;
    areas.forEach(a => { const flow = outOf(a) * k, h = Math.max(flow, minBand); L[a] = { y, h, flow, used: 0 }; y += h + gapL; });
    const H = y - gapL;
    const rightTotal = classKeys.reduce((s, c) => s + Math.max(into(c) * k, minRight), 0) + gapR * (classKeys.length - 1);
    let yr = top + (H - top - rightTotal) / 2;
    classKeys.forEach(c => { const flow = into(c) * k, h = Math.max(flow, minRight); R[c] = { y: yr, h, flow, used: 0 }; yr += h + gapR; });

    host.innerHTML = '';
    const svg = el('svg', { viewBox: `0 0 ${W} ${H}`, width: W, height: H, 'aria-hidden': 'true' }, host);
    const defs = el('defs', {}, svg);
    const grad = el('linearGradient', { id: 'sk-grad', gradientUnits: 'userSpaceOnUse', x1: x0, y1: 0, x2: x1, y2: 0 }, defs);
    el('stop', { offset: '0%', 'stop-color': '#4f7cf0' }, grad);
    el('stop', { offset: '100%', 'stop-color': '#9b7bf2' }, grad);

    // Column headers sit over their own node column, so each label names only its side.
    const heads = el('g', { class: 'sk-heads' }, svg);
    el('text', { x: x0 + nodeW, y: 14, 'text-anchor': 'end', class: 'sk-head atomic' }, heads, 'Atomic');
    el('text', { x: x0 + nodeW, y: 30, 'text-anchor': 'end', class: 'sk-head-sub' }, heads, narrow ? `${nAtomic} tasks` : `${nAtomic} tasks · one command each`);
    el('text', { x: x1, y: 14, class: 'sk-head comp' }, heads, 'Compositional');
    el('text', { x: x1, y: 30, class: 'sk-head-sub' }, heads, narrow ? `${nComp} tasks` : `${nComp} tasks · 2–14 edits each`);
    // On phones the columns are too close for a middle label; the caption carries it.
    if (!narrow) el('text', { x: (x0 + nodeW + x1) / 2, y: 30, 'text-anchor': 'middle', class: 'sk-head-mid' }, heads, `${total} atomic intents`);

    const linkLayer = el('g', { class: 'sk-links' }, svg), sparkLayer = el('g', { class: 'sk-sparks' }, svg), nodeLayer = el('g', { class: 'sk-nodes' }, svg);
    const ordered = [...links].sort((p, q) => classKeys.indexOf(p.cls) - classKeys.indexOf(q.cls) || areas.indexOf(p.area) - areas.indexOf(q.area));
    ordered.forEach(l => {
      const w = l.v * k, a = L[l.area], r = R[l.cls];
      const ya = a.y + (a.h - a.flow) / 2 + a.used + w / 2; a.used += w;
      const yb = r.y + (r.h - r.flow) / 2 + r.used + w / 2; r.used += w;
      const xa = x0 + nodeW, mx = (xa + x1) / 2;
      const d = `M${xa},${ya}C${mx},${ya} ${mx},${yb} ${x1},${yb}`;
      const path = el('path', { d, class: 'sk-link', 'stroke-width': Math.max(w, 1.2), 'data-area': l.area, 'data-cls': l.cls }, linkLayer);
      el('title', {}, path, `${l.area} → ${classes[l.cls].name}: ${l.v} atomic intents`);
      el('path', { d, class: 'sk-spark', 'data-area': l.area, 'data-cls': l.cls }, sparkLayer);
    });

    areas.forEach(a => {
      const n = L[a], g = el('g', { class: 'sk-node atomic', 'data-seg': `area:${a}`, tabindex: '0', role: 'button', 'aria-label': `${a}: ${atomicIn(a).length} Atomic tasks, ${outOf(a)} atomic intents. Show its Word commands.` }, nodeLayer);
      el('rect', { class: 'sk-hit', x: 0, y: n.y - gapL / 2, width: x0 + nodeW, height: n.h + gapL }, g);
      el('rect', { class: 'sk-bar', x: x0, y: n.y, width: nodeW, height: n.h, rx: 2 }, g);
      const t = el('text', { x: x0 - 10, y: n.y + n.h / 2 + 4, 'text-anchor': 'end', class: 'sk-label' }, g, a);
      el('tspan', { class: 'sk-count', dx: 6 }, t, atomicIn(a).length);
    });
    classKeys.forEach(key => {
      const n = R[key], m = classes[key], tasks = D.cases.filter(c => c.class === key).length;
      const g = el('g', { class: 'sk-node comp', 'data-seg': `class:${key}`, tabindex: '0', role: 'button', 'aria-label': `${m.name}: ${tasks} Compositional tasks, ${into(key)} atomic intents. Show its scenarios.` }, nodeLayer);
      el('rect', { class: 'sk-hit', x: x1, y: n.y - 8, width: W - x1, height: n.h + 16 }, g);
      el('rect', { class: 'sk-bar', x: x1, y: n.y, width: nodeW, height: n.h, rx: 2 }, g);
      const cy = n.y + n.h / 2;
      el('text', { x: x1 + nodeW + 12, y: cy - (narrow ? 2 : 4), class: 'sk-mark' }, g, m.mark);
      const t = el('text', { x: x1 + nodeW + 12, y: cy + (narrow ? 13 : 15), class: 'sk-label strong' }, g, m.name);
      // On phones the class count is left to the detail panel so the name fits.
      if (!narrow) el('tspan', { class: 'sk-count', dx: 6 }, t, tasks);
    });
    host.style.height = `${H}px`;
    // Beside the Sankey, the detail scrolls within the figure's height.
    box.style.maxHeight = matchMedia('(max-width: 900px)').matches ? '' : `${H}px`;

    if (!drawn && !reduce) {
      $$('.sk-link', host).forEach((p, i) => {
        const len = p.getTotalLength();
        p.style.strokeDasharray = `${len}`;
        p.style.strokeDashoffset = `${len}`;
        p.style.animation = `sk-grow 1.1s cubic-bezier(.3,.7,.2,1) ${i * 18}ms forwards`;
      });
    }
    drawn = true;
    light();
  }

  // Highlight the flows of the hovered (or pinned) node; with neither, all flows show evenly.
  function light() {
    const key = hovered || pinned;
    host.classList.toggle('focused', !!key);
    if (!key) {
      $$('.sk-link, .sk-spark, .sk-node', host).forEach(n => n.classList.remove('on', 'pinned'));
      return;
    }
    const [kind, id] = key.split(':');
    const on = p => kind === 'area' ? p.dataset.area === id : p.dataset.cls === id;
    $$('.sk-link, .sk-spark', host).forEach(p => p.classList.toggle('on', on(p)));
    const linkedNodes = new Set(links.filter(l => kind === 'area' ? l.area === id : l.cls === id).map(l => kind === 'area' ? `class:${l.cls}` : `area:${l.area}`));
    $$('.sk-node', host).forEach(n => {
      n.classList.toggle('on', n.dataset.seg === key || linkedNodes.has(n.dataset.seg));
      n.classList.toggle('pinned', n.dataset.seg === pinned);
      n.setAttribute('aria-pressed', n.dataset.seg === pinned);
    });
  }

  function detail() {
    const [kind, id] = pinned.split(':');
    const ts = kind === 'area' ? atomicIn(id) : D.cases.filter(c => c.class === id);
    const counts = new Map();
    ts.forEach(t => { const label = kind === 'area' ? commandLabel(t.command) : templateLabel(t.command); counts.set(label, (counts.get(label) || 0) + 1); });
    const chips = `<ul class="cmd-chips">${[...counts].map(([label, n]) => `<li>${esc(label)}<small>${n}</small></li>`).join('')}</ul>`;
    const close = '<button type="button" class="sk-close" aria-label="Close and show the full landscape">×</button>';
    box.dataset.track = kind === 'area' ? 'atomic' : 'comp';
    let html;
    if (kind === 'area') {
      const feeds = links.filter(l => l.area === id).sort((p, q) => q.v - p.v).map(l => `${esc(classes[l.cls].name)} <b>${l.v}</b>`).join(' · ');
      html = `<p class="sk-kind">Atomic · feature area</p><p class="dist-detail-head"><strong>${esc(id)}</strong><span>${ts.length} tasks · ${counts.size} Word commands</span></p><p class="dist-detail-desc">${esc(D.atomic_previews.find(p => p.area === id).description)}</p>${chips}<p class="tl-draws">Feeds ${feeds} atomic intents in Compositional tasks.</p>`;
    } else {
      const m = classes[id];
      const from = links.filter(l => l.cls === id).sort((p, q) => q.v - p.v).map(l => `${esc(l.area)} <b>${l.v}</b>`).join(' · ');
      html = `<p class="sk-kind">Compositional · composition class <span class="class-mark">${esc(m.mark)}</span></p><p class="dist-detail-head"><strong>${esc(m.name)}</strong><span>${ts.length} tasks · ${counts.size} scenarios</span></p><p class="dist-detail-desc">${esc(m.description)} For example: ${esc(m.recipe)}.</p>${chips}<p class="tl-draws">Draws on ${from} atomic intents.</p>`;
    }
    box.innerHTML = `<div class="sk-inner">${close}${html}</div>`;
  }

  const open = n => {
    if (!n) return;
    pinned = n.dataset.seg; hovered = null;
    detail();
    panel.classList.add('open');
    box.setAttribute('aria-hidden', 'false');
    light();
  };
  const closeDetail = () => {
    if (!pinned) return;
    const back = $(`[data-seg="${CSS.escape(pinned)}"]`, host);
    pinned = null;
    panel.classList.remove('open');
    box.setAttribute('aria-hidden', 'true');
    light();
    back?.focus({ preventScroll: true });
  };

  host.addEventListener('pointerover', e => { const n = e.target.closest('[data-seg]'); if (n && n.dataset.seg !== hovered) { hovered = n.dataset.seg; light(); } });
  host.addEventListener('pointerleave', () => { hovered = null; light(); });
  host.addEventListener('click', e => open(e.target.closest('[data-seg]')));
  host.addEventListener('keydown', e => { if ((e.key === 'Enter' || e.key === ' ') && e.target.closest('[data-seg]')) { e.preventDefault(); open(e.target.closest('[data-seg]')); } });
  box.addEventListener('click', e => { if (e.target.closest('.sk-close')) closeDetail(); });
  panel.addEventListener('keydown', e => { if (e.key === 'Escape') closeDetail(); });

  // Draw (and animate the bands in) once the figure reaches the viewport; redraw as its width changes.
  let frame = 0, ready = reduce || !('IntersectionObserver' in window);
  new ResizeObserver(() => { if (!ready) return; cancelAnimationFrame(frame); frame = requestAnimationFrame(draw); }).observe(host);
  if (ready) draw();
  else {
    const io = new IntersectionObserver(entries => { if (entries.some(e => e.isIntersecting)) { io.disconnect(); ready = true; draw(); } }, { rootMargin: '0px 0px -10% 0px' });
    io.observe(host);
  }
})();

/* ==================== v12.js ==================== */
/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* v12: Figure 8, the paper's two-panel failure analysis, drawn in HTML so labels keep their size at any width.
   (A) Output status of non-passing cases, as a share of all cases per system.
   (B) Mismatch types of the incorrect outputs of Opus 5, one type per case.
   Data: web/data/failure-data.js (paper build); names and order: research-data.js (Table 1). */
(() => {
  'use strict';
  const D = window.WORDEVAL_RESEARCH, F = window.WORDEVAL_FAILURES, host = document.getElementById('outcomes');
  if (!D || !F || !host) return;
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const short = m => m.name.replace(', high', '');
  const pct = (k, n) => (k / n * 100).toFixed(1);
  const status = [['incorrect', 'Incorrect output'], ['invalid', 'Invalid output'], ['missing', 'Missing output']];
  const typeName = { property_value_mismatch: 'Property value mismatch', structure_mismatch: 'Structure mismatch', missing_element: 'Missing element', off_target_change: 'Off-target change', content_mismatch: 'Content mismatch', unrequested_change: 'Unrequested change' };
  const lead = F.mismatch.id;
  const leadName = short(D.models.find(m => m.id === lead));
  const axis = (ticks, title) => `<div class="fa-axis" aria-hidden="true"><span></span><span class="fa-ticks">${ticks.map((t, i) => `<i style="left:${i / (ticks.length - 1) * 100}%">${t}</i>`).join('')}</span><span></span></div><p class="fa-axis-title" aria-hidden="true">${title}</p>`;

  // Each row is name | bar | value; gridlines are drawn once behind the bar column (CSS).
  // Panel A: 0–60% of all cases, gridlines every 20 points; the value is the total non-passing share.
  const maxA = 60;
  const rowsA = D.models.map(m => {
    const r = F.systems.find(s => s.id === m.id), total = status.reduce((s, [k]) => s + r[k], 0);
    const said = status.map(([k, label]) => `${label.toLowerCase()} ${r[k]}`).join(', ');
    const segs = status.filter(([k]) => r[k] > 0).map(([k, label]) => `<span class="fa-seg ${k}" style="flex-grow:${r[k]}" title="${esc(`${label} · ${r[k]} cases (${pct(r[k], F.n)}%)`)}"></span>`).join('');
    return `<li class="${m.id === lead ? 'lead' : ''}"><span class="fa-name">${esc(short(m))}</span><span class="fa-track" aria-hidden="true"><span class="fa-stack" style="width:${total / F.n / maxA * 10000}%">${segs}</span></span><span class="fa-val" aria-hidden="true">${pct(total, F.n)}%</span><span class="sr-only">: ${esc(said)} of ${F.n} cases</span></li>`;
  }).join('');

  // Panel B: counts, 0–30, gridlines every 10.
  const maxB = 30;
  const rowsB = F.mismatch.types.map(t => `<li><span class="fa-name">${typeName[t.key]}</span><span class="fa-track" aria-hidden="true"><span class="fa-stack" style="width:${t.n / maxB * 100}%"><span class="fa-seg type"></span></span></span><span class="fa-val">${t.n}</span></li>`).join('');

  host.classList.add('fa');
  host.innerHTML = `<div class="fa-grid">
    <section class="fa-panel fa-a" aria-labelledby="fa-a-label">
      <p class="fa-label" id="fa-a-label"><b>A</b>Output status of non-passing cases</p>
      <ul class="fa-legend" aria-hidden="true">${status.map(([k, label]) => `<li><i class="${k}"></i>${label}</li>`).join('')}</ul>
      <ul class="fa-bars">${rowsA}</ul>
      ${axis([0, 20, 40, 60], 'Share of all cases (%)')}
    </section>
    <section class="fa-panel fa-b" aria-labelledby="fa-b-label">
      <p class="fa-label" id="fa-b-label"><b>B</b>Mismatch types of ${esc(leadName)} incorrect outputs</p>
      <ul class="fa-legend" aria-hidden="true"><li><i class="incorrect"></i>Incorrect outputs, one type per case</li></ul>
      <ul class="fa-bars">${rowsB}</ul>
      ${axis([0, 10, 20, 30], `Cases (n = ${F.mismatch.n})`)}
    </section>
  </div>`;
})();

/* ==================== stage.js ==================== */
/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* Figure 1, the stage: the agent's output.docx, the task's reference.docx and the request go into WordEval; what
   comes back is the property readout of each file and the comparison, then the verdict. Tabs switch between real
   cases. The other figures are drawn by the v2, v4 and v12 parts. */
(function () {
  'use strict';
  const REDUCED = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; };
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];

  /* ---------------- stage cases ----------------
     Every case is a real run from the paper's evaluation (or, for the first, from the oracle-validation cohort).
     WordEval opens the agent's output.docx and the task's reference.docx in Word, reads the properties the request
     names from each, and compares the two readouts. `rows` are those readouts (key, output value, reference value,
     match); they come from archive/web/data/case-gallery.js `stored`, archive/web/data/hero-case.json `word_state`,
     or the .docx XML. */
  const assetRoot = new URL('../../', document.currentScript.src);
  const CASES = [
    {
      id: 'bullets', tab: 'Bullets and spacing', track: 'Compositional', cls: 'comp', agent: 'validation cohort',
      request: 'Add solid circle bullet points to all paragraphs and set the line spacing to 1.5 lines.',
      output: 'web/media/cases/case-compositional-01455c0e70ff/output.png', reference: 'web/media/cases/case-compositional-01455c0e70ff/reference.png',
      rows: [
        { key: 'ListType', out: 'wdListNoNumbering ×3', ref: 'wdListBullet ×3', ok: false },
        { key: 'ListString', out: '', ref: 'U+F0B7 ×3', ok: false },
        { key: 'LineSpacingRule', out: '1.5 lines ×3', ref: '1.5 lines ×3', ok: true },
        { key: 'LineSpacing', out: 18, ref: 18, ok: true }
      ],
      pass: false,
      caption: 'Both documents show a bullet before each paragraph. Word reads no list in the output: each “•” is a typed character, so ListString is empty. The reference stores a bullet list. Line spacing matches, the list does not, so the task fails.'
    },
    {
      id: 'rotate-align', tab: 'Rotate, then align', track: 'Compositional', cls: 'comp', agent: 'Claude Code · Opus 5',
      request: 'Rotate the middle rectangle 90° clockwise, then align all three rectangles by their top edges.',
      output: 'web/media/gallery/case-compositional-57100b52fdbb/output.png', reference: 'web/media/gallery/case-compositional-57100b52fdbb/reference.png',
      rows: [
        { key: 'Rotation', out: [0, 90, 0], ref: [0, 90, 0], ok: true },
        { key: 'Size', out: '60 × 120 pt ×3', ref: '60 × 120 pt ×3', ok: true },
        { key: 'Top', out: [-22, -22, -22], ref: [-22, -52, -22], ok: false },
        { key: 'VisibleTop', out: [-22, 8, -22], ref: [-22, -22, -22], ok: false }
      ],
      pass: false,
      caption: 'Two edits in one request. Rotation and sizes match. The top edges do not: a 60 × 120 pt rectangle turned 90° is 60 pt tall on the page, so its visible top edge sits 30 pt lower than the stored frame top says. The same agent passes each edit when asked for it alone.'
    },
    {
      id: 'rotate', tab: 'Rotate a shape', track: 'Atomic', cls: 'atomic', agent: 'Claude Code · Opus 5',
      request: 'Rotate the teardrop shape 90° clockwise.',
      output: 'web/media/cases/case-atomic-e8a20bbf072c/output.png', reference: 'web/media/cases/case-atomic-e8a20bbf072c/reference.png',
      rows: [
        { key: 'Rotation', out: 90, ref: 90, ok: true },
        { key: 'HorizontalFlip', out: false, ref: false, ok: true },
        { key: 'VerticalFlip', out: false, ref: false, ok: true }
      ],
      pass: true,
      caption: 'One Word command in one request. The rotation and flip properties Word reads from the two files agree, and the task passes. This is the single-edit test linked to the combined request.'
    },
    {
      id: 'footnotes', tab: 'Footnote numbering', track: 'Atomic', cls: 'atomic', agent: 'Claude Code · Opus 5',
      request: 'Change the footnote numbering format from numbers (1, 2, 3) to lowercase letters (a, b, c).',
      output: 'web/media/gallery/case-atomic-e55055512896/output.png', reference: 'web/media/gallery/case-atomic-e55055512896/reference.png',
      rows: [
        { key: 'SectionNumberStyle', out: null, ref: 'lowerLetter', ok: false },
        { key: 'DocumentNumberStyle', out: 'lowerLetter', ref: 'lowerLetter', ok: true },
        { key: 'FootnoteNumbers', out: '1, 2, 3', ref: 'a, b, c', ok: false }
      ],
      pass: false,
      caption: 'The agent saved the a, b, c format in the document-wide settings only. The section’s own setting stays unset, so Word still numbers the footnotes 1, 2, 3. WordEval reads the number style Word applies, not the setting the agent wrote.'
    },
    {
      id: 'table', tab: 'Fit, distribute, align', track: 'Compositional', cls: 'comp', agent: 'Kimi Code · K3',
      request: 'Auto-fit the table to its contents, distribute all rows to equal height, and set all cells to middle-center alignment.',
      output: 'web/media/gallery/case-compositional-de7a118f26d6/output.png', reference: 'web/media/gallery/case-compositional-de7a118f26d6/reference.png',
      rows: [
        { key: 'PreferredWidth', out: '100%', ref: 'auto', ok: false },
        { key: 'ColumnWidths', out: [156, 156, 156], ref: [32, 142, 294], ok: false },
        { key: 'RowHeights', out: [72, 72, 72], ref: [72, 72, 72], ok: true },
        { key: 'CellAlignment', out: 'middle, center ×9', ref: 'middle, center ×9', ok: true }
      ],
      pass: false,
      caption: 'Three edits in one request. Row heights and cell alignment match. The table stays at 100% of the page width with three equal columns, so the auto-fit edit did not take effect. Every requested edit is checked, and one missed edit fails the task.'
    },
    {
      id: 'comment', tab: 'Reply to a comment', track: 'Atomic', cls: 'atomic', agent: 'OpenCode · DeepSeek V4 Pro',
      request: "Add a reply 'This is a reply.' to the existing comment on 'TARGET' as user2.",
      output: 'web/media/gallery/case-atomic-ff6e8a3d8b6b/output.png', reference: 'web/media/gallery/case-atomic-ff6e8a3d8b6b/reference.png',
      rows: [
        { key: 'CommentsOnTarget', out: 1, ref: 2, ok: false },
        { key: 'ReplyByUser2', out: null, ref: 'This is a reply.', ok: false },
        { key: 'SavedIn', out: 'commentsExtensible.xml', ref: 'comments.xml', ok: false }
      ],
      pass: false,
      caption: 'The reply text is on the page. Word finds one comment and no reply: the text was saved as a feedback entry, not as a comment object, and shows up as a paragraph under the line. WordEval counts comment objects in Word rather than searching the text.'
    }
  ];

  /* render one JSON value with syntax colours */
  const jval = v => {
    if (Array.isArray(v)) return `<span class="p">[</span>${v.map(jval).join('<span class="p">, </span>')}<span class="p">]</span>`;
    if (typeof v === 'string') return `<span class="s">"${esc(v)}"</span>`;
    return `<span class="n">${esc(String(v))}</span>`;
  };

  const stage = {
    timers: [], current: -1, auto: true, visible: true, cycleTimer: null,
    els: {
      root: $('#stage'), track: $('#demo-track'), agent: $('#demo-agent'),
      outDoc: $('#demo-output'), outImg: $('#demo-output-img'), refDoc: $('#demo-reference'), refImg: $('#demo-reference-img'),
      bubble: $('#bubble'), prompt: $('#demo-prompt'), node: $('#node'), nodeZone: $('.zone-node'),
      state: $('#demo-state'), jsonOut: $('#json-out-body'), jsonRef: $('#json-ref-body'),
      stamp: $('#demo-stamp'), note: $('#demo-note'), tabs: $('#demo-tabs')
    }
  };
  const later = (fn, ms) => { const t = setTimeout(fn, REDUCED ? 0 : ms); stage.timers.push(t); return t; };
  const reflow = node => void node.offsetWidth;

  function resetStage() {
    stage.timers.forEach(clearTimeout); stage.timers = [];
    if (stage.cycleTimer) { clearTimeout(stage.cycleTimer); stage.cycleTimer = null; }
    const E = stage.els;
    E.node.classList.remove('active'); E.nodeZone.classList.remove('flow-in', 'flow-out');
    E.state.classList.remove('show'); E.stamp.classList.remove('show', 'pass', 'fail');
    E.jsonOut.innerHTML = ''; E.jsonRef.innerHTML = '';
    [E.outDoc, E.refDoc, E.bubble].forEach(n => { n.classList.add('notrans'); n.classList.remove('show', 'sent', 'ghost', 'done'); reflow(n); n.classList.remove('notrans'); });
    E.prompt.textContent = '';
  }

  function runStage(idx) {
    const c = CASES[idx]; if (!c || !stage.els.root) return;
    resetStage();
    stage.current = idx;
    const E = stage.els;
    E.tabs.querySelectorAll('.tab').forEach((t, k) => { t.classList.toggle('active', k === idx); t.classList.toggle('paused', !stage.auto); t.setAttribute('aria-selected', k === idx); });
    E.track.textContent = c.track; E.track.className = `chip ${c.cls}`;
    E.agent.textContent = c.agent;
    E.outImg.src = new URL(c.output, assetRoot).href; E.refImg.src = new URL(c.reference, assetRoot).href;
    E.note.innerHTML = `<b>${esc(c.tab)}.</b> ${esc(c.caption)}`;

    /* the two documents appear, then the request types out */
    later(() => E.outDoc.classList.add('show'), 60);
    later(() => E.refDoc.classList.add('show'), 260);
    const text = c.request, speed = Math.max(7, Math.min(16, 1500 / (text.length / 3)));
    let i = 0; const tType = 600;
    const type = () => { i += 3; E.prompt.textContent = text.slice(0, i); if (i < text.length) later(type, speed); else E.bubble.classList.add('done'); };
    later(type, tType);
    const tTyped = tType + Math.ceil(text.length / 3) * speed;

    /* all three go into WordEval */
    const tMerge = tTyped + 500;
    later(() => E.nodeZone.classList.add('flow-in'), tTyped + 200);
    later(() => { E.outDoc.classList.add('sent'); E.refDoc.classList.add('sent'); E.bubble.classList.add('sent'); }, tMerge);
    later(() => { E.node.classList.add('active'); E.nodeZone.classList.remove('flow-in'); }, tMerge + 700);
    later(() => E.nodeZone.classList.add('flow-out'), tMerge + 1000);

    /* the two readouts come out, one property at a time, then the verdict */
    const tOut = tMerge + 1350;
    later(() => E.state.classList.add('show'), tOut);
    const open = [E.jsonOut, E.jsonRef].map(b => { const l = el('span', 'ln brace p', '{'); b.appendChild(l); return l; });
    const items = c.rows.map((r, k) => {
      const last = k === c.rows.length - 1 ? '' : '<span class="p">,</span>';
      const line = (v, mark) => el('span', `ln ${r.ok ? 'ok' : 'no'}`,
        `  <span class="k">"${esc(r.key)}"</span><span class="p">: </span>${jval(v)}${last}${mark ? `<span class="mark" aria-label="${r.ok ? 'match' : 'differ'}">${r.ok ? '✓' : '✕'}</span>` : ''}`);
      const a = line(r.out, true), b = line(r.ref, false);
      E.jsonOut.appendChild(a); E.jsonRef.appendChild(b); return [a, b];
    });
    const close = [E.jsonOut, E.jsonRef].map(b => { const l = el('span', 'ln brace p', '}'); b.appendChild(l); return l; });
    later(() => open.forEach(l => l.classList.add('show')), tOut + 400);
    items.forEach((pair, k) => later(() => pair.forEach(l => l.classList.add('show')), tOut + 650 + k * 520));
    later(() => close.forEach(l => l.classList.add('show')), tOut + 650 + items.length * 520);
    const tStamp = tOut + 650 + items.length * 520 + 150;
    later(() => {
      E.stamp.className = `stamp ${c.pass ? 'pass' : 'fail'}`;
      E.stamp.innerHTML = `<i aria-hidden="true">${c.pass ? '✓' : '✕'}</i><span>WordEval</span><b>${c.pass ? 'Pass' : 'Fail'}</b>`;
      E.stamp.classList.add('show');
      E.nodeZone.classList.remove('flow-out');
    }, tStamp);

    /* ghosts of the documents come back so the readouts can be compared with the renders */
    later(() => {
      [E.outDoc, E.refDoc, E.bubble].forEach(n => { n.classList.add('notrans'); n.classList.remove('sent'); reflow(n); n.classList.remove('notrans'); n.classList.add('ghost'); });
    }, tOut + 700);

    /* auto-advance */
    const cycle = tStamp + 7000;
    const tab = E.tabs.querySelectorAll('.tab')[idx]; if (tab) tab.style.setProperty('--cycle', `${cycle}ms`);
    if (stage.auto && !REDUCED) stage.cycleTimer = setTimeout(() => {
      if (stage.visible) runStage((idx + 1) % CASES.length);
      else stage.cycleTimer = setTimeout(() => runStage((idx + 1) % CASES.length), 4000);
    }, cycle);
  }

  function renderTabs() {
    const T = stage.els.tabs; if (!T) return;
    CASES.forEach((c, idx) => {
      const b = el('button', `tab ${c.pass ? 'pass' : 'fail'}`); b.type = 'button'; b.setAttribute('role', 'tab');
      b.innerHTML = `<i aria-hidden="true">${c.pass ? '✓' : '✕'}</i>${esc(c.tab)}`;
      b.setAttribute('aria-label', `${c.tab}, ${c.track}, ${c.pass ? 'pass' : 'fail'}`);
      b.addEventListener('click', () => { stage.auto = false; runStage(idx); });
      T.appendChild(b);
    });
  }

  /* ---------------- boot ---------------- */
  renderTabs();
  if (stage.els.root) {
    new IntersectionObserver(es => { stage.visible = es[0].isIntersecting; }, { threshold: 0.2 }).observe(stage.els.root);
    runStage(0);
  }
  document.documentElement.dataset.ready = 'true';
})();

/* ==================== coverage.js ==================== */
/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* Object–property coverage of the oracle on the Atomic tasks: feature area → runtime object → object property, the
   hierarchy of the paper's appendix table (tab:capability-hierarchy). Mounts on every [data-coverage] host:
     data-coverage="sunburst"  rings from feature area to runtime object to object property, a slice's angle being the
                               number of Atomic tasks whose oracle reads that property; selecting an area opens it
     data-coverage="columns"   the table's three columns, one feature area opened at a time
   It also exposes WORDEVAL_COVERAGE_UI.areaBlock(area, caseId) for Figure 5's area view (explorer-tree.js).
   Counts are per Word command, as in the paper table; preservation checks are not counted. Figure 5 fills only the
   properties the shown task's check compares, which for a few variants is fewer (comparedByCase).
   Data: web/data/coverage-data.js (scripts/export_coverage_data.py); command names from web/data/explorer-data.js.
   Motion is limited to opacity and colour fades. */
(() => {
  'use strict';
  const C = window.WORDEVAL_COVERAGE, X = window.WORDEVAL_EXPLORER, UI = window.WORDEVAL_UI;
  if (!C || !UI) return;
  const { esc, commandLabel } = UI;
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const NS = 'http://www.w3.org/2000/svg';
  const plural = (n, w, ws) => `${n} ${n === 1 ? w : ws || `${w}s`}`;
  const reduce = () => matchMedia('(prefers-reduced-motion: reduce)').matches;

  // Feature areas in the heatmap's order (mean pass rate across systems), as in every figure that lists them.
  const areas = (UI.areaOrder || C.areas.map(a => a.name)).map(n => C.areas.find(a => a.name === n)).filter(Boolean);
  const T = C.totals;
  const reads = p => p.cases.length;
  const objReads = o => o.props.reduce((s, p) => s + reads(p), 0);
  const areaReads = a => a.objects.reduce((s, o) => s + objReads(o), 0);
  const nProps = a => a.objects.reduce((s, o) => s + o.props.length, 0);
  const commandOf = new Map(), homeOf = new Map();
  (X ? X.areas : []).forEach(a => a.commands.forEach(cmd => cmd.variants.forEach(v => { commandOf.set(v.case, cmd.id); homeOf.set(v.case, a.name); })));
  const homeTasks = name => ((X && X.areas.find(a => a.name === name)) || {}).tasks;
  const descOf = name => ((X && X.areas.find(a => a.name === name)) || {}).description || '';

  // One colour family, blue to violet in each figure's area order; rings lighten outward.
  const hue = i => 212 + i * (72 / (areas.length - 1));
  const shade = (i, ring) => `hsl(${hue(i).toFixed(1)} ${[58, 54, 50][ring]}% ${[52, 66, 80][ring] + (i % 2) * 3}%)`;
  const idx = name => areas.findIndex(a => a.name === name);

  // "Read by" line for one property: the Word commands of the Atomic tasks whose oracle reads it.
  function readBy(area, obj, p) {
    const counts = new Map();
    p.cases.forEach(c => { const k = commandOf.has(c) ? commandLabel(commandOf.get(c)) : c; counts.set(k, (counts.get(k) || 0) + 1); });
    const list = [...counts].sort((x, y) => y[1] - x[1] || x[0].localeCompare(y[0]));
    return `<code>${esc(obj.name)}.${esc(p.name)}</code> is read by ${plural(reads(p), 'Atomic task')} in ${esc(area.name)}: ${list.map(([k, n]) => `${esc(k)}${n > 1 ? ` <small>×${n}</small>` : ''}`).join(', ')}.`;
  }
  // Tasks whose oracle reads this area's properties but that belong to another feature area (Shape text: 5 Shape geometry tasks).
  function outside(a) {
    const all = new Set(a.objects.flatMap(o => o.props.flatMap(p => p.cases))), from = new Map();
    all.forEach(c => { const h = homeOf.get(c); if (h && h !== a.name) from.set(h, (from.get(h) || 0) + 1); });
    if (!from.size) return '';
    const own = all.size - [...from.values()].reduce((s, n) => s + n, 0);
    return `<p class="cov-note">${plural(all.size, 'Atomic task')} read these properties: ${own} from ${esc(a.name)} and ${[...from].map(([h, n]) => `${n} from ${esc(h)}`).join(', ')}.</p>`;
  }
  const areaHead = a => `<p class="cov-head"><strong>${esc(a.name)}</strong><span>${plural(homeTasks(a.name) ?? a.tasks, 'Atomic task')} · ${plural(a.objects.length, 'runtime object')} · ${plural(nProps(a), 'property', 'properties')}</span></p>`;
  const chips = (o, lit) => o.props.map(p => `<button type="button" class="cov-prop${lit && lit.has(`${o.name}.${p.name}`) ? ' read' : ''}" data-o="${esc(o.name)}" data-p="${esc(p.name)}">${esc(p.name)}<small>${reads(p)}</small></button>`).join('');
  const idleRead = 'Numbers count the Atomic tasks whose oracle reads each property. Point at one to see their Word commands.';

  /* ---------- sunburst ---------- */
  function sunburst(host) {
    host.classList.add('cov', 'cov-sun');
    // Ring and list share the page's area order (the heatmap's), so the list reads as the ring's legend.
    const order = areas, ix = idx;
    host.innerHTML = `<div class="cov-sun-body"><div class="cov-chart"><svg class="cov-svg" role="group" aria-label="Object–property coverage: feature areas, runtime objects, and object properties"></svg>
        <div class="cov-center"></div><div class="cov-tip" role="tooltip"></div></div>
      <div class="cov-panel" aria-live="polite"></div></div>
      <p class="cov-hint">Point at a slice to read it; select a feature area to open its runtime objects and properties.</p>`;
    const svg = $('.cov-svg', host), center = $('.cov-center', host), tip = $('.cov-tip', host), panel = $('.cov-panel', host), chart = $('.cov-chart', host);
    let focus = null, R = 0;

    const arc = (cx, cy, r0, r1, a0, a1) => {
      const p = (r, a) => `${(cx + r * Math.sin(a)).toFixed(2)},${(cy - r * Math.cos(a)).toFixed(2)}`;
      if (a1 - a0 > 2 * Math.PI - 1e-6) return `M${p(r1, 0)}A${r1},${r1} 0 1 1 ${p(r1, Math.PI)}A${r1},${r1} 0 1 1 ${p(r1, 0)}M${p(r0, 0)}A${r0},${r0} 0 1 0 ${p(r0, Math.PI)}A${r0},${r0} 0 1 0 ${p(r0, 0)}Z`;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      return `M${p(r1, a0)}A${r1},${r1} 0 ${large} 1 ${p(r1, a1)}L${p(r0, a1)}A${r0},${r0} 0 ${large} 0 ${p(r0, a0)}Z`;
    };
    const ctx = document.createElement('canvas').getContext('2d');
    const textW = (s, px, weight = 500) => { ctx.font = `${weight} ${px}px Inter, system-ui, sans-serif`; return ctx.measureText(s).width; };
    const el = (tag, attrs, parent, text) => { const n = document.createElementNS(NS, tag); Object.entries(attrs).forEach(([k, v]) => n.setAttribute(k, v)); if (text !== undefined) n.textContent = text; parent.appendChild(n); return n; };

    // A label set radially inside a ring slice when it fits, flipped on the left half so it reads left to right.
    function radial(g, cx, cy, r0, r1, a0, a1, text, px, cls) {
      const mid = (a0 + a1) / 2, room = r1 - r0 - 10, w = textW(text, px);
      if (w > room || (a1 - a0) * (r0 + r1) / 2 < px + 3) return;
      const rm = (r0 + r1) / 2, x = cx + rm * Math.sin(mid), y = cy - rm * Math.cos(mid);
      let deg = mid * 180 / Math.PI - 90;
      if (mid > Math.PI) deg += 180;
      el('text', { x: x.toFixed(1), y: y.toFixed(1), class: cls, 'text-anchor': 'middle', 'dominant-baseline': 'central', transform: `rotate(${deg.toFixed(1)} ${x.toFixed(1)} ${y.toFixed(1)})` }, g, text);
    }

    function draw() {
      const W = chart.clientWidth;
      if (!W) return;
      // Phones: no outer labels (the list below the chart names the areas by colour), so the rings take the full width.
      const all = !focus, narrow = W < 560, labelRoom = all && !narrow ? Math.min(128, W * .21) : 8;
      R = Math.max(110, Math.min(250, W / 2 - labelRoom));
      const H = 2 * R + (all && !narrow ? 36 : 16), cx = W / 2, cy = H / 2;
      host.classList.toggle('narrow', narrow);
      svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('width', W); svg.setAttribute('height', H);
      chart.style.height = `${H}px`;
      // The overview list shows all 15 areas; an opened area's object list scrolls within the chart's height.
      panel.style.maxHeight = matchMedia('(max-width: 900px)').matches || !focus ? '' : `${H}px`;
      svg.innerHTML = '';
      const g = el('g', { class: 'cov-rings' }, svg), labels = el('g', { class: 'cov-labels' }, svg);
      const list = all ? order : [focus];
      const total = list.reduce((s, a) => s + areaReads(a), 0);
      const rr = all ? [.3, .5, .72, 1] : [.34, .34, .62, 1];
      // Reading order. In the full view the first areas run down the left side from 12 o'clock and the rest down the
      // right side, split where the angle is closest to half, so both label columns read top to bottom in the list's order
      // (left column, then right). Inside an area, objects and properties also run top to bottom.
      const TAU = 2 * Math.PI, ang = v => v / total * TAU;
      let split = list.length, best = Infinity, cum = 0;
      if (all) list.forEach((area, k) => { cum += areaReads(area); const d = Math.abs(cum - total / 2); if (d < best) { best = d; split = k + 1; } });
      const run = (start, dir) => { let pos = start; return w => { const r = dir > 0 ? [pos, pos + w] : [pos - w, pos]; pos += dir * w; return r; }; };
      const left = run(TAU, -1), right = run(0, 1);
      const ext = [];
      list.forEach((area, k) => {
        const i = ix(area.name), dir = all && k < split ? -1 : 1;
        const [a0, a1] = all ? (dir < 0 ? left : right)(ang(areaReads(area))) : [0, TAU];
        if (all) {
          el('path', { d: arc(cx, cy, rr[0] * R, rr[1] * R, a0, a1), fill: shade(i, 0), class: 'cov-arc area', 'data-a': area.name, tabindex: '0', role: 'button', 'aria-label': `${area.name}: ${plural(area.objects.length, 'runtime object')}, ${plural(nProps(area), 'object property', 'object properties')}. Open it.` }, g);
          ext.push({ area, mid: (a0 + a1) / 2 });
        }
        const nextObj = run(dir < 0 ? a1 : a0, dir);
        area.objects.forEach(o => {
          const [o0, o1] = nextObj(ang(objReads(o)));
          el('path', { d: arc(cx, cy, rr[1] * R, rr[2] * R, o0, o1), fill: shade(i, 1), class: 'cov-arc obj', 'data-a': area.name, 'data-o': o.name }, g);
          radial(labels, cx, cy, rr[1] * R, rr[2] * R, o0, o1, o.name, all ? 10 : 12, `cov-lbl obj${all ? ' small' : ''}`);
          const nextProp = run(dir < 0 ? o1 : o0, dir);
          o.props.forEach(p => {
            const [p0, p1] = nextProp(ang(reads(p)));
            el('path', { d: arc(cx, cy, rr[2] * R, rr[3] * R, p0, p1), fill: shade(i, 2), class: 'cov-arc prop', 'data-a': area.name, 'data-o': o.name, 'data-p': p.name }, g);
            if (!all) radial(labels, cx, cy, rr[2] * R, rr[3] * R, p0, p1, p.name, 11, 'cov-lbl prop');
          });
        });
      });
      if (all && !narrow) outerLabels(labels, ext, cx, cy);
      centerText();
      light(null);
    }

    // Feature-area names outside the rings, with leader lines; labels on each side are pushed apart to stay legible.
    function outerLabels(g, ext, cx, cy) {
      const gap = 15;
      const items = ext.map(e => ({ ...e, side: Math.sin(e.mid) >= 0 ? 1 : -1, y0: cy - (R + 14) * Math.cos(e.mid) }));
      [1, -1].forEach(side => {
        const s = items.filter(e => e.side === side).sort((p, q) => p.y0 - q.y0);
        s.forEach((e, k) => { e.y = Math.max(e.y0, k ? s[k - 1].y + gap : 8); });
        for (let k = s.length - 1; k >= 0; k--) { const lim = k === s.length - 1 ? 2 * cy - 8 : s[k + 1].y - gap; if (s[k].y > lim) s[k].y = lim; }
      });
      items.forEach(e => {
        const px = cx + (R + 8) * Math.sin(e.mid), py = cy - (R + 8) * Math.cos(e.mid);
        const xEdge = cx + e.side * (Math.abs(px - cx) + 14), tx = xEdge + e.side * 5;
        el('path', { d: `M${(cx + (R + 2) * Math.sin(e.mid)).toFixed(1)},${(cy - (R + 2) * Math.cos(e.mid)).toFixed(1)}L${px.toFixed(1)},${py.toFixed(1)}L${xEdge.toFixed(1)},${e.y.toFixed(1)}`, class: 'cov-lead', 'data-a': e.area.name }, g);
        el('text', { x: tx.toFixed(1), y: e.y.toFixed(1), 'text-anchor': e.side > 0 ? 'start' : 'end', 'dominant-baseline': 'central', class: 'cov-out', 'data-a': e.area.name }, g, e.area.name);
      });
    }

    function centerText() {
      center.innerHTML = focus
        ? `<span class="cov-c-k">Feature area</span><b>${esc(focus.name)}</b><span>${plural(focus.objects.length, 'runtime object')}<br>${plural(nProps(focus), 'property', 'properties')}</span><button type="button" class="cov-back">← All areas</button>`
        : `<b class="cov-c-n">${T.atomic_tasks}</b><span>Atomic tasks</span>`;
      const d = 2 * (focus ? .34 : .3) * R - 10;
      center.style.width = center.style.height = `${d}px`;
    }

    function panelHtml() {
      if (!focus) {
        return `<p class="cov-kind">Atomic tasks · ${T.feature_areas} feature areas</p>
          <p class="cov-head"><strong>Object–property coverage</strong><span>${T.runtime_objects} runtime objects · ${T.object_properties} object properties</span></p>
          <ul class="cov-areas">${order.map((a, i) => `<li><button type="button" class="cov-area" data-a="${esc(a.name)}"><i style="background:${shade(i, 0)}"></i><b>${esc(a.name)}</b><span>${plural(a.objects.length, 'object')} · ${plural(nProps(a), 'property', 'properties')}</span></button></li>`).join('')}</ul>`;
      }
      return `<p class="cov-kind">Feature area<button type="button" class="cov-back link">← All feature areas</button></p>${areaHead(focus)}
        <p class="cov-desc">${esc(descOf(focus.name))}</p>
        <dl class="cov-objs stack">${focus.objects.map(o => `<div class="cov-obj" data-o="${esc(o.name)}"><dt><code>${esc(o.name)}</code></dt><dd>${chips(o)}</dd></div>`).join('')}</dl>
        ${outside(focus)}<p class="cov-read">${idleRead}</p>`;
    }
    const setPanel = () => {
      panel.classList.remove('is-in');
      panel.innerHTML = panelHtml();
      panel.scrollTop = 0;
      requestAnimationFrame(() => panel.classList.add('is-in'));
    };

    // Highlight one area / object / property and the path back to the centre; everything else fades.
    function light(t) {
      host.classList.toggle('focused', !!t);
      $$('.cov-arc', svg).forEach(p => {
        const on = !t || (p.dataset.a === t.a && (!t.o || !p.dataset.o || p.dataset.o === t.o) && (!t.p || !p.dataset.p || p.dataset.p === t.p));
        p.classList.toggle('off', !!t && !on);
      });
      $$('.cov-out, .cov-lead', svg).forEach(n => n.classList.toggle('off', !!t && n.dataset.a !== t.a));
      $$('.cov-area', panel).forEach(b => b.classList.toggle('on', !!t && b.dataset.a === t.a));
      $$('.cov-prop', panel).forEach(b => b.classList.toggle('on', !!t && !!t.o && b.dataset.o === t.o && (!t.p || b.dataset.p === t.p)));
    }
    const target = n => (n && n.dataset.a ? { a: n.dataset.a, o: n.dataset.o, p: n.dataset.p } : null);
    function tipFor(t) {
      const a = C.areas.find(x => x.name === t.a);
      if (t.p) { const o = a.objects.find(x => x.name === t.o), p = o.props.find(x => x.name === t.p); return `<b>${esc(o.name)}.${esc(p.name)}</b><span>${esc(a.name)} · read by ${plural(reads(p), 'Atomic task')}</span>`; }
      if (t.o) { const o = a.objects.find(x => x.name === t.o); return `<b>${esc(o.name)}</b><span>${esc(a.name)} · ${plural(o.props.length, 'property', 'properties')} · read by ${plural(o.tasks, 'Atomic task')}</span>`; }
      return `<b>${esc(a.name)}</b><span>${plural(a.objects.length, 'runtime object')} · ${plural(nProps(a), 'property', 'properties')}</span>`;
    }
    function open(name) {
      focus = name ? C.areas.find(a => a.name === name) : null;
      tip.classList.remove('on');
      chart.classList.add('swap'); panel.classList.remove('is-in');
      setTimeout(() => { draw(); setPanel(); chart.classList.remove('swap'); }, reduce() ? 0 : 200);
    }

    svg.addEventListener('pointermove', e => {
      const t = target(e.target.closest('.cov-arc, .cov-out'));
      if (!t) { tip.classList.remove('on'); light(null); return; }
      light(t);
      tip.innerHTML = tipFor(t);
      const r = chart.getBoundingClientRect();
      let x = e.clientX - r.left + 14;
      const y = e.clientY - r.top + 14;
      if (x + tip.offsetWidth > r.width) x = e.clientX - r.left - tip.offsetWidth - 14;
      tip.style.transform = `translate(${x}px, ${y}px)`;
      tip.classList.add('on');
    });
    svg.addEventListener('pointerleave', () => { tip.classList.remove('on'); light(null); });
    svg.addEventListener('click', e => { const t = target(e.target.closest('.cov-arc, .cov-out')); if (t && !focus) open(t.a); });
    svg.addEventListener('keydown', e => { const t = target(e.target.closest('.cov-arc')); if (t && !focus && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); open(t.a); } });
    host.addEventListener('click', e => {
      if (e.target.closest('.cov-back')) { open(null); return; }
      const b = e.target.closest('.cov-area'); if (b) open(b.dataset.a);
    });
    host.addEventListener('keydown', e => { if (e.key === 'Escape' && focus) open(null); });
    panel.addEventListener('pointerover', e => {
      const b = e.target.closest('.cov-area, .cov-prop');
      if (!b) return;
      if (b.classList.contains('cov-area')) { light({ a: b.dataset.a }); return; }
      const o = focus.objects.find(x => x.name === b.dataset.o), p = o.props.find(x => x.name === b.dataset.p);
      light({ a: focus.name, o: o.name, p: p.name });
      $('.cov-read', panel).innerHTML = readBy(focus, o, p);
    });
    panel.addEventListener('pointerleave', () => { light(null); const r = $('.cov-read', panel); if (r) r.innerHTML = idleRead; });

    let frame = 0;
    new ResizeObserver(() => { cancelAnimationFrame(frame); frame = requestAnimationFrame(draw); }).observe(chart);
    draw(); setPanel();
  }

  /* ---------- three columns: feature area | runtime object | object properties ---------- */
  function columns(host) {
    host.classList.add('cov', 'cov-cols');
    const most = Math.max(...areas.map(nProps));
    host.innerHTML = `<div class="cv-heads" aria-hidden="true"><p>Feature area<small>${T.feature_areas}</small><span>objects · properties</span></p><p>Runtime object<small>${T.runtime_objects}</small></p><p>Object properties<small>${T.object_properties}</small></p></div>
      <div class="cv-body"><ol class="cv-areas">${areas.map((a, i) => `<li><button type="button" class="cv-area" data-a="${esc(a.name)}" aria-pressed="false" style="--sw:${shade(i, 0)};--bar:${(nProps(a) / most * 100).toFixed(1)}%"><b>${esc(a.name)}</b><span aria-label="${plural(a.objects.length, 'runtime object')}, ${plural(nProps(a), 'property', 'properties')}">${a.objects.length} · ${nProps(a)}</span><i aria-hidden="true"></i></button></li>`).join('')}</ol>
        <svg class="cv-wires" aria-hidden="true"></svg><div class="cv-right"></div></div>
      <p class="cov-read cv-read"></p>`;
    const body = $('.cv-body', host), wires = $('.cv-wires', host);
    let right = $('.cv-right', host), cur = null;

    function wire() {
      const btn = $(`.cv-area[data-a="${CSS.escape(cur.name)}"]`, host);
      if (!btn || matchMedia('(max-width: 860px)').matches) { wires.innerHTML = ''; return; }
      const br = body.getBoundingClientRect(), b = btn.getBoundingClientRect();
      const x0 = b.right - br.left, y0 = b.top + b.height / 2 - br.top;
      let s = '';
      $$('.cv-obj', right).forEach(row => {
        const r = $('dt', row).getBoundingClientRect(), x1 = r.left - br.left - 4, y1 = r.top + 11 - br.top, mx = (x0 + x1) / 2;
        s += `<path data-o="${esc(row.dataset.o)}" d="M${x0.toFixed(1)},${y0.toFixed(1)}C${mx.toFixed(1)},${y0.toFixed(1)} ${mx.toFixed(1)},${y1.toFixed(1)} ${x1.toFixed(1)},${y1.toFixed(1)}"/>`;
      });
      wires.setAttribute('width', br.width); wires.setAttribute('height', br.height); wires.innerHTML = s;
    }
    function open(name) {
      cur = C.areas.find(a => a.name === name);
      $$('.cv-area', host).forEach(b => { b.classList.toggle('active', b.dataset.a === name); b.setAttribute('aria-pressed', b.dataset.a === name); });
      const fresh = right.cloneNode(false); right.replaceWith(fresh); right = fresh;
      right.innerHTML = `${areaHead(cur)}<p class="cov-desc">${esc(descOf(cur.name))}</p>
        <dl class="cov-objs">${cur.objects.map(o => `<div class="cov-obj cv-obj" data-o="${esc(o.name)}"><dt><code>${esc(o.name)}</code><small>${plural(o.tasks, 'task')}</small></dt><dd>${chips(o)}</dd></div>`).join('')}</dl>${outside(cur)}`;
      $('.cv-read', host).innerHTML = idleRead;
      wire();
      requestAnimationFrame(() => { right.classList.add('is-in'); wire(); });
    }
    const lit = (o, p) => {
      body.classList.toggle('focus', !!o);
      $$('.cv-wires path', host).forEach(w => w.classList.toggle('lit', w.dataset.o === o));
      $$('.cv-obj', host).forEach(r => r.classList.toggle('lit', r.dataset.o === o));
      $$('.cov-prop', host).forEach(b => b.classList.toggle('on', b.dataset.o === o && (!p || b.dataset.p === p)));
    };
    host.addEventListener('click', e => { const b = e.target.closest('.cv-area'); if (b) open(b.dataset.a); });
    host.addEventListener('pointerover', e => {
      const chip = e.target.closest('.cov-prop'), row = e.target.closest('.cv-obj');
      if (chip) {
        const o = cur.objects.find(x => x.name === chip.dataset.o), p = o.props.find(x => x.name === chip.dataset.p);
        lit(o.name, p.name); $('.cv-read', host).innerHTML = readBy(cur, o, p);
      } else if (row) lit(row.dataset.o);
    });
    body.addEventListener('pointerleave', () => { lit(null); $('.cv-read', host).innerHTML = idleRead; });
    new ResizeObserver(() => cur && wire()).observe(body);
    open(host.dataset.open || 'Shape geometry');
  }

  /* ---------- Figure 5 area view: the area's runtime objects, this task's properties lit ---------- */
  function areaBlock(name, caseId) {
    const a = C.areas.find(x => x.name === name);
    if (!a) return '';
    // Specific to general: first the runtime objects this task is checked on (its properties filled), then the
    // area's other objects, so the row leads back to Figure 4's view of every area.
    const lit = new Set((C.comparedByCase || {})[caseId] || C.byCase[caseId] || []);
    const has = o => o.props.some(p => lit.has(`${o.name}.${p.name}`));
    const mine = a.objects.filter(has), rest = a.objects.filter(o => !has(o));
    const elsewhere = [...lit].filter(k => !a.objects.some(o => o.props.some(p => `${o.name}.${p.name}` === k)));
    const rows = list => `<dl class="cov-objs inline">${list.map(o => `<div class="cov-obj" data-o="${esc(o.name)}"><dt><code>${esc(o.name)}</code></dt><dd>${chips(o, lit)}</dd></div>`).join('')}</dl>`;
    return `<div class="xt-row"><div class="cov-block">
        <p class="cov-sub">This task<span>${plural(lit.size, 'property', 'properties')} checked on ${plural(mine.length, 'runtime object')}${elsewhere.length ? `, ${elsewhere.map(k => `<code>${esc(k)}</code>`).join(', ')} in another area` : ''}</span></p>
        ${rows(mine)}
        ${rest.length ? `<p class="cov-sub rest">Also read in ${esc(name)}<span>${plural(nProps(a), 'property', 'properties')} on ${plural(a.objects.length, 'runtime object')} across the area</span></p>${rows(rest)}` : `<p class="cov-note">These are all the runtime objects the oracle reads in ${esc(name)}: ${plural(nProps(a), 'property', 'properties')} in all.</p>`}
        <p class="cov-read" data-area="${esc(name)}">${idleRead}</p></div></div>`;
  }
  // Figure 5's block lives inside explorer-tree.js's card, so its readout is delegated.
  document.addEventListener('pointerover', e => {
    const chip = e.target.closest('.cov-block .cov-prop'), block = e.target.closest('.cov-block');
    if (!block) return;
    const out = $('.cov-read', block), a = C.areas.find(x => x.name === out.dataset.area);
    if (!chip) { out.innerHTML = idleRead; return; }
    const o = a.objects.find(x => x.name === chip.dataset.o);
    out.innerHTML = readBy(a, o, o.props.find(x => x.name === chip.dataset.p));
  });
  window.WORDEVAL_COVERAGE_UI = { areaBlock, readBy, totals: T };

  $$('[data-coverage]').forEach(host => (host.dataset.coverage === 'columns' ? columns : sunburst)(host));
})();

/* ==================== explorer-tree.js ==================== */
/* Copyright (c) Microsoft Corporation. Licensed under the MIT License. */
/* Figure 5: one task from each part of the benchmark, drawn as an upside-down tree.
   Atomic tasks, by feature area:          feature-area axis → the area's representative Atomic task.
   Compositional tasks, by composition class: feature-area axis → the Atomic tasks linked to the task's atomic
   intents → their wires funnelling into one Compositional task, whose instruction numbers each intent's phrase.
   Terms follow the paper: an atomic intent is one editing unit of a Compositional task, linked to the Atomic task
   that applies the same command.
   Data: web/data/explorer-data.js (areas, commands, previews, scenarios with resolved intents, oracle checks, Atomic
   renders from scripts/prepare_intent_renders.py). Hand-placed here: the phrase of each representative instruction
   that states each intent, in the instruction's own words. */
(() => {
  'use strict';
  const X = window.WORDEVAL_EXPLORER, UI = window.WORDEVAL_UI;
  const host = document.getElementById('explorer-tree');
  if (!X || !UI || !host) return;
  const { classes, commandLabel, templateLabel, esc } = UI;
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const classKeys = ['additive', 'prerequisite', 'scope_changing', 'layout_coupled'];
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
  const narrow = () => matchMedia('(max-width: 860px)').matches;

  // One Compositional task per class. Additive (Lists, Paragraphs) and Prerequisite-chain (Shape geometry, Shape text) draw on
  // two feature areas; the other two are the featured examples with rendered pages.
  const REP = {
    additive: 'case-compositional-85e60bdc76e2',
    prerequisite: 'case-compositional-58c445dc2e3b',
    scope_changing: 'case-compositional-5d6dbee05ed8',
    layout_coupled: 'case-compositional-57100b52fdbb',
  };
  const PHRASES = {
    'case-compositional-58c445dc2e3b': ['draw a text box by dragging from (100, 100) to (300, 180)', 'rotate all text in the text box 90° clockwise'],
    'case-compositional-85e60bdc76e2': ['Add a numbered list (1. 2. 3.) to all paragraphs', 'set double line spacing', 'center all text'],
    'case-compositional-46cde88f71d0': ['Add a title "Chart Title" above the bar chart', 'move the legend to the right side', 'add data labels at the center of each bar', 'change the chart to a purple monochromatic color scheme'],
    'case-compositional-0e67821a76bf': ["insert a clustered column chart with categories 'Q1' and 'Q2', and one series 'Revenue' with values 100 and 150", "add data labels at the 'Outside End' position"],
    'case-compositional-5d6dbee05ed8': ['Insert a new row below the first row', 'a new column to the right of the first column', 'split the table at the inserted row', 'distribute the columns in the upper table', 'distribute the rows in the lower table'],
    'case-compositional-57100b52fdbb': ['Rotate the middle rectangle 90° clockwise', 'align all three rectangles by their top edges'],
  };
  // Feature areas in the heatmap's order (mean pass rate across systems), as in Figures 3 and 5.
  const areas = UI.areaOrder ? UI.areaOrder.map(n => X.areas.find(a => a.name === n)).filter(Boolean) : X.areas;
  const allComp = X.scenarios.flatMap(s => s.cases.map(c => ({ ...c, scenario: s.id, class: s.class })));
  const caseOf = k => allComp.find(c => c.id === REP[k]) || allComp.find(c => c.class === k);
  const atomicInfo = new Map();
  X.areas.forEach(a => a.commands.forEach(cmd => cmd.variants.forEach(v => atomicInfo.set(v.case, { command: cmd.id, key: `${cmd.id}/${v.variant}`, instruction: v.instruction }))));
  const nComp = k => allComp.filter(c => c.class === k).length;
  const nScen = k => X.scenarios.filter(s => s.class === k).length;

  const marked = c => {
    let out = '', pos = 0;
    (PHRASES[c.id] || []).forEach((ph, i) => { const at = c.instruction.indexOf(ph, pos); if (at < 0) return; out += esc(c.instruction.slice(pos, at)) + `<mark class="xt-ph" data-n="${i + 1}">${esc(ph)}<sup>${i + 1}</sup></mark>`; pos = at + ph.length; });
    return out + esc(c.instruction.slice(pos));
  };
  const assetRoot = new URL('../../', document.currentScript.src);
  const pair = (input, reference, alt) => input && reference ? `<div class="xt-pair"><figure><img src="${new URL(input, assetRoot).href}" alt="${esc(alt)}: input" loading="lazy"><figcaption>input</figcaption></figure><span aria-hidden="true">→</span><figure><img src="${new URL(reference, assetRoot).href}" alt="${esc(alt)}: reference" loading="lazy"><figcaption>reference</figcaption></figure></div>` : '';
  const check = id => { const p = (X.checks[id] || {}).primary; return p ? `<p class="xt-check"><b>Oracle check.</b> ${esc(p)}.</p>` : ''; };
  const axis = used => `<div class="xt-axis">${areas.map(a => `<span class="xt-area${used.has(a.name) ? ' used' : ''}" data-area="${esc(a.name)}"><i></i><span>${esc(a.name)}</span></span>`).join('')}</div>`;
  const level = (name, sub, cls = '') => `<p class="xt-lv ${cls}">${name}<small>${sub}</small></p>`;
  // Opt-in (data-oracle-reads on the host): the area view adds the area's runtime objects and properties from coverage.js.
  const cov = host.hasAttribute('data-oracle-reads') && window.WORDEVAL_COVERAGE_UI;

  function areaView(a) {
    const id = a.preview.case, info = atomicInfo.get(id);
    return {
      kind: `<p class="xt-kind">${esc(a.name)} <span>${plural(a.commands.length, 'Word command')} · ${plural(a.tasks, 'Atomic task')}</span></p>`,
      body: `<div class="xt-tree">
        ${level('Feature area', 'groups the Word commands')}${axis(new Set([a.name]))}
        ${level('Atomic task', 'one editing unit, one command')}<div class="xt-row"><div class="xt-node solo" data-area="${esc(a.name)}">
          <p class="xt-h"><span class="xt-n solo">A</span><b>${esc(commandLabel(info.command))}</b><small>${esc(a.name)}</small></p>
          <p class="xt-key">Atomic task<code>${esc(info.key)}</code></p>
          <p class="xt-instr">${esc(info.instruction)}</p>${pair(a.preview.input, a.preview.reference, a.preview.title)}${check(id)}</div></div>
        ${cov ? `${level('Object properties', 'the oracle reads, by runtime object')}${cov.areaBlock(a.name, id)}` : ''}
        <svg class="xt-wires" aria-hidden="true"></svg></div>
        <p class="xt-note">One of the ${plural(a.tasks, 'Atomic task')} in ${esc(a.name)}. Each Atomic task asks for one Word command on a document built for it.</p>`,
    };
  }

  function classView(k) {
    const c = caseOf(k), m = classes[k], n = c.intents.length;
    return {
      kind: `<p class="xt-kind comp"><span class="class-mark">${esc(m.mark)}</span>${esc(m.name)} <span>${plural(nScen(k), 'scenario')} · ${plural(nComp(k), 'Compositional task')} · shown: ${esc(templateLabel(c.scenario))}</span></p>`,
      body: `<div class="xt-tree">
        ${level('Feature area', 'of each intent')}${axis(new Set(c.intents.map(i => i.area)))}
        ${level('Atomic intents', 'one command each')}<div class="xt-row" style="--cols:${n}">${c.intents.map((i, j) => `<div class="xt-node" data-n="${j + 1}" data-area="${esc(i.area)}">
          <p class="xt-h"><span class="xt-n" data-n="${j + 1}">${j + 1}</span><b>${esc(commandLabel(i.command))}</b></p>
          <p class="xt-key">Intent ${j + 1}<code>${esc(i.key)}</code></p>${i.media ? pair(i.media.input, i.media.reference, commandLabel(i.command)) : ''}</div>`).join('')}</div>
        ${level('Compositional task', `${plural(n, 'intent')}, one instruction`, 'comp')}<div class="xt-row"><div class="xt-root">
          <p class="xt-ports" aria-hidden="true">${c.intents.map((_, j) => `<span class="xt-port" data-n="${j + 1}">${j + 1}</span>`).join('')}</p>
          <p class="xt-sent">${marked(c)}</p>
          <div class="xt-root-b">${c.media ? pair(c.media.input, c.media.reference, c.media.title) : ''}
            <div><p class="xt-root-k"><span class="class-mark">${esc(m.mark)}</span>${esc(m.name)} · ${esc(templateLabel(c.scenario))}</p>${check(c.id)}</div></div></div></div>
        <svg class="xt-wires" aria-hidden="true"></svg></div>
        <p class="xt-note">Read from the top. Each atomic intent is one editing unit of the Compositional task, and its number matches its phrase in the instruction. The same command is also released as an Atomic task, whose input and reference pages are shown on the intent. The Compositional task is checked on its own document.</p>`,
    };
  }

  // Wires: area tick → task card; in class view, card → numbered port on the Compositional task.
  function wire(t) {
    const svg = $('.xt-wires', t); if (!svg) return;
    if (narrow()) { svg.innerHTML = ''; return; }
    const tr = t.getBoundingClientRect(), box = r => ({ x: r.left + r.width / 2 - tr.left, top: r.top - tr.top, bottom: r.bottom - tr.top });
    const curve = (a, b, attrs) => { const my = (a.bottom + b.top) / 2; return `<path ${attrs} d="M${a.x},${a.bottom} C${a.x},${my} ${b.x},${my} ${b.x},${b.top}"/>`; };
    let s = '';
    $$('.xt-node', t).forEach(node => {
      const n = node.dataset.n || '', nb = box(node.getBoundingClientRect());
      const tick = $(`.xt-area[data-area="${CSS.escape(node.dataset.area)}"] i`, t);
      if (tick) s += curve(box(tick.getBoundingClientRect()), nb, `class="up" data-n="${n}"`);
      const port = n && $(`.xt-port[data-n="${n}"]`, t);
      if (port) s += curve(nb, box(port.getBoundingClientRect()), `class="down" data-n="${n}"`);
    });
    const solo = $('.xt-node.solo', t), reads = $('.cov-block', t);
    if (solo && reads) s += curve(box(solo.getBoundingClientRect()), box(reads.getBoundingClientRect()), 'class="down" data-n=""');
    svg.setAttribute('width', tr.width); svg.setAttribute('height', tr.height); svg.innerHTML = s;
  }

  const state = host.dataset.openArea ? { kind: 'area', key: host.dataset.openArea } : { kind: 'class', key: 'layout_coupled' };
  host.innerHTML = `<div class="xt">
    <div class="xt-nav">
      <p class="xt-label">Atomic tasks <span>by feature area · one shown per area</span></p><p class="xt-pills">${areas.map(a => `<button type="button" class="xt-pill" data-a="${esc(a.name)}">${esc(a.name)}</button>`).join('')}</p>
      <p class="xt-label comp">Compositional tasks <span>by composition class · one shown per class</span></p><p class="xt-pills">${classKeys.map(k => `<button type="button" class="xt-pill comp" data-k="${k}"><span class="class-mark">${esc(classes[k].mark)}</span>${esc(classes[k].name)}</button>`).join('')}</p>
    </div><div class="xt-card"></div></div>`;
  let card = $('.xt-card', host), ro = null;

  function render() {
    $$('.xt-pill', host).forEach(b => b.classList.toggle('active', state.kind === 'area' ? b.dataset.a === state.key : b.dataset.k === state.key));
    const view = state.kind === 'area' ? areaView(X.areas.find(a => a.name === state.key)) : classView(state.key);
    const fresh = card.cloneNode(false); card.replaceWith(fresh); card = fresh; card.classList.remove('is-in');
    card.innerHTML = view.kind + view.body;
    const t = $('.xt-tree', card), go = () => requestAnimationFrame(() => wire(t));
    $$('img', t).forEach(im => { if (!im.complete) im.addEventListener('load', go, { once: true }); });
    if (ro) ro.disconnect();
    ro = new ResizeObserver(go); ro.observe(t);
    const lit = (n, on) => { $$(`[data-n="${n}"]`, card).forEach(el => el.classList.toggle('lit', on)); t.classList.toggle('focus', on); };
    card.addEventListener('pointerover', e => { const x = e.target.closest('.xt-node[data-n], .xt-ph, .xt-port'); if (x && !x.contains(e.relatedTarget)) lit(x.dataset.n, true); });
    card.addEventListener('pointerout', e => { const x = e.target.closest('.xt-node[data-n], .xt-ph, .xt-port'); if (x && !x.contains(e.relatedTarget)) lit(x.dataset.n, false); });
    wire(t);
    requestAnimationFrame(() => card.classList.add('is-in'));
  }
  $('.xt-nav', host).addEventListener('click', e => {
    const b = e.target.closest('[data-a], [data-k]'); if (!b) return;
    Object.assign(state, b.dataset.a ? { kind: 'area', key: b.dataset.a } : { kind: 'class', key: b.dataset.k });
    render();
  });
  render();
})();
