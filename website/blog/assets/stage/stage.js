/* Figure 1 stage, copied from WordEval/web/js/stage.js (stage part only); image paths point to assets/stage/. */
/* WordEval "stage" page. Vanilla JS, no build step.
   1. Hero stage: the agent's output.docx, the task's reference.docx and the request go into WordEval; what comes
      back is the property readout of each file and the comparison, then the verdict. Tabs switch between real cases.
   2. Figures: Table 1, the capability heatmap, the composition figure, and the failure analysis,
      all read from web/data/research-data.js and web/data/failure-data.js. */
(function () {
  'use strict';
  const assetRoot = new URL('./', document.currentScript.src);
  const REDUCED = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; };
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];

  /* ---------------- stage cases ----------------
     Every case is a real run from the paper's evaluation (or, for the first, from the oracle-validation cohort).
     WordEval opens the agent's output.docx and the task's reference.docx in Word, reads the properties the request
     names from each, and compares the two readouts. `rows` are those readouts (key, output value, reference value,
     match); they come from web/data/case-gallery.js `stored`, web/data/hero-case.json `word_state`, or the .docx XML. */
  const CASES = [
    {
      id: 'bullets', tab: 'Bullets and spacing', track: 'Compositional', cls: 'comp', agent: 'validation cohort',
      request: 'Add solid circle bullet points to all paragraphs and set the line spacing to 1.5 lines.',
      output: 'assets/stage/case-compositional-01455c0e70ff/output.png', reference: 'assets/stage/case-compositional-01455c0e70ff/reference.png',
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
      output: 'assets/stage/case-compositional-57100b52fdbb/output.png', reference: 'assets/stage/case-compositional-57100b52fdbb/reference.png',
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
      output: 'assets/stage/case-atomic-e8a20bbf072c/output.png', reference: 'assets/stage/case-atomic-e8a20bbf072c/reference.png',
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
      output: 'assets/stage/case-atomic-e55055512896/output.png', reference: 'assets/stage/case-atomic-e55055512896/reference.png',
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
      output: 'assets/stage/case-compositional-de7a118f26d6/output.png', reference: 'assets/stage/case-compositional-de7a118f26d6/reference.png',
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
      output: 'assets/stage/case-atomic-ff6e8a3d8b6b/output.png', reference: 'assets/stage/case-atomic-ff6e8a3d8b6b/reference.png',
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
    E.outImg.src = new URL(c.output.replace(/^assets\/stage\//, ''), assetRoot).href;
    E.refImg.src = new URL(c.reference.replace(/^assets\/stage\//, ''), assetRoot).href;
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
})();
