/* Shared page navigation; manuscript content and chart runtimes stay independent. */
(() => {
  'use strict';
  const directory = document.querySelector('.page-directory');
  const links = [...document.querySelectorAll('.site-section-nav [data-section]')];
  const targets = links.map(link => document.getElementById(link.dataset.section));
  const toggle = directory?.querySelector('.directory-toggle');
  const closeDirectory = () => {
    directory?.classList.remove('is-open');
    toggle?.setAttribute('aria-expanded', 'false');
  };
  toggle?.addEventListener('click', () => {
    const open = directory.classList.toggle('is-open');
    toggle.setAttribute('aria-expanded', String(open));
  });
  links.forEach(link => link.addEventListener('click', closeDirectory));
  document.addEventListener('click', event => {
    if (directory && !directory.contains(event.target)) closeDirectory();
  });
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && directory?.classList.contains('is-open')) {
      closeDirectory();
      toggle.focus();
    }
  });
  let frame = 0;
  const updateCurrent = () => {
    const opening = document.querySelector('.home-page > .hero, .post > .article-header');
    const inOpening = opening && opening.getBoundingClientRect().bottom > 72;
    directory?.classList.toggle('is-opening', Boolean(inOpening));
    if (inOpening) closeDirectory();
    let selected = 0;
    targets.forEach((target, index) => {
      if (!target) return;
      const scrollMargin = Number.parseFloat(getComputedStyle(target).scrollMarginTop) || 0;
      if (target.getBoundingClientRect().top <= 72 + scrollMargin) selected = index;
    });
    if (scrollY > 0 && innerHeight + scrollY >= document.documentElement.scrollHeight - 4) selected = links.length - 1;
    links.forEach((link, index) => {
      if (index === selected) link.setAttribute('aria-current', 'location');
      else link.removeAttribute('aria-current');
    });
  };
  const schedule = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(updateCurrent); };
  schedule();
  addEventListener('scroll', schedule, { passive: true });
  addEventListener('resize', schedule, { passive: true });
  addEventListener('hashchange', schedule);
  addEventListener('load', schedule);

  const copyButton = document.querySelector('[data-copy-citation]');
  copyButton?.addEventListener('click', async () => {
    const citation = document.querySelector(copyButton.dataset.copyCitation);
    const status = document.getElementById('copy-status');
    try {
      await navigator.clipboard.writeText(citation.textContent);
      copyButton.textContent = 'Copied ✓';
      status.textContent = 'Citation copied.';
    } catch {
      getSelection().selectAllChildren(citation);
      copyButton.textContent = 'Selected';
      status.textContent = 'Citation selected. Press Control+C or Command+C to copy.';
    }
    setTimeout(() => { copyButton.textContent = 'Copy'; }, 2200);
  });
})();
