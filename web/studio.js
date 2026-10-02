(() => {
  'use strict';
  const workspace = document.getElementById('workspacePanel');
  const learning = document.getElementById('learningPanel');
  const workspaceTab = document.getElementById('workspaceTab');
  const learnTab = document.getElementById('learnTab');
  const welcome = document.getElementById('studioWelcome');
  let started = false;
  const overviewTab = document.getElementById('overviewTab');
  const reducedMotion = matchMedia('(prefers-reduced-motion: reduce)');
  const demo = document.getElementById('transferDemo');
  const preview = document.getElementById('togglePreview');
  preview.addEventListener('click', () => {
    const updated = demo.classList.toggle('is-transferring');
    preview.setAttribute('aria-pressed', String(updated));
    preview.textContent = updated ? 'Reset illustration' : 'Preview value transfer';
  });
  function openOverview() {
    welcome.hidden = false;
    workspace.hidden = true;
    learning.hidden = true;
    overviewTab.setAttribute('aria-pressed', 'true');
    workspaceTab.setAttribute('aria-pressed', 'false');
    learnTab.setAttribute('aria-pressed', 'false');
  }
  overviewTab.addEventListener('click', openOverview);
  function openWorkspace(animate = false) {
    started = true;
    welcome.hidden = true;
    learning.hidden = true;
    workspace.hidden = false;
    overviewTab.setAttribute('aria-pressed', 'false');
    workspaceTab.setAttribute('aria-pressed', 'true');
    learnTab.setAttribute('aria-pressed', 'false');
    if (animate && !reducedMotion.matches && window.Motion) {
      Motion.animate(workspace, {opacity:[.65,1], transform:['translateY(8px)','translateY(0)']}, {type:'spring',bounce:0,duration:.25});
    }
    window.dispatchEvent(new Event('resize'));
  }
  welcome.hidden = started;
  workspace.hidden = !started;
  document.getElementById('welcomeStart').addEventListener('click', () => {
    openWorkspace();
    document.getElementById('dropX').focus();
  });
  let sampleRunning = false;
  async function playSample() {
    if (sampleRunning) return;
    sampleRunning = true;
    workspace.hidden = true;
    learning.hidden = true;
    welcome.hidden = false;
    const button = document.getElementById('welcomeSample');
    button.disabled = true;
    document.getElementById('welcomeStart').disabled = true;
    const status = document.getElementById('welcomeStatus');
    status.textContent = 'Preparing the sample workbook and presentation…';
    welcome.setAttribute('aria-busy', 'true');
    try {
      const complete = new Promise(resolve => document.addEventListener('deck-sample-complete', event => resolve(event.detail), {once:true}));
      document.getElementById('btnSample').click();
      const outcome = await complete;
      if (outcome.success) {
        status.textContent = '';
        openWorkspace(true);
        document.getElementById('reviewTitle').focus({preventScroll:true});
      } else status.textContent = outcome.error || 'The sample could not finish. Try again, or start with your own files.';
    } finally {
      welcome.removeAttribute('aria-busy');
      button.disabled = false;
      document.getElementById('welcomeStart').disabled = false;
      sampleRunning = false;
    }
  }
  document.getElementById('welcomeSample').addEventListener('click', playSample);
  function showLearning(show) {
    welcome.hidden = show || started;
    workspace.hidden = show || !started;
    learning.hidden = !show;
    overviewTab.setAttribute('aria-pressed', 'false');
    workspaceTab.setAttribute('aria-pressed', String(!show));
    learnTab.setAttribute('aria-pressed', String(show));
    // The preview measures its real width; recalculate after returning from the guide.
    if (!show) window.dispatchEvent(new Event('resize'));
  }
  workspaceTab.addEventListener('click', event => openWorkspace(event.detail > 0));
  learnTab.addEventListener('click', () => showLearning(true));
  document.getElementById('sidebarLearn').addEventListener('click', () => {
    showLearning(true);
    learnTab.focus();
  });
  document.getElementById('emptySample').addEventListener('click', () => {
    document.getElementById('btnSample').click();
  });
  document.getElementById('guideSample').addEventListener('click', () => {
    openWorkspace();
    document.getElementById('btnSample').click();
    workspaceTab.focus();
  });
  const result = document.getElementById('result');
  const empty = document.getElementById('studioEmpty');
  const syncEmpty = () => { empty.hidden = !result.classList.contains('hidden'); };
  new MutationObserver(syncEmpty).observe(result, { attributes: true, attributeFilter: ['class'] });
  syncEmpty();
})();
