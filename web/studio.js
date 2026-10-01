(() => {
  'use strict';
  const workspace = document.getElementById('workspacePanel');
  const learning = document.getElementById('learningPanel');
  const workspaceTab = document.getElementById('workspaceTab');
  const learnTab = document.getElementById('learnTab');
  const welcome = document.getElementById('studioWelcome');
  let started = false;
  function openWorkspace(animate = false) {
    started = true;
    welcome.hidden = true;
    learning.hidden = true;
    workspace.hidden = false;
    workspaceTab.setAttribute('aria-pressed', 'true');
    learnTab.setAttribute('aria-pressed', 'false');
    if (animate && !matchMedia('(prefers-reduced-motion: reduce)').matches) {
      workspace.animate([{opacity:0, transform:'translateY(16px)'},{opacity:1,transform:'translateY(0)'}], {duration:350,easing:'cubic-bezier(.22,1,.36,1)'});
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
    document.getElementById('welcomeStatus').textContent = 'Loading the real sample while the illustration shows how values transfer…';
    const demo = document.getElementById('transferDemo');
    demo.classList.remove('is-transferring', 'is-previewing');
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    demo.classList.add('is-transferring');
    document.getElementById('btnSample').click();
    const delay = matchMedia('(prefers-reduced-motion: reduce)').matches ? 0 : 2300;
    await new Promise(resolve => setTimeout(resolve, delay));
    // Wait for the actual processor, rather than presenting a fabricated result.
    if (result.classList.contains('hidden')) {
      await new Promise(resolve => {
        const observer = new MutationObserver(() => {
          if (!result.classList.contains('hidden')) { observer.disconnect(); resolve(); }
        });
        observer.observe(result, {attributes:true,attributeFilter:['class']});
        setTimeout(() => { observer.disconnect(); resolve(); }, 12000);
      });
    }
    openWorkspace(true);
    workspaceTab.focus();
    button.disabled = false;
    document.getElementById('welcomeStart').disabled = false;
    sampleRunning = false;
  }
  document.getElementById('welcomeSample').addEventListener('click', playSample);
  function showLearning(show) {
    welcome.hidden = show || started;
    workspace.hidden = show || !started;
    learning.hidden = !show;
    workspaceTab.setAttribute('aria-pressed', String(!show));
    learnTab.setAttribute('aria-pressed', String(show));
    // The preview measures its real width; recalculate after returning from the guide.
    if (!show) window.dispatchEvent(new Event('resize'));
  }
  workspaceTab.addEventListener('click', () => openWorkspace());
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
