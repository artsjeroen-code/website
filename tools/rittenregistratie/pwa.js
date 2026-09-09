(() => {
  'use strict';

  function simplifyPageHeader() {
    const header = document.querySelector('.page-header');
    if (!header) return;
    const eyebrow = header.querySelector('.eyebrow');
    const subtitle = header.querySelector('.subtitle');
    if (eyebrow) eyebrow.remove();
    if (subtitle) subtitle.remove();
  }

  function addQuickCapturePanel() {
    const dashboard = document.querySelector('[data-view-panel="dashboard"]');
    const ridePanel = dashboard && dashboard.querySelector('.ride-entry-panel');
    if (!dashboard || !ridePanel || document.getElementById('quickStart')) return;

    const panel = document.createElement('section');
    panel.className = 'panel quick-capture-panel';
    panel.innerHTML = `
      <div class="panel-heading">
        <div><h2>Snelle invoer</h2></div>
      </div>
      <div class="quick-capture-actions">
        <button class="quick-capture-button start" id="quickStart" type="button">Beginpunt registreren</button>
        <button class="quick-capture-button end" id="quickEnd" type="button">Eindpunt registreren</button>
      </div>
      <div id="quickMessage" class="form-message" role="status" aria-live="polite"></div>
      <div class="quick-ride-list" id="quickRideList"></div>
      <p class="empty-state" id="quickRideEmpty">Nog geen snelle registraties.</p>
    `;
    dashboard.insertBefore(panel, ridePanel);
  }

  function loadQuickCaptureAssets() {
    if (!document.querySelector('link[data-quick-capture]')) {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = 'quick-capture.css';
      link.dataset.quickCapture = 'true';
      document.head.appendChild(link);
    }

    if (!document.querySelector('script[data-quick-capture]')) {
      const script = document.createElement('script');
      script.src = 'quick-capture.js';
      script.defer = true;
      script.dataset.quickCapture = 'true';
      document.body.appendChild(script);
    }
  }

  simplifyPageHeader();
  addQuickCapturePanel();
  loadQuickCaptureAssets();

  if (!('serviceWorker' in navigator)) return;
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./service-worker.js', { scope: './' })
      .catch((error) => console.warn('PWA service worker kon niet worden geregistreerd.', error));
  });
})();
