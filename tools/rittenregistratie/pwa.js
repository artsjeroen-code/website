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

  function arrangeRideForm() {
    const ridePanel = document.querySelector('.ride-entry-panel');
    const form = document.getElementById('rideForm');
    if (!ridePanel || !form || form.querySelector('.ride-point-grid')) return;

    const eyebrow = ridePanel.querySelector('.eyebrow');
    if (eyebrow) eyebrow.remove();

    const formGrid = form.querySelector('.form-grid');
    const addressGrid = form.querySelector('.address-grid');
    const dateLabel = document.getElementById('rideDate')?.closest('label');
    const typeLabel = document.getElementById('rideType')?.closest('label');
    const startLabel = document.getElementById('startOdometer')?.closest('label');
    const endLabel = document.getElementById('endOdometer')?.closest('label');
    const departureBlock = document.getElementById('departureAddress')?.closest('.address-block');
    const arrivalBlock = document.getElementById('arrivalAddress')?.closest('.address-block');

    if (!formGrid || !addressGrid || !dateLabel || !typeLabel || !startLabel || !endLabel || !departureBlock || !arrivalBlock) return;

    const metaGrid = document.createElement('div');
    metaGrid.className = 'ride-meta-grid';
    metaGrid.append(dateLabel, typeLabel);

    const pointGrid = document.createElement('div');
    pointGrid.className = 'ride-point-grid';

    const departurePoint = document.createElement('div');
    departurePoint.className = 'ride-point';
    departurePoint.append(startLabel, departureBlock);

    const arrivalPoint = document.createElement('div');
    arrivalPoint.className = 'ride-point';
    arrivalPoint.append(endLabel, arrivalBlock);

    pointGrid.append(departurePoint, arrivalPoint);
    formGrid.replaceWith(metaGrid, pointGrid);
    addressGrid.remove();
  }

  function addQuickCapturePanel() {
    const dashboard = document.querySelector('[data-view-panel="dashboard"]');
    const vehiclePanel = dashboard && dashboard.querySelector('.dashboard-vehicle-panel');
    if (!dashboard || !vehiclePanel || document.getElementById('quickStart')) return;

    const panel = document.createElement('section');
    panel.className = 'panel quick-capture-panel';
    panel.innerHTML = `
      <div class="quick-capture-actions">
        <button class="quick-capture-button start" id="quickStart" type="button">Beginpunt registreren</button>
        <button class="quick-capture-button end" id="quickEnd" type="button">Eindpunt registreren</button>
      </div>
      <div id="quickMessage" class="form-message" role="status" aria-live="polite"></div>
      <div class="quick-ride-list" id="quickRideList"></div>
      <p class="empty-state" id="quickRideEmpty">Nog geen snelle registraties.</p>
    `;
    dashboard.insertBefore(panel, vehiclePanel);
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

  function loadRegistrationPolicy() {
    if (document.querySelector('script[data-registration-policy]')) return;
    const script = document.createElement('script');
    script.src = 'registration-policy.js';
    script.defer = true;
    script.dataset.registrationPolicy = 'true';
    document.body.appendChild(script);
  }

  simplifyPageHeader();
  arrangeRideForm();
  addQuickCapturePanel();
  loadQuickCaptureAssets();
  loadRegistrationPolicy();

  if (!('serviceWorker' in navigator)) return;
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./service-worker.js', { scope: './' })
      .catch((error) => console.warn('PWA service worker kon niet worden geregistreerd.', error));
  });
})();
