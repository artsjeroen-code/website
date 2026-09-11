(() => {
  'use strict';

  const ASSET_VERSION = 'v16';
  const versioned = (path) => `${path}?${ASSET_VERSION}`;

  function simplifyPageHeader() {
    const header = document.querySelector('.page-header');
    if (!header) return;
    const eyebrow = header.querySelector('.eyebrow');
    const subtitle = header.querySelector('.subtitle');
    if (eyebrow) eyebrow.remove();
    if (subtitle) subtitle.remove();
  }

  function customizeShell() {
    const favicon = document.querySelector('link[rel="icon"]');
    if (favicon) {
      favicon.href = '../../icons/rittenregistratie.svg';
      favicon.type = 'image/svg+xml';
    }

    const nav = document.querySelector('.app-menu nav');
    const overview = nav?.querySelector('[data-view="overview"]');
    const corrections = nav?.querySelector('[data-view="corrections"]');
    if (nav && overview && corrections) nav.insertBefore(overview, corrections);
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

  function loadMenuLayoutFix() {
    if (document.querySelector('link[data-menu-layout-fix]')) return;
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = versioned('menu-layout-fix.css');
    link.dataset.menuLayoutFix = 'true';
    document.head.appendChild(link);
  }

  function loadQuickCaptureAssets() {
    if (!document.querySelector('link[data-quick-capture]')) {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = versioned('quick-capture.css');
      link.dataset.quickCapture = 'true';
      document.head.appendChild(link);
    }

    if (!document.querySelector('script[data-quick-capture]')) {
      const script = document.createElement('script');
      script.src = versioned('quick-capture.js');
      script.defer = true;
      script.dataset.quickCapture = 'true';
      document.body.appendChild(script);
    }
  }

  function loadAddressPresetAssets() {
    if (!document.querySelector('link[data-address-presets]')) {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = versioned('address-presets.css');
      link.dataset.addressPresets = 'true';
      document.head.appendChild(link);
    }

    if (!document.querySelector('script[data-address-presets]')) {
      const script = document.createElement('script');
      script.src = versioned('address-presets.js');
      script.defer = true;
      script.dataset.addressPresets = 'true';
      document.body.appendChild(script);
    }
  }

  function loadDatePickerAssets() {
    if (!document.querySelector('link[data-date-picker]')) {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = versioned('date-picker.css');
      link.dataset.datePicker = 'true';
      document.head.appendChild(link);
    }

    if (!document.querySelector('script[data-date-picker]')) {
      const script = document.createElement('script');
      script.src = versioned('date-picker.js');
      script.defer = true;
      script.dataset.datePicker = 'true';
      document.body.appendChild(script);
    }
  }

  function loadRegistrationPolicy() {
    if (document.querySelector('script[data-registration-policy]')) return;
    const script = document.createElement('script');
    script.src = versioned('registration-policy.js');
    script.defer = true;
    script.dataset.registrationPolicy = 'true';
    document.body.appendChild(script);
  }

  function addEmployerExport() {
    const csvButton = document.getElementById('exportCsv');
    if (!csvButton || document.getElementById('exportEmployer')) return;

    const exportButton = document.createElement('button');
    exportButton.className = 'secondary-button';
    exportButton.id = 'exportEmployer';
    exportButton.type = 'button';
    exportButton.textContent = 'Download Rittenregistratie';
    csvButton.insertAdjacentElement('afterend', exportButton);

    const script = document.createElement('script');
    script.src = versioned('employer-export.js');
    script.defer = true;
    script.dataset.employerExport = 'true';
    document.body.appendChild(script);
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
      <p class="empty-state" id="quickRideEmpty">Nog geen concept ritten opgeslagen</p>
    `;
    dashboard.insertBefore(panel, vehiclePanel);
  }

  simplifyPageHeader();
  customizeShell();
  arrangeRideForm();
  addQuickCapturePanel();
  loadMenuLayoutFix();
  loadQuickCaptureAssets();
  loadAddressPresetAssets();
  loadDatePickerAssets();
  loadRegistrationPolicy();
  addEmployerExport();

  if (!('serviceWorker' in navigator)) return;
  window.addEventListener('load', () => {
    navigator.serviceWorker.register(versioned('./service-worker.js'), { scope: './' })
      .catch((error) => console.warn('PWA service worker kon niet worden geregistreerd.', error));
  });
})();