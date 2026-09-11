(() => {
  'use strict';

  const API_BASE = './api';
  const rideSelect = document.getElementById('correctionRide');
  const correctionForm = document.getElementById('correctionForm');
  const correctionDate = document.getElementById('correctionDate');
  const correctionType = document.getElementById('correctionType');
  const correctionStart = document.getElementById('correctionStart');
  const correctionEnd = document.getElementById('correctionEnd');
  const correctionDepartureTime = document.getElementById('correctionDepartureTime');
  const correctionArrivalTime = document.getElementById('correctionArrivalTime');
  const correctionDeparture = document.getElementById('correctionDeparture');
  const correctionArrival = document.getElementById('correctionArrival');
  const correctionNotes = document.getElementById('correctionNotes');
  const correctionReason = document.getElementById('correctionReason');
  const correctionMessage = document.getElementById('correctionMessage');
  const auditBody = document.getElementById('auditBody');
  const auditEmpty = document.getElementById('auditEmpty');
  const saveCorrection = document.getElementById('saveCorrection');

  let rides = [];
  let audit = [];

  async function api(path, options = {}) {
    const response = await fetch(`${API_BASE}${path}`, {
      cache: 'no-store',
      headers: {
        Accept: 'application/json',
        ...(options.body ? { 'Content-Type': 'application/json' } : {})
      },
      ...options
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
    return payload;
  }

  function setMessage(text, success = false) {
    correctionMessage.textContent = text;
    correctionMessage.classList.toggle('success', success);
  }

  function showCorrectionToast(ride) {
    document.getElementById('correctionSuccessToast')?.remove();
    const toast = document.createElement('div');
    toast.id = 'correctionSuccessToast';
    toast.setAttribute('role', 'status');
    toast.textContent = `Rit gecorrigeerd${ride.vehiclePlate ? ` · ${ride.vehiclePlate}` : ''}`;
    Object.assign(toast.style, {
      position: 'fixed',
      left: '50%',
      top: '18px',
      transform: 'translateX(-50%)',
      zIndex: '3000',
      padding: '10px 14px',
      borderRadius: '10px',
      background: 'var(--panel)',
      color: 'var(--text)',
      border: '1px solid var(--green, #34a853)',
      boxShadow: '0 8px 24px rgba(0,0,0,.25)',
      fontWeight: '700'
    });
    document.body.appendChild(toast);
    window.setTimeout(() => toast.remove(), 3200);
  }

  function editableTime(value) {
    return value ? String(value).slice(0, 5) : '';
  }

  function normalizeTimeField(input) {
    const raw = input.value.trim();
    if (/^\d{4}$/.test(raw)) input.value = `${raw.slice(0, 2)}:${raw.slice(2)}`;
  }

  function rideLabel(ride) {
    const plate = ride.vehiclePlate || 'zonder kenteken';
    return `${ride.date} · ${plate} · ${ride.startOdometer}-${ride.endOdometer} km · ${ride.departureAddress} → ${ride.arrivalAddress}`;
  }

  function populateRideSelect(selectedRideId = '') {
    rideSelect.replaceChildren();
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = 'Kies een rit…';
    rideSelect.appendChild(placeholder);

    [...rides].reverse().forEach((ride) => {
      const option = document.createElement('option');
      option.value = String(ride.id);
      option.textContent = rideLabel(ride);
      rideSelect.appendChild(option);
    });

    if (selectedRideId && rides.some((ride) => String(ride.id) === String(selectedRideId))) {
      rideSelect.value = String(selectedRideId);
    }
  }

  function fillCorrectionForm() {
    const ride = rides.find((item) => String(item.id) === rideSelect.value);
    if (!ride) {
      correctionForm.reset();
      rideSelect.value = '';
      return;
    }
    correctionDate.value = ride.date;
    correctionType.value = ride.type;
    correctionStart.value = ride.startOdometer;
    correctionEnd.value = ride.endOdometer;
    correctionDepartureTime.value = editableTime(ride.departureTime);
    correctionArrivalTime.value = editableTime(ride.arrivalTime);
    correctionDeparture.value = ride.departureAddress;
    correctionArrival.value = ride.arrivalAddress;
    correctionNotes.value = ride.notes || '';
    correctionReason.value = '';
    setMessage(`Correctie voor ${ride.vehiclePlate || 'dit voertuig'}; de kilometerketen wordt alleen binnen dit voertuig gecontroleerd.`);
  }

  function summaryChange(entry) {
    const oldRide = entry.old;
    const newRide = entry.new;
    const changed = [];
    const labels = {
      date: 'datum', type: 'type', startOdometer: 'beginstand', endOdometer: 'eindstand',
      departureTime: 'vertrektijd', arrivalTime: 'aankomsttijd',
      departureAddress: 'vertrekadres', arrivalAddress: 'aankomstadres', notes: 'toelichting'
    };
    Object.keys(labels).forEach((key) => {
      if (String(oldRide[key] ?? '') !== String(newRide[key] ?? '')) changed.push(labels[key]);
    });
    return changed.length ? changed.join(', ') : 'geen inhoudelijk verschil';
  }

  function renderAudit() {
    auditBody.replaceChildren();
    auditEmpty.hidden = audit.length > 0;
    audit.forEach((entry) => {
      const row = document.createElement('tr');
      const plate = entry.new?.vehiclePlate || entry.old?.vehiclePlate || '—';
      const values = [
        new Date(entry.correctedAt).toLocaleString('nl-NL'),
        `#${entry.rideId} · ${plate}`,
        summaryChange(entry),
        entry.reason
      ];
      values.forEach((value) => {
        const cell = document.createElement('td');
        cell.textContent = value;
        row.appendChild(cell);
      });
      auditBody.appendChild(row);
    });
  }

  async function load(selectedRideId = '') {
    try {
      const [ridesPayload, auditPayload] = await Promise.all([api('/rides'), api('/audit')]);
      rides = ridesPayload.rides || [];
      audit = auditPayload.audit || [];
      populateRideSelect(selectedRideId);
      renderAudit();
      if (selectedRideId) fillCorrectionForm();
    } catch (error) {
      setMessage(`Correcties konden niet worden geladen: ${error.message}`);
    }
  }

  function overviewRowMatches(row, ride) {
    const cells = row.cells;
    if (!cells || cells.length < 8) return false;
    const plate = ride.vehiclePlate || '—';
    const start = String(ride.startOdometer);
    const end = String(ride.endOdometer);
    return cells[1].textContent.trim() === plate
      && cells[5].textContent.trim().replaceAll('.', '') === start
      && cells[6].textContent.trim().replaceAll('.', '') === end
      && cells[3].textContent.includes(ride.departureAddress)
      && cells[4].textContent.includes(ride.arrivalAddress);
  }

  function focusCorrectedRide(ride, attempts = 0) {
    const row = [...document.querySelectorAll('#ridesBody tr')].find((candidate) => overviewRowMatches(candidate, ride));
    if (!row) {
      if (attempts < 20) window.setTimeout(() => focusCorrectedRide(ride, attempts + 1), 100);
      return;
    }

    row.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
    const previousBackground = row.style.backgroundColor;
    const previousTransition = row.style.transition;
    row.style.transition = 'background-color .45s ease';
    row.style.backgroundColor = 'rgba(251, 188, 4, .30)';
    window.setTimeout(() => {
      row.style.backgroundColor = previousBackground;
      window.setTimeout(() => { row.style.transition = previousTransition; }, 500);
    }, 1800);
  }

  function showCorrectedRideInOverview(ride) {
    const yearFilter = document.getElementById('yearFilter');
    const monthFilter = document.getElementById('monthFilter');
    const overviewButton = document.querySelector('.menu-item[data-view="overview"]');
    const rideYear = String(ride.date || '').slice(0, 4);

    if (yearFilter && rideYear) {
      yearFilter.value = rideYear;
      yearFilter.dispatchEvent(new Event('change', { bubbles: true }));
    }
    if (monthFilter) {
      monthFilter.value = 'all';
      monthFilter.dispatchEvent(new Event('change', { bubbles: true }));
    }
    overviewButton?.click();
    showCorrectionToast(ride);
    window.setTimeout(() => focusCorrectedRide(ride), 120);
  }

  async function submitCorrection(event) {
    event.preventDefault();
    const id = rideSelect.value;
    if (!id) {
      setMessage('Kies eerst een rit om te corrigeren.');
      return;
    }

    normalizeTimeField(correctionDepartureTime);
    normalizeTimeField(correctionArrivalTime);
    if (!correctionForm.reportValidity()) {
      setMessage('Gebruik voor tijden het 24-uurs formaat UU:MM, bijvoorbeeld 08:30 of 17:45.');
      return;
    }

    saveCorrection.disabled = true;
    const oldText = saveCorrection.textContent;
    saveCorrection.textContent = 'Correctie opslaan…';
    try {
      const payload = await api(`/rides/${id}`, {
        method: 'PATCH',
        body: JSON.stringify({
          date: correctionDate.value,
          type: correctionType.value,
          startOdometer: Number(correctionStart.value),
          endOdometer: Number(correctionEnd.value),
          departureTime: correctionDepartureTime.value.trim() || null,
          arrivalTime: correctionArrivalTime.value.trim() || null,
          departureAddress: correctionDeparture.value.trim(),
          arrivalAddress: correctionArrival.value.trim(),
          notes: correctionNotes.value.trim(),
          reason: correctionReason.value.trim()
        })
      });

      const correctedRide = payload.ride;
      await load(id);
      correctionReason.value = '';
      setMessage('Rit is gecorrigeerd en de wijziging is in de auditlog vastgelegd.', true);
      document.dispatchEvent(new CustomEvent('rittenregistratie:data-changed'));
      window.setTimeout(() => showCorrectedRideInOverview(correctedRide), 120);
    } catch (error) {
      setMessage(`Correctie niet opgeslagen: ${error.message}`);
    } finally {
      saveCorrection.disabled = false;
      saveCorrection.textContent = oldText;
    }
  }

  [correctionDepartureTime, correctionArrivalTime].forEach((input) => {
    input.addEventListener('blur', () => normalizeTimeField(input));
  });

  document.addEventListener('rittenregistratie:edit-ride', (event) => {
    const rideId = event.detail && event.detail.rideId;
    if (!rideId) return;
    rideSelect.value = String(rideId);
    fillCorrectionForm();
  });

  rideSelect.addEventListener('change', fillCorrectionForm);
  correctionForm.addEventListener('submit', submitCorrection);
  load();
})();
