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

  function rideLabel(ride) {
    const plate = ride.vehiclePlate || 'zonder kenteken';
    return `${ride.date} · ${plate} · ${ride.startOdometer}-${ride.endOdometer} km · ${ride.departureAddress} → ${ride.arrivalAddress}`;
  }

  function populateRideSelect() {
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
    correctionDepartureTime.value = ride.departureTime || '';
    correctionArrivalTime.value = ride.arrivalTime || '';
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

  async function load() {
    try {
      const [ridesPayload, auditPayload] = await Promise.all([api('/rides'), api('/audit')]);
      rides = ridesPayload.rides || [];
      audit = auditPayload.audit || [];
      populateRideSelect();
      renderAudit();
    } catch (error) {
      setMessage(`Correcties konden niet worden geladen: ${error.message}`);
    }
  }

  async function submitCorrection(event) {
    event.preventDefault();
    const id = rideSelect.value;
    if (!id) {
      setMessage('Kies eerst een rit om te corrigeren.');
      return;
    }
    if (!correctionForm.reportValidity()) return;

    saveCorrection.disabled = true;
    const oldText = saveCorrection.textContent;
    saveCorrection.textContent = 'Correctie opslaan…';
    try {
      await api(`/rides/${id}`, {
        method: 'PATCH',
        body: JSON.stringify({
          date: correctionDate.value,
          type: correctionType.value,
          startOdometer: Number(correctionStart.value),
          endOdometer: Number(correctionEnd.value),
          departureTime: correctionDepartureTime.value || null,
          arrivalTime: correctionArrivalTime.value || null,
          departureAddress: correctionDeparture.value.trim(),
          arrivalAddress: correctionArrival.value.trim(),
          notes: correctionNotes.value.trim(),
          reason: correctionReason.value.trim()
        })
      });
      setMessage('Correctie opgeslagen en in de auditlog vastgelegd. Pagina wordt ververst.', true);
      window.setTimeout(() => window.location.reload(), 600);
    } catch (error) {
      setMessage(`Correctie niet opgeslagen: ${error.message}`);
    } finally {
      saveCorrection.disabled = false;
      saveCorrection.textContent = oldText;
    }
  }

  rideSelect.addEventListener('change', fillCorrectionForm);
  correctionForm.addEventListener('submit', submitCorrection);
  load();
})();
