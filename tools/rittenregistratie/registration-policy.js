(() => {
  'use strict';

  const DEFAULT_FULL_FROM = '2027-01-01';
  const DEFAULT_PRIVATE_LIMIT = 500;
  let fullRegistrationFrom = DEFAULT_FULL_FROM;
  let privateKmLimit = DEFAULT_PRIVATE_LIMIT;

  const form = document.getElementById('rideForm');
  const rideDate = document.getElementById('rideDate');
  const rideType = document.getElementById('rideType');
  const vehicleSelect = document.getElementById('vehicleSelect');
  const startOdometer = document.getElementById('startOdometer');
  const endOdometer = document.getElementById('endOdometer');
  const departureAddress = document.getElementById('departureAddress');
  const arrivalAddress = document.getElementById('arrivalAddress');
  const notes = document.getElementById('notes');
  const formMessage = document.getElementById('formMessage');
  const privateKm = document.getElementById('privateKm');
  const fillPrevious = document.getElementById('fillFromPrevious');

  if (!form || !rideDate || !rideType) return;

  function isFullRegistration(dateValue = rideDate.value) {
    return String(dateValue || '') >= fullRegistrationFrom;
  }

  function setMessage(text, success = false) {
    if (!formMessage) return;
    formMessage.textContent = text;
    formMessage.classList.toggle('success', success);
  }

  function coordsFrom(input) {
    const lat = Number(input.dataset.latitude);
    const lon = Number(input.dataset.longitude);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    return { lat, lon };
  }

  function syncFormPolicy() {
    const full = isFullRegistration();
    const privateOption = rideType.querySelector('option[value="private"]');
    if (privateOption) {
      privateOption.hidden = !full;
      privateOption.disabled = !full;
    }
    if (!full && rideType.value !== 'business') rideType.value = 'business';
    if (fillPrevious) {
      fillPrevious.hidden = !full;
      fillPrevious.title = full
        ? 'Neem de vorige eindstand over'
        : 'Tot en met 2026 kunnen niet-geregistreerde privéritten tussen zakelijke ritten zitten';
    }
  }

  async function refreshPrivateCounter() {
    if (!privateKm) return;
    const year = new Date().getFullYear();
    if (`${year}-01-01` < fullRegistrationFrom) {
      privateKm.textContent = 'n.v.t.';
      return;
    }
    try {
      const response = await fetch('./api/rides', {
        cache: 'no-store',
        credentials: 'same-origin',
        headers: { Accept: 'application/json' }
      });
      if (!response.ok) return;
      const payload = await response.json();
      const rides = Array.isArray(payload.rides) ? payload.rides : [];
      const total = rides
        .filter((ride) => String(ride.date).startsWith(`${year}-`) && ride.type === 'private')
        .reduce((sum, ride) => sum + Number(ride.distance || 0), 0);
      privateKm.textContent = `${new Intl.NumberFormat('nl-NL').format(total)} / ${privateKmLimit} km`;
    } catch {
      // De normale app toont zijn eigen waarde als deze extra beleidsweergave niet kan laden.
    }
  }

  async function loadPolicy() {
    try {
      const response = await fetch('./api/policy', {
        cache: 'no-store',
        credentials: 'same-origin',
        headers: { Accept: 'application/json' }
      });
      if (!response.ok) return;
      const payload = await response.json();
      if (payload.fullRegistrationFrom) fullRegistrationFrom = String(payload.fullRegistrationFrom);
      if (Number.isFinite(Number(payload.privateKmLimit))) privateKmLimit = Number(payload.privateKmLimit);
    } catch {
      // Veilige defaults blijven actief.
    }
    syncFormPolicy();
    refreshPrivateCounter();
  }

  async function submitRide(event) {
    event.preventDefault();
    event.stopImmediatePropagation();

    if (!form.reportValidity()) return;
    const vehicleId = Number(vehicleSelect && vehicleSelect.value);
    if (!Number.isInteger(vehicleId) || vehicleId <= 0) {
      setMessage('Kies eerst een kenteken.');
      return;
    }

    const start = Number(startOdometer.value);
    const end = Number(endOdometer.value);
    if (!Number.isInteger(start) || !Number.isInteger(end)) {
      setMessage('Gebruik hele kilometers voor de kilometerstanden.');
      return;
    }
    if (start < 0 || end < start) {
      setMessage('De eindkilometerstand kan niet lager zijn dan de beginstand.');
      return;
    }

    const full = isFullRegistration();
    const type = full ? rideType.value : 'business';
    const payload = {
      vehicleId,
      date: rideDate.value,
      type,
      startOdometer: start,
      endOdometer: end,
      departureAddress: departureAddress.value.trim(),
      arrivalAddress: arrivalAddress.value.trim(),
      departureTime: departureAddress.dataset.capturedTime || null,
      arrivalTime: arrivalAddress.dataset.capturedTime || null,
      departureCoords: coordsFrom(departureAddress),
      arrivalCoords: coordsFrom(arrivalAddress),
      notes: notes.value.trim()
    };

    const button = form.querySelector('button[type="submit"]');
    const oldText = button ? button.textContent : '';
    if (button) {
      button.disabled = true;
      button.textContent = 'Opslaan…';
    }
    setMessage('');

    try {
      const response = await fetch('./api/rides', {
        method: 'POST',
        cache: 'no-store',
        credentials: 'same-origin',
        headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      let result = {};
      try { result = await response.json(); }
      catch { throw new Error(`API gaf geen geldige JSON (HTTP ${response.status})`); }
      if (!response.ok) throw new Error(result.error || `API-fout HTTP ${response.status}`);
      setMessage(`Rit opgeslagen: ${result.ride.distance} km.`, true);
      window.setTimeout(() => window.location.reload(), 450);
    } catch (error) {
      setMessage(`Rit niet opgeslagen: ${error.message}`);
      if (button) {
        button.disabled = false;
        button.textContent = oldText;
      }
    }
  }

  rideDate.addEventListener('change', syncFormPolicy);
  rideDate.addEventListener('input', syncFormPolicy);
  if (vehicleSelect) {
    vehicleSelect.addEventListener('change', () => {
      if (!isFullRegistration()) {
        window.setTimeout(() => { startOdometer.value = ''; }, 0);
      }
    });
  }
  form.addEventListener('reset', () => window.setTimeout(syncFormPolicy, 0));
  form.addEventListener('submit', submitRide, true);

  syncFormPolicy();
  loadPolicy();
  window.addEventListener('load', () => {
    window.setTimeout(() => {
      syncFormPolicy();
      refreshPrivateCounter();
    }, 700);
  });
})();
