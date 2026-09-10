(() => {
  'use strict';

  const API_BASE = './api';
  const startButton = document.getElementById('quickStart');
  const endButton = document.getElementById('quickEnd');
  const message = document.getElementById('quickMessage');
  const list = document.getElementById('quickRideList');
  const empty = document.getElementById('quickRideEmpty');

  if (!startButton || !endButton || !message || !list || !empty) return;

  let quickRides = [];
  let appliedQuickRideId = null;

  const nativeFetch = window.fetch.bind(window);
  window.fetch = async (...args) => {
    const response = await nativeFetch(...args);
    const request = args[0];
    const options = args[1] || {};
    const url = typeof request === 'string' ? request : request?.url || '';
    const method = String(options.method || request?.method || 'GET').toUpperCase();

    if (response.ok && method === 'POST' && url.endsWith('/api/rides') && appliedQuickRideId !== null) {
      const quickRideId = appliedQuickRideId;
      try {
        const archiveResponse = await nativeFetch(`${API_BASE}/quick-rides/${quickRideId}/archive`, {
          method: 'POST',
          cache: 'no-store',
          credentials: 'same-origin',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json'
          },
          body: '{}'
        });
        if (!archiveResponse.ok) throw new Error(`HTTP ${archiveResponse.status}`);
        appliedQuickRideId = null;
        await loadQuickRides();
      } catch (error) {
        setMessage(`Rit is opgeslagen, maar de concept-rit kon niet worden verwijderd: ${error.message}`);
      }
    }

    return response;
  };

  function localCapturedAt(date = new Date()) {
    const parts = [
      date.getFullYear(),
      String(date.getMonth() + 1).padStart(2, '0'),
      String(date.getDate()).padStart(2, '0')
    ];
    const time = [date.getHours(), date.getMinutes(), date.getSeconds()]
      .map((value) => String(value).padStart(2, '0'))
      .join(':');
    return `${parts.join('-')}T${time}`;
  }

  function formatCaptured(value) {
    if (!value) return '—';
    const [date, time = ''] = String(value).split('T');
    const [year, month, day] = date.split('-');
    return `${day}-${month}-${year} ${time.slice(0, 5)}`;
  }

  function setMessage(text, success = false) {
    message.textContent = text;
    message.classList.toggle('success', success);
  }

  async function api(path, options = {}) {
    const response = await fetch(`${API_BASE}${path}`, {
      cache: 'no-store',
      credentials: 'same-origin',
      headers: {
        Accept: 'application/json',
        ...(options.body ? { 'Content-Type': 'application/json' } : {})
      },
      ...options
    });
    let payload = {};
    try { payload = await response.json(); }
    catch { throw new Error(`API gaf geen geldige JSON (HTTP ${response.status})`); }
    if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
    return payload;
  }

  function getLocation() {
    return new Promise((resolve, reject) => {
      if (!navigator.geolocation) {
        reject(new Error('Deze browser ondersteunt geen locatiebepaling.'));
        return;
      }
      navigator.geolocation.getCurrentPosition(
        (position) => resolve({
          lat: Number(position.coords.latitude.toFixed(6)),
          lon: Number(position.coords.longitude.toFixed(6)),
          accuracy: Math.round(position.coords.accuracy || 0)
        }),
        (error) => {
          const messages = {
            1: 'Locatietoegang is geweigerd.',
            2: 'De huidige locatie kon niet worden bepaald.',
            3: 'Locatiebepaling duurde te lang.'
          };
          reject(new Error(messages[error.code] || 'Locatie bepalen is mislukt.'));
        },
        { enableHighAccuracy: true, timeout: 15000, maximumAge: 15000 }
      );
    });
  }

  async function reverseGeocode(coords) {
    const params = new URLSearchParams({
      format: 'jsonv2',
      lat: String(coords.lat),
      lon: String(coords.lon),
      zoom: '18',
      addressdetails: '1',
      'accept-language': 'nl'
    });
    try {
      const response = await fetch(`https://nominatim.openstreetmap.org/reverse?${params}`, {
        headers: { Accept: 'application/json' }
      });
      if (!response.ok) return `GPS ${coords.lat}, ${coords.lon}`;
      const result = await response.json();
      const address = result.address || {};
      const road = address.road || address.pedestrian || address.residential || '';
      const street = [road, address.house_number || ''].filter(Boolean).join(' ');
      const locality = [address.postcode || '', address.city || address.town || address.village || address.municipality || ''].filter(Boolean).join(' ');
      return [street, locality].filter(Boolean).join(', ') || result.display_name || `GPS ${coords.lat}, ${coords.lon}`;
    } catch {
      return `GPS ${coords.lat}, ${coords.lon}`;
    }
  }

  async function capture(kind) {
    const button = kind === 'start' ? startButton : endButton;
    const original = button.textContent;
    startButton.disabled = true;
    endButton.disabled = true;
    button.textContent = 'Locatie bepalen…';
    setMessage('');

    try {
      const capturedAt = localCapturedAt();
      const coords = await getLocation();
      button.textContent = 'Adres zoeken…';
      const address = await reverseGeocode(coords);
      await api(`/quick-rides/${kind}`, {
        method: 'POST',
        body: JSON.stringify({ capturedAt, coords, address })
      });
      setMessage(kind === 'start' ? 'Beginpunt opgeslagen.' : 'Eindpunt opgeslagen. Concept-rit is klaar om later aan te vullen.', true);
      await loadQuickRides();
    } catch (error) {
      setMessage(error.message);
    } finally {
      button.textContent = original;
      updateButtons();
    }
  }

  function applyQuickRide(ride) {
    if (!ride.complete) return;
    const dateInput = document.getElementById('rideDate');
    const departure = document.getElementById('departureAddress');
    const arrival = document.getElementById('arrivalAddress');
    const startOdometer = document.getElementById('startOdometer');
    const endOdometer = document.getElementById('endOdometer');
    const form = document.getElementById('rideForm');
    if (!dateInput || !departure || !arrival || !form) return;

    appliedQuickRideId = ride.id;
    dateInput.value = String(ride.startCapturedAt).slice(0, 10);
    departure.value = ride.startAddress || `GPS ${ride.startCoords.lat}, ${ride.startCoords.lon}`;
    arrival.value = ride.endAddress || `GPS ${ride.endCoords.lat}, ${ride.endCoords.lon}`;
    if (startOdometer && ride.startOdometer !== null && ride.startOdometer !== undefined) {
      startOdometer.value = String(ride.startOdometer);
      startOdometer.dispatchEvent(new Event('input', { bubbles: true }));
    }
    if (endOdometer && ride.endOdometer !== null && ride.endOdometer !== undefined) {
      endOdometer.value = String(ride.endOdometer);
      endOdometer.dispatchEvent(new Event('input', { bubbles: true }));
    }
    departure.dataset.latitude = String(ride.startCoords.lat);
    departure.dataset.longitude = String(ride.startCoords.lon);
    departure.dataset.capturedTime = String(ride.startCapturedAt).slice(11, 19);
    arrival.dataset.latitude = String(ride.endCoords.lat);
    arrival.dataset.longitude = String(ride.endCoords.lon);
    arrival.dataset.capturedTime = String(ride.endCapturedAt).slice(11, 19);

    const departureMeta = document.getElementById('departureMeta');
    const arrivalMeta = document.getElementById('arrivalMeta');
    if (departureMeta) departureMeta.textContent = `Snelle registratie · ${formatCaptured(ride.startCapturedAt)} · GPS ±${Math.round(ride.startCoords.accuracy || 0)} m`;
    if (arrivalMeta) arrivalMeta.textContent = `Snelle registratie · ${formatCaptured(ride.endCapturedAt)} · GPS ±${Math.round(ride.endCoords.accuracy || 0)} m`;

    document.dispatchEvent(new CustomEvent('rittenregistratie:location-filled', { detail: { targetId: 'departureAddress' } }));
    document.dispatchEvent(new CustomEvent('rittenregistratie:location-filled', { detail: { targetId: 'arrivalAddress' } }));
    form.scrollIntoView({ behavior: 'smooth', block: 'start' });
    const hasOdometers = ride.startOdometer !== null && ride.startOdometer !== undefined && ride.endOdometer !== null && ride.endOdometer !== undefined;
    setMessage(
      hasOdometers
        ? 'Locaties, tijden en kilometerstanden zijn overgenomen. Kies nu het kenteken en controleer de rit.'
        : 'Locaties en tijden zijn overgenomen. Kies nu het kenteken en vul de ontbrekende kilometerstanden in.',
      true
    );
  }

  async function abortQuickRide(id) {
    try {
      await api(`/quick-rides/${id}/archive`, { method: 'POST', body: '{}' });
      if (appliedQuickRideId === id) appliedQuickRideId = null;
      await loadQuickRides();
      setMessage('Snelle registratie afgebroken.', true);
    } catch (error) {
      setMessage(`Afbreken mislukt: ${error.message}`);
    }
  }

  function updateButtons() {
    const open = quickRides.find((ride) => !ride.complete);
    startButton.disabled = Boolean(open);
    endButton.disabled = !open;
  }

  function render() {
    list.replaceChildren();
    empty.hidden = quickRides.length > 0;

    quickRides.forEach((ride) => {
      const item = document.createElement('article');
      item.className = `quick-ride-item${ride.complete ? ' complete' : ' open'}`;

      const copy = document.createElement('div');
      copy.className = 'quick-ride-copy';
      const title = document.createElement('strong');
      title.textContent = ride.complete ? `Concept-rit ${formatCaptured(ride.startCapturedAt)}` : `Open beginpunt ${formatCaptured(ride.startCapturedAt)}`;
      const details = document.createElement('span');
      const startKm = ride.startOdometer === null || ride.startOdometer === undefined ? '' : ` · ${ride.startOdometer} km`;
      const endKm = ride.endOdometer === null || ride.endOdometer === undefined ? '' : ` · ${ride.endOdometer} km`;
      details.textContent = ride.complete
        ? `${ride.startAddress || 'Beginlocatie'}${startKm} → ${ride.endAddress || 'Eindlocatie'}${endKm} · aankomst ${formatCaptured(ride.endCapturedAt)}`
        : `${ride.startAddress || 'Beginlocatie opgeslagen'}${startKm} · wacht op eindpunt`;
      copy.append(title, details);

      const actions = document.createElement('div');
      actions.className = 'quick-ride-actions';
      if (ride.complete) {
        const use = document.createElement('button');
        use.type = 'button';
        use.className = 'primary-button';
        use.textContent = 'Overnemen';
        use.addEventListener('click', () => applyQuickRide(ride));
        actions.appendChild(use);
      }
      const abort = document.createElement('button');
      abort.type = 'button';
      abort.className = 'secondary-button';
      abort.textContent = 'Afbreken';
      abort.addEventListener('click', () => abortQuickRide(ride.id));
      actions.appendChild(abort);

      item.append(copy, actions);
      list.appendChild(item);
    });
    updateButtons();
  }

  async function loadQuickRides() {
    try {
      const payload = await api('/quick-rides');
      quickRides = Array.isArray(payload.quickRides) ? payload.quickRides : [];
      render();
    } catch (error) {
      setMessage(`Snelle registraties konden niet worden geladen: ${error.message}`);
      quickRides = [];
      render();
    }
  }

  const resetButton = document.getElementById('resetForm');
  if (resetButton) {
    resetButton.addEventListener('click', () => {
      appliedQuickRideId = null;
    });
  }

  startButton.addEventListener('click', () => capture('start'));
  endButton.addEventListener('click', () => capture('end'));
  loadQuickRides();
})();
