(() => {
  'use strict';

  const API_ROUTE = './api/route';
  const NOMINATIM_SEARCH_URL = 'https://nominatim.openstreetmap.org/search';
  const button = document.getElementById('checkRoute');
  const result = document.getElementById('routeCheckResult');
  const icon = document.getElementById('routeStatusIcon');
  const startInput = document.getElementById('startOdometer');
  const endInput = document.getElementById('endOdometer');
  const departureInput = document.getElementById('departureAddress');
  const arrivalInput = document.getElementById('arrivalAddress');

  if (!button || !result || !icon || !startInput || !endInput || !departureInput || !arrivalInput) return;

  let running = false;
  let lastSignature = '';

  function coords(input) {
    const lat = Number(input.dataset.latitude);
    const lon = Number(input.dataset.longitude);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    return { lat, lon };
  }

  function odometerDistance() {
    const start = Number(startInput.value);
    const end = Number(endInput.value);
    if (!Number.isFinite(start) || !Number.isFinite(end) || end < start) return null;
    return end - start;
  }

  function formatKm(value) {
    return new Intl.NumberFormat('nl-NL', { maximumFractionDigits: 1 }).format(value);
  }

  function show(message, status = 'idle', symbol = '?') {
    result.textContent = message;
    result.dataset.status = status;
    icon.dataset.status = status;
    icon.textContent = symbol;
  }

  function signature() {
    const departure = coords(departureInput);
    const arrival = coords(arrivalInput);
    const drivenKm = odometerDistance();
    if (!departure || !arrival || drivenKm === null) return '';
    return [departure.lat, departure.lon, arrival.lat, arrival.lon, drivenKm].join('|');
  }

  async function forwardGeocode(input, label) {
    const address = input.value.trim();
    if (!address) throw new Error(`Vul eerst het ${label}adres in`);

    const params = new URLSearchParams({
      format: 'jsonv2',
      q: address,
      limit: '1',
      addressdetails: '0',
      countrycodes: 'nl',
      'accept-language': 'nl'
    });

    const response = await fetch(`${NOMINATIM_SEARCH_URL}?${params.toString()}`, {
      method: 'GET',
      cache: 'no-store',
      headers: { Accept: 'application/json' }
    });

    if (!response.ok) throw new Error(`Adres zoeken gaf HTTP ${response.status}`);
    const rows = await response.json();
    if (!Array.isArray(rows) || !rows.length) throw new Error(`${label === 'vertrek' ? 'Vertrek' : 'Aankomst'}adres niet gevonden`);

    const lat = Number(rows[0].lat);
    const lon = Number(rows[0].lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) throw new Error(`Geen geldige coördinaten voor het ${label}adres`);

    input.dataset.latitude = String(lat);
    input.dataset.longitude = String(lon);
    return { lat, lon };
  }

  async function ensureCoordsForManualAddresses() {
    let departure = coords(departureInput);
    let arrival = coords(arrivalInput);

    if (!departure) {
      show('Vertrekadres wordt opgezocht…', 'checking', '…');
      departure = await forwardGeocode(departureInput, 'vertrek');
    }

    if (!arrival) {
      show('Aankomstadres wordt opgezocht…', 'checking', '…');
      arrival = await forwardGeocode(arrivalInput, 'aankomst');
    }

    return { departure, arrival };
  }

  async function checkRoute({ silentIfIncomplete = false, geocodeManual = false } = {}) {
    if (running) return;

    let departure = coords(departureInput);
    let arrival = coords(arrivalInput);
    const drivenKm = odometerDistance();

    if (drivenKm === null) {
      if (!silentIfIncomplete) show('Vul eerst een geldige begin- en eindkilometerstand in.', 'question', '?');
      return;
    }

    running = true;
    const original = button.textContent;
    button.disabled = true;
    button.textContent = 'Bezig…';

    try {
      if ((!departure || !arrival) && geocodeManual) {
        ({ departure, arrival } = await ensureCoordsForManualAddresses());
      }

      if (!departure || !arrival) {
        if (!silentIfIncomplete) show('Gebruik “Gebruik locatie”, een snelkeuze of klik op “Opnieuw” bij handmatig ingevoerde adressen.', 'question', '?');
        return;
      }

      const currentSignature = [departure.lat, departure.lon, arrival.lat, arrival.lon, drivenKm].join('|');
      if (!currentSignature || currentSignature === lastSignature) return;

      show('Normale autoroute wordt berekend…', 'checking', '…');

      const response = await fetch(API_ROUTE, {
        method: 'POST',
        cache: 'no-store',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ departure, arrival })
      });

      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);

      const routeKm = Number(payload.route && payload.route.distanceKm);
      if (!Number.isFinite(routeKm)) throw new Error('Geen geldige routeafstand ontvangen');

      const difference = drivenKm - routeKm;
      const absoluteDifference = Math.abs(difference);
      const greenThreshold = Math.max(1, routeKm * 0.05);
      const largeThreshold = Math.max(3, routeKm * 0.20);
      const direction = difference > 0 ? 'meer' : 'minder';

      if (absoluteDifference <= greenThreshold) {
        show(
          `Teller ${formatKm(drivenKm)} km · route ${formatKm(routeKm)} km · verschil ${formatKm(absoluteDifference)} km.`,
          'ok',
          '✓'
        );
      } else if (absoluteDifference <= largeThreshold) {
        show(
          `Kleine afwijking: teller ${formatKm(drivenKm)} km · route ${formatKm(routeKm)} km · ${formatKm(absoluteDifference)} km ${direction}.`,
          'question',
          '?'
        );
      } else {
        show(
          `Flinke afwijking: teller ${formatKm(drivenKm)} km · route ${formatKm(routeKm)} km · ${formatKm(absoluteDifference)} km ${direction}. Controleer invoer of noteer de afwijkende route.`,
          'warning',
          '!'
        );
      }
      lastSignature = currentSignature;
    } catch (error) {
      console.warn('Routecontrole mislukt.', error);
      show(`Routecontrole tijdelijk niet beschikbaar: ${error.message}.`, 'question', '?');
    } finally {
      running = false;
      button.disabled = false;
      button.textContent = original;
    }
  }

  function resetAndMaybeCheck() {
    lastSignature = '';
    show('Wordt automatisch uitgevoerd zodra beide locaties en kilometerstanden bekend zijn. Bij handmatige adressen: klik op “Opnieuw”.', 'idle', '?');
    window.setTimeout(() => checkRoute({ silentIfIncomplete: true }), 0);
  }

  button.addEventListener('click', () => {
    lastSignature = '';
    checkRoute({ geocodeManual: true });
  });

  [startInput, endInput].forEach((input) => {
    input.addEventListener('input', resetAndMaybeCheck);
  });

  [departureInput, arrivalInput].forEach((input) => {
    input.addEventListener('input', () => {
      delete input.dataset.latitude;
      delete input.dataset.longitude;
      resetAndMaybeCheck();
    });
  });

  document.addEventListener('rittenregistratie:location-filled', (event) => {
    lastSignature = '';
    if (event.detail && event.detail.targetId === 'arrivalAddress') {
      checkRoute({ silentIfIncomplete: true });
    } else {
      window.setTimeout(() => checkRoute({ silentIfIncomplete: true }), 0);
    }
  });
})();
