(() => {
  'use strict';

  const API_ROUTE = './api/route';
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

  async function checkRoute({ silentIfIncomplete = false } = {}) {
    if (running) return;

    const departure = coords(departureInput);
    const arrival = coords(arrivalInput);
    const drivenKm = odometerDistance();

    if (!departure || !arrival) {
      if (!silentIfIncomplete) show('Gebruik voor vertrek en aankomst “Gebruik locatie” om de route automatisch te controleren.', 'question', '?');
      return;
    }

    if (drivenKm === null) {
      if (!silentIfIncomplete) show('Vul eerst een geldige begin- en eindkilometerstand in.', 'question', '?');
      return;
    }

    const currentSignature = signature();
    if (!currentSignature || currentSignature === lastSignature) return;

    running = true;
    const original = button.textContent;
    button.disabled = true;
    button.textContent = 'Bezig…';
    show('Normale autoroute wordt berekend…', 'checking', '…');

    try {
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
    show('Wordt automatisch uitgevoerd zodra beide locaties en kilometerstanden bekend zijn.', 'idle', '?');
    window.setTimeout(() => checkRoute({ silentIfIncomplete: true }), 0);
  }

  button.addEventListener('click', () => {
    lastSignature = '';
    checkRoute();
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
