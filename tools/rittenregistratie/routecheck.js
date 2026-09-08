(() => {
  'use strict';

  const API_ROUTE = './api/route';
  const button = document.getElementById('checkRoute');
  const result = document.getElementById('routeCheckResult');
  const startInput = document.getElementById('startOdometer');
  const endInput = document.getElementById('endOdometer');
  const departureInput = document.getElementById('departureAddress');
  const arrivalInput = document.getElementById('arrivalAddress');

  if (!button || !result || !startInput || !endInput || !departureInput || !arrivalInput) return;

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

  function show(message, status = '') {
    result.textContent = message;
    result.dataset.status = status;
  }

  async function checkRoute() {
    const departure = coords(departureInput);
    const arrival = coords(arrivalInput);
    const drivenKm = odometerDistance();

    if (!departure || !arrival) {
      show('Gebruik eerst bij vertrek én aankomst de knop “Gebruik locatie”, zodat GPS-coördinaten beschikbaar zijn.', 'warning');
      return;
    }

    if (drivenKm === null) {
      show('Vul eerst een geldige begin- en eindkilometerstand in.', 'warning');
      return;
    }

    const original = button.textContent;
    button.disabled = true;
    button.textContent = 'Route berekenen…';
    show('Normale autoroute wordt berekend…');

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
      const threshold = Math.max(3, routeKm * 0.2);

      if (absoluteDifference <= threshold) {
        show(
          `Controle OK: teller ${formatKm(drivenKm)} km · normale route ${formatKm(routeKm)} km · verschil ${formatKm(absoluteDifference)} km.`,
          'ok'
        );
      } else {
        const direction = difference > 0 ? 'meer' : 'minder';
        show(
          `Let op: teller ${formatKm(drivenKm)} km · normale route ${formatKm(routeKm)} km. Je reed ${formatKm(absoluteDifference)} km ${direction} dan de berekende route. Controleer invoer of noteer een afwijkende route.`,
          'warning'
        );
      }
    } catch (error) {
      console.warn('Routecontrole mislukt.', error);
      show(`Routecontrole tijdelijk niet beschikbaar: ${error.message}. De rit kan wel gewoon worden opgeslagen.`, 'warning');
    } finally {
      button.disabled = false;
      button.textContent = original;
    }
  }

  button.addEventListener('click', checkRoute);

  [startInput, endInput, departureInput, arrivalInput].forEach((input) => {
    input.addEventListener('input', () => show('Nog niet gecontroleerd.'));
  });
})();
