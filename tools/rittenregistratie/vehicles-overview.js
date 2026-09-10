(() => {
  'use strict';

  const API_BASE = './api';
  const body = document.getElementById('vehiclesOverviewBody');
  const empty = document.getElementById('vehiclesOverviewEmpty');
  if (!body || !empty) return;

  function formatDate(value) {
    if (!value) return 'Doorlopend';
    const [year, month, day] = String(value).split('-');
    return year && month && day ? `${day}-${month}-${year}` : String(value);
  }

  function formatNumber(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return '—';
    return new Intl.NumberFormat('nl-NL', { maximumFractionDigits: 0 }).format(number);
  }

  async function api(path) {
    const response = await fetch(`${API_BASE}${path}`, {
      cache: 'no-store',
      credentials: 'same-origin',
      headers: { Accept: 'application/json' }
    });
    let payload = {};
    try { payload = await response.json(); }
    catch { throw new Error(`API gaf geen geldige JSON (HTTP ${response.status})`); }
    if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
    return payload;
  }

  function createCell(text, className = '') {
    const cell = document.createElement('td');
    cell.textContent = text;
    if (className) cell.className = className;
    return cell;
  }

  function lastOdometerFor(vehicle, rides) {
    const matches = rides
      .filter((ride) => ride.vehicleId === vehicle.id)
      .map((ride) => Number(ride.endOdometer))
      .filter(Number.isFinite);
    if (matches.length) return Math.max(...matches);
    return Number.isFinite(Number(vehicle.initialOdometer)) ? Number(vehicle.initialOdometer) : null;
  }

  function render(vehicles, rides) {
    body.replaceChildren();
    empty.hidden = vehicles.length > 0;

    vehicles.forEach((vehicle) => {
      const vehicleRides = rides.filter((ride) => ride.vehicleId === vehicle.id);
      const row = document.createElement('tr');
      row.appendChild(createCell(vehicle.plate || '—'));
      row.appendChild(createCell([vehicle.make, vehicle.model].filter(Boolean).join(' ') || '—'));
      row.appendChild(createCell(formatDate(vehicle.useFrom)));
      row.appendChild(createCell(`${formatNumber(vehicle.initialOdometer)} km`, 'numeric'));
      row.appendChild(createCell(vehicle.useTo ? formatDate(vehicle.useTo) : 'Doorlopend'));
      row.appendChild(createCell(String(vehicleRides.length), 'numeric'));
      const last = lastOdometerFor(vehicle, rides);
      row.appendChild(createCell(last === null ? '—' : `${formatNumber(last)} km`, 'numeric'));
      body.appendChild(row);
    });
  }

  async function load() {
    try {
      const [vehiclesPayload, ridesPayload] = await Promise.all([api('/vehicles'), api('/rides')]);
      const vehicles = Array.isArray(vehiclesPayload.vehicles) ? vehiclesPayload.vehicles : [];
      const rides = Array.isArray(ridesPayload.rides) ? ridesPayload.rides : [];
      render(vehicles, rides);
    } catch (error) {
      body.replaceChildren();
      empty.hidden = false;
      empty.textContent = `Voertuigen konden niet worden geladen: ${error.message}`;
    }
  }

  document.querySelector('[data-view="vehicles"]')?.addEventListener('click', load);
  document.addEventListener('rittenregistratie:data-changed', load);
  load();
})();
