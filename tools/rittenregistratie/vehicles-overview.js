(() => {
  'use strict';

  const API_BASE = './api';
  const body = document.getElementById('vehiclesOverviewBody');
  const empty = document.getElementById('vehiclesOverviewEmpty');
  if (!body || !empty) return;

  const headerRow = body.closest('table')?.querySelector('thead tr');
  if (headerRow && !headerRow.querySelector('[data-vehicles-total-km]')) {
    const totalHeader = document.createElement('th');
    totalHeader.textContent = 'Totaal km';
    totalHeader.dataset.vehiclesTotalKm = 'true';
    const actionHeader = document.createElement('th');
    actionHeader.textContent = 'Actie';
    actionHeader.dataset.vehiclesAction = 'true';
    headerRow.append(totalHeader, actionHeader);
  }

  function formatDate(value) {
    if (!value) return 'Doorlopend';
    const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return String(value);
    return `${match[3]}-${match[2]}-${match[1]}`;
  }

  function formatNumber(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return '—';
    return new Intl.NumberFormat('nl-NL', { maximumFractionDigits: 0 }).format(number);
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

  function totalKmFor(vehicle, rides) {
    const initial = Number(vehicle.initialOdometer);
    const last = lastOdometerFor(vehicle, rides);
    if (!Number.isFinite(initial) || !Number.isFinite(last)) return null;
    return Math.max(0, last - initial);
  }

  function dateInput(value, label) {
    const input = document.createElement('input');
    input.type = 'date';
    input.value = value || '';
    input.setAttribute('aria-label', label);
    return input;
  }

  function numberInput(value, label) {
    const input = document.createElement('input');
    input.type = 'number';
    input.min = '0';
    input.step = '1';
    input.inputMode = 'numeric';
    input.value = value ?? '';
    input.setAttribute('aria-label', label);
    return input;
  }

  function renderEditRow(row, vehicle, rides, reload) {
    const cells = row.children;
    const useFromInput = dateInput(vehicle.useFrom, 'In gebruik vanaf');
    const initialInput = numberInput(vehicle.initialOdometer, 'Beginstand');
    const useToInput = dateInput(vehicle.useTo, 'In gebruik tot');

    cells[2].replaceChildren(useFromInput);
    cells[3].replaceChildren(initialInput);
    cells[4].replaceChildren(useToInput);

    const actionCell = cells[cells.length - 1];
    const save = document.createElement('button');
    save.type = 'button';
    save.className = 'primary-button';
    save.textContent = 'Opslaan';

    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'secondary-button';
    cancel.textContent = 'Annuleren';

    const message = document.createElement('div');
    message.className = 'field-meta';

    save.addEventListener('click', async () => {
      const initialOdometer = Number(initialInput.value);
      if (!useFromInput.value) {
        message.textContent = 'Begindatum is verplicht.';
        return;
      }
      if (!Number.isInteger(initialOdometer) || initialOdometer < 0) {
        message.textContent = 'Vul een geldige beginstand in.';
        return;
      }

      save.disabled = true;
      cancel.disabled = true;
      message.textContent = 'Opslaan…';
      try {
        await api(`/vehicles/${vehicle.id}`, {
          method: 'PUT',
          body: JSON.stringify({
            make: vehicle.make,
            model: vehicle.model,
            plate: vehicle.plate,
            useFrom: useFromInput.value,
            useTo: useToInput.value || null,
            initialOdometer
          })
        });
        document.dispatchEvent(new CustomEvent('rittenregistratie:data-changed'));
        await reload();
      } catch (error) {
        message.textContent = `Niet opgeslagen: ${error.message}`;
        save.disabled = false;
        cancel.disabled = false;
      }
    });

    cancel.addEventListener('click', reload);
    actionCell.replaceChildren(save, cancel, message);
  }

  function render(vehicles, rides, reload) {
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

      const total = totalKmFor(vehicle, rides);
      row.appendChild(createCell(total === null ? '—' : `${formatNumber(total)} km`, 'numeric'));

      const actionCell = document.createElement('td');
      const edit = document.createElement('button');
      edit.type = 'button';
      edit.className = 'secondary-button';
      edit.textContent = 'Aanpassen';
      edit.addEventListener('click', () => renderEditRow(row, vehicle, rides, reload));
      actionCell.appendChild(edit);
      row.appendChild(actionCell);

      body.appendChild(row);
    });
  }

  async function load() {
    try {
      const [vehiclesPayload, ridesPayload] = await Promise.all([api('/vehicles'), api('/rides')]);
      const vehicles = Array.isArray(vehiclesPayload.vehicles) ? vehiclesPayload.vehicles : [];
      const rides = Array.isArray(ridesPayload.rides) ? ridesPayload.rides : [];
      render(vehicles, rides, load);
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
