(() => {
  'use strict';

  const API_BASE = './api';
  const RDW_API = 'https://opendata.rdw.nl/resource/m9d7-ebf2.json';
  const form = document.getElementById('rideForm');
  const vehicleForm = document.getElementById('vehicleForm');
  const vehicleSelect = document.getElementById('vehicleSelect');
  const selectedVehicleLabel = document.getElementById('selectedVehicleLabel');
  const vehicleStatusMessage = document.getElementById('vehicleStatusMessage');
  const toggleVehicleForm = document.getElementById('toggleVehicleForm');
  const cancelVehicle = document.getElementById('cancelVehicle');
  const saveVehicleButton = document.getElementById('saveVehicle');
  const vehicleMessage = document.getElementById('vehicleMessage');
  const vehicleMake = document.getElementById('vehicleMake');
  const vehicleModel = document.getElementById('vehicleModel');
  const vehiclePlate = document.getElementById('vehiclePlate');
  const vehicleUseFrom = document.getElementById('vehicleUseFrom');
  const vehicleUseTo = document.getElementById('vehicleUseTo');
  const vehicleInitialOdometer = document.getElementById('vehicleInitialOdometer');
  const lookupRdwButton = document.getElementById('lookupRdw');
  const rdwStatus = document.getElementById('rdwStatus');

  const rideDate = document.getElementById('rideDate');
  const rideType = document.getElementById('rideType');
  const startOdometer = document.getElementById('startOdometer');
  const endOdometer = document.getElementById('endOdometer');
  const departureAddress = document.getElementById('departureAddress');
  const arrivalAddress = document.getElementById('arrivalAddress');
  const notes = document.getElementById('notes');
  const distancePreview = document.getElementById('distancePreview');
  const formMessage = document.getElementById('formMessage');
  const ridesBody = document.getElementById('ridesBody');
  const emptyState = document.getElementById('emptyState');
  const totalRides = document.getElementById('totalRides');
  const businessKm = document.getElementById('businessKm');
  const privateKm = document.getElementById('privateKm');
  const lastOdometer = document.getElementById('lastOdometer');
  const departureMeta = document.getElementById('departureMeta');
  const arrivalMeta = document.getElementById('arrivalMeta');
  const yearFilter = document.getElementById('yearFilter');
  const submitButton = form.querySelector('button[type="submit"]');

  let rides = [];
  let vehicles = [];
  let selectedVehicleId = null;
  let apiAvailable = false;
  let selectedYear = String(new Date().getFullYear());

  function localDateValue(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat('nl-NL', { maximumFractionDigits: 1 }).format(value);
  }

  function numberValue(input) {
    if (input.value.trim() === '') return null;
    const value = Number(input.value);
    return Number.isFinite(value) ? value : null;
  }

  function setMessage(message, success = false) {
    formMessage.textContent = message;
    formMessage.classList.toggle('success', success);
  }

  function setVehicleMessage(message, success = false) {
    vehicleMessage.textContent = message;
    vehicleMessage.classList.toggle('success', success);
  }

  function normalizePlate(value) {
    return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  function selectedVehicle() {
    return vehicles.find((vehicle) => vehicle.id === selectedVehicleId) || null;
  }

  function ridesForSelectedVehicle() {
    return rides.filter((ride) => ride.vehicleId === selectedVehicleId);
  }

  function latestRideForSelectedVehicle() {
    const matches = ridesForSelectedVehicle();
    return matches.length ? matches[matches.length - 1] : null;
  }

  function filteredRides() {
    return rides.filter((ride) => String(ride.date).startsWith(`${selectedYear}-`));
  }

  function calculateDistance() {
    const start = numberValue(startOdometer);
    const end = numberValue(endOdometer);
    if (start === null || end === null || end < start) {
      distancePreview.textContent = '— km';
      return null;
    }
    const distance = end - start;
    distancePreview.textContent = `${formatNumber(distance)} km`;
    return distance;
  }

  function clearLocationDataset(input) {
    delete input.dataset.latitude;
    delete input.dataset.longitude;
  }

  function coordsFromInput(input) {
    const lat = Number(input.dataset.latitude);
    const lon = Number(input.dataset.longitude);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    return { lat, lon };
  }

  function resetForm({ keepDate = true } = {}) {
    const currentDate = rideDate.value || localDateValue();
    form.reset();
    rideType.value = 'business';
    rideDate.value = keepDate ? currentDate : localDateValue();
    departureMeta.textContent = '';
    arrivalMeta.textContent = '';
    clearLocationDataset(departureAddress);
    clearLocationDataset(arrivalAddress);
    distancePreview.textContent = '— km';
    setMessage('');
    const previous = latestRideForSelectedVehicle();
    const vehicle = selectedVehicle();
    if (previous) {
      startOdometer.value = previous.endOdometer;
    } else if (vehicle && Number.isInteger(vehicle.initialOdometer)) {
      startOdometer.value = vehicle.initialOdometer;
    }
  }

  function fillPreviousOdometer() {
    const vehicle = selectedVehicle();
    if (!vehicle) {
      setMessage('Kies eerst een kenteken.');
      return;
    }
    const previous = latestRideForSelectedVehicle();
    if (!previous) {
      if (Number.isInteger(vehicle.initialOdometer)) {
        startOdometer.value = vehicle.initialOdometer;
        calculateDistance();
        setMessage(`Beginstand voor ${vehicle.plate} is de vastgelegde stand bij ingebruikname: ${formatNumber(vehicle.initialOdometer)} km.`, true);
      } else {
        setMessage(`${vehicle.plate} heeft nog geen eerdere rit. Vul de beginstand van dit voertuig in.`);
      }
      return;
    }
    startOdometer.value = previous.endOdometer;
    calculateDistance();
    setMessage(`Beginstand voor ${vehicle.plate}: ${formatNumber(previous.endOdometer)} km.`, true);
  }

  async function apiRequest(path, options = {}) {
    const response = await fetch(`${API_BASE}${path}`, {
      cache: 'no-store',
      headers: {
        Accept: 'application/json',
        ...(options.body ? { 'Content-Type': 'application/json' } : {}),
        ...(options.headers || {})
      },
      ...options
    });
    let payload = {};
    try {
      payload = await response.json();
    } catch (error) {
      throw new Error(`API gaf geen geldige JSON (HTTP ${response.status})`);
    }
    if (!response.ok) throw new Error(payload.error || `API-fout HTTP ${response.status}`);
    return payload;
  }

  async function lookupRdwVehicle() {
    const plate = normalizePlate(vehiclePlate.value);
    if (!plate) {
      setVehicleMessage('Vul eerst een kenteken in.');
      vehiclePlate.focus();
      return;
    }

    lookupRdwButton.disabled = true;
    const oldText = lookupRdwButton.textContent;
    lookupRdwButton.textContent = 'Ophalen…';
    rdwStatus.textContent = '';
    setVehicleMessage('');
    vehicleMake.value = '';
    vehicleModel.value = '';

    try {
      const response = await fetch(`${RDW_API}?kenteken=${encodeURIComponent(plate)}`, {
        headers: { Accept: 'application/json' },
        cache: 'no-store'
      });
      if (!response.ok) throw new Error(`RDW gaf HTTP ${response.status}`);
      const rows = await response.json();
      if (!Array.isArray(rows) || !rows.length) {
        throw new Error('Kenteken niet gevonden in RDW Open Data');
      }
      const item = rows[0];
      if (!item.merk || !item.handelsbenaming) {
        throw new Error('RDW gaf geen bruikbaar merk/type terug');
      }
      vehiclePlate.value = plate;
      vehicleMake.value = item.merk;
      vehicleModel.value = item.handelsbenaming;
      rdwStatus.textContent = `RDW gevonden: ${item.merk} ${item.handelsbenaming}${item.voertuigsoort ? ` · ${item.voertuigsoort}` : ''}.`;
      vehicleUseFrom.focus();
    } catch (error) {
      setVehicleMessage(`RDW-gegevens konden niet worden opgehaald: ${error.message}`);
    } finally {
      lookupRdwButton.disabled = false;
      lookupRdwButton.textContent = oldText;
    }
  }

  function populateVehicleSelect() {
    vehicleSelect.replaceChildren();
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = vehicles.length ? 'Kies kenteken…' : 'Nog geen voertuigen';
    vehicleSelect.appendChild(placeholder);

    vehicles.forEach((vehicle) => {
      const option = document.createElement('option');
      option.value = String(vehicle.id);
      option.textContent = `${vehicle.plate} · ${vehicle.make} ${vehicle.model}`;
      option.selected = vehicle.id === selectedVehicleId;
      vehicleSelect.appendChild(option);
    });

    if (selectedVehicleId) vehicleSelect.value = String(selectedVehicleId);
    updateSelectedVehicleUi();
  }

  function updateSelectedVehicleUi() {
    const vehicle = selectedVehicle();
    selectedVehicleLabel.textContent = vehicle ? `${vehicle.plate} · ${vehicle.make} ${vehicle.model}` : '—';
    submitButton.disabled = !apiAvailable || !vehicle;
    const previous = latestRideForSelectedVehicle();
    lastOdometer.textContent = previous ? `${formatNumber(previous.endOdometer)} km` : '—';
    if (vehicle) {
      vehicleStatusMessage.textContent = previous
        ? `Kilometerketen actief voor ${vehicle.plate}; laatste stand ${formatNumber(previous.endOdometer)} km.`
        : Number.isInteger(vehicle.initialOdometer)
          ? `${vehicle.plate} start op ${formatNumber(vehicle.initialOdometer)} km per ${vehicle.useFrom}.`
          : `${vehicle.plate} start een eigen kilometerketen.`;
      resetForm();
    } else {
      vehicleStatusMessage.textContent = 'Kies een kenteken om een rit te registreren.';
    }
  }

  function populateYearFilter() {
    const years = new Set(rides.map((ride) => String(ride.date).slice(0, 4)));
    years.add(String(new Date().getFullYear()));
    years.add(String(new Date().getFullYear() + 1));
    if (!years.has(selectedYear)) selectedYear = [...years].sort().reverse()[0];
    yearFilter.replaceChildren();
    [...years].sort().reverse().forEach((year) => {
      const option = document.createElement('option');
      option.value = year;
      option.textContent = year;
      option.selected = year === selectedYear;
      yearFilter.appendChild(option);
    });
  }

  function createCell(text, className = '') {
    const cell = document.createElement('td');
    cell.textContent = text;
    if (className) cell.className = className;
    return cell;
  }

  function renderTable() {
    const visibleRides = filteredRides();
    ridesBody.replaceChildren();
    emptyState.hidden = visibleRides.length > 0;
    emptyState.textContent = `Nog geen ritten opgeslagen voor ${selectedYear}.`;

    visibleRides.forEach((ride) => {
      const row = document.createElement('tr');
      row.appendChild(createCell(formatDate(ride.date)));
      row.appendChild(createCell(ride.vehiclePlate || '—'));
      const typeCell = document.createElement('td');
      const pill = document.createElement('span');
      pill.className = 'type-pill';
      pill.textContent = ride.type === 'private' ? 'Privé' : 'Zakelijk';
      typeCell.appendChild(pill);
      row.appendChild(typeCell);
      row.appendChild(createCell(ride.departureAddress));
      row.appendChild(createCell(ride.arrivalAddress));
      row.appendChild(createCell(formatNumber(ride.startOdometer), 'numeric'));
      row.appendChild(createCell(formatNumber(ride.endOdometer), 'numeric'));
      row.appendChild(createCell(formatNumber(ride.distance), 'numeric'));
      row.appendChild(createCell(ride.notes || '—'));
      ridesBody.appendChild(row);
    });
  }

  function renderSummary() {
    const visibleRides = filteredRides();
    const businessTotal = visibleRides.filter((ride) => ride.type === 'business').reduce((sum, ride) => sum + ride.distance, 0);
    const privateTotal = visibleRides.filter((ride) => ride.type === 'private').reduce((sum, ride) => sum + ride.distance, 0);
    totalRides.textContent = String(visibleRides.length);
    businessKm.textContent = `${formatNumber(businessTotal)} km`;
    privateKm.textContent = `${formatNumber(privateTotal)} km`;
    const previous = latestRideForSelectedVehicle();
    lastOdometer.textContent = previous ? `${formatNumber(previous.endOdometer)} km` : '—';
  }

  function render() {
    renderTable();
    renderSummary();
  }

  function validateRide() {
    if (!form.reportValidity()) return null;
    const vehicle = selectedVehicle();
    if (!vehicle) {
      setMessage('Kies eerst een kenteken.');
      return null;
    }
    const start = numberValue(startOdometer);
    const end = numberValue(endOdometer);
    if (start === null || end === null || !Number.isInteger(start) || !Number.isInteger(end)) {
      setMessage('Gebruik hele kilometers voor de kilometerstanden.');
      return null;
    }
    if (end < start) {
      setMessage('De eindkilometerstand kan niet lager zijn dan de beginstand.');
      return null;
    }
    const previous = latestRideForSelectedVehicle();
    if (previous && start !== previous.endOdometer) {
      setMessage(`Niet sluitend voor ${vehicle.plate}: de vorige rit eindigde op ${formatNumber(previous.endOdometer)} km.`);
      return null;
    }
    if (!previous && Number.isInteger(vehicle.initialOdometer) && start !== vehicle.initialOdometer) {
      setMessage(`De eerste rit van ${vehicle.plate} moet beginnen op de vastgelegde stand bij ingebruikname: ${formatNumber(vehicle.initialOdometer)} km.`);
      return null;
    }
    return {
      vehicleId: vehicle.id,
      date: rideDate.value,
      type: rideType.value,
      startOdometer: start,
      endOdometer: end,
      departureAddress: departureAddress.value.trim(),
      arrivalAddress: arrivalAddress.value.trim(),
      departureCoords: coordsFromInput(departureAddress),
      arrivalCoords: coordsFromInput(arrivalAddress),
      notes: notes.value.trim()
    };
  }

  async function handleSubmit(event) {
    event.preventDefault();
    setMessage('');
    if (!apiAvailable) {
      setMessage('Centrale opslag is niet bereikbaar.');
      return;
    }
    const ride = validateRide();
    if (!ride) return;

    submitButton.disabled = true;
    const oldText = submitButton.textContent;
    submitButton.textContent = 'Opslaan…';
    try {
      const payload = await apiRequest('/rides', { method: 'POST', body: JSON.stringify(ride) });
      rides.push(payload.ride);
      selectedYear = String(payload.ride.date).slice(0, 4);
      populateYearFilter();
      render();
      resetForm();
      setMessage(`Rit voor ${payload.ride.vehiclePlate} opgeslagen: ${formatNumber(payload.ride.distance)} km.`, true);
    } catch (error) {
      setMessage(`Rit niet opgeslagen: ${error.message}`);
    } finally {
      submitButton.disabled = !apiAvailable || !selectedVehicle();
      submitButton.textContent = oldText;
    }
  }

  function showVehicleForm(show) {
    vehicleForm.hidden = !show;
    toggleVehicleForm.hidden = show;
    if (show) {
      vehicleForm.reset();
      vehicleMake.value = '';
      vehicleModel.value = '';
      rdwStatus.textContent = '';
      vehicleUseFrom.value = localDateValue();
      setVehicleMessage('');
      vehiclePlate.focus();
    }
  }

  async function handleVehicleSubmit(event) {
    event.preventDefault();
    if (!vehicleForm.reportValidity()) return;
    const initialOdometer = Number(vehicleInitialOdometer.value);
    if (!Number.isInteger(initialOdometer) || initialOdometer < 0) {
      setVehicleMessage('Vul een geldige kilometerstand bij ingebruikname in.');
      return;
    }
    if (!vehicleMake.value || !vehicleModel.value) {
      setVehicleMessage('Haal eerst de voertuiggegevens bij RDW op.');
      return;
    }

    saveVehicleButton.disabled = true;
    const oldText = saveVehicleButton.textContent;
    saveVehicleButton.textContent = 'Toevoegen…';
    try {
      const payload = await apiRequest('/vehicles', {
        method: 'POST',
        body: JSON.stringify({
          make: vehicleMake.value.trim(),
          model: vehicleModel.value.trim(),
          plate: normalizePlate(vehiclePlate.value),
          useFrom: vehicleUseFrom.value,
          initialOdometer,
          useTo: vehicleUseTo.value || null
        })
      });
      vehicles.push(payload.vehicle);
      selectedVehicleId = payload.vehicle.id;
      populateVehicleSelect();
      showVehicleForm(false);
      vehicleStatusMessage.textContent = `${payload.vehicle.plate} toegevoegd met beginstand ${formatNumber(payload.vehicle.initialOdometer)} km per ${payload.vehicle.useFrom}.`;
      resetForm();
    } catch (error) {
      setVehicleMessage(`Voertuig niet toegevoegd: ${error.message}`);
    } finally {
      saveVehicleButton.disabled = false;
      saveVehicleButton.textContent = oldText;
    }
  }

  function formatDate(value) {
    if (!value) return '—';
    const [year, month, day] = value.split('-');
    return `${day}-${month}-${year}`;
  }

  function csvEscape(value) {
    const text = String(value ?? '');
    return `"${text.replaceAll('"', '""')}"`;
  }

  function exportCsv() {
    const visibleRides = filteredRides();
    if (!visibleRides.length) {
      setMessage(`Er zijn geen ritten voor ${selectedYear} om te exporteren.`);
      return;
    }
    const businessTotal = visibleRides.filter((ride) => ride.type === 'business').reduce((sum, ride) => sum + ride.distance, 0);
    const privateTotal = visibleRides.filter((ride) => ride.type === 'private').reduce((sum, ride) => sum + ride.distance, 0);
    const usedVehicleIds = new Set(visibleRides.map((ride) => ride.vehicleId));
    const usedVehicles = vehicles.filter((vehicle) => usedVehicleIds.has(vehicle.id));

    const rows = [
      ['Rittenregistratie', selectedYear],
      [],
      ['Voertuigen in dit jaar'],
      ['Kenteken', 'Merk', 'Type / model', 'In gebruik vanaf', 'Beginstand bij ingebruikname', 'In gebruik tot'],
      ...usedVehicles.map((vehicle) => [vehicle.plate, vehicle.make, vehicle.model, vehicle.useFrom, vehicle.initialOdometer ?? '', vehicle.useTo || 'doorlopend']),
      [],
      ['Jaaroverzicht'],
      ['Aantal ritten', visibleRides.length],
      ['Zakelijke kilometers', businessTotal],
      ['Privékilometers', privateTotal],
      ['Totaal kilometers', businessTotal + privateTotal],
      [],
      ['Datum', 'Kenteken', 'Type', 'Vertrekadres', 'Aankomstadres', 'Begin km-stand', 'Eind km-stand', 'Kilometers', 'Toelichting'],
      ...visibleRides.map((ride) => [
        ride.date,
        ride.vehiclePlate || '',
        ride.type === 'private' ? 'Privé' : 'Zakelijk',
        ride.departureAddress,
        ride.arrivalAddress,
        ride.startOdometer,
        ride.endOdometer,
        ride.distance,
        ride.notes
      ])
    ];

    const content = rows.map((row) => row.map(csvEscape).join(';')).join('\r\n');
    const blob = new Blob([`\uFEFF${content}`], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `rittenregistratie-${selectedYear}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setMessage(`Jaaroverzicht ${selectedYear} geëxporteerd.`, true);
  }

  async function loadData() {
    try {
      const [ridesPayload, vehiclesPayload] = await Promise.all([
        apiRequest('/rides'),
        apiRequest('/vehicles')
      ]);
      rides = Array.isArray(ridesPayload.rides) ? ridesPayload.rides : [];
      vehicles = Array.isArray(vehiclesPayload.vehicles) ? vehiclesPayload.vehicles : [];
      apiAvailable = true;
      selectedVehicleId = vehicles.length ? vehicles[vehicles.length - 1].id : null;
      populateVehicleSelect();
      populateYearFilter();
      render();
      resetForm();
      if (!vehicles.length) {
        vehicleStatusMessage.textContent = 'Voeg eerst een voertuig toe.';
        showVehicleForm(true);
      }
    } catch (error) {
      console.error('Kon centrale gegevens niet laden.', error);
      rides = [];
      vehicles = [];
      apiAvailable = false;
      submitButton.disabled = true;
      saveVehicleButton.disabled = true;
      render();
      setMessage(`Centrale opslag niet bereikbaar: ${error.message}.`);
    }
  }

  function setupThemeToggle() {
    const button = document.getElementById('railTheme');
    if (!button) return;
    button.addEventListener('click', () => {
      const next = document.documentElement.dataset.theme === 'light' ? 'dark' : 'light';
      document.documentElement.dataset.theme = next;
      localStorage.setItem('startpagina-theme', next);
    });
  }

  rideDate.value = localDateValue();
  submitButton.disabled = true;
  startOdometer.addEventListener('input', calculateDistance);
  endOdometer.addEventListener('input', calculateDistance);
  form.addEventListener('submit', handleSubmit);
  vehicleForm.addEventListener('submit', handleVehicleSubmit);
  vehicleSelect.addEventListener('change', () => {
    selectedVehicleId = vehicleSelect.value ? Number(vehicleSelect.value) : null;
    updateSelectedVehicleUi();
    render();
  });
  toggleVehicleForm.addEventListener('click', () => showVehicleForm(true));
  cancelVehicle.addEventListener('click', () => showVehicleForm(false));
  lookupRdwButton.addEventListener('click', lookupRdwVehicle);
  vehiclePlate.addEventListener('blur', () => {
    if (normalizePlate(vehiclePlate.value) && !vehicleMake.value) lookupRdwVehicle();
  });
  document.getElementById('fillFromPrevious').addEventListener('click', fillPreviousOdometer);
  document.getElementById('resetForm').addEventListener('click', () => resetForm());
  document.getElementById('exportCsv').addEventListener('click', exportCsv);
  yearFilter.addEventListener('change', () => {
    selectedYear = yearFilter.value;
    render();
  });

  setupThemeToggle();
  populateYearFilter();
  render();
  loadData();
})();
