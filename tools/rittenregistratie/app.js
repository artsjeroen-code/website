(() => {
  'use strict';

  const API_BASE = './api';

  const form = document.getElementById('rideForm');
  const vehicleForm = document.getElementById('vehicleForm');
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
  const saveVehicleButton = document.getElementById('saveVehicle');
  const vehicleMessage = document.getElementById('vehicleMessage');
  const vehicleMake = document.getElementById('vehicleMake');
  const vehicleModel = document.getElementById('vehicleModel');
  const vehiclePlate = document.getElementById('vehiclePlate');
  const vehicleUseFrom = document.getElementById('vehicleUseFrom');
  const vehicleUseTo = document.getElementById('vehicleUseTo');

  let rides = [];
  let vehicle = null;
  let apiAvailable = false;
  let selectedYear = String(new Date().getFullYear());

  function localDateValue(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function numberValue(input) {
    if (input.value.trim() === '') return null;
    const value = Number(input.value);
    return Number.isFinite(value) ? value : null;
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

  function formatNumber(value) {
    return new Intl.NumberFormat('nl-NL', { maximumFractionDigits: 1 }).format(value);
  }

  function setMessage(message, kind = 'warning') {
    formMessage.textContent = message;
    formMessage.classList.toggle('success', kind === 'success');
  }

  function setVehicleMessage(message, kind = 'warning') {
    vehicleMessage.textContent = message;
    vehicleMessage.classList.toggle('success', kind === 'success');
  }

  function clearMessage() {
    setMessage('');
  }

  function latestRide() {
    return rides.length ? rides[rides.length - 1] : null;
  }

  function filteredRides() {
    return rides.filter((ride) => String(ride.date).startsWith(`${selectedYear}-`));
  }

  function fillPreviousOdometer() {
    const previous = latestRide();
    if (!previous) {
      setMessage('Er is nog geen vorige rit om een eindstand van over te nemen.');
      return;
    }
    startOdometer.value = previous.endOdometer;
    calculateDistance();
    setMessage(`Beginstand ingevuld met vorige eindstand: ${formatNumber(previous.endOdometer)} km.`, 'success');
  }

  function clearLocationDataset(input) {
    delete input.dataset.latitude;
    delete input.dataset.longitude;
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
    clearMessage();
    const previous = latestRide();
    if (previous) startOdometer.value = previous.endOdometer;
  }

  function coordsFromInput(input) {
    const lat = Number(input.dataset.latitude);
    const lon = Number(input.dataset.longitude);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    return { lat, lon };
  }

  function validateRide() {
    if (!form.reportValidity()) return null;
    if (!vehicle) {
      setMessage('Sla eerst de voertuiggegevens op.');
      return null;
    }

    const start = numberValue(startOdometer);
    const end = numberValue(endOdometer);
    if (start === null || end === null) {
      setMessage('Vul een geldige begin- en eindkilometerstand in.');
      return null;
    }
    if (!Number.isInteger(start) || !Number.isInteger(end)) {
      setMessage('Gebruik hele kilometers voor de kilometerstanden.');
      return null;
    }
    if (end < start) {
      setMessage('De eindkilometerstand kan niet lager zijn dan de beginstand.');
      return null;
    }

    const previous = latestRide();
    if (previous && start !== previous.endOdometer) {
      setMessage(`Niet sluitend: de vorige rit eindigde op ${formatNumber(previous.endOdometer)} km. Pas de beginstand aan voordat je opslaat.`);
      return null;
    }

    return {
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

  function fillVehicleForm() {
    if (!vehicle) return;
    vehicleMake.value = vehicle.make || '';
    vehicleModel.value = vehicle.model || '';
    vehiclePlate.value = vehicle.plate || '';
    vehicleUseFrom.value = vehicle.useFrom || '';
    vehicleUseTo.value = vehicle.useTo || '';
  }

  async function loadData() {
    try {
      const [ridesPayload, vehiclePayload] = await Promise.all([
        apiRequest('/rides'),
        apiRequest('/vehicle')
      ]);
      rides = Array.isArray(ridesPayload.rides) ? ridesPayload.rides : [];
      vehicle = vehiclePayload.vehicle || null;
      apiAvailable = true;
      fillVehicleForm();
      populateYearFilter();
      render();
      resetForm();
      submitButton.disabled = !vehicle;
      if (!vehicle) setVehicleMessage('Vul de voertuiggegevens in en sla ze op voordat je een nieuwe rit registreert.');
    } catch (error) {
      console.error('Kon centrale gegevens niet laden.', error);
      rides = [];
      vehicle = null;
      apiAvailable = false;
      submitButton.disabled = true;
      saveVehicleButton.disabled = true;
      populateYearFilter();
      render();
      setMessage(`Centrale opslag niet bereikbaar: ${error.message}. Er wordt niets lokaal opgeslagen.`);
    }
  }

  async function handleVehicleSubmit(event) {
    event.preventDefault();
    setVehicleMessage('');
    if (!apiAvailable) {
      setVehicleMessage('Centrale opslag is niet bereikbaar.');
      return;
    }
    if (!vehicleForm.reportValidity()) return;

    saveVehicleButton.disabled = true;
    const oldText = saveVehicleButton.textContent;
    saveVehicleButton.textContent = 'Opslaan…';
    try {
      const payload = await apiRequest('/vehicle', {
        method: 'PUT',
        body: JSON.stringify({
          make: vehicleMake.value.trim(),
          model: vehicleModel.value.trim(),
          plate: vehiclePlate.value.trim(),
          useFrom: vehicleUseFrom.value,
          useTo: vehicleUseTo.value || null
        })
      });
      vehicle = payload.vehicle;
      fillVehicleForm();
      submitButton.disabled = false;
      setVehicleMessage(`Voertuig opgeslagen: ${vehicle.make} ${vehicle.model} · ${vehicle.plate}.`, 'success');
    } catch (error) {
      setVehicleMessage(`Voertuig niet opgeslagen: ${error.message}`);
    } finally {
      saveVehicleButton.disabled = false;
      saveVehicleButton.textContent = oldText;
    }
  }

  async function handleSubmit(event) {
    event.preventDefault();
    clearMessage();
    if (!apiAvailable) {
      setMessage('Centrale opslag is niet bereikbaar. Vernieuw de pagina nadat de API weer beschikbaar is.');
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
      setMessage(`Rit centraal opgeslagen: ${formatNumber(payload.ride.distance)} km.`, 'success');
    } catch (error) {
      setMessage(`Rit niet opgeslagen: ${error.message}`);
    } finally {
      submitButton.disabled = !apiAvailable || !vehicle;
      submitButton.textContent = oldText;
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
      row.appendChild(createCell('—'));
      ridesBody.appendChild(row);
    });
  }

  function renderSummary() {
    const visibleRides = filteredRides();
    const businessTotal = visibleRides.filter((ride) => ride.type === 'business').reduce((sum, ride) => sum + ride.distance, 0);
    const privateTotal = visibleRides.filter((ride) => ride.type === 'private').reduce((sum, ride) => sum + ride.distance, 0);
    const lastVisibleRide = visibleRides.length ? visibleRides[visibleRides.length - 1] : null;
    totalRides.textContent = String(visibleRides.length);
    businessKm.textContent = `${formatNumber(businessTotal)} km`;
    privateKm.textContent = `${formatNumber(privateTotal)} km`;
    lastOdometer.textContent = lastVisibleRide ? `${formatNumber(lastVisibleRide.endOdometer)} km` : '—';
  }

  function render() {
    renderTable();
    renderSummary();
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
    if (!vehicle) {
      setMessage('Voertuiggegevens ontbreken; sla die eerst op voordat je exporteert.');
      return;
    }

    const businessTotal = visibleRides.filter((ride) => ride.type === 'business').reduce((sum, ride) => sum + ride.distance, 0);
    const privateTotal = visibleRides.filter((ride) => ride.type === 'private').reduce((sum, ride) => sum + ride.distance, 0);
    const firstRide = visibleRides[0];
    const lastRide = visibleRides[visibleRides.length - 1];

    const rows = [
      ['Rittenregistratie', selectedYear],
      [],
      ['Voertuig'],
      ['Merk', vehicle.make],
      ['Type / model', vehicle.model],
      ['Kenteken', vehicle.plate],
      ['In gebruik vanaf', vehicle.useFrom],
      ['In gebruik tot', vehicle.useTo || 'doorlopend'],
      [],
      ['Jaaroverzicht'],
      ['Aantal ritten', visibleRides.length],
      ['Begin kilometerstand', firstRide.startOdometer],
      ['Eind kilometerstand', lastRide.endOdometer],
      ['Zakelijke kilometers', businessTotal],
      ['Privékilometers', privateTotal],
      ['Totaal kilometers', businessTotal + privateTotal],
      [],
      ['Datum', 'Type', 'Vertrekadres', 'Aankomstadres', 'Begin km-stand', 'Eind km-stand', 'Kilometers', 'Toelichting'],
      ...visibleRides.map((ride) => [
        ride.date,
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
    link.download = `rittenregistratie-${vehicle.plate}-${selectedYear}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setMessage(`Jaaroverzicht ${selectedYear} geëxporteerd voor ${vehicle.plate}.`, 'success');
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
