(() => {
  'use strict';

  const STORAGE_KEY = 'rittenregistratie-demo-v1';

  const form = document.getElementById('rideForm');
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

  let rides = loadRides();
  let locationState = {
    departureAddress: null,
    arrivalAddress: null
  };

  function localDateValue(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function loadRides() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
      return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      console.warn('Kon lokale ritten niet lezen.', error);
      return [];
    }
  }

  function saveRides() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rides));
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

  function clearMessage() {
    setMessage('');
  }

  function latestRide() {
    return rides.length ? rides[rides.length - 1] : null;
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

  function resetForm({ keepDate = true } = {}) {
    const currentDate = rideDate.value || localDateValue();
    form.reset();
    rideType.value = 'business';
    rideDate.value = keepDate ? currentDate : localDateValue();
    locationState = { departureAddress: null, arrivalAddress: null };
    departureMeta.textContent = '';
    arrivalMeta.textContent = '';
    distancePreview.textContent = '— km';
    clearMessage();

    const previous = latestRide();
    if (previous) startOdometer.value = previous.endOdometer;
  }

  function validateRide() {
    if (!form.reportValidity()) return null;

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
      id: typeof crypto !== 'undefined' && crypto.randomUUID
        ? crypto.randomUUID()
        : `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      date: rideDate.value,
      type: rideType.value,
      startOdometer: start,
      endOdometer: end,
      distance: end - start,
      departureAddress: departureAddress.value.trim(),
      arrivalAddress: arrivalAddress.value.trim(),
      departureCoords: locationState.departureAddress,
      arrivalCoords: locationState.arrivalAddress,
      notes: notes.value.trim(),
      createdAt: new Date().toISOString()
    };
  }

  function handleSubmit(event) {
    event.preventDefault();
    clearMessage();

    const ride = validateRide();
    if (!ride) return;

    rides.push(ride);
    saveRides();
    render();
    resetForm();
    setMessage(`Rit opgeslagen: ${formatNumber(ride.distance)} km.`, 'success');
  }

  function createCell(text, className = '') {
    const cell = document.createElement('td');
    cell.textContent = text;
    if (className) cell.className = className;
    return cell;
  }

  function renderTable() {
    ridesBody.replaceChildren();
    emptyState.hidden = rides.length > 0;

    rides.forEach((ride) => {
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

      const actions = document.createElement('td');
      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.className = 'icon-button';
      deleteButton.textContent = 'Wis';
      deleteButton.setAttribute('aria-label', `Wis rit van ${formatDate(ride.date)}`);
      deleteButton.addEventListener('click', () => deleteRide(ride.id));
      actions.appendChild(deleteButton);
      row.appendChild(actions);

      ridesBody.appendChild(row);
    });
  }

  function renderSummary() {
    const businessTotal = rides
      .filter((ride) => ride.type === 'business')
      .reduce((sum, ride) => sum + ride.distance, 0);
    const privateTotal = rides
      .filter((ride) => ride.type === 'private')
      .reduce((sum, ride) => sum + ride.distance, 0);
    const previous = latestRide();

    totalRides.textContent = String(rides.length);
    businessKm.textContent = `${formatNumber(businessTotal)} km`;
    privateKm.textContent = `${formatNumber(privateTotal)} km`;
    lastOdometer.textContent = previous ? `${formatNumber(previous.endOdometer)} km` : '—';
  }

  function render() {
    renderTable();
    renderSummary();
  }

  function formatDate(value) {
    const [year, month, day] = value.split('-');
    return `${day}-${month}-${year}`;
  }

  function deleteRide(id) {
    const index = rides.findIndex((ride) => ride.id === id);
    if (index < 0) return;

    const isLast = index === rides.length - 1;
    const message = isLast
      ? 'Deze rit uit de lokale testopslag wissen?'
      : 'Deze rit staat midden in de registratie. Wissen kan de kilometerreeks niet-sluitend maken. Toch wissen?';

    if (!window.confirm(message)) return;

    rides.splice(index, 1);
    saveRides();
    render();
    resetForm();
  }

  function csvEscape(value) {
    const text = String(value ?? '');
    return `"${text.replaceAll('"', '""')}"`;
  }

  function exportCsv() {
    if (!rides.length) {
      setMessage('Er zijn nog geen ritten om te exporteren.');
      return;
    }

    const header = [
      'Datum', 'Type', 'Vertrekadres', 'Aankomstadres',
      'Begin km-stand', 'Eind km-stand', 'Kilometers', 'Toelichting'
    ];

    const rows = rides.map((ride) => [
      ride.date,
      ride.type === 'private' ? 'Privé' : 'Zakelijk',
      ride.departureAddress,
      ride.arrivalAddress,
      ride.startOdometer,
      ride.endOdometer,
      ride.distance,
      ride.notes
    ]);

    const content = [header, ...rows]
      .map((row) => row.map(csvEscape).join(';'))
      .join('\r\n');

    const blob = new Blob([`\uFEFF${content}`], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `rittenregistratie-${new Date().getFullYear()}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function clearAll() {
    if (!rides.length) return;
    if (!window.confirm('Alle lokale testgegevens wissen? Dit kan niet ongedaan worden gemaakt.')) return;

    rides = [];
    saveRides();
    render();
    resetForm({ keepDate: false });
    setMessage('Alle lokale testgegevens zijn gewist.', 'success');
  }

  function requestLocation(targetId, button) {
    if (!navigator.geolocation) {
      setMessage('Deze browser ondersteunt geen locatiebepaling.');
      return;
    }

    const originalText = button.textContent;
    const restoreButton = () => {
      button.disabled = false;
      button.textContent = originalText;
    };

    button.disabled = true;
    button.textContent = 'Locatie bepalen…';
    clearMessage();

    navigator.geolocation.getCurrentPosition(
      (position) => {
        const { latitude, longitude, accuracy } = position.coords;
        const coords = {
          lat: Number(latitude.toFixed(6)),
          lon: Number(longitude.toFixed(6)),
          accuracy: Math.round(accuracy)
        };
        locationState[targetId] = coords;

        const target = document.getElementById(targetId);
        target.value = `GPS ${coords.lat}, ${coords.lon}`;
        const meta = targetId === 'departureAddress' ? departureMeta : arrivalMeta;
        meta.textContent = `GPS gevonden (nauwkeurigheid ±${coords.accuracy} m). Adresomzetting volgt in stap 2.`;
        setMessage('Locatie gevonden. In de volgende ontwikkelstap wordt deze automatisch naar een straatadres omgezet.', 'success');
        restoreButton();
      },
      (error) => {
        const messages = {
          1: 'Locatietoegang is geweigerd. Geef de website locatietoestemming of vul het adres handmatig in.',
          2: 'De telefoon kon de huidige locatie niet bepalen.',
          3: 'Het bepalen van de locatie duurde te lang. Probeer het opnieuw.'
        };
        setMessage(messages[error.code] || 'Locatie bepalen is mislukt.');
        restoreButton();
      },
      {
        enableHighAccuracy: true,
        timeout: 12000,
        maximumAge: 30000
      }
    );
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
  startOdometer.addEventListener('input', calculateDistance);
  endOdometer.addEventListener('input', calculateDistance);
  form.addEventListener('submit', handleSubmit);
  document.getElementById('fillFromPrevious').addEventListener('click', fillPreviousOdometer);
  document.getElementById('resetForm').addEventListener('click', () => resetForm());
  document.getElementById('exportCsv').addEventListener('click', exportCsv);
  document.getElementById('clearAll').addEventListener('click', clearAll);

  document.querySelectorAll('[data-location-target]').forEach((button) => {
    button.addEventListener('click', () => requestLocation(button.dataset.locationTarget, button));
  });

  setupThemeToggle();
  render();

  if (latestRide()) startOdometer.value = latestRide().endOdometer;
})();
