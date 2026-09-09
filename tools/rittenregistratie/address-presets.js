(() => {
  'use strict';

  const NOMINATIM_SEARCH_URL = 'https://nominatim.openstreetmap.org/search';
  const PRESETS = [
    { label: 'Thuis', address: 'Olieslagerstraat 29, 5975 VP Sevenum' },
    { label: 'Station Horst-Sevenum', address: 'Stationsstraat 151, 5963 AA Hegelsom' },
    { label: 'Kantoor Roermond', address: 'Laurentiusplein 8, 6043 CS Roermond' },
    { label: 'Kantoor Eindhoven', address: 'Karel de Grotelaan 4, 5616 CA Eindhoven' }
  ];

  function metaFor(input) {
    return document.getElementById(input.id === 'departureAddress' ? 'departureMeta' : 'arrivalMeta');
  }

  function notifyFilled(input) {
    document.dispatchEvent(new CustomEvent('rittenregistratie:location-filled', {
      detail: { targetId: input.id }
    }));
  }

  async function geocode(address) {
    const params = new URLSearchParams({
      format: 'jsonv2',
      limit: '1',
      countrycodes: 'nl',
      q: address
    });
    const response = await fetch(`${NOMINATIM_SEARCH_URL}?${params}`, {
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const rows = await response.json();
    if (!Array.isArray(rows) || !rows.length) throw new Error('Adres niet gevonden');
    const lat = Number(rows[0].lat);
    const lon = Number(rows[0].lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) throw new Error('Geen geldige coördinaten');
    return { lat, lon };
  }

  async function applyPreset(input, preset, button) {
    input.value = preset.address;
    delete input.dataset.capturedTime;
    delete input.dataset.latitude;
    delete input.dataset.longitude;

    const meta = metaFor(input);
    const original = button.textContent;
    button.disabled = true;
    button.textContent = '…';

    try {
      const coords = await geocode(preset.address);
      input.dataset.latitude = String(coords.lat);
      input.dataset.longitude = String(coords.lon);
      if (meta) meta.textContent = `Sneladres · ${preset.label} · locatie via OpenStreetMap`;
    } catch (error) {
      console.warn(`Sneladres kon niet worden gegeocodeerd: ${preset.label}`, error);
      if (meta) meta.textContent = `Sneladres · ${preset.label} · adres ingevuld, GPS niet gevonden`;
    } finally {
      button.disabled = false;
      button.textContent = original;
      input.dispatchEvent(new Event('change', { bubbles: true }));
      notifyFilled(input);
    }
  }

  function addPresets(input) {
    const block = input.closest('.address-block');
    if (!block || block.querySelector('.address-presets')) return;

    const row = document.createElement('div');
    row.className = 'address-presets';
    row.setAttribute('aria-label', `Sneladressen voor ${input.id === 'departureAddress' ? 'vertrek' : 'aankomst'}`);

    PRESETS.forEach((preset) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'address-preset-button';
      button.textContent = preset.label;
      button.title = preset.address;
      button.addEventListener('click', () => applyPreset(input, preset, button));
      row.appendChild(button);
    });

    const meta = metaFor(input);
    if (meta) block.insertBefore(row, meta);
    else block.appendChild(row);
  }

  ['departureAddress', 'arrivalAddress'].forEach((id) => {
    const input = document.getElementById(id);
    if (input) addPresets(input);
  });
})();
