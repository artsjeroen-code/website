(() => {
  'use strict';

  const DEFAULT_FULL_FROM = '2027-01-01';
  const DEFAULT_PRIVATE_LIMIT = 500;
  let fullRegistrationFrom = DEFAULT_FULL_FROM;
  let privateKmLimit = DEFAULT_PRIVATE_LIMIT;

  const rideDate = document.getElementById('rideDate');
  const rideType = document.getElementById('rideType');
  const vehicleSelect = document.getElementById('vehicleSelect');
  const privateKm = document.getElementById('privateKm');
  const fillPrevious = document.getElementById('fillFromPrevious');

  if (!rideDate || !rideType) return;

  function isFullRegistration(dateValue = rideDate.value) {
    return String(dateValue || '') >= fullRegistrationFrom;
  }

  function syncFormPolicy() {
    const full = isFullRegistration();
    const privateOption = rideType.querySelector('option[value="private"]');
    if (privateOption) {
      privateOption.hidden = !full;
      privateOption.disabled = !full;
    }
    if (!full && rideType.value !== 'business') rideType.value = 'business';
    if (fillPrevious) {
      fillPrevious.hidden = !full;
      fillPrevious.title = full
        ? 'Neem de vorige eindstand over'
        : 'Tot en met 2026 kunnen niet-geregistreerde privéritten tussen zakelijke ritten zitten';
    }
  }

  async function fetchJson(path) {
    const response = await fetch(path, {
      cache: 'no-store',
      credentials: 'same-origin',
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.json();
  }

  function derivedPrivateKm(rides, vehicles, year) {
    let total = 0;
    for (const vehicle of vehicles) {
      const vehicleRides = rides.filter((ride) =>
        ride.vehicleId === vehicle.id &&
        String(ride.date).startsWith(`${year}-`) &&
        ride.type === 'business'
      );
      if (!vehicleRides.length || !Number.isFinite(Number(vehicle.initialOdometer))) continue;

      const highest = Math.max(...vehicleRides.map((ride) => Number(ride.endOdometer)));
      const business = vehicleRides.reduce((sum, ride) => sum + Number(ride.distance || 0), 0);
      total += Math.max(0, highest - Number(vehicle.initialOdometer) - business);
    }
    return total;
  }

  async function refreshPrivateCounter() {
    if (!privateKm) return;
    const year = new Date().getFullYear();
    try {
      const [ridesPayload, vehiclesPayload] = await Promise.all([
        fetchJson('./api/rides'),
        fetchJson('./api/vehicles')
      ]);
      const rides = Array.isArray(ridesPayload.rides) ? ridesPayload.rides : [];
      const vehicles = Array.isArray(vehiclesPayload.vehicles) ? vehiclesPayload.vehicles : [];

      if (`${year}-01-01` < fullRegistrationFrom) {
        const total = derivedPrivateKm(rides, vehicles, year);
        privateKm.textContent = `${new Intl.NumberFormat('nl-NL').format(total)} km`;
        return;
      }

      const total = rides
        .filter((ride) => String(ride.date).startsWith(`${year}-`) && ride.type === 'private')
        .reduce((sum, ride) => sum + Number(ride.distance || 0), 0);
      privateKm.textContent = `${new Intl.NumberFormat('nl-NL').format(total)} / ${privateKmLimit} km`;
    } catch {
      // De normale app toont zijn eigen waarde als deze aanvullende beleidsweergave niet kan laden.
    }
  }

  async function loadPolicy() {
    try {
      const payload = await fetchJson('./api/policy');
      if (payload.fullRegistrationFrom) fullRegistrationFrom = String(payload.fullRegistrationFrom);
      if (Number.isFinite(Number(payload.privateKmLimit))) privateKmLimit = Number(payload.privateKmLimit);
    } catch {
      // Veilige defaults blijven actief.
    }
    syncFormPolicy();
    refreshPrivateCounter();
  }

  rideDate.addEventListener('change', syncFormPolicy);
  rideDate.addEventListener('input', syncFormPolicy);
  if (vehicleSelect) vehicleSelect.addEventListener('change', syncFormPolicy);

  syncFormPolicy();
  loadPolicy();
  window.addEventListener('load', () => {
    window.setTimeout(() => {
      syncFormPolicy();
      refreshPrivateCounter();
    }, 700);
  });
})();
