(() => {
  'use strict';

  const NOMINATIM_REVERSE_URL = 'https://nominatim.openstreetmap.org/reverse';

  function setFormMessage(message, success = false) {
    const element = document.getElementById('formMessage');
    if (!element) return;
    element.textContent = message;
    element.classList.toggle('success', success);
  }

  function metaForTarget(targetId) {
    return document.getElementById(targetId === 'departureAddress' ? 'departureMeta' : 'arrivalMeta');
  }

  function formatAddress(result) {
    const address = result && result.address ? result.address : {};
    const road = address.road || address.pedestrian || address.residential || address.cycleway || address.footway || '';
    const houseNumber = address.house_number || '';
    const postcode = address.postcode || '';
    const city = address.city || address.town || address.village || address.municipality || '';

    const street = [road, houseNumber].filter(Boolean).join(' ');
    const locality = [postcode, city].filter(Boolean).join(' ');
    const compact = [street, locality].filter(Boolean).join(', ');

    return compact || result.display_name || '';
  }

  async function reverseGeocode(latitude, longitude) {
    const params = new URLSearchParams({
      format: 'jsonv2',
      lat: String(latitude),
      lon: String(longitude),
      zoom: '18',
      addressdetails: '1',
      'accept-language': 'nl'
    });

    const response = await fetch(`${NOMINATIM_REVERSE_URL}?${params.toString()}`, {
      method: 'GET',
      headers: {
        Accept: 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Reverse geocoding gaf HTTP ${response.status}`);
    }

    const result = await response.json();
    const formatted = formatAddress(result);
    if (!formatted) throw new Error('Geen bruikbaar adres gevonden');
    return formatted;
  }

  function locateAndFill(targetId, button) {
    if (!navigator.geolocation) {
      setFormMessage('Deze browser ondersteunt geen locatiebepaling.');
      return;
    }

    const target = document.getElementById(targetId);
    const meta = metaForTarget(targetId);
    if (!target || !meta) return;

    const originalText = button.textContent;
    button.disabled = true;
    button.textContent = 'Locatie bepalen…';
    setFormMessage('');

    const restore = () => {
      button.disabled = false;
      button.textContent = originalText;
    };

    navigator.geolocation.getCurrentPosition(
      async (position) => {
        const { latitude, longitude, accuracy } = position.coords;
        const lat = Number(latitude.toFixed(6));
        const lon = Number(longitude.toFixed(6));

        try {
          button.textContent = 'Adres zoeken…';
          const address = await reverseGeocode(lat, lon);
          target.value = address;
          target.dataset.latitude = String(lat);
          target.dataset.longitude = String(lon);
          meta.textContent = `GPS ±${Math.round(accuracy)} m · adres via OpenStreetMap`;
          setFormMessage(`Adres gevonden: ${address}`, true);
        } catch (error) {
          console.warn('Adresomzetting mislukt.', error);
          target.value = `GPS ${lat}, ${lon}`;
          target.dataset.latitude = String(lat);
          target.dataset.longitude = String(lon);
          meta.textContent = `GPS ±${Math.round(accuracy)} m · adres kon niet automatisch worden bepaald`;
          setFormMessage('Locatie gevonden, maar het straatadres kon niet worden opgehaald. Je kunt het adres handmatig aanpassen.');
        } finally {
          restore();
        }
      },
      (error) => {
        const messages = {
          1: 'Locatietoegang is geweigerd. Geef de website locatietoestemming of vul het adres handmatig in.',
          2: 'De telefoon kon de huidige locatie niet bepalen.',
          3: 'Het bepalen van de locatie duurde te lang. Probeer het opnieuw.'
        };
        setFormMessage(messages[error.code] || 'Locatie bepalen is mislukt.');
        restore();
      },
      {
        enableHighAccuracy: true,
        timeout: 12000,
        maximumAge: 30000
      }
    );
  }

  document.querySelectorAll('[data-location-target]').forEach((button) => {
    button.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      locateAndFill(button.dataset.locationTarget, button);
    }, { capture: true });
  });
})();
