(() => {
  'use strict';

  const form = document.getElementById('passwordLoginForm');
  const passwordInput = document.getElementById('passwordInput');
  const statusText = document.getElementById('authStatus');
  const message = document.getElementById('authMessage');

  async function api(path, options = {}) {
    const response = await fetch(`./api/auth/${path}`, {
      cache: 'no-store',
      credentials: 'same-origin',
      headers: {
        Accept: 'application/json',
        ...(options.body ? { 'Content-Type': 'application/json' } : {})
      },
      ...options
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
    return payload;
  }

  async function init() {
    try {
      const status = await api('status');
      if (status.authenticated) {
        window.location.replace('./');
        return;
      }
      if (!status.passwordEnabled) {
        statusText.textContent = 'Wachtwoord-login is nog niet ingesteld op de server.';
        message.textContent = 'Stel eerst RITTEN_AUTH_PASSWORD_HASH in op de Raspberry Pi. De bestaande authenticatie is niet gewijzigd.';
        return;
      }
      statusText.textContent = 'Log in met je wachtwoord.';
      form.hidden = false;
    } catch (error) {
      statusText.textContent = 'Inloggen niet beschikbaar.';
      message.textContent = `Authenticatieservice niet bereikbaar: ${error.message}`;
    }
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const password = passwordInput.value;
    if (!password) return;
    const submitButton = form.querySelector('button[type="submit"]');
    submitButton.disabled = true;
    message.textContent = 'Wachtwoord controleren…';
    try {
      await api('password', {
        method: 'POST',
        body: JSON.stringify({ password })
      });
      passwordInput.value = '';
      window.location.replace('./');
    } catch (error) {
      message.textContent = `Inloggen mislukt: ${error.message}`;
      passwordInput.select();
    } finally {
      submitButton.disabled = false;
    }
  });

  init();
})();
