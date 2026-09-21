(() => {
  const form = document.getElementById('loginForm');
  const password = document.getElementById('password');
  const button = document.getElementById('loginButton');
  const message = document.getElementById('authMessage');

  async function authStatus() {
    const response = await fetch('./api/auth/status', {
      cache: 'no-store',
      credentials: 'same-origin'
    });
    if (!response.ok) throw new Error('Status kon niet worden opgehaald');
    return response.json();
  }

  async function redirectIfAuthenticated() {
    try {
      const status = await authStatus();
      if (status.authenticated) {
        window.location.replace('./');
        return true;
      }
    } catch {}
    return false;
  }

  form.addEventListener('submit', async event => {
    event.preventDefault();
    message.textContent = '';
    button.disabled = true;
    button.textContent = 'Inloggen…';

    try {
      const response = await fetch('./api/auth/login', {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password: password.value })
      });

      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error || 'Inloggen mislukt');
      }

      password.value = '';
      window.location.replace('./');
    } catch (error) {
      message.textContent = error.message || 'Inloggen mislukt';
      password.select();
    } finally {
      button.disabled = false;
      button.textContent = 'Inloggen';
    }
  });

  const params = new URLSearchParams(window.location.search);
  if (params.get('error') === 'connection') {
    message.textContent = 'De synchronisatieservice is momenteel niet bereikbaar.';
  }

  redirectIfAuthenticated();
})();
