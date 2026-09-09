(() => {
  'use strict';

  const message = document.getElementById('authMessage');
  const loginButton = document.getElementById('passkeyLogin');
  const registerButton = document.getElementById('passkeyRegister');
  const statusText = document.getElementById('authStatus');
  const passwordForm = document.getElementById('passwordLoginForm');
  const passwordInput = document.getElementById('passwordInput');

  function setMessage(text, success = false) {
    message.textContent = text;
    message.classList.toggle('success', success);
  }

  function b64ToBytes(value) {
    const text = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
    const padded = text + '='.repeat((4 - (text.length % 4)) % 4);
    const binary = atob(padded);
    return Uint8Array.from(binary, (char) => char.charCodeAt(0));
  }

  function bytesToB64(value) {
    const bytes = new Uint8Array(value);
    let binary = '';
    bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
    return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
  }

  function requestOptionsFromJson(options) {
    const publicKey = { ...options.publicKey };
    publicKey.challenge = b64ToBytes(publicKey.challenge);
    if (Array.isArray(publicKey.allowCredentials)) {
      publicKey.allowCredentials = publicKey.allowCredentials.map((item) => ({
        ...item,
        id: b64ToBytes(item.id)
      }));
    }
    return publicKey;
  }

  function authenticationToJson(credential) {
    return {
      id: credential.id,
      rawId: bytesToB64(credential.rawId),
      type: credential.type,
      authenticatorAttachment: credential.authenticatorAttachment || undefined,
      clientExtensionResults: credential.getClientExtensionResults(),
      response: {
        clientDataJSON: bytesToB64(credential.response.clientDataJSON),
        authenticatorData: bytesToB64(credential.response.authenticatorData),
        signature: bytesToB64(credential.response.signature),
        userHandle: credential.response.userHandle
          ? bytesToB64(credential.response.userHandle)
          : null
      }
    };
  }

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
    let payload = {};
    try { payload = await response.json(); }
    catch { throw new Error(`Geen geldig antwoord (HTTP ${response.status})`); }
    if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
    return payload;
  }

  function redirectToApp() {
    window.location.replace('./');
  }

  async function login() {
    loginButton.disabled = true;
    registerButton.disabled = true;
    setMessage('Passkey wordt opgevraagd…');
    try {
      const begin = await api('login/begin', { method: 'POST' });
      const credential = await navigator.credentials.get({
        publicKey: requestOptionsFromJson(begin.options)
      });
      if (!credential) throw new Error('Geen passkey ontvangen');
      await api('login/complete', {
        method: 'POST',
        body: JSON.stringify({
          transaction: begin.transaction,
          credential: authenticationToJson(credential)
        })
      });
      setMessage('Inloggen gelukt.', true);
      redirectToApp();
    } catch (error) {
      setMessage(error.name === 'NotAllowedError'
        ? 'Passkey-aanmelding is geannuleerd of niet toegestaan.'
        : `Inloggen mislukt: ${error.message}`);
    } finally {
      loginButton.disabled = false;
      registerButton.disabled = false;
    }
  }

  async function passwordLogin(event) {
    event.preventDefault();
    const password = passwordInput.value;
    if (!password) return;

    const submitButton = passwordForm.querySelector('button[type="submit"]');
    submitButton.disabled = true;
    setMessage('Wachtwoord controleren…');
    try {
      await api('password', {
        method: 'POST',
        body: JSON.stringify({ password })
      });
      passwordInput.value = '';
      setMessage('Inloggen gelukt.', true);
      redirectToApp();
    } catch (error) {
      setMessage(`Inloggen mislukt: ${error.message}`);
      passwordInput.select();
    } finally {
      submitButton.disabled = false;
    }
  }

  function register() {
    window.location.href = './register.html';
  }

  async function init() {
    try {
      const status = await api('status');
      if (status.authenticated) {
        redirectToApp();
        return;
      }

      passwordForm.hidden = !status.passwordEnabled;

      const webauthnSupported = Boolean(window.PublicKeyCredential && navigator.credentials);
      if (!webauthnSupported) {
        statusText.textContent = status.passwordEnabled
          ? 'Passkeys worden door deze browser niet ondersteund. Gebruik je wachtwoord.'
          : 'Deze browser ondersteunt geen passkeys/WebAuthn.';
        loginButton.hidden = true;
        registerButton.hidden = true;
        return;
      }

      if (status.passkeyCount > 0) {
        statusText.textContent = `${status.passkeyCount} passkey${status.passkeyCount === 1 ? '' : 's'} geregistreerd.`;
        loginButton.hidden = false;
        registerButton.textContent = 'Extra passkey registreren';
      } else {
        statusText.textContent = status.passwordEnabled
          ? 'Nog geen passkey geregistreerd. Je kunt inloggen met je wachtwoord of eerst een passkey instellen.'
          : 'Nog geen passkey geregistreerd. Stel eerst Face ID, Touch ID of vingerafdruk in.';
        loginButton.hidden = true;
        registerButton.textContent = 'Eerste passkey instellen';
      }
      registerButton.hidden = false;
    } catch (error) {
      setMessage(`Authenticatieservice niet bereikbaar: ${error.message}`);
    }
  }

  loginButton.addEventListener('click', login);
  registerButton.addEventListener('click', register);
  passwordForm.addEventListener('submit', passwordLogin);
  init();
})();
