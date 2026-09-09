(() => {
  'use strict';

  const message = document.getElementById('authMessage');
  const loginButton = document.getElementById('passkeyLogin');
  const registerButton = document.getElementById('passkeyRegister');
  const statusText = document.getElementById('authStatus');

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

  function creationOptionsFromJson(options) {
    const publicKey = { ...options.publicKey };
    publicKey.challenge = b64ToBytes(publicKey.challenge);
    publicKey.user = { ...publicKey.user, id: b64ToBytes(publicKey.user.id) };
    if (Array.isArray(publicKey.excludeCredentials)) {
      publicKey.excludeCredentials = publicKey.excludeCredentials.map((item) => ({
        ...item,
        id: b64ToBytes(item.id)
      }));
    }
    return publicKey;
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

  function registrationToJson(credential) {
    return {
      id: credential.id,
      rawId: bytesToB64(credential.rawId),
      type: credential.type,
      authenticatorAttachment: credential.authenticatorAttachment || undefined,
      clientExtensionResults: credential.getClientExtensionResults(),
      response: {
        clientDataJSON: bytesToB64(credential.response.clientDataJSON),
        attestationObject: bytesToB64(credential.response.attestationObject),
        transports: typeof credential.response.getTransports === 'function'
          ? credential.response.getTransports()
          : undefined
      }
    };
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

  async function register() {
    loginButton.disabled = true;
    registerButton.disabled = true;
    setMessage('Je wachtwoord wordt alleen nu gevraagd om een passkey veilig te registreren.');
    try {
      const begin = await api('register/begin', { method: 'POST' });
      const credential = await navigator.credentials.create({
        publicKey: creationOptionsFromJson(begin.options)
      });
      if (!credential) throw new Error('Geen passkey aangemaakt');
      await api('register/complete', {
        method: 'POST',
        body: JSON.stringify({
          transaction: begin.transaction,
          credential: registrationToJson(credential)
        })
      });
      setMessage('Passkey geregistreerd.', true);
      redirectToApp();
    } catch (error) {
      setMessage(error.name === 'NotAllowedError'
        ? 'Passkey-registratie is geannuleerd of niet toegestaan.'
        : `Registreren mislukt: ${error.message}`);
    } finally {
      loginButton.disabled = false;
      registerButton.disabled = false;
    }
  }

  async function init() {
    if (!window.PublicKeyCredential || !navigator.credentials) {
      statusText.textContent = 'Deze browser ondersteunt geen passkeys/WebAuthn.';
      loginButton.hidden = true;
      registerButton.hidden = true;
      return;
    }

    try {
      const status = await api('status');
      if (status.authenticated) {
        redirectToApp();
        return;
      }
      if (status.passkeyCount > 0) {
        statusText.textContent = `${status.passkeyCount} passkey${status.passkeyCount === 1 ? '' : 's'} geregistreerd.`;
        loginButton.hidden = false;
        registerButton.textContent = 'Extra passkey registreren';
      } else {
        statusText.textContent = 'Nog geen passkey geregistreerd. Stel eerst Face ID, Touch ID of vingerafdruk in.';
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
  init();
})();
