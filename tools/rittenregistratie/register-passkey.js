(() => {
  'use strict';

  const button = document.getElementById('registerPasskey');
  const message = document.getElementById('registerMessage');

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

  async function register() {
    if (!window.PublicKeyCredential || !navigator.credentials) {
      setMessage('Deze browser ondersteunt geen passkeys/WebAuthn.');
      return;
    }

    button.disabled = true;
    setMessage('Passkey wordt aangemaakt…');
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
      setMessage('Passkey geregistreerd. Je bent ingelogd.', true);
      window.setTimeout(() => window.location.replace('./'), 450);
    } catch (error) {
      setMessage(error.name === 'NotAllowedError'
        ? 'Passkey-registratie is geannuleerd of niet toegestaan.'
        : `Registreren mislukt: ${error.message}`);
    } finally {
      button.disabled = false;
    }
  }

  button.addEventListener('click', register);
})();
