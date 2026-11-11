"use strict";

(() => {
  const sendButton = document.getElementById("railClipboardSend");
  const copyButton = document.getElementById("railClipboardCopy");
  const status = document.getElementById("railClipboardStatus");
  if (!sendButton || !copyButton || !status) return;

  const API = "./tools/klembord/api";
  const REFRESH_MS = 5000;
  const MAX_CACHE_AGE_MS = 8000;
  const MAX_MESSAGE_BYTES = 16384;
  const encoder = new TextEncoder();

  let latestMessage = null;
  let lastChecked = 0;
  let loading = null;
  let authenticated = false;
  let statusTimer = null;

  function feedback(message) {
    status.textContent = message;
    status.classList.add("visible");
    if (statusTimer) clearTimeout(statusTimer);
    statusTimer = setTimeout(() => {
      status.classList.remove("visible");
    }, 4500);
  }

  async function api(path, options) {
    const response = await fetch(API + path, {
      credentials: "same-origin",
      cache: "no-store",
      ...options
    });

    if (response.status === 401) {
      authenticated = false;
      latestMessage = null;
      throw new Error("Log eerst in via Beam Me Up.");
    }
    if (!response.ok) {
      throw new Error(response.status === 429
        ? "Te veel verzoeken. Probeer het later opnieuw."
        : "Beam Me Up is tijdelijk niet beschikbaar.");
    }
    return response.json();
  }

  function updateButtons() {
    copyButton.disabled = !authenticated || !latestMessage;
    copyButton.title = latestMessage
      ? "Nieuwste bericht naar systeemklembord kopiëren"
      : "Geen bericht beschikbaar (log zo nodig in bij Beam Me Up)";
  }

  async function updateLatest() {
    if (loading) return loading;
    loading = (async () => {
      try {
        const response = await api("/messages");
        authenticated = true;
        latestMessage = Array.isArray(response.messages) && response.messages.length
          ? response.messages[0]
          : null;
        lastChecked = Date.now();
        updateButtons();
      } catch {
        lastChecked = 0;
        latestMessage = null;
        updateButtons();
      }
    })();

    try { await loading; } finally { loading = null; }
  }

  sendButton.addEventListener("click", async () => {
    if (!navigator.clipboard?.readText) {
      feedback("Deze browser ondersteunt het lezen van het systeemklembord niet.");
      return;
    }

    // Het systeemklembord wordt uitsluitend gelezen na een expliciete klik.
    let content;
    try {
      content = await navigator.clipboard.readText();
    } catch {
      feedback("Geen klembordtoegang. Sta klembordlezen toe in je browser.");
      return;
    }

    if (!content || !content.trim()) {
      feedback("Het systeemklembord bevat geen tekst.");
      return;
    }
    if (encoder.encode(content).length > MAX_MESSAGE_BYTES) {
      feedback("Tekst te groot (maximum 16 KB).");
      return;
    }

    sendButton.disabled = true;
    try {
      await api("/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ content })
      });
      feedback("Klembordtekst verstuurd.");
      await updateLatest();
    } catch (error) {
      feedback(error.message);
    } finally {
      sendButton.disabled = false;
    }
  });

  copyButton.addEventListener("click", () => {
    if (!latestMessage || !authenticated) {
      feedback("Geen bericht beschikbaar. Log zo nodig in via Beam Me Up.");
      return;
    }

    // Niet na een asynchronous fetch schrijven: sommige browsers verliezen
    // anders de gebruikersactivatie voor clipboard.writeText().
    if (Date.now() - lastChecked > MAX_CACHE_AGE_MS) {
      feedback("Berichten worden bijgewerkt. Klik nogmaals op kopiëren.");
      void updateLatest();
      return;
    }

    if (!navigator.clipboard?.writeText) {
      feedback("Deze browser ondersteunt schrijven naar het systeemklembord niet.");
      return;
    }

    const text = latestMessage.content;
    navigator.clipboard.writeText(text)
      .then(() => feedback("Nieuwste bericht gekopieerd naar het systeemklembord."))
      .catch(() => feedback("Kopiëren geblokkeerd. Controleer de browserrechten."));
  });

  updateButtons();
  void updateLatest();

  setInterval(() => {
    if (!document.hidden) void updateLatest();
  }, REFRESH_MS);

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) void updateLatest();
  });
})();
