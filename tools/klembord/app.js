
"use strict";

const API = "./api";

const $ = (id) => document.getElementById(id);

const loginPanel = $("loginPanel");
const clipboardPanel = $("clipboardPanel");
const logoutBtn = $("logoutBtn");
const messagesList = $("messagesList");
const statusMessage = $("statusMessage");

let authenticated = false;

function setStatus(message) {
  statusMessage.textContent = message;
}

function showAuthenticated(value) {
  authenticated = value;

  loginPanel.hidden = value;
  clipboardPanel.hidden = !value;
  logoutBtn.hidden = !value;

  if (!value) {
    messagesList.replaceChildren();
    $("password").value = "";
    $("message").value = "";
  }
}

async function apiRequest(path, options = {}) {
  const response = await fetch(API + path, {
    credentials: "same-origin",
    cache: "no-store",
    ...options
  });

  let data;

  try {
    data = await response.json();
  } catch {
    throw new Error("Ongeldig serverantwoord");
  }

  if (!response.ok) {
    if (response.status === 401 && path !== "/auth/login") {
      showAuthenticated(false);
    }

    throw new Error(data.error || "Verzoek mislukt");
  }

  return data;
}

async function checkSession() {
  try {
    const result = await apiRequest("/auth/status");

    showAuthenticated(result.authenticated === true);

    if (authenticated) {
      await loadMessages();
    }
  } catch {
    showAuthenticated(false);
    setStatus("Kan geen verbinding maken met de server.");
  }
}

async function login(event) {
  event.preventDefault();

  const password = $("password").value;
  const button = $("loginForm").querySelector("button");

  button.disabled = true;

  try {
    await apiRequest("/auth/login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ password })
    });

    showAuthenticated(true);
    setStatus("Ingelogd.");

    await loadMessages();
  } catch (error) {
    setStatus("Inloggen mislukt: " + error.message);
  } finally {
    button.disabled = false;
  }
}

async function logout() {
  try {
    await apiRequest("/auth/logout", {
      method: "POST"
    });

    showAuthenticated(false);
    setStatus("Uitgelogd.");
  } catch (error) {
    setStatus("Uitloggen mislukt: " + error.message);
  }
}

async function sendMessage(event) {
  event.preventDefault();

  const content = $("message").value;

  if (!content.trim()) {
    setStatus("Vul eerst een bericht in.");
    return;
  }

  const button = $("messageForm").querySelector("button");
  button.disabled = true;

  try {
    await apiRequest("/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ content })
    });

    $("message").value = "";
    await loadMessages();

    setStatus("Bericht opgeslagen.");
  } catch (error) {
    setStatus("Versturen mislukt: " + error.message);
  } finally {
    button.disabled = false;
  }
}

function createMessageCard(message) {
  const card = document.createElement("article");
  card.className = "message-card";

  const content = document.createElement("div");
  content.className = "message-content";

  // Gebruik textContent, nooit innerHTML voor berichtinhoud.
  content.textContent = message.content;

  const meta = document.createElement("p");
  meta.className = "message-meta";

  const date = new Date(message.created_at * 1000);
  meta.textContent = date.toLocaleString("nl-NL");

  const actions = document.createElement("div");
  actions.className = "message-actions";

  const copyButton = document.createElement("button");
  copyButton.type = "button";
  copyButton.textContent = "Kopiëren";

  copyButton.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(message.content);
      setStatus("Tekst gekopieerd.");
    } catch {
      setStatus("Kopiëren mislukt. Controleer de browserrechten.");
    }
  });


  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.textContent = "Verwijderen";

  deleteButton.addEventListener("click", async () => {
    const confirmed = window.confirm(
      "Weet je zeker dat je dit bericht wilt verwijderen?"
    );

    if (!confirmed) return;

    deleteButton.disabled = true;

    try {
      await apiRequest(
        "/messages/" + encodeURIComponent(message.id),
        {
          method: "DELETE"
        }
      );

      await loadMessages();
      setStatus("Bericht verwijderd.");
    } catch (error) {
      setStatus(
        "Verwijderen mislukt: " + error.message
      );

      deleteButton.disabled = false;
    }
  });

  actions.append(copyButton, deleteButton);
  card.append(meta, content, actions);


  return card;
}

async function loadMessages() {
  if (!authenticated) return;

  try {
    const result = await apiRequest("/messages");

    messagesList.replaceChildren();

    if (!result.messages.length) {
      const empty = document.createElement("p");
      empty.className = "empty-state";
      empty.textContent = "Geen berichten beschikbaar.";
      messagesList.append(empty);
      return;
    }

    for (const message of result.messages) {
      messagesList.append(createMessageCard(message));
    }
  } catch (error) {
    setStatus("Berichten ophalen mislukt: " + error.message);
  }
}

$("loginForm").addEventListener("submit", login);
$("messageForm").addEventListener("submit", sendMessage);
logoutBtn.addEventListener("click", logout);
$("refreshBtn").addEventListener("click", loadMessages);

$("railTheme").addEventListener("click", () => {
  const root = document.documentElement;
  const next = root.dataset.theme === "dark" ? "light" : "dark";
  root.dataset.theme = next;
  localStorage.setItem("startpagina-theme", next);
});

checkSession();
