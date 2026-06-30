const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const STORAGE_KEY = "school-mail-dashboard-settings";
const SCOPES = ["User.Read", "Mail.Read", "Mail.ReadWrite"];
const LOCAL_HOSTNAMES = new Set(["localhost", "127.0.0.1", "::1", ""]);
const IS_HOSTED_DEMO = !LOCAL_HOSTNAMES.has(window.location.hostname);

const priorities = [
  {
    id: "urgent",
    label: "Urgent",
    tone: "tone-urgent",
    hint: "Deadlines, high importance, flagged, or time-sensitive messages",
  },
  {
    id: "high",
    label: "High",
    tone: "tone-high",
    hint: "Teacher messages, tasks, meetings, and academic action items",
  },
  {
    id: "normal",
    label: "Normal",
    tone: "tone-normal",
    hint: "Useful messages that may need a read or light follow-up",
  },
  {
    id: "low",
    label: "Low",
    tone: "tone-low",
    hint: "Newsletters, notices, automated messages, and FYI items",
  },
];

const demoMessages = [
  {
    id: "demo-1",
    subject: "Coursework deadline reminder",
    from: { emailAddress: { name: "Mr Carter", address: "teacher@school.example" } },
    receivedDateTime: new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString(),
    bodyPreview: "Please submit your coursework draft by Friday afternoon. Bring questions to tomorrow's lesson.",
    bodyText: "Please submit your coursework draft by Friday afternoon. Bring questions to tomorrow's lesson.",
    importance: "high",
    isRead: false,
    flag: { flagStatus: "flagged" },
    webLink: "",
  },
  {
    id: "demo-2",
    subject: "History society trip form",
    from: { emailAddress: { name: "Trips Office", address: "trips@school.example" } },
    receivedDateTime: new Date(Date.now() - 3 * 60 * 60 * 1000).toISOString(),
    bodyPreview: "Please complete the attached consent form before next Wednesday.",
    bodyText: "Please complete the attached consent form before next Wednesday.",
    importance: "normal",
    isRead: true,
    flag: { flagStatus: "notFlagged" },
    webLink: "",
  },
  {
    id: "demo-3",
    subject: "Weekly bulletin",
    from: { emailAddress: { name: "School Office", address: "office@school.example" } },
    receivedDateTime: new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString(),
    bodyPreview: "This week's notices include sports fixtures, music rehearsals, and library opening hours.",
    bodyText: "This week's notices include sports fixtures, music rehearsals, and library opening hours.",
    importance: "low",
    isRead: true,
    flag: { flagStatus: "notFlagged" },
    webLink: "",
  },
  {
    id: "demo-4",
    subject: "Can we meet after class?",
    from: { emailAddress: { name: "Ms Ahmed", address: "teacher@school.example" } },
    receivedDateTime: new Date(Date.now() - 45 * 60 * 1000).toISOString(),
    bodyPreview: "I would like to quickly talk through your essay structure after period 5 today.",
    bodyText: "I would like to quickly talk through your essay structure after period 5 today.",
    importance: "normal",
    isRead: false,
    flag: { flagStatus: "notFlagged" },
    webLink: "",
  },
];

const state = {
  settings: {
    clientId: "",
    tenant: "organizations",
    outlookMailbox: "",
    openaiKey: "",
    openaiModel: "gpt-4.1-mini",
  },
  msalInstance: null,
  account: null,
  messages: [],
  filteredMessages: [],
  search: "",
  mode: "demo",
  consentMode: "ask",
  serverAiConfigured: false,
  hostedDemo: IS_HOSTED_DEMO,
  latestDraft: null,
};

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  bindElements();
  loadSettings();
  bindEvents();
  configureHostedDemo();
  if (!state.hostedDemo) {
    initialiseAuth();
  }
  loadDemo();
  if (!state.hostedDemo) {
    loadServerStatus();
  }
  render();
  refreshIcons();
});

function bindElements() {
  Object.assign(els, {
    accountLabel: document.getElementById("account-label"),
    aiStatus: document.getElementById("ai-status"),
    clearSettings: document.getElementById("clear-settings-button"),
    clientId: document.getElementById("client-id-input"),
    connectionStatus: document.getElementById("connection-status"),
    consentDialog: document.getElementById("consent-dialog"),
    consentText: document.getElementById("consent-text"),
    demoButton: document.getElementById("demo-button"),
    dialog: document.getElementById("mail-dialog"),
    dialogBody: document.getElementById("dialog-body"),
    dialogClose: document.getElementById("dialog-close"),
    dialogCopy: document.getElementById("dialog-copy"),
    dialogDraft: document.getElementById("dialog-draft"),
    dialogSubtitle: document.getElementById("dialog-subtitle"),
    dialogTitle: document.getElementById("dialog-title"),
    hostedBanner: document.getElementById("hosted-banner"),
    lastUpdated: document.getElementById("last-updated"),
    messageCount: document.getElementById("message-count"),
    openaiKey: document.getElementById("openai-key-input"),
    openaiModel: document.getElementById("openai-model-input"),
    outlookMailbox: document.getElementById("outlook-mailbox-input"),
    outlookDesktop: document.getElementById("outlook-desktop-button"),
    priorityList: document.getElementById("priority-list"),
    refresh: document.getElementById("refresh-button"),
    search: document.getElementById("search-input"),
    settingsForm: document.getElementById("settings-form"),
    signin: document.getElementById("signin-button"),
    signout: document.getElementById("signout-button"),
    tenant: document.getElementById("tenant-input"),
    toast: document.getElementById("toast"),
  });
}

function bindEvents() {
  document.querySelectorAll("[data-view]").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.view));
  });

  document.querySelectorAll("[data-consent-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      state.consentMode = button.dataset.consentMode;
      document.querySelectorAll("[data-consent-mode]").forEach((item) => item.classList.toggle("active", item === button));
      showToast(state.consentMode === "local" ? "AI calls are paused. Local helpers will be used." : "AI approval will be requested for each email.");
    });
  });

  on(els.settingsForm, "submit", (event) => {
    event.preventDefault();
    state.settings = {
      clientId: valueOf(els.clientId),
      tenant: valueOf(els.tenant) || "organizations",
      outlookMailbox: state.hostedDemo ? "" : valueOf(els.outlookMailbox),
      openaiKey: state.hostedDemo ? "" : valueOf(els.openaiKey),
      openaiModel: valueOf(els.openaiModel) || "gpt-4.1-mini",
    };
    saveSettings();
    if (!state.hostedDemo) {
      initialiseAuth();
    }
    render();
    showToast("Settings saved.");
  });

  on(els.clearSettings, "click", () => {
    localStorage.removeItem(STORAGE_KEY);
    state.settings = { clientId: "", tenant: "organizations", outlookMailbox: "", openaiKey: "", openaiModel: "gpt-4.1-mini" };
    state.msalInstance = null;
    state.account = null;
    setSettingsForm();
    loadDemo();
    render();
    showToast("Settings cleared.");
  });

  on(els.signin, "click", signIn);
  on(els.signout, "click", signOut);
  on(els.demoButton, "click", () => {
    loadDemo();
    render();
    showToast("Demo inbox loaded.");
  });
  on(els.outlookDesktop, "click", loadOutlookDesktop);
  on(els.refresh, "click", refreshInbox);
  on(els.search, "input", () => {
    state.search = valueOf(els.search).toLowerCase();
    renderPriorityList();
  });
  on(els.dialogClose, "click", () => els.dialog.close());
  on(els.dialogCopy, "click", copyDialogText);
  on(els.dialogDraft, "click", createDraftFromDialog);
}

function on(element, eventName, handler) {
  if (element) {
    element.addEventListener(eventName, handler);
  }
}

function valueOf(element) {
  return element?.value?.trim() || "";
}

function configureHostedDemo() {
  if (!state.hostedDemo) {
    return;
  }

  els.hostedBanner?.classList.remove("hidden");
  document.querySelectorAll("[data-local-only]").forEach((element) => {
    element.classList.add("hidden");
    element.querySelectorAll?.("input, button, textarea, select").forEach((control) => {
      control.disabled = true;
    });
  });
  state.consentMode = "local";
  document.querySelectorAll("[data-consent-mode]").forEach((button) => {
    const isLocal = button.dataset.consentMode === "local";
    button.classList.toggle("active", isLocal);
    button.disabled = !isLocal;
  });
}

function loadSettings() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      state.settings = { ...state.settings, ...JSON.parse(saved) };
    } catch {
      localStorage.removeItem(STORAGE_KEY);
    }
  }
  setSettingsForm();
}

function saveSettings() {
  const { openaiKey, ...safeSettings } = state.settings;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(safeSettings));
}

function setSettingsForm() {
  if (els.clientId) els.clientId.value = state.settings.clientId;
  if (els.tenant) els.tenant.value = state.settings.tenant;
  if (els.outlookMailbox) els.outlookMailbox.value = state.hostedDemo ? "" : state.settings.outlookMailbox;
  if (els.openaiKey) els.openaiKey.value = state.hostedDemo ? "" : state.settings.openaiKey;
  if (els.openaiModel) els.openaiModel.value = state.settings.openaiModel;
}

function initialiseAuth() {
  if (!state.settings.clientId || !window.msal) {
    return;
  }

  state.msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: state.settings.clientId,
      authority: `https://login.microsoftonline.com/${state.settings.tenant || "organizations"}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  });

  const accounts = state.msalInstance.getAllAccounts();
  state.account = accounts[0] || null;
  if (state.account) {
    state.mode = "graph";
  }
}

async function signIn() {
  if (state.hostedDemo) {
    showToast("Hosted demo only. Run locally to sign in.");
    return;
  }

  if (!state.settings.clientId) {
    switchView("settings");
    showToast("Add your Microsoft client ID first.");
    return;
  }

  if (!state.msalInstance) {
    initialiseAuth();
  }

  try {
    const response = await state.msalInstance.loginPopup({ scopes: SCOPES, prompt: "select_account" });
    state.account = response.account;
    state.mode = "graph";
    await refreshInbox();
  } catch (error) {
    showToast(readableError(error));
  }
  render();
}

async function signOut() {
  if (state.msalInstance && state.account) {
    try {
      await state.msalInstance.logoutPopup({ account: state.account });
    } catch {
      await state.msalInstance.logoutRedirect({ account: state.account });
    }
  }
  state.account = null;
  loadDemo();
  render();
}

async function refreshInbox() {
  if (state.hostedDemo) {
    loadDemo();
    render();
    return;
  }

  if (state.mode === "outlook-desktop") {
    await loadOutlookDesktop();
    return;
  }

  if (!state.account || !state.msalInstance) {
    loadDemo();
    render();
    return;
  }

  setBusy(true);
  try {
    const encodedSelect = encodeURIComponent("id,subject,from,receivedDateTime,bodyPreview,importance,isRead,flag,conversationId,webLink,categories,body");
    const data = await graphFetch(`/me/mailFolders/inbox/messages?$top=40&$orderby=receivedDateTime desc&$select=${encodedSelect}`);
    state.messages = (data.value || []).map(normaliseMessage);
    state.mode = "graph";
    els.lastUpdated.textContent = `Updated ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`;
  } catch (error) {
    showToast(readableError(error));
  } finally {
    setBusy(false);
    render();
  }
}

async function loadOutlookDesktop() {
  if (state.hostedDemo) {
    showToast("Hosted demo only. Run locally for Outlook Desktop.");
    return;
  }

  if (!state.settings.outlookMailbox) {
    showToast("Add an Outlook Desktop mailbox in Settings first.");
    switchView("settings");
    return;
  }

  setBusy(true);
  try {
    const mailbox = encodeURIComponent(state.settings.outlookMailbox);
    const response = await fetch(`/api/outlook/messages?mailbox=${mailbox}`);
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Could not read Outlook Desktop inbox");
    }
    state.account = null;
    state.mode = "outlook-desktop";
    state.messages = (data.value || []).map(normaliseMessage);
    els.lastUpdated.textContent = `Outlook Desktop · ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`;
    showToast("Loaded from Outlook Desktop.");
  } catch (error) {
    showToast(readableError(error));
  } finally {
    setBusy(false);
    render();
  }
}

async function loadServerStatus() {
  if (state.hostedDemo) {
    state.serverAiConfigured = false;
    return;
  }

  try {
    const response = await fetch("/api/status");
    if (!response.ok) return;
    const data = await response.json();
    state.serverAiConfigured = Boolean(data.aiConfigured);
    render();
  } catch {
    state.serverAiConfigured = false;
  }
}

async function getAccessToken() {
  const request = { scopes: SCOPES, account: state.account };
  try {
    const response = await state.msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch {
    const response = await state.msalInstance.acquireTokenPopup(request);
    state.account = response.account;
    return response.accessToken;
  }
}

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  headers.set("Accept", "application/json");
  headers.set("Prefer", 'outlook.body-content-type="text"');
  if (options.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(`${GRAPH_ROOT}${path}`, { ...options, headers });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Microsoft Graph returned ${response.status}: ${detail.slice(0, 220)}`);
  }

  if (response.status === 204) {
    return null;
  }
  return response.json();
}

function loadDemo() {
  state.mode = "demo";
  state.messages = demoMessages.map(normaliseMessage);
  els.lastUpdated.textContent = "Demo inbox";
}

function normaliseMessage(message) {
  const bodyText = stripHtml(message.body?.content || message.bodyText || message.bodyPreview || "");
  const normalised = {
    ...message,
    bodyText,
    senderName: message.from?.emailAddress?.name || "Unknown sender",
    senderAddress: message.from?.emailAddress?.address || "",
  };
  return { ...normalised, priority: classifyMessage(normalised) };
}

function classifyMessage(message) {
  const text = `${message.subject || ""} ${message.senderName || ""} ${message.senderAddress || ""} ${message.bodyPreview || ""}`.toLowerCase();
  const isFlagged = message.flag?.flagStatus === "flagged";
  const isHighImportance = message.importance === "high";
  const urgentWords = ["urgent", "deadline", "today", "tomorrow", "asap", "overdue", "due by", "due today", "final reminder", "immediately"];
  const highWords = ["homework", "coursework", "assignment", "assessment", "exam", "test", "meeting", "teacher", "lesson", "feedback", "consent form", "bring"];
  const lowWords = ["newsletter", "bulletin", "digest", "notice", "no-reply", "noreply", "unsubscribe", "library opening", "fixture"];

  if (isFlagged || isHighImportance || containsAnyKeyword(text, urgentWords)) {
    return "urgent";
  }
  if (containsAnyKeyword(text, highWords)) {
    return "high";
  }
  if (containsAnyKeyword(text, lowWords)) {
    return "low";
  }
  return "normal";
}

function containsAnyKeyword(text, keywords) {
  return keywords.some((keyword) => {
    const escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (keyword.includes(" ") || keyword.includes("-")) {
      return text.includes(keyword);
    }
    return new RegExp(`(^|[^a-z])${escaped}s?([^a-z]|$)`, "i").test(text);
  });
}

function render() {
  const connected = Boolean(state.account);
  els.accountLabel.textContent = state.hostedDemo ? "Hosted demo" : connected ? state.account.username : state.mode === "outlook-desktop" ? "Outlook Desktop" : "Demo workspace";
  els.connectionStatus.textContent = state.hostedDemo ? "Demo" : connected ? "Graph" : state.mode === "outlook-desktop" ? "Desktop" : "Demo";
  els.messageCount.textContent = String(state.messages.length);
  els.aiStatus.textContent = (state.settings.openaiKey || state.serverAiConfigured) && state.consentMode !== "local" ? "OpenAI" : "Local";
  els.signin?.classList.toggle("hidden", connected || state.hostedDemo);
  els.signout?.classList.toggle("hidden", !connected || state.hostedDemo);
  renderPriorityList();
  refreshIcons();
}

function renderPriorityList() {
  const query = state.search;
  state.filteredMessages = state.messages.filter((message) => {
    if (!query) return true;
    return `${message.senderName} ${message.senderAddress} ${message.subject} ${message.bodyPreview}`.toLowerCase().includes(query);
  });

  els.priorityList.innerHTML = priorities
    .map((priority) => {
      const items = state.filteredMessages.filter((message) => message.priority === priority.id);
      const open = priority.id === "urgent" || priority.id === "high" ? "open" : "";
      return `
        <details class="priority-group" ${open}>
          <summary>
            <span class="priority-tone ${priority.tone}" aria-hidden="true"></span>
            <span class="priority-title">
              <strong>${escapeHtml(priority.label)}</strong>
              <span>${escapeHtml(priority.hint)}</span>
            </span>
            <span class="priority-count">${items.length}</span>
          </summary>
          <div class="message-list">
            ${items.length ? items.map(renderMessage).join("") : `<div class="empty-state">No messages here.</div>`}
          </div>
        </details>
      `;
    })
    .join("");

  els.priorityList.querySelectorAll("[data-action]").forEach((button) => {
    button.addEventListener("click", () => handleMailAction(button.dataset.action, button.dataset.id));
  });
  refreshIcons();
}

function renderMessage(message) {
  const date = formatDate(message.receivedDateTime);
  const unread = message.isRead ? "" : "unread";
  const flagged = message.flag?.flagStatus === "flagged" ? `<span class="tag">Flagged</span>` : "";
  const importance = message.importance && message.importance !== "normal" ? `<span class="tag">${escapeHtml(message.importance)}</span>` : "";
  const link = message.webLink ? `<a class="tag" href="${escapeAttribute(message.webLink)}" target="_blank" rel="noreferrer">Outlook</a>` : "";

  return `
    <article class="email-row ${unread}">
      <div class="email-main">
        <div class="email-meta">
          <span class="email-sender">${escapeHtml(message.senderName)}</span>
          <span>${escapeHtml(date)}</span>
          <span class="email-tags">${flagged}${importance}${link}</span>
        </div>
        <div class="email-subject">${escapeHtml(message.subject || "(No subject)")}</div>
        <div class="email-preview">${escapeHtml(message.bodyPreview || message.bodyText || "")}</div>
      </div>
      <div class="email-actions">
        <button class="mail-action" type="button" data-action="reply" data-id="${escapeAttribute(message.id)}" title="Reply draft">
          <i data-lucide="reply"></i><span>Reply</span>
        </button>
        <button class="mail-action" type="button" data-action="summary" data-id="${escapeAttribute(message.id)}" title="Summarise">
          <i data-lucide="list-checks"></i><span>Summarise</span>
        </button>
        <button class="mail-action" type="button" data-action="alternate" data-id="${escapeAttribute(message.id)}" title="Alternate reply">
          <i data-lucide="sparkles"></i><span>Alternate</span>
        </button>
      </div>
    </article>
  `;
}

async function handleMailAction(action, id) {
  const message = state.messages.find((item) => item.id === id);
  if (!message) return;

  try {
    const requiresAi = action === "summary" || action === "reply" || action === "alternate";
    if (requiresAi && state.settings.openaiKey && state.consentMode === "ask") {
      const approved = await requestAiConsent(action, message);
      if (!approved) return;
    }

    setBusy(true);
    const content = await generateAssistantText(action, message);
    showMailDialog(action, message, content);
  } catch (error) {
    showToast(readableError(error));
  } finally {
    setBusy(false);
  }
}

function requestAiConsent(action, message) {
  return new Promise((resolve) => {
    els.consentText.textContent = `${actionLabel(action)} will send this email's sender, subject, preview, and body text to OpenAI. Nothing will be sent if you cancel.`;
    const handleClose = () => {
      els.consentDialog.removeEventListener("close", handleClose);
      resolve(els.consentDialog.returnValue === "approve");
    };
    els.consentDialog.addEventListener("close", handleClose);
    els.consentDialog.showModal();
  });
}

async function generateAssistantText(action, message) {
  if (state.hostedDemo) {
    return localAssistantText(action, message);
  }

  if ((state.settings.openaiKey || state.serverAiConfigured) && state.consentMode !== "local") {
    return callOpenAi(action, message);
  }
  return localAssistantText(action, message);
}

async function callOpenAi(action, message) {
  if (state.hostedDemo) {
    return localAssistantText(action, message);
  }

  const response = await fetch("/api/ai", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      action,
      apiKey: state.settings.openaiKey || undefined,
      model: state.settings.openaiModel || "gpt-4.1-mini",
      message: {
        senderName: message.senderName,
        senderAddress: message.senderAddress,
        subject: message.subject,
        bodyPreview: message.bodyPreview,
        bodyText: message.bodyText,
      },
    }),
  });

  if (!response.ok) {
    const detail = await response.json().catch(() => ({}));
    throw new Error(detail.error || `AI request failed with ${response.status}`);
  }

  const data = await response.json();
  return data.text || localAssistantText(action, message);
}

function localAssistantText(action, message) {
  const sender = message.senderName || "there";
  const subject = message.subject || "your email";
  const preview = message.bodyPreview || message.bodyText || "No preview available.";
  const deadline = findDeadline(`${subject} ${preview} ${message.bodyText}`);

  if (action === "summary") {
    return [
      `- ${subject}`,
      `- Main point: ${preview}`,
      deadline ? `- Deadline or timing: ${deadline}` : "- Deadline or timing: none spotted",
      `- Suggested next step: ${message.priority === "low" ? "Read when convenient." : "Review and respond if action is needed."}`,
    ].join("\n");
  }

  if (action === "alternate") {
    return [
      `Dear ${sender},`,
      "",
      "Thank you for your email. I have seen this and will review the details carefully.",
      deadline ? `I have noted the timing around ${deadline}.` : "I will let you know if I have any questions.",
      "",
      "Kind regards,",
    ].join("\n");
  }

  return [
    `Hi ${sender},`,
    "",
    "Thanks for your email. I have seen this and will follow up as needed.",
    deadline ? `I have noted the deadline around ${deadline}.` : "",
    "",
    "Best,",
  ]
    .filter((line, index, arr) => !(line === "" && arr[index - 1] === ""))
    .join("\n");
}

function showMailDialog(action, message, content) {
  state.latestDraft = { action, message, content };
  els.dialogTitle.textContent = actionLabel(action);
  els.dialogSubtitle.textContent = `${message.senderName} · ${message.subject || "(No subject)"}`;
  els.dialogBody.textContent = content;
  els.dialogDraft.classList.toggle("hidden", action === "summary" || (state.mode !== "graph" && state.mode !== "outlook-desktop"));
  els.dialog.showModal();
}

async function createDraftFromDialog() {
  const draft = state.latestDraft;
  if (!draft || state.hostedDemo || (state.mode !== "graph" && state.mode !== "outlook-desktop")) {
    if (state.hostedDemo) {
      showToast("Hosted demo only. Run locally to create drafts.");
    }
    return;
  }

  setBusy(true);
  try {
    if (state.mode === "outlook-desktop") {
      const response = await fetch("/api/outlook/draft", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          id: draft.message.id,
          content: draft.content,
        }),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || "Could not create Outlook Desktop draft");
      }
      showToast("Draft reply saved in Outlook Desktop.");
      els.dialog.close();
      return;
    }

    const created = await graphFetch(`/me/messages/${encodeURIComponent(draft.message.id)}/createReply`, {
      method: "POST",
      body: "{}",
    });

    await graphFetch(`/me/messages/${encodeURIComponent(created.id)}`, {
      method: "PATCH",
      body: JSON.stringify({
        body: {
          contentType: "Text",
          content: draft.content,
        },
      }),
    });

    showToast("Draft reply created in Outlook.");
    els.dialog.close();
  } catch (error) {
    showToast(readableError(error));
  } finally {
    setBusy(false);
  }
}

async function copyDialogText() {
  const text = els.dialogBody.textContent || "";
  try {
    await navigator.clipboard.writeText(text);
    showToast("Copied.");
  } catch {
    showToast("Select the text and copy it manually.");
  }
}

function switchView(view) {
  document.querySelectorAll("[data-view]").forEach((button) => button.classList.toggle("active", button.dataset.view === view));
  document.querySelectorAll(".view").forEach((section) => section.classList.toggle("active-view", section.id === `${view}-view`));
}

function setBusy(isBusy) {
  if (els.refresh) els.refresh.disabled = isBusy;
  if (els.signin) els.signin.disabled = isBusy;
}

function actionLabel(action) {
  if (action === "summary") return "Summary";
  if (action === "alternate") return "Alternate reply";
  return "Reply draft";
}

function findDeadline(text) {
  const lowered = text.toLowerCase();
  const patterns = [
    /\b(today|tomorrow|tonight|friday|monday|tuesday|wednesday|thursday|saturday|sunday)\b/i,
    /\bby\s+([a-z]+day|[0-9]{1,2}(?:st|nd|rd|th)?\s+[a-z]+)\b/i,
    /\bdue\s+(?:by|on)?\s*([a-z]+day|today|tomorrow|[0-9]{1,2}(?:st|nd|rd|th)?\s+[a-z]+)\b/i,
  ];
  for (const pattern of patterns) {
    const match = lowered.match(pattern);
    if (match) return match[1] || match[0];
  }
  return "";
}

function formatDate(value) {
  if (!value) return "";
  const date = new Date(value);
  const now = new Date();
  const sameDay = date.toDateString() === now.toDateString();
  if (sameDay) {
    return date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  }
  return date.toLocaleDateString([], { day: "2-digit", month: "short" });
}

function stripHtml(value) {
  const template = document.createElement("template");
  template.innerHTML = value;
  return (template.content.textContent || template.innerText || value).replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}

function readableError(error) {
  const message = error?.message || String(error);
  if (message.includes("spawn EPERM")) return "Windows blocked the Outlook helper. Start the local dashboard from your own terminal and try again.";
  if (message.includes("user_cancelled")) return "Sign-in was cancelled.";
  if (message.includes("AADSTS65001")) return "Microsoft needs permission approval before this app can read mail.";
  if (message.includes("AADSTS700016")) return "Microsoft could not find that client ID. Check the app registration.";
  return message;
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.remove("hidden");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => els.toast.classList.add("hidden"), 5200);
}

function refreshIcons() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}
