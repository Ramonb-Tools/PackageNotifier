import { PublicClientApplication } from "@azure/msal-browser";
import "./styles.css";

/* Package Pulse
 * Creates a Microsoft 365 Outlook draft with a captured package image.
 * No recipient is added by the app: the user chooses the coworker in Outlook.
 */

const cfg = window.PACKAGE_PULSE_CONFIG;

const els = {
  signInButton: document.getElementById("signInButton"),
  signInLabel: document.getElementById("signInLabel"),
  authBadge: document.getElementById("authBadge"),
  statusDot: document.querySelector(".status-dot"),
  photoInput: document.getElementById("photoInput"),
  previewImage: document.getElementById("previewImage"),
  capturePlaceholder: document.getElementById("capturePlaceholder"),
  retakeButton: document.getElementById("retakeButton"),
  createDraftButton: document.getElementById("createDraftButton"),
  imageBadge: document.getElementById("imageBadge"),
  imageMeta: document.getElementById("imageMeta"),
  attachmentName: document.getElementById("attachmentName"),
  attachmentSize: document.getElementById("attachmentSize"),
  statusPanel: document.getElementById("statusPanel"),
  statusTitle: document.getElementById("statusTitle"),
  statusMessage: document.getElementById("statusMessage"),
  openDraftButton: document.getElementById("openDraftButton")
};

let msalApp = null;
let account = null;
let processedPhoto = null;
let lastDraftLink = null;

function isConfigured() {
  return cfg?.clientId &&
    !cfg.clientId.includes("PASTE-") &&
    cfg?.tenantId &&
    !cfg.tenantId.includes("PASTE-");
}

function showStatus(title, message, mode = "info") {
  els.statusPanel.hidden = false;
  els.statusTitle.textContent = title;
  els.statusMessage.textContent = message;
  els.statusPanel.dataset.mode = mode;
}

function setAuthUI(isConnected, username = "") {
  els.statusDot.classList.toggle("online", isConnected);
  els.authBadge.className = isConnected ? "badge badge-online" : "badge badge-offline";
  els.authBadge.textContent = isConnected ? "CONNECTED" : "OFFLINE";
  els.signInLabel.textContent = isConnected ? (username || "Outlook connected") : "Connect Outlook";
}

async function initializeAuth() {
  if (!isConfigured()) {
    showStatus(
      "CONFIGURATION REQUIRED",
      "Add your Microsoft Entra client ID and tenant ID in config.js before Outlook sign-in can work."
    );
    return;
  }

  msalApp = new PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
      redirectUri: cfg.redirectUri
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false
    }
  });

  await msalApp.initialize();
  const redirectResult = await msalApp.handleRedirectPromise();

  account = redirectResult?.account || msalApp.getAllAccounts()[0] || null;
  if (account) {
    msalApp.setActiveAccount(account);
    setAuthUI(true, account.name || account.username);
  }
}

async function signIn() {
  if (!isConfigured()) {
    showStatus(
      "CONFIGURATION REQUIRED",
      "Open config.js and add the Microsoft Entra Application (client) ID and tenant ID."
    );
    return;
  }

  if (!msalApp) {
    await initializeAuth();
    if (!msalApp) return;
  }

  if (account) {
    showStatus("OUTLOOK CONNECTED", `Signed in as ${account.username || account.name}.`);
    return;
  }

  await msalApp.loginRedirect({
    scopes: cfg.scopes,
    prompt: "select_account"
  });
}

async function getAccessToken() {
  if (!account) {
    await signIn();
    throw new Error("Sign-in required.");
  }

  try {
    const result = await msalApp.acquireTokenSilent({
      account,
      scopes: cfg.scopes
    });
    return result.accessToken;
  } catch {
    const result = await msalApp.acquireTokenPopup({
      account,
      scopes: cfg.scopes
    });
    return result.accessToken;
  }
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => resolve(reader.result);
    reader.readAsDataURL(file);
  });
}

function dataUrlToBase64(dataUrl) {
  return dataUrl.split(",")[1];
}

async function compressImage(file) {
  const source = await fileToDataUrl(file);
  const img = new Image();
  img.src = source;
  await img.decode();

  const maxDimension = 1600;
  const scale = Math.min(1, maxDimension / Math.max(img.naturalWidth, img.naturalHeight));
  const width = Math.round(img.naturalWidth * scale);
  const height = Math.round(img.naturalHeight * scale);

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d", { alpha: false });
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.drawImage(img, 0, 0, width, height);

  let quality = 0.84;
  let blob = await new Promise(resolve => canvas.toBlob(resolve, "image/jpeg", quality));

  // Keep the attachment comfortably below the small-attachment range.
  while (blob && blob.size > 2.3 * 1024 * 1024 && quality > 0.5) {
    quality -= 0.08;
    blob = await new Promise(resolve => canvas.toBlob(resolve, "image/jpeg", quality));
  }

  if (!blob) throw new Error("Could not process image.");

  const dataUrl = await fileToDataUrl(blob);
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");

  return {
    blob,
    dataUrl,
    contentBytes: dataUrlToBase64(dataUrl),
    name: `package-${stamp}.jpg`,
    contentType: "image/jpeg"
  };
}

async function handlePhoto(file) {
  if (!file) return;

  showStatus("PROCESSING IMAGE", "Optimizing the package photo for the Outlook attachment…");

  try {
    processedPhoto = await compressImage(file);

    els.previewImage.src = processedPhoto.dataUrl;
    els.previewImage.style.display = "block";
    els.capturePlaceholder.style.display = "none";
    els.retakeButton.disabled = false;
    els.imageBadge.textContent = "IMAGE READY";
    els.imageBadge.classList.add("badge-ready");
    els.imageMeta.textContent = `${processedPhoto.name} · ${formatBytes(processedPhoto.blob.size)}`;
    els.attachmentName.textContent = processedPhoto.name;
    els.attachmentSize.textContent = formatBytes(processedPhoto.blob.size);
    els.createDraftButton.disabled = false;

    showStatus("IMAGE READY", "Photo attached locally. Connect Outlook, then create the draft.");
  } catch (error) {
    processedPhoto = null;
    showStatus("IMAGE ERROR", error.message || "Could not process the selected image.", "error");
  }
}

function buildDraftPayload() {
  return {
    subject: cfg.message.subject,
    body: {
      contentType: "HTML",
      content: cfg.message.bodyHtml
    },
    toRecipients: [],
    attachments: [
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: processedPhoto.name,
        contentType: processedPhoto.contentType,
        contentBytes: processedPhoto.contentBytes
      }
    ]
  };
}

function buildComposeCandidate(draft) {
  // Microsoft documents message.webLink for opening a message.
  // A compose-mode variant is used here as a progressive enhancement.
  // If Outlook changes the route, the documented webLink remains available as fallback.
  if (!draft?.webLink) return null;

  try {
    const url = new URL(draft.webLink);

    if (url.pathname.includes("/deeplink/read/")) {
      url.pathname = url.pathname.replace("/deeplink/read/", "/deeplink/compose/");
      return url.toString();
    }

    // Older OWA-style links can still open the draft, though they may initially use a read view.
    return draft.webLink;
  } catch {
    return draft.webLink;
  }
}

async function createDraft() {
  if (!processedPhoto) {
    showStatus("PHOTO REQUIRED", "Take a package photo first.");
    return;
  }

  els.createDraftButton.disabled = true;
  els.createDraftButton.classList.add("loading");
  els.createDraftButton.querySelector(".button-text").textContent = "Creating Draft…";

  try {
    if (!account) {
      await signIn();
      return;
    }

    showStatus("OUTLOOK SYNC", "Creating your prefilled draft in Microsoft 365…");
    const token = await getAccessToken();

    const response = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(buildDraftPayload())
    });

    if (!response.ok) {
      const details = await response.text();
      throw new Error(`Microsoft Graph returned ${response.status}. ${details.slice(0, 260)}`);
    }

    const draft = await response.json();
    lastDraftLink = buildComposeCandidate(draft) || draft.webLink || "https://outlook.office.com/mail/drafts";

    els.openDraftButton.hidden = false;
    showStatus(
      "DRAFT CREATED",
      "Subject, message, and photo are ready. Opening Outlook so you can choose the coworker and send."
    );

    // Let the success state paint before navigating away.
    setTimeout(() => {
      window.location.href = lastDraftLink;
    }, 650);

  } catch (error) {
    console.error(error);
    showStatus(
      "DRAFT ERROR",
      error.message || "The Outlook draft could not be created.",
      "error"
    );
  } finally {
    els.createDraftButton.disabled = false;
    els.createDraftButton.classList.remove("loading");
    els.createDraftButton.querySelector(".button-text").textContent = "Create Outlook Draft";
  }
}

els.signInButton.addEventListener("click", signIn);
els.photoInput.addEventListener("change", event => handlePhoto(event.target.files?.[0]));
els.retakeButton.addEventListener("click", () => els.photoInput.click());
els.createDraftButton.addEventListener("click", createDraft);
els.openDraftButton.addEventListener("click", () => {
  if (lastDraftLink) window.location.href = lastDraftLink;
});

if ("serviceWorker" in navigator && location.protocol === "https:") {
  navigator.serviceWorker.register("./sw.js").catch(console.warn);
}

initializeAuth().catch(error => {
  console.error(error);
  showStatus("AUTH INITIALIZATION ERROR", error.message || "Could not initialize Microsoft sign-in.");
});
