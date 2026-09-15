# Package Pulse

A GitHub Pages-ready Progressive Web App that turns a package photo into a prefilled Outlook draft.

## The workflow

1. Open Package Pulse on your phone.
2. Tap **Take package photo**.
3. The browser opens the rear camera.
4. Package Pulse compresses the image locally.
5. Tap **Create Outlook Draft**.
6. Microsoft Graph creates a draft in your Microsoft 365 mailbox containing:
   - **Subject:** `You Have Mail`
   - **Message:** `Hello, You have received a package, please pick it up at your earliest convenience.`
   - The package photo as an attachment.
7. The app opens the Outlook draft.
8. Use Outlook to find the coworker, review everything, and press **Send**.

Recipient lookup intentionally stays in Outlook; Package Pulse does not request directory access.

## Design

Stanford-inspired Cardinal red meets restrained neon-noir / futuristic terminal UI. No Stanford logo is included, and the prototype does not present itself as an official Stanford product.

## Stack

- Vite
- Vanilla JavaScript
- `@azure/msal-browser`
- Microsoft Graph
- PWA manifest + service worker
- GitHub Actions / GitHub Pages

## Microsoft Entra setup

Your organization may require an administrator to create or approve this registration.

### 1. Register the SPA

In Microsoft Entra:

1. **App registrations** → **New registration**
2. Name: `Package Pulse`
3. Choose the account type appropriate for your organization.
4. Under **Authentication**, add a **Single-page application (SPA)** redirect URI.

Local development:

```text
http://localhost:5173/
```

GitHub Pages:

```text
https://YOUR-GITHUB-USERNAME.github.io/package-pulse/
```

### 2. Add delegated Microsoft Graph permissions

Add:

- `Mail.ReadWrite`
- `User.Read`

`Mail.ReadWrite` is what allows the signed-in user to create a draft in their own mailbox.

Depending on tenant policy, administrator consent may be required.

### 3. Configure the app

Edit `public/config.js`:

```js
window.PACKAGE_PULSE_CONFIG = {
  clientId: "YOUR-APPLICATION-CLIENT-ID",
  tenantId: "YOUR-TENANT-ID",
  // ...
};
```

The client ID and tenant ID are identifiers, not secrets.

**Never put a client secret in this repository.**

## Local development

Install dependencies:

```bash
npm install
```

Start Vite:

```bash
npm run dev
```

Open:

```text
http://localhost:5173/
```

Make sure the exact URL is registered as an SPA redirect URI in Entra.

## GitHub deployment

Create a repository named `package-pulse`, then:

```bash
git init
git add .
git commit -m "Initial Package Pulse PWA"
git branch -M main
git remote add origin https://github.com/YOUR-GITHUB-USERNAME/package-pulse.git
git push -u origin main
```

Then open:

**GitHub → repository → Settings → Pages → Source → GitHub Actions**

The included `.github/workflows/deploy-pages.yml` builds and deploys the site whenever `main` is updated.

After GitHub gives you the final Pages address, make sure that exact address is also registered as an SPA redirect URI in Microsoft Entra.

## Draft-opening behavior

Microsoft Graph exposes a `webLink` property whose documented purpose is to open the message in Outlook on the web.

Package Pulse uses that supported link as the handoff to Outlook.

There is also a progressive enhancement in `src/app.js` that attempts to turn a modern Outlook `/deeplink/read/` route into `/deeplink/compose/` so the draft opens directly in editing mode. The compose-route behavior is not a documented Microsoft Graph contract, so the normal Graph `webLink` remains the fallback.

This prototype therefore targets **Outlook on the web** as the reliable handoff. Whether a phone launches the native Outlook app for that URL is controlled by the device/app and should not be assumed.

## Image handling

Before uploading to Graph, the app:

- scales the longest image dimension to a maximum of 1600 px,
- converts it to JPEG,
- progressively lowers JPEG quality if necessary,
- aims to keep the attachment comfortably below the small-attachment range.

The package photo is not saved by Package Pulse to a custom database or server.

## PWA behavior

The app includes:

- a manifest,
- an installable app icon,
- a service worker that caches the application shell.

The service worker does **not** intentionally cache Graph API responses or captured package photos.

## Automatic message

Edit `public/config.js` if you want to change it:

```js
message: {
  subject: "You Have Mail",
  bodyHtml:
    "<p>Hello,</p>" +
    "<p>You have received a package, please pick it up at your earliest convenience.</p>"
}
```

## Organizational review

Before treating this as a production internal tool, confirm your organization's requirements for:

- Microsoft Entra app registration and consent
- `Mail.ReadWrite`
- GitHub Pages / external hosting
- package-label photography and privacy
- internal application branding
