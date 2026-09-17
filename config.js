// Package Pulse configuration
//
// The values below are placeholders until the Apps/Identity team selects and
// approves the Outlook integration method.
//
// If the existing Microsoft Graph SPA prototype is used, the client ID and
// tenant ID are identifiers, not client secrets.
//
// IMPORTANT: never place a client secret, password, or private credential in
// a browser/PWA repository.

window.PACKAGE_PULSE_CONFIG = {
  clientId: "PASTE-APPROVED-APPLICATION-CLIENT-ID-HERE",
  tenantId: "PASTE-APPROVED-TENANT-ID-HERE",
  redirectUri: window.location.origin + window.location.pathname,
  scopes: ["Mail.ReadWrite", "User.Read"],

  message: {
    subject: "You Have a Package",
    bodyHtml:
      "<p>Hello,</p>" +
      "<p>You have received a package, please pick it up at your earliest convenience.</p>"
  }
};
