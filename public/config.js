// Package Pulse configuration
//
// Your Microsoft Entra Application (client) ID is a public identifier, not a client secret.
// For a Stanford-only deployment, replace TENANT_ID with your organization's tenant ID.
// If your administrator prefers multi-tenant sign-in, authority can use "organizations".
//
// IMPORTANT: never put a client secret in a browser/PWA repository.

window.PACKAGE_PULSE_CONFIG = {
  clientId: "PASTE-YOUR-APPLICATION-CLIENT-ID-HERE",
  tenantId: "PASTE-YOUR-TENANT-ID-HERE",
  redirectUri: window.location.origin + window.location.pathname,
  scopes: ["Mail.ReadWrite", "User.Read"],

  message: {
    subject: "You Have Mail",
    bodyHtml:
      "<p>Hello,</p>" +
      "<p>You have received a package, please pick it up at your earliest convenience.</p>"
  }
};
