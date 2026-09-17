# Package Pulse

Package Pulse is a mobile-first PWA for preparing package-notification emails.

## Workflow

1. Take a package photo.
2. Create the prepared Outlook draft.
3. Select the coworker in Outlook.
4. Review the message and attachment.
5. Send.

## Email template

**Subject:** You Have a Package

**Message:**

Hello,

You have received a package, please pick it up at your earliest convenience.

## Current status

- Mobile camera/photo workflow is implemented.
- Responsive desktop and mobile interface is implemented.
- Outlook/Microsoft 365 integration code is included as a prototype and still requires the approved organization configuration.
- The redundant Outlook summary panel has been removed from the interface.

## Development

```bash
npm install
npm run dev
```

## Production build

```bash
npm run build
```

The GitHub Pages workflow is included under:

```text
.github/workflows/deploy-pages.yml
```
