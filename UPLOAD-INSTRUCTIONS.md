# GitHub Replacement Instructions

This folder is ready to replace the current Package Pulse repository contents.

## What changed
- Removed the redundant **Outlook Link** summary panel.
- Updated the email subject everywhere to **You Have a Package**.
- Kept the Outlook connection button in the top navigation.
- Simplified the hero layout.
- Bumped the PWA cache version so browsers are less likely to keep the old interface.
- Kept the existing mobile camera and Outlook draft prototype behavior.

## Recommended upload
1. Keep your GitHub repository itself.
2. Delete the old project files/folders from the repository.
3. Upload the contents of this folder to the repository root.
4. Keep **Settings → Pages → Source** set to **GitHub Actions**.
5. Commit to `main`.
6. Wait for the Pages deployment workflow to finish.
7. Refresh the live site.

Expected root structure:

```text
.github/
public/
src/
.gitignore
README.md
index.html
package.json
vite.config.js
```

The `.github/workflows/deploy-pages.yml` file is included.
