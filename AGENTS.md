# Repository working instructions

This repository is the production source for the price research support PWA.

- Keep changes scoped to the user's current request and preserve unrelated work.
- Before finishing a code or data change, run `npm run build`.
- Unless the user explicitly asks not to publish, commit only the files changed for the current task and push the completed change to `origin/main`.
- Pull/rebase against the latest `origin/main` before pushing. Do not force-push.
- A push to `main` triggers `.github/workflows/deploy.yml` and publishes `dist` to the `gh-pages` branch.
- Do not commit `node_modules` or `dist`.
- Browser survey records and photos live in IndexedDB; source deployment does not upload or migrate a user's field data.
