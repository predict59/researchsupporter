# Repository working instructions

This repository is the production source for the price research support PWA.

- Keep changes scoped to the user's current request and preserve unrelated work.
- Before finishing a code or data change, run `npm run build`.
- For every functional change, add a Korean entry at the top of `CHANGELOG.md`. Record the change purpose, user-visible behavior, data-structure impact, cautions or rollback point, and the commit message (or commit hash when available).
- When a change affects the app's architecture, data model, business rules, supported workflow, or operational risks, update `docs/PROJECT_STATUS.md` and its `최종 갱신` date.
- When a user-visible screen, label, button, input rule, export procedure, backup procedure, or field workflow changes, update `docs/USER_GUIDE.md` in the same task.
- For deployed app-shell or runtime changes, increment `CACHE_NAME` in `public/service-worker.js` so installed PWAs receive the new version. A reference-data-only update does not require a cache bump unless the cached file must be invalidated.
- Unless the user explicitly asks not to publish, commit only the files changed for the current task and push the completed change to `origin/main`.
- Pull/rebase against the latest `origin/main` before pushing. Do not force-push.
- A push to `main` triggers `.github/workflows/deploy.yml` and publishes `dist` to the `gh-pages` branch.
- Do not commit `node_modules` or `dist`.
- Browser survey records and photos live in IndexedDB; source deployment does not upload or migrate a user's field data.
