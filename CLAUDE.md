# Notes for working on this repo

- Standalone app. Never share code, data, keys or storage names with the MMT dashboard.
- `index.html` is the whole app; `apps-script/Code.gs` is the Google Sheet backend.
- Every change bumps the version tag (`BUILD` + `#build-tag` in index.html, `CACHE` in sw.js,
  `BACKEND_VERSION` in Code.gs when the backend changes). Format `M.DD vN`.
- Test with the built-in demo mode (no real Sheet). Never write to the real Sheet when testing.
- All writes go through `queue()` and must be safe to repeat (upsert by id / delete by id).
