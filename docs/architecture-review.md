# Architecture and Production Readiness Review

## Resolved in this pass

- Added strict startup validation for MongoDB, JWT, port, frontend origin, and
  reverse-proxy configuration.
- Added database-aware readiness (`/healthz`) and graceful SIGTERM/SIGINT
  handling for safe PM2/systemd reloads.
- Added an Ubuntu CI quality gate and a filename-case check that catches imports
  which work on Windows but fail on Linux.
- Pinned the supported Node.js baseline and made one root `npm run verify`
  command cover backend tests, frontend tests, tooling tests, lint, and build.
- Updated vulnerable backend dependency locks to remove all known production
  audit findings.
- Replaced the vulnerable, unpatched `xlsx` package with an isolated lazy-loaded
  ExcelJS adapter and capped imported workbooks at 15 MB.
- Added PM2 readiness, graceful-stop, memory, and restart limits.
- Added a complete Linux deployment, backup, restore, smoke-test, and rollback
  runbook.

## Current boundaries

- MongoDB is the system of record. Application hosts are disposable as long as
  the database, `.env`, DNS/TLS, and deployment commit are backed up.
- PM2 must stay at one application instance because in-memory presence and
  Socket.IO rooms are process-local. Horizontal scaling requires a shared
  Socket.IO adapter and shared presence storage.
- Spreadsheet import supports `.xlsx` and `.csv`. Legacy binary `.xls` parsing
  was removed with the vulnerable SheetJS dependency.

## Remaining maintainability work

`backend/server.js` and `frontend/src/App.jsx` are still large orchestration
files. Splitting them is worthwhile, but doing a big-bang rewrite during a
server migration would add unnecessary regression risk. Continue incrementally:

1. Move sheet REST handlers from `server.js` into route modules grouped by
   rows, sharing, versions, and exports.
2. Move Socket.IO handlers into a realtime service with explicit authorization
   and persistence dependencies.
3. Split `App.jsx` into dashboard, toolbar, virtualized grid, sharing, export,
   and version-history feature components.
4. Add API integration tests with an isolated MongoDB test database, then add
   browser tests for login, open sheet, edit, realtime sync, and export.
5. Add a shared Socket.IO adapter before changing PM2 `instances` above one.

These are architectural improvements, not blockers for the documented
single-instance Linux deployment. They should land as small behavior-preserving
changes with tests, not as one rewrite.
