# Linux Production Readiness and Clean-Code Plan

## Goal

Prepare Sheet SaaS for a safe move from Windows to a Linux production server without changing existing business behavior or risking MongoDB data.

## Assumptions

- Production runs Node.js behind Nginx and is managed by PM2.
- MongoDB remains external (for example MongoDB Atlas); the application server does not own the primary database files.
- Existing API behavior and sheet data formats must remain backward compatible.
- Deployment to GitHub is a separate action and will only happen after explicit approval.

## Phase 1 - Baseline and audit

- Inventory runtime, build, test, environment, and deployment files.
- Run the existing backend/frontend test, lint, and build commands.
- Audit dependency vulnerabilities without applying unsafe forced upgrades.
- Check for Windows-only paths, commands, and case-insensitive import mistakes.

## Phase 2 - Linux runtime foundation

- Validate required environment variables at startup with test coverage.
- Add a health endpoint that reports application and database readiness without leaking secrets.
- Add graceful shutdown for SIGTERM/SIGINT so PM2 and systemd can deploy safely.
- Harden the PM2 configuration for Linux restart and shutdown behavior.

## Phase 3 - Automated Linux verification

- Add one root verification command covering both applications.
- Add a deterministic import-case check for Linux compatibility.
- Add GitHub Actions CI on Ubuntu using the supported Node.js version.

## Phase 4 - Operations documentation

- Document first-time Linux installation, Nginx/WebSocket proxying, PM2 startup, SSL, firewall, and environment permissions.
- Document deployment, health verification, MongoDB backup/restore, and rollback.
- Update the root README and environment examples.

## Phase 5 - Final quality gate

- Run tests, lint, build, case checks, and dependency audit.
- Review the complete diff for security, portability, behavior changes, and secrets.
- Record any remaining architectural debt that should be split into a future low-risk refactor.
