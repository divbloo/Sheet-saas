# Sheet SaaS

Realtime collaborative spreadsheet web app built with React, Node.js, Express, Socket.IO, and MongoDB.

## Features

- User signup and login with JWT authentication
- Workspaces, sheets, sharing, and collaborator roles
- Realtime cell updates, style updates, cursors, and presence
- Large sheet storage through paginated MongoDB row documents
- ERP item-master template and configurable dropdown options
- CSV import/export and PDF export
- Backend tests for security helpers, validation, and sheet row utilities

## Project Structure

```text
backend/   Express API, Socket.IO server, Mongoose models, tests
frontend/  React/Vite client
scripts/   Local development helper scripts
```

In production, the backend can serve the built frontend from `frontend/dist`.

## Requirements

- Node.js 20 or newer
- MongoDB database

## Environment

Create `backend/.env` for local development:

```env
NODE_ENV=development
PORT=5000
MONGO_URI=mongodb://127.0.0.1:27017/sheet-saas
JWT_SECRET=replace-with-a-long-random-secret
FRONTEND_URL=http://localhost:5173
```

For production:

```env
NODE_ENV=production
PORT=5000
MONGO_URI=your-production-mongodb-uri
JWT_SECRET=at-least-32-characters-long
FRONTEND_URL=https://your-domain.com
```

`VITE_API_URL` is optional. If it is not set, the frontend uses:

- `http://localhost:5000` in development
- the current site origin in production

## Install

```bash
npm --prefix backend install
npm --prefix frontend install
```

## Development

Run backend and frontend together:

```bash
npm run dev
```

Or run them separately:

```bash
npm run backend
npm run frontend
```

## Verification

```bash
npm test
npm run lint
npm run build
```

On Windows PowerShell, if `npm` is blocked by execution policy, use `npm.cmd`:

```powershell
npm.cmd test
npm.cmd run lint
npm.cmd run build
```

## Production Build

Build the frontend:

```bash
npm run build
```

Start the backend in production mode:

```bash
NODE_ENV=production npm --prefix backend start
```

The backend serves `frontend/dist` automatically when `NODE_ENV=production`.

## Data Migration

If older sheets still store row data inside the sheet document, run:

```bash
npm run migrate:rows
```

## Deployment Notes

- Set `NODE_ENV=production`.
- Set a strong `JWT_SECRET` with at least 32 characters.
- Set `FRONTEND_URL` to the deployed public origin.
- Make sure the server allows WebSocket traffic for Socket.IO.
- Do not commit `.env`, `node_modules`, `dist`, build output, or temporary Excel files.
