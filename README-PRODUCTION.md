# MaliDesk — Next Production Version

This release keeps the existing MaliDesk screens and business logic while moving the authoritative rental dataset behind the authenticated API/PostgreSQL layer.

## What changed
- Rental state is stored in PostgreSQL instead of being browser-only when online.
- Browser storage remains as an offline cache.
- Optimistic versioning prevents silent last-write-wins overwrites when two users edit at once.
- Server-side permission checks are derived from the actual data diff; the client cannot grant itself a role.
- Backend audit logging records login, user administration, password changes, and rental-data mutations.
- Temporary passwords force a password change at first login.
- Password reset revokes existing sessions.
- Remember-me now has a longer persistent session; normal sessions are shorter-lived browser sessions.
- Health endpoint added for deployment monitoring.
- Dataset size and basic structure limits are enforced at the API boundary.
- LocalStorage fallback makes offline operation resilient.

## Important
This is a production architecture upgrade, but a security review, HTTPS reverse proxy, managed PostgreSQL backups, secret management, and deployment-specific penetration/regression testing are still required before a public launch.

## Windows local run
1. Install Node.js 22+ and Docker Desktop.
2. In `server/`, copy `.env.example` to `.env` and set a strong `SESSION_SECRET`.
3. Run `docker compose up -d postgres`.
4. Apply `server/schema.sql` to PostgreSQL.
5. Run `node server/bootstrap-admin.js` and create the first Administrator.
6. Start API: `cd server; npm start`.
7. Start UI: `cd frontend; npm install; npm run dev`.

## First data initialisation
The first Administrator opening a new empty MaliDesk instance will bootstrap the existing browser seed/local dataset into PostgreSQL. After that, PostgreSQL is authoritative.

## Recommended public deployment
Use HTTPS and serve the frontend and API from the same origin through a reverse proxy/load balancer. Keep PostgreSQL private. Set:
- `NODE_ENV=production`
- `COOKIE_SECURE=true`
- `CORS_ORIGIN` only if your deployment genuinely requires cross-origin requests
- a long random `SESSION_SECRET`
- a managed PostgreSQL connection with TLS

Do not commit `.env` or database credentials.


## Render: separate frontend + API
If the frontend is deployed at `https://malidesk-frontend.onrender.com` and the API at `https://malidesk.onrender.com`, configure these environment variables:

**Frontend service**
- `VITE_API_BASE_URL=https://malidesk.onrender.com`

**API service**
- `NODE_ENV=production`
- `COOKIE_SECURE=true`
- `CORS_ORIGIN=https://malidesk-frontend.onrender.com`
- `DATABASE_URL=<your Render PostgreSQL external/internal connection string>`
- `SESSION_SECRET=<long random secret>`

After changing frontend environment variables, rebuild/redeploy the frontend because Vite injects `VITE_*` variables at build time. The frontend authentication client also has a production fallback to `https://malidesk.onrender.com`, so `/api/auth/*` requests are never intentionally sent to the frontend origin in a production build. The API health check is `/api/health`.

### Deployment verification
1. On the frontend Render service, set `VITE_API_BASE_URL=https://malidesk.onrender.com`.
2. Trigger a fresh frontend deploy/build.
3. Open `https://malidesk.onrender.com/api/health` and confirm the API responds.
4. In browser DevTools, confirm authentication requests go to `https://malidesk.onrender.com/api/auth/me` rather than `https://malidesk-frontend.onrender.com/api/auth/me`.
5. A logged-out `/api/auth/me` response of `401` is expected; a `404` from the frontend domain indicates an incorrectly built/deployed frontend.

