# MaliDesk — corrected ready-to-run package

This package restores the original MaliDesk interface and behavior while providing a proper Vite + React + Tailwind build and the authentication shell from the production security upgrade.

## Why the previous hosted version looked wrong
The JSX uses Tailwind utility classes throughout the interface. The previous package did not load Tailwind correctly, so the React data rendered but the styling did not. This package includes the correct Vite entry point and loads the Tailwind runtime before the React app, preserving the original utility-class styling.

## Windows local setup

### 1. Install dependencies
Open PowerShell in the `MaliDesk` folder:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup-windows.ps1
```

### 2. Start PostgreSQL
If Docker Desktop is installed:

```powershell
docker compose up -d postgres
```

The included local database uses:

- Database: `malidesk`
- User: `malidesk`
- Password: `malidesk_local_change_me`
- Host: `127.0.0.1`
- Port: `5432`

Update `server\.env` with:

```env
DATABASE_URL=postgresql://malidesk:malidesk_local_change_me@127.0.0.1:5432/malidesk
SESSION_SECRET=replace-with-a-long-random-secret
COOKIE_SECURE=false
CORS_ORIGIN=http://127.0.0.1:5173
```

For production, use a managed PostgreSQL connection string and set `COOKIE_SECURE=true` behind HTTPS.

### 3. Create database tables
Use your PostgreSQL client (psql, pgAdmin, Supabase SQL Editor, etc.) and run `server\schema.sql` against the MaliDesk database.

With psql:

```powershell
psql "postgresql://malidesk:malidesk_local_change_me@127.0.0.1:5432/malidesk" -f .\server\schema.sql
```

### 4. Create the first Administrator

```powershell
cd server
node bootstrap-admin.js
```

Do this once. After creating the Administrator, disable/remove the bootstrap script from any production deployment.

### 5. Start MaliDesk
Use two PowerShell windows.

API:

```powershell
cd C:\MaliDesk\server
npm start
```

Frontend:

```powershell
cd C:\MaliDesk\frontend
npm run dev
```

Open:

`http://127.0.0.1:5173`

The Vite development server proxies `/api/*` to `http://127.0.0.1:3001`, so `/api/auth/login` no longer returns the frontend 404.

## Production build

```powershell
cd frontend
npm run build
```

The production files are created in `frontend\dist`. The UI styling is loaded from the Tailwind runtime referenced in `frontend\index.html`, so the deployed site needs internet access to load that runtime.

## Important production note
The original MaliDesk application still contains its browser/local data layer for offline compatibility. The authentication API secures account access, but a true multi-device production deployment also requires routing sensitive rental-data mutations through authenticated server endpoints and making the server database the source of truth. Do not treat browser-side role hiding as a security boundary.

## If you already have PostgreSQL installed
You can skip Docker and set `DATABASE_URL` in `server\.env` to your PostgreSQL instance.

## If the API says ECONNREFUSED 127.0.0.1:5432
That means PostgreSQL is not reachable at the address in `DATABASE_URL`. This corrected package loads `server\.env` automatically; the previous package did not, which could make Node fall back to localhost.

Production v2 documentation: see README-PRODUCTION.md.
