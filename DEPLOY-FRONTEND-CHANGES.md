# MaliDesk frontend update — deployment instructions

## Changes checked in the source
- `frontend/src/main.jsx` loads `frontend/src/App.jsx`, which loads the canonical `frontend/src/MaliDesk.jsx`.
- The canonical Import & export component contains no `Download CSV` action.
- `Download Excel` remains and is styled green (`#1F8A4C`).
- The canonical source has no top-bar download icon or notification bell controls.
- The Excel template is served from `frontend/public/rental-income-template.xlsx`.

## Deploy the FRONTEND service on Render
The screenshot's requested controls are part of the browser UI. Deploying only the API service (`https://malidesk.onrender.com`, which runs `node index.js`) will not update the page at `https://malidesk-frontend.onrender.com`.

In the Render dashboard, open the **existing frontend Static Site service** for MaliDesk and confirm it is connected to the intended repository and branch. Set:
- **Root Directory:** `frontend`
- **Build Command:** `npm ci && npm run build`
- **Publish Directory:** `dist`
- **Environment Variable:** `VITE_API_BASE_URL=https://malidesk.onrender.com`

Then save and trigger a fresh deploy of that frontend service. Do not change the backend service or database for these UI-only changes. If the existing frontend service uses a different root directory or a different deployment type, confirm its current settings before changing them.

After deployment, open `https://malidesk-frontend.onrender.com` and do a hard refresh (Ctrl+Shift+R). Confirm the top-right download and bell icons are gone, Download CSV is absent, and Download Excel is green. Also test an Excel export and the login/data flow.

## Build verification status
The source was statically checked for the requested UI controls and the frontend entry point. A production build could not be completed in the current environment because dependencies were not installed and `npm ci` did not finish before the environment timeout. Run the build command above in Render or on a machine with npm registry access before treating the release as build-verified.
