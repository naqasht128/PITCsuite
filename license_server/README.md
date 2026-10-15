# PITC Suite — Licensing, Device Locking & Auto-Update Server

Serverless Node.js API backed by Supabase PostgreSQL for device-locked licensing and private GitHub release update proxying.

## 🚀 Setup Steps

### 1. Supabase Database
1. Create a project at [supabase.com](https://supabase.com).
2. Open **SQL Editor** and execute the contents of `schema.sql`.
3. Go to **Project Settings** -> **API**:
   - Copy **Project URL**
   - Copy **`service_role` secret key** (bypasses RLS safely from the backend)

### 2. Vercel Deployment
1. Import the repository in [vercel.com](https://vercel.com).
2. **Important:** Set **Root Directory** to `license_server`.
3. Add the following environment variables:
   - `SUPABASE_URL`: `https://your-project.supabase.co`
   - `SUPABASE_SERVICE_ROLE_KEY`: `eyJhbGciOi...`
   - `ADMIN_SECRET`: Strong secret password (used by `generate_key.py`)
   - `GITHUB_TOKEN`: Fine-grained Personal Access Token with **Contents: Read-only** access to `naqasht128/PitcSuite`.
   - `GITHUB_REPO`: `naqasht128/PitcSuite`
4. Deploy! Your API will be live at `https://pit-csuite.vercel.app` (or your assigned Vercel domain).

## 📡 Endpoints

- `GET /health` — Server health check
- `POST /api/activate` — Activate key for a hardware device
- `POST /api/verify` — Verify key and device binding
- `POST /api/admin/generate` — Generate new serial key (`PS-XXXX-XXXX-XXXX`), protected by `x-admin-secret`
- `GET /api/updates/latest` — Check latest version metadata
- `GET /api/updates/download?asset_id=...` — Download release asset without exposing GitHub credentials
- `POST /api/admin/set-update` — Manually override or register version info
