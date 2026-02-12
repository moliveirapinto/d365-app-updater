# Supabase Setup for Usage Tracking

This guide explains how to set up **Supabase** (free tier) to track who uses the D365 App Updater and what they update.

## 1. Create a Supabase Account & Project

1. Go to [supabase.com](https://supabase.com) and sign up (free).
2. Click **New Project**.
3. Choose an organization, set a project name (e.g. `d365-app-updater`), and set a database password.
4. Pick a region close to your users and click **Create new project**.
5. Wait for the project to finish provisioning (< 1 minute).

## 2. Create the `usage_logs` Table

Go to **SQL Editor** (left sidebar) and run this SQL:

```sql
CREATE TABLE usage_logs (
    id          BIGSERIAL PRIMARY KEY,
    timestamp   TIMESTAMPTZ NOT NULL DEFAULT now(),
    user_email  TEXT,
    org_url     TEXT,
    environment_id TEXT,
    success_count INTEGER DEFAULT 0,
    fail_count    INTEGER DEFAULT 0,
    total_apps    INTEGER DEFAULT 0,
    app_names     TEXT
);

-- Index for fast querying
CREATE INDEX idx_usage_logs_timestamp ON usage_logs (timestamp DESC);
CREATE INDEX idx_usage_logs_user ON usage_logs (user_email);
CREATE INDEX idx_usage_logs_org ON usage_logs (org_url);

-- Enable Row Level Security (required by Supabase)
ALTER TABLE usage_logs ENABLE ROW LEVEL SECURITY;

-- Allow inserts from the anon key (app users logging usage)
CREATE POLICY "Allow anonymous inserts"
    ON usage_logs FOR INSERT
    TO anon
    WITH CHECK (true);

-- Allow reads from the anon key (admin dashboard viewing data)
CREATE POLICY "Allow anonymous reads"
    ON usage_logs FOR SELECT
    TO anon
    USING (true);
```

Click **Run** to execute.

## 2b. Create the `update_schedules` Table (for Auto-Updates)

If you want to enable the **Auto-Update Scheduling** feature, also run this SQL:

```sql
CREATE TABLE update_schedules (
    id              BIGSERIAL PRIMARY KEY,
    user_email      TEXT NOT NULL,
    environment_id  TEXT NOT NULL,
    org_url         TEXT,
    enabled         BOOLEAN DEFAULT false,
    day_of_week     INTEGER CHECK (day_of_week >= 0 AND day_of_week <= 6), -- 0=Sunday, 6=Saturday
    time_utc        TEXT DEFAULT '03:00', -- HH:MM format in user's timezone
    timezone        TEXT DEFAULT 'UTC',
    client_id       TEXT,              -- Azure AD App Registration Client ID
    tenant_id       TEXT,              -- Azure AD Tenant ID
    has_secret      BOOLEAN DEFAULT false, -- True if secret is stored in schedule_secrets table
    last_run_at     TIMESTAMPTZ,
    last_run_status TEXT,
    last_run_result JSONB,
    created_at      TIMESTAMPTZ DEFAULT now(),
    updated_at      TIMESTAMPTZ DEFAULT now(),
    UNIQUE(user_email, environment_id)
);

-- Indexes for efficient querying by the scheduler
CREATE INDEX idx_schedules_enabled ON update_schedules (enabled) WHERE enabled = true;
CREATE INDEX idx_schedules_user ON update_schedules (user_email);

-- Enable Row Level Security
ALTER TABLE update_schedules ENABLE ROW LEVEL SECURITY;

-- Allow inserts and updates from anon key (users configuring their schedules)
CREATE POLICY "Allow anonymous inserts"
    ON update_schedules FOR INSERT
    TO anon
    WITH CHECK (true);

CREATE POLICY "Allow anonymous updates"
    ON update_schedules FOR UPDATE
    TO anon
    USING (true)
    WITH CHECK (true);

CREATE POLICY "Allow anonymous reads"
    ON update_schedules FOR SELECT
    TO anon
    USING (true);
```

### SECURITY: Create a Separate Secrets Table (REQUIRED)

The client secrets are stored in a **separate protected table** that the public anon key **cannot read**:

```sql
-- Create a secure table for storing client secrets
-- Only service_role can read this table (used by GitHub Actions)
CREATE TABLE schedule_secrets (
    id              BIGSERIAL PRIMARY KEY,
    schedule_id     BIGINT NOT NULL REFERENCES update_schedules(id) ON DELETE CASCADE,
    client_secret   TEXT NOT NULL,
    created_at      TIMESTAMPTZ DEFAULT now(),
    updated_at      TIMESTAMPTZ DEFAULT now(),
    UNIQUE(schedule_id)
);

-- Enable Row Level Security
ALTER TABLE schedule_secrets ENABLE ROW LEVEL SECURITY;

-- CRITICAL: Anon can INSERT (create new secrets)
CREATE POLICY "Anon can insert secrets"
    ON schedule_secrets FOR INSERT
    TO anon
    WITH CHECK (true);

-- CRITICAL: Anon can UPDATE (change existing secrets)
CREATE POLICY "Anon can upsert secrets"
    ON schedule_secrets FOR UPDATE
    TO anon
    USING (true)
    WITH CHECK (true);

-- NO SELECT POLICY FOR ANON = anon key CANNOT read secrets
-- Only service_role (used by GitHub Actions) can read secrets
```

### If you already have client_secret in update_schedules, migrate it:

```sql
-- 1. Create the schedule_secrets table (run SQL above first)

-- 2. Migrate existing secrets
INSERT INTO schedule_secrets (schedule_id, client_secret)
SELECT id, client_secret FROM update_schedules 
WHERE client_secret IS NOT NULL AND client_secret != '';

-- 3. Remove secret from main table (optional but recommended)
ALTER TABLE update_schedules DROP COLUMN IF EXISTS client_secret;
```

## 3. Get Your Project URL and Keys

1. Go to **Settings** → **API** (left sidebar).
2. Copy:
   - **Project URL**: `https://xxxxxxxx.supabase.co`
   - **anon (public) key**: Used by the web app (public, safe to expose)
   - **service_role key**: Used ONLY by GitHub Actions (KEEP SECRET!)

**⚠️ IMPORTANT:** Add the `service_role` key to your GitHub repository secrets as `SUPABASE_SERVICE_KEY`

## 4. Configure the Admin Dashboard

1. Open the admin dashboard: `https://moliveirapinto.github.io/d365-app-updater/admin.html`
2. Enter your **Project URL** and **anon key** in the config bar at the top.
3. Click **Connect**.
4. The config is saved in your browser's localStorage.

## 5. Configure the Main App

The main app also needs the Supabase config to log usage data. It reads from the same localStorage key (`d365_admin_supabase`), so:

- If you've already connected via the admin dashboard on the same browser, the main app will automatically log usage.
- Alternatively, you can set it manually in the browser console:
  ```js
  localStorage.setItem('d365_admin_supabase', JSON.stringify({
      url: 'https://YOUR_PROJECT.supabase.co',
      key: 'YOUR_ANON_KEY'
  }));
  ```

## 6. Security Notes

- The **anon key** only allows `INSERT` and `SELECT` on `usage_logs` (via RLS policies above).
- No one can `UPDATE` or `DELETE` data with the anon key.
- For stricter security, create a **service_role** policy that only allows `SELECT` and use the service_role key only in the admin dashboard.
- The anon key is safe to embed in client-side code — it only has the permissions you grant via RLS.

## 7. What Gets Logged

Each update session logs:

| Field | Description |
|-------|-------------|
| `timestamp` | When the update was performed (UTC) |
| `user_email` | Email of the authenticated user (from MSAL) |
| `org_url` | The Dynamics 365 organization URL |
| `environment_id` | Power Platform environment GUID |
| `success_count` | Number of apps updated successfully |
| `fail_count` | Number of apps that failed to update |
| `total_apps` | Total apps attempted |
| `app_names` | Comma-separated list of app names |

## 8. Admin Dashboard Features

- **Summary cards**: Total sessions, unique users, environments, success/fail totals
- **Dashboard view**: Timeline chart, success/fail donut, top environments & users bar charts, recent activity feed
- **Table view**: Full sortable/filterable/paginated table with search, date range, status filters, and CSV export
