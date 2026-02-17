const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');

// Supabase configuration (only these need to be in env vars)
const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_KEY;

// ═══════════════════════════════════════════════════════════════
// Timer trigger: runs every hour at :05
// Checks Supabase for schedules matching the current UTC day/time,
// then uses each schedule's own stored credentials to authenticate
// and update apps via the Power Platform API — the exact same API
// the "Update All Apps" button uses in the browser.
// ═══════════════════════════════════════════════════════════════
app.timer('ScheduledUpdateTrigger', {
    schedule: '0 5 * * * *',
    handler: async (myTimer, context) => {
        context.log('⏰ Scheduled update check starting...');

        if (!supabaseUrl || !supabaseKey) {
            context.error('❌ SUPABASE_URL and SUPABASE_KEY env vars are required');
            return;
        }

        const now = new Date();
        const currentDayOfWeek = now.getUTCDay();
        const currentHour = now.getUTCHours();
        const currentTimeUtc = `${currentHour.toString().padStart(2, '0')}:00`;
        const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        context.log(`Current UTC: ${days[currentDayOfWeek]} ${currentTimeUtc} (${now.toISOString()})`);

        try {
            const schedules = await getMatchingSchedules(currentDayOfWeek, currentTimeUtc, context);
            if (schedules.length === 0) {
                context.log('No schedules match current UTC slot. Done.');
                return;
            }
            context.log(`Found ${schedules.length} schedule(s) to process`);

            for (const schedule of schedules) {
                context.log(`── Processing schedule #${schedule.id} for ${schedule.user_email} ──`);
                await processSchedule(schedule, context);
            }
            context.log('✅ Scheduled update check completed');
        } catch (error) {
            context.error('❌ Scheduled update check failed:', error.message);
        }
    }
});

// ── Supabase helpers ──────────────────────────────────────────

async function getMatchingSchedules(dayOfWeek, timeUtc, context) {
    // Fetch enabled schedules whose UTC day + time match now
    const url = `${supabaseUrl}/rest/v1/update_schedules?enabled=eq.true&day_of_week=eq.${dayOfWeek}&time_utc=eq.${timeUtc}&select=*`;
    const resp = await fetch(url, { headers: sbHeaders() });
    if (!resp.ok) throw new Error(`Supabase schedules query failed: ${resp.status}`);
    return resp.json();
}

async function getScheduleSecret(scheduleId, context) {
    // Read the client_secret stored in the secure table
    const url = `${supabaseUrl}/rest/v1/schedule_secrets?schedule_id=eq.${scheduleId}&select=client_secret`;
    const resp = await fetch(url, { headers: sbHeaders() });
    if (!resp.ok) {
        context.warn(`Could not read schedule_secrets (${resp.status}), trying update_schedules.client_secret`);
        return null;
    }
    const rows = await resp.json();
    return rows.length > 0 ? rows[0].client_secret : null;
}

async function getSecretFallback(scheduleId, context) {
    // Fallback: read client_secret directly from update_schedules row
    const url = `${supabaseUrl}/rest/v1/update_schedules?id=eq.${scheduleId}&select=client_secret`;
    const resp = await fetch(url, { headers: sbHeaders() });
    if (!resp.ok) return null;
    const rows = await resp.json();
    return rows.length > 0 ? rows[0].client_secret : null;
}

async function updateScheduleResult(scheduleId, status, result, context) {
    const url = `${supabaseUrl}/rest/v1/update_schedules?id=eq.${scheduleId}`;
    const resp = await fetch(url, {
        method: 'PATCH',
        headers: { ...sbHeaders(), 'Content-Type': 'application/json' },
        body: JSON.stringify({
            last_run_at: new Date().toISOString(),
            last_run_status: status,
            last_run_result: result
        })
    });
    if (!resp.ok) context.warn(`Failed to save run result: ${resp.status}`);
}

function sbHeaders() {
    return { 'apikey': supabaseKey, 'Authorization': `Bearer ${supabaseKey}` };
}

// ── Core processing ───────────────────────────────────────────

async function processSchedule(schedule, context) {
    let status = 'success';
    let result = { appsUpdated: 0, appsFailed: 0, apps: [] };

    try {
        // 1. Resolve credentials from the schedule record
        const tenantId = schedule.tenant_id;
        const clientId = schedule.client_id;
        if (!tenantId || !clientId) throw new Error('Schedule is missing tenant_id or client_id');

        let clientSecret = await getScheduleSecret(schedule.id, context);
        if (!clientSecret) clientSecret = await getSecretFallback(schedule.id, context);
        if (!clientSecret) throw new Error('No client_secret found for this schedule (check schedule_secrets table or update_schedules.client_secret)');

        context.log(`Credentials: tenant=${tenantId}, clientId=${clientId}, env=${schedule.environment_id}`);

        // 2. Get Power Platform token using THIS schedule's credentials
        const ppToken = await getPowerPlatformToken(tenantId, clientId, clientSecret, context);

        // 3. Discover apps with updates (same API the browser uses)
        const appsToUpdate = await getAppsWithUpdates(schedule.environment_id, ppToken, context);
        context.log(`Found ${appsToUpdate.length} app(s) with updates`);

        if (appsToUpdate.length === 0) {
            context.log('✅ All apps are already up to date');
            result.message = 'All apps are up to date';
        } else {
            // 4. Update each app (same POST the browser uses)
            for (const a of appsToUpdate) {
                const name = a.localizedName || a.applicationName || a.uniqueName || 'Unknown';
                try {
                    await updateApp(schedule.environment_id, a, ppToken, context);
                    result.appsUpdated++;
                    result.apps.push({ name, status: 'success', version: a.latestVersion });
                } catch (err) {
                    result.appsFailed++;
                    result.apps.push({ name, status: 'failed', error: err.message });
                    context.warn(`  ❌ ${name}: ${err.message}`);
                }
            }
        }
    } catch (error) {
        status = 'failed';
        result.error = error.message;
        context.error(`Schedule failed: ${error.message}`);
    }

    await updateScheduleResult(schedule.id, status, result, context);
    context.log(`Result: ${result.appsUpdated} updated, ${result.appsFailed} failed`);
}

// ── Auth ──────────────────────────────────────────────────────

async function getPowerPlatformToken(tenantId, clientId, clientSecret, context) {
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const token = await credential.getToken('https://api.powerplatform.com/.default');
    context.log('Power Platform token acquired');
    return token.token;
}

// ── App discovery (identical to the browser's loadApplications) ──

async function getAppsWithUpdates(environmentId, ppToken, context) {
    const baseUrl = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages`;
    const qs = 'api-version=2022-03-01-preview';

    // Step 1 — installed apps
    context.log('Fetching installed apps…');
    const installed = await fetchAllPages(`${baseUrl}?appInstallState=Installed&${qs}`, ppToken, context);
    context.log(`  ${installed.length} installed apps`);

    // Step 2 — full catalog (contains newer versions)
    context.log('Fetching catalog…');
    const catalog = await fetchAllPages(`${baseUrl}?${qs}`, ppToken, context);
    context.log(`  ${catalog.length} catalog entries`);

    // Build highest-version map keyed by applicationId
    const catalogById = new Map();
    for (const c of catalog) {
        if (!c.applicationId) continue;
        const prev = catalogById.get(c.applicationId);
        if (!prev || cmpVer(c.version, prev.version) > 0) catalogById.set(c.applicationId, c);
    }

    // Build highest-version map keyed by uniqueName base
    const catalogByName = new Map();
    for (const c of catalog) {
        if (!c.uniqueName) continue;
        const base = c.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
        const prev = catalogByName.get(base);
        if (!prev || cmpVer(c.version, prev.version) > 0) catalogByName.set(base, c);
    }

    // Build highest-version map keyed by display name
    const catalogByDisplay = new Map();
    for (const c of catalog) {
        const n = (c.localizedName || c.applicationName || '').toLowerCase();
        if (!n) continue;
        const prev = catalogByDisplay.get(n);
        if (!prev || cmpVer(c.version, prev.version) > 0) catalogByDisplay.set(n, c);
    }

    // Detect updates — same multi-check logic as app.js
    const out = [];
    for (const a of installed) {
        if (a.singlePageApplicationUrl) continue; // SPA = Admin Center only

        let hit = false, latest = null, catName = null;

        // Check state field
        const st = (a.state || '').toLowerCase();
        if (st.includes('update') || st === 'installedwithupdateavailable') hit = true;

        // Check direct version fields
        if (!hit) {
            const dv = a.catalogVersion || a.availableVersion || a.latestVersion || a.newVersion;
            if (dv && cmpVer(dv, a.version) > 0) { hit = true; latest = dv; }
            if (!hit && a.updateAvailable === true) hit = true;
        }

        // Check catalog by applicationId
        if (!hit && a.applicationId) {
            const ce = catalogById.get(a.applicationId);
            if (ce && cmpVer(ce.version, a.version) > 0) { hit = true; latest = ce.version; catName = ce.uniqueName; }
        }

        // Check catalog by uniqueName base
        if (!hit && a.uniqueName) {
            const base = a.uniqueName.replace(/_upgrade$/i, '').replace(/_\d+$/, '');
            const ce = catalogByName.get(base);
            if (ce && cmpVer(ce.version, a.version) > 0) { hit = true; latest = ce.version; catName = ce.uniqueName; }
        }

        // Check catalog by display name
        if (!hit) {
            const n = (a.localizedName || a.applicationName || '').toLowerCase();
            if (n) {
                const ce = catalogByDisplay.get(n);
                if (ce && cmpVer(ce.version, a.version) > 0) { hit = true; latest = ce.version; catName = ce.uniqueName; }
            }
        }

        if (hit) {
            const name = a.localizedName || a.applicationName || a.uniqueName;
            context.log(`  ✓ ${name}  ${a.version} → ${latest || 'newer'}`);
            out.push({ ...a, catalogUniqueName: catName || a.uniqueName, latestVersion: latest });
        }
    }
    return out;
}

// ── App update (identical POST to the browser's reinstallAllApps) ──

async function updateApp(environmentId, a, ppToken, context) {
    const pkg = a.catalogUniqueName || a.uniqueName;
    if (!pkg) throw new Error('No package uniqueName');
    const name = a.localizedName || a.applicationName || a.uniqueName;

    const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${pkg}/install?api-version=2022-03-01-preview`;
    context.log(`  POST ${pkg}`);

    const resp = await fetch(url, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${ppToken}`, 'Content-Type': 'application/json' }
    });

    if (!resp.ok) {
        const body = await resp.text();
        throw new Error(`${resp.status} ${body.substring(0, 200)}`);
    }

    context.log(`  ✅ ${name} — update submitted`);
    await new Promise(r => setTimeout(r, 1500)); // rate-limit guard
}

// ── Utilities ─────────────────────────────────────────────────

async function fetchAllPages(url, token, context) {
    let items = [], next = url, page = 1;
    while (next) {
        const r = await fetch(next, { headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json' } });
        if (!r.ok) { const t = await r.text(); throw new Error(`Page ${page}: ${r.status} ${t.substring(0, 200)}`); }
        const d = await r.json();
        items = items.concat(d.value || []);
        next = d.nextLink || null;
        if (++page > 20) break;
    }
    return items;
}

function cmpVer(a, b) {
    if (!a || !b) return 0;
    const pa = a.split('.').map(n => parseInt(n, 10) || 0);
    const pb = b.split('.').map(n => parseInt(n, 10) || 0);
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
        if ((pa[i] || 0) > (pb[i] || 0)) return 1;
        if ((pa[i] || 0) < (pb[i] || 0)) return -1;
    }
    return 0;
}
