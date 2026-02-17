const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');

// Configuration from environment
const config = {
    tenantId: process.env.AZURE_TENANT_ID,
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    supabaseUrl: process.env.SUPABASE_URL,
    supabaseKey: process.env.SUPABASE_KEY
};

// Timer trigger: runs every hour at minute 5
app.timer('ScheduledUpdateTrigger', {
    schedule: '0 5 * * * *', // Every hour at :05
    handler: async (myTimer, context) => {
        context.log('⏰ Scheduled update check starting...');
        
        const now = new Date();
        const currentDayOfWeek = now.getUTCDay(); // 0 = Sunday
        const currentHour = now.getUTCHours();
        const currentTimeUtc = `${currentHour.toString().padStart(2, '0')}:00`;
        
        const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        context.log(`Current UTC: ${days[currentDayOfWeek]} ${currentTimeUtc} (${now.toISOString()})`);
        
        try {
            // Get schedules that match current day/time
            const schedules = await getMatchingSchedules(currentDayOfWeek, currentTimeUtc, context);
            
            if (schedules.length === 0) {
                context.log(`No schedules match current time (Day: ${currentDayOfWeek}, Time: ${currentTimeUtc}). Done.`);
                return;
            }
            
            context.log(`Found ${schedules.length} schedule(s) to process`);
            
            for (const schedule of schedules) {
                context.log(`Processing schedule ID ${schedule.id} for ${schedule.user_email} (${schedule.timezone || 'UTC'})`);
                await processSchedule(schedule, context);
            }
            
            context.log('✅ Scheduled update check completed');
        } catch (error) {
            context.error('❌ Scheduled update check failed:', error);
        }
    }
});

async function getMatchingSchedules(dayOfWeek, timeUtc, context) {
    const url = `${config.supabaseUrl}/rest/v1/update_schedules?enabled=eq.true&day_of_week=eq.${dayOfWeek}&time_utc=eq.${timeUtc}&select=*`;
    
    context.log(`Querying schedules: ${url}`);
    
    const response = await fetch(url, {
        headers: {
            'apikey': config.supabaseKey,
            'Authorization': `Bearer ${config.supabaseKey}`
        }
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        context.error(`Failed to fetch schedules: ${response.status} - ${errorText}`);
        throw new Error(`Failed to fetch schedules: ${response.status}`);
    }
    
    const schedules = await response.json();
    context.log(`Query returned ${schedules.length} schedule(s)`);
    
    return schedules;
}

async function processSchedule(schedule, context) {
    context.log(`Processing schedule for ${schedule.user_email} in environment ${schedule.environment_id}`);
    
    const startTime = new Date();
    let status = 'success';
    let result = { appsUpdated: 0, appsFailed: 0, apps: [] };
    
    try {
        // Get Power Platform API token
        const ppToken = await getPowerPlatformToken(context);
        
        // Get apps with available updates (using Power Platform API - same as main app)
        const apps = await getAppsWithUpdates(schedule.environment_id, ppToken, context);
        
        context.log(`Found ${apps.length} app(s) with updates available`);
        
        if (apps.length === 0) {
            context.log('✅ All apps are up to date!');
            status = 'success';
            result.message = 'All apps are up to date';
        } else {
            // Update each app
            for (const app of apps) {
                const appName = app.localizedName || app.applicationName || app.uniqueName || 'Unknown';
                try {
                    await updateApp(schedule.environment_id, app, ppToken, context);
                    result.appsUpdated++;
                    result.apps.push({ name: appName, status: 'success', version: app.latestVersion });
                    context.log(`  ✅ ${appName} - update submitted`);
                } catch (appError) {
                    result.appsFailed++;
                    result.apps.push({ name: appName, status: 'failed', error: appError.message });
                    context.warn(`  ❌ ${appName} - failed: ${appError.message}`);
                }
            }
        }
        
    } catch (error) {
        status = 'failed';
        result.error = error.message;
        context.error(`Schedule processing failed: ${error.message}`);
    }
    
    // Update schedule with results
    await updateScheduleResult(schedule.id, status, result, context);
    
    context.log(`Schedule completed: ${result.appsUpdated} updated, ${result.appsFailed} failed`);
}

// NOTE: This function is kept for potential future use if Dataverse API access is needed
// Currently, we use only the Power Platform API for app detection and updates
async function getAccessToken(orgUrl, context) {
    // Use client credentials flow with service principal
    const credential = new ClientSecretCredential(
        config.tenantId,
        config.clientId,
        config.clientSecret
    );
    
    // Get token for the Dynamics CRM API
    // The scope should be the org URL + /.default for client credentials
    const scope = `${orgUrl}/.default`;
    
    const token = await credential.getToken(scope);
    return token.token;
}

async function getPowerPlatformToken(context) {
    // Get token for Power Platform API
    const credential = new ClientSecretCredential(
        config.tenantId,
        config.clientId,
        config.clientSecret
    );
    
    // Power Platform API scope
    const token = await credential.getToken('https://api.powerplatform.com/.default');
    return token.token;
}

async function getAppsWithUpdates(environmentId, ppToken, context) {
    // Use the same Power Platform API that the main app uses
    const baseUrl = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages`;
    const apiVersion = 'api-version=2022-03-01-preview';
    
    context.log('Fetching installed apps from Power Platform API...');
    const installedApps = await fetchAllPages(`${baseUrl}?appInstallState=Installed&${apiVersion}`, ppToken, context);
    context.log(`Found ${installedApps.length} installed apps`);
    
    context.log('Fetching catalog packages...');
    const allCatalog = await fetchAllPages(`${baseUrl}?${apiVersion}`, ppToken, context);
    context.log(`Found ${allCatalog.length} catalog packages`);
    
    // Build version map from catalog by applicationId
    const catalogMapById = new Map();
    for (const app of allCatalog) {
        if (!app.applicationId) continue;
        const existing = catalogMapById.get(app.applicationId);
        if (!existing || compareVersions(app.version, existing.version) > 0) {
            catalogMapById.set(app.applicationId, app);
        }
    }
    
    // Detect apps with updates (same logic as main app)
    const appsWithUpdates = [];
    for (const app of installedApps) {
        // Skip apps that require Admin Center (SPA)
        if (app.singlePageApplicationUrl) {
            context.log(`  Skipping ${app.localizedName || app.uniqueName} - requires Admin Center`);
            continue;
        }
        
        let hasUpdate = false;
        let latestVersion = null;
        let catalogUniqueName = null;
        
        // Check 1: Direct API fields
        if (app.updateAvailable || app.catalogVersion || app.availableVersion) {
            const directVersion = app.catalogVersion || app.availableVersion;
            if (directVersion && compareVersions(directVersion, app.version) > 0) {
                hasUpdate = true;
                latestVersion = directVersion;
            } else if (app.updateAvailable === true) {
                hasUpdate = true;
            }
        }
        
        // Check 2: Compare with catalog by applicationId
        if (!hasUpdate && app.applicationId) {
            const catalogEntry = catalogMapById.get(app.applicationId);
            if (catalogEntry && compareVersions(catalogEntry.version, app.version) > 0) {
                hasUpdate = true;
                latestVersion = catalogEntry.version;
                catalogUniqueName = catalogEntry.uniqueName;
            }
        }
        
        if (hasUpdate) {
            context.log(`  ✓ Update available: ${app.localizedName || app.uniqueName} ${app.version} → ${latestVersion || 'newer'}`);
            appsWithUpdates.push({
                ...app,
                catalogUniqueName: catalogUniqueName || app.uniqueName,
                latestVersion: latestVersion
            });
        }
    }
    
    return appsWithUpdates;
}

async function fetchAllPages(url, token, context) {
    // Fetch paginated results from Power Platform API
    let allItems = [];
    let nextLink = url;
    let pageNum = 1;
    
    while (nextLink) {
        const response = await fetch(nextLink, {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Accept': 'application/json'
            }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`API error on page ${pageNum}: ${response.status} - ${errorText}`);
        }
        
        const data = await response.json();
        const items = data.value || [];
        allItems = allItems.concat(items);
        
        nextLink = data.nextLink || null;
        pageNum++;
        
        // Safety limit
        if (pageNum > 20) {
            context.warn('Reached page limit of 20, stopping pagination');
            break;
        }
    }
    
    return allItems;
}

function compareVersions(v1, v2) {
    // Compare two version strings (e.g., "1.2.3" vs "1.2.4")
    if (!v1 || !v2) return 0;
    
    const parts1 = v1.split('.').map(p => parseInt(p, 10) || 0);
    const parts2 = v2.split('.').map(p => parseInt(p, 10) || 0);
    const maxLen = Math.max(parts1.length, parts2.length);
    
    for (let i = 0; i < maxLen; i++) {
        const p1 = parts1[i] || 0;
        const p2 = parts2[i] || 0;
        if (p1 > p2) return 1;
        if (p1 < p2) return -1;
    }
    
    return 0;
}

async function updateApp(environmentId, app, ppToken, context) {
    // Use the Power Platform API to install/update the app (same as main app)
    const installUniqueName = app.catalogUniqueName || app.uniqueName;
    
    if (!installUniqueName) {
        const appName = app.localizedName || app.applicationName || 'Unknown';
        context.warn(`No unique name found for app: ${appName}`);
        throw new Error('App missing unique name');
    }
    
    const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${installUniqueName}/install?api-version=2022-03-01-preview`;
    
    const appName = app.localizedName || app.applicationName || app.uniqueName;
    context.log(`Installing/updating: ${appName} (${installUniqueName})`);
    
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${ppToken}`,
            'Content-Type': 'application/json'
        }
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        context.error(`Failed to update ${appName}: ${response.status} - ${errorText}`);
        throw new Error(`Update failed: ${response.status} - ${errorText}`);
    }
    
    context.log(`✅ Successfully submitted update for ${appName}`);
    
    // Small delay to avoid rate limiting
    await new Promise(resolve => setTimeout(resolve, 1500));
}

async function updateScheduleResult(scheduleId, status, result, context) {
    const url = `${config.supabaseUrl}/rest/v1/update_schedules?id=eq.${scheduleId}`;
    
    const response = await fetch(url, {
        method: 'PATCH',
        headers: {
            'apikey': config.supabaseKey,
            'Authorization': `Bearer ${config.supabaseKey}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            last_run_at: new Date().toISOString(),
            last_run_status: status,
            last_run_result: result
        })
    });
    
    if (!response.ok) {
        context.warn(`Failed to update schedule result: ${response.status}`);
    }
}
