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
        // Get access tokens - one for Dataverse queries, one for Power Platform API
        const dataverseToken = await getAccessToken(schedule.org_url, context);
        const ppToken = await getPowerPlatformToken(context);
        
        // Get apps with available updates
        const apps = await getAppsWithUpdates(schedule.org_url, schedule.environment_id, dataverseToken, context);
        
        context.log(`Found ${apps.length} app(s) with updates available`);
        
        if (apps.length === 0) {
            context.log('No apps need updating. Skipping update phase.');
        }
        
        // Update each app
        for (const app of apps) {
            try {
                await updateApp(schedule.org_url, schedule.environment_id, app, ppToken, context);
                result.appsUpdated++;
                result.apps.push({ name: app.msdyn_name, status: 'success' });
            } catch (appError) {
                result.appsFailed++;
                result.apps.push({ name: app.msdyn_name, status: 'failed', error: appError.message });
                context.warn(`Failed to update ${app.msdyn_name}: ${appError.message}`);
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

async function getAppsWithUpdates(orgUrl, environmentId, accessToken, context) {
    // Query the organization for apps with updates
    const url = `${orgUrl}/api/data/v9.2/msdyn_solutioncomponentcountssummaries?$filter=msdyn_componenttype eq 300 and msdyn_upgradeavailable eq true`;
    
    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json'
        }
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        context.error(`Failed to get apps: ${response.status} - ${errorText}`);
        throw new Error(`Failed to get apps: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    context.log(`Found ${data.value?.length || 0} apps with msdyn_upgradeavailable=true`);
    return data.value || [];
}

async function updateApp(orgUrl, environmentId, app, ppToken, context) {
    // Use the Power Platform API to install/update the app
    const installUniqueName = app.msdyn_uniquename || app.uniqueName;
    
    if (!installUniqueName) {
        context.warn(`No unique name found for app: ${app.msdyn_name || 'Unknown'}`);
        throw new Error('App missing unique name');
    }
    
    const url = `https://api.powerplatform.com/appmanagement/environments/${environmentId}/applicationPackages/${installUniqueName}/install?api-version=2022-03-01-preview`;
    
    context.log(`Installing/updating app: ${app.msdyn_name} (${installUniqueName})`);
    
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${ppToken}`,
            'Content-Type': 'application/json'
        }
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        context.error(`Failed to update ${app.msdyn_name}: ${response.status} - ${errorText}`);
        throw new Error(`Update failed: ${response.status} - ${errorText}`);
    }
    
    context.log(`✅ Successfully submitted update for ${app.msdyn_name}`);
    
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
