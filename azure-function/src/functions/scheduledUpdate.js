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
        
        context.log(`Current UTC: Day ${currentDayOfWeek}, Time ${currentTimeUtc}`);
        
        try {
            // Get schedules that match current day/time
            const schedules = await getMatchingSchedules(currentDayOfWeek, currentTimeUtc, context);
            
            if (schedules.length === 0) {
                context.log('No schedules match current time. Done.');
                return;
            }
            
            context.log(`Found ${schedules.length} schedule(s) to process`);
            
            for (const schedule of schedules) {
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
    
    const response = await fetch(url, {
        headers: {
            'apikey': config.supabaseKey,
            'Authorization': `Bearer ${config.supabaseKey}`
        }
    });
    
    if (!response.ok) {
        throw new Error(`Failed to fetch schedules: ${response.status}`);
    }
    
    return response.json();
}

async function processSchedule(schedule, context) {
    context.log(`Processing schedule for ${schedule.user_email} in environment ${schedule.environment_id}`);
    
    const startTime = new Date();
    let status = 'success';
    let result = { appsUpdated: 0, appsFailed: 0, apps: [] };
    
    try {
        // Get access token for the environment
        const accessToken = await getAccessToken(schedule.org_url, context);
        
        // Get apps with available updates
        const apps = await getAppsWithUpdates(schedule.org_url, schedule.environment_id, accessToken, context);
        
        context.log(`Found ${apps.length} app(s) with updates available`);
        
        // Update each app
        for (const app of apps) {
            try {
                await updateApp(schedule.org_url, schedule.environment_id, app, accessToken, context);
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
        throw new Error(`Failed to get apps: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    return data.value || [];
}

async function updateApp(orgUrl, environmentId, app, accessToken, context) {
    // Use the AppModuleComponentSource API to trigger update
    // This is a simplified example - actual implementation may need adjustment
    // based on your Power Platform API configuration
    
    context.log(`Updating app: ${app.msdyn_name}`);
    
    // The actual update mechanism depends on how your app integrates with Power Platform
    // This could be:
    // 1. Installing/updating a solution
    // 2. Calling the App Module update API
    // 3. Using the Power Platform Admin API
    
    // For now, log that we would update
    context.log(`Would update app ${app.msdyn_name} (${app.msdyn_componentid})`);
    
    // Simulated delay
    await new Promise(resolve => setTimeout(resolve, 1000));
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
