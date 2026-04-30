// Global variables
let msalInstance = null;

let ppToken = null; // Power Platform API tok
en
let environmentId = null;
let currentOrgUr
l = null;
let apps = [];
let allEnvironments 
= []; // Cached list of all environments
let 
selectedApps = new Set(); // Multi-select tra
cking

// ═══════════�
�══════════════�
�══════════════�
�══════════════�
�══════════════�
�═══
// LOGGING SYSTEM - Persists acros
s redirects for debugging auth flows
// ═�
�══════════════�
�══════════════�
�══════════════�
�══════════════�
�═════════════
cons
t LOG_STORAGE_KEY = 'd365_app_logs';
const MA
X_LOGS = 200;

function appLog(level, message
, data = null) {
    const timestamp = new Da
te().toISOString();
    const logEntry = { ti
mestamp, level, message, data: data ? JSON.st
ringify(data) : null };
    
    // Console o
utput
    const consoleMsg = `[${timestamp}] 
[${level}] ${message}`;
    if (level === 'ER
ROR') {
        console.error(consoleMsg, dat
a || '');
    } else if (level === 'WARN') {

        console.warn(consoleMsg, data || '');

    } else {
        console.log(consoleMsg,
 data || '');
    }
    
    // Persist to se
ssionStorage (survives redirects within same 
session)
    try {
        const logs = JSON.
parse(sessionStorage.getItem(LOG_STORAGE_KEY)
 || '[]');
        logs.push(logEntry);
     
   // Keep only last MAX_LOGS entries
       
 while (logs.length > MAX_LOGS) logs.shift();

        sessionStorage.setItem(LOG_STORAGE_K
EY, JSON.stringify(logs));
    } catch (e) {

        console.error('Failed to persist log:
', e);
    }
    
    // Update UI log panel 
if visible
    updateLogPanel();
}

function 
updateLogPanel() {
    const panel = document
.getElementById('logPanel');
    const conten
t = document.getElementById('logContent');
  
  if (!panel || !content) return;
    
    tr
y {
        const logs = JSON.parse(sessionSt
orage.getItem(LOG_STORAGE_KEY) || '[]');
    
    content.innerHTML = logs.map(log => {
   
         const color = log.level === 'ERROR' 
? '#ff4444' : log.level === 'WARN' ? '#ffaa00
' : '#00cc00';
            const time = log.t
imestamp.split('T')[1].split('.')[0];
       
     const dataStr = log.data ? `\n    └─
 ${log.data}` : '';
            return `<div 
style="color:${color};margin:2px 0;"><span st
yle="color:#888">[${time}]</span> <b>[${log.l
evel}]</b> ${log.message}${dataStr}</div>`;
 
       }).join('');
        content.scrollTop
 = content.scrollHeight;
    } catch (e) {}
}


function clearLogs() {
    sessionStorage.r
emoveItem(LOG_STORAGE_KEY);
    updateLogPane
l();
    appLog('INFO', 'Logs cleared');
}

f
unction exportLogs() {
    try {
        cons
t logs = JSON.parse(sessionStorage.getItem(LO
G_STORAGE_KEY) || '[]');
        const text =
 logs.map(l => `[${l.timestamp}] [${l.level}]
 ${l.message}${l.data ? ' | ' + l.data : ''}`
).join('\n');
        const blob = new Blob([
text], { type: 'text/plain' });
        const
 url = URL.createObjectURL(blob);
        con
st a = document.createElement('a');
        a
.href = url;
        a.download = `d365-app-l
ogs-${new Date().toISOString().replace(/[:.]/
g, '-')}.txt`;
        a.click();
        URL
.revokeObjectURL(url);
    } catch (e) {
    
    alert('Failed to export logs: ' + e.messa
ge);
    }
}

function toggleLogPanel() {
   
 const panel = document.getElementById('logPa
nel');
    if (panel) {
        const isHidde
n = panel.style.display === 'none';
        p
anel.style.display = isHidden ? 'block' : 'no
ne';
        if (isHidden) updateLogPanel();

    }
}

// Shorthand logging functions
const
 logInfo = (msg, data) => appLog('INFO', msg,
 data);
const logWarn = (msg, data) => appLog
('WARN', msg, data);
const logError = (msg, d
ata) => appLog('ERROR', msg, data);
const log
Debug = (msg, data) => appLog('DEBUG', msg, d
ata);

// ═══════════�
�══════════════�
�══════════════�
�══════════════�
�══════════════�
�═══

// Supabase config for usage trac
king
const SUPABASE_URL = 'https://fpekzltxuk
ikaixebeeu.supabase.co';
const SUPABASE_KEY =
 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3
MiOiJzdXBhYmFzZSIsInJlZiI6ImZwZWt6bHR4dWtpa2F
peGViZWV1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzA0
MDU0ODEsImV4cCI6MjA4NTk4MTQ4MX0.uH4JgKbf_-Al_
iArzEy6UZ3edJNzFSCBVlMNI04li0Y';

// MSAL Con
figuration
function createMsalConfig(tenantId
, clientId) {
    // Compute redirect URI fro
m current path (handles GitHub Pages subpaths
 like /d365-app-updater/)
    const pathDir =
 window.location.pathname.substring(0, window
.location.pathname.lastIndexOf('/') + 1);
   
 // Use 'organizations' endpoint if no tenant
 ID provided (allows any Azure AD account)
  
  const authority = tenantId 
        ? `http
s://login.microsoftonline.com/${tenantId}`
  
      : 'https://login.microsoftonline.com/or
ganizations';
    return {
        auth: {
  
          clientId: clientId,
            aut
hority: authority,
            redirectUri: w
indow.location.origin + pathDir,
        },
 
       cache: {
            cacheLocation: 'l
ocalStorage',
            storeAuthStateInCoo
kie: false,
        },
    };
}

// Flag to t
rack if we're resuming from a redirect
let _p
endingRedirectAuth = false;

// Initialize on
 page load
document.addEventListener('DOMCont
entLoaded', async function() {
    // ══�
�� EMERGENCY RESET: Add ?reset to URL to clea
r all auth state ═══
    const urlParam
s = new URLSearchParams(window.location.searc
h);
    if (urlParams.has('reset')) {
       
 console.log('🔄 RESET MODE: Clearing all a
uthentication state...');
        localStorag
e.removeItem('d365_app_updater_creds');
     
   sessionStorage.removeItem('d365_auth_step'
);
        sessionStorage.removeItem('d365_re
direct_count');
        sessionStorage.remove
Item(LOG_STORAGE_KEY);
        // Clear MSAL 
cache
        localStorage.removeItem('msal.t
oken.keys.' + urlParams.get('reset'));
      
  // Clear all MSAL-related items
        Obj
ect.keys(localStorage).forEach(key => {
     
       if (key.startsWith('msal.')) localStor
age.removeItem(key);
        });
        Obje
ct.keys(sessionStorage).forEach(key => {
    
        if (key.startsWith('msal.') || key.st
artsWith('d365_') || key.startsWith('wizard_'
)) {
                sessionStorage.removeIte
m(key);
            }
        });
        ale
rt('✅ All authentication data cleared! You 
can now start fresh.');
        // Redirect t
o clean URL
        window.location.href = wi
ndow.location.origin + window.location.pathna
me;
        return;
    }

    // â•â�
�â• SSO from data-gen: override stale 
localStorage with URL params BEFORE any auth 
init â•â•â•
    if (urlParams
.get('autoConnect') === '1' && urlParams.get(
'orgUrl')) {
        const ssoOrgUrl   = url
Params.get('orgUrl');
        const ssoClien
tId = urlParams.get('clientId') || (typeof SH
ARED_CLIENT_ID !== 'undefined' ? SHARED_CLIEN
T_ID : '');
        const ssoTenantId = urlP
arams.get('tenantId') || 'organizations';
  
      // Clobber any stale cached creds so ha
ndleRedirectResponse() uses the right client

        localStorage.setItem('d365_app_updat
er_creds', JSON.stringify({
            orgU
rl: ssoOrgUrl, tenantId: ssoTenantId, clientI
d: ssoClientId
        }));
        // Also
 wipe any stale MSAL cache that might trigger
 handleRedirectPromise on the wrong client
 
       Object.keys(localStorage).forEach(k =>
 { if (k.startsWith('msal.')) localStorage.re
moveItem(k); });
        Object.keys(session
Storage).forEach(k => { if (k.startsWith('msa
l.')) sessionStorage.removeItem(k); });
    
    console.log('[SSO] Overrode cached creds 
with URL params, clientId=' + (ssoClientId||'
').substring(0,8) + '...');
    }

    log
Info('=== APP INITIALIZATION ===');
    logIn
fo('URL', { href: window.location.href, hash:
 window.location.hash ? 'present' : 'none', p
athname: window.location.pathname });
    log
Info('Auth step in storage', sessionStorage.g
etItem('d365_auth_step'));
    logInfo('Saved
 creds', { 
        sessionStorage: !!session
Storage.getItem('d365_app_updater_creds_temp'
), 
        localStorage: !!localStorage.getI
tem('d365_app_updater_creds') 
    });

    /
/ If returning from wizard auth redirect, han
dle MSAL here on the root page
    // (MSAL r
equires redirect response to be processed on 
the same URL that was used as redirectUri)
  
  const wizardClientId = sessionStorage.getIt
em('wizard_clientId');
    if (wizardClientId
) {
        logInfo('Wizard flow detected', {
 wizardClientId: wizardClientId.substring(0, 
8) + '...', hasHash: !!window.location.hash }
);
        
        // First check if the has
h contains an error
        if (window.locati
on.hash && window.location.hash.includes('err
or')) {
            const hashParams = new UR
LSearchParams(window.location.hash.substring(
1));
            const error = hashParams.get
('error');
            const errorDesc = deco
deURIComponent(hashParams.get('error_descript
ion') || 'Unknown error');
            logErr
or('Wizard auth error in URL', { error, error
Desc });
            sessionStorage.setItem('
wizard_error', errorDesc);
            sessio
nStorage.removeItem('wizard_clientId');
     
       window.location.replace('setup-wizard.
html');
            return;
        }
       
 
        // Only process if we have a hash (
returning from redirect)
        if (window.l
ocation.hash) {
            logInfo('Processi
ng wizard redirect...');
            try {
  
              const pathDir = window.location
.pathname.substring(0, window.location.pathna
me.lastIndexOf('/') + 1);
                con
st wizardMsalConfig = {
                    a
uth: {
                        clientId: wiza
rdClientId,
                        authority
: 'https://login.microsoftonline.com/organiza
tions',
                        redirectUri: 
window.location.origin + pathDir
            
        },
                    cache: { cache
Location: 'sessionStorage', storeAuthStateInC
ookie: false }
                };
           
     const wizardMsal = new msal.PublicClient
Application(wizardMsalConfig);
              
  await wizardMsal.initialize();
            
    logInfo('Wizard MSAL initialized, calling
 handleRedirectPromise...');
                

                const response = await wizar
dMsal.handleRedirectPromise();
              
  logInfo('Wizard handleRedirectPromise resul
t', response ? { hasToken: !!response.accessT
oken, scopes: response.scopes } : 'null');
  
              
                if (response &
& response.accessToken) {
                   
 sessionStorage.setItem('wizard_accessToken',
 response.accessToken);
                    l
ogInfo('Wizard token saved to sessionStorage'
);
                } else {
                 
   // No response from redirect, try to get t
oken silently if there's an account
         
           const accounts = wizardMsal.getAll
Accounts();
                    logInfo('Wiza
rd accounts', accounts.length);
             
       if (accounts.length > 0) {
           
             try {
                          
  const silentResult = await wizardMsal.acqui
reTokenSilent({
                             
   scopes: ['https://graph.microsoft.com/Appl
ication.ReadWrite.All', 'https://graph.micros
oft.com/DelegatedPermissionGrant.ReadWrite.Al
l'],
                                account:
 accounts[0]
                            });

                            if (silentResult 
&& silentResult.accessToken) {
              
                  sessionStorage.setItem('wiz
ard_accessToken', silentResult.accessToken);

                                logInfo('Wiza
rd token acquired silently');
               
             }
                        } catc
h (silentErr) {
                            l
ogWarn('Wizard silent token acquisition faile
d', silentErr.message);
                     
   }
                    }
                }

                
                // Check if 
we got a token
                if (!sessionSt
orage.getItem('wizard_accessToken')) {
      
              logError('No wizard token acqui
red after redirect processing');
            
        sessionStorage.setItem('wizard_error'
, 'Failed to acquire access token. Please ens
ure you have admin permissions and try again.
');
                }
            } catch (er
r) {
                logError('Wizard redirec
t handling error', err.message);
            
    sessionStorage.setItem('wizard_error', er
r.message);
            }
            // Forw
ard to setup-wizard.html (without the hash)
 
           window.location.replace('setup-wiz
ard.html');
            return;
        } els
e {
            // wizardClientId is set but 
no hash - user might be stuck, clear and let 
them restart
            logWarn('Wizard clie
ntId set but no hash, clearing wizard state')
;
            sessionStorage.removeItem('wiza
rd_clientId');
        }
    }

    hideLoadi
ng();
    
    if (typeof msal === 'undefined
') {
        logError('MSAL library failed to
 load');
        alert('Error: MSAL library f
ailed to load.');
        return;
    }
    l
ogInfo('MSAL library loaded successfully');
 
   
    const redirectUriElement = document.g
etElementById('redirectUri');
    if (redirec
tUriElement) {
        redirectUriElement.tex
tContent = window.location.origin;
    }
    

    loadSavedCredentials();
    
    documen
t.getElementById('authForm').addEventListener
('submit', handleAuthentication);
    documen
t.getElementById('logoutBtn').addEventListene
r('click', handleLogout);
    document.getEle
mentById('refreshAppsBtn').addEventListener('
click', loadApplications);
    document.getEl
ementById('updateAllBtn').addEventListener('c
lick', updateAllApps);
    document.getElemen
tById('reinstallAllBtn').addEventListener('cl
ick', reinstallAllApps);
    document.getElem
entById('updateSelectedBtn').addEventListener
('click', updateSelectedApps);
    
    // Cl
ose environment dropdown when clicking outsid
e
    document.addEventListener('click', func
tion(e) {
        const switcher = document.q
uerySelector('.env-switcher');
        if (sw
itcher && !switcher.contains(e.target)) {
   
         closeEnvDropdown();
        }
    })
;
    
    // Try to handle redirect response
 first, then fall back to auto-login
    hand
leRedirectResponse().then(() => {
        con
sole.log('App initialized');
        trySsoAu
toConnect();
    });
});

// Load saved crede
ntials
function loadSavedCredentials() {
    
const savedCreds = localStorage.getItem('d365
_app_updater_creds');
    if (savedCreds) {
 
       try {
            const creds = JSON.p
arse(savedCreds);
            const orgUrlEl 
= document.getElementById('orgUrl');
        
    if (orgUrlEl) orgUrlEl.value = creds.orgU
rl || creds.organizationId || creds.environme
ntId || '';
            document.getElementBy
Id('clientId').value = creds.clientId || '';

            document.getElementById('remember
Me').checked = true;
        } catch (e) {}
 
   }
}

// Handle redirect response when retu
rning from Microsoft login redirect
async fun
ction handleRedirectResponse() {
    logInfo(
'=== handleRedirectResponse START ===');
    

    // Check if URL hash contains an error r
esponse from Azure AD
    const hash = window
.location.hash;
    if (hash && (hash.include
s('error=') || hash.includes('error_descripti
on='))) {
        logError('Azure AD returned
 an error in the URL');
        
        // T
ry to decode the error
        const hashPara
ms = new URLSearchParams(hash.substring(1));

        const error = hashParams.get('error')
;
        const errorDesc = decodeURIComponen
t(hashParams.get('error_description') || 'Unk
nown error');
        
        logError('Azur
e AD Error', { error, errorDesc });
        

        // Clear all auth state
        local
Storage.removeItem('d365_app_updater_creds');

        sessionStorage.removeItem('d365_auth
_step');
        sessionStorage.removeItem('d
365_redirect_count');
        Object.keys(loc
alStorage).forEach(key => {
            if (k
ey.startsWith('msal.')) localStorage.removeIt
em(key);
        });
        Object.keys(sess
ionStorage).forEach(key => {
            if (
key.startsWith('msal.')) sessionStorage.remov
eItem(key);
        });
        
        // C
lean the URL
        history.replaceState(nul
l, '', window.location.pathname);
        
  
      // Show error
        const resetButton
 = `<br><br><button onclick="window.location.
reload()" style="background:#0078d4;color:whi
te;border:none;padding:10px 20px;border-radiu
s:6px;cursor:pointer;font-weight:600;">🔄 S
tart Fresh</button>`;
        
        let fr
iendlyMessage = errorDesc;
        if (errorD
esc.includes('AADSTS650057') || errorDesc.inc
ludes('not listed in the requested permission
s')) {
            friendlyMessage = `<strong
>Missing API Permissions</strong><br><br>
You
r Azure AD app registration is missing requir
ed permissions.<br><br><strong>Raw error:</st
rong> <code style='font-size:11px;word-break:
break-all'>${errorDesc}</code><br><br>
<stron
g>To fix this:</strong><br>
1. Go to <a href=
"https://portal.azure.com/#view/Microsoft_AAD
_RegisteredApps/ApplicationsListBlade" target
="_blank" rel="noopener">Azure Portal → App
 Registrations</a><br>
2. Find and click on y
our app<br>
3. Go to <strong>API permissions<
/strong> → <strong>Add a permission</strong
><br>
4. Add: <strong>Power Platform API</str
ong> → Delegated → <code>user_impersonati
on</code><br>
5. Add: <strong>Dynamics CRM</s
trong> → Delegated → <code>user_impersona
tion</code><br>
6. Click <strong>Grant admin 
consent</strong>${resetButton}`;
        } el
se {
            friendlyMessage = `<strong>A
uthentication Error</strong><br><br>${errorDe
sc}${resetButton}`;
        }
        
      
  showError(friendlyMessage);
        return;

    }
    
    // Try to get credentials fro
m sessionStorage first (survives redirect, wo
rks with tracking prevention)
    // then fal
l back to localStorage (for "remember me" fun
ctionality)
    let savedCreds = sessionStora
ge.getItem('d365_app_updater_creds_temp');
  
  let credsSource = 'sessionStorage';
    
  
  if (!savedCreds) {
        savedCreds = loc
alStorage.getItem('d365_app_updater_creds');

        credsSource = 'localStorage';
    }
 
   
    if (!savedCreds) {
        logInfo('N
o saved credentials in sessionStorage or loca
lStorage, user must log in manually');
      
  return;
    }
    
    logInfo('Found crede
ntials in ' + credsSource);

    let creds;
 
   try {
        creds = JSON.parse(savedCred
s);
        logDebug('Parsed saved credential
s', { orgUrl: creds.orgUrl, tenantId: creds.t
enantId?.substring(0,8) + '...', hasClientId:
 !!creds.clientId });
    } catch (e) {
     
   logError('Failed to parse saved credential
s', e.message);
        return;
    }

    co
nst orgUrlValue = creds.orgUrl || creds.organ
izationId || creds.environmentId || '';
    c
onst tenantId = creds.tenantId || '';
    con
st clientId = creds.clientId || '';
    
    
// Tenant ID is optional - only Client ID and
 Org URL are required
    if (!clientId || !o
rgUrlValue) {
        logWarn('Missing requir
ed credentials', { hasOrgUrl: !!orgUrlValue, 
hasTenantId: !!tenantId, hasClientId: !!clien
tId });
        return;
    }

    try {
    
    logInfo('Creating MSAL instance...');
   
     const msalConfig = createMsalConfig(tena
ntId, clientId);
        logDebug('MSAL confi
g', { redirectUri: msalConfig.auth.redirectUr
i, authority: msalConfig.auth.authority });
 
       
        msalInstance = new msal.Publi
cClientApplication(msalConfig);
        await
 msalInstance.initialize();
        logInfo('
MSAL initialized');

        // Check if we'r
e returning from a redirect login
        log
Info('Calling handleRedirectPromise...');
   
     const redirectResult = await msalInstanc
e.handleRedirectPromise();
        logInfo('h
andleRedirectPromise result', redirectResult 
? { 
            hasToken: !!redirectResult.a
ccessToken, 
            scopes: redirectResu
lt.scopes,
            account: redirectResul
t.account?.username 
        } : 'null');

  
      const accounts = msalInstance.getAllAcc
ounts();
        logInfo('MSAL accounts found
', { count: accounts.length, accounts: accoun
ts.map(a => a.username) });
        
        
if (accounts.length === 0) {
            logI
nfo('No accounts in cache, user must log in m
anually');
            msalInstance = null;
 
           sessionStorage.removeItem('d365_au
th_step');
            return;
        }

   
     const account = accounts[0];
        log
Info('Using account', account.username);

   
     // Track which step of auth we're on to 
avoid redirect loops
        const authStep =
 sessionStorage.getItem('d365_auth_step') || 
'none';
        logInfo('Current auth step', 
authStep);
        
        // Check redirect
 counter to prevent infinite loops
        le
t redirectCount = parseInt(sessionStorage.get
Item('d365_redirect_count') || '0', 10);
    
    logInfo('Redirect count', redirectCount);

        
        if (redirectCount > 5) {
  
          logError('Too many redirects, break
ing loop');
            sessionStorage.remove
Item('d365_auth_step');
            sessionSt
orage.removeItem('d365_redirect_count');
    
        throw new Error('Authentication faile
d after multiple attempts. Please clear your 
browser cache and try again.');
        }

  
      // If returning from a runtime BAP toke
n redirect, clear step and continue
        i
f (authStep === 'acquiring_bap_runtime' && re
directResult) {
            logInfo('Returned
 from runtime BAP token redirect, clearing st
ep');
            sessionStorage.removeItem('
d365_auth_step');
        }

        // If we
 just came back from login (initial or login_
redirect step), the redirectResult is for the
 login
        // We need to proceed to acqui
re PP and BAP tokens
        if ((authStep ==
= 'login_redirect' || authStep === 'initial')
 && redirectResult) {
            logInfo('Re
turned from initial login redirect, proceedin
g to acquire API tokens');
            sessio
nStorage.removeItem('d365_auth_step'); // Cle
ar so we can proceed
        }

        showL
oading('Authenticating...', 'Restoring your s
ession');

        // ─── Acquire Power
 Platform token ─────────�
�──────────────�
�──────────────
 
       logInfo('Acquiring Power Platform toke
n...');
        const ppRequest = { scopes: [
'https://api.powerplatform.com/.default'], ac
count };
        let ppResult;
        
     
   // Check if redirect result contains PP sc
ope
        const isPPRedirectResult = redire
ctResult && redirectResult.scopes && 
       
     redirectResult.scopes.some(s => s.includ
es('api.powerplatform.com'));
        
      
  if (authStep === 'acquiring_pp' && isPPRedi
rectResult && redirectResult.accessToken) {
 
           ppResult = redirectResult;
       
     logInfo('Using PP token from redirect re
sult', { scopes: redirectResult.scopes });
  
          sessionStorage.removeItem('d365_aut
h_step');
            sessionStorage.removeIt
em('d365_redirect_count');
        } else {
 
           try {
                logDebug('Tr
ying acquireTokenSilent for PP...');
        
        ppResult = await msalInstance.acquire
TokenSilent(ppRequest);
                logIn
fo('PP token acquired silently');
           
 } catch (e) {
                logWarn('acqui
reTokenSilent for PP failed', { error: e.mess
age, errorCode: e.errorCode });
             
   // Only redirect if we haven't tried too m
any times
                if (authStep !== 'a
cquiring_pp') {
                    logInfo('
Redirecting for PP token consent...');
      
              sessionStorage.setItem('d365_au
th_step', 'acquiring_pp');
                  
  sessionStorage.setItem('d365_redirect_count
', String(redirectCount + 1));
              
      await msalInstance.acquireTokenRedirect
(ppRequest);
                    return;
    
            } else {
                    logE
rror('Already tried PP redirect, still failin
g', e.message);
                    sessionSt
orage.removeItem('d365_auth_step');
         
           sessionStorage.removeItem('d365_re
direct_count');
                    throw new
 Error('Failed to acquire Power Platform toke
n. Error: ' + e.message + '. Please check you
r app registration has the correct API permis
sions.');
                }
            }
   
     }
        ppToken = ppResult.accessToken
;
        logInfo('PP token acquired successf
ully');

        // ─── Acquire BAP tok
en ──────────────
───────────────
───────────────
──────
        logInfo('Acquiring
 BAP token...');
        const bapRequest = {
 scopes: ['https://api.bap.microsoft.com/.def
ault'], account };
        
        // Check 
if redirect result contains BAP scope
       
 const isBAPRedirectResult = redirectResult &
& redirectResult.scopes && 
            redir
ectResult.scopes.some(s => s.includes('api.ba
p.microsoft.com'));
        
        if (auth
Step === 'acquiring_bap' && isBAPRedirectResu
lt && redirectResult.accessToken) {
         
   logInfo('Using BAP token from redirect res
ult', { scopes: redirectResult.scopes });
   
         sessionStorage.removeItem('d365_auth
_step');
            sessionStorage.removeIte
m('d365_redirect_count');
        } else {
  
          try {
                logDebug('Try
ing acquireTokenSilent for BAP...');
        
        await msalInstance.acquireTokenSilent
(bapRequest);
                logInfo('BAP to
ken acquired silently');
            } catch 
(e) {
                logWarn('acquireTokenSi
lent for BAP failed', { error: e.message, err
orCode: e.errorCode });
                if (a
uthStep !== 'acquiring_bap') {
              
      logInfo('Redirecting for BAP token cons
ent...');
                    sessionStorage.
setItem('d365_auth_step', 'acquiring_bap');
 
                   sessionStorage.setItem('d3
65_redirect_count', String(redirectCount + 1)
);
                    await msalInstance.acq
uireTokenRedirect(bapRequest);
              
      return;
                } else {
      
              logError('Already tried BAP red
irect, still failing', e.message);
          
          sessionStorage.removeItem('d365_aut
h_step');
                    sessionStorage.
removeItem('d365_redirect_count');
          
          throw new Error('Failed to acquire 
BAP token. Error: ' + e.message + '. Please c
heck your app registration has the correct AP
I permissions.');
                }
         
   }
        }

        // Clear auth step - 
we're done with redirects
        sessionStor
age.removeItem('d365_auth_step');
        log
Info('Auth step cleared, proceeding to resolv
e environment');

        showLoading('Authen
ticating...', 'Resolving environment');

    
    // Normalize org URL
        let normaliz
edOrgUrl = orgUrlValue;
        if (!normaliz
edOrgUrl.startsWith('https://')) {
          
  normalizedOrgUrl = 'https://' + normalizedO
rgUrl;
        }
        normalizedOrgUrl = n
ormalizedOrgUrl.replace(/\/+$/, '');
        
logInfo('Resolving environment for URL', norm
alizedOrgUrl);

        environmentId = await
 resolveOrgUrlToEnvironmentId(normalizedOrgUr
l);
        if (!environmentId) {
           
 throw new Error('Could not resolve environme
nt. Please verify the Organization URL and yo
ur permissions.');
        }
        logInfo(
'Environment resolved', environmentId);

    
    currentOrgUrl = normalizedOrgUrl;

      
  showLoading('Authenticating...', 'Loading e
nvironment details');
        await getEnviro
nmentName();

        hideLoading();

       
 document.getElementById('authSection').class
List.add('hidden');
        document.getEleme
ntById('appsSection').classList.remove('hidde
n');

        // Load schedule settings
     
   loadSchedule();

        // Clean up temp 
credentials from sessionStorage
        sessi
onStorage.removeItem('d365_app_updater_creds_
temp');
        
        // If user wanted to
 remember, try to save to localStorage (may f
ail with tracking prevention)
        if (cre
ds.rememberMe) {
            try {
          
      localStorage.setItem('d365_app_updater_
creds', JSON.stringify({ orgUrl: creds.orgUrl
, tenantId: creds.tenantId, clientId: creds.c
lientId }));
            } catch (e) {
      
          logWarn('Could not save to localSto
rage for remember me', e.message);
          
  }
        }

        logInfo('=== AUTH SUCC
ESS ===', account.username);
        await lo
adApplications();

    } catch (e) {
        
sessionStorage.removeItem('d365_auth_step');

        sessionStorage.removeItem('d365_redir
ect_count');
        sessionStorage.removeIte
m('d365_app_updater_creds_temp'); // Clean up
 temp creds on error too
        logError('==
= AUTH FAILED ===', e.message);
        hideL
oading();
        
        // CRITICAL: Clear
 ALL saved state to prevent auto-login loop
 
       localStorage.removeItem('d365_app_upda
ter_creds');
        
        // Clear ALL MS
AL cache (localStorage AND sessionStorage)
  
      Object.keys(localStorage).forEach(key =
> {
            if (key.startsWith('msal.')) 
localStorage.removeItem(key);
        });
   
     Object.keys(sessionStorage).forEach(key 
=> {
            if (key.startsWith('msal.'))
 sessionStorage.removeItem(key);
        });

        
        // Clean URL hash if present
 (prevents re-processing on next attempt)
   
     if (window.location.hash) {
            
history.replaceState(null, '', window.locatio
n.pathname);
        }
        
        // Ma
ke sure auth section is visible
        docum
ent.getElementById('authSection').classList.r
emove('hidden');
        document.getElementB
yId('appsSection').classList.add('hidden');
 
       
        // Clear the form's "remember
 me" checkbox
        const rememberMe = docu
ment.getElementById('rememberMe');
        if
 (rememberMe) rememberMe.checked = false;
   
     
        // Provide helpful error messag
es for common issues
        let errorMessage
 = e.message;
        const resetButton = `<b
r><br><button onclick="window.location.reload
()" style="background:#0078d4;color:white;bor
der:none;padding:10px 20px;border-radius:6px;
cursor:pointer;font-weight:600;">🔄 Start F
resh</button>`;
        
        if (e.messag
e.includes('AADSTS650057') || e.message.inclu
des('Invalid resource') || e.message.includes
('not listed in the requested permissions')) 
{
            errorMessage = `<strong>Missing
 API Permissions</strong><br><br>
Your Azure 
AD app registration is missing required permi
ssions.<br><br><strong>Raw error:</strong> <c
ode style='font-size:11px;word-break:break-al
l'>${errorDesc}</code><br><br>
<strong>To fix
 this:</strong><br>
1. Go to <a href="https:/
/portal.azure.com/#view/Microsoft_AAD_Registe
redApps/ApplicationsListBlade" target="_blank
" rel="noopener">Azure Portal → App Registr
ations</a><br>
2. Find and click on your app<
br>
3. Go to <strong>API permissions</strong>
 → <strong>Add a permission</strong><br>
4.
 Add: <strong>Power Platform API</strong> →
 Delegated → <code>user_impersonation</code
><br>
5. Add: <strong>Dynamics CRM</strong> �
�� Delegated → <code>user_impersonation</co
de><br>
6. Click <strong>Grant admin consent<
/strong>${resetButton}<br><br>
<small style="
color:#888">Error: ${e.message.substring(0, 1
50)}...</small>`;
        } else if (e.messag
e.includes('AADSTS700016') || e.message.inclu
des('not found in the directory')) {
        
    errorMessage = `<strong>Application Not F
ound</strong><br><br>
The Client ID does not 
exist in the specified tenant.<br>
Please ver
ify your Tenant ID and Client ID are correct.
${resetButton}`;
        } else if (e.message
.includes('AADSTS50011') || e.message.include
s('reply URL') || e.message.includes('redirec
t')) {
            errorMessage = `<strong>In
valid Redirect URI</strong><br><br>
The redir
ect URI is not configured in your app registr
ation.<br><br>
Add this URI to your app's red
irect URIs:<br>
<code>${window.location.origi
n + window.location.pathname.substring(0, win
dow.location.pathname.lastIndexOf('/') + 1)}<
/code>${resetButton}`;
        } else {
     
       errorMessage = `<strong>Authentication
 Failed</strong><br><br>${e.message}${resetBu
tton}`;
        }
        
        showError(
errorMessage);
        msalInstance = null;
 
       ppToken = null;
        environmentId 
= null;
    }
}

// Handle authentication
asy
nc function handleAuthentication(event) {
   
 event.preventDefault();
    
    logInfo('==
= handleAuthentication START (user clicked Co
nnect) ===');
    
    let orgUrlValue = docu
ment.getElementById('orgUrl').value.trim();
 
   const tenantId = ''; // Tenant ID removed 
from UI - will use 'organizations' endpoint
 
   const clientId = document.getElementById('
clientId').value.trim();
    const rememberMe
 = document.getElementById('rememberMe').chec
ked;
    
    logInfo('Form values', { orgUrl
: orgUrlValue, clientId: clientId.substring(0
,8) + '...', rememberMe });
    
    const gu
idRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{
4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!guid
Regex.test(clientId)) {
        logError('Inv
alid Client ID GUID format');
        showErr
or('Invalid Client ID format. Client ID must 
be a valid GUID.');
        return;
    }
   
 
    // Normalize the org URL
    if (!orgUr
lValue.startsWith('https://')) {
        orgU
rlValue = 'https://' + orgUrlValue;
    }
   
 orgUrlValue = orgUrlValue.replace(/\/+$/, ''
); // remove trailing slashes
    
    if (!o
rgUrlValue.includes('.dynamics.com')) {
     
   logError('Invalid Organization URL');
    
    showError('Invalid Organization URL. It s
hould look like https://yourorg.crm.dynamics.
com');
        return;
    }
    
    // ALWA
YS save to sessionStorage for redirect flow (
survives redirect, works with tracking preven
tion)
    const credsObj = { orgUrl: orgUrlVa
lue, tenantId, clientId, rememberMe };
    se
ssionStorage.setItem('d365_app_updater_creds_
temp', JSON.stringify(credsObj));
    logInfo
('Credentials saved to sessionStorage for red
irect');
    
    // Also try localStorage if
 user wants to remember (may fail with tracki
ng prevention)
    if (rememberMe) {
        
try {
            localStorage.setItem('d365_
app_updater_creds', JSON.stringify({ orgUrl: 
orgUrlValue, tenantId, clientId }));
        
    logInfo('Credentials also saved to localS
torage');
        } catch (e) {
            l
ogWarn('Could not save to localStorage (track
ing prevention?)', e.message);
        }
    
}
    
    // Clear any previous auth step st
ate to start fresh
    sessionStorage.removeI
tem('d365_auth_step');
    logInfo('Cleared p
revious auth step');
    
    try {
        s
howLoading('Authenticating...', 'Connecting t
o Microsoft');
        
        const msalCon
fig = createMsalConfig(tenantId, clientId);
 
       logDebug('MSAL config', { redirectUri:
 msalConfig.auth.redirectUri, authority: msal
Config.auth.authority });
        
        ms
alInstance = new msal.PublicClientApplication
(msalConfig);
        await msalInstance.init
ialize();
        logInfo('MSAL initialized')
;
        
        // Redirect to Microsoft l
ogin — the page will reload and handleRedir
ectResponse() will continue
        showLoadi
ng('Authenticating...', 'Redirecting to Micro
soft sign-in...');
        
        // Save t
he pending auth state so we know to continue 
after redirect
        sessionStorage.setItem
('d365_auth_step', 'login_redirect');
       
 logInfo('Set auth step to login_redirect, ca
lling loginRedirect...');
        
        //
 SSO: pass loginHint if companion app provide
d one
        const _ssoLoginHint = sessionSt
orage.getItem('d365_login_hint');
        con
st _loginRequest = { scopes: ['openid', 'prof
ile'] };
        if (_ssoLoginHint) {
       
     _loginRequest.loginHint = _ssoLoginHint;

            logInfo('SSO loginHint applied',
 { loginHint: _ssoLoginHint.substring(0,8) + 
'...' });
        }
        await msalInstanc
e.loginRedirect(_loginRequest);
        
    
    // Execution stops here — browser navig
ates to Microsoft login
        return;
     
   
    } catch (error) {
        hideLoading
();
        console.error('Auth error:', erro
r);
        showError('Authentication failed:
 ' + error.message);
    }
}

// Resolve an O
rganization URL (e.g. https://orgname.crm.dyn
amics.com) to a Power Platform Environment ID

async function resolveOrgUrlToEnvironmentId(
orgUrl) {
    const bapToken = await getBAPTo
ken();
    
    // Normalize for comparison: 
lowercase, no trailing slash
    const normal
izedInput = orgUrl.toLowerCase().replace(/\/+
$/, '');
    
    // List all environments an
d find the one whose instanceUrl matches
    
const response = await fetch('https://api.bap
.microsoft.com/providers/Microsoft.BusinessAp
pPlatform/scopes/admin/environments?api-versi
on=2021-04-01', {
        headers: { 'Authori
zation': `Bearer ${bapToken}` }
    });
    

    if (!response.ok) {
        console.error
('Failed to list environments:', response.sta
tus);
        throw new Error('Failed to list
 environments. Make sure you have Power Platf
orm admin access.');
    }
    
    const dat
a = await response.json();
    const environm
ents = data.value || [];
    console.log('Fou
nd', environments.length, 'environments, sear
ching for URL:', orgUrl);
    
    // Cache a
ll environments for the switcher
    allEnvir
onments = environments.filter(env => env.prop
erties?.linkedEnvironmentMetadata?.instanceUr
l).map(env => ({
        id: env.name,
      
  name: env.properties?.displayName || env.na
me,
        instanceUrl: (env.properties?.lin
kedEnvironmentMetadata?.instanceUrl || '').re
place(/\/+$/, ''),
        type: env.properti
es?.environmentType || '',
    })).sort((a, b
) => a.name.localeCompare(b.name));
    
    
for (const env of environments) {
        con
st instanceUrl = env.properties?.linkedEnviro
nmentMetadata?.instanceUrl;
        const env
Name = env.properties?.displayName || env.nam
e;
        
        if (instanceUrl) {
      
      const normalizedInstance = instanceUrl.
toLowerCase().replace(/\/+$/, '');
          
  console.log(`  Environment: ${envName} (${e
nv.name}), instanceUrl: ${instanceUrl}`);
   
         
            if (normalizedInstance 
=== normalizedInput) {
                consol
e.log('  ✓ Match found! Environment ID:', e
nv.name);
                return env.name;
  
          }
        }
    }
    
    console.
error('No environment found matching URL:', o
rgUrl);
    return null;
}

// Get environmen
t name from BAP API (for display purposes)
as
ync function getEnvironmentName() {
    // Us
e shared BAP token helper
    const bapToken 
= await getBAPToken();
    
    // Get enviro
nment details by ID
    const response = awai
t fetch(`https://api.bap.microsoft.com/provid
ers/Microsoft.BusinessAppPlatform/scopes/admi
n/environments/${environmentId}?api-version=2
021-04-01`, {
        headers: { 'Authorizati
on': `Bearer ${bapToken}` }
    });
    
    
if (response.ok) {
        const env = await 
response.json();
        const displayName = 
env.properties?.displayName || environmentId;

        document.getElementById('environment
Name').textContent = displayName;
        con
sole.log('Environment:', displayName);
    } 
else {
        // If we can't get the name, j
ust use the ID
        document.getElementByI
d('environmentName').textContent = environmen
tId;
        console.log('Could not get envir
onment name, using ID');
    }
    
    // Re
nder environment switcher
    renderEnvSwitch
er();
}

// ── Environment Switcher ─�
�──────────────�
�──────────────�
�──────
function renderEnvSwitche
r() {
    const list = document.getElementByI
d('envList');
    if (!list || allEnvironment
s.length === 0) return;
    
    let html = '
';
    for (const env of allEnvironments) {
 
       const isActive = env.id === environmen
tId;
        const shortUrl = env.instanceUrl
.replace(/^https?:\/\//, '');
        html +=
 '<div class="env-item' + (isActive ? ' activ
e' : '') + '" onclick="switchEnvironment(\'' 
+ env.id + '\')" title="' + escapeHtml(env.in
stanceUrl) + '">';
        html += '<div clas
s="env-item-icon"><i class="fas fa-' + (isAct
ive ? 'check' : 'globe') + '"></i></div>';
  
      html += '<div class="env-item-details">
';
        html += '<div class="env-item-name
">' + escapeHtml(env.name) + '</div>';
      
  html += '<div class="env-item-url">' + esca
peHtml(shortUrl) + '</div>';
        html += 
'</div>';
        html += '</div>';
    }
   
 list.innerHTML = html;
}

function toggleEnv
Dropdown() {
    const dropdown = document.ge
tElementById('envDropdown');
    const btn = 
document.getElementById('envSwitcherBtn');
  
  const isOpen = dropdown.classList.contains(
'show');
    if (isOpen) {
        closeEnvDr
opdown();
    } else {
        dropdown.class
List.add('show');
        btn.classList.add('
open');
        document.getElementById('envS
earchInput').value = '';
        document.get
ElementById('envSearchInput').focus();
      
  filterEnvList();
    }
}

function closeEnv
Dropdown() {
    const dropdown = document.ge
tElementById('envDropdown');
    const btn = 
document.getElementById('envSwitcherBtn');
  
  if (dropdown) dropdown.classList.remove('sh
ow');
    if (btn) btn.classList.remove('open
');
}

function filterEnvList() {
    const s
earch = (document.getElementById('envSearchIn
put').value || '').toLowerCase();
    const l
ist = document.getElementById('envList');
   
 const filtered = allEnvironments.filter(env 
=> {
        if (!search) return true;
      
  return env.name.toLowerCase().includes(sear
ch) || env.instanceUrl.toLowerCase().includes
(search);
    });
    
    if (filtered.lengt
h === 0) {
        list.innerHTML = '<div cla
ss="env-empty"><i class="fas fa-search me-1">
</i> No environments found</div>';
        re
turn;
    }
    
    let html = '';
    for (
const env of filtered) {
        const isActi
ve = env.id === environmentId;
        const 
shortUrl = env.instanceUrl.replace(/^https?:\
/\//, '');
        html += '<div class="env-i
tem' + (isActive ? ' active' : '') + '" oncli
ck="switchEnvironment(\'' + env.id + '\')" ti
tle="' + escapeHtml(env.instanceUrl) + '">';

        html += '<div class="env-item-icon"><
i class="fas fa-' + (isActive ? 'check' : 'gl
obe') + '"></i></div>';
        html += '<div
 class="env-item-details">';
        html += 
'<div class="env-item-name">' + escapeHtml(en
v.name) + '</div>';
        html += '<div cla
ss="env-item-url">' + escapeHtml(shortUrl) + 
'</div>';
        html += '</div>';
        h
tml += '</div>';
    }
    list.innerHTML = h
tml;
}

async function switchEnvironment(envI
d) {
    if (envId === environmentId) {
     
   closeEnvDropdown();
        return;
    }

    
    closeEnvDropdown();
    
    const e
nv = allEnvironments.find(e => e.id === envId
);
    if (!env) return;
    
    environment
Id = envId;
    currentOrgUrl = env.instanceU
rl;
    selectedApps.clear();
    
    // Upd
ate saved credentials with new org URL
    co
nst savedCreds = localStorage.getItem('d365_a
pp_updater_creds');
    if (savedCreds) {
   
     try {
            const creds = JSON.par
se(savedCreds);
            creds.orgUrl = en
v.instanceUrl;
            localStorage.setIt
em('d365_app_updater_creds', JSON.stringify(c
reds));
        } catch (e) {}
    }
    
   
 document.getElementById('environmentName').t
extContent = env.name;
    renderEnvSwitcher(
);
    
    // Reset schedule state for the n
ew environment
    scheduleLoaded = false;
  
  document.getElementById('scheduleEnabled').
checked = false;
    document.getElementById(
'scheduleDetails').style.display = 'none';
  
  document.getElementById('scheduleClientSecr
et').value = '';
    document.getElementById(
'scheduleClientSecret').placeholder = 'Enter 
client secret';
    document.getElementById('
scheduleStatus').innerHTML = '<i class="fas f
a-info-circle"></i> Schedule not configured';

    document.getElementById('scheduleStatus'
).className = 'schedule-status';
    // Load 
schedule for the new environment
    loadSche
dule();
    
    console.log('Switching to en
vironment:', env.name, '(' + envId + ')');
  
  await loadApplications();
}

// Compare two
 version strings (e.g., "1.2.3.4" vs "1.2.3.5
")
function compareVersions(v1, v2) {
    if 
(!v1 || !v2) return 0;
    const parts1 = v1.
split('.').map(Number);
    const parts2 = v2
.split('.').map(Number);
    for (let i = 0; 
i < Math.max(parts1.length, parts2.length); i
++) {
        const p1 = parts1[i] || 0;
    
    const p2 = parts2[i] || 0;
        if (p1
 > p2) return 1;
        if (p1 < p2) return 
-1;
    }
    return 0;
}

// Helper: fetch a
ll pages from a paginated Power Platform API 
endpoint
async function fetchAllPages(url, to
ken) {
    let allItems = [];
    let nextUrl
 = url;
    let pageCount = 0;
    
    while
 (nextUrl) {
        pageCount++;
        con
sole.log(`Fetching page ${pageCount}: ${nextU
rl.substring(0, 120)}...`);
        
        
const controller = new AbortController();
   
     const timeoutId = setTimeout(() => contr
oller.abort(), 60000);
        
        const
 response = await fetch(nextUrl, {
          
  headers: { 'Authorization': `Bearer ${token
}` },
            signal: controller.signal
 
       });
        
        clearTimeout(time
outId);
        
        if (!response.ok) {

            const errorText = await response.
text();
            console.error('API Error 
on page', pageCount, ':', response.status, er
rorText);
            throw new Error('Failed
 to fetch apps (page ' + pageCount + '): ' + 
response.status);
        }
        
        
const data = await response.json();
        c
onst items = data.value || [];
        allIte
ms = allItems.concat(items);
        nextUrl 
= data['@odata.nextLink'] || null;
        
 
       console.log(`Page ${pageCount}: got ${
items.length} items (total so far: ${allItems
.length})`);
    }
    
    return allItems;

}

// Helper: refresh the Power Platform API 
token silently
async function refreshPPToken(
) {
    try {
        const accounts = msalIn
stance.getAllAccounts();
        if (!account
s || accounts.length === 0) return false;
   
     const ppRequest = { scopes: ['https://ap
i.powerplatform.com/.default'], account: acco
unts[0] };
        const result = await msalI
nstance.acquireTokenSilent(ppRequest);
      
  ppToken = result.accessToken;
        conso
le.log('PP token refreshed successfully');
  
      return true;
    } catch (e) {
        
console.warn('Failed to refresh PP token:', e
.message);
        return false;
    }
}

// 
Helper: get BAP token for admin API calls
asy
nc function getBAPToken() {
    const account
s = msalInstance.getAllAccounts();
    const 
bapRequest = { scopes: ['https://api.bap.micr
osoft.com/.default'], account: accounts[0] };

    try {
        const result = await msalI
nstance.acquireTokenSilent(bapRequest);
     
   return result.accessToken;
    } catch (e)
 {
        // Silent failed — redirect for 
consent, but track step to avoid loops
      
  const currentStep = sessionStorage.getItem(
'd365_auth_step');
        if (currentStep ==
= 'acquiring_bap_runtime') {
            thro
w new Error('Failed to acquire BAP token. Ple
ase check your app registration permissions.'
);
        }
        sessionStorage.setItem('
d365_auth_step', 'acquiring_bap_runtime');
  
      await msalInstance.acquireTokenRedirect
(bapRequest);
        // Page will reload, th
row to stop current execution
        throw n
ew Error('Redirecting for BAP API consent...'
);
    }
}

// Load applications from Power P
latform API
async function loadApplications()
 {
    showLoading('Loading applications...',
 'Fetching from Power Platform');
    
    co
nst appsList = document.getElementById('appsL
ist');
    appsList.innerHTML = '<div class="
text-center py-5"><div class="spinner-border 
text-primary"></div></div>';
    
    try {
 
       const baseUrl = `https://api.powerplat
form.com/appmanagement/environments/${environ
mentId}/applicationPackages`;
        const a
piVersion = 'api-version=2022-03-01-preview';

        
        // ── Step 1: Fetch INS
TALLED apps explicitly ───────�
��──────────
        show
Loading('Loading applications...', 'Fetching 
installed apps...');
        const installedA
ppsRaw = await fetchAllPages(
            `${
baseUrl}?appInstallState=Installed&${apiVersi
on}`, ppToken
        );
        console.log(
'Installed apps fetched:', installedAppsRaw.l
ength);
        
        // ── Step 2: Fe
tch ALL catalog packages (includes newer vers
ions) ──
        showLoading('Loading app
lications...', 'Fetching available catalog ve
rsions...');
        const allAppsRaw = await
 fetchAllPages(
            `${baseUrl}?${api
Version}`, ppToken
        );
        console
.log('All catalog packages fetched:', allApps
Raw.length);
        
        // ── Step 
2b: Fetch NotInstalled packages specifically 
(update packages) ──
        showLoading(
'Loading applications...', 'Fetching update p
ackages...');
        let notInstalledRaw = [
];
        try {
            notInstalledRaw 
= await fetchAllPages(
                `${bas
eUrl}?appInstallState=NotInstalled&${apiVersi
on}`, ppToken
            );
            cons
ole.log('NotInstalled packages fetched:', not
InstalledRaw.length);
        } catch (e) {
 
           console.warn('NotInstalled fetch f
ailed (non-critical):', e.message);
        }

        
        // Merge all catalog source
s
        const allCatalogEntries = [...allAp
psRaw, ...notInstalledRaw];
        
        
// Debug: log unique states and sample data
 
       const states = [...new Set(allCatalogE
ntries.map(a => a.state))];
        console.l
og('All states found in catalog:', states);
 
       if (installedAppsRaw.length > 0) {
   
         console.log('Sample installed app fi
elds:', Object.keys(installedAppsRaw[0]));
  
          console.log('Sample installed app:'
, JSON.stringify(installedAppsRaw[0], null, 2
));
        }
        
        // ── Step
 3: Build version maps from ALL catalog entri
es ──────
        // Map by appli
cationId → keep highest version
        con
st catalogMapById = new Map();
        for (c
onst app of allCatalogEntries) {
            
if (!app.applicationId) continue;
           
 const existing = catalogMapById.get(app.appl
icationId);
            if (!existing || comp
areVersions(app.version, existing.version) > 
0) {
                catalogMapById.set(app.a
pplicationId, app);
            }
        }
 
       
        // Map by uniqueName base →
 keep highest version (fallback matching)
   
     const catalogByName = new Map();
       
 for (const app of allCatalogEntries) {
     
       if (!app.uniqueName) continue;
       
     const baseName = app.uniqueName.replace(
/_upgrade$/i, '').replace(/_\d+$/, '');
     
       const existing = catalogByName.get(bas
eName);
            if (!existing || compareV
ersions(app.version, existing.version) > 0) {

                catalogByName.set(baseName, 
app);
            }
        }
        
      
  // Map by display name → keep highest ver
sion
        const catalogByDisplayName = new
 Map();
        for (const app of allCatalogE
ntries) {
            const name = (app.local
izedName || app.applicationName || '').toLowe
rCase();
            if (!name) continue;
   
         const existing = catalogByDisplayNam
e.get(name);
            if (!existing || com
pareVersions(app.version, existing.version) >
 0) {
                catalogByDisplayName.se
t(name, app);
            }
        }
       
 
        console.log('Catalog map by ID entr
ies:', catalogMapById.size);
        console.
log('Catalog map by name entries:', catalogBy
Name.size);
        
        // ── Step 4
: Detect updates for each installed app ─�
�──────────
        let u
pdatesFound = 0;
        apps = installedApps
Raw.map(app => {
            let hasUpdate = 
false;
            let latestVersion = null;

            let catalogUniqueName = null;
   
         let spaOnly = false;
            
  
          // Skip update detection for apps t
hat require Custom Install Experience (SPA)
 
           // These cannot be updated via the
 API — they must be updated through the Adm
in Center
            if (app.singlePageAppli
cationUrl) {
                spaOnly = true;

            }
            
            // Che
ck 0: State-based detection — API may direc
tly flag updates
            const stateLower
 = (app.state || '').toLowerCase();
         
   if (!spaOnly && (stateLower.includes('upda
te') || stateLower === 'updateavailable' || s
tateLower === 'installedwithupdateavailable')
) {
                hasUpdate = true;
       
         console.log(`  [by state="${app.stat
e}"] ${app.localizedName || app.uniqueName}`)
;
            }
            
            // C
heck 1: Direct API fields that might indicate
 update availability
            if (!spaOnly
 && (app.updateAvailable || app.catalogVersio
n || app.availableVersion || app.latestVersio
n || app.newVersion || app.updateVersion)) {

                const directVersion = app.cat
alogVersion || app.availableVersion || app.la
testVersion || app.newVersion || app.updateVe
rsion;
                if (directVersion && c
ompareVersions(directVersion, app.version) > 
0) {
                    hasUpdate = true;
  
                  latestVersion = directVersi
on;
                    console.log(`  [direc
t field] ${app.localizedName || app.uniqueNam
e}: ${app.version} → ${latestVersion}`);
  
              }
                // updateAvai
lable might be a boolean
                if (
app.updateAvailable === true && !latestVersio
n) {
                    hasUpdate = true;
  
                  console.log(`  [updateAvail
able=true] ${app.localizedName || app.uniqueN
ame}`);
                }
            }
     
       
            // ALWAYS look up catalog
 uniqueName — even if hasUpdate was already
 set by
            // state/flags above. The
 install API needs the CATALOG package's uniq
ueName
            // (often has _upgrade suf
fix), NOT the installed app's uniqueName. Usi
ng the
            // installed uniqueName re
sults in a no-op (200 OK but no actual update
).
            
            // Check 2: Catal
og entry by applicationId
            if (!sp
aOnly && app.applicationId) {
               
 const catalogEntry = catalogMapById.get(app.
applicationId);
                if (catalogEn
try && compareVersions(catalogEntry.version, 
app.version) > 0) {
                    if (!
hasUpdate) hasUpdate = true;
                
    latestVersion = catalogEntry.version;
   
                 catalogUniqueName = catalogE
ntry.uniqueName;
                    console.
log(`  [by appId] ${app.localizedName || app.
uniqueName}: ${app.version} → ${latestVersi
on} (pkg: ${catalogUniqueName})`);
          
      }
            }
            
          
  // Check 3: Catalog entry by uniqueName bas
e
            if (!spaOnly && !catalogUniqueN
ame && app.uniqueName) {
                cons
t baseName = app.uniqueName.replace(/_upgrade
$/i, '').replace(/_\d+$/, '');
              
  const byName = catalogByName.get(baseName);

                if (byName && compareVersion
s(byName.version, app.version) > 0) {
       
             if (!hasUpdate) hasUpdate = true
;
                    latestVersion = byName.
version;
                    catalogUniqueNam
e = byName.uniqueName;
                    co
nsole.log(`  [by name] ${app.localizedName ||
 app.uniqueName}: ${app.version} → ${latest
Version} (pkg: ${catalogUniqueName})`);
     
           }
            }
            
     
       // Check 4: Catalog entry by localized
Name / applicationName
            if (!spaOn
ly && !catalogUniqueName) {
                c
onst appName = (app.localizedName || app.appl
icationName || '').toLowerCase();
           
     if (appName) {
                    for (
const [, catApp] of catalogMapById) {
       
                 const catName = (catApp.loca
lizedName || catApp.applicationName || '').to
LowerCase();
                        if (catN
ame === appName && compareVersions(catApp.ver
sion, app.version) > 0) {
                   
         if (!hasUpdate) hasUpdate = true;
  
                          latestVersion = cat
App.version;
                            cata
logUniqueName = catApp.uniqueName;
          
                  console.log(`  [by displayN
ame] ${app.localizedName || app.uniqueName}: 
${app.version} → ${latestVersion} (pkg: ${c
atalogUniqueName})`);
                       
     break;
                        }
       
             }
                }
            
}
            
            if (hasUpdate) upd
atesFound++;
            if (spaOnly) {
     
           console.log(`  [skipped SPA] ${app
.localizedName || app.uniqueName} — require
s Admin Center`);
            }
            

            return {
                id: app.
id,
                uniqueName: app.uniqueNam
e,
                catalogUniqueName: catalog
UniqueName || app.uniqueName,
               
 name: app.localizedName || app.applicationNa
me || app.uniqueName || 'Unknown',
          
      version: app.version || 'Unknown',
    
            latestVersion: latestVersion,
   
             state: app.state || 'Installed',

                hasUpdate: hasUpdate,
      
          publisher: app.publisherName || 'Mi
crosoft',
                description: app.ap
plicationDescription || '',
                l
earnMoreUrl: app.learnMoreUrl || null,
      
          instancePackageId: app.instancePack
ageId,
                applicationId: app.app
licationId,
                spaOnly: spaOnly

            };
        });
        
        c
onsole.log('Updates found from PP API:', upda
tesFound);
        
        // ── Step 5:
 ALWAYS check BAP Admin API for additional up
dates ──
        // The BAP API can detec
t updates that the PP API misses
        cons
ole.log('Checking BAP Admin API for additiona
l updates...');
        showLoading('Loading 
applications...', 'Cross-checking updates via
 Admin API...');
        
        try {
     
       const bapToken = await getBAPToken();

            const bapUrl = `https://api.bap.m
icrosoft.com/providers/Microsoft.BusinessAppP
latform/scopes/admin/environments/${environme
ntId}/applicationPackages?api-version=2021-04
-01`;
            const bapApps = await fetch
AllPages(bapUrl, bapToken);
            
    
        console.log('BAP API returned:', bapA
pps.length, 'packages');
            if (bapA
pps.length > 0) {
                console.log
('BAP sample app fields:', Object.keys(bapApp
s[0]));
                // Log first 3 sample
s for debugging
                bapApps.slice
(0, 3).forEach((a, i) => console.log(`BAP sam
ple ${i}:`, JSON.stringify(a, null, 2)));
   
         }
            
            // Build 
BAP catalog map by applicationId (keep highes
t version)
            const bapCatalogMap = 
new Map();
            for (const bapApp of b
apApps) {
                if (!bapApp.applica
tionId) continue;
                const exist
ing = bapCatalogMap.get(bapApp.applicationId)
;
                if (!existing || compareVer
sions(bapApp.version, existing.version) > 0) 
{
                    bapCatalogMap.set(bapAp
p.applicationId, bapApp);
                }
 
           }
            
            // Also
 build by uniqueName base
            const b
apByName = new Map();
            for (const 
bapApp of bapApps) {
                if (!bap
App.uniqueName) continue;
                con
st baseName = bapApp.uniqueName.replace(/_upg
rade$/i, '').replace(/_\d+$/, '');
          
      const existing = bapByName.get(baseName
);
                if (!existing || compareVe
rsions(bapApp.version, existing.version) > 0)
 {
                    bapByName.set(baseName
, bapApp);
                }
            }
  
          
            // Also build by displ
ay name
            const bapByDisplayName = 
new Map();
            for (const bapApp of b
apApps) {
                const name = (bapAp
p.localizedName || bapApp.applicationName || 
'').toLowerCase();
                if (!name)
 continue;
                const existing = b
apByDisplayName.get(name);
                if
 (!existing || compareVersions(bapApp.version
, existing.version) > 0) {
                  
  bapByDisplayName.set(name, bapApp);
       
         }
            }
            
       
     // Check installed apps that DON'T alrea
dy have an update detected
            for (c
onst app of apps) {
                if (app.h
asUpdate) continue;
                
        
        let found = false;
                
 
               // Check direct fields from BA
P response for this app
                const
 bapInstalled = bapApps.find(b => 
          
          (b.applicationId === app.applicatio
nId) && 
                    (b.state === 'In
stalled' || b.instancePackageId)
            
    );
                if (bapInstalled) {
  
                  // State-based detection
  
                  const bapState = (bapInstal
led.state || '').toLowerCase();
             
       if (bapState.includes('update') || bap
State === 'updateavailable') {
              
          app.hasUpdate = true;
             
           found = true;
                    
    console.log(`  [BAP state="${bapInstalled
.state}"] ${app.name}`);
                    
}
                    // Check if BAP provide
s update info directly
                    co
nst directVer = bapInstalled.catalogVersion |
| bapInstalled.availableVersion || bapInstall
ed.latestVersion || bapInstalled.newVersion |
| bapInstalled.updateVersion;
               
     if (directVer && compareVersions(directV
er, app.version) > 0) {
                     
   app.hasUpdate = true;
                    
    app.latestVersion = directVer;
          
              found = true;
                 
       console.log(`  [BAP direct] ${app.name
}: ${app.version} → ${directVer}`);
       
             }
                    if (bapIns
talled.updateAvailable === true) {
          
              app.hasUpdate = true;
         
               found = true;
                
        console.log(`  [BAP updateAvailable=t
rue] ${app.name}`);
                    }
   
             }
                
             
   // Compare by applicationId
              
  if (!found && app.applicationId) {
        
            const bapEntry = bapCatalogMap.ge
t(app.applicationId);
                    if 
(bapEntry && compareVersions(bapEntry.version
, app.version) > 0) {
                       
 app.hasUpdate = true;
                      
  app.latestVersion = bapEntry.version;
     
                   app.catalogUniqueName = ba
pEntry.uniqueName || app.uniqueName;
        
                found = true;
               
         console.log(`  [BAP by appId] ${app.
name}: ${app.version} → ${bapEntry.version}
`);
                    }
                }
 
               
                // Compare by
 uniqueName
                if (!found && app
.uniqueName) {
                    const base
Name = app.uniqueName.replace(/_upgrade$/i, '
').replace(/_\d+$/, '');
                    
const bapByNameEntry = bapByName.get(baseName
);
                    if (bapByNameEntry && 
compareVersions(bapByNameEntry.version, app.v
ersion) > 0) {
                        app.ha
sUpdate = true;
                        app.l
atestVersion = bapByNameEntry.version;
      
                  app.catalogUniqueName = bap
ByNameEntry.uniqueName || app.uniqueName;
   
                     found = true;
          
              console.log(`  [BAP by name] ${
app.name}: ${app.version} → ${bapByNameEntr
y.version}`);
                    }
         
       }
                
                // 
Compare by display name
                if (!
found) {
                    const appDisplay
Name = (app.name || '').toLowerCase();
      
              if (appDisplayName) {
         
               const bapByDN = bapByDisplayNa
me.get(appDisplayName);
                     
   if (bapByDN && compareVersions(bapByDN.ver
sion, app.version) > 0) {
                   
         app.hasUpdate = true;
              
              app.latestVersion = bapByDN.ver
sion;
                            app.catalog
UniqueName = bapByDN.uniqueName || app.unique
Name;
                            found = tru
e;
                            console.log(` 
 [BAP by displayName] ${app.name}: ${app.vers
ion} → ${bapByDN.version}`);
              
          }
                    }
           
     }
                
                if (f
ound) updatesFound++;
            }
         
   
            // ── Step 5b: Check for 
installed apps that BAP knows but PP API miss
ed entirely ──
            const knownApp
Ids = new Set(apps.map(a => a.applicationId).
filter(Boolean));
            const knownName
s = new Set(apps.map(a => (a.name || '').toLo
werCase()).filter(Boolean));
            
   
         for (const bapApp of bapApps) {
    
            const bapState = (bapApp.state ||
 '').toLowerCase();
                const isI
nstalled = bapState === 'installed' || bapSta
te.includes('update') || bapApp.instancePacka
geId;
                if (!isInstalled) conti
nue;
                
                // Skip
 if we already know about this app
          
      if (bapApp.applicationId && knownAppIds
.has(bapApp.applicationId)) continue;
       
         const bapName = (bapApp.localizedNam
e || bapApp.applicationName || '').toLowerCas
e();
                if (bapName && knownName
s.has(bapName)) continue;
                
  
              // This is an installed app the
 PP API missed
                let hasUpdate 
= false;
                let latestVersion = 
null;
                
                if (ba
pState.includes('update') || bapApp.updateAva
ilable === true) {
                    hasUpd
ate = true;
                }
               
 const directVer = bapApp.catalogVersion || b
apApp.availableVersion || bapApp.latestVersio
n || bapApp.newVersion;
                if (d
irectVer && compareVersions(directVer, bapApp
.version) > 0) {
                    hasUpdat
e = true;
                    latestVersion =
 directVer;
                }
               
 // Check if BAP catalog has a higher version

                if (!hasUpdate && bapApp.app
licationId) {
                    const bapCa
tEntry = bapCatalogMap.get(bapApp.application
Id);
                    if (bapCatEntry && c
ompareVersions(bapCatEntry.version, bapApp.ve
rsion) > 0) {
                        hasUpda
te = true;
                        latestVers
ion = bapCatEntry.version;
                  
  }
                }
                
      
          if (hasUpdate) {
                  
  updatesFound++;
                    console
.log(`  [BAP new app] ${bapApp.localizedName 
|| bapApp.uniqueName}: ${bapApp.version} → 
${latestVersion || 'update flagged'}`);
     
           }
                
               
 apps.push({
                    id: bapApp.i
d,
                    uniqueName: bapApp.uni
queName,
                    catalogUniqueNam
e: bapApp.uniqueName,
                    nam
e: bapApp.localizedName || bapApp.application
Name || bapApp.uniqueName || 'Unknown',
     
               version: bapApp.version || 'Un
known',
                    latestVersion: la
testVersion,
                    state: bapAp
p.state || 'Installed',
                    h
asUpdate: hasUpdate,
                    publ
isher: bapApp.publisherName || 'Microsoft',
 
                   description: bapApp.applic
ationDescription || '',
                    l
earnMoreUrl: bapApp.learnMoreUrl || null,
   
                 instancePackageId: bapApp.in
stancePackageId,
                    applicat
ionId: bapApp.applicationId
                }
);
            }
            
            con
sole.log('Total updates found after BAP cross
-check:', updatesFound);
        } catch (bap
Error) {
            console.warn('BAP API cr
oss-check failed (non-critical):', bapError.m
essage);
        }
        
        // ──
 Step 6: Sort — updates first, then alphabe
tically ────────
        apps
.sort((a, b) => {
            if (a.hasUpdate
 && !b.hasUpdate) return -1;
            if (
!a.hasUpdate && b.hasUpdate) return 1;
      
      return a.name.localeCompare(b.name);
  
      });
        
        // Store not-insta
lled apps for browsing
        const knownIns
talledAppIds = new Set(apps.map(a => a.applic
ationId).filter(Boolean));
        const notI
nstalledApps = [];
        for (const [appId,
 app] of catalogMapById) {
            if (!k
nownInstalledAppIds.has(appId) && app.state !
== 'Installed' && !app.instancePackageId) {
 
               notInstalledApps.push({
      
              id: app.id,
                   
 uniqueName: app.uniqueName,
                
    catalogUniqueName: app.uniqueName,
      
              name: app.localizedName || app.
applicationName || app.uniqueName || 'Unknown
',
                    version: app.version |
| 'Unknown',
                    latestVersio
n: null,
                    state: 'Availabl
e',
                    hasUpdate: false,
   
                 publisher: app.publisherName
 || 'Microsoft',
                    descript
ion: app.applicationDescription || '',
      
              learnMoreUrl: app.learnMoreUrl 
|| null,
                    instancePackageI
d: null,
                    applicationId: a
pp.applicationId
                });
        
    }
        }
        notInstalledApps.sort
((a, b) => a.name.localeCompare(b.name));
   
     window.availableApps = notInstalledApps;

        
        console.log('Final result:'
, apps.length, 'installed apps,', updatesFoun
d, 'with updates');
        displayApplicatio
ns();
        hideLoading();
        
    } c
atch (error) {
        hideLoading();
       
 console.error('Error loading applications:',
 error);
        
        let errorMsg = erro
r.message;
        if (error.name === 'AbortE
rror') {
            errorMsg = 'Request time
d out. The Power Platform API took too long t
o respond. Please try again.';
        }
    
    
        const appsList = document.getEle
mentById('appsList');
        appsList.innerH
TML = '<div class="alert alert-danger"><i cla
ss="fas fa-exclamation-triangle me-2"></i>Fai
led to load applications: ' + errorMsg + '</d
iv>';
    }
}

// Display applications
functi
on displayApplications() {
    const appsList
 = document.getElementById('appsList');
    

    if (apps.length === 0) {
        appsList
.innerHTML = '<div class="text-center py-5"><
p class="text-muted">No applications found.</
p></div>';
        return;
    }
    
    con
st appsWithUpdates = apps.filter(a => a.hasUp
date && !a.updateState);
    const appsUpdati
ng = apps.filter(a => a.updateState === 'subm
itted' || a.updateState === 'updating');
    
const appsFailed = apps.filter(a => a.updateS
tate === 'failed');
    const installedApps =
 apps.filter(a => a.instancePackageId);
    c
onst updateCount = appsWithUpdates.length;
  
  const updatingCount = appsUpdating.length;

    const failedCount = appsFailed.length;
  
  
    // Update summary text
    let summary
Parts = [installedApps.length + ' apps instal
led'];
    if (updateCount > 0) {
        sum
maryParts.push('<span style="color: #28a745; 
font-weight: 600;">' + updateCount + ' update
' + (updateCount !== 1 ? 's' : '') + ' availa
ble</span>');
    }
    if (updatingCount > 0
) {
        summaryParts.push('<span style="c
olor: #0d6efd; font-weight: 600;"><span class
="spinner-updating"></span>' + updatingCount 
+ ' updating</span>');
    }
    if (failedCo
unt > 0) {
        summaryParts.push('<span s
tyle="color: #dc3545; font-weight: 600;">' + 
failedCount + ' failed</span>');
    }
    if
 (updateCount === 0 && updatingCount === 0 &&
 failedCount === 0) {
        summaryParts.pu
sh('all up to date');
    }
    document.getE
lementById('appCountText').innerHTML = summar
yParts.join(' &nbsp;|&nbsp; ');
    
    docu
ment.getElementById('updateAllBtn').disabled 
= updateCount === 0;
    
    // Update selec
ted button visibility
    updateSelectedButto
n();
    
    // Show installed apps (include
 updating/failed states)
    const installedO
rUpdatable = apps.filter(a => a.hasUpdate || 
a.instancePackageId || a.updateState);
    co
nst appsToShow = installedOrUpdatable.length 
> 0 ? installedOrUpdatable : apps.slice(0, 50
);
    
    // Sort: failed first, then updat
ing, then updates available, then installed
 
   appsToShow.sort((a, b) => {
        const 
order = s => s === 'failed' ? 0 : (s === 'sub
mitted' || s === 'updating') ? 1 : 2;
       
 const oa = order(a.updateState), ob = order(
b.updateState);
        if (oa !== ob) return
 oa - ob;
        if (a.hasUpdate && !b.hasUp
date) return -1;
        if (!a.hasUpdate && 
b.hasUpdate) return 1;
        return a.name.
localeCompare(b.name);
    });
    
    let h
tml = '';
    
    // Status banners
    if (
failedCount > 0) {
        html += '<div clas
s="alert alert-danger mb-3" style="border-lef
t: 4px solid #dc3545;">';
        html += '<d
iv class="d-flex align-items-center">';
     
   html += '<i class="fas fa-exclamation-tria
ngle fa-2x me-3 text-danger"></i>';
        h
tml += '<div>';
        html += '<strong>' + 
failedCount + ' update' + (failedCount !== 1 
? 's' : '') + ' failed</strong><br>';
       
 html += '<small>Scroll down to see details. 
You can retry individual apps or check the Po
wer Platform Admin Center.</small>';
        
html += '</div>';
        html += '</div>';
 
       html += '</div>';
    }
    if (updati
ngCount > 0) {
        html += '<div class="a
lert alert-info mb-3" style="border-left: 4px
 solid #0d6efd;">';
        html += '<div cla
ss="d-flex align-items-center">';
        htm
l += '<i class="fas fa-sync-alt fa-spin fa-2x
 me-3 text-primary"></i>';
        html += '<
div>';
        html += '<strong>' + updatingC
ount + ' update' + (updatingCount !== 1 ? 's'
 : '') + ' in progress</strong><br>';
       
 html += '<small>Updates are running in the b
ackground. Click <strong>"Refresh"</strong> t
o check current status.</small>';
        htm
l += '</div>';
        html += '</div>';
    
    html += '</div>';
    }
    if (updateCou
nt > 0) {
        html += '<div class="alert 
alert-warning mb-3" style="border-left: 4px s
olid #ffc107;">';
        html += '<div class
="d-flex align-items-center">';
        html 
+= '<i class="fas fa-arrow-circle-up fa-2x me
-3 text-warning"></i>';
        html += '<div
>';
        html += '<strong>' + updateCount 
+ ' update' + (updateCount !== 1 ? 's' : '') 
+ ' available</strong><br>';
        html += 
'<small>Click <strong>"Update All Apps"</stro
ng> to apply all updates, or update apps indi
vidually below.</small>';
        html += '</
div>';
        html += '</div>';
        html
 += '</div>';
    }
    if (updateCount === 0
 && updatingCount === 0 && failedCount === 0)
 {
        html += '<div class="alert alert-s
uccess mb-3">';
        html += '<i class="fa
s fa-check-circle me-2"></i>';
        html +
= 'All installed applications are up to date.
';
        html += '</div>';
    }
    
    f
or (const app of appsToShow) {
        let st
ateClass, stateIcon, stateText, cardClass, ca
rdStyle;
        
        if (app.updateState
 === 'submitted' || app.updateState === 'upda
ting') {
            stateClass = 'primary';

            stateIcon = '';
            state
Text = 'Updating...';
            cardClass =
 'app-card state-updating';
            cardS
tyle = '';
        } else if (app.updateState
 === 'failed') {
            stateClass = 'da
nger';
            stateIcon = 'exclamation-t
riangle';
            stateText = 'Failed';
 
           cardClass = 'app-card state-failed
';
            cardStyle = '';
        } else
 if (app.hasUpdate) {
            stateClass 
= 'warning';
            stateIcon = 'arrow-c
ircle-up';
            stateText = 'Update Av
ailable';
            cardClass = 'app-card app-card-update';

            cardStyle = 'border-left: 4px solid #e67e22;';
        }
 else {
            stateClass = 'secondary';

            stateIcon = 'check-circle';
    
        stateText = app.instancePackageId ? '
Installed' : 'Available';
            cardCla
ss = 'app-card';
            cardStyle = '';

        }
        
        html += '<div clas
s="' + cardClass + '" style="' + cardStyle + 
'">';
        html += '<div class="row align-
items-center">';
        html += '<div class=
"col-md-6">';
        // Show checkbox for ap
ps with available updates or failed state
   
     const showCheckbox = app.hasUpdate || ap
p.updateState === 'failed';
        const isC
hecked = selectedApps.has(app.uniqueName);
  
      if (showCheckbox) {
            html +=
 '<div class="d-flex align-items-start">';
  
          html += '<input type="checkbox" cla
ss="app-select-cb" ' + (isChecked ? 'checked'
 : '') + ' onchange="toggleAppSelection(\'' +
 escapeHtml(app.uniqueName) + '\', this.check
ed)" title="Select for bulk update">';
      
      html += '<div>';
        }
        html
 += '<div class="app-name"><i class="fas fa-c
ube me-2"></i>' + escapeHtml(app.name) + '</d
iv>';
        html += '<div class="app-versio
n mt-2">';
        html += '<i class="fas fa-
tag"></i> Version: <strong>' + escapeHtml(app
.version) + '</strong>';
        if ((app.has
Update || app.updateState === 'submitted' || 
app.updateState === 'updating') && app.latest
Version) {
            html += ' <i class="fa
s fa-long-arrow-alt-right text-' + (app.updat
eState ? 'primary' : 'success') + ' mx-1"></i
> <strong class="text-' + (app.updateState ? 
'primary' : 'success') + '">' + escapeHtml(ap
p.latestVersion) + '</strong>';
        }
   
     html += '</div>';
        html += '<div 
class="text-muted small mt-1"><i class="fas f
a-building"></i> ' + escapeHtml(app.publisher
) + '</div>';
        if (showCheckbox) {
   
         html += '</div></div>'; // close che
ckbox wrapper divs
        }
        html += 
'</div>';
        html += '<div class="col-md
-3 text-center">';
        if (app.updateStat
e === 'submitted' || app.updateState === 'upd
ating') {
            html += '<span class="b
adge bg-primary"><span class="spinner-updatin
g"></span> Updating...</span>';
        } els
e {
            html += '<span class="badge b
g-' + stateClass + '">';
            if (stat
eIcon) html += '<i class="fas fa-' + stateIco
n + '"></i> ';
            html += stateText 
+ '</span>';
        }
        html += '</div
>';
        html += '<div class="col-md-3 tex
t-end">';
        if (app.updateState === 'su
bmitted' || app.updateState === 'updating') {

            html += '<span class="text-prima
ry"><i class="fas fa-sync-alt fa-spin"></i> I
n progress</span>';
        } else if (app.up
dateState === 'failed') {
            html +=
 '<button class="btn btn-outline-danger btn-s
m" onclick="updateSingleApp(\'' + escapeHtml(
app.uniqueName) + '\')"><i class="fas fa-redo
"></i> Retry</button>';
        } else if (ap
p.hasUpdate) {
            html += '<button c
lass="btn btn-success btn-sm" onclick="update
SingleApp(\'' + escapeHtml(app.uniqueName) + 
'\')"><i class="fas fa-download"></i> Update<
/button>';
        } else if (!app.instancePa
ckageId) {
            html += '<button class
="btn btn-primary btn-sm" onclick="installApp
(\'' + escapeHtml(app.uniqueName) + '\')"><i 
class="fas fa-plus"></i> Install</button>';
 
       } else {
            html += '<span cl
ass="text-success"><i class="fas fa-check-cir
cle"></i> Up to date</span>';
        }
     
   html += '</div>'; // close col-md-3
      
  html += '</div>'; // close row
        // E
rror detail spans full card width, below the 
row
        if (app.updateState === 'failed' 
&& app.updateError) {
            const error
Info = parseErrorMessage(app.updateError);
  
          html += '<div class="error-detail m
t-2">';
            html += '<div class="erro
r-summary"><i class="fas fa-exclamation-circl
e me-1"></i>' + escapeHtml(errorInfo.summary)
 + '</div>';
            if (errorInfo.detail
) {
                const rawId = 'err-' + (a
pp.uniqueName || '').replace(/[^a-zA-Z0-9]/g,
 '_');
                html += '<span class="
error-toggle" onclick="document.getElementByI
d(\'' + rawId + '\').classList.toggle(\'show\
')">';
                html += '<i class="fas
 fa-chevron-down me-1"></i>Show details</span
>';
                html += '<div class="erro
r-raw" id="' + rawId + '">' + escapeHtml(erro
rInfo.detail) + '</div>';
            }
     
       html += '</div>';
        }
        ht
ml += '</div>'; // close app-card
    }
    

    appsList.innerHTML = html;
}

// Fetch wi
th retry for transient server errors (500/502
/503/504)
async function fetchInstallWithRetr
y(url, appName, maxRetries = 3) {
    for (le
t attempt = 0; attempt <= maxRetries; attempt
++) {
        let response = await fetch(url,
 {
            method: 'POST',
            he
aders: {
                'Authorization': `Be
arer ${ppToken}`,
                'Content-Ty
pe': 'application/json'
            }
       
 });

        // Retry on 401 with a fresh to
ken
        if (response.status === 401) {
  
          console.log(`  ${appName}: 401, ref
reshing token...`);
            await refresh
PPToken();
            response = await fetch
(url, {
                method: 'POST',
     
           headers: {
                    'Au
thorization': `Bearer ${ppToken}`,
          
          'Content-Type': 'application/json'

                }
            });
           
 if (response.status < 500) return response;

        }

        // Retry on transient serv
er errors
        if (response.status >= 500 
&& attempt < maxRetries) {
            const 
delay = Math.min(2000 * Math.pow(2, attempt),
 15000);
            console.warn(`  ⚠ ${ap
pName}: ${response.status} on attempt ${attem
pt + 1}/${maxRetries + 1}, retrying in ${dela
y}ms...`);
            await new Promise(r =>
 setTimeout(r, delay));
            continue;

        }

        return response;
    }
}


// Update a single app
async function update
SingleApp(uniqueName) {
    const app = apps.
find(a => a.uniqueName === uniqueName);
    i
f (!app) return;
    
    // If retrying a fa
iled update, skip confirmation
    if (app.up
dateState !== 'failed') {
        if (!(await
 showModal({ title: 'Update App', message: 'I
nstall update for "' + app.name + '"?\n\nCurr
ent: ' + app.version + '\nNew: ' + (app.lates
tVersion || 'latest'), type: 'update', okText
: 'Update', okClass: 'btn-success-modal' })))
 {
            return;
        }
    }
    
 
   // Mark as updating and refresh display im
mediately
    app.updateState = 'submitted';

    app.updateError = null;
    displayApplic
ations();
    
    try {
        // Refresh t
oken to avoid stale auth
        await refres
hPPToken();
        
        const installUni
queName = app.catalogUniqueName || app.unique
Name;
        const url = `https://api.powerp
latform.com/appmanagement/environments/${envi
ronmentId}/applicationPackages/${installUniqu
eName}/install?api-version=2022-03-01-preview
`;
        
        console.log('Installing u
pdate:', installUniqueName, 'for', app.name);

        
        const response = await fetc
hInstallWithRetry(url, app.name);
        
  
      if (!response.ok) {
            const e
rrorText = await response.text();
           
 // Check if this is a SPA-only app that can'
t be updated via API
            if (response
.status === 400 && errorText.includes('Custom
 Install Experience')) {
                thro
w new Error('This app requires manual update 
through the Power Platform Admin Center. It c
annot be updated via API.');
            }
  
          throw new Error(response.status + '
 - ' + errorText);
        }
        
       
 // Keep as 'submitted' — user can Refresh 
to check later
        app.updateState = 'sub
mitted';
        app.hasUpdate = false; // Do
n't count it as a pending update anymore
    
    displayApplications();
        logUsage(1
, 0, [app.name]);
        
    } catch (error
) {
        console.error('Update error:', er
ror);
        app.updateState = 'failed';
   
     app.updateError = error.message;
       
 app.hasUpdate = true; // Keep as updatable s
o retry is possible
        displayApplicatio
ns();
        logUsage(0, 1, [app.name]);
   
 }
}

// Install an app
async function instal
lApp(uniqueName) {
    const app = apps.find(
a => a.uniqueName === uniqueName);
    if (!a
pp) return;
    
    if (!(await showModal({ 
title: 'Install App', message: 'Install "' + 
app.name + '"?', type: 'info', okText: 'Insta
ll', okClass: 'btn-success-modal' }))) {
    
    return;
    }
    
    app.updateState = 
'submitted';
    app.updateError = null;
    
displayApplications();
    
    try {
       
 await refreshPPToken();
        const url = 
`https://api.powerplatform.com/appmanagement/
environments/${environmentId}/applicationPack
ages/${app.uniqueName}/install?api-version=20
22-03-01-preview`;
        
        const res
ponse = await fetchInstallWithRetry(url, app.
name);
        
        if (!response.ok) {
 
           const errorText = await response.t
ext();
            throw new Error(response.s
tatus + ' - ' + errorText);
        }
       
 
        app.updateState = 'submitted';
    
    displayApplications();
        
    } cat
ch (error) {
        console.error('Install e
rror:', error);
        app.updateState = 'fa
iled';
        app.updateError = error.messag
e;
        displayApplications();
    }
}

//
 Update an installed app
async function reins
tallApp(uniqueName) {
    const app = apps.fi
nd(a => a.uniqueName === uniqueName);
    if 
(!app) return;
    
    if (app.updateState !
== 'failed') {
        if (!(await showModal(
{ title: 'Update App', message: 'Update "' + 
app.name + '"?\n\nCurrent version: ' + app.ve
rsion, type: 'update', okText: 'Update', okCl
ass: 'btn-success-modal' }))) {
            r
eturn;
        }
    }
    
    app.updateSta
te = 'submitted';
    app.updateError = null;

    displayApplications();
    
    try {
  
      await refreshPPToken();
        const u
rl = `https://api.powerplatform.com/appmanage
ment/environments/${environmentId}/applicatio
nPackages/${app.uniqueName}/install?api-versi
on=2022-03-01-preview`;
        
        cons
ole.log('Updating:', app.name);
        
    
    const response = await fetchInstallWithRe
try(url, app.name);
        
        if (!res
ponse.ok) {
            const errorText = awa
it response.text();
            throw new Err
or(response.status + ' - ' + errorText);
    
    }
        
        app.updateState = 'sub
mitted';
        app.hasUpdate = false;
     
   displayApplications();
        
    } catc
h (error) {
        console.error('Update err
or:', error);
        app.updateState = 'fail
ed';
        app.updateError = error.message;

        app.hasUpdate = true;
        displa
yApplications();
    }
}

// Update all apps

async function updateAllApps() {
    const ap
psToUpdate = apps.filter(a => a.hasUpdate && 
a.updateState !== 'submitted' && a.updateStat
e !== 'updating');
    
    if (appsToUpdate.
length === 0) {
        await showAlert('No U
pdates', 'No updates available.', 'info');
  
      return;
    }
    
    if (!(await show
UpdateConfirm(appsToUpdate))) {
        retur
n;
    }
    
    let successCount = 0;
    l
et failCount = 0;
    
    // Mark all as upd
ating immediately
    for (const app of appsT
oUpdate) {
        app.updateState = 'submitt
ed';
        app.updateError = null;
    }
  
  displayApplications();
    
    showLoading
('Installing updates...', '0 of ' + appsToUpd
ate.length);
    
    // Refresh token before
 starting the batch
    await refreshPPToken(
);
    let completed = 0;
    const BATCH_SIZ
E = 5;
    
    for (let i = 0; i < appsToUpd
ate.length; i += BATCH_SIZE) {
        const 
batch = appsToUpdate.slice(i, i + BATCH_SIZE)
;
        await Promise.allSettled(batch.map(
async (app) => {
            try {
          
      const installUniqueName = app.catalogUn
iqueName || app.uniqueName;
                c
onst url = `https://api.powerplatform.com/app
management/environments/${environmentId}/appl
icationPackages/${installUniqueName}/install?
api-version=2022-03-01-preview`;
            
    
                console.log('Updating:',
 app.name, 'using package:', installUniqueNam
e);
                
                const re
sponse = await fetchInstallWithRetry(url, app
.name);
                
                if (
response.ok) {
                    successCou
nt++;
                    app.updateState = '
submitted';
                    app.hasUpdate
 = false;
                } else {
          
          failCount++;
                    co
nst errorText = await response.text();
      
              app.updateState = 'failed';
   
                 app.updateError = response.s
tatus + ' - ' + errorText;
                  
  app.hasUpdate = true;
                    c
onsole.error('Failed to update ' + app.name +
 ':', response.status);
                }
   
         } catch (error) {
                fa
ilCount++;
                app.updateState = 
'failed';
                app.updateError = e
rror.message;
                app.hasUpdate =
 true;
                console.error('Error u
pdating ' + app.name + ':', error);
         
   }
            completed++;
            doc
ument.getElementById('loadingDetails').textCo
ntent = completed + ' of ' + appsToUpdate.len
gth + ' completed';
        }));
    }
    
 
   hideLoading();
    displayApplications();

    
    // Log usage to Supabase
    logUsag
e(successCount, failCount, appsToUpdate.map(a
 => a.name));
    
    if (failCount === 0) {

        await showAlert('Updates Started', '
All ' + successCount + ' updates submitted su
ccessfully! Updates are running in the backgr
ound and may take several minutes. Click "Ref
resh" to check progress.', 'success');
    } 
else {
        await showAlert('Updates Submi
tted', successCount + ' update' + (successCou
nt !== 1 ? 's' : '') + ' submitted successful
ly. ' + failCount + ' failed — see details 
below. You can retry failed updates individua
lly.', 'warning');
    }
}

// Update all ins
talled apps
async function reinstallAllApps()
 {
    const appsToUpdate = apps.filter(a => 
a.hasUpdate && a.updateState !== 'submitted' 
&& a.updateState !== 'updating');
    
    if
 (appsToUpdate.length === 0) {
        await 
showAlert('All Up to Date', 'No updates avail
able. All apps are up to date.', 'success');

        return;
    }
    
    if (!(await sh
owUpdateConfirm(appsToUpdate))) {
        ret
urn;
    }
    
    let successCount = 0;
   
 let failCount = 0;
    
    // Mark all as u
pdating immediately
    for (const app of app
sToUpdate) {
        app.updateState = 'submi
tted';
        app.updateError = null;
    }

    displayApplications();
    
    showLoadi
ng('Updating apps...', '0 of ' + appsToUpdate
.length);
    
    // Refresh token before st
arting the batch to avoid 401s mid-loop
    a
wait refreshPPToken();
    let completed = 0;

    const BATCH_SIZE = 5;
    
    for (let 
i = 0; i < appsToUpdate.length; i += BATCH_SI
ZE) {
        const batch = appsToUpdate.slic
e(i, i + BATCH_SIZE);
        await Promise.a
llSettled(batch.map(async (app) => {
        
    try {
                const installUnique
Name = app.catalogUniqueName || app.uniqueNam
e;
                const url = `https://api.p
owerplatform.com/appmanagement/environments/$
{environmentId}/applicationPackages/${install
UniqueName}/install?api-version=2022-03-01-pr
eview`;
                
                cons
ole.log(`Updating: ${app.name}`);
           
     console.log(`  installed pkg: ${app.uniq
ueName}`);
                console.log(`  ins
tall pkg:   ${installUniqueName}${installUniq
ueName !== app.uniqueName ? ' ← CATALOG' : 
' ← SAME AS INSTALLED (may be no-op!)'}`); 

                console.log(`  version: ${ap
p.version} → ${app.latestVersion || 'latest
'}`);
                
                const 
response = await fetchInstallWithRetry(url, a
pp.name);
                
                if
 (response.ok) {
                    const re
sponseBody = await response.text();
         
           console.log(`  ✓ ${response.stat
us} OK`, responseBody ? responseBody.substrin
g(0, 200) : '(empty body)');
                
    successCount++;
                    app.u
pdateState = 'submitted';
                   
 app.hasUpdate = false;
                } els
e {
                    failCount++;
        
            const errorText = await response.
text();
                    app.updateState =
 'failed';
                    app.updateErro
r = response.status + ' - ' + errorText;
    
                app.hasUpdate = true;
       
             console.error('Failed to update 
' + app.name + ':', response.status);
       
         }
            } catch (error) {
    
            failCount++;
                app.
updateState = 'failed';
                app.u
pdateError = error.message;
                a
pp.hasUpdate = true;
                console.
error('Error updating ' + app.name + ':', err
or);
            }
            completed++;
 
           document.getElementById('loadingDe
tails').textContent = completed + ' of ' + ap
psToUpdate.length + ' completed';
        }))
;
    }
    
    hideLoading();
    displayAp
plications();
    
    // Log usage to Supaba
se
    logUsage(successCount, failCount, apps
ToUpdate.map(a => a.name));
    
    if (fail
Count === 0) {
        await showAlert('Updat
es Submitted', 'All ' + successCount + ' upda
te requests submitted! Updates are running in
 the background. Click "Refresh" to check pro
gress.', 'success');
    } else {
        awa
it showAlert('Updates Submitted', successCoun
t + ' update' + (successCount !== 1 ? 's' : '
') + ' submitted. ' + failCount + ' failed �
� see details below. You can retry failed upd
ates individually.', 'warning');
    }
}

// 
Toggle individual app selection for multi-sel
ect
function toggleAppSelection(uniqueName, c
hecked) {
    if (checked) {
        selected
Apps.add(uniqueName);
    } else {
        se
lectedApps.delete(uniqueName);
    }
    upda
teSelectedButton();
}

// Show/hide the "Upda
te Selected" button and update count
function
 updateSelectedButton() {
    const btn = doc
ument.getElementById('updateSelectedBtn');
  
  const countSpan = document.getElementById('
selectedCount');
    if (!btn || !countSpan) 
return;
    const count = selectedApps.size;

    countSpan.textContent = count;
    if (co
unt > 0) {
        btn.classList.remove('d-no
ne');
    } else {
        btn.classList.add(
'd-none');
    }
}

// Update only the select
ed apps
async function updateSelectedApps() {

    if (selectedApps.size === 0) return;

  
  const appsToUpdate = apps.filter(a =>
     
   selectedApps.has(a.uniqueName) &&
        
(a.hasUpdate || a.updateState === 'failed') &
&
        a.updateState !== 'submitted' &&
  
      a.updateState !== 'updating'
    );

  
  if (appsToUpdate.length === 0) {
        aw
ait showAlert('Nothing to Update', 'The selec
ted apps have no pending updates.', 'info');

        return;
    }

    if (!(await showUp
dateConfirm(appsToUpdate))) {
        return;

    }

    let successCount = 0;
    let fai
lCount = 0;

    // Mark all as updating imme
diately
    for (const app of appsToUpdate) {

        app.updateState = 'submitted';
     
   app.updateError = null;
    }
    displayA
pplications();

    showLoading('Updating sel
ected apps...', '0 of ' + appsToUpdate.length
);

    // Refresh token before starting the 
batch
    await refreshPPToken();
    let com
pleted = 0;
    const BATCH_SIZE = 5;

    fo
r (let i = 0; i < appsToUpdate.length; i += B
ATCH_SIZE) {
        const batch = appsToUpda
te.slice(i, i + BATCH_SIZE);
        await Pr
omise.allSettled(batch.map(async (app) => {
 
           try {
                const instal
lUniqueName = app.catalogUniqueName || app.un
iqueName;
                const url = `https:
//api.powerplatform.com/appmanagement/environ
ments/${environmentId}/applicationPackages/${
installUniqueName}/install?api-version=2022-0
3-01-preview`;

                console.log(`
Updating: ${app.name}`);
                cons
ole.log(`  installed pkg: ${app.uniqueName}`)
;
                console.log(`  install pkg:
   ${installUniqueName}${installUniqueName !=
= app.uniqueName ? ' ← CATALOG' : ' ← SAM
E AS INSTALLED (may be no-op!)'}`);
         
       console.log(`  version: ${app.version}
 → ${app.latestVersion || 'latest'}`);

   
             const response = await fetchInst
allWithRetry(url, app.name);

               
 if (response.ok) {
                    const
 responseBody = await response.text();
      
              console.log(`  ✓ ${response.s
tatus} OK`, responseBody ? responseBody.subst
ring(0, 200) : '(empty body)');
             
       successCount++;
                    ap
p.updateState = 'submitted';
                
    app.hasUpdate = false;
                } 
else {
                    failCount++;
     
               const errorText = await respon
se.text();
                    app.updateStat
e = 'failed';
                    app.updateE
rror = response.status + ' - ' + errorText;
 
                   app.hasUpdate = true;
    
                console.error('Failed to upda
te ' + app.name + ':', response.status);
    
            }
            } catch (error) {
 
               failCount++;
                a
pp.updateState = 'failed';
                ap
p.updateError = error.message;
              
  app.hasUpdate = true;
                conso
le.error('Error updating ' + app.name + ':', 
error);
            }
            completed++
;
            document.getElementById('loadin
gDetails').textContent = completed + ' of ' +
 appsToUpdate.length + ' completed';
        
}));
    }

    hideLoading();
    selectedAp
ps.clear();
    displayApplications();

    /
/ Log usage to Supabase
    logUsage(successC
ount, failCount, appsToUpdate.map(a => a.name
));

    if (failCount === 0) {
        await
 showAlert('Updates Submitted', 'All ' + succ
essCount + ' selected update(s) submitted! Up
dates are running in the background. Click "R
efresh" to check progress.', 'success');
    
} else {
        await showAlert('Updates Sub
mitted', successCount + ' submitted, ' + fail
Count + ' failed — see details below.', 'wa
rning');
    }
}

// Logout
function handleLo
gout() {
    showModal({ title: 'Logout', mes
sage: 'Are you sure you want to logout?', typ
e: 'warning', okText: 'Logout', okClass: 'btn
-danger-modal' }).then(confirmed => {
       
 if (!confirmed) return;
        accessToken 
= null;
        ppToken = null;
        envir
onmentId = null;
        apps = [];
        

        if (msalInstance) {
            const
 accounts = msalInstance.getAllAccounts();
  
          if (accounts.length > 0) {
        
        // Clear MSAL cache so auto-login won
't fire again
                msalInstance.lo
goutRedirect({ account: accounts[0] }).catch(
() => {});
            } else {
             
   // No accounts but clear cache anyway
    
            msalInstance.clearCache().catch((
) => {});
            }
        }
        
  
      document.getElementById('appsSection').
classList.add('hidden');
        document.get
ElementById('authSection').classList.remove('
hidden');
    });
}

// ── Usage Tracking
 (Supabase) ───────────
───────────────
───────
function getSupabaseCon
fig() {
    return { url: SUPABASE_URL, key: 
SUPABASE_KEY };
}

function getCurrentUserEma
il() {
    if (!msalInstance) return null;
  
  const accounts = msalInstance.getAllAccount
s();
    if (accounts.length > 0) {
        r
eturn accounts[0].username || accounts[0].nam
e || null;
    }
    return null;
}

function
 getCurrentClientId() {
    // Get the client
 ID from stored credentials (used at login)
 
   const savedCreds = localStorage.getItem('d
365_app_updater_creds') || 
                 
      sessionStorage.getItem('d365_app_update
r_creds');
    if (savedCreds) {
        try 
{
            const creds = JSON.parse(savedC
reds);
            return creds.clientId || '
';
        } catch (e) {
            console.
warn('Could not parse saved credentials');
  
      }
    }
    return '';
}

function getC
urrentTenantId() {
    // Get the tenant ID f
rom stored credentials or from MSAL account
 
   const savedCreds = localStorage.getItem('d
365_app_updater_creds') || 
                 
      sessionStorage.getItem('d365_app_update
r_creds');
    if (savedCreds) {
        try 
{
            const creds = JSON.parse(savedC
reds);
            if (creds.tenantId) return
 creds.tenantId;
        } catch (e) {
      
      console.warn('Could not parse saved cre
dentials for tenant');
        }
    }
    //
 Fallback: get from MSAL account
    if (msal
Instance) {
        const accounts = msalInst
ance.getAllAccounts();
        if (accounts.l
ength > 0 && accounts[0].tenantId) {
        
    return accounts[0].tenantId;
        }
  
  }
    return '';
}

async function logUsage
(successCount, failCount, appNames) {
    con
st cfg = getSupabaseConfig();
    if (!cfg) {

        console.log('Usage tracking: Supabas
e not configured, skipping.');
        return
;
    }

    const record = {
        timesta
mp: new Date().toISOString(),
        user_em
ail: getCurrentUserEmail() || 'unknown',
    
    org_url: currentOrgUrl || '',
        env
ironment_id: environmentId || '',
        suc
cess_count: successCount || 0,
        fail_c
ount: failCount || 0,
        total_apps: (su
ccessCount || 0) + (failCount || 0),
        
app_names: (appNames || []).join(', ')
    };


    try {
        const resp = await fetch(
`${cfg.url}/rest/v1/usage_logs`, {
          
  method: 'POST',
            headers: {
    
            'apikey': cfg.key,
              
  'Authorization': `Bearer ${cfg.key}`,
     
           'Content-Type': 'application/json'
,
                'Prefer': 'return=minimal'

            },
            body: JSON.stringi
fy(record)
        });

        if (resp.ok) 
{
            console.log('Usage logged succe
ssfully:', record);
        } else {
        
    console.warn('Usage logging failed:', res
p.status, await resp.text());
        }
    }
 catch (error) {
        console.warn('Usage 
logging error (non-critical):', error.message
);
    }
}

// ── Auto-Update Scheduling 
(Supabase) ───────────�
��──────────────�
��──────
let scheduleLoaded = fal
se;

function toggleScheduleDetails() {
    c
onst details = document.getElementById('sched
uleDetails');
    const enabled = document.ge
tElementById('scheduleEnabled').checked;
    
if (enabled) {
        details.style.display 
= details.style.display === 'none' ? 'block' 
: 'none';
    }
}

function handleScheduleTog
gle() {
    const enabled = document.getEleme
ntById('scheduleEnabled').checked;
    const 
details = document.getElementById('scheduleDe
tails');
    details.style.display = enabled 
? 'block' : 'none';
    
    if (!enabled) {

        // Disable schedule in Supabase
     
   disableSchedule();
    }
}

function toggl
eSecretVisibility() {
    const secretInput =
 document.getElementById('scheduleClientSecre
t');
    const icon = document.getElementById
('secretToggleIcon');
    if (secretInput.typ
e === 'password') {
        secretInput.type 
= 'text';
        icon.className = 'fas fa-ey
e-slash';
    } else {
        secretInput.ty
pe = 'password';
        icon.className = 'fa
s fa-eye';
    }
}

function showCredentialsH
elp() {
    const clientId = getCurrentClient
Id();
    const clientIdDisplay = clientId ? 
`<code style="background:#e5e7eb; padding: 2p
x 6px; border-radius: 4px;">${clientId}</code
>` : '(from your login)';
    
    showModal(
{
        title: 'How to Create a Client Secr
et',
        body: `<div style="text-align: l
eft; font-size: 0.9rem;">
<p style="backgroun
d: #f0f9ff; padding: 10px; border-radius: 6px
; border-left: 3px solid #0078d4;">
<strong>Y
our Client ID:</strong> ${clientIdDisplay}<br
>
<small>This is the App Registration you use
d to log in.</small>
</p>

<p><strong>Step 1:
 Open Your App Registration</strong><br>
<a h
ref="https://portal.azure.com/#blade/Microsof
t_AAD_RegisteredApps/ApplicationsListBlade" t
arget="_blank">Click here to open Azure App R
egistrations</a><br>
Find and click on your a
pp (the one with the Client ID above)</p>

<p
><strong>Step 2: Create a Client Secret</stro
ng><br>
- In the left menu, click <strong>"Ce
rtificates & secrets"</strong><br>
- Click <s
trong>"+ New client secret"</strong><br>
- De
scription: <code>Scheduler</code><br>
- Expir
es: Choose 24 months (recommended)<br>
- Clic
k <strong>"Add"</strong></p>

<p><strong>Step
 3: Copy the Secret Value</strong><br>
<span 
style="color: #dc2626;">⚠️ IMPORTANT: Cop
y the <strong>Value</strong> key (not the Sec
ret ID) immediately!</span><br>
It will only 
be shown once. If you lose it, you'll need to
 create a new one.</p>

<p><strong>Step 4: Pa
ste Here and Save</strong><br>
Paste the secr
et value in the field above and click "Save S
chedule"</p>

<hr style="margin: 15px 0;">
<p
 style="color: #059669;"><strong>✅ What hap
pens automatically:</strong><br>
- API permis
sions are added to your app<br>
- Admin conse
nt is granted<br>
- Application User is creat
ed in Dataverse<br>
- System Administrator ro
le is assigned</p>
</div>`,
        type: 'in
fo',
        confirmOnly: true
    });
}

asy
nc function loadSchedule() {
    if (schedule
Loaded) return;
    
    const cfg = getSupab
aseConfig();
    if (!cfg) return;
    
    c
onst userEmail = getCurrentUserEmail();
    c
onst envId = environmentId || '';
    
    if
 (!userEmail || !envId) return;
    
    try 
{
        const resp = await fetch(
         
   `${cfg.url}/rest/v1/update_schedules?user_
email=eq.${encodeURIComponent(userEmail)}&env
ironment_id=eq.${encodeURIComponent(envId)}&s
elect=*`,
            {
                heade
rs: {
                    'apikey': cfg.key,

                    'Authorization': `Bearer 
${cfg.key}`
                }
            }
 
       );
        
        if (resp.ok) {
   
         const schedules = await resp.json();

            if (schedules.length > 0) {
    
            const schedule = schedules[0];
  
              document.getElementById('schedu
leEnabled').checked = schedule.enabled;
     
           
                // Convert UTC va
lues back to local timezone for display
     
           const localSchedule = convertFromU
TC(schedule.day_of_week, schedule.time_utc, s
chedule.timezone || 'UTC');
                d
ocument.getElementById('scheduleDay').value =
 localSchedule.day_of_week_local;
           
     document.getElementById('scheduleTime').
value = localSchedule.time_local;
           
     document.getElementById('scheduleTimezon
e').value = schedule.timezone || 'UTC';
     
           // Secret is stored securely - jus
t indicate it's set
                const sec
retInput = document.getElementById('scheduleC
lientSecret');
                if (schedule.h
as_secret) {
                    secretInput.
placeholder = '(secret securely saved - leave
 blank to keep)';
                    secretI
nput.value = ''; // Clear any value
         
       } else {
                    // Check 
if user has a secret saved for another enviro
nment
                    try {
             
           const otherResp = await fetch(
   
                         `${cfg.url}/rest/v1/
update_schedules?user_email=eq.${encodeURICom
ponent(userEmail)}&has_secret=eq.true&select=
id`,
                            { headers: {
 'apikey': cfg.key, 'Authorization': `Bearer 
${cfg.key}` } }
                        );
  
                      const others = await ot
herResp.json();
                        if (o
thers.length > 0) {
                         
   secretInput.placeholder = '(will reuse sec
ret from another environment)';
             
               secretInput.value = '';
      
                  }
                    } cat
ch (e) { /* ignore */ }
                }
   
             document.getElementById('schedul
eDetails').style.display = schedule.enabled ?
 'block' : 'none';
                updateSche
duleStatus(schedule);
            }
        }

        scheduleLoaded = true;
    } catch (
error) {
        console.warn('Failed to load
 schedule:', error.message);
    }
}

functio
n updateScheduleStatus(schedule) {
    const 
statusEl = document.getElementById('scheduleS
tatus');
    if (!schedule || !schedule.enabl
ed) {
        statusEl.innerHTML = '<i class=
"fas fa-info-circle"></i> Schedule not config
ured';
        statusEl.className = 'schedule
-status';
        return;
    }
    
    cons
t days = ['Sunday', 'Monday', 'Tuesday', 'Wed
nesday', 'Thursday', 'Friday', 'Saturday'];
 
   
    // Convert UTC values back to local t
imezone for display
    const localSchedule =
 convertFromUTC(schedule.day_of_week, schedul
e.time_utc, schedule.timezone || 'UTC');
    
const dayName = days[localSchedule.day_of_wee
k_local];
    const timeDisplay = formatTimeD
isplay(localSchedule.time_local);
    
    //
 Also show UTC for clarity
    const utcDayNa
me = days[schedule.day_of_week];
    const ut
cTimeDisplay = formatTimeDisplay(schedule.tim
e_utc);
    const lastRun = schedule.last_run
_at ? new Date(schedule.last_run_at).toLocale
String() : 'Never';
    
    // If UTC is dif
ferent from display, show both
    const show
UtcInfo = (schedule.day_of_week !== localSche
dule.day_of_week_local || schedule.time_utc !
== localSchedule.time_local);
    
    status
El.innerHTML = `<i class="fas fa-check-circle
"></i> Scheduled: Every <strong>${dayName}</s
trong> at <strong>${timeDisplay}</strong> (${
schedule.timezone || 'UTC'})` +
        (show
UtcInfo ? `<br><small style="color: #6b7280;"
>Runs at: ${utcDayName} ${utcTimeDisplay} UTC
</small>` : '') +
        `<br><small>Last ru
n: ${lastRun}</small>`;
    statusEl.classNam
e = 'schedule-status active';
}

function for
matTimeDisplay(time24) {
    const [hours, mi
nutes] = time24.split(':');
    const h = par
seInt(hours, 10);
    const ampm = h >= 12 ? 
'PM' : 'AM';
    const h12 = h % 12 || 12;
  
  return `${h12}:${minutes} ${ampm}`;
}

/**

 * Convert a local time in a specific timezon
e to UTC day and time
 * @param {number} dayO
fWeek - Day of week (0=Sunday, 6=Saturday) in
 local timezone
 * @param {string} time - Tim
e in HH:MM format in local timezone
 * @param
 {string} timezone - IANA timezone (e.g., 'Am
erica/New_York')
 * @returns {object} { day_o
f_week_utc, time_utc } - Converted to UTC
 */

function convertToUTC(dayOfWeek, time, timez
one) {
    // Parse the time
    const [hours
, minutes] = time.split(':').map(Number);
   
 
    // Use a reference date - find next occ
urrence of the target day
    // Start with a
 known date and adjust to the target day of w
eek
    const now = new Date();
    const cur
rentDay = now.getDay();
    let daysUntilTarg
et = (dayOfWeek - currentDay + 7) % 7;
    if
 (daysUntilTarget === 0) daysUntilTarget = 7;
 // Next week if today
    
    const targetD
ate = new Date(now);
    targetDate.setDate(n
ow.getDate() + daysUntilTarget);
    targetDa
te.setHours(0, 0, 0, 0);
    
    // Format t
his date in the target timezone to get the ye
ar-month-day
    const formatter = new Intl.D
ateTimeFormat('en-CA', { // en-CA gives YYYY-
MM-DD format
        timeZone: timezone,
    
    year: 'numeric',
        month: '2-digit'
,
        day: '2-digit'
    });
    const lo
calDate = formatter.format(targetDate);
    

    // Create the full datetime string: YYYY-
MM-DDTHH:MM in the target timezone
    const 
dateTimeString = `${localDate}T${hours.toStri
ng().padStart(2, '0')}:${minutes.toString().p
adStart(2, '0')}:00`;
    
    // Parse this 
as if it's in the target timezone by using th
e locale-specific parsing
    // We'll create
 a date and check what UTC time corresponds t
o this local time
    
    // Use this approa
ch: create UTC dates and check which one show
s the right time in the target timezone
    /
/ Binary search would be most efficient, but 
let's use a simpler approach for clarity
    

    // Get the approximate UTC time by tryin
g different offsets
    for (let offsetHours 
= -12; offsetHours <= 14; offsetHours++) {
  
      const testDate = new Date(`${localDate}
T${hours.toString().padStart(2, '0')}:${minut
es.toString().padStart(2, '0')}:00Z`);
      
  testDate.setUTCHours(testDate.getUTCHours()
 - offsetHours);
        
        // Check wh
at this UTC time shows in the target timezone

        const checkFormatter = new Intl.Date
TimeFormat('en-CA', {
            timeZone: t
imezone,
            year: 'numeric',
       
     month: '2-digit',
            day: '2-di
git',
            hour: '2-digit',
          
  minute: '2-digit',
            hour12: fals
e
        });
        
        const parts = 
checkFormatter.formatToParts(testDate);
     
   const partsMap = {};
        parts.forEach
(p => { if (p.type !== 'literal') partsMap[p.
type] = p.value; });
        
        const c
heckDate = `${partsMap.year}-${partsMap.month
}-${partsMap.day}`;
        const checkHour =
 parseInt(partsMap.hour, 10);
        const c
heckMinute = parseInt(partsMap.minute, 10);
 
       
        if (checkDate === localDate &
& checkHour === hours && checkMinute === minu
tes) {
            // Found the matching UTC 
time!
            const utcDay = testDate.get
UTCDay();
            const utcHour = testDat
e.getUTCHours();
            const utcMinute 
= testDate.getUTCMinutes();
            
    
        return {
                day_of_week_
utc: utcDay,
                time_utc: `${utc
Hour.toString().padStart(2, '0')}:${utcMinute
.toString().padStart(2, '0')}`
            };

        }
    }
    
    // Fallback: no con
version (shouldn't happen with valid timezone
s)
    return {
        day_of_week_utc: dayO
fWeek,
        time_utc: time
    };
}

/**
 
* Convert UTC day and time to local timezone

 * @param {number} dayOfWeekUtc - Day of week
 in UTC (0=Sunday, 6=Saturday)
 * @param {str
ing} timeUtc - Time in HH:MM format in UTC
 *
 @param {string} timezone - IANA timezone (e.
g., 'America/New_York')
 * @returns {object} 
{ day_of_week_local, time_local } - Converted
 to local timezone
 */
function convertFromUT
C(dayOfWeekUtc, timeUtc, timezone) {
    // P
arse the UTC time
    const [hours, minutes] 
= timeUtc.split(':').map(Number);
    
    //
 Create a UTC date with the specified day and
 time
    const now = new Date();
    const c
urrentDay = now.getUTCDay();
    let daysUnti
lTarget = (dayOfWeekUtc - currentDay + 7) % 7
;
    if (daysUntilTarget === 0) daysUntilTar
get = 7; // Next week if today
    
    const
 utcDate = new Date(now);
    utcDate.setUTCD
ate(now.getUTCDate() + daysUntilTarget);
    
utcDate.setUTCHours(hours, minutes, 0, 0);
  
  
    // Format this UTC date in the target 
timezone
    const formatter = new Intl.DateT
imeFormat('en-CA', {
        timeZone: timezo
ne,
        year: 'numeric',
        month: '
2-digit',
        day: '2-digit',
        hou
r: '2-digit',
        minute: '2-digit',
    
    hour12: false,
        weekday: 'short'
 
   });
    
    const parts = formatter.forma
tToParts(utcDate);
    const partsMap = {};
 
   parts.forEach(p => { if (p.type !== 'liter
al') partsMap[p.type] = p.value; });
    
   
 const localHour = parseInt(partsMap.hour, 10
);
    const localMinute = parseInt(partsMap.
minute, 10);
    
    // Get the day of week 
in the local timezone
    const weekdayMap = 
{ 'Sun': 0, 'Mon': 1, 'Tue': 2, 'Wed': 3, 'Th
u': 4, 'Fri': 5, 'Sat': 6 };
    const localD
ay = weekdayMap[partsMap.weekday] ?? dayOfWee
kUtc;
    
    return {
        day_of_week_l
ocal: localDay,
        time_local: `${localH
our.toString().padStart(2, '0')}:${localMinut
e.toString().padStart(2, '0')}`
    };
}

asy
nc function saveSchedule() {
    const cfg = 
getSupabaseConfig();
    if (!cfg) {
        
showError('Scheduling requires Supabase confi
guration.');
        return;
    }
    
    c
onst userEmail = getCurrentUserEmail();
    c
onst envId = environmentId || '';
    const o
rgUrl = currentOrgUrl || '';
    
    if (!us
erEmail || !envId) {
        showError('Pleas
e connect to an environment first.');
       
 return;
    }
    
    const saveBtn = docum
ent.getElementById('scheduleSaveBtn');
    co
nst originalText = saveBtn.innerHTML;
    sav
eBtn.disabled = true;
    saveBtn.innerHTML =
 '<i class="fas fa-spinner fa-spin"></i> Savi
ng...';
    
    // Get user's selected value
s (in their local timezone)
    const selecte
dDay = parseInt(document.getElementById('sche
duleDay').value, 10);
    const selectedTime 
= document.getElementById('scheduleTime').val
ue;
    const selectedTimezone = document.get
ElementById('scheduleTimezone').value;
    
 
   // Convert to UTC
    const utcSchedule = 
convertToUTC(selectedDay, selectedTime, selec
tedTimezone);
    
    const schedule = {
   
     user_email: userEmail,
        environme
nt_id: envId,
        org_url: orgUrl,
      
  enabled: document.getElementById('scheduleE
nabled').checked,
        day_of_week: utcSch
edule.day_of_week_utc,  // CONVERTED to UTC
 
       time_utc: utcSchedule.time_utc,       
      // CONVERTED to UTC
        timezone: s
electedTimezone,                 // Store ori
ginal timezone for display
        client_id:
 getCurrentClientId(),
        tenant_id: get
CurrentTenantId(),
        updated_at: new Da
te().toISOString()
    };
    
    // Only in
clude client_secret if user entered a new one

    const newSecret = document.getElementByI
d('scheduleClientSecret').value.trim();
    /
/ Note: secret is stored in separate secure t
able, not in schedule
    
    // Track if se
cret is set (for validation and UI)
    if (n
ewSecret) {
        schedule.has_secret = tru
e;
    }
    
    try {
        // Upsert: tr
y to update first, then insert if not exists

        const checkResp = await fetch(
      
      `${cfg.url}/rest/v1/update_schedules?us
er_email=eq.${encodeURIComponent(userEmail)}&
environment_id=eq.${encodeURIComponent(envId)
}&select=id,has_secret`,
            {
      
          headers: {
                    'api
key': cfg.key,
                    'Authoriza
tion': `Bearer ${cfg.key}`
                }

            }
        );
        
        con
st existing = await checkResp.json();
       
 let resp;
        
        // Check if user 
already has a secret saved for ANY environmen
t (same client_id)
        let hasExistingSec
ret = existing.length > 0 && existing[0].has_
secret;
        let existingSecretScheduleId 
= null;
        
        if (!hasExistingSecr
et && !newSecret) {
            // Look for a
 secret from another environment with same cl
ient_id
            const otherResp = await f
etch(
                `${cfg.url}/rest/v1/upd
ate_schedules?user_email=eq.${encodeURICompon
ent(userEmail)}&has_secret=eq.true&select=id,
client_id`,
                {
               
     headers: {
                        'apik
ey': cfg.key,
                        'Author
ization': `Bearer ${cfg.key}`
               
     }
                }
            );
     
       const otherSchedules = await otherResp
.json();
            if (otherSchedules.lengt
h > 0) {
                hasExistingSecret = 
true;
                existingSecretScheduleI
d = otherSchedules[0].id;
            }
     
   }
        
        // Validate credentials
 - only require secret if none exists anywher
e
        if (schedule.enabled && !newSecret 
&& !hasExistingSecret) {
            saveBtn.
disabled = false;
            saveBtn.innerHT
ML = originalText;
            showError('Cli
ent Secret is required for scheduled updates.
');
            return;
        }
        
  
      if (existing.length > 0) {
            
// Update - don't overwrite secret if not pro
vided
            const scheduleId = existing
[0].id;
            resp = await fetch(
     
           `${cfg.url}/rest/v1/update_schedul
es?id=eq.${scheduleId}`,
                {
  
                  method: 'PATCH',
          
          headers: {
                        
'apikey': cfg.key,
                        'A
uthorization': `Bearer ${cfg.key}`,
         
               'Content-Type': 'application/j
son',
                        'Prefer': 'retu
rn=representation'
                    },
   
                 body: JSON.stringify(schedul
e)
                }
            );
         
   
            // If new secret provided, sa
ve it directly to the schedule record AND the
 secure secrets table
            if (resp.ok
 && newSecret) {
                // Primary: 
save directly to update_schedules.client_secr
et (workflow reads this)
                awai
t fetch(
                    `${cfg.url}/rest
/v1/update_schedules?id=eq.${scheduleId}`,
  
                  {
                        m
ethod: 'PATCH',
                        heade
rs: {
                            'apikey': c
fg.key,
                            'Authoriz
ation': `Bearer ${cfg.key}`,
                
            'Content-Type': 'application/json
'
                        },
                
        body: JSON.stringify({ client_secret:
 newSecret, has_secret: true })
             
       }
                );
                /
/ Secondary: also try schedule_secrets table

                await saveSecretSecurely(cfg,
 scheduleId, newSecret);
            }
      
  } else {
            // Insert
            
schedule.created_at = new Date().toISOString(
);
            resp = await fetch(
          
      `${cfg.url}/rest/v1/update_schedules`,

                {
                    method:
 'POST',
                    headers: {
     
                   'apikey': cfg.key,
       
                 'Authorization': `Bearer ${c
fg.key}`,
                        'Content-Ty
pe': 'application/json',
                    
    'Prefer': 'return=representation'
       
             },
                    body: JSO
N.stringify(schedule)
                }
     
       );
            
            // Save se
cret directly to the schedule record AND secu
re secrets table
            if (resp.ok && n
ewSecret) {
                const savedSchedu
le = await resp.json();
                const
 scheduleId = Array.isArray(savedSchedule) ? 
savedSchedule[0].id : savedSchedule.id;
     
           // Primary: save directly to updat
e_schedules.client_secret (workflow reads thi
s)
                await fetch(
             
       `${cfg.url}/rest/v1/update_schedules?i
d=eq.${scheduleId}`,
                    {
  
                      method: 'PATCH',
      
                  headers: {
                
            'apikey': cfg.key,
              
              'Authorization': `Bearer ${cfg.
key}`,
                            'Content-T
ype': 'application/json'
                    
    },
                        body: JSON.str
ingify({ client_secret: newSecret, has_secret
: true })
                    }
             
   );
                // Secondary: also try 
schedule_secrets table
                await 
saveSecretSecurely(cfg, scheduleId, newSecret
);
                // Re-wrap for later use
 
               resp = { ok: true, json: async
 () => savedSchedule };
            } else if
 (resp.ok && existingSecretScheduleId) {
    
            // Reuse secret from another envi
ronment - copy the actual secret value
      
          const savedSchedule = await resp.js
on();
                const scheduleId = Arra
y.isArray(savedSchedule) ? savedSchedule[0].i
d : savedSchedule.id;
                
      
          // Read the secret from the other s
chedule
                try {
               
     const srcResp = await fetch(
           
             `${cfg.url}/rest/v1/update_sched
ules?id=eq.${existingSecretScheduleId}&select
=client_secret`,
                        { he
aders: { 'apikey': cfg.key, 'Authorization': 
`Bearer ${cfg.key}` } }
                    )
;
                    const srcData = await s
rcResp.json();
                    if (srcDat
a.length > 0 && srcData[0].client_secret) {
 
                       // Copy secret to the 
new schedule
                        await fe
tch(
                            `${cfg.url}/
rest/v1/update_schedules?id=eq.${scheduleId}`
,
                            {
             
                   method: 'PATCH',
         
                       headers: {
           
                         'apikey': cfg.key,
 
                                   'Authoriza
tion': `Bearer ${cfg.key}`,
                 
                   'Content-Type': 'applicati
on/json'
                                },
 
                               body: JSON.str
ingify({ client_secret: srcData[0].client_sec
ret, has_secret: true })
                    
        }
                        );
        
                console.log('✅ Secret copie
d from schedule', existingSecretScheduleId);

                    }
                } catch
 (e) {
                    console.warn('Coul
d not copy secret:', e.message);
            
    }
                resp = { ok: true, json
: async () => savedSchedule };
            }

        }
        
        if (resp.ok) {
   
         const saved = await resp.json();
   
         updateScheduleStatus(Array.isArray(s
aved) ? saved[0] : saved);
            
     
       // Build conversion info message
     
       const days = ['Sunday', 'Monday', 'Tue
sday', 'Wednesday', 'Thursday', 'Friday', 'Sa
turday'];
            const selectedDayName =
 days[selectedDay];
            const selecte
dTimeDisplay = formatTimeDisplay(selectedTime
);
            const utcDayName = days[utcSch
edule.day_of_week_utc];
            const utc
TimeDisplay = formatTimeDisplay(utcSchedule.t
ime_utc);
            
            const isDi
fferent = (selectedDay !== utcSchedule.day_of
_week_utc || selectedTime !== utcSchedule.tim
e_utc);
            const conversionInfo = is
Different 
                ? `<div style="bac
kground: #eff6ff; padding: 12px; border-radiu
s: 6px; margin: 10px 0; border-left: 3px soli
d #3b82f6;">
                    <strong>📅
 Your Schedule:</strong> ${selectedDayName} a
t ${selectedTimeDisplay} (${selectedTimezone}
)<br>
                    <strong>⏰ Runs at
:</strong> ${utcDayName} at ${utcTimeDisplay}
 UTC
                   </div>`
             
   : `<p>Scheduled for: <strong>${selectedDay
Name} at ${selectedTimeDisplay} UTC</strong><
/p>`;
            
            // If scheduli
ng is enabled AND a new secret was provided, 
set up the app registration
            if (s
chedule.enabled && schedule.client_id && newS
ecret) {
                saveBtn.innerHTML = 
'<i class="fas fa-spinner fa-spin"></i> Setti
ng up permissions...';
                
     
           const setupResult = await setupApp
Registration(schedule.client_id);
           
     
                if (setupResult.success
) {
                    let bodyHtml = '<p>Yo
ur auto-update schedule has been saved.</p>';

                    bodyHtml += conversionIn
fo;
                    if (setupResult.permi
ssionsAdded) {
                        bodyHt
ml += '<p>✅ Dynamics CRM permission added t
o your app registration.</p>';
              
      }
                    if (setupResult.a
ppUserCreated) {
                        body
Html += '<p>✅ Application user created in D
ataverse.</p>';
                    }
       
             bodyHtml += '<p>Updates will run
 automatically at the scheduled time.</p>';
 
                   
                    showM
odal({
                        title: 'Schedu
le Saved',
                        body: body
Html,
                        type: 'success'
,
                        confirmOnly: true
 
                   });
                } else
 {
                    showModal({
          
              title: 'Schedule Saved (Manual 
Setup Needed)',
                        body:
 `<p>Your schedule has been saved.</p>${conve
rsionInfo}
                            <p><st
rong>⚠️ Automatic setup failed: ${setupRe
sult.error}</strong></p>
                    
        <p>Please manually:</p>
             
               <ol>
                         
       <li>Add "Dynamics CRM → user_imperso
nation" permission to your app</li>
         
                       <li>Grant admin consen
t</li>
                                <li>Cr
eate an Application User in Power Platform Ad
min Center</li>
                            <
/ol>`,
                        type: 'warning
',
                        confirmOnly: true

                    });
                }
   
         } else {
                showModal({

                    title: 'Schedule Saved',

                    body: `<p>Your auto-upda
te schedule has been saved.</p>${conversionIn
fo}<p>Updates will run automatically at the s
cheduled time.</p>`,
                    type
: 'success',
                    confirmOnly:
 true
                });
            }
     
   } else {
            const errorText = awa
it resp.text();
            console.error('Fa
iled to save schedule:', errorText);
        
    showError('Failed to save schedule. Pleas
e try again.');
        }
    } catch (error)
 {
        console.error('Schedule save error
:', error);
        showError('Failed to save
 schedule: ' + error.message);
    } finally 
{
        saveBtn.disabled = false;
        s
aveBtn.innerHTML = originalText;
    }
}

// 
── Automatic App Registration Setup ─�
�──────────────�
�──────────────�
�─
// Dynamics CRM API Resource ID (constan
t across all tenants)
const DYNAMICS_CRM_RESO
URCE_ID = '00000007-0000-0000-c000-0000000000
00';
// user_impersonation scope ID for Dynam
ics CRM
const USER_IMPERSONATION_SCOPE_ID = '
78ce3f0f-a1ce-49c2-8cde-64b5c0896db4';

async
 function setupAppRegistration(clientId) {
  
  // This function configures the user's app 
registration with required permissions
    //
 and creates the Application User in Datavers
e
    
    if (!msalInstance) {
        conso
le.warn('MSAL not initialized');
        retu
rn { success: false, error: 'Not authenticate
d' };
    }
    
    const accounts = msalIns
tance.getAllAccounts();
    if (accounts.leng
th === 0) {
        return { success: false, 
error: 'No authenticated account' };
    }
  
  
    try {
        // Get Graph API token
 
       const graphToken = await msalInstance.
acquireTokenSilent({
            scopes: ['ht
tps://graph.microsoft.com/Application.ReadWri
te.All'],
            account: accounts[0]
  
      }).catch(async () => {
            // F
allback to popup if silent fails
            
return await msalInstance.acquireTokenPopup({

                scopes: ['https://graph.micr
osoft.com/Application.ReadWrite.All'],
      
          account: accounts[0]
            })
;
        });
        
        if (!graphToke
n || !graphToken.accessToken) {
            r
eturn { success: false, error: 'Could not get
 Graph API permission. Please grant admin con
sent.' };
        }
        
        console.
log('Got Graph API token, configuring app reg
istration...');
        
        // Step 1: F
ind the app registration by client ID
       
 const appResp = await fetch(
            `ht
tps://graph.microsoft.com/v1.0/applications?$
filter=appId eq '${clientId}'`,
            {

                headers: {
                 
   'Authorization': `Bearer ${graphToken.acce
ssToken}`
                }
            }
   
     );
        
        if (!appResp.ok) {
 
           const errText = await appResp.text
();
            console.error('Failed to find
 app registration:', errText);
            re
turn { success: false, error: 'Could not find
 app registration. Make sure you have Applica
tion.ReadWrite.All permission.' };
        }

        
        const appData = await appRes
p.json();
        if (!appData.value || appDa
ta.value.length === 0) {
            return {
 success: false, error: `App registration wit
h ID ${clientId} not found in your tenant.` }
;
        }
        
        const app = appD
ata.value[0];
        const appObjectId = app
.id;
        console.log('Found app registrat
ion:', app.displayName, 'Object ID:', appObje
ctId);
        
        // Step 2: Check if D
ynamics CRM permission already exists
       
 const existingPermissions = app.requiredReso
urceAccess || [];
        const hasDynamicsCR
M = existingPermissions.some(
            ra 
=> ra.resourceAppId === DYNAMICS_CRM_RESOURCE
_ID &&
                  ra.resourceAccess.so
me(a => a.id === USER_IMPERSONATION_SCOPE_ID)

        );
        
        if (!hasDynamics
CRM) {
            console.log('Adding Dynami
cs CRM user_impersonation permission...');
  
          
            // Add Dynamics CRM pe
rmission
            const newPermissions = [
...existingPermissions];
            const dy
namicsCrmEntry = newPermissions.find(ra => ra
.resourceAppId === DYNAMICS_CRM_RESOURCE_ID);

            
            if (dynamicsCrmEntr
y) {
                // Add scope to existing
 entry
                if (!dynamicsCrmEntry.
resourceAccess.some(a => a.id === USER_IMPERS
ONATION_SCOPE_ID)) {
                    dyna
micsCrmEntry.resourceAccess.push({
          
              id: USER_IMPERSONATION_SCOPE_ID
,
                        type: 'Scope'
     
               });
                }
        
    } else {
                // Add new entry
 for Dynamics CRM
                newPermissi
ons.push({
                    resourceAppId:
 DYNAMICS_CRM_RESOURCE_ID,
                  
  resourceAccess: [{
                        
id: USER_IMPERSONATION_SCOPE_ID,
            
            type: 'Scope'
                   
 }]
                });
            }
       
     
            // Update the app registrat
ion
            const updateResp = await fetc
h(
                `https://graph.microsoft.c
om/v1.0/applications/${appObjectId}`,
       
         {
                    method: 'PATCH
',
                    headers: {
           
             'Authorization': `Bearer ${graph
Token.accessToken}`,
                        
'Content-Type': 'application/json'
          
          },
                    body: JSON.s
tringify({
                        requiredRe
sourceAccess: newPermissions
                
    })
                }
            );
     
       
            if (!updateResp.ok) {
   
             const errText = await updateResp
.text();
                console.error('Faile
d to update app permissions:', errText);
    
            return { success: false, error: '
Could not add Dynamics CRM permission. You ma
y need to add it manually.' };
            }

            
            console.log('✅ Dyn
amics CRM permission added to app registratio
n');
        } else {
            console.log
('✅ Dynamics CRM permission already exists'
);
        }
        
        // Step 3: Gran
t admin consent (requires Directory.ReadWrite
.All or admin privileges)
        // This cre
ates a service principal if it doesn't exist 
and grants consent
        try {
            
await grantAdminConsent(graphToken.accessToke
n, clientId);
        } catch (consentError) 
{
            console.warn('Admin consent may
 need to be granted manually:', consentError.
message);
        }
        
        // Step 
4: Create Application User in Dataverse
     
   const appUserResult = await createApplicat
ionUser(clientId);
        
        return { 

            success: true, 
            perm
issionsAdded: !hasDynamicsCRM,
            ap
pUserCreated: appUserResult.success,
        
    appUserMessage: appUserResult.message
   
     };
        
    } catch (error) {
      
  console.error('Setup error:', error);
     
   return { success: false, error: error.mess
age };
    }
}

async function grantAdminCons
ent(graphToken, clientId) {
    // First ensu
re the service principal exists
    let spRes
p = await fetch(
        `https://graph.micro
soft.com/v1.0/servicePrincipals?$filter=appId
 eq '${clientId}'`,
        {
            hea
ders: { 'Authorization': `Bearer ${graphToken
}` }
        }
    );
    
    let spData = a
wait spResp.json();
    let servicePrincipalI
d;
    
    if (!spData.value || spData.value
.length === 0) {
        // Create service pr
incipal
        const createSpResp = await fe
tch(
            'https://graph.microsoft.com
/v1.0/servicePrincipals',
            {
     
           method: 'POST',
                he
aders: {
                    'Authorization':
 `Bearer ${graphToken}`,
                    
'Content-Type': 'application/json'
          
      },
                body: JSON.stringify
({ appId: clientId })
            }
        )
;
        
        if (createSpResp.ok) {
   
         const newSp = await createSpResp.jso
n();
            servicePrincipalId = newSp.i
d;
            console.log('✅ Service princ
ipal created:', servicePrincipalId);
        
} else {
            throw new Error('Could n
ot create service principal');
        }
    
} else {
        servicePrincipalId = spData.
value[0].id;
    }
    
    // Get Dynamics C
RM service principal
    const crmSpResp = aw
ait fetch(
        `https://graph.microsoft.c
om/v1.0/servicePrincipals?$filter=appId eq '$
{DYNAMICS_CRM_RESOURCE_ID}'`,
        {
     
       headers: { 'Authorization': `Bearer ${
graphToken}` }
        }
    );
    
    cons
t crmSpData = await crmSpResp.json();
    if 
(!crmSpData.value || crmSpData.value.length =
== 0) {
        console.warn('Dynamics CRM se
rvice principal not found - consent may need 
to be granted manually');
        return;
   
 }
    
    const crmServicePrincipalId = crm
SpData.value[0].id;
    
    // Grant oauth2P
ermissionGrant (delegated permission consent)

    const grantResp = await fetch(
        '
https://graph.microsoft.com/v1.0/oauth2Permis
sionGrants',
        {
            method: 'P
OST',
            headers: {
                
'Authorization': `Bearer ${graphToken}`,
    
            'Content-Type': 'application/json
'
            },
            body: JSON.strin
gify({
                clientId: servicePrinc
ipalId,
                consentType: 'AllPrin
cipals',
                resourceId: crmServi
cePrincipalId,
                scope: 'user_i
mpersonation'
            })
        }
    );

    
    if (grantResp.ok || grantResp.statu
s === 409) { // 409 = already exists
        
console.log('✅ Admin consent granted for Dy
namics CRM');
    } else {
        const errT
ext = await grantResp.text();
        console
.warn('Could not grant admin consent:', errTe
xt);
    }
}

async function createApplicatio
nUser(clientId) {
    if (!currentOrgUrl || !
msalInstance) {
        return { success: fal
se, message: 'Not connected to environment' }
;
    }
    
    try {
        const accounts
 = msalInstance.getAllAccounts();
        if 
(accounts.length === 0) {
            return 
{ success: false, message: 'No authenticated 
account' };
        }
        
        const 
tokenResponse = await msalInstance.acquireTok
enSilent({
            scopes: [`${currentOrg
Url}/.default`],
            account: account
s[0]
        });
        
        const acces
sToken = tokenResponse.accessToken;
        

        // Check if application user already 
exists
        const checkUrl = `${currentOrg
Url}/api/data/v9.2/systemusers?$filter=applic
ationid eq ${clientId}&$select=systemuserid,f
ullname`;
        const checkResp = await fet
ch(checkUrl, {
            headers: {
       
         'Authorization': `Bearer ${accessTok
en}`,
                'OData-MaxVersion': '4.
0',
                'OData-Version': '4.0',
 
               'Accept': 'application/json'
 
           }
        });
        
        if 
(checkResp.ok) {
            const checkData 
= await checkResp.json();
            if (che
ckData.value && checkData.value.length > 0) {

                console.log('✅ Application
 user already exists:', checkData.value[0].fu
llname);
                return { success: tr
ue, message: 'Application user already exists
' };
            }
        }
        
       
 // Get root business unit
        const buRe
sp = await fetch(
            `${currentOrgUr
l}/api/data/v9.2/businessunits?$filter=parent
businessunitid eq null&$select=businessunitid
`,
            {
                headers: {
 
                   'Authorization': `Bearer $
{accessToken}`,
                    'OData-Ma
xVersion': '4.0',
                    'OData-
Version': '4.0',
                    'Accept'
: 'application/json'
                }
      
      }
        );
        
        if (!buRe
sp.ok) {
            return { success: false,
 message: 'Could not get business unit' };
  
      }
        
        const buData = await
 buResp.json();
        if (!buData.value || 
buData.value.length === 0) {
            retu
rn { success: false, message: 'No root busine
ss unit found' };
        }
        
        
const businessUnitId = buData.value[0].busine
ssunitid;
        
        // Create the appl
ication user
        const createResp = await
 fetch(`${currentOrgUrl}/api/data/v9.2/system
users`, {
            method: 'POST',
       
     headers: {
                'Authorizatio
n': `Bearer ${accessToken}`,
                
'OData-MaxVersion': '4.0',
                'O
Data-Version': '4.0',
                'Accept
': 'application/json',
                'Conte
nt-Type': 'application/json'
            },
 
           body: JSON.stringify({
           
     'applicationid': clientId,
             
   'fullname': 'D365 App Updater Scheduler',

                'internalemailaddress': `app-
updater-${clientId.substring(0, 8)}@automatio
n.local`,
                'businessunitid@oda
ta.bind': `/businessunits(${businessUnitId})`
,
                'accessmode': 4 // Non-inte
ractive (application user)
            })
   
     });
        
        if (createResp.ok |
| createResp.status === 204) {
            co
nsole.log('✅ Application user created succe
ssfully');
            
            // Get th
e created user ID and assign System Administr
ator role
            const userUrl = createR
esp.headers.get('OData-EntityId');
          
  if (userUrl) {
                const userId
Match = userUrl.match(/systemusers\(([^)]+)\)
/);
                if (userIdMatch) {
      
              const userId = userIdMatch[1];

                    await assignSystemAdminRo
le(accessToken, userId);
                }
  
          }
            
            return {
 success: true, message: 'Application user cr
eated with System Administrator role' };
    
    } else {
            const errorText = aw
ait createResp.text();
            console.er
ror('Could not create application user:', cre
ateResp.status, errorText);
            retur
n { success: false, message: `Could not creat
e application user: ${errorText}` };
        
}
        
    } catch (error) {
        cons
ole.error('Error creating application user:',
 error);
        return { success: false, mes
sage: error.message };
    }
}

async functio
n assignSystemAdminRole(accessToken, userId) 
{
    try {
        // Get the System Adminis
trator role ID
        const roleResp = await
 fetch(
            `${currentOrgUrl}/api/dat
a/v9.2/roles?$filter=name eq 'System Administ
rator'&$select=roleid`,
            {
       
         headers: {
                    'Auth
orization': `Bearer ${accessToken}`,
        
            'OData-MaxVersion': '4.0',
      
              'OData-Version': '4.0',
       
             'Accept': 'application/json'
   
             }
            }
        );
     
   
        if (!roleResp.ok) {
            c
onsole.warn('Could not get System Administrat
or role');
            return;
        }
    
    
        const roleData = await roleResp.
json();
        if (!roleData.value || roleDa
ta.value.length === 0) {
            console.
warn('System Administrator role not found');

            return;
        }
        
      
  const roleId = roleData.value[0].roleid;
  
      
        // Associate the role with the
 user
        const associateResp = await fet
ch(
            `${currentOrgUrl}/api/data/v9
.2/systemusers(${userId})/systemuserroles_ass
ociation/$ref`,
            {
               
 method: 'POST',
                headers: {
 
                   'Authorization': `Bearer $
{accessToken}`,
                    'OData-Ma
xVersion': '4.0',
                    'OData-
Version': '4.0',
                    'Content
-Type': 'application/json'
                },

                body: JSON.stringify({
     
               '@odata.id': `${currentOrgUrl}
/api/data/v9.2/roles(${roleId})`
            
    })
            }
        );
        
    
    if (associateResp.ok || associateResp.sta
tus === 204) {
            console.log('✅ S
ystem Administrator role assigned');
        
} else {
            console.warn('Could not 
assign role:', associateResp.status);
       
 }
    } catch (error) {
        console.warn
('Error assigning role:', error.message);
   
 }
}

// Securely save client secret to a sep
arate protected table
// The anon key can INS
ERT/UPDATE but CANNOT read from this table
as
ync function saveSecretSecurely(cfg, schedule
Id, secret) {
    try {
        // Try to ups
ert the secret
        // First check if a ro
w exists (we can't read it, but we can try to
 update)
        const upsertResp = await fet
ch(
            `${cfg.url}/rest/v1/schedule_
secrets`,
            {
                metho
d: 'POST',
                headers: {
       
             'apikey': cfg.key,
             
       'Authorization': `Bearer ${cfg.key}`,

                    'Content-Type': 'applicat
ion/json',
                    'Prefer': 'res
olution=merge-duplicates'  // Upsert: update 
if exists
                },
                
body: JSON.stringify({
                    sc
hedule_id: scheduleId,
                    cl
ient_secret: secret,
                    upda
ted_at: new Date().toISOString()
            
    })
            }
        );
        
    
    if (upsertResp.ok) {
            console.
log('✅ Secret saved securely');
        } e
lse {
            // If upsert fails, try upd
ate
            const updateResp = await fetc
h(
                `${cfg.url}/rest/v1/schedu
le_secrets?schedule_id=eq.${scheduleId}`,
   
             {
                    method: 'P
ATCH',
                    headers: {
       
                 'apikey': cfg.key,
         
               'Authorization': `Bearer ${cfg
.key}`,
                        'Content-Type
': 'application/json'
                    },

                    body: JSON.stringify({
  
                      client_secret: secret,

                        updated_at: new Date(
).toISOString()
                    })
      
          }
            );
            
     
       if (updateResp.ok) {
                c
onsole.log('✅ Secret updated securely');
  
          } else {
                console.wa
rn('Could not save secret securely:', await u
pdateResp.text());
            }
        }
  
  } catch (error) {
        console.warn('Err
or saving secret:', error.message);
    }
}


// Check if a secret exists for a schedule (w
ithout being able to read it)
async function 
hasSecretStored(cfg, scheduleId) {
    try {

        // We can't read the secret, but we c
an check if rows exist by trying a count
    
    // Actually, with RLS blocking SELECT, we
 can't even count
        // Instead, we trac
k this in the main schedule table
        ret
urn true; // Assume exists if schedule is alr
eady enabled
    } catch (error) {
        re
turn false;
    }
}

async function disableSc
hedule() {
    const cfg = getSupabaseConfig(
);
    if (!cfg) return;
    
    const userE
mail = getCurrentUserEmail();
    const envId
 = environmentId || '';
    
    if (!userEma
il || !envId) return;
    
    try {
        
await fetch(
            `${cfg.url}/rest/v1/
update_schedules?user_email=eq.${encodeURICom
ponent(userEmail)}&environment_id=eq.${encode
URIComponent(envId)}`,
            {
        
        method: 'PATCH',
                head
ers: {
                    'apikey': cfg.key,

                    'Authorization': `Bearer
 ${cfg.key}`,
                    'Content-Ty
pe': 'application/json'
                },
  
              body: JSON.stringify({ enabled:
 false, updated_at: new Date().toISOString() 
})
            }
        );
        updateSch
eduleStatus(null);
    } catch (error) {
    
    console.warn('Failed to disable schedule:
', error.message);
    }
}

// UI Helpers
fun
ction showLoading(message, details) {
    con
st overlay = document.getElementById('loading
Overlay');
    document.getElementById('loadi
ngMessage').textContent = message;
    docume
nt.getElementById('loadingDetails').textConte
nt = details || '';
    overlay.classList.rem
ove('hidden');
    overlay.style.display = 'f
lex';
}

function hideLoading() {
    const o
verlay = document.getElementById('loadingOver
lay');
    overlay.classList.add('hidden');
 
   overlay.style.display = 'none';
}

functio
n showError(message) {
    hideLoading();
   
 // Use 'body' for HTML content, 'message' fo
r plain text
    if (message.includes('<') &&
 message.includes('>')) {
        showModal({
 title: 'Error', body: message, type: 'danger
', confirmOnly: true });
    } else {
       
 showModal({ title: 'Error', message: message
, type: 'danger', confirmOnly: true });
    }

}

function escapeHtml(text) {
    const div
 = document.createElement('div');
    div.tex
tContent = text || '';
    return div.innerHT
ML;
}

// Parse API error messages into user-
friendly summary + raw detail
function parseE
rrorMessage(errorStr) {
    if (!errorStr) re
turn { summary: 'Unknown error', detail: '' }
;
    
    // Extract HTTP status code if pre
sent (e.g. "401 - {...}")
    const statusMat
ch = errorStr.match(/^(\d{3})\s*-\s*(.*)/s);

    const statusCode = statusMatch ? parseInt
(statusMatch[1]) : null;
    const body = sta
tusMatch ? statusMatch[2].trim() : errorStr;

    
    // Try to parse JSON error body
    
let parsed = null;
    try {
        parsed =
 JSON.parse(body);
    } catch (e) {
        
// Not JSON
    }
    
    // Build user-frie
ndly summary based on status code
    let sum
mary = '';
    if (statusCode === 401 || (par
sed && parsed.code === 'AuthorizationHeaderIn
valid')) {
        summary = 'Authorization f
ailed — your session may have expired. Try 
signing out and back in, or check that the ap
p registration has the required API permissio
ns.';
    } else if (statusCode === 403) {
  
      summary = 'Access denied — you don\'t
 have permission to update this app. Check yo
ur Power Platform admin role.';
    } else if
 (statusCode === 404) {
        summary = 'Pa
ckage not found — the update package may no
 longer be available in the catalog.';
    } 
else if (statusCode === 429) {
        summar
y = 'Too many requests — the API is rate-li
miting. Please wait a few minutes and try aga
in.';
    } else if (statusCode >= 500) {
   
     summary = 'Server error (' + statusCode 
+ ') — the Power Platform service encounter
ed an issue. Try again later.';
    } else if
 (parsed && parsed.message) {
        // Trun
cate the parsed message to first sentence
   
     const msg = parsed.message;
        cons
t firstSentence = msg.split(/\.\s/)[0];
     
   summary = firstSentence.length < msg.lengt
h ? firstSentence + '.' : msg;
    } else {
 
       summary = statusCode ? 'Error ' + stat
usCode + ' — update request failed.' : erro
rStr.substring(0, 120);
    }
    
    // Raw
 detail is the full original error for those 
who want it
    const detail = errorStr.lengt
h > summary.length + 20 ? errorStr : '';
    

    return { summary, detail };
}

// ──
 Custom Modal System ────────
───────────────
───────────────
────────
let _modalResolve = 
null;

/**
 * Show a custom modal dialog. Ret
urns a Promise<boolean>.
 * Options:
 *   tit
le     – Modal title
 *   message   – Tex
t or HTML message
 *   body      – Full HTM
L body (overrides message)
 *   type      –
 'info' | 'warning' | 'success' | 'danger' | 
'update'
 *   icon      – FontAwesome icon 
class (auto-selected from type if omitted)
 *
   okText    – OK button text (default "OK"
)
 *   cancelText– Cancel button text (defa
ult "Cancel")
 *   okClass   – Extra class 
for OK button (e.g. 'btn-success-modal')
 *  
 confirmOnly – If true, hide Cancel button 
(alert-style)
 */
function showModal(opts) {

    return new Promise(resolve => {
        _
modalResolve = resolve;
        
        cons
t overlay = document.getElementById('customMo
dal');
        const iconWrap = document.getE
lementById('modalIconWrap');
        const ic
on = document.getElementById('modalIcon');
  
      const title = document.getElementById('
modalTitle');
        const body = document.g
etElementById('modalBody');
        const okB
tn = document.getElementById('modalOkBtn');
 
       const cancelBtn = document.getElementB
yId('modalCancelBtn');
        
        const
 typeIcons = {
            info: 'fas fa-info
-circle',
            warning: 'fas fa-exclam
ation-triangle',
            success: 'fas fa
-check-circle',
            danger: 'fas fa-t
imes-circle',
            update: 'fas fa-arr
ow-circle-up'
        };
        
        con
st t = opts.type || 'info';
        iconWrap.
className = 'modal-icon-wrap icon-' + t;
    
    icon.className = opts.icon || typeIcons[t
] || typeIcons.info;
        title.textConten
t = opts.title || 'Notice';
        
        
if (opts.body) {
            body.innerHTML =
 opts.body;
        } else {
            body
.innerHTML = '<p class="mb-0">' + escapeHtml(
opts.message || '') + '</p>';
        }
     
   
        okBtn.textContent = opts.okText |
| 'OK';
        okBtn.className = 'btn btn-mo
dal-ok' + (opts.okClass ? ' ' + opts.okClass 
: '');
        cancelBtn.textContent = opts.c
ancelText || 'Cancel';
        cancelBtn.styl
e.display = opts.confirmOnly ? 'none' : '';
 
       
        overlay.style.display = 'flex
';
    });
}

function closeModal(result) {
 
   document.getElementById('customModal').sty
le.display = 'none';
    if (_modalResolve) {

        _modalResolve(result);
        _moda
lResolve = null;
    }
}

/**
 * Helper: show
 a confirm modal for updating apps.
 * @param
 {Array} appsToUpdate – array of {name, ver
sion, latestVersion}
 * @returns Promise<bool
ean>
 */
function showUpdateConfirm(appsToUpd
ate) {
    let listHtml = '<ul class="update-
list">';
    for (const app of appsToUpdate) 
{
        listHtml += '<li>';
        listHtm
l += '<span class="app-label" title="' + esca
peHtml(app.name) + '">' + escapeHtml(app.name
) + '</span>';
        listHtml += '<span cla
ss="version-badge">' + escapeHtml(app.version
) + '<span class="arrow">→</span>' + escape
Html(app.latestVersion || 'latest') + '</span
>';
        listHtml += '</li>';
    }
    li
stHtml += '</ul>';
    
    const bodyHtml = 
'<p class="modal-message">The following ' + a
ppsToUpdate.length + ' app' + (appsToUpdate.l
ength !== 1 ? 's' : '') + ' will be updated:<
/p>' + listHtml;
    
    return showModal({

        title: 'Update Apps',
        body: b
odyHtml,
        type: 'update',
        okTe
xt: 'Update All',
        okClass: 'btn-succe
ss-modal',
        cancelText: 'Cancel'
    }
);
}

/**
 * Helper: show a simple alert moda
l (no Cancel button).
 */
function showAlert(
title, message, type) {
    return showModal(
{ title, message, type: type || 'info', confi
rmOnly: true });
}

window.updateSingleApp = 
updateSingleApp;
window.installApp = installA
pp;
window.closeModal = closeModal;

// â•
â•â• SSO AUTO-CONNECT FROM COMPAN
ION APPS (e.g. D365 DataGen) â•â•â
•
// Shared multi-tenant Entra app regist
ration. Users still see the same
// one-time 
consent prompt on first use in their tenant; 
after that, no
// app registration or client 
ID is ever needed.
const SHARED_CLIENT_ID = '
c4ff0dc1-4cf0-44e1-8a26-7d265772484a';

funct
ion trySsoAutoConnect() {
    if (msalInstanc
e) { logInfo('SSO skip: already authenticated
'); return; }
    const params = new URLSearc
hParams(window.location.search);
    const pa
ramOrgUrl    = params.get('orgUrl');
    cons
t paramClientId  = params.get('clientId') || 
SHARED_CLIENT_ID;
    const paramLoginHint = 
params.get('loginHint');
    const paramAutoC
onnect = params.get('autoConnect') === '1';


    if (!paramAutoConnect || !paramOrgUrl) re
turn;

    logInfo('SSO auto-connect detected
', {
        orgUrl: paramOrgUrl,
        cli
entId: paramClientId.substring(0,8) + '...',

        hasLoginHint: !!paramLoginHint
    })
;

    const orgUrlEl   = document.getElement
ById('orgUrl');
    const clientIdEl = docume
nt.getElementById('clientId');
    if (orgUrl
El)   orgUrlEl.value   = paramOrgUrl;
    if 
(clientIdEl) clientIdEl.value = paramClientId
;
    if (paramLoginHint) {
        sessionSt
orage.setItem('d365_login_hint', paramLoginHi
nt);
    }
    // Strip URL params so a reloa
d doesn't loop
    const cleanUrl = window.lo
cation.origin + window.location.pathname;
   
 window.history.replaceState({}, '', cleanUrl
);

    // Hide the form chrome so the user s
ees a clean "Authenticating..."
    const aut
hCard = document.querySelector('#authForm')?.
closest('.card');
    if (authCard) authCard.
style.display = 'none';

    showLoading('Sig
ning you in...', 'Connecting to Microsoft via
 SSO');

    // Submit the form programmatica
lly â€” handleAuthentication will pick u
p loginHint from sessionStorage
    const aut
hForm = document.getElementById('authForm');

    if (authForm) {
        logInfo('Auto-sub
mitting auth form');
        authForm.request
Submit ? authForm.requestSubmit() :
         
   authForm.dispatchEvent(new Event('submit',
 { cancelable: true, bubbles: true }));
    }

}
// â•â•â• END SSO BLOCK â�
��â•â•


