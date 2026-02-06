// ═══════════════════════════════════════════════════════════════
// Admin Dashboard — D365 App Updater
// Connects to Supabase to display usage analytics
// ═══════════════════════════════════════════════════════════════

const DEFAULT_SUPABASE_URL = 'https://fpekzltxukikaixebeeu.supabase.co';
const DEFAULT_SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZwZWt6bHR4dWtpa2FpeGViZWV1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzA0MDU0ODEsImV4cCI6MjA4NTk4MTQ4MX0.uH4JgKbf_-Al_iArzEy6UZ3edJNzFSCBVlMNI04li0Y';
let supabaseUrl = DEFAULT_SUPABASE_URL;
let supabaseKey = DEFAULT_SUPABASE_KEY;
let allRecords = [];
let filteredRecords = [];
let currentPage = 0;
let sortCol = 'timestamp';
let sortAsc = false;
let chartTimeline = null;
let chartSuccessFail = null;
let chartEnvs = null;
let chartUsers = null;

// ── Init ──────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function () {
    // Pre-fill config fields with hardcoded defaults
    document.getElementById('cfgUrl').value = supabaseUrl;
    document.getElementById('cfgKey').value = supabaseKey;
    // Auto-load data immediately since config is hardcoded
    loadData();
});

function saveConfig() {
    supabaseUrl = document.getElementById('cfgUrl').value.trim().replace(/\/+$/, '');
    supabaseKey = document.getElementById('cfgKey').value.trim();
    if (!supabaseUrl || !supabaseKey) return;
    localStorage.setItem('d365_admin_supabase', JSON.stringify({ url: supabaseUrl, key: supabaseKey }));
    document.getElementById('configSaved').style.display = 'inline';
    setTimeout(() => document.getElementById('configSaved').style.display = 'none', 2000);
    loadData();
}

// ── Data fetch ────────────────────────────────────────────────
async function loadData() {
    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('adminLoading').style.display = '';
    document.getElementById('mainContent').style.display = 'none';

    try {
        // Fetch all records ordered by timestamp desc
        const url = `${supabaseUrl}/rest/v1/usage_logs?select=*&order=timestamp.desc&limit=5000`;
        const resp = await fetch(url, {
            headers: {
                'apikey': supabaseKey,
                'Authorization': `Bearer ${supabaseKey}`,
                'Content-Type': 'application/json'
            }
        });

        if (!resp.ok) {
            const err = await resp.text();
            throw new Error(`Supabase error ${resp.status}: ${err}`);
        }

        allRecords = await resp.json();
        filteredRecords = [...allRecords];

        document.getElementById('adminLoading').style.display = 'none';
        document.getElementById('mainContent').style.display = '';

        renderStats();
        renderCharts();
        renderRecentActivity();
        applyFilters();

    } catch (error) {
        document.getElementById('adminLoading').style.display = 'none';
        document.getElementById('emptyState').style.display = '';
        document.getElementById('emptyState').innerHTML =
            '<i class="fas fa-exclamation-triangle text-danger"></i>' +
            '<h4>Connection Failed</h4>' +
            '<p>' + escapeHtml(error.message) + '</p>' +
            '<p class="text-muted" style="font-size:0.85rem;">Check your Supabase URL and key, then click Connect again.</p>';
        console.error('Failed to load data:', error);
    }
}

// ── Stats ─────────────────────────────────────────────────────
function renderStats() {
    const totalSessions = allRecords.length;
    const uniqueUsers = new Set(allRecords.map(r => (r.user_email || '').toLowerCase()).filter(Boolean)).size;
    const uniqueEnvs = new Set(allRecords.map(r => (r.org_url || '').toLowerCase()).filter(Boolean)).size;
    const totalSuccess = allRecords.reduce((s, r) => s + (r.success_count || 0), 0);
    const totalFailed = allRecords.reduce((s, r) => s + (r.fail_count || 0), 0);

    document.getElementById('statSessions').textContent = totalSessions.toLocaleString();
    document.getElementById('statUsers').textContent = uniqueUsers.toLocaleString();
    document.getElementById('statEnvs').textContent = uniqueEnvs.toLocaleString();
    document.getElementById('statSuccess').textContent = totalSuccess.toLocaleString();
    document.getElementById('statFailed').textContent = totalFailed.toLocaleString();
}

// ── Charts ────────────────────────────────────────────────────
function renderCharts() {
    renderTimelineChart();
    renderSuccessFailChart();
    renderEnvsChart();
    renderUsersChart();
}

function renderTimelineChart() {
    // Group by date
    const byDate = {};
    allRecords.forEach(r => {
        const d = (r.timestamp || '').substring(0, 10);
        if (!d) return;
        if (!byDate[d]) byDate[d] = { success: 0, failed: 0 };
        byDate[d].success += r.success_count || 0;
        byDate[d].failed += r.fail_count || 0;
    });

    const dates = Object.keys(byDate).sort();
    const last30 = dates.slice(-30);

    if (chartTimeline) chartTimeline.destroy();
    chartTimeline = new Chart(document.getElementById('chartTimeline'), {
        type: 'bar',
        data: {
            labels: last30.map(d => formatDateShort(d)),
            datasets: [
                {
                    label: 'Succeeded',
                    data: last30.map(d => byDate[d].success),
                    backgroundColor: '#28a74580',
                    borderColor: '#28a745',
                    borderWidth: 1,
                    borderRadius: 4,
                },
                {
                    label: 'Failed',
                    data: last30.map(d => byDate[d].failed),
                    backgroundColor: '#dc354580',
                    borderColor: '#dc3545',
                    borderWidth: 1,
                    borderRadius: 4,
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { intersect: false, mode: 'index' },
            scales: {
                x: { stacked: true, grid: { display: false } },
                y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } }
            },
            plugins: { legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle' } } }
        }
    });
}

function renderSuccessFailChart() {
    const totalSuccess = allRecords.reduce((s, r) => s + (r.success_count || 0), 0);
    const totalFailed = allRecords.reduce((s, r) => s + (r.fail_count || 0), 0);

    if (chartSuccessFail) chartSuccessFail.destroy();
    chartSuccessFail = new Chart(document.getElementById('chartSuccessFail'), {
        type: 'doughnut',
        data: {
            labels: ['Succeeded', 'Failed'],
            datasets: [{
                data: [totalSuccess, totalFailed],
                backgroundColor: ['#28a745', '#dc3545'],
                borderWidth: 0,
                spacing: 2,
                borderRadius: 4,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '65%',
            plugins: {
                legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', padding: 16 } }
            }
        }
    });
}

function renderEnvsChart() {
    const envMap = {};
    allRecords.forEach(r => {
        const env = r.org_url || 'Unknown';
        envMap[env] = (envMap[env] || 0) + (r.success_count || 0) + (r.fail_count || 0);
    });

    const sorted = Object.entries(envMap).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const colors = ['#0078d4', '#00a4ef', '#7fba00', '#f25022', '#ffb900', '#737373', '#b4009e', '#00b7c3'];

    if (chartEnvs) chartEnvs.destroy();
    chartEnvs = new Chart(document.getElementById('chartEnvs'), {
        type: 'bar',
        data: {
            labels: sorted.map(e => shortenUrl(e[0])),
            datasets: [{
                label: 'Total Updates',
                data: sorted.map(e => e[1]),
                backgroundColor: colors.slice(0, sorted.length),
                borderRadius: 6,
                borderSkipped: false,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            scales: { x: { beginAtZero: true, ticks: { precision: 0 } }, y: { grid: { display: false } } },
            plugins: { legend: { display: false } }
        }
    });
}

function renderUsersChart() {
    const userMap = {};
    allRecords.forEach(r => {
        const user = r.user_email || 'Unknown';
        userMap[user] = (userMap[user] || 0) + (r.success_count || 0) + (r.fail_count || 0);
    });

    const sorted = Object.entries(userMap).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const colors = ['#7c3aed', '#a855f7', '#c084fc', '#e9d5ff', '#0078d4', '#00a4ef', '#f25022', '#ffb900'];

    if (chartUsers) chartUsers.destroy();
    chartUsers = new Chart(document.getElementById('chartUsers'), {
        type: 'bar',
        data: {
            labels: sorted.map(e => e[0]),
            datasets: [{
                label: 'Total Updates',
                data: sorted.map(e => e[1]),
                backgroundColor: colors.slice(0, sorted.length),
                borderRadius: 6,
                borderSkipped: false,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            scales: { x: { beginAtZero: true, ticks: { precision: 0 } }, y: { grid: { display: false } } },
            plugins: { legend: { display: false } }
        }
    });
}

// ── Recent Activity ───────────────────────────────────────────
function renderRecentActivity() {
    const recent = allRecords.slice(0, 10);
    if (recent.length === 0) {
        document.getElementById('recentActivity').innerHTML = '<p class="text-muted text-center py-3">No activity yet.</p>';
        return;
    }

    let html = '<div class="list-group list-group-flush">';
    for (const r of recent) {
        const timeAgo = getRelativeTime(r.timestamp);
        const total = (r.success_count || 0) + (r.fail_count || 0);
        const failBadge = r.fail_count > 0
            ? ' <span class="badge badge-fail-count">' + r.fail_count + ' failed</span>'
            : '';

        html += '<div class="list-group-item px-0 py-3 border-0 border-bottom">';
        html += '<div class="d-flex align-items-center">';
        html += '<div class="stat-icon ' + (r.fail_count > 0 ? 'red' : 'green') + ' me-3" style="width:36px;height:36px;font-size:0.9rem;">';
        html += '<i class="fas fa-' + (r.fail_count > 0 ? 'exclamation-triangle' : 'check') + '"></i></div>';
        html += '<div class="flex-grow-1">';
        html += '<div style="font-weight:500;font-size:0.88rem;">' + escapeHtml(r.user_email || 'Unknown user') + '</div>';
        html += '<div class="text-muted" style="font-size:0.78rem;">' + escapeHtml(shortenUrl(r.org_url || '')) + '</div>';
        html += '</div>';
        html += '<div class="text-end">';
        html += '<div><span class="badge badge-success-count">' + (r.success_count || 0) + ' updated</span> ' + failBadge + '</div>';
        html += '<div class="text-muted" style="font-size:0.75rem;" title="' + escapeHtml(r.timestamp || '') + '">' + timeAgo + '</div>';
        html += '</div>';
        html += '</div>';
        html += '</div>';
    }
    html += '</div>';
    document.getElementById('recentActivity').innerHTML = html;
}

// ── View toggle ───────────────────────────────────────────────
function switchView(view) {
    document.getElementById('viewDashboard').style.display = view === 'dashboard' ? '' : 'none';
    document.getElementById('viewTable').style.display = view === 'table' ? '' : 'none';
    document.getElementById('btnDashboard').classList.toggle('active', view === 'dashboard');
    document.getElementById('btnTable').classList.toggle('active', view === 'table');
}

// ── Table: Filters, Sort, Pagination ──────────────────────────
function applyFilters() {
    const search = (document.getElementById('filterSearch').value || '').toLowerCase();
    const from = document.getElementById('filterFrom').value;
    const to = document.getElementById('filterTo').value;
    const status = document.getElementById('filterStatus').value;

    filteredRecords = allRecords.filter(r => {
        if (search) {
            const hay = ((r.user_email || '') + ' ' + (r.org_url || '') + ' ' + (r.app_names || '')).toLowerCase();
            if (!hay.includes(search)) return false;
        }
        if (from) {
            const d = (r.timestamp || '').substring(0, 10);
            if (d < from) return false;
        }
        if (to) {
            const d = (r.timestamp || '').substring(0, 10);
            if (d > to) return false;
        }
        if (status === 'success' && !(r.success_count > 0)) return false;
        if (status === 'failed' && !(r.fail_count > 0)) return false;
        return true;
    });

    // Sort
    filteredRecords.sort((a, b) => {
        let va = a[sortCol], vb = b[sortCol];
        if (typeof va === 'string') va = va.toLowerCase();
        if (typeof vb === 'string') vb = vb.toLowerCase();
        if (va == null) va = '';
        if (vb == null) vb = '';
        if (va < vb) return sortAsc ? -1 : 1;
        if (va > vb) return sortAsc ? 1 : -1;
        return 0;
    });

    currentPage = 0;
    renderTable();
}

function sortTable(col) {
    if (sortCol === col) {
        sortAsc = !sortAsc;
    } else {
        sortCol = col;
        sortAsc = col === 'timestamp' ? false : true;
    }

    // Update header icons
    document.querySelectorAll('.table-card thead th').forEach(th => th.classList.remove('sorted'));
    applyFilters();
}

function renderTable() {
    const pageSize = parseInt(document.getElementById('pageSize').value) || 25;
    const start = currentPage * pageSize;
    const page = filteredRecords.slice(start, start + pageSize);

    const tbody = document.getElementById('tableBody');

    if (filteredRecords.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted py-4">No records match your filters.</td></tr>';
        document.getElementById('pageInfo').textContent = '0 of 0';
        document.getElementById('btnPrev').disabled = true;
        document.getElementById('btnNext').disabled = true;
        return;
    }

    let html = '';
    for (const r of page) {
        const ts = formatTimestamp(r.timestamp);
        const total = (r.success_count || 0) + (r.fail_count || 0);
        const appNames = r.app_names || '';
        const truncatedNames = appNames.length > 60 ? appNames.substring(0, 60) + '...' : appNames;

        html += '<tr>';
        html += '<td class="text-nowrap">' + ts + '</td>';
        html += '<td>' + escapeHtml(r.user_email || '—') + '</td>';
        html += '<td title="' + escapeHtml(r.org_url || '') + '">' + escapeHtml(shortenUrl(r.org_url || '—')) + '</td>';
        html += '<td class="text-center"><span class="badge badge-success-count">' + (r.success_count || 0) + '</span></td>';
        html += '<td class="text-center">';
        if (r.fail_count > 0) {
            html += '<span class="badge badge-fail-count">' + r.fail_count + '</span>';
        } else {
            html += '<span class="text-muted">0</span>';
        }
        html += '</td>';
        html += '<td class="text-center">' + total + '</td>';
        html += '<td><span class="text-muted" style="font-size:0.78rem;" title="' + escapeHtml(appNames) + '">' + escapeHtml(truncatedNames || '—') + '</span></td>';
        html += '</tr>';
    }
    tbody.innerHTML = html;

    // Pagination info
    const end = Math.min(start + pageSize, filteredRecords.length);
    document.getElementById('pageInfo').textContent = (start + 1) + '-' + end + ' of ' + filteredRecords.length;
    document.getElementById('btnPrev').disabled = currentPage === 0;
    document.getElementById('btnNext').disabled = end >= filteredRecords.length;
}

function changePage(delta) {
    const pageSize = parseInt(document.getElementById('pageSize').value) || 25;
    const maxPage = Math.ceil(filteredRecords.length / pageSize) - 1;
    currentPage = Math.max(0, Math.min(maxPage, currentPage + delta));
    renderTable();
}

function clearFilters() {
    document.getElementById('filterSearch').value = '';
    document.getElementById('filterFrom').value = '';
    document.getElementById('filterTo').value = '';
    document.getElementById('filterStatus').value = '';
    applyFilters();
}

// ── Export CSV ─────────────────────────────────────────────────
function exportCSV() {
    const headers = ['Timestamp', 'User', 'Organization URL', 'Succeeded', 'Failed', 'Total', 'App Names'];
    const rows = filteredRecords.map(r => [
        r.timestamp || '',
        r.user_email || '',
        r.org_url || '',
        r.success_count || 0,
        r.fail_count || 0,
        (r.success_count || 0) + (r.fail_count || 0),
        (r.app_names || '').replace(/,/g, ';')
    ]);

    let csv = headers.join(',') + '\n';
    for (const row of rows) {
        csv += row.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',') + '\n';
    }

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'usage_log_' + new Date().toISOString().substring(0, 10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
}

// ── Helpers ───────────────────────────────────────────────────
function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function shortenUrl(url) {
    return (url || '').replace(/^https?:\/\//, '').replace(/\/$/, '');
}

function formatTimestamp(ts) {
    if (!ts) return '—';
    try {
        const d = new Date(ts);
        return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
            + ' ' + d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
    } catch (e) {
        return ts.substring(0, 16);
    }
}

function formatDateShort(dateStr) {
    try {
        const d = new Date(dateStr + 'T00:00:00');
        return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
    } catch (e) {
        return dateStr;
    }
}

function getRelativeTime(ts) {
    if (!ts) return '';
    try {
        const now = Date.now();
        const then = new Date(ts).getTime();
        const diff = now - then;
        const mins = Math.floor(diff / 60000);
        if (mins < 1) return 'just now';
        if (mins < 60) return mins + 'm ago';
        const hrs = Math.floor(mins / 60);
        if (hrs < 24) return hrs + 'h ago';
        const days = Math.floor(hrs / 24);
        if (days < 30) return days + 'd ago';
        return Math.floor(days / 30) + 'mo ago';
    } catch (e) {
        return '';
    }
}
