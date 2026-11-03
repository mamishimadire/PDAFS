/*
 * ╔══════════════════════════════════════════════════════════════════════════════╗
 * ║  AFS PDF COMPARISON ANALYZER  v4.3 — Front-end application script           ║
 * ║  SNG Grant Thornton | CAATs Platform                                         ║
 * ║  Author  : Mamishi Tonny Madire                                              ║
 * ║  Date    : 2026-03-15                                                        ║
 * ║                                                                              ║
 * ║  Handles all accordion step interactions via AJAX (no full page reloads).    ║
 * ║  Mirrors the Python notebook widget event handlers:                          ║
 * ║    on_eng() · on_add() · on_clr() · on_extract() · on_compare() · on_save() ║
 * ╚══════════════════════════════════════════════════════════════════════════════╝
 */

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 1 · UTILITIES
// ─────────────────────────────────────────────────────────────────────────────

function getToken() {
    const el = document.querySelector('#__afs-af-form input[name="__RequestVerificationToken"]');
    return el ? el.value : '';
}

function esc(s) {
    if (!s) return '';
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

const dashboardState = {
    engagementSaved: false,
    queueCount: 0,
    extractedCount: 0,
    pairCount: 0,
    changedCount: 0,
    exportSaved: false,
};

function setBadge(id, text, tone) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = text;
    el.classList.remove('is-done', 'is-live', 'is-pending');
    const cls = tone === 'done' ? 'is-done' : tone === 'live' ? 'is-live' : 'is-pending';
    el.classList.add(cls);
    // Also update the sibling nav badge (sidebar) if this is a step-state element
    if (id.startsWith('step-state-')) {
        const step = id.replace('step-state-', '');
        const navBadge = document.getElementById('nav-status-' + step);
        if (navBadge) {
            navBadge.textContent = text;
            navBadge.classList.remove('is-done', 'is-live', 'is-pending');
            navBadge.classList.add(cls);
        }
    }
}

function setMetric(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function setActiveStep(step) {
    document.querySelectorAll('[data-step-nav]').forEach(el => {
        const isActive = el.dataset.stepNav === step;
        el.classList.toggle('active', isActive);
    });
}

function currentWorkflowStage() {
    if (dashboardState.exportSaved) return 'Working papers saved successfully';
    if (dashboardState.pairCount > 0) return 'Comparison complete. Review results and export the pack.';
    if (dashboardState.extractedCount >= 2) return 'Extraction complete. Run comparison next.';
    if (dashboardState.queueCount >= 2) return 'PDFs queued. Extract content to prepare comparison.';
    if (dashboardState.engagementSaved) return 'Engagement saved. Upload the AFS PDFs to continue.';
    return 'Awaiting engagement details';
}

function refreshDashboard() {
    setMetric('metric-queue', dashboardState.queueCount);
    setMetric('metric-extracted', dashboardState.extractedCount);
    setMetric('metric-pairs', dashboardState.pairCount);
    setMetric('metric-changed', dashboardState.changedCount);

    setBadge('nav-status-step1', dashboardState.engagementSaved ? 'Saved' : 'Ready',
        dashboardState.engagementSaved ? 'done' : 'live');
    setBadge('step-state-step1', dashboardState.engagementSaved ? 'Saved' : 'Ready',
        dashboardState.engagementSaved ? 'done' : 'live');

    const uploadText = dashboardState.queueCount > 0 ? `${dashboardState.queueCount} queued` : 'Pending';
    setBadge('nav-status-step2', uploadText, dashboardState.queueCount > 0 ? 'done' : 'pending');
    setBadge('step-state-step2', uploadText, dashboardState.queueCount > 0 ? 'done' : 'pending');

    const extractReady = dashboardState.queueCount >= 2;
    const extractText = dashboardState.extractedCount > 0
        ? `${dashboardState.extractedCount} ready`
        : extractReady ? 'Ready' : 'Pending';
    setBadge('nav-status-step3', extractText, dashboardState.extractedCount > 0 ? 'done' : extractReady ? 'live' : 'pending');
    setBadge('step-state-step3', extractText, dashboardState.extractedCount > 0 ? 'done' : extractReady ? 'live' : 'pending');

    const compareReady = dashboardState.extractedCount >= 2;
    const compareText = dashboardState.pairCount > 0
        ? `${dashboardState.pairCount} pairs`
        : compareReady ? 'Ready' : 'Pending';
    setBadge('nav-status-step4', compareText, dashboardState.pairCount > 0 ? 'done' : compareReady ? 'live' : 'pending');
    setBadge('step-state-step4', compareText, dashboardState.pairCount > 0 ? 'done' : compareReady ? 'live' : 'pending');

    const exportText = dashboardState.exportSaved ? 'Saved'
        : dashboardState.pairCount > 0 ? 'Ready' : 'Pending';
    setBadge('nav-status-step5', exportText, dashboardState.exportSaved ? 'done' : dashboardState.pairCount > 0 ? 'live' : 'pending');
    setBadge('step-state-step5', exportText, dashboardState.exportSaved ? 'done' : dashboardState.pairCount > 0 ? 'live' : 'pending');

    const stageEl = document.getElementById('workflow-stage');
    if (stageEl) stageEl.textContent = currentWorkflowStage();
}

function renderModelStatus(status) {
    const card = document.getElementById('boost-status-card');
    const valueEl = document.getElementById('boost-status-value');
    const summaryEl = document.getElementById('boost-status-summary');
    const lightsEl = document.getElementById('boost-status-lights');
    if (!card || !valueEl || !summaryEl || !lightsEl) return;

    const providers = Array.isArray(status && status.providers) ? status.providers : [];
    const boostActive = !!(status && status.boostActive);

    card.classList.toggle('is-active', boostActive);
    card.classList.toggle('is-inactive', !boostActive);
    valueEl.textContent = boostActive ? 'Active' : 'Inactive';
    summaryEl.textContent = '';

    lightsEl.innerHTML = providers.map(provider => {
        const active = !!provider.active;
        const stateClass = active ? 'is-active' : 'is-inactive';
        return `<span class="boost-model-light ${stateClass}" data-key="${esc((provider.key || '').toLowerCase())}">
            <span class="boost-model-light-dot"></span>
        </span>`;
    }).join('');

    if (!providers.length) {
        lightsEl.innerHTML = `<span class="boost-model-light is-inactive"><span class="boost-model-light-dot"></span></span>`;
    }
}

async function loadModelStatus() {
    try {
        const status = await apiGetJson('/api/model-status');
        renderModelStatus(status);
    } catch (e) {
        renderModelStatus({
            boostActive: false,
            summary: 'Model status unavailable right now.',
            providers: []
        });
    }
}

function focusStep(step) {
    openAcc(step);
    setActiveStep(step);
    const container = document.getElementById('acc-' + step);
    if (container) container.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function setStatus(id, msg, type) {
    const el = document.getElementById(id);
    if (!el) return;
    const colors = { ok: '#00C9A7', error: '#EF4444', warn: '#F59E0B', info: '#A78BFA' };
    el.style.color = colors[type] || '#CBD5E1';
    el.textContent = msg;
}

function showBanner(containerId, msg, type) {
    const colors = { ok: '#00C9A7', error: '#EF4444', warn: '#F59E0B', info: '#7B4FFF' };
    const c = colors[type] || '#7B4FFF';
    const el = document.getElementById(containerId);
    if (!el) return;
    el.innerHTML = `<div style="background:${c}22;border-left:4px solid ${c};padding:9px 14px;border-radius:5px;margin:5px 0;font-family:'IBM Plex Sans',sans-serif;font-size:13px;font-weight:600;color:${c}">${esc(msg)}</div>`;
}

function friendlyRequestError(err, timeoutMs) {
    if (err && err.name === 'AbortError') {
        return `Request timed out after ${Math.round(timeoutMs / 1000)}s.`;
    }

    const msg = String(err && err.message ? err.message : '').trim();
    if (!msg) {
        return 'The request could not be completed. Please retry.';
    }

    if (msg.includes('Failed to fetch') || msg.includes('NetworkError') || msg.includes('Load failed')) {
        return 'Unable to reach the server. Please retry.';
    }

    return msg;
}

async function apiGetJson(url, options = {}) {
    const timeoutMs = options.timeoutMs || 15000;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

    try {
        const r = await fetch(url, {
            method: 'GET',
            signal: controller.signal,
            headers: { 'Accept': 'application/json' }
        });

        if (!r.ok) {
            throw new Error(`Server returned HTTP ${r.status}.`);
        }

        const body = (await r.text()).trim();
        if (!body) {
            throw new Error('Server returned an empty response.');
        }

        try {
            return JSON.parse(body);
        } catch {
            throw new Error('Server returned an invalid response.');
        }
    } catch (err) {
        throw new Error(friendlyRequestError(err, timeoutMs));
    } finally {
        clearTimeout(timeout);
    }
}

async function apiPost(url, formData, options = {}) {
    const timeoutMs = options.timeoutMs || 30000;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

    if (!formData.has('__RequestVerificationToken')) {
        formData.append('__RequestVerificationToken', getToken());
    }

    try {
        const r = await fetch(url, {
            method: 'POST',
            body: formData,
            signal: controller.signal,
            headers: { 'Accept': 'application/json' }
        });

        const raw = await r.text();
        const body = (raw || '').trim();

        if (!r.ok) {
            throw new Error(`Server returned HTTP ${r.status}.`);
        }

        if (!body) {
            throw new Error('Server returned an empty response.');
        }

        try {
            return JSON.parse(body);
        } catch {
            throw new Error('Server returned an invalid response.');
        }
    } catch (err) {
        throw new Error(friendlyRequestError(err, timeoutMs));
    } finally {
        clearTimeout(timeout);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 2 · ACCORDION TOGGLE
// Mirrors Python acc.selected_index behaviour.
// ─────────────────────────────────────────────────────────────────────────────

function toggleAcc(step) {
    const body   = document.getElementById('body-' + step);
    const arrow  = document.getElementById('arrow-' + step);
    const header = body ? body.previousElementSibling : null;
    if (!body) return;
    const open = body.style.display !== 'none';
    body.style.display = open ? 'none' : 'block';
    if (arrow) arrow.classList.toggle('is-open', !open);
    if (header) header.classList.toggle('is-open', !open);
    if (!open) setActiveStep(step);
}

function openAcc(step) {
    const body   = document.getElementById('body-' + step);
    const arrow  = document.getElementById('arrow-' + step);
    const header = body ? body.previousElementSibling : null;
    if (body)   body.style.display = 'block';
    if (arrow)  arrow.classList.add('is-open');
    if (header) header.classList.add('is-open');
    setActiveStep(step);
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 3 · STEP 1 — ENGAGEMENT DETAILS
// Reference: Python on_eng() handler
// ─────────────────────────────────────────────────────────────────────────────

async function saveEngagement() {
    const fd = new FormData(document.getElementById('engForm'));
    // explicitly add named fields (form has no action, use ids)
    const fields = {
        'parent':              'eng_parent',
        'client':              'eng_client',
        'engagementName':      'eng_engname',
        'engagementNumber':    'eng_engnum',
        'financialYearStart':  'eng_fystart',
        'financialYearEnd':    'eng_fyend',
        'preparedBy':          'eng_auditor',
        'director':            'eng_director',
        'manager':             'eng_manager',
        'objective':           'eng_objective',
    };
    const data = new FormData();
    for (const [name, id] of Object.entries(fields)) {
        data.append(name, document.getElementById(id)?.value || '');
    }
    setStatus('eng-status', 'Saving...', 'info');
    try {
        const res = await apiPost('/api/engagement', data);
        if (res.success) {
            dashboardState.engagementSaved = true;
            refreshDashboard();
            openAcc('step2');
        }
        setStatus('eng-status', res.message || (res.success ? '✓ Saved' : 'Error'), res.success ? 'ok' : 'error');
    } catch (e) {
        setStatus('eng-status', 'Error: ' + e.message, 'error');
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 4 · STEP 2 — UPLOAD & QUEUE
// Reference: Python on_add() / on_clr() handlers
// ─────────────────────────────────────────────────────────────────────────────

async function addToQueue() {
    const input = document.getElementById('pdfInput');
    if (!input || !input.files.length) return;
    const data = new FormData();
    for (const f of input.files) data.append('files', f);
    setStatus('upload-status', 'Uploading...', 'info');
    try {
        const res = await apiPost('/api/upload', data);
        setStatus('upload-status', res.message, res.success ? 'ok' : 'error');
        renderQueue(res.queue || []);
        resetDownstreamResults();
        input.value = '';
        // Auto-extract once at least 2 PDFs are in the queue
        if (res.success && (res.queue || []).length >= 2) {
            openAcc('step3');
            await extractPdfs();
        }
    } catch (e) {
        setStatus('upload-status', 'Error: ' + e.message, 'error');
    }
}

async function clearQueue() {
    const data = new FormData();
    setStatus('upload-status', 'Clearing...', 'info');
    try {
        const res = await apiPost('/api/clear', data);
        setStatus('upload-status', res.message, 'warn');
        renderQueue([]);
        resetDownstreamResults();
    } catch (e) {
        setStatus('upload-status', 'Error: ' + e.message, 'error');
    }
}

function resetDownstreamResults() {
    window._cmpData = null;
    dashboardState.extractedCount = 0;
    dashboardState.pairCount = 0;
    dashboardState.changedCount = 0;
    dashboardState.exportSaved = false;
    refreshDashboard();
    document.getElementById('extract-results').innerHTML = '';
    document.getElementById('compare-results').innerHTML = '';
    document.getElementById('export-results').innerHTML = '';
}

function renderQueue(queue) {
    const wrap = document.getElementById('queue-table-body');
    dashboardState.queueCount = queue ? queue.length : 0;
    dashboardState.extractedCount = (queue || []).filter(f => f.docType && f.docType !== 'queued').length;
    dashboardState.pairCount = 0;
    dashboardState.changedCount = 0;
    dashboardState.exportSaved = false;
    refreshDashboard();
    if (!queue || queue.length === 0) {
        wrap.innerHTML = '<div class="afs-empty-state">No PDFs queued yet.</div>';
        return;
    }
    let rows = '';
    queue.forEach((f, i) => {
        const bg   = i % 2 === 0 ? '#162347' : '#111827';
        const badge = f.docType === 'digital' ? '#059669' : f.docType === 'scanned' ? '#D97706' : '#374151';
        rows += `<tr style="background:${bg}">
            <td style="${tdStyle}">${esc(f.label)}</td>
            <td style="${tdStyle}">${esc(f.filename)}</td>
            <td style="${tdStyle}">${f.kb} KB</td>
            <td style="${tdStyle}"><span style="background:${badge}33;color:${badge};border:1px solid ${badge}66;border-radius:3px;padding:1px 7px;font-size:10px;font-family:'IBM Plex Mono',monospace">${esc(f.docType || 'queued')}</span></td>
        </tr>`;
    });
    wrap.innerHTML = `
        <table style="width:100%;border-collapse:collapse">
            <thead><tr style="background:#1E2F5A">
                <th style="${thStyle}">LABEL</th>
                <th style="${thStyle}">FILENAME</th>
                <th style="${thStyle}">SIZE</th>
                <th style="${thStyle}">TYPE</th>
            </tr></thead>
            <tbody>${rows}</tbody>
        </table>`;
}

const thStyle = 'padding:6px 10px;color:#93C5FD;font-size:11px;font-family:"IBM Plex Sans",sans-serif;border-bottom:1px solid rgba(255,255,255,0.08);text-align:left';
const tdStyle = 'padding:6px 10px;color:#CBD5E1;font-size:12px;font-family:"IBM Plex Sans",sans-serif;border-bottom:1px solid rgba(255,255,255,0.05)';

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 5 · STEP 3 — EXTRACT
// Reference: Python on_extract() handler
// ─────────────────────────────────────────────────────────────────────────────

async function extractPdfs() {
    setStatus('extract-status', 'Extracting... OCR-heavy PDFs can take a few minutes.', 'info');
    document.getElementById('extract-results').innerHTML = '';
    try {
        const res = await apiPost('/api/extract', new FormData(), { timeoutMs: 600000 });
        setStatus('extract-status', res.message, res.success ? 'ok' : 'warn');
        renderExtractResults(res.reports || [], res.errors || []);
        if (res.success) openAcc('step4');
    } catch (e) {
        setStatus('extract-status', 'Error: ' + e.message, 'error');
    }
}

function renderExtractResults(reports, errors) {
    const wrap = document.getElementById('extract-results');
    dashboardState.queueCount = reports.length;
    dashboardState.extractedCount = reports.filter(r => r.docType !== 'error').length;
    dashboardState.pairCount = 0;
    dashboardState.changedCount = 0;
    dashboardState.exportSaved = false;
    refreshDashboard();
    if (!reports.length) { wrap.innerHTML = ''; return; }

    let rows = '';
    reports.forEach((r, i) => {
        const bg = i % 2 === 0 ? '#162347' : '#111827';
        const docType = (r.docType || '').toLowerCase();
        const tc = docType.includes('digital')
            ? '#059669'
            : (docType.includes('scanned') || docType.includes('ocr'))
                ? '#D97706'
                : docType.includes('corrupted')
                    ? '#EF4444'
                    : '#EF4444';
        rows += `<tr style="background:${bg}">
            <td style="${tdStyle}">${esc(r.label)}</td>
            <td style="${tdStyle}">${esc(r.filename)}</td>
            <td style="${tdStyle}"><span style="color:${tc}">${esc(r.docType)}</span></td>
            <td style="${tdStyle}">${esc(r.year || '?')}</td>
            <td style="${tdStyle}">${r.pages}</td>
            <td style="${tdStyle}">${r.words.toLocaleString()}</td>
        </tr>`;
    });

    let errHtml = '';
    errors.forEach(e => {
        errHtml += `<div style="color:#FCA5A5;font-size:11px;margin:3px 0;font-family:'IBM Plex Mono',monospace">⚠ ${esc(e)}</div>`;
    });

    wrap.innerHTML = `
        <table style="width:100%;border-collapse:collapse;margin-top:8px">
            <thead><tr style="background:#1E2F5A">
                <th style="${thStyle}">LABEL</th><th style="${thStyle}">FILE</th>
                <th style="${thStyle}">TYPE</th><th style="${thStyle}">YEAR</th>
                <th style="${thStyle}">PAGES</th><th style="${thStyle}">WORDS</th>
            </tr></thead>
            <tbody>${rows}</tbody>
        </table>${errHtml}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 6 · STEP 4 — COMPARE
// Reference: Python on_compare() / _render_comparison() handlers
// ─────────────────────────────────────────────────────────────────────────────

async function runComparison() {
    const statusEl = document.getElementById('compare-status');
    const progWrap = document.getElementById('compare-progress');
    const progBar  = document.getElementById('compare-prog-bar');
    const progLbl  = document.getElementById('compare-prog-label');
    const results  = document.getElementById('compare-results');
    const runBtn   = document.querySelector('button[onclick="runComparison()"]');

    try {
        const q = await fetch('/api/queue').then(r => r.json());
        const queue = q?.queue || [];
        const extracted = queue.filter(f => f.docType && f.docType !== 'queued').length;

        renderQueue(queue);

        if (queue.length < 2) {
            document.getElementById('extract-results').innerHTML = '';
            results.innerHTML = '';
            setStatus('compare-status', 'Session expired or server restarted. Upload the PDFs again.', 'error');
            openAcc('step2');
            return;
        }

        if (extracted < 2) {
            document.getElementById('extract-results').innerHTML = '';
            results.innerHTML = '';
            setStatus('compare-status', 'Server session no longer has extracted reports. Run Extract again.', 'error');
            openAcc('step3');
            return;
        }
    } catch (e) {
        setStatus('compare-status', 'Could not confirm current session state. Refresh and try again.', 'error');
        return;
    }

    setStatus('compare-status', 'Running comparison... Core page diff can take a few minutes on large PDFs.', 'info');
    progWrap.style.display = 'block';
    results.innerHTML      = '';
    if (runBtn) runBtn.disabled = true;
    progBar.classList.remove('is-error');

    // Animate progress bar with rotating stage messages so user sees active progress
    const stages = [
        { pct: 10,  msg: 'Aligning pages between documents…' },
        { pct: 22,  msg: 'Extracting text from all pages…' },
        { pct: 38,  msg: 'Running line-by-line comparison…' },
        { pct: 52,  msg: 'Computing word-level differences…' },
        { pct: 64,  msg: 'Building snapshot highlights...' },
        { pct: 75,  msg: 'Rendering page snapshots...' },
        { pct: 84,  msg: 'Finalising comparison results...' },
        { pct: 91,  msg: 'Core page diff still running...' },
    ];
    let pct = 0, stageIdx = 0;
    const startTime = Date.now();
    const formatElapsed = (elapsedSeconds) => {
        const mins = Math.floor(elapsedSeconds / 60);
        const secs = elapsedSeconds % 60;
        return `${mins}:${String(secs).padStart(2, '0')}`;
    };
    const timer = setInterval(() => {
        const elapsed = Math.floor((Date.now() - startTime) / 1000);
        // Advance through stages based on elapsed time (each stage ~7-10 s)
        while (stageIdx < stages.length - 1 && elapsed > (stageIdx + 1) * 8) stageIdx++;
        const target = stages[stageIdx].pct;
        pct += (target - pct) * 0.12;
        const shown = Math.min(pct, 91);
        progBar.style.width = shown + '%';
        progLbl.textContent = `${stages[stageIdx].msg} Elapsed ${formatElapsed(elapsed)}.`;
    }, 300);

    const finishProgress = (success) => {
        clearInterval(timer);
        progBar.classList.toggle('is-error', !success);
        progBar.style.width = '100%';
        progLbl.textContent = success ? 'Done' : 'Failed';
    };

    try {
        const res = await apiPost('/api/compare', new FormData(), { timeoutMs: 900000 });
        finishProgress(!!res.success);
        setTimeout(() => { progWrap.style.display = 'none'; }, 600);

        if (!res.success) {
            setStatus('compare-status', res.message, 'error');
            return;
        }
        setStatus('compare-status', `${res.message} Deep analysis running in background…`, 'ok');
        renderComparison(res);
        openAcc('step5');
        // Fire AI analysis in background — does NOT delay comparison display
        runAiAnalysis();
    } catch (e) {
        finishProgress(false);
        progWrap.style.display = 'none';
        setStatus('compare-status', 'Error: ' + e.message, 'error');
    } finally {
        if (runBtn) runBtn.disabled = false;
    }
}

// Fires after comparison renders — calls /api/ai-analyze in background
// then patches the Insights tab without a full re-render.
async function runAiAnalysis(isRetry = false) {
    const insightsTab = document.getElementById('tab-ai');
    if (insightsTab) {
        const msg = isRetry ? 'Switching to backup model — retrying…' : 'Deep analysis in progress…';
        const sub = isRetry ? 'A different AI model is now processing your documents…' : 'Reading both documents and comparing changes…';
        insightsTab.innerHTML = `<div style="padding:20px;text-align:center;color:#A78BFA">
            <div style="font-size:13px;font-weight:600;margin-bottom:8px">${msg}</div>
            <div style="width:100%;height:4px;background:#1E2F5A;border-radius:2px;overflow:hidden">
                <div style="height:4px;background:linear-gradient(90deg,#7B4FFF,#C8102E);border-radius:2px;animation:afs-scan 1.6s ease-in-out infinite"></div>
            </div>
            <div style="margin-top:8px;color:#475569;font-size:11.5px">${sub}</div>
        </div>`;
    }
    try {
        const res = await apiPost('/api/ai-analyze', new FormData(), { timeoutMs: 360000 });
        const aiReady   = res.aiAssist      && res.aiAssist.status      === 'ready';
        const readReady = res.aiContentRead && res.aiContentRead.status === 'ready';

        if (!res.success || (!aiReady && !readReady)) {
            if (!isRetry) {
                setTimeout(() => runAiAnalysis(true), 8000);
                return;
            }
            if (insightsTab) insightsTab.innerHTML = renderAiUnavailableNotice(
                'Intelligent review could not be completed.',
                'All AI models returned errors. The document comparison is still available. You can retry below.',
                true);
            return;
        }
        if (window._cmpData) {
            window._cmpData.aiAssist      = res.aiAssist      || null;
            window._cmpData.aiContentRead = res.aiContentRead || null;
        }
        const newHtml = renderAiReviewTab(res.aiAssist || null, res.aiContentRead || null);
        const wrap = document.createElement('div');
        wrap.innerHTML = newHtml;
        const newPanel = wrap.querySelector('#tab-ai');
        if (newPanel && insightsTab) {
            newPanel.style.display = insightsTab.style.display;
            insightsTab.replaceWith(newPanel);
        }
        const findCount = res.aiAssist && res.aiAssist.materialFindings ? res.aiAssist.materialFindings.length : 0;
        setStatus('compare-status', `Comparison complete. ${findCount} insight(s) found.`, 'ok');
        document.querySelectorAll('.afs-tab').forEach(btn => {
            if (btn.textContent.includes('Review')) {
                btn.innerHTML = `Review <span class="afs-badge">${findCount}</span>`;
            }
        });
    } catch (e) {
        if (!isRetry) {
            setTimeout(() => runAiAnalysis(true), 8000);
            return;
        }
        if (insightsTab) insightsTab.innerHTML = renderAiUnavailableNotice(
            'Intelligent review could not be completed.',
            'A network error occurred. The document comparison is still available. You can retry below.',
            true);
    }
}

function renderComparison(data) {
    const wrap = document.getElementById('compare-results');
    const c    = data.counts;
    const al   = data.alignment;
    const nc   = data.numCmp;
    const ai         = data.aiAssist || null;
    const aiRead     = data.aiContentRead || null;
    const autoReview         = data.autoReview || [];
    const materialityThresh  = data.materialityThreshold || 0;
    const total = c.total || 1;
    const pctAffected = ((c.changed + c.added + c.removed) / total * 100).toFixed(1);
    dashboardState.pairCount = data.pageDiffs ? data.pageDiffs.length : 0;
    dashboardState.changedCount = (c.changed || 0) + (c.added || 0) + (c.removed || 0);
    dashboardState.exportSaved = false;
    refreshDashboard();

    // ── Smart page alignment info bar ──────────────────────────────────────
    let html = `
    <div style="background:#0D1B2A;border:1px solid #7B4FFF44;border-radius:9px;padding:10px 14px;margin-bottom:10px;font-family:'IBM Plex Mono',monospace;font-size:11px">
        <span style="color:#A78BFA;font-weight:700">🔗 SMART PAGE ALIGNMENT</span>  |
        <span style="color:#10B981"> AFS1: ${al.afs1Pages} pages</span>
        <span style="color:#3B82F6"> AFS2: ${al.afs2Pages} pages</span>  |
        <span style="color:#FBBF24"> ${al.matched} content-matched pairs</span>
        <span style="color:#EF4444"> ${al.unmatched} AFS1 pages unmatched</span>
    </div>`;

    // ── KPI cards ────────────────────────────────────────────────────────
    html += `<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px">
        ${kpiCard('✓ Same Lines',   c.same,    Math.round(c.same/total*100)+'% of all', '#10B981')}
        ${kpiCard('~ Changed',      c.changed, 'content differs',                        '#F59E0B')}
        ${kpiCard('+ Added',        c.added,   'in AFS 2 only',                          '#3B82F6')}
        ${kpiCard('− Removed',      c.removed, 'in AFS 1 only',                          '#EF4444')}
        ${kpiCard('Numbers Match',  nc.simPct+'%', '% of all numbers',                   '#7B4FFF')}
        ${kpiCard('% Affected',     pctAffected+'%', 'changed+added+removed',            '#F59E0B')}
    </div>`;

    // ── Tab bar ──────────────────────────────────────────────────────────
    html += `
    <div class="afs-tabs" id="cmpTabs">
        <button class="afs-tab active" onclick="switchTab('tab-changed')">~ Changed <span class="afs-badge">${data.totalChanged}</span></button>
        <button class="afs-tab" onclick="switchTab('tab-added')">+ Added <span class="afs-badge">${data.totalAdded}</span></button>
        <button class="afs-tab" onclick="switchTab('tab-removed')">− Removed <span class="afs-badge">${data.totalRemoved}</span></button>
        <button class="afs-tab" onclick="switchTab('tab-numbers')">Numbers</button>
        <button class="afs-tab" onclick="switchTab('tab-pagemap')">Page Map</button>
        <button class="afs-tab" onclick="switchTab('tab-visual')">📄 Visual</button>
        <button class="afs-tab" onclick="switchTab('tab-ai')">Review <span class="afs-badge">${ai && ai.materialFindings ? ai.materialFindings.length : 0}</span></button>
        <button class="afs-tab" onclick="switchTab('tab-sysreview')">System Review <span class="afs-badge">${autoReview.length}</span></button>
        <button class="afs-tab" onclick="switchTab('tab-comments')">Comments</button>
    </div>
    <div class="afs-tab-content">`;

    // ── Tab: Changed lines ────────────────────────────────────────────────
    html += `<div id="tab-changed" class="afs-tab-pane">
        <div style="color:#F59E0B;font-weight:700;font-size:13px;margin:10px 0 6px">~ CHANGED LINES — word diff + auditor comment</div>`;
    if (data.changedLines && data.changedLines.length) {
        html += `<table style="width:100%;border-collapse:collapse">
            <thead><tr style="background:#1E2F5A">
                <th style="${thStyle};width:32px">#</th>
                <th style="${thStyle}">WORD DIFF (red=removed from AFS1, blue=added in AFS2)</th>
                <th style="${thStyle};width:110px">NUM CHANGES</th>
            </tr></thead><tbody>`;
        data.changedLines.forEach((d, i) => {
            const bg  = i % 2 === 0 ? '#1C1400' : '#14100A';
            const wd  = renderWordDiff(d.wordDiff);
            const nd  = d.numDiff && d.numDiff.length ? esc(d.numDiff.join(', ')) : '—';
            html += `<tr style="background:${bg};vertical-align:top">
                <td style="${tdStyle};color:#F59E0B;font-family:'IBM Plex Mono',monospace;white-space:nowrap">${i+1}</td>
                <td style="${tdStyle};font-family:'IBM Plex Mono',monospace;max-width:640px">${wd}</td>
                <td style="${tdStyle};color:#FBBF24;font-family:'IBM Plex Mono',monospace">${nd}</td>
            </tr>`;
        });
        html += '</tbody></table>';

        // Comment boxes for first 20 changed lines
        html += `<div style="margin:10px 0 4px;color:#A78BFA;font-size:11px;font-family:'IBM Plex Mono',monospace">✏ AUDITOR COMMENT BOXES — first 20 changed lines</div>`;
        data.changedLines.slice(0, 20).forEach((d, i) => {
            const lbl  = `~${i+1}  ${(d.line1 || '').substring(0, 55)}`;
            const key  = `changed:${i+1}`;
            html += commentBox(key, lbl);
        });
    } else {
        html += '<p style="color:#10B981;padding:10px">✓ No changed lines detected.</p>';
    }
    html += '</div>';

    // ── Tab: Added lines ─────────────────────────────────────────────────
    html += `<div id="tab-added" class="afs-tab-pane" style="display:none">
        <div style="color:#3B82F6;font-weight:700;font-size:13px;margin:10px 0 6px">+ LINES ADDED IN ${esc(data.afs2Label)}</div>`;
    if (data.addedLines && data.addedLines.length) {
        html += `<table style="width:100%;border-collapse:collapse">
            <thead><tr style="background:#1E3A5F">
                <th style="${thStyle};width:40px">#</th>
                <th style="${thStyle}">ADDED LINE</th>
            </tr></thead><tbody>`;
        data.addedLines.forEach((d, i) => {
            const bg = i % 2 === 0 ? '#0D1B2A' : '#081424';
            html += `<tr style="background:${bg}">
                <td style="${tdStyle};color:#60A5FA;font-family:'IBM Plex Mono',monospace">${d.i}</td>
                <td style="${tdStyle};color:#BAE6FD;font-family:'IBM Plex Mono',monospace">${esc(d.line)}</td>
            </tr>`;
        });
        html += '</tbody></table>';
    } else {
        html += '<p style="color:#94A3B8;padding:10px">No added lines.</p>';
    }
    html += '</div>';

    // ── Tab: Removed lines ───────────────────────────────────────────────
    html += `<div id="tab-removed" class="afs-tab-pane" style="display:none">
        <div style="color:#EF4444;font-weight:700;font-size:13px;margin:10px 0 6px">− LINES REMOVED FROM ${esc(data.afs1Label)}</div>`;
    if (data.removedLines && data.removedLines.length) {
        html += `<table style="width:100%;border-collapse:collapse">
            <thead><tr style="background:#5F1D1D">
                <th style="${thStyle};width:40px">#</th>
                <th style="${thStyle}">REMOVED LINE</th>
            </tr></thead><tbody>`;
        data.removedLines.forEach((d, i) => {
            const bg = i % 2 === 0 ? '#2A0D0D' : '#1A0808';
            html += `<tr style="background:${bg}">
                <td style="${tdStyle};color:#F87171;font-family:'IBM Plex Mono',monospace">${d.i}</td>
                <td style="${tdStyle};color:#FECACA;font-family:'IBM Plex Mono',monospace">${esc(d.line)}</td>
            </tr>`;
        });
        html += '</tbody></table>';
    } else {
        html += '<p style="color:#94A3B8;padding:10px">No removed lines.</p>';
    }
    html += '</div>';

    // ── Tab: Number comparison ───────────────────────────────────────────
    html += `<div id="tab-numbers" class="afs-tab-pane" style="display:none">
        <div style="color:#FBBF24;font-weight:700;font-size:13px;margin:10px 0 6px">NUMBER COMPARISON</div>
        <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px">
            ${kpiCard('AFS1 numbers', nc.countAfs1, 'unique numeric tokens', '#EF4444')}
            ${kpiCard('AFS2 numbers', nc.countAfs2, 'unique numeric tokens', '#3B82F6')}
            ${kpiCard('In both',      nc.inBoth.length, 'shared numbers',     '#10B981')}
            ${kpiCard('Similarity',   nc.simPct+'%', 'Jaccard overlap',       '#7B4FFF')}
        </div>`;
    if (nc.onlyAfs1.length) {
        html += numTable('Only in AFS 1 (missing from AFS 2)', nc.onlyAfs1, '#EF4444');
    }
    if (nc.onlyAfs2.length) {
        html += numTable('Only in AFS 2 (new in AFS 2)', nc.onlyAfs2, '#3B82F6');
    }
    html += '</div>';

    // ── Tab: Page map ────────────────────────────────────────────────────
    html += `<div id="tab-pagemap" class="afs-tab-pane" style="display:none">
        <div style="color:#A78BFA;font-weight:700;font-size:13px;margin:10px 0 6px">PAGE ALIGNMENT SUMMARY</div>
        <table style="width:100%;border-collapse:collapse">
            <thead><tr style="background:#1E2F5A">
                <th style="${thStyle}">AFS1 PAGE</th><th style="${thStyle}">AFS2 PAGE</th>
                <th style="${thStyle}">ALIGN SIM</th><th style="${thStyle}">% SAME</th>
                <th style="${thStyle}">SAME</th><th style="${thStyle}">CHANGED</th>
                <th style="${thStyle}">ADDED</th><th style="${thStyle}">REMOVED</th>
            </tr></thead><tbody>`;
    if (data.pageDiffs && data.pageDiffs.length) {
        data.pageDiffs.forEach((p, i) => {
            const bg  = i % 2 === 0 ? '#162347' : '#111827';
            const pct = p.pctSame;
            const fc  = pct === 100 ? '#10B981' : pct >= 70 ? '#F59E0B' : '#EF4444';
            const flag = pct === 100 ? '' : (pct >= 70 ? '~' : '!!');
            const aiBadge = renderAiFocusBadge(ai, p.pairIndex ?? i);
            html += `<tr style="background:${bg}">
                <td style="${tdStyle}">AFS1 p${p.pageAfs1 || '?'} ${aiBadge}</td>
                <td style="${tdStyle}">${p.pageAfs2 ? 'AFS2 p'+p.pageAfs2 : '<span style="color:#EF4444">—</span>'}</td>
                <td style="${tdStyle};font-family:'IBM Plex Mono',monospace">${p.alignSim.toFixed(3)}</td>
                <td style="${tdStyle};color:${fc};font-weight:600">${flag} ${pct}%</td>
                <td style="${tdStyle};color:#10B981">${p.same}</td>
                <td style="${tdStyle};color:#F59E0B">${p.changed}</td>
                <td style="${tdStyle};color:#3B82F6">${p.added}</td>
                <td style="${tdStyle};color:#EF4444">${p.removed}</td>
            </tr>`;
        });
    }
    html += '</tbody></table>';

    // Page pair detail dropdown
    if (data.pageDiffs && data.pageDiffs.length) {
        html += `<div style="margin:14px 0 4px;color:#A78BFA;font-weight:700;font-size:13px">📄 PAGE DIFF DETAIL</div>`;
        html += `<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px">
            <select id="pairSelect" class="afs-input" style="width:340px">`;
        data.pageDiffs.forEach((p, k) => {
            const issues = p.changed + p.added + p.removed;
            const pct    = p.pctSame ?? 100;
            const flag   = issues === 0        ? ' ✓'
                         : pct >= 98           ? ' ~ minor'
                         : issues > 5          ? ' !!'
                                               : ' !';
            const changeLabel = issues === 0 ? 'no changes'
                              : issues === 1 ? '1 change'
                                             : `${issues} changes`;
            const relation = p.secondaryMatch
                ? `AFS1 p${p.pageAfs1 ?? '?'} ≈ AFS2 p${p.pageAfs2 ?? '?'}`
                : `AFS1 p${p.pageAfs1} ↔ ${p.pageAfs2 ? 'AFS2 p'+p.pageAfs2 : 'NO MATCH'}`;
            const tag = p.secondaryMatch ? ' [cont.]' : '';
            html += `<option value="${p.k ?? k}">${relation}${tag}${flag} (${changeLabel})</option>`;
        });
        html += `</select>
            <button class="afs-btn afs-btn-primary" onclick="showPageDiff()">Show Diff</button>
        </div>
        <div id="pair-detail"></div>`;
    }
    html += '</div>';

    // ── Tab: Page Visual Viewer ──────────────────────────────────────────
    html += `<div id="tab-visual" class="afs-tab-pane" style="display:none">
        <div style="color:#A78BFA;font-weight:700;font-size:13px;margin:10px 0 4px">📄 PAGE VISUAL VIEWER</div>
        <div class="afs-info afs-info-violet" style="margin-bottom:10px">
            Pages are matched by <b>content similarity</b>, not page number.
            Yellow bands highlight changed / added / removed lines.
        </div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:10px">
            <select id="visualPairSelect" class="afs-input" style="width:380px">`;
    if (data.pageDiffs && data.pageDiffs.length) {
        data.pageDiffs.forEach((p, k) => {
            const issues = p.changed + p.added + p.removed;
            const pct    = p.pctSame ?? 100;
            const flag   = issues === 0        ? ' ✓'
                         : pct >= 98           ? ' ~ minor'
                         : issues > 5          ? ' !!'
                                               : ' !';
            const changeLabel = issues === 0 ? 'no changes'
                              : issues === 1 ? '1 change'
                                             : `${issues} changes`;
            const relation = p.secondaryMatch
                ? `AFS1 p${p.pageAfs1 ?? '?'} ≈ AFS2 p${p.pageAfs2 ?? '?'}`
                : `AFS1 p${p.pageAfs1} ↔ ${p.pageAfs2 ? 'AFS2 p'+p.pageAfs2 : 'NO MATCH'}`;
            const tag = p.secondaryMatch ? ' [cont.]' : '';
            html += `<option value="${p.k ?? k}">${relation}${tag}${flag} (${changeLabel})</option>`;
        });
    }
    html += `</select>
            <button class="afs-btn" onclick="snapNav(-1)" id="snap-prev"
                style="padding:6px 12px;background:#1E2F5A;color:#A78BFA;border:1px solid #7B4FFF44;border-radius:8px;font-size:13px;cursor:pointer"
                title="Previous page pair">&#8592; Prev</button>
            <button class="afs-btn afs-btn-primary" onclick="showSnapshot()">Show Snapshot</button>
            <button class="afs-btn" onclick="snapNav(1)" id="snap-next"
                style="padding:6px 12px;background:#1E2F5A;color:#A78BFA;border:1px solid #7B4FFF44;border-radius:8px;font-size:13px;cursor:pointer"
                title="Next page pair">Next &#8594;</button>
            <span id="snap-status" style="font-size:11px;color:#A78BFA;font-family:'IBM Plex Mono',monospace"></span>
        </div>
        <div id="snap-counter" style="font-size:11px;color:#475569;font-family:'IBM Plex Mono',monospace;margin-bottom:6px"></div>
        <div id="snap-results"></div>
    </div>`;

    // ── Tab: System Review (deterministic — no AI required) ──────────────
    html += renderSystemReviewTab(autoReview);

    // ── Tab: Comments ─────────────────────────────────────────────────────
    html += renderAiReviewTab(ai, aiRead);

    html += `<div id="tab-comments" class="afs-tab-pane" style="display:none">
        <div style="color:#FBBF24;font-weight:700;font-size:13px;margin:10px 0 6px">📋 CONSOLIDATED AUDITOR COMMENTS</div>
        ${commentBox('overall', 'Overall Audit Conclusion / Sign-off Notes')}
    </div>`;

    html += '</div></div>'; // end tab-content + compare-results

    wrap.innerHTML = html;

    // Store data for page diff viewer
    data._snapshotNonce = Date.now();
    window._cmpData = data;

    // Pre-populate materiality threshold from server state
    const matInput = document.getElementById('materiality-input');
    if (matInput && materialityThresh > 0) {
        matInput.value = materialityThresh;
        applyMaterialityFilter();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7 · RENDER HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function kpiCard(title, value, sub, color) {
    return `<div style="background:${color}18;border:1px solid ${color}44;border-radius:9px;padding:12px 16px;min-width:140px">
        <div style="color:#888;font-size:10px;text-transform:uppercase;letter-spacing:1px">${esc(title)}</div>
        <div style="color:${color};font-size:22px;font-weight:800">${value}</div>
        <div style="color:#999;font-size:10px">${esc(sub)}</div>
    </div>`;
}

function numTable(title, nums, color) {
    let rows = '';
    nums.forEach((n, i) => {
        const bg = i % 2 === 0 ? '#162347' : '#111827';
        rows += `<tr style="background:${bg}"><td style="${tdStyle}">${i+1}</td><td style="${tdStyle};font-family:'IBM Plex Mono',monospace;color:${color}">${esc(n)}</td></tr>`;
    });
    return `<div style="margin:6px 0 2px;color:${color};font-size:12px;font-weight:600">${esc(title)} (${nums.length})</div>
    <table style="width:48%;border-collapse:collapse;display:inline-table;margin-right:2%">
        <thead><tr style="background:#1E2F5A"><th style="${thStyle};width:40px">#</th><th style="${thStyle}">NUMBER</th></tr></thead>
        <tbody>${rows}</tbody>
    </table>`;
}

function getAiPageFocus(ai, pairIndex) {
    if (!ai || !Array.isArray(ai.pageFocus)) return null;
    return ai.pageFocus.find(p => p.pairIndex === pairIndex) || null;
}

function getAiFindingCountForPair(ai, pairIndex) {
    if (!ai || !Array.isArray(ai.materialFindings)) return 0;
    return ai.materialFindings.filter(f => f.pairIndex === pairIndex).length;
}

function renderAiFocusBadge(ai, pairIndex) {
    const focus = getAiPageFocus(ai, pairIndex);
    if (!focus) return '';
    const tone = (focus.priority || '').toLowerCase() === 'high'
        ? '#F59E0B'
        : (focus.priority || '').toLowerCase() === 'medium'
            ? '#60A5FA'
            : '#94A3B8';
    const findings = getAiFindingCountForPair(ai, pairIndex);
    const suffix = findings > 0 ? ` • ${findings} finding${findings === 1 ? '' : 's'}` : '';
    return `<span style="display:inline-flex;align-items:center;gap:6px;margin-left:6px;background:${tone}22;border:1px solid ${tone}55;color:${tone};border-radius:999px;padding:1px 8px;font-size:10px;font-family:'IBM Plex Mono',monospace">AI ${esc((focus.priority || 'focus').toUpperCase())}${suffix}</span>`;
}

// Builds clickable page reference pills for the Review tab.
// Clicking AFS1/AFS2 pill opens the snapshot for that page.
// Renders one side of a word-diff line with changed words highlighted inline
function renderWordDiffLine(tokens, targetTag) {
    if (!Array.isArray(tokens) || tokens.length === 0) return '<span style="color:#475569;font-style:italic">—</span>';
    let html = '';
    tokens.forEach(t => {
        const word = esc(t.word || '');
        if (!word) return;
        if (t.tag === 'same') {
            html += word + ' ';
        } else if (t.tag === targetTag) {
            const isRem = targetTag === 'removed';
            const clr = isRem ? '#FCA5A5' : '#86EFAC';
            const bg  = isRem ? 'rgba(220,38,38,0.25)' : 'rgba(34,197,94,0.25)';
            html += `<span style="background:${bg};color:${clr};border-radius:3px;padding:0 3px;font-weight:600">${word}</span> `;
        }
    });
    return html.trim() || '<span style="color:#475569;font-style:italic">—</span>';
}

// Returns the changed diff lines for a given pairIndex + ordinal list
function getFindingDiffRows(pairIndex, focusDiffOrdinals) {
    if (typeof pairIndex !== 'number' || !window._cmpData) return [];
    const pd = (window._cmpData.pageDiffs || []).find(p => p.pairIndex === pairIndex);
    if (!pd) return [];
    const issues = (pd.diffs || []).filter(d =>
        d.status !== 'same' && d.status !== 'same-cross-page' && d.status !== 'same-different-page');
    if (Array.isArray(focusDiffOrdinals) && focusDiffOrdinals.length > 0)
        return focusDiffOrdinals.map(o => issues[o - 1]).filter(Boolean).slice(0, 4);
    return issues.slice(0, 3);
}

function buildPageRefHtml(pairIndex, pageAfs1, pageAfs2) {
    let html = '';
    if (typeof pairIndex === 'number') {
        html += `<span style="background:#1E2F5A;color:#A78BFA;border:1px solid #7B4FFF44;border-radius:999px;padding:1px 8px;font-size:10.5px;font-family:'IBM Plex Mono',monospace">
            Pair ${pairIndex + 1}
        </span>`;
    }
    if (pageAfs1) {
        html += `<button onclick="jumpToPageSnapshot(${pairIndex ?? 0}, 1)"
            style="background:#0D2137;color:#93C5FD;border:1px solid #3B82F644;border-radius:999px;padding:1px 9px;font-size:10.5px;font-family:'IBM Plex Mono',monospace;cursor:pointer"
            title="Open AFS 1 page ${pageAfs1} snapshot">
            AFS 1 · p${pageAfs1}
        </button>`;
    }
    if (pageAfs2) {
        html += `<button onclick="jumpToPageSnapshot(${pairIndex ?? 0}, 2)"
            style="background:#0D2137;color:#6EE7B7;border:1px solid #10B98144;border-radius:999px;padding:1px 9px;font-size:10.5px;font-family:'IBM Plex Mono',monospace;cursor:pointer"
            title="Open AFS 2 page ${pageAfs2} snapshot">
            AFS 2 · p${pageAfs2}
        </button>`;
    }
    return html || `<span style="color:#475569;font-size:10.5px;font-family:'IBM Plex Mono',monospace">No page reference</span>`;
}

// Jump to a page pair snapshot from anywhere in the UI (e.g. Review tab findings).
function jumpToPageSnapshot(pairIndex, afsNum) {
    // Switch to Visual tab so the snapshot viewer is visible
    switchTab('tab-visual');
    const pairSelect = document.getElementById('pairSelect');
    if (pairSelect) {
        // Find the option that matches this pairIndex
        const data = window._cmpData;
        if (data && data.pageDiffs) {
            const matchIdx = data.pageDiffs.findIndex(p => (p.pairIndex ?? p.k) === pairIndex);
            if (matchIdx >= 0) {
                pairSelect.value = String(matchIdx);
                pairSelect.dispatchEvent(new Event('change'));
            }
        }
    }
    // Open snapshot view for the correct AFS side then scroll into view
    setTimeout(() => {
        const data = window._cmpData;
        if (!data || !data.pageDiffs) return;
        const p = data.pageDiffs.find(pd => (pd.pairIndex ?? pd.k) === pairIndex);
        if (p) {
            showSnapshot(p, pairIndex, afsNum);
            // Scroll snapshot result into view smoothly
            setTimeout(() => {
                const el = document.getElementById('snapshot-result');
                if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 400);
        }
    }, 150);
}

// ─────────────────────────────────────────────────────────────────────────────
// SYSTEM REVIEW TAB — Deterministic audit findings (no AI required)
// ─────────────────────────────────────────────────────────────────────────────

function renderSystemReviewTab(findings) {
    let html = `<div id="tab-sysreview" class="afs-tab-pane" style="display:none">`;

    // Header row with materiality control (#6)
    html += `<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:10px 0 8px">
        <span style="color:#A78BFA;font-weight:700;font-size:13px">&#128269; SYSTEM REVIEW</span>
        <span style="background:#7B4FFF22;border:1px solid #7B4FFF55;color:#C4B5FD;border-radius:999px;padding:1px 10px;font-size:10.5px;font-family:'IBM Plex Mono',monospace">Deterministic &#8212; no AI</span>
        <span style="margin-left:auto;display:flex;align-items:center;gap:6px;font-size:11px;color:#94A3B8">
            Materiality threshold:
            <input id="materiality-input" type="number" min="0" step="1000" value="0"
                style="width:90px;background:#0D1B2A;border:1px solid #334155;color:#E2E8F0;border-radius:5px;padding:2px 6px;font-size:11px;font-family:'IBM Plex Mono',monospace"
                title="Hide findings where the absolute change is below this amount (0 = show all)"
                oninput="applyMaterialityFilter()">
            <span style="color:#475569">R</span>
        </span>
    </div>`;

    if (!findings || findings.length === 0) {
        html += `<div style="padding:24px;text-align:center;color:#64748B;font-size:13px;background:#0D1B2A;border-radius:8px;border:1px dashed #334155">
            Documents are identical &#8212; no differences were detected.
        </div></div>`;
        return html;
    }

    // Summary bar
    const high    = findings.filter(f => f.severity === 'high').length;
    const medium  = findings.filter(f => f.severity === 'medium').length;
    const low     = findings.filter(f => f.severity === 'low').length;
    const changed = findings.filter(f => ['numeric-change','text-change','sign-change','policy-change','director-change','comparative-mismatch'].includes(f.category)).length;
    const added   = findings.filter(f => ['content-added','section-added'].includes(f.category)).length;
    const removed = findings.filter(f => ['content-removed','section-removed'].includes(f.category)).length;

    html += `<div style="background:#0D1B2A;border:1px solid #1E3A5F;border-radius:10px;padding:10px 14px;margin-bottom:14px">
        <div style="display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin-bottom:8px">
            <span style="color:#E2E8F0;font-weight:700;font-size:12px">${findings.length} system finding(s)</span>
            ${high   > 0 ? `<span style="color:#EF4444;font-weight:700;font-size:11px">&#9679; ${high} HIGH</span>` : ''}
            ${medium > 0 ? `<span style="color:#F59E0B;font-weight:700;font-size:11px">&#9679; ${medium} MEDIUM</span>` : ''}
            ${low    > 0 ? `<span style="color:#60A5FA;font-size:11px">&#9679; ${low} LOW</span>` : ''}
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
            ${changed  > 0 ? `<span style="background:#F59E0B18;border:1px solid #F59E0B44;color:#FDE68A;border-radius:6px;padding:3px 9px;font-size:10.5px;font-family:'IBM Plex Mono',monospace">~ ${changed} Changed</span>` : ''}
            ${added    > 0 ? `<span style="background:#10B98118;border:1px solid #10B98144;color:#6EE7B7;border-radius:6px;padding:3px 9px;font-size:10.5px;font-family:'IBM Plex Mono',monospace">+ ${added} Added</span>` : ''}
            ${removed  > 0 ? `<span style="background:#EF444418;border:1px solid #EF444444;color:#FCA5A5;border-radius:6px;padding:3px 9px;font-size:10.5px;font-family:'IBM Plex Mono',monospace">&#8722; ${removed} Removed</span>` : ''}
        </div>
    </div>`;

    html += renderAutoReviewFindings(findings);
    html += `</div>`;
    return html;
}

// #6 — Apply materiality filter: hide findings below threshold client-side
function applyMaterialityFilter() {
    const threshold = parseFloat(document.getElementById('materiality-input')?.value || 0) || 0;
    document.querySelectorAll('.sysreview-finding').forEach(el => {
        const abs = parseFloat(el.dataset.abschange || 0);
        // Show if: no threshold, no absChange (non-numeric), or above threshold
        const hide = threshold > 0 && abs > 0 && abs < threshold;
        el.style.display = hide ? 'none' : '';
    });
    // Also persist to server
    if (threshold > 0) {
        const fd = new FormData();
        fd.append('threshold', threshold);
        fetch('/api/materiality', { method:'POST', body: fd,
            headers:{'RequestVerificationToken': document.querySelector('input[name="__RequestVerificationToken"]')?.value || ''} });
    }
}

function renderAutoReviewFindings(findings) {
    if (!findings || findings.length === 0) return '';

    const catMeta = {
        'sign-change':          { label:'Sign Reversals',        color:'#F87171', icon:'&#8645;' },
        'policy-change':        { label:'Policy Changes',        color:'#FB923C', icon:'&#128196;' },
        'director-change':      { label:'Key Personnel',         color:'#FBBF24', icon:'&#128100;' },
        'section-added':        { label:'Pages Added',           color:'#34D399', icon:'+' },
        'section-removed':      { label:'Pages Removed',         color:'#F87171', icon:'&#8722;' },
        'numeric-change':       { label:'Numeric Changes',       color:'#A78BFA', icon:'#' },
        'comparative-mismatch': { label:'Comparative Mismatch',  color:'#F472B6', icon:'&#8596;' },
        'stale-reference':      { label:'Stale References',      color:'#FB923C', icon:'&#9875;' },
        'content-added':        { label:'Content Added',         color:'#34D399', icon:'+' },
        'content-removed':      { label:'Content Removed',       color:'#F87171', icon:'&#8722;' },
        'text-change':          { label:'Text Changes',          color:'#94A3B8', icon:'T' },
    };

    // #1 — Group by sectionGroup (outer), then category (inner)
    const sectionMap = {};
    findings.forEach(f => {
        const sec = f.sectionGroup || 'Other';
        const cat = f.category || 'text-change';
        if (!sectionMap[sec]) sectionMap[sec] = {};
        if (!sectionMap[sec][cat]) sectionMap[sec][cat] = [];
        sectionMap[sec][cat].push(f);
    });

    // Sort sections: "Other" last
    const sections = Object.keys(sectionMap).sort((a,b) =>
        a === 'Other' ? 1 : b === 'Other' ? -1 : a.localeCompare(b));

    let html = '';
    let globalIdx = 0;

    sections.forEach(sec => {
        const cats = sectionMap[sec];
        const secTotal = Object.values(cats).reduce((s,a) => s + a.length, 0);
        const secId = 'sec-' + sec.replace(/[^a-z0-9]/gi,'_');

        html += `<div style="margin-bottom:18px">
            <div onclick="toggleSection('${secId}')" style="display:flex;align-items:center;gap:8px;margin-bottom:8px;padding:5px 10px;background:#0D1B2A;border:1px solid #1E3A5F;border-radius:8px;cursor:pointer;user-select:none">
                <span id="${secId}-arrow" style="color:#7B4FFF;font-size:12px;transition:transform 0.2s">&#9660;</span>
                <span style="color:#C4B5FD;font-weight:700;font-size:12px;flex:1">${esc(sec)}</span>
                <span style="background:#7B4FFF22;color:#A78BFA;border-radius:999px;padding:1px 9px;font-size:10px;font-family:'IBM Plex Mono',monospace">${secTotal}</span>
            </div>
            <div id="${secId}">`;

        Object.keys(cats).forEach(cat => {
            const items = cats[cat];
            const meta  = catMeta[cat] || { label: cat, color:'#A78BFA', icon:'•' };

            html += `<div style="margin-bottom:10px;padding-left:8px;border-left:2px solid ${meta.color}44">
                <div style="display:flex;align-items:center;gap:6px;margin-bottom:6px">
                    <span style="background:${meta.color}22;color:${meta.color};border-radius:5px;padding:1px 7px;font-size:10.5px;font-weight:700;font-family:'IBM Plex Mono',monospace">${meta.icon} ${esc(meta.label)}</span>
                    <span style="color:#475569;font-size:10px">${items.length}</span>
                </div>`;

            items.forEach(f => {
                globalIdx++;
                const sev      = (f.severity || 'low').toLowerCase();
                const sevColor = sev === 'high' ? '#EF4444' : sev === 'medium' ? '#F59E0B' : '#60A5FA';
                const pairIdx  = typeof f.pairIndex === 'number' ? f.pairIndex : null;
                const signOff  = f.signOff || '';

                // #7 — Confidence badge
                const conf = typeof f.confidence === 'number' ? f.confidence : null;
                const confColor = conf === null ? '' : conf >= 80 ? '#10B981' : conf >= 60 ? '#F59E0B' : '#EF4444';
                const confBadge = conf !== null
                    ? `<span title="Confidence score: ${conf}/100" style="background:${confColor}22;border:1px solid ${confColor}55;color:${confColor};border-radius:999px;padding:1px 6px;font-size:9px;font-family:'IBM Plex Mono',monospace">${conf}%</span>`
                    : '';

                // Snapshot buttons
                const snapBtnStyle = 'border-radius:5px;padding:2px 7px;font-size:9.5px;cursor:pointer;font-family:\'IBM Plex Mono\',monospace;white-space:nowrap;border:1px solid';
                const snapBtn = pairIdx !== null ? [
                    f.pageAfs1 ? `<button onclick="jumpToPageSnapshot(${pairIdx},1)" style="${snapBtnStyle} #3B82F655;background:#0D213718;color:#93C5FD">AFS1·p${f.pageAfs1} &#8594;</button>` : '',
                    f.pageAfs2 ? `<button onclick="jumpToPageSnapshot(${pairIdx},2)" style="${snapBtnStyle} #10B98155;background:#06291818;color:#6EE7B7">AFS2·p${f.pageAfs2} &#8594;</button>` : '',
                ].filter(Boolean).join(' ') : '';

                const signOffButtons = '';

                // Before/After diff block
                let diffBlock = '';
                if (f.before && f.after && !['content-added','content-removed'].includes(f.category)) {
                    diffBlock = `<div style="margin-top:7px;border-radius:6px;overflow:hidden;font-family:'IBM Plex Mono',monospace;font-size:10.5px">
                        <div style="background:rgba(239,68,68,0.08);padding:4px 8px;color:#FCA5A5;border-left:3px solid #EF4444">&#8722; ${esc(f.before)}</div>
                        <div style="background:rgba(34,197,94,0.08);padding:4px 8px;color:#86EFAC;border-left:3px solid #10B981">+ ${esc(f.after)}</div>
                        ${typeof f.pctChange === 'number' && f.pctChange > 0
                            ? `<div style="background:rgba(251,191,36,0.06);padding:3px 8px;color:#FDE68A;font-size:10px">Change: ${f.pctChange}%${f.absChange ? ' | Abs: R ' + (f.absChange >= 1000000 ? (f.absChange/1000000).toFixed(2)+'M' : f.absChange >= 1000 ? (f.absChange/1000).toFixed(0)+'K' : f.absChange) : ''}</div>`
                            : ''}
                    </div>`;
                } else if (f.before && !f.after) {
                    diffBlock = `<div style="margin-top:6px;background:rgba(239,68,68,0.07);border-left:3px solid #EF4444;border-radius:5px;padding:4px 8px;font-family:'IBM Plex Mono',monospace;font-size:10.5px;color:#FCA5A5">&#8722; ${esc(f.before)}</div>`;
                } else if (!f.before && f.after) {
                    diffBlock = `<div style="margin-top:6px;background:rgba(34,197,94,0.07);border-left:3px solid #10B981;border-radius:5px;padding:4px 8px;font-family:'IBM Plex Mono',monospace;font-size:10.5px;color:#86EFAC">+ ${esc(f.after)}</div>`;
                }

                // Card container — data-abschange for materiality filter
                const cardOpacity = signOff === 'cleared' ? '0.45' : '1';
                const cardBorderExtra = signOff === 'escalate' ? `;box-shadow:0 0 0 1px #EF444488` : '';
                html += `<div class="sysreview-finding" data-abschange="${f.absChange || 0}"
                    style="background:#0F172A;border-left:3px solid ${sevColor};border-radius:7px;padding:9px 12px;margin-bottom:7px;opacity:${cardOpacity}${cardBorderExtra};${signOff==='cleared'?'text-decoration:line-through;text-decoration-color:#334155':''}"
                    >
                    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:4px">
                        <div style="color:#F1F5F9;font-weight:600;font-size:11.5px;line-height:1.4;flex:1;${signOff==='cleared'?'text-decoration:none':''}">${globalIdx}. ${esc(f.title || 'Untitled')}</div>
                        <div style="display:flex;gap:4px;align-items:center;flex-shrink:0">
                            <span style="background:${sevColor}22;border:1px solid ${sevColor}55;color:${sevColor};border-radius:999px;padding:1px 6px;font-size:9px;font-family:'IBM Plex Mono',monospace">${sev.toUpperCase()}</span>
                            ${confBadge}
                            ${snapBtn}
                        </div>
                    </div>
                    <div style="color:#94A3B8;font-size:11px;line-height:1.5">${esc(f.description || '')}</div>
                    ${diffBlock}
                    ${signOffButtons}
                </div>`;
            });

            html += `</div>`;  // cat block
        });

        html += `</div></div>`;  // sectionMap block + section wrapper
    });

    return html;
}

// Toggle section collapse (#1)
function toggleSection(secId) {
    const el    = document.getElementById(secId);
    const arrow = document.getElementById(secId + '-arrow');
    if (!el) return;
    const open = el.style.display !== 'none';
    el.style.display  = open ? 'none' : '';
    if (arrow) arrow.style.transform = open ? 'rotate(-90deg)' : '';
}

// #5 — Sign-off action: POST to /api/finding-signoff, update card visually
function findingSignOff(id, action, btn) {
    const fd = new FormData();
    fd.append('id', id);
    fd.append('action', action);
    const token = document.querySelector('input[name="__RequestVerificationToken"]')?.value || '';
    fetch('/api/finding-signoff', { method:'POST', body: fd, headers:{'RequestVerificationToken': token} })
        .then(r => r.ok ? r.json() : null)
        .then(() => {
            const card = document.getElementById('finding-card-' + id);
            if (!card) return;
            // Update visual state
            const isCleared   = action === 'cleared';
            const isEscalate  = action === 'escalate';
            card.style.opacity = isCleared ? '0.45' : '1';
            card.style.boxShadow = isEscalate ? '0 0 0 1px #EF444488' : '';
            // Update button active states within this card
            card.querySelectorAll('[data-action]').forEach(b => {
                const a = b.dataset.action;
                const active = a === action;
                const colors = { cleared:'#10B981', query:'#F59E0B', escalate:'#EF4444' };
                const c = colors[a] || '#64748B';
                b.style.borderColor = active ? c : c + '55';
                b.style.background  = active ? c + '22' : 'transparent';
                b.style.color       = active ? c : c;
            });
        })
        .catch(() => {});
}

function renderAiReviewTab(ai, aiRead) {
    const aiReady = ai && ai.status === 'ready';
    const readReady = aiRead && aiRead.status === 'ready';

    if (!aiReady && !readReady) {
        return `<div id="tab-ai" class="afs-tab-pane" style="display:none">
            ${renderAiUnavailableNotice(
                'Intelligent review could not be completed.',
                'All AI models returned errors. The document comparison is still available.',
                true)}
        </div>`;
    }

    let html = `<div id="tab-ai" class="afs-tab-pane" style="display:none">`;

    // ── Sub-tab selector ──────────────────────────────────────────────────────
    html += `<div style="display:flex;gap:6px;margin-bottom:12px;flex-wrap:wrap;align-items:center">
        <button onclick="aiTab('diff')" id="aibtn-diff"
            style="padding:4px 14px;border-radius:999px;border:1px solid #7B4FFF55;background:#7B4FFF22;color:#C4B5FD;font-size:11.5px;cursor:pointer;font-weight:700">
            Diff Analysis
        </button>
        <button onclick="aiTab('content')" id="aibtn-content"
            style="padding:4px 14px;border-radius:999px;border:1px solid rgba(148,163,184,0.25);background:transparent;color:#94A3B8;font-size:11.5px;cursor:pointer">
            Document Reading
        </button>
        <span style="margin-left:auto;color:#475569;font-size:10.5px;font-family:'IBM Plex Mono',monospace">
            Click a page reference to open snapshot
        </span>
    </div>`;

    // ══════════════════════════════════════════════════════════════════════════
    // PANEL 1: DIFF ANALYSIS (AiAssistResult)
    // ══════════════════════════════════════════════════════════════════════════
    html += `<div id="aipanel-diff">`;

    if (!aiReady) {
        html += renderAiUnavailableNotice(
            'Diff review could not be completed.',
            'The comparison result is still available. You can retry the AI analysis.',
            true);
    } else {
        const riskColor = (ai.overallRisk || '').toLowerCase() === 'high' ? '#EF4444'
            : (ai.overallRisk || '').toLowerCase() === 'medium' ? '#F59E0B' : '#10B981';

        html += `<div style="background:#0D1B2A;border:1px solid #7B4FFF44;border-radius:10px;padding:12px 14px;margin-bottom:12px">
            <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:8px">
                <span style="background:#7B4FFF22;border:1px solid #7B4FFF55;color:#C4B5FD;border-radius:999px;padding:2px 10px;font-size:11px;font-family:'IBM Plex Mono',monospace">
                    Intelligence Review Active
                </span>
                <span style="background:${riskColor}22;border:1px solid ${riskColor}55;color:${riskColor};border-radius:999px;padding:2px 10px;font-size:11px;font-family:'IBM Plex Mono',monospace">
                    Risk: ${esc((ai.overallRisk || 'unspecified').toUpperCase())}
                </span>
            </div>
            <div style="color:#E5E7EB;font-size:12.5px;line-height:1.6">${esc(ai.executiveSummary || 'No executive summary.')}</div>
        </div>`;

        // Highlight guidance
        if (Array.isArray(ai.highlightGuidance) && ai.highlightGuidance.length) {
            html += `<div style="margin:10px 0 6px;color:#FBBF24;font-weight:700;font-size:11.5px;letter-spacing:.05em">HIGHLIGHT GUIDANCE</div>`;
            ai.highlightGuidance.forEach(item => {
                html += `<div style="background:#111827;border:1px solid rgba(123,79,255,0.18);border-radius:8px;padding:8px 10px;margin-bottom:5px;color:#CBD5E1;font-size:12px">${esc(item)}</div>`;
            });
        }

        // Material findings
        html += `<div style="margin:14px 0 6px;color:#F59E0B;font-weight:700;font-size:11.5px;letter-spacing:.05em">MATERIAL FINDINGS</div>`;
        if (Array.isArray(ai.materialFindings) && ai.materialFindings.length) {
            ai.materialFindings.forEach((f, idx) => {
                const sev = (f.severity || '').toLowerCase();
                const isFalsePositive = (f.title || '').toLowerCase().includes('false positive');
                const sevColor = isFalsePositive ? '#475569'
                    : sev === 'high' ? '#EF4444'
                    : sev === 'medium' ? '#F59E0B'
                    : '#60A5FA';
                const pairIdx = typeof f.pairIndex === 'number' ? f.pairIndex : null;
                const pageRef = buildPageRefHtml(pairIdx, f.pageAfs1, f.pageAfs2);
                const nums = Array.isArray(f.changedNumbers) && f.changedNumbers.length
                    ? f.changedNumbers.slice(0, 6).join(', ') : '';

                // AFS 1 vs AFS 2 diff rows for this finding
                const diffRows = getFindingDiffRows(pairIdx, f.focusDiffOrdinals);
                let diffHtml = '';
                if (diffRows.length > 0) {
                    diffHtml = `<div style="margin-top:10px;border-top:1px solid rgba(255,255,255,0.07);padding-top:9px">
                        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:4px">
                            <div style="color:#94A3B8;font-size:10px;font-weight:700;letter-spacing:.08em">AFS 1</div>
                            <div style="color:#94A3B8;font-size:10px;font-weight:700;letter-spacing:.08em">AFS 2</div>
                        </div>`;
                    diffRows.forEach(d => {
                        const statusIs = d.status || '';
                        const l1bg = statusIs === 'removed' ? 'rgba(239,68,68,0.08)' : statusIs === 'changed' ? 'rgba(220,38,38,0.08)' : 'rgba(255,255,255,0.03)';
                        const l2bg = statusIs === 'added'   ? 'rgba(34,197,94,0.08)'  : statusIs === 'changed' ? 'rgba(34,197,94,0.08)'  : 'rgba(255,255,255,0.03)';
                        const l1 = d.wordDiff ? renderWordDiffLine(d.wordDiff, 'removed') : esc(d.line1 || '—');
                        const l2 = d.wordDiff ? renderWordDiffLine(d.wordDiff, 'added')   : esc(d.line2 || '—');
                        diffHtml += `<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:4px">
                            <div style="background:${l1bg};border-radius:5px;padding:5px 8px;font-size:11px;color:#E5E7EB;font-family:'IBM Plex Mono',monospace;line-height:1.5">${l1}</div>
                            <div style="background:${l2bg};border-radius:5px;padding:5px 8px;font-size:11px;color:#E5E7EB;font-family:'IBM Plex Mono',monospace;line-height:1.5">${l2}</div>
                        </div>`;
                    });
                    diffHtml += `</div>`;
                }

                // View Snapshot button
                const snapBtn = pairIdx !== null
                    ? `<button onclick="jumpToPageSnapshot(${pairIdx}, 1)"
                        style="background:#7B4FFF22;border:1px solid #7B4FFF66;color:#C4B5FD;border-radius:7px;padding:5px 12px;font-size:11px;cursor:pointer;font-family:'IBM Plex Mono',monospace;white-space:nowrap">
                        View Snapshot →
                       </button>` : '';

                html += `<div style="background:#111827;border-left:3px solid ${sevColor};border-radius:8px;padding:10px 12px;margin-bottom:8px">
                    <div style="display:flex;justify-content:space-between;gap:8px;align-items:flex-start;margin-bottom:6px">
                        <div style="color:#F8FAFC;font-weight:700;font-size:12.5px">${idx + 1}. ${esc(f.title || 'Untitled finding')}</div>
                        <span style="background:${sevColor}22;border:1px solid ${sevColor}55;color:${sevColor};border-radius:999px;padding:1px 8px;font-size:10px;font-family:'IBM Plex Mono',monospace;white-space:nowrap">${isFalsePositive ? 'FALSE +' : esc((f.severity || 'INFO').toUpperCase())}</span>
                    </div>
                    <div style="display:flex;flex-wrap:wrap;align-items:center;gap:5px;margin-bottom:7px">
                        ${pageRef}
                        ${snapBtn}
                    </div>
                    <div style="color:#CBD5E1;font-size:12px;margin-bottom:5px">${esc(f.auditRationale || '')}</div>
                    ${f.recommendedProcedure ? `<div style="color:#93C5FD;font-size:12px;margin-bottom:5px">${esc(f.recommendedProcedure)}</div>` : ''}
                    ${nums ? `<div style="color:#94A3B8;font-size:10.5px;font-family:'IBM Plex Mono',monospace;margin-top:3px">Values: ${esc(nums)}</div>` : ''}
                    ${diffHtml}
                </div>`;
            });
        } else {
            html += `<div class="afs-empty-state">No material findings returned.</div>`;
        }

        // Page priorities
        html += `<div style="margin:14px 0 6px;color:#A78BFA;font-weight:700;font-size:11.5px;letter-spacing:.05em">PAGE REVIEW PRIORITIES</div>`;
        if (Array.isArray(ai.pageFocus) && ai.pageFocus.length) {
            ai.pageFocus.forEach((f, i) => {
                const pColor = (f.priority || '').toLowerCase() === 'high' ? '#EF4444'
                    : (f.priority || '').toLowerCase() === 'medium' ? '#F59E0B' : '#60A5FA';
                const ordinals = Array.isArray(f.focusDiffOrdinals) && f.focusDiffOrdinals.length
                    ? f.focusDiffOrdinals.join(', ') : '—';
                const fpIdx = f.pairIndex ?? null;
                const pageRef = buildPageRefHtml(fpIdx, f.pageAfs1, f.pageAfs2);
                const fpSnap = fpIdx !== null
                    ? `<button onclick="jumpToPageSnapshot(${fpIdx}, 1)"
                        style="background:#7B4FFF22;border:1px solid #7B4FFF55;color:#C4B5FD;border-radius:7px;padding:3px 10px;font-size:10.5px;cursor:pointer;font-family:'IBM Plex Mono',monospace">
                        View Snapshot →</button>` : '';
                html += `<div style="background:${i % 2 === 0 ? '#162347' : '#111827'};border-radius:7px;padding:8px 10px;margin-bottom:5px;display:flex;align-items:flex-start;gap:10px;flex-wrap:wrap">
                    <span style="color:${pColor};font-weight:700;font-size:11px;font-family:'IBM Plex Mono',monospace;min-width:54px">${esc((f.priority || '—').toUpperCase())}</span>
                    <div style="flex:1;min-width:180px">
                        <div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:4px">${pageRef} ${fpSnap}</div>
                        <div style="color:#CBD5E1;font-size:11.5px">${esc(f.why || '')}</div>
                    </div>
                    <span style="color:#475569;font-size:10.5px;font-family:'IBM Plex Mono',monospace;white-space:nowrap">diffs: ${esc(ordinals)}</span>
                </div>`;
            });
        } else {
            html += `<div class="afs-empty-state">No page priorities returned.</div>`;
        }

        if (Array.isArray(ai.limitations) && ai.limitations.length) {
            html += `<div style="margin:12px 0 4px;color:#475569;font-size:10.5px;font-weight:700;letter-spacing:.05em">NOTES</div>`;
            ai.limitations.forEach(item => {
                html += `<div style="color:#475569;font-size:10.5px;margin-bottom:3px">${esc(item)}</div>`;
            });
        }
    }

    html += `</div>`; // end aipanel-diff

    // ══════════════════════════════════════════════════════════════════════════
    // PANEL 2: CONTENT READING (AiContentReadResult)
    // ══════════════════════════════════════════════════════════════════════════
    html += `<div id="aipanel-content" style="display:none">`;

    if (!readReady) {
        html += renderAiUnavailableNotice(
            'Content reading could not be completed.',
            'The comparison result is still available. You can retry the AI analysis.',
            true);
    } else {
        const riskClr = (aiRead.riskLevel || '').toLowerCase() === 'high' ? '#EF4444'
            : (aiRead.riskLevel || '').toLowerCase() === 'medium' ? '#F59E0B' : '#10B981';

        // Overall assessment
        html += `<div style="background:#0D1B2A;border:1px solid #7B4FFF44;border-radius:10px;padding:12px 14px;margin-bottom:12px">
            <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px">
                <span style="background:#7B4FFF22;border:1px solid #7B4FFF55;color:#C4B5FD;border-radius:999px;padding:2px 10px;font-size:11px;font-family:'IBM Plex Mono',monospace">Document Reading</span>
                <span style="background:${riskClr}22;border:1px solid ${riskClr}55;color:${riskClr};border-radius:999px;padding:2px 10px;font-size:11px;font-family:'IBM Plex Mono',monospace">Risk: ${esc((aiRead.riskLevel || '—').toUpperCase())}</span>
            </div>
            <div style="color:#E5E7EB;font-size:12.5px;line-height:1.6;margin-bottom:8px">${esc(aiRead.overallAssessment || 'No assessment returned.')}</div>
            ${aiRead.auditConclusion ? `<div style="border-top:1px solid #1E3A5F;padding-top:8px;color:#93C5FD;font-size:12px;line-height:1.5">${esc(aiRead.auditConclusion)}</div>` : ''}
        </div>`;

        // Section-level diffs
        if (Array.isArray(aiRead.sectionDiffs) && aiRead.sectionDiffs.length) {
            html += `<div style="margin:10px 0 6px;color:#F59E0B;font-weight:700;font-size:11.5px;letter-spacing:.05em">SECTION-BY-SECTION CHANGES</div>`;
            aiRead.sectionDiffs.forEach((s, si) => {
                const sev = (s.severity || '').toLowerCase();
                const sevClr = sev === 'high' ? '#EF4444' : sev === 'medium' ? '#F59E0B' : '#10B981';
                const ctClr  = s.changeType === 'numeric' ? '#FBBF24' : s.changeType === 'structural' ? '#A78BFA' : '#CBD5E1';
                const ref = resolveSnapshotRef(s.pairIndex, s.pageAfs1, s.pageAfs2, s.sectionName);
                html += `<div style="background:#111827;border-left:3px solid ${sevClr};border-radius:8px;padding:10px 12px;margin-bottom:7px">
                    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:5px">
                        <div style="color:#F8FAFC;font-weight:700;font-size:12.5px">${esc(s.sectionName)}</div>
                        <div style="display:flex;gap:5px;flex-shrink:0">
                            <span style="color:${ctClr};background:${ctClr}18;border:1px solid ${ctClr}44;border-radius:999px;padding:1px 7px;font-size:10px;font-family:'IBM Plex Mono',monospace">${esc(s.changeType || '—')}</span>
                            <span style="color:${sevClr};background:${sevClr}18;border:1px solid ${sevClr}44;border-radius:999px;padding:1px 7px;font-size:10px;font-family:'IBM Plex Mono',monospace">${esc(s.severity || '—')}</span>
                        </div>
                    </div>
                    <div style="color:#CBD5E1;font-size:12px;margin-bottom:5px">${esc(s.changeDescription || '')}</div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:11px">
                        <div style="background:#0D1B2A;border-radius:5px;padding:5px 8px"><div style="color:#64748B;font-size:10px;margin-bottom:2px">AFS 1</div><div style="color:#94A3B8">${esc(s.afs1Summary || '—')}</div></div>
                        <div style="background:#0D1B2A;border-radius:5px;padding:5px 8px"><div style="color:#64748B;font-size:10px;margin-bottom:2px">AFS 2</div><div style="color:#94A3B8">${esc(s.afs2Summary || '—')}</div></div>
                    </div>
                    ${Array.isArray(s.changedValues) && s.changedValues.length ? `<div style="margin-top:6px;display:flex;flex-wrap:wrap;gap:4px">${s.changedValues.map(v => `<span style="background:#1E3A5F;color:#FBBF24;border-radius:5px;padding:2px 6px;font-size:10.5px;font-family:'IBM Plex Mono',monospace">${esc(v)}</span>`).join('')}</div>` : ''}
                    ${snapshotRefHtml(ref, ref ? undefined : si)}
                </div>`;
            });
        }

        // Line item changes — table with page ref column and snapshot button per row
        if (Array.isArray(aiRead.lineItemChanges) && aiRead.lineItemChanges.length) {
            html += `<div style="margin:14px 0 6px;color:#A78BFA;font-weight:700;font-size:11.5px;letter-spacing:.05em">NUMERIC LINE ITEMS</div>`;
            html += `<table style="width:100%;border-collapse:collapse;font-size:11.5px">
                <thead><tr style="background:#1E2F5A">
                    <th style="${thStyle}">LINE ITEM</th><th style="${thStyle}">AFS 1</th><th style="${thStyle}">AFS 2</th><th style="${thStyle}">CHANGE</th><th style="${thStyle}">PAGE</th>
                </tr></thead><tbody>`;
            aiRead.lineItemChanges.forEach((l, i) => {
                const bg = i % 2 === 0 ? '#162347' : '#111827';
                const chgClr = (l.pctChange || '').startsWith('-') ? '#EF4444' : '#10B981';
                const ref = resolveSnapshotRef(l.pairIndex, l.pageAfs1, l.pageAfs2, l.lineItem);
                const pageCell = ref
                    ? `<span style="color:#64748B;font-size:10px;font-family:'IBM Plex Mono',monospace;display:block">p.${ref.pageAfs1||'?'}↔p.${ref.pageAfs2||'?'}</span>
                       <button onclick="jumpToPageSnapshot(${ref.pairIndex},1)" style="margin-top:2px;padding:1px 8px;font-size:10px;border-radius:999px;border:1px solid #7B4FFF55;background:#7B4FFF18;color:#C4B5FD;cursor:pointer">View →</button>`
                    : `<span style="color:#334155;font-size:10px">—</span>`;
                html += `<tr style="background:${bg}">
                    <td style="${tdStyle};font-weight:600">${esc(l.lineItem)}</td>
                    <td style="${tdStyle};font-family:'IBM Plex Mono',monospace">${esc(l.valueInAfs1 || '—')}</td>
                    <td style="${tdStyle};font-family:'IBM Plex Mono',monospace">${esc(l.valueInAfs2 || '—')}</td>
                    <td style="${tdStyle};color:${chgClr};font-family:'IBM Plex Mono',monospace;font-weight:700">${esc(l.pctChange || '—')}</td>
                    <td style="${tdStyle}">${pageCell}</td>
                </tr>`;
            });
            html += `</tbody></table>`;
        }

        // New / removed disclosures & policy changes
        const disclosures = [
            { title: 'NEW DISCLOSURES', items: aiRead.newDisclosures, color: '#10B981' },
            { title: 'REMOVED DISCLOSURES', items: aiRead.removedDisclosures, color: '#EF4444' },
            { title: 'POLICY CHANGES', items: aiRead.policyChanges, color: '#F59E0B' },
        ];
        disclosures.forEach(({ title, items, color }) => {
            if (!Array.isArray(items) || !items.length) return;
            html += `<div style="margin:14px 0 6px;color:${color};font-weight:700;font-size:11.5px;letter-spacing:.05em">${title}</div>`;
            items.forEach(item => {
                html += `<div style="background:#111827;border-left:3px solid ${color};border-radius:6px;padding:7px 10px;margin-bottom:5px;color:#CBD5E1;font-size:12px">${esc(item)}</div>`;
            });
        });

        if (Array.isArray(aiRead.limitations) && aiRead.limitations.length) {
            html += `<div style="margin:12px 0 4px;color:#475569;font-size:10.5px;font-weight:700;letter-spacing:.05em">NOTES</div>`;
            aiRead.limitations.forEach(item => {
                html += `<div style="color:#475569;font-size:10.5px;margin-bottom:3px">${esc(item)}</div>`;
            });
        }
    }

    html += `</div>`; // end aipanel-content
    html += `</div>`; // end tab-ai
    return html;
}

// Resolves a pairIndex from optional page numbers, falling back to a text search
// through already-loaded pageDiffs. Returns { pairIndex, pageAfs1, pageAfs2 } or null.
function resolveSnapshotRef(pairIndex, pageAfs1, pageAfs2, searchText) {
    const data = window._cmpData;
    if (!data || !data.pageDiffs) return null;

    // Direct pairIndex provided
    if (typeof pairIndex === 'number') {
        const p = data.pageDiffs.find(d => (d.pairIndex ?? d.k) === pairIndex);
        if (p) return { pairIndex, pageAfs1: p.pageAfs1, pageAfs2: p.pageAfs2 };
    }

    // Match by page numbers
    if (pageAfs1 || pageAfs2) {
        const p = data.pageDiffs.find(d =>
            (pageAfs1 && d.pageAfs1 === pageAfs1) ||
            (pageAfs2 && d.pageAfs2 === pageAfs2));
        if (p) return { pairIndex: p.pairIndex ?? p.k, pageAfs1: p.pageAfs1, pageAfs2: p.pageAfs2 };
    }

    // Fallback: text search in already-loaded diffs
    if (searchText) {
        const lower = searchText.toLowerCase().trim().substring(0, 40);
        for (const p of data.pageDiffs) {
            if (!Array.isArray(p.diffs)) continue;
            const found = p.diffs.some(d =>
                (d.line1 || '').toLowerCase().includes(lower) ||
                (d.line2 || '').toLowerCase().includes(lower));
            if (found) return { pairIndex: p.pairIndex ?? p.k, pageAfs1: p.pageAfs1, pageAfs2: p.pageAfs2 };
        }
    }
    return null;
}

// Renders a "View Snapshot →" button for Document Reading section diffs.
// Always shows the button — uses the best available ref or falls back to pairIndex 0.
function snapshotRefHtml(ref, fallbackPairIndex) {
    const r = ref || null;
    const pairIdx = r?.pairIndex ?? fallbackPairIndex ?? 0;
    const p1 = r?.pageAfs1 ? `AFS 1 p.${r.pageAfs1}` : '';
    const p2 = r?.pageAfs2 ? `AFS 2 p.${r.pageAfs2}` : '';
    const pageLabel = [p1, p2].filter(Boolean).join(' ↔ ');
    return `<div style="display:flex;align-items:center;gap:8px;margin-top:8px;flex-wrap:wrap;border-top:1px solid #1E3A5F;padding-top:7px">
        ${pageLabel ? `<span style="color:#64748B;font-size:10.5px;font-family:'IBM Plex Mono',monospace">${esc(pageLabel)}</span>` : ''}
        <button onclick="jumpToPageSnapshot(${pairIdx},1)"
            style="padding:3px 12px;font-size:11px;border-radius:999px;border:1px solid #7B4FFF88;background:#7B4FFF28;color:#C4B5FD;cursor:pointer;font-family:'IBM Plex Mono',monospace;font-weight:600;letter-spacing:.02em">
            View Snapshot →
        </button>
    </div>`;
}

function aiTab(name) {
    ['diff','content','sysreview'].forEach(n => {
        const panel = document.getElementById('aipanel-' + n);
        const btn   = document.getElementById('aibtn-' + n);
        if (!panel || !btn) return;
        const active = n === name;
        panel.style.display = active ? '' : 'none';
        btn.style.background = active ? '#7B4FFF22' : 'transparent';
        btn.style.color      = active ? '#C4B5FD' : '#94A3B8';
        btn.style.border     = active ? '1px solid #7B4FFF55' : '1px solid rgba(148,163,184,0.25)';
        btn.style.fontWeight = active ? '700' : '400';
    });
}

function renderAiRunBanner(ai) {
    if (!ai) {
        return `<div class="afs-info afs-info-violet" style="margin-bottom:10px">
            AI assist data was not returned for this comparison run. The system is using deterministic comparison and highlighting only.
        </div>`;
    }

    if (ai.status !== 'ready') {
        return `<div class="afs-info afs-info-violet" style="margin-bottom:10px">
            <div><b>Review currently unavailable.</b> Please try again later.</div>
            <div style="margin-top:4px;color:#94A3B8">The comparison result is still available — only the intelligent materiality review could not be completed at this time.</div>
        </div>`;
    }

    const findings = Array.isArray(ai.materialFindings) ? ai.materialFindings.length : 0;
    const focusedPages = Array.isArray(ai.pageFocus) ? ai.pageFocus.length : 0;
    return `<div class="afs-info afs-info-violet" style="margin-bottom:10px">
        <div><b>Intelligent review complete.</b></div>
        <div>Prioritised ${focusedPages} page pair(s) and identified ${findings} material finding(s).</div>
        <div style="margin-top:4px;color:#94A3B8">Highlights marked in red indicate high-priority changes. Orange indicates medium priority. Yellow indicates standard changes.</div>
    </div>`;
}

function renderAiUnavailableNotice(title, detail, showRetry = false) {
    const retryBtn = showRetry
        ? `<div style="margin-top:10px">
               <button onclick="runAiAnalysis(false)" style="background:#7B4FFF;color:#fff;border:none;padding:6px 16px;border-radius:6px;font-size:12px;cursor:pointer;font-weight:600">
                   Retry AI Analysis
               </button>
           </div>`
        : '';
    return `<div class="afs-info afs-info-violet" style="margin-bottom:10px">
        <div><b>${esc(title)}</b></div>
        <div style="margin-top:4px;color:#94A3B8">${esc(detail)}</div>
        ${retryBtn}
    </div>`;
}

/**
 * Renders word-diff tokens.
 * Tag values from LineComparatorService: "same" | "removed" | "added"
 * Reference: Python LineComparator._word_diff()
 */
function renderWordDiff(wordDiff) {
    if (!wordDiff || !wordDiff.length) return '';
    let out = '';
    wordDiff.forEach(({ w, t }) => {
        const word = esc(w);
        if (t === 'same') {
            out += word + ' ';
        } else if (t === 'removed') {
            out += `<span style="background:#EF444433;color:#FCA5A5;text-decoration:line-through;padding:0 2px">${word}</span> `;
        } else if (t === 'added') {
            out += `<span style="background:#3B82F633;color:#93C5FD;font-weight:700;padding:0 2px">${word}</span> `;
        }
    });
    return out;
}

function renderDiffLine(d) {
    if (d.status === 'same-different-page') {
        return `<div class="diff-line diff-cross-page">
            <span style="display:inline-flex;align-items:center;gap:6px;background:#0F172A;border:1px solid rgba(59,130,246,0.35);color:#93C5FD;border-radius:999px;padding:2px 8px;font-size:10px;font-family:'IBM Plex Mono',monospace">Different Page</span>
            <span class="diff-text">${esc(d.line1 || d.line2)}</span>
            ${d.crossPageNote ? `<div style="margin-top:4px;color:#94A3B8;font-size:10px">${esc(d.crossPageNote)}</div>` : ''}
        </div>`;
    }

    const base = d.status === 'changed'
        ? (renderWordDiff(d.wordDiff) || esc(d.line1) + ' -> ' + esc(d.line2))
        : d.status === 'added'
            ? `<span style="color:#BAE6FD">${esc(d.line2)}</span>`
            : `<span style="color:#FECACA">${esc(d.line1)}</span>`;

    if (!d.crossPageNote)
        return base;

    return `${base}<div style="margin-top:4px;color:#94A3B8;font-size:10px">${esc(d.crossPageNote)}</div>`;
}

/**
 * Renders an auditor comment box for a given key.
 * Saves via POST /comment (AJAX).
 */
function commentBox(key, label) {
    const safeKey   = esc(key);
    const safeLabel = esc(label);
    const inputId   = 'cmt_' + key.replace(/[^a-z0-9]/gi, '_');
    return `
    <div style="margin:4px 0;padding:6px;border:1px solid #1e3a5f;border-radius:5px">
        <div style="color:#FBBF24;font-size:10px;font-weight:700;font-family:'IBM Plex Mono',monospace;margin-bottom:3px">
            ✏ AUDITOR NOTE — ${safeLabel}
        </div>
        <textarea id="${inputId}" data-key="${safeKey}" class="afs-input afs-textarea" rows="2" placeholder="Add auditor comment / finding note..." style="width:99%;font-size:12px;"></textarea>
        <div style="display:flex;align-items:center;gap:8px;margin-top:4px">
            <button class="afs-btn" style="background:#1e3a5f;color:#93C5FD;padding:3px 12px;font-size:11px" onclick="saveComment('${inputId}','${safeKey}')">Save Comment</button>
            <span id="cmt_ok_${inputId}" style="color:#00C9A7;font-size:11px;font-family:'IBM Plex Mono',monospace"></span>
        </div>
    </div>`;

    html += renderAiRunBanner(ai);
}

function getPageDiffByRawIndex(data, rawIndex) {
    if (!data || !data.pageDiffs || !data.pageDiffs.length) return null;
    return data.pageDiffs.find(p => (p.k ?? -1) === rawIndex) || data.pageDiffs[rawIndex] || null;
}

async function ensurePageDiffDetails(p) {
    if (!p) return null;
    if (Array.isArray(p.diffs)) return p;

    const res = await fetch(`/api/page-diff-detail?pairIndex=${encodeURIComponent(p.k ?? 0)}`, {
        cache: 'no-store'
    }).then(r => r.json());

    if (!res.success) {
        throw new Error(res.message || 'Could not load page diff detail.');
    }

    p.diffs = Array.isArray(res.diffs) ? res.diffs : [];
    if (typeof res.pairIndex === 'number') {
        p.k = res.pairIndex;
    }
    return p;
}

async function showPageDiff() {
    const data  = window._cmpData;
    if (!data) return;
    const k     = parseInt(document.getElementById('pairSelect').value, 10);
    const p     = getPageDiffByRawIndex(data, k);
    const wrap  = document.getElementById('pair-detail');
    if (!p) { wrap.innerHTML = ''; return; }

    wrap.innerHTML = '<div style="color:#A78BFA;font-family:\'IBM Plex Mono\',monospace;font-size:11px;padding:6px">Loading page diff detail...</div>';

    try {
        await ensurePageDiffDetails(p);
    } catch (e) {
        wrap.innerHTML = `<div style="color:#FCA5A5;font-size:12px;padding:8px">Error: ${esc(e.message)}</div>`;
        return;
    }

    const pct  = p.pctSame;
    const fc   = pct === 100 ? '#10B981' : pct >= 70 ? '#F59E0B' : '#EF4444';
    const lbl  = `AFS1 p${p.pageAfs1} ↔ ${p.pageAfs2 ? 'AFS2 p'+p.pageAfs2 : 'UNMATCHED'}`;

    let html = `<div style="background:#111827;border-radius:7px;padding:8px 12px;font-family:'IBM Plex Mono',monospace;font-size:11px;color:#CBD5E1;margin-bottom:8px">
        <b style="color:#A78BFA">Pair ${(p.k ?? k)+1}</b>  |
        <span style="color:#7B4FFF"> ${esc(lbl)}</span>  |
        <span style="color:#10B981">✓ ${p.same} same</span>
        <span style="color:#F59E0B"> ~ ${p.changed} changed</span>
        <span style="color:#3B82F6"> + ${p.added} added</span>
        <span style="color:#EF4444"> − ${p.removed} removed</span>  |
        <span style="color:${fc}">${pct}% same</span>  |
        <span style="color:#64748B">align: ${p.alignSim.toFixed(3)}</span>
    </div>`;

    const aiFocus = getAiPageFocus(data.aiAssist, p.pairIndex ?? (p.k ?? k));
    if (aiFocus) {
        const ordinals = Array.isArray(aiFocus.focusDiffOrdinals) && aiFocus.focusDiffOrdinals.length
            ? ` Focus diffs: ${aiFocus.focusDiffOrdinals.join(', ')}.`
            : '';
        html += `<div class="afs-info afs-info-violet" style="margin-bottom:8px"><b>AI focus:</b> ${esc(aiFocus.why || 'Review this page pair closely.')}${esc(ordinals)}</div>`;
    }

    if (p.diffs && p.diffs.length) {
        html += `<div style="margin:8px 0 3px;color:#F59E0B;font-weight:700;font-size:12px">CHANGES ON THIS PAGE PAIR</div>
        <table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:#1E2F5A">
            <th style="${thStyle};width:32px"></th>
            <th style="${thStyle}">DIFF (red=removed, blue=added in AFS2)</th>
            <th style="${thStyle};width:100px">NUM</th>
        </tr></thead><tbody>`;
        p.diffs.forEach((d, j) => {
            const sym = (d.status === 'changed' ? '~' : d.status === 'added' ? '+' : '−') + (j+1);
            const bgr = d.status === 'changed' ? '#1C1400' : d.status === 'added' ? '#081424' : '#1A0808';
            const symCol = d.status === 'changed' ? '#F59E0B' : d.status === 'added' ? '#3B82F6' : '#EF4444';
            const content = renderDiffLine(d);
            const nd = d.numDiff && d.numDiff.length ? esc(d.numDiff.join(', ')) : '—';
            html += `<tr style="background:${bgr};vertical-align:top">
                <td style="${tdStyle};color:${symCol};font-family:'IBM Plex Mono',monospace;white-space:nowrap">${sym}</td>
                <td style="${tdStyle};font-family:'IBM Plex Mono',monospace;font-size:11px">${content}</td>
                <td style="${tdStyle};color:#FBBF24;font-family:'IBM Plex Mono',monospace">${nd}</td>
            </tr>`;
        });
        html += '</tbody></table>';
    } else {
        html += '<p style="color:#10B981;font-size:12px">✓ No changes on this page pair.</p>';
    }

    // Page-pair comment box
    html += commentBox('page:' + (p.k ?? k), lbl + ' findings');
    wrap.innerHTML = html;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7b · PAGE VISUAL VIEWER — renders PDF page images via /api/snapshot
// Reference: Python _show_page_snapshot() with Poppler pdftoppm
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7c · PAGE SNAPSHOT (v2 — white-card layout with AI focus banner)
// ─────────────────────────────────────────────────────────────────────────────

// Step to the previous (-1) or next (+1) page pair and load snapshot automatically.
function snapNav(dir) {
    const sel = document.getElementById('visualPairSelect');
    if (!sel || !sel.options.length) return;
    const cur = parseInt(sel.value, 10);
    const opts = Array.from(sel.options).map(o => parseInt(o.value, 10));
    const curIdx = opts.indexOf(cur);
    const nextIdx = Math.max(0, Math.min(opts.length - 1, curIdx + dir));
    if (nextIdx === curIdx) return;         // already at boundary
    sel.value = String(opts[nextIdx]);
    showSnapshot();
}

async function showSnapshot() {
    const sel = document.getElementById('visualPairSelect');
    if (!sel) return;

    const k = parseInt(sel.value, 10);

    // Update counter & prev/next disabled state
    const opts = Array.from(sel.options);
    const curIdx = opts.findIndex(o => parseInt(o.value, 10) === k);
    const total  = opts.length;
    const counter = document.getElementById('snap-counter');
    const prevBtn = document.getElementById('snap-prev');
    const nextBtn = document.getElementById('snap-next');
    if (counter) counter.textContent = total > 0 ? `Page pair ${curIdx + 1} of ${total}` : '';
    if (prevBtn) prevBtn.disabled = curIdx <= 0;
    if (nextBtn) nextBtn.disabled = curIdx >= total - 1;
    const statusEl = document.getElementById('snap-status');
    const wrap = document.getElementById('snap-results');

    if (statusEl) statusEl.textContent = 'Loading...';
    if (wrap) {
        wrap.innerHTML = '<div style="color:#A78BFA;font-family:\'IBM Plex Mono\',monospace;font-size:11px;padding:6px">Rendering page snapshots with highlights...</div>';
    }

    try {
        const snapshotNonce = window._cmpData?._snapshotNonce || Date.now();
        const fetchSnapshot = side =>
            fetch(`/api/snapshot?pairIndex=${k}&afsNum=${side}&v=${snapshotNonce}`, { cache: 'no-store' })
                .then(r => r.json());

        const [r1, r2] = await Promise.all([
            fetchSnapshot(1),
            fetchSnapshot(2),
        ]);

        if (statusEl) statusEl.textContent = '';

        if (!r1.success && !r2.success) {
            wrap.innerHTML = `<div style="background:#1C1A00;border:1px solid #F59E0B;color:#FCD34D;padding:10px 14px;border-radius:7px;font-size:12px;font-family:'IBM Plex Mono',monospace">
                ${esc(r1.message || r2.message || 'Snapshots unavailable.')}<br>
                <span style="color:#94A3B8;font-size:11px">The text-based page diff is still available in the Page Map tab.</span>
            </div>`;
            return;
        }

        const pct = r1.pctSame ?? r2.pctSame ?? 0;
        const fc = pct === 100 ? '#10B981' : pct >= 70 ? '#F59E0B' : '#EF4444';
        const p = getPageDiffByRawIndex(window._cmpData, k);
        if (p) {
            await ensurePageDiffDetails(p);
        }
        const resolvedPairIndex = p?.k ?? r1.pairIndex ?? r2.pairIndex ?? k;
        const pg1 = (r1.pageAfs1 ?? r2.pageAfs1 ?? 0) + 1;
        const pg2Ref = r1.pageAfs2 ?? r2.pageAfs2;
        const pg2 = pg2Ref != null ? pg2Ref + 1 : null;
        const relationLabel = p?.secondaryMatch ? 'continuation match' : 'content match';
        const cardStyle = 'background:#FFFFFF;border:1px solid rgba(148,163,184,0.28);border-radius:12px;padding:10px;box-shadow:0 10px 24px rgba(15,23,42,0.12);min-width:0;width:100%;overflow:hidden';
        const imgStyle = 'display:block;width:100%;max-width:100%;height:auto;background:#fff;border-radius:8px;border:1px solid rgba(148,163,184,0.24);image-rendering:auto;object-fit:contain;object-position:top center';
        const linkStyle = 'display:inline-flex;align-items:center;gap:6px;padding:7px 10px;border-radius:999px;border:1px solid rgba(148,163,184,0.24);background:rgba(15,23,42,0.72);color:#E2E8F0;font-size:11px;font-family:\'IBM Plex Mono\',monospace;text-decoration:none';

        const opts2 = document.getElementById('visualPairSelect')
            ? Array.from(document.getElementById('visualPairSelect').options) : [];
        const curIdx2 = opts2.findIndex(o => parseInt(o.value, 10) === k);
        const total2  = opts2.length;

        let html = `<div style="background:#111827;border-radius:7px;padding:8px 12px;font-family:'IBM Plex Mono',monospace;font-size:11px;color:#CBD5E1;margin-bottom:8px;display:flex;align-items:center;flex-wrap:wrap;gap:10px">
            <div style="flex:1;min-width:0">
                <b style="color:#A78BFA">Pair ${resolvedPairIndex + 1}</b>
                &nbsp;|&nbsp; AFS 1 p${pg1} ↔ AFS 2 p${pg2 ?? '?'}
                &nbsp;|&nbsp; <span style="color:#10B981">✓${r1.same ?? 0}</span>
                <span style="color:#F59E0B"> ~${r1.changed ?? 0}</span>
                <span style="color:#3B82F6"> +${r1.added ?? 0}</span>
                <span style="color:#EF4444"> −${r1.removed ?? 0}</span>
                &nbsp;|&nbsp; <span style="color:${fc}">${pct}% same</span>
                &nbsp;|&nbsp; <span style="color:#475569">${curIdx2 + 1} / ${total2}</span>
            </div>
            <div style="display:flex;gap:6px;flex-shrink:0">
                <button onclick="snapNav(-1)" ${curIdx2 <= 0 ? 'disabled' : ''}
                    style="padding:3px 10px;background:#1E2F5A;color:#A78BFA;border:1px solid #7B4FFF44;border-radius:6px;font-size:12px;cursor:pointer;opacity:${curIdx2 <= 0 ? 0.4 : 1}">
                    ← Prev
                </button>
                <button onclick="snapNav(1)" ${curIdx2 >= total2 - 1 ? 'disabled' : ''}
                    style="padding:3px 10px;background:#1E2F5A;color:#A78BFA;border:1px solid #7B4FFF44;border-radius:6px;font-size:12px;cursor:pointer;opacity:${curIdx2 >= total2 - 1 ? 0.4 : 1}">
                    Next →
                </button>
            </div>
        </div>`;

        const aiFocus = getAiPageFocus(window._cmpData?.aiAssist, resolvedPairIndex);
        if (aiFocus) {
            const ordinals = Array.isArray(aiFocus.focusDiffOrdinals) && aiFocus.focusDiffOrdinals.length
                ? ` Focus diffs: ${aiFocus.focusDiffOrdinals.join(', ')}.`
                : '';
            html += `<div class="afs-info afs-info-violet" style="margin-bottom:10px"><b>AI focus:</b> ${esc(aiFocus.why || 'Review this page pair closely.')}${esc(ordinals)}</div>`;
        }

        const issueCount = (r1.changed ?? 0) + (r1.added ?? 0) + (r1.removed ?? 0);
        const isMajor   = issueCount > 2 || pct < 95;
        const diffNote  = issueCount === 0
            ? '✓ No differences detected on this page pair.'
            : pct >= 98
                ? `~ ${issueCount} minor difference${issueCount === 1 ? '' : 's'} detected (≥98% identical). May be whitespace or punctuation.`
                : `${issueCount} difference${issueCount === 1 ? '' : 's'} detected. Pages are paired by ${relationLabel}.`;
        const noteColor = issueCount === 0 ? '#10B981' : pct >= 98 ? '#F59E0B' : '#94A3B8';

        html += `<div style="color:${noteColor};font-size:11px;font-family:'IBM Plex Mono',monospace;margin:0 0 10px">
            ${esc(diffNote)}
        </div>`;

        html += `<div style="width:100%;overflow-x:auto;padding-bottom:4px">
            <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;align-items:start;width:100%;min-width:0">`;

        html += r1.success && r1.b64
            ? `<section style="${cardStyle}">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px;min-width:0">
                    <div style="min-width:0;flex:1 1 auto">
                        <div style="font-size:12px;color:#0F172A;font-weight:700">AFS 1 - page ${pg1}</div>
                        <div style="font-size:10px;color:#F59E0B;font-family:'IBM Plex Mono',monospace">yellow = changed or removed content</div>
                    </div>
                    <div style="display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end;flex:0 0 auto">
                        ${r1.pdfUrl ? `<a href="${r1.pdfUrl}" target="_blank" rel="noopener" style="${linkStyle}">Open PDF</a>` : ''}
                        <a href="data:image/png;base64,${r1.b64}" target="_blank" rel="noopener" style="${linkStyle}">Open image</a>
                    </div>
                </div>
                <img src="data:image/png;base64,${r1.b64}" alt="AFS 1 page ${pg1}" style="${imgStyle}"/>
            </section>`
            : `<section style="${cardStyle}"><div style="border:1px dashed #374151;border-radius:10px;padding:40px;text-align:center;color:#64748B">AFS 1 unavailable<br><small>${esc(r1.message || '')}</small></div></section>`;

        html += r2.success && r2.b64
            ? `<section style="${cardStyle}">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px;min-width:0">
                    <div style="min-width:0;flex:1 1 auto">
                        <div style="font-size:12px;color:#0F172A;font-weight:700">AFS 2 - page ${pg2 ?? '?'}</div>
                        <div style="font-size:10px;color:#F59E0B;font-family:'IBM Plex Mono',monospace">yellow = changed or added content</div>
                    </div>
                    <div style="display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end;flex:0 0 auto">
                        ${r2.pdfUrl ? `<a href="${r2.pdfUrl}" target="_blank" rel="noopener" style="${linkStyle}">Open PDF</a>` : ''}
                        <a href="data:image/png;base64,${r2.b64}" target="_blank" rel="noopener" style="${linkStyle}">Open image</a>
                    </div>
                </div>
                <img src="data:image/png;base64,${r2.b64}" alt="AFS 2 page ${pg2 ?? '?'}" style="${imgStyle}"/>
            </section>`
            : `<section style="${cardStyle}"><div style="border:1px dashed #7C2D12;border-radius:10px;padding:40px;text-align:center;color:#F59E0B;font-size:12px">Page has no match in AFS 2</div></section>`;

        html += '</div></div>';

        if (p && p.diffs && p.diffs.length) {
            html += `<div style="margin:10px 0 4px;color:#F59E0B;font-weight:700;font-size:12px">CHANGES ON THIS PAGE PAIR</div>`;
            html += `<table style="width:100%;border-collapse:collapse"><thead><tr style="background:#1E2F5A"><th style="${thStyle};width:32px"></th><th style="${thStyle}">DIFF</th><th style="${thStyle};width:100px">NUM</th></tr></thead><tbody>`;
            p.diffs.forEach((d, j) => {
                const sym = (d.status === 'changed' ? '~' : d.status === 'added' ? '+' : '-') + (j + 1);
                const bgr = d.status === 'changed' ? '#1C1400' : d.status === 'added' ? '#081424' : '#1A0808';
                const sc = d.status === 'changed' ? '#F59E0B' : d.status === 'added' ? '#3B82F6' : '#EF4444';
                const ct = renderDiffLine(d);
                const nd = d.numDiff && d.numDiff.length ? esc(d.numDiff.join(', ')) : '-';
                html += `<tr style="background:${bgr};vertical-align:top"><td style="${tdStyle};color:${sc};font-family:'IBM Plex Mono',monospace;white-space:nowrap">${sym}</td><td style="${tdStyle};font-family:'IBM Plex Mono',monospace;font-size:11px">${ct}</td><td style="${tdStyle};color:#FBBF24;font-family:'IBM Plex Mono',monospace">${nd}</td></tr>`;
            });
            html += '</tbody></table>';
        }

        html += commentBox('page:' + resolvedPairIndex, `AFS1 p${pg1} <-> AFS2 p${pg2 || '?'} findings`);
        wrap.innerHTML = html;
    } catch (e) {
        if (statusEl) statusEl.textContent = '';
        if (wrap) wrap.innerHTML = `<div style="color:#FCA5A5;font-size:12px;padding:8px">Error: ${esc(e.message)}</div>`;
    }
}

function toggleDarkMode() {
    const isDark = document.body.classList.toggle('dark-mode');
    localStorage.setItem('afs-theme', isDark ? 'dark' : 'light');
    const btn = document.getElementById('themeTog');
    if (btn) btn.innerHTML = isDark ? '<span>☀</span> Light Mode' : '<span>🌙</span> Dark Mode';
}

// Keep legacy alias for any stale cached references
const toggleLightMode = toggleDarkMode;

// Restore saved theme immediately to avoid flash
(function () {
    if (localStorage.getItem('afs-theme') === 'dark') {
        document.body.classList.add('dark-mode');
        document.addEventListener('DOMContentLoaded', function () {
            const btn = document.getElementById('themeTog');
            if (btn) btn.innerHTML = '<span>☀</span> Light Mode';
        });
    }
})();

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 8 · TAB SWITCHER
// ─────────────────────────────────────────────────────────────────────────────

function switchTab(tabId) {
    document.querySelectorAll('.afs-tab-pane').forEach(p => p.style.display = 'none');
    const target = document.getElementById(tabId);
    if (target) target.style.display = 'block';
    document.querySelectorAll('#compare-results .afs-tab').forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('onclick') && btn.getAttribute('onclick').includes(tabId)) {
            btn.classList.add('active');
        }
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 9 · AUDITOR COMMENT SAVE
// Reference: POST /comment endpoint
// ─────────────────────────────────────────────────────────────────────────────

async function saveComment(inputId, key) {
    const ta  = document.getElementById(inputId);
    const okEl = document.getElementById('cmt_ok_' + inputId);
    if (!ta) return;
    const data = new FormData();
    data.append('key',     key);
    data.append('comment', ta.value);
    if (okEl) okEl.textContent = 'Saving...';
    try {
        const res = await apiPost('/comment', data);
        if (okEl) okEl.textContent = res.success ? '✓ Saved' : '✗ Error';
    } catch (e) {
        if (okEl) okEl.textContent = '✗ ' + e.message;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 10 · STEP 5 — EXPORT
// Reference: Python on_save() handler
// ─────────────────────────────────────────────────────────────────────────────

async function saveWorkingPapers() {
    const saveBtn = document.querySelector('button[onclick="saveWorkingPapers()"]');
    if (saveBtn) saveBtn.disabled = true;

    setStatus('export-status', 'Saving...', 'info');
    const resultsEl = document.getElementById('export-results');
    resultsEl.innerHTML = '';

    const data = new FormData();
    data.append('outputFolder', document.getElementById('export_folder').value || '');
    data.append('exportPdf',    document.getElementById('chk_pdf').checked   ? 'true' : 'false');
    data.append('exportWord',   document.getElementById('chk_word').checked  ? 'true' : 'false');
    data.append('exportExcel',  document.getElementById('chk_excel').checked ? 'true' : 'false');
    data.append('exportText',   document.getElementById('chk_text').checked  ? 'true' : 'false');

    try {
        const res = await apiPost('/api/export', data, { timeoutMs: 600000 });
        setStatus('export-status', res.message, res.success ? 'ok' : 'error');
        renderExportResults(res);
    } catch (e) {
        setStatus('export-status', 'Error: ' + e.message, 'error');
    } finally {
        if (saveBtn) saveBtn.disabled = false;
    }
}

function renderExportResults(res) {
    const wrap = document.getElementById('export-results');
    dashboardState.exportSaved = !!(res.saved && res.saved.length);
    refreshDashboard();
    let html = '';

    if (res.saved && res.saved.length) {
        html += `<div style="background:#052E16;border:2px solid #10B981;border-radius:8px;padding:10px 16px;font-family:'IBM Plex Mono',monospace;margin:8px 0">
            <div style="color:#10B981;font-size:13px;font-weight:700;margin-bottom:4px">✓ ${res.saved.length} file(s) saved</div>
            <div style="color:#6EE7B7;font-size:10px">Folder: ${esc(res.folder)}</div>`;
        res.saved.forEach(f => {
            html += `<div style="color:#A7F3D0;font-size:11px;margin-top:2px">✔ ${esc(f)}</div>`;
        });
        html += '</div>';
    }

    if (res.errors && res.errors.length) {
        html += `<div style="background:#2D0000;border:2px solid #EF4444;border-radius:8px;padding:10px 16px;font-family:'IBM Plex Mono',monospace;margin:8px 0">
            <div style="color:#EF4444;font-size:13px;font-weight:700">✗ ${res.errors.length} error(s)</div>`;
        res.errors.forEach(e => {
            html += `<div style="color:#FCA5A5;font-size:11px;margin-top:2px">• ${esc(e)}</div>`;
        });
        html += '</div>';
    }

    if (!res.saved || !res.saved.length) {
        html += `<div style="background:#2D0000;border:2px solid #EF4444;border-radius:8px;padding:10px 16px;font-family:'IBM Plex Mono',monospace">
            <div style="color:#EF4444;font-size:13px;font-weight:700">✗ Nothing was saved — see errors above.</div>
            <div style="color:#FCA5A5;font-size:10px;margin-top:4px">Common fixes:<br>
            • Run Step 4 (Compare) first<br>
            • Check that the output folder path is valid</div>
        </div>`;
    }

    wrap.innerHTML = html;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 11 · PAGE LOAD — restore session state
// Pre-fills engagement fields and queue from the server session
// ─────────────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
    refreshDashboard();
    setActiveStep('step1');
    await loadModelStatus();
    window.setInterval(() => { loadModelStatus(); }, 30000);

    try {
        // Pre-fill engagement form
        const eng = await fetch('/api/engagement').then(r => r.json());
        if (eng) {
            const map = {
                eng_parent:    eng.parent,
                eng_client:    eng.client,
                eng_engname:   eng.engagementName,
                eng_engnum:    eng.engagementNumber,
                eng_fystart:   eng.financialYearStart,
                eng_fyend:     eng.financialYearEnd,
                eng_auditor:   eng.preparedBy,
                eng_director:  eng.director,
                eng_manager:   eng.manager,
                eng_objective: eng.objective,
            };
            for (const [id, val] of Object.entries(map)) {
                const el = document.getElementById(id);
                if (el && val) el.value = val;
            }
            dashboardState.engagementSaved = !!(eng.client || eng.engagementName || eng.engagementNumber);
            refreshDashboard();
        }
    } catch (e) { /* ignore */ }

    try {
        // Restore queue
        const q = await fetch('/api/queue').then(r => r.json());
        if (q && q.queue && q.queue.length) {
            renderQueue(q.queue);
        }
    } catch (e) { /* ignore */ }

    // Set default export folder — prefer the server-injected data-default attribute,
    // fall back to a best-guess path only if the input has no value.
    const expFld = document.getElementById('export_folder');
    if (expFld && !expFld.value) {
        expFld.value = expFld.dataset.default || 'C:\\Users\\Desktop\\AFS_Comparison';
    }

    refreshDashboard();
});
