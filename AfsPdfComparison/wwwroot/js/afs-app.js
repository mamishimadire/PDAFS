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

async function apiPost(url, formData) {
    formData.append('__RequestVerificationToken', getToken());
    const r = await fetch(url, { method: 'POST', body: formData });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.json();
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 2 · ACCORDION TOGGLE
// Mirrors Python acc.selected_index behaviour.
// ─────────────────────────────────────────────────────────────────────────────

function toggleAcc(step) {
    const body  = document.getElementById('body-' + step);
    const arrow = document.getElementById('arrow-' + step);
    if (!body) return;
    const open = body.style.display !== 'none';
    body.style.display  = open ? 'none' : 'block';
    arrow.innerHTML     = open ? '&#9658;' : '&#9660;';
}

function openAcc(step) {
    const body  = document.getElementById('body-' + step);
    const arrow = document.getElementById('arrow-' + step);
    if (body)  body.style.display  = 'block';
    if (arrow) arrow.innerHTML     = '&#9660;';
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
        input.value = '';
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
        document.getElementById('extract-results').innerHTML = '';
        document.getElementById('compare-results').innerHTML = '';
        document.getElementById('export-results').innerHTML  = '';
    } catch (e) {
        setStatus('upload-status', 'Error: ' + e.message, 'error');
    }
}

function renderQueue(queue) {
    const wrap = document.getElementById('queue-table-body');
    if (!queue || queue.length === 0) {
        wrap.innerHTML = '<div style="color:#4B5563;font-size:12px">No PDFs queued yet.</div>';
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
    setStatus('extract-status', 'Extracting...', 'info');
    document.getElementById('extract-results').innerHTML = '';
    try {
        const res = await apiPost('/api/extract', new FormData());
        setStatus('extract-status', res.message, res.success ? 'ok' : 'warn');
        renderExtractResults(res.reports || [], res.errors || []);
        if (res.success) openAcc('step4');
    } catch (e) {
        setStatus('extract-status', 'Error: ' + e.message, 'error');
    }
}

function renderExtractResults(reports, errors) {
    const wrap = document.getElementById('extract-results');
    if (!reports.length) { wrap.innerHTML = ''; return; }

    let rows = '';
    reports.forEach((r, i) => {
        const bg = i % 2 === 0 ? '#162347' : '#111827';
        const tc = r.docType === 'digital' ? '#059669' : r.docType === 'scanned' ? '#D97706' : '#EF4444';
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

    setStatus('compare-status', 'Running comparison...', 'info');
    progWrap.style.display = 'block';
    results.innerHTML      = '';

    // Animate progress bar (indeterminate — server is busy)
    let pct = 0;
    const timer = setInterval(() => {
        pct = Math.min(pct + 2, 92);
        progBar.style.width = pct + '%';
        progLbl.textContent = `Running... ${pct}%`;
    }, 150);

    try {
        const res = await apiPost('/api/compare', new FormData());
        clearInterval(timer);
        progBar.style.width = '100%';
        progLbl.textContent = 'Done';
        setTimeout(() => { progWrap.style.display = 'none'; }, 600);

        if (!res.success) {
            setStatus('compare-status', res.message, 'error');
            return;
        }
        setStatus('compare-status', res.message, 'ok');
        renderComparison(res);
        openAcc('step5');
    } catch (e) {
        clearInterval(timer);
        progWrap.style.display = 'none';
        setStatus('compare-status', 'Error: ' + e.message, 'error');
    }
}

function renderComparison(data) {
    const wrap = document.getElementById('compare-results');
    const c    = data.counts;
    const al   = data.alignment;
    const nc   = data.numCmp;
    const total = c.total || 1;
    const pctAffected = ((c.changed + c.added + c.removed) / total * 100).toFixed(1);

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
            html += `<tr style="background:${bg}">
                <td style="${tdStyle}">AFS1 p${p.pageAfs1 || '?'}</td>
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
            const flag  = issues > 5 ? ' !!' : (issues > 0 ? ' !' : ' OK');
            html += `<option value="${k}">AFS1 p${p.pageAfs1} ↔ ${p.pageAfs2 ? 'AFS2 p'+p.pageAfs2 : 'NO MATCH'}${flag} (${issues} changes)</option>`;
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
            const flag = issues > 5 ? ' !!' : issues > 0 ? ' !' : ' OK';
            html += `<option value="${k}">AFS1 p${p.pageAfs1} ↔ ${p.pageAfs2 ? 'AFS2 p'+p.pageAfs2 : 'NO MATCH'}${flag} (${issues} changes)</option>`;
        });
    }
    html += `</select>
            <button class="afs-btn afs-btn-primary" onclick="showSnapshot()">Show Snapshot</button>
            <span id="snap-status" style="font-size:11px;color:#A78BFA;font-family:'IBM Plex Mono',monospace"></span>
        </div>
        <div id="snap-results"></div>
    </div>`;

    // ── Tab: Comments ─────────────────────────────────────────────────────
    html += `<div id="tab-comments" class="afs-tab-pane" style="display:none">
        <div style="color:#FBBF24;font-weight:700;font-size:13px;margin:10px 0 6px">📋 CONSOLIDATED AUDITOR COMMENTS</div>
        ${commentBox('overall', 'Overall Audit Conclusion / Sign-off Notes')}
    </div>`;

    html += '</div></div>'; // end tab-content + compare-results

    wrap.innerHTML = html;

    // Store data for page diff viewer
    window._cmpData = data;
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
}

function showPageDiff() {
    const data  = window._cmpData;
    if (!data) return;
    const k     = parseInt(document.getElementById('pairSelect').value, 10);
    const p     = data.pageDiffs[k];
    const wrap  = document.getElementById('pair-detail');
    if (!p) { wrap.innerHTML = ''; return; }

    const pct  = p.pctSame;
    const fc   = pct === 100 ? '#10B981' : pct >= 70 ? '#F59E0B' : '#EF4444';
    const lbl  = `AFS1 p${p.pageAfs1} ↔ ${p.pageAfs2 ? 'AFS2 p'+p.pageAfs2 : 'UNMATCHED'}`;

    let html = `<div style="background:#111827;border-radius:7px;padding:8px 12px;font-family:'IBM Plex Mono',monospace;font-size:11px;color:#CBD5E1;margin-bottom:8px">
        <b style="color:#A78BFA">Pair ${k+1}</b>  |
        <span style="color:#7B4FFF"> ${esc(lbl)}</span>  |
        <span style="color:#10B981">✓ ${p.same} same</span>
        <span style="color:#F59E0B"> ~ ${p.changed} changed</span>
        <span style="color:#3B82F6"> + ${p.added} added</span>
        <span style="color:#EF4444"> − ${p.removed} removed</span>  |
        <span style="color:${fc}">${pct}% same</span>  |
        <span style="color:#64748B">align: ${p.alignSim.toFixed(3)}</span>
    </div>`;

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
            let content = '';
            if (d.status === 'changed') content = renderWordDiff(d.wordDiff) || esc(d.line1) + ' → ' + esc(d.line2);
            else if (d.status === 'added') content = `<span style="color:#BAE6FD">${esc(d.line2)}</span>`;
            else content = `<span style="color:#FECACA">${esc(d.line1)}</span>`;
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
    html += commentBox('page:' + k, lbl + ' findings');
    wrap.innerHTML = html;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7b · PAGE VISUAL VIEWER — renders PDF page images via /api/snapshot
// Reference: Python _show_page_snapshot() with Poppler pdftoppm
// ─────────────────────────────────────────────────────────────────────────────

async function showSnapshot() {
    const sel = document.getElementById('visualPairSelect');
    if (!sel) return;
    const k = parseInt(sel.value, 10);
    const statusEl = document.getElementById('snap-status');
    const wrap = document.getElementById('snap-results');
    if (statusEl) statusEl.textContent = 'Loading...';
    if (wrap) wrap.innerHTML = '<div style="color:#A78BFA;font-family:\'IBM Plex Mono\',monospace;font-size:11px;padding:6px">Rendering page images via Poppler…</div>';
    try {
        const res = await fetch('/api/snapshot?pairIndex=' + k).then(r => r.json());
        if (statusEl) statusEl.textContent = '';
        if (!res.b64Afs1 && !res.b64Afs2) {
            wrap.innerHTML = `<div style="background:#1C1A00;border:1px solid #F59E0B;color:#FCD34D;padding:10px 14px;border-radius:7px;font-size:12px;font-family:'IBM Plex Mono',monospace">
                ⚠ ${esc(res.message || 'Page images unavailable — Poppler not installed.')}<br>
                <span style="color:#94A3B8;font-size:11px">The page diff text view is available in the "Page Map" tab.</span>
            </div>`;
            return;
        }
        const pct = res.pctSame;
        const fc  = pct === 100 ? '#10B981' : pct >= 70 ? '#F59E0B' : '#EF4444';
        const p   = window._cmpData?.pageDiffs?.[k];
        let html  = `<div style="background:#111827;border-radius:7px;padding:8px 12px;font-family:'IBM Plex Mono',monospace;font-size:11px;color:#CBD5E1;margin-bottom:8px">
            <b style="color:#A78BFA">Pair ${k+1}</b>  |
            <span style="color:#7B4FFF"> ${esc(res.afs1Label)}</span>  |
            <span style="color:#10B981"> ✓ ${res.same} same</span>
            <span style="color:#F59E0B"> ~ ${res.changed} changed</span>
            <span style="color:#3B82F6"> + ${res.added} added</span>
            <span style="color:#EF4444"> − ${res.removed} removed</span>  |
            <span style="color:${fc}">${pct}% same</span>
        </div>`;
        // Side-by-side page images
        html += '<div style="color:#64748B;font-size:10px;font-family:\'IBM Plex Mono\',monospace;margin-bottom:4px">Yellow bands = changed / added / removed lines — from the Page Map text diff</div>';
        html += '<table style="width:100%;border-collapse:collapse"><tr>';
        html += res.b64Afs1
            ? `<td style="width:50%;padding:3px;vertical-align:top"><img src="data:image/png;base64,${res.b64Afs1}" style="width:100%;border-radius:6px;border:1px solid rgba(255,255,255,0.1);box-shadow:0 4px 16px rgba(0,0,0,0.5)"/><div style="font-size:10px;color:#64748B;text-align:center;margin-top:3px;font-family:'IBM Plex Mono',monospace">${esc(res.afs1Label)}</div></td>`
            : `<td style="width:50%;padding:3px;vertical-align:top"><div style="background:#111827;border:1px dashed #374151;border-radius:6px;padding:40px;text-align:center;color:#64748B">Image unavailable</div></td>`;
        html += res.b64Afs2
            ? `<td style="width:50%;padding:3px;vertical-align:top"><img src="data:image/png;base64,${res.b64Afs2}" style="width:100%;border-radius:6px;border:1px solid rgba(255,255,255,0.1);box-shadow:0 4px 16px rgba(0,0,0,0.5)"/><div style="font-size:10px;color:#64748B;text-align:center;margin-top:3px;font-family:'IBM Plex Mono',monospace">${esc(res.afs2Label)}</div></td>`
            : `<td style="width:50%;padding:3px;vertical-align:top"><div style="background:#1C1400;border-radius:6px;padding:40px;text-align:center;color:#F59E0B;font-size:12px">Page has no match in AFS 2</div></td>`;
        html += '</tr></table>';
        // Change list for this page pair
        if (p && p.diffs && p.diffs.length) {
            html += `<div style="margin:10px 0 4px;color:#F59E0B;font-weight:700;font-size:12px">CHANGES ON THIS PAGE PAIR</div>`;
            html += `<table style="width:100%;border-collapse:collapse"><thead><tr style="background:#1E2F5A"><th style="${thStyle};width:32px"></th><th style="${thStyle}">DIFF</th><th style="${thStyle};width:100px">NUM</th></tr></thead><tbody>`;
            p.diffs.forEach((d, j) => {
                const sym = (d.status==='changed'?'~':d.status==='added'?'+':'−')+(j+1);
                const bgr = d.status==='changed'?'#1C1400':d.status==='added'?'#081424':'#1A0808';
                const sc  = d.status==='changed'?'#F59E0B':d.status==='added'?'#3B82F6':'#EF4444';
                const ct  = d.status==='changed' ? (renderWordDiff(d.wordDiff)||esc(d.line1)+' → '+esc(d.line2))
                          : d.status==='added'  ? `<span style="color:#BAE6FD">${esc(d.line2)}</span>`
                          : `<span style="color:#FECACA">${esc(d.line1)}</span>`;
                const nd  = d.numDiff&&d.numDiff.length ? esc(d.numDiff.join(', ')) : '—';
                html += `<tr style="background:${bgr};vertical-align:top"><td style="${tdStyle};color:${sc};font-family:'IBM Plex Mono',monospace;white-space:nowrap">${sym}</td><td style="${tdStyle};font-family:'IBM Plex Mono',monospace;font-size:11px">${ct}</td><td style="${tdStyle};color:#FBBF24;font-family:'IBM Plex Mono',monospace">${nd}</td></tr>`;
            });
            html += '</tbody></table>';
        }
        html += commentBox('page:'+k, `AFS1 p${p?.pageAfs1||k+1} ↔ AFS2 p${p?.pageAfs2||'?'} findings`);
        wrap.innerHTML = html;
    } catch (e) {
        if (statusEl) statusEl.textContent = '';
        if (wrap) wrap.innerHTML = `<div style="color:#FCA5A5;font-size:12px;padding:8px">Error: ${esc(e.message)}</div>`;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7c · LIGHT MODE TOGGLE
// ─────────────────────────────────────────────────────────────────────────────

function toggleLightMode() {
    const isLight = document.body.classList.toggle('light-mode');
    localStorage.setItem('afs-theme', isLight ? 'light' : 'dark');
    const btn = document.getElementById('themeTog');
    if (btn) btn.textContent = isLight ? '🌙 Dark Mode' : '☀️ Light Mode';
}

// Restore saved theme immediately (before DOMContentLoaded) to avoid flash
(function () {
    if (localStorage.getItem('afs-theme') === 'light') {
        document.body.classList.add('light-mode');
        document.addEventListener('DOMContentLoaded', function () {
            const btn = document.getElementById('themeTog');
            if (btn) btn.textContent = '🌙 Dark Mode';
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
        const res = await apiPost('/api/export', data);
        setStatus('export-status', res.message, res.success ? 'ok' : 'error');
        renderExportResults(res);
    } catch (e) {
        setStatus('export-status', 'Error: ' + e.message, 'error');
    }
}

function renderExportResults(res) {
    const wrap = document.getElementById('export-results');
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
        }
    } catch (e) { /* ignore */ }

    try {
        // Restore queue
        const q = await fetch('/api/queue').then(r => r.json());
        if (q && q.queue && q.queue.length) {
            renderQueue(q.queue);
        }
    } catch (e) { /* ignore */ }

    // Set default export folder
    const expFld = document.getElementById('export_folder');
    if (expFld && !expFld.value) {
        expFld.value = 'C:\\Users\\' + (navigator.userAgent.includes('Win') ? '' : '') + 'Desktop\\AFS_Comparison';
    }
});
