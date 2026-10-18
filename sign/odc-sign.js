/* ═══════════════════════════════════════════════════════════════════════
   ODC client-sign overlay
   Loaded at the end of contract/index.html and onboarding/index.html.

   Two modes, decided by the URL:
     • coordinator mode  (no #sid=)  — the normal form; adds "Create client link"
     • client mode       (#sid=<id>) — loads the session, locks the documents,
                                       lets the participant sign and submit.
   All server work goes through the Supabase Edge Function in
   supabase/functions/sign-session (Microsoft Graph → SharePoint + e-mail).
   ═══════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  const PAGE = window.ODC_SIGN_PAGE;                 // 'contract' | 'onboarding'
  const CFG  = window.ODC_SIGN_CONFIG || {};
  const FN   = CFG.functionUrl || '';
  const q    = new URLSearchParams(location.search);
  const SID  = ((location.hash.match(/^#sid=([0-9a-f-]{36})/i) || [])[1] || q.get('sid') || '').toLowerCase();
  const MODE = SID ? 'client' : 'coordinator';
  const LS_COORD = 'odcs_coord_v1';                   // remembered coordinator e-mail + PIN

  document.documentElement.classList.add('odc-sign-' + MODE);
  console.log('[odc-sign] page=' + PAGE + ' mode=' + MODE + (SID ? ' sid=' + SID : ''));

  // Onboarding slot roles (contract has STAFF_SLOTS / REP_SLOTS of its own)
  const OB_STAFF = new Set(['sig_1', 'sig_2', 'sig_10']);
  const OB_REP   = new Set(['sig_4', 'sig_9']);
  const OB_CONSENT_PAIR = new Set(['sig_6', 'sig_7']);
  const OB_FORMS = {   // form id → [title for the PDF footer, file name]
    client:      ['Client Information Form',                'Client information form.pdf'],
    safety:      ['Home Safety Checklist',                  '04. Home Safety Checklist.pdf'],
    intake:      ['Participant Intake Checklist',           '01. Participant Intake Checklist.pdf'],
    consent:     ['Participant Consent Form',               '02. Participant Consent Form.pdf'],
    emergency:   ['Emergency and Disaster Management Plan', '03. Emergency and Disaster Management Plan.pdf'],
    eligibility: ['Self-Managed Eligibility Checklist',     '00-1. Self-Managed Eligibility Checklist.pdf']
  };

  // ── small helpers ─────────────────────────────────────────────────────
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const onLoad = new Promise(r => (document.readyState === 'complete' ? r() : window.addEventListener('load', r)));
  const onReady = new Promise(r => (document.readyState !== 'loading' ? r() : document.addEventListener('DOMContentLoaded', r)));

  function toast(html, kind) {
    if (typeof window._showToast === 'function') return window._showToast(html, kind);
    const c = { info: ['#eaf3ff', '#6b9bd1', '#1a4a7a'], success: ['#eef9ea', '#4a7a3c', '#2f4a2a'], error: ['#fef3ef', '#d99656', '#7a4a1a'] }[kind || 'info'];
    const t = document.createElement('div');
    t.className = 'no-print';
    t.style.cssText = 'position:fixed;top:14px;right:14px;max-width:360px;padding:12px 16px;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,.18);z-index:99999;font-size:.86rem;line-height:1.45;background:' + c[0] + ';border:1px solid ' + c[1] + ';color:' + c[2];
    t.innerHTML = html + ' <button onclick="this.parentNode.remove()" style="float:right;background:none;border:none;cursor:pointer;font-size:1.1rem;padding:0 0 0 8px;color:inherit;">&times;</button>';
    document.body.appendChild(t);
    if (kind !== 'error') setTimeout(() => { try { t.remove(); } catch (e) {} }, 7000);
    return t;
  }

  // Full-screen progress overlay — reuses the page's #cfOverlay markup
  const ov = {
    show(title, msg) { const o = $('cfOverlay'); if (!o) return; o.style.display = 'flex'; this.set(title, msg, 0); },
    set(title, msg, frac) {
      if (title != null && $('cfOverlayTitle')) $('cfOverlayTitle').textContent = title;
      if (msg != null && $('cfOverlayMsg')) $('cfOverlayMsg').textContent = msg;
      if (typeof frac === 'number' && $('cfOverlayBar')) $('cfOverlayBar').style.width = Math.round(Math.max(0, Math.min(1, frac)) * 100) + '%';
    },
    hide() { const o = $('cfOverlay'); if (o) o.style.display = 'none'; }
  };

  async function api(body, opts) {
    if (!FN || /YOUR-PROJECT-REF/.test(FN)) throw new Error('The e-sign service URL is not configured yet (sign/config.js).');
    const headers = {};
    if (CFG.anonKey) { headers['apikey'] = CFG.anonKey; headers['Authorization'] = 'Bearer ' + CFG.anonKey; }
    let res;
    if (body instanceof FormData) res = await fetch(FN, { method: 'POST', headers, body });
    else { headers['Content-Type'] = 'application/json'; res = await fetch(FN, { method: 'POST', headers, body: JSON.stringify(body) }); }
    let data = null;
    try { data = await res.json(); } catch (e) {}
    if (!res.ok || !data || data.ok === false) {
      throw new Error((data && (data.error || data.message)) || ('Service error ' + res.status));
    }
    return data;
  }

  // ── Page adapters: what differs between the two apps ──────────────────
  const adapter = PAGE === 'contract' ? {
    participantName()  { return ($('f_name') || {}).value || ''; },
    participantEmail() { return ($('f_email') || {}).value || ''; },
    // Everything a client page needs to show exactly what the coordinator prepared
    collect() {
      const full = window._collectFull();
      let ff = {}; try { ff = (typeof ffCache !== 'undefined') ? ffCache : {}; } catch (e) {}
      return { v: full.v, sig: full.sig, doc: full.doc || {}, ff: ff };
    },
    async restore(p) {
      try { if (p.ff && typeof ffCache !== 'undefined') ffCache = Object.assign({}, p.ff); } catch (e) {}
      window.docInputCache = Object.assign({}, p.doc || {});
      window._restoreAllValues(p.v || []);
      if (p.sig) window._restoreSigState(p.sig);
      if (p.doc) window._restoreDocInputCache(p.doc);
      await wait(900);
      try { window._restoreDocInputs && window._restoreDocInputs(); } catch (e) {}
    },
    hasParticipantSignature() {
      try { return appliedClientIds.size > 0 || appliedRepIds.size > 0; } catch (e) { return false; }
    },
    async buildPdfs(progress) {
      const out = await window.odcBuildPdfForCurrentSelection();   // one PDF: Service Agreement (+ Letter of Variation)
      progress(1, 1);
      return [{ blob: out.blob, filename: out.filename, pages: out.pages }];
    },
    docLabel: 'Service Agreement'
  } : {
    participantName()  { const e = document.querySelector('[data-shared="p_name"]'); return e ? e.value : ''; },
    participantEmail() { const e = document.querySelector('[data-shared="p_email"]'); return e ? e.value : ''; },
    collect() {
      const full = window._collectFull();
      let picks = []; try { picks = Array.from(_picks); } catch (e) {}
      return { v: full.v, sigs: full.sigs || {}, k: full.k || null, ai: full.ai || null, picks: picks };
    },
    async restore(p) {
      try {
        if (Array.isArray(p.picks) && p.picks.length) {
          _picks = new Set(p.picks); _picks.add('client');
          if (typeof setPicks === 'function') setPicks(_picks);
          if (typeof applyPicksToTabs === 'function') applyPicksToTabs();
        }
      } catch (e) { console.warn('[odc-sign] picks restore:', e); }
      orig.restoreAllValues(p.v || [], p.k || null, p.ai || null);
      if (p.sigs) orig.restoreSignatures(p.sigs);
      await wait(300);
    },
    hasParticipantSignature() {
      return Array.prototype.some.call(document.querySelectorAll('[data-sigid].applied'), s => !OB_STAFF.has(s.dataset.sigid));
    },
    pickedIds() {
      let picks = ['client']; try { picks = Array.from(_picks); } catch (e) {}
      const order = ['client', 'safety', 'intake', 'consent', 'emergency', 'eligibility'];
      return order.filter(id => picks.indexOf(id) !== -1 && $('form-' + id));
    },
    async buildPdfs(progress) {
      const ids = this.pickedIds();
      const labels = {}; Object.keys(OB_FORMS).forEach(id => { labels['form-' + id] = OB_FORMS[id][0]; });
      const name = (this.participantName() || 'Client').replace(/[\\/:*?"<>|#%&{}]/g, '').trim();
      const out = [];
      for (let i = 0; i < ids.length; i++) {
        const id = ids[i];
        ov.set('Creating PDF ' + (i + 1) + ' of ' + ids.length + ' — ' + OB_FORMS[id][0], 'Laying out pages…', i / ids.length);
        const r = await window.odcRenderPdf(['form-' + id], {
          footerLabels: labels,
          onProgress: p => ov.set(null, 'Page ' + Math.min(p.done + 1, p.total) + ' of ' + p.total, (i + p.done / Math.max(1, p.total)) / ids.length)
        });
        out.push({ blob: r.blob, filename: name + '_' + OB_FORMS[id][1], pages: r.pages });
        progress(i + 1, ids.length);
      }
      return out;
    },
    docLabel: 'Onboarding Forms'
  };

  // Originals we wrap (function declarations in the host page are global bindings)
  const orig = {};

  // ═══════════════════════════════════════════════════════════════════════
  // CLIENT MODE
  // ═══════════════════════════════════════════════════════════════════════
  if (MODE === 'client') {
    let session = null;

    // 1. Nothing from a previous visitor on this device may leak into the view
    try { if (typeof window.clearLocalSave === 'function') window.clearLocalSave(); else if (typeof clearLocalSave === 'function') clearLocalSave(); } catch (e) { console.warn('[odc-sign] clearLocalSave:', e); }
    try { Object.keys(localStorage).forEach(k => { if (/^odc_(shared_|sig_|all_forms|contracts|sa_)/.test(k)) localStorage.removeItem(k); }); } catch (e) {}

    // 2. Stop the host page restoring its own (now empty / stale) saves
    try { if (typeof loadSavedData === 'function') window.removeEventListener('DOMContentLoaded', loadSavedData); } catch (e) {}
    let gate = true;                                   // block host restores until the session is in
    try { orig.restoreAllValues = _restoreAllValues; _restoreAllValues = function () { if (!gate) return orig.restoreAllValues.apply(this, arguments); }; } catch (e) {}
    try { orig.restoreSignatures = _restoreSignatures; _restoreSignatures = function () { if (!gate) return orig.restoreSignatures.apply(this, arguments); }; } catch (e) {}
    try { orig.showRestoreToast = _showRestoreToast; _showRestoreToast = function () {}; } catch (e) {}
    try { autoSaveDebounced = function () {}; } catch (e) {}

    // 3. Participants may not sign or remove staff fields
    if (PAGE === 'contract') {
      try {
        orig.applySigTo = applySigTo; orig.removeSigFrom = removeSigFrom;
        applySigTo = function (id) { if (STAFF_SLOTS.has(id)) { toast('This field is signed by One Dream Community staff.', 'info'); return; } return orig.applySigTo(id); };
        removeSigFrom = function (id) { if (STAFF_SLOTS.has(id)) return; return orig.removeSigFrom(id); };
      } catch (e) { console.warn('[odc-sign] slot guard:', e); }
    } else {
      try {
        orig.openSigModal = window.openSigModal;
        window.openSigModal = function (id, label) { if (OB_STAFF.has(id)) { toast('This field is signed by One Dream Community staff.', 'info'); return; } return orig.openSigModal(id, label); };
        orig.clearSlot = window._clearSlotSignature;
        window._clearSlotSignature = function (id) { if (OB_STAFF.has(id)) return; return orig.clearSlot(id); };
        // "Apply to All" → every participant field on the picked forms, except the consent / do-not-consent pair
        window.applySigToAll = function () {
          if (!_sigPad || _sigPad.isEmpty()) { alert('Please draw a signature first.'); return; }
          const dataUrl = _sigToJPEG(_sigPad);
          let n = 0;
          document.querySelectorAll('[data-sigid]').forEach(slot => {
            const id = slot.dataset.sigid;
            if (OB_STAFF.has(id) || OB_REP.has(id) || OB_CONSENT_PAIR.has(id)) return;
            if (!slot.closest('.form-section')) return;
            _setSlotSignature(id, dataUrl); n++;
          });
          closeSigModal();
          toast('✓ Signature placed in ' + n + ' participant field' + (n === 1 ? '' : 's') + '.<br><small>Representative fields and the consent choice are signed individually.</small>', 'success');
        };
      } catch (e) { console.warn('[odc-sign] slot guard:', e); }
    }

    // 4. Load the session, then restore and adjust the UI
    onReady.then(() => {
      ov.show('Loading your documents…', 'Please wait a moment.');
      document.body.classList.add('odcs-loading');
    });

    Promise.all([api({ action: 'get', id: SID }), onLoad]).then(async ([data]) => {
      session = data;
      if (data.page && data.page !== PAGE) {           // link opened on the wrong page
        location.replace('../' + data.page + '/#sid=' + SID); return;
      }
      if (data.status === 'signed') { showDone({ already: true, signedAt: data.signedAt, emailedTo: data.emailedTo || [] }); ov.hide(); return; }
      if (data.status === 'expired') { showBlocked('This link has expired.', 'Please contact One Dream Community for a new link.'); ov.hide(); return; }
      if (data.status === 'submitting') { showBlocked('Your submission is being processed.', 'Please wait a few minutes and open the link again. If you did not submit anything, contact One Dream Community on ' + esc(CFG.officePhone || '') + '.'); ov.hide(); return; }
      gate = false;
      ov.set(null, 'Preparing the documents…', .5);
      await adapter.restore(data.payload || {});
      installClientUi(data);
      ov.hide();
      document.body.classList.remove('odcs-loading');
    }).catch(e => {
      console.error('[odc-sign] load failed:', e);
      ov.hide();
      showBlocked('We could not open this link.', esc(e.message || e) + '<br><br>Please check the link is complete, or contact One Dream Community on ' + esc(CFG.officePhone || '') + '.');
    });

    function installClientUi(data) {
      const name = esc(data.participant && data.participant.name || adapter.participantName());
      // Banner
      const b = document.createElement('div');
      b.className = 'odcs-banner client no-print';
      b.innerHTML = '<div><strong>Hello' + (name ? ' ' + name : '') + ',</strong> please review your ' + adapter.docLabel + ' below and sign where marked. '
        + 'The information has been prepared by your One Dream Community coordinator — if anything is incorrect, please contact us before signing.</div>'
        + (PAGE === 'contract'
          ? '<div class="odcs-steps"><span>1 · Draw your signature</span><span>2 · Review &amp; place it</span><span>3 · Submit</span></div>'
          : '<div class="odcs-steps"><span>1 · Open each form tab</span><span>2 · Click the signature boxes</span><span>3 · Submit</span></div>');
      const anchor = document.querySelector('.app-header, .topbar');
      if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(b, anchor.nextSibling); else document.body.prepend(b);

      if (PAGE === 'contract') {
        // Steps: 1 signature → 2 documents
        const t2 = $('tab2'), t3 = $('tab3');
        if (t2) t2.innerHTML = '<span class="step-badge">1</span> Your Signature';
        if (t3) t3.innerHTML = '<span class="step-badge">2</span> Review &amp; Sign Documents';
        // staff pad is not for the participant
        const staffCol = $('sigCanvasStaff') && $('sigCanvasStaff').closest('.col-md-4'); if (staffCol) staffCol.style.display = 'none';
        const repCol = $('sigCanvasRep') && $('sigCanvasRep').closest('.col-md-4');
        const repTitle = repCol && repCol.querySelector(':scope > div > div:first-child'); if (repTitle) repTitle.textContent = '🧑‍🤝‍🧑 Representative / Attorney (only if signing on the participant\'s behalf)';
        const alert2 = document.querySelector('#panel2 .alert');
        if (alert2) alert2.innerHTML = '<strong>✍️ Draw your signature</strong> with your finger, stylus or mouse, then click <strong>✓ Accept</strong>. On the next step you can place it on every signature line in one click.';
        document.querySelectorAll('#panel2 button').forEach(btn => { if (/goTo\(1\)/.test(btn.getAttribute('onclick') || '')) btn.style.display = 'none'; });
        const alert1 = document.querySelector('#panel3 .apply-all-bar');
        if (alert1) {
          const grp = alert1.querySelector('.d-flex.gap-2');
          if (grp) {
            grp.querySelectorAll('button').forEach(btn => { if (/applyAll|openPrintModal/.test(btn.getAttribute('onclick') || '')) return; btn.style.display = 'none'; });
            const applyBtn = grp.querySelector('button[onclick*="applyAll"]'); if (applyBtn) applyBtn.textContent = '✍️ Place my signature on all fields';
            const sb = document.createElement('button'); sb.className = 'odcs-btn blue'; sb.id = 'odcsSubmitBar'; sb.textContent = '📤 Submit signed agreement'; sb.onclick = submit; grp.appendChild(sb);
          }
        }
        goTo(2);
      } else {
        // first picked form open, submit button in the nav bar
        try { const ids = adapter.pickedIds(); if (ids.length) showForm(ids[0]); } catch (e) {}
        const nav = $('navBar');
        if (nav) { const sb = document.createElement('button'); sb.className = 'odcs-btn blue'; sb.id = 'odcsSubmitBar'; sb.textContent = '📤 Submit signed forms'; sb.onclick = submit; nav.appendChild(sb); }
        try { const modalApply = document.querySelector('#sigModalBackdrop button[onclick*="applySigToAll"]'); if (modalApply) modalApply.textContent = '✍️ Apply to all my fields'; } catch (e) {}
      }
      // Floating submit
      const f = document.createElement('div'); f.id = 'odcsFloat'; f.className = 'no-print';
      f.innerHTML = '<button class="odcs-btn blue big" id="odcsSubmitFloat">📤 Submit signed ' + (PAGE === 'contract' ? 'agreement' : 'forms') + '</button>';
      document.body.appendChild(f);
      $('odcsSubmitFloat').onclick = submit;
    }

    let submitting = false;
    async function submit() {
      if (submitting) return;
      if (!adapter.hasParticipantSignature()) {
        toast('⚠ <strong>No signature placed yet.</strong> ' + (PAGE === 'contract' ? 'Go to step 1 to draw your signature, then place it on the documents.' : 'Click the signature boxes on each form to sign.'), 'error');
        if (PAGE === 'contract') { try { if (!clientSigUrl && !repSigUrl) goTo(2); else goTo(3); } catch (e) {} }
        return;
      }
      const email = session && session.participant && session.participant.email;
      const msg = 'Submit your signed ' + adapter.docLabel + ' to One Dream Community?\n\n'
        + (email ? 'A copy will be e-mailed to ' + email + ' and to One Dream Community.' : 'A copy will be e-mailed to One Dream Community.')
        + '\n\nThis cannot be undone.';
      if (!confirm(msg)) return;
      submitting = true;
      document.querySelectorAll('#odcsSubmitBar, #odcsSubmitFloat').forEach(b => b.disabled = true);
      try {
        ov.show('Creating your signed PDF…', 'This can take up to a minute. Please keep this page open.');
        const files = await adapter.buildPdfs((done, total) => ov.set(null, 'Document ' + done + ' of ' + total + ' ready', done / total));
        ov.set('Sending to One Dream Community…', 'Uploading ' + files.length + ' file' + (files.length === 1 ? '' : 's') + '…', 0.95);
        const fd = new FormData();
        fd.append('action', 'submit');
        fd.append('id', SID);
        fd.append('page', PAGE);
        try { fd.append('snapshot', JSON.stringify(adapter.collect())); } catch (e) {}
        fd.append('tz', Intl.DateTimeFormat().resolvedOptions().timeZone || '');
        files.forEach((f, i) => fd.append('file' + i, f.blob, f.filename));
        const res = await api(fd);
        ov.hide();
        showDone({ emailedTo: res.emailedTo || [], files: files, signedAt: res.signedAt });
      } catch (e) {
        console.error('[odc-sign] submit failed:', e);
        ov.hide();
        submitting = false;
        document.querySelectorAll('#odcsSubmitBar, #odcsSubmitFloat').forEach(b => b.disabled = false);
        toast('✗ <strong>Could not submit.</strong> ' + esc(e.message || e) + '<br><small>Please try again. If it keeps failing, call One Dream Community on ' + esc(CFG.officePhone || '') + '.</small>', 'error');
      }
    }

    function showDone(o) {
      const d = document.createElement('div'); d.id = 'odcsDone';
      const when = o.signedAt ? new Date(o.signedAt).toLocaleString('en-AU', { dateStyle: 'medium', timeStyle: 'short' }) : '';
      let files = '';
      if (o.files && o.files.length) {
        files = '<div class="files"><p style="margin:14px 0 4px;"><strong>Your copy:</strong></p>' + o.files.map(f => {
          const url = URL.createObjectURL(f.blob); return '<a href="' + url + '" download="' + esc(f.filename) + '">⬇ ' + esc(f.filename) + ' (' + f.pages + ' pages)</a>';
        }).join('') + '</div>';
      }
      d.innerHTML = '<div class="card"><div class="tick">✅</div>'
        + '<h2>' + (o.already ? 'Already submitted' : 'Thank you — your ' + adapter.docLabel + ' has been submitted') + '</h2>'
        + (o.already ? '<p>These documents were signed and submitted' + (when ? ' on ' + esc(when) : '') + '. Nothing more is needed.</p>' : '')
        + (!o.already ? '<p>A signed copy has been sent to ' + (o.emailedTo && o.emailedTo.length ? '<strong>' + o.emailedTo.map(esc).join('</strong>, <strong>') + '</strong>' : 'One Dream Community') + '.' + (when ? ' Signed ' + esc(when) + '.' : '') + '</p>' : '')
        + files
        + '<p style="margin-top:18px;font-size:.88rem;color:#555;">Questions? Call One Dream Community on <strong>' + esc(CFG.officePhone || '') + '</strong> or e-mail <strong>' + esc(CFG.officeEmail || '') + '</strong>.</p>'
        + '</div></div>';
      document.body.appendChild(d);
      document.body.style.overflow = 'hidden';
    }
    function showBlocked(title, html) {
      const d = document.createElement('div'); d.id = 'odcsDone';
      d.innerHTML = '<div class="card"><div class="tick">⚠️</div><h2>' + esc(title) + '</h2><p>' + html + '</p></div>';
      document.body.appendChild(d);
      document.body.style.overflow = 'hidden';
    }
    return;
  }

  // ═══════════════════════════════════════════════════════════════════════
  // COORDINATOR MODE — the normal app plus "Create client link"
  // ═══════════════════════════════════════════════════════════════════════
  onReady.then(() => {
    const b = document.createElement('div');
    b.className = 'odcs-banner no-print';
    b.innerHTML = '<div><strong>🔗 Client e-sign link builder.</strong> Fill in the participant\'s details' + (PAGE === 'contract' ? ', choose the agreement variant, and (optionally) draw and apply the <strong>staff signature</strong>' : ' and tick the forms they need (✓ Pick Forms); sign any staff fields') + '. '
      + 'Then click <strong>Create client link</strong> — the participant opens the link, signs, and the signed PDF is e-mailed to them, to ' + esc(CFG.officeEmail || 'agedcare@') + ' and to you, and filed in SharePoint.</div>';
    const anchor = document.querySelector('.app-header, .topbar');
    if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(b, anchor.nextSibling); else document.body.prepend(b);

    const f = document.createElement('div'); f.id = 'odcsFloat'; f.className = 'no-print';
    f.innerHTML = '<button class="odcs-btn big" id="odcsCreateBtn">✉️ Create client link</button>';
    document.body.appendChild(f);
    $('odcsCreateBtn').onclick = openCreateModal;

    if (PAGE === 'contract') {
      const grp = document.querySelector('#panel3 .apply-all-bar .d-flex.gap-2');
      if (grp) { const sb = document.createElement('button'); sb.className = 'odcs-btn'; sb.textContent = '✉️ Create client link'; sb.onclick = openCreateModal; grp.appendChild(sb); }
    } else {
      const nav = $('navBar') && $('navBar').querySelector('div:last-child');
      if (nav) { const sb = document.createElement('button'); sb.className = 'nav-btn'; sb.style.cssText = 'background:#4a7c3f;color:#fff;border-color:#4a7c3f;'; sb.textContent = '✉️ Create client link'; sb.onclick = openCreateModal; nav.insertBefore(sb, nav.firstChild); }
    }

    const m = document.createElement('div'); m.id = 'odcsModalBackdrop'; m.className = 'no-print';
    m.innerHTML = '<div class="odcs-modal">'
      + '<h3>✉️ Create client e-sign link</h3>'
      + '<p class="sub">The participant opens the link in any browser — no login, no app. Everything you have entered (including staff signatures) is sent with it. Documents are read-only for the participant; they can only sign.</p>'
      + '<div class="row"><div><label>Participant name</label><input type="text" id="odcsPName"></div><div><label>Participant e-mail (copy of the signed PDF goes here)</label><input type="email" id="odcsPEmail" placeholder="participant@example.com"></div></div>'
      + '<div class="row"><div><label>Your e-mail (coordinator — you receive a copy)</label><input type="email" id="odcsCEmail" placeholder="you@onedreamcommunity.com.au"></div><div><label>Coordinator PIN</label><input type="password" id="odcsPin" autocomplete="off"></div></div>'
      + '<div class="msg" id="odcsMsg"></div>'
      + '<div class="link-box" id="odcsLinkBox" style="display:none;"><label style="margin-top:0;">Client link</label><textarea id="odcsLink" readonly></textarea>'
      + '<div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap;"><button class="odcs-btn ghost" id="odcsCopy">📋 Copy link</button><a class="odcs-btn ghost" id="odcsMail" href="#" style="text-decoration:none;">✉️ E-mail to participant</a><a class="odcs-btn ghost" id="odcsOpen" href="#" target="_blank" rel="noopener" style="text-decoration:none;">👁 Preview as participant</a></div>'
      + '<div style="font-size:.78rem;color:#666;margin-top:8px;">Valid for ' + (CFG.linkValidDays || 30) + ' days · can be submitted once · the participant sees the documents exactly as prepared here.</div></div>'
      + '<div class="actions"><button class="odcs-btn ghost" id="odcsClose">Close</button><button class="odcs-btn" id="odcsCreate">Create link</button></div>'
      + '</div>';
    document.body.appendChild(m);
    m.addEventListener('click', e => { if (e.target === m) closeModal(); });
    $('odcsClose').onclick = closeModal;
    $('odcsCreate').onclick = createLink;
    $('odcsCopy').onclick = () => { const t = $('odcsLink'); t.select(); try { navigator.clipboard.writeText(t.value); } catch (e) { document.execCommand('copy'); } toast('Link copied.', 'success'); };
  });

  function openCreateModal() {
    const saved = JSON.parse(localStorage.getItem(LS_COORD) || '{}');
    $('odcsPName').value = adapter.participantName();
    $('odcsPEmail').value = adapter.participantEmail();
    $('odcsCEmail').value = saved.email || '';
    $('odcsPin').value = saved.pin || '';
    $('odcsMsg').className = 'msg'; $('odcsMsg').innerHTML = '';
    $('odcsLinkBox').style.display = 'none';
    $('odcsCreate').disabled = false; $('odcsCreate').textContent = 'Create link';
    // make sure the documents exist so the snapshot includes everything typed into them
    if (PAGE === 'contract') { try { if (typeof renderDocs === 'function' && !$('docsContainer').children.length) { goTo(3); } } catch (e) {} }
    $('odcsModalBackdrop').classList.add('show');
    document.body.style.overflow = 'hidden';
  }
  function closeModal() { $('odcsModalBackdrop').classList.remove('show'); document.body.style.overflow = ''; }

  async function createLink() {
    const msg = $('odcsMsg'); msg.className = 'msg'; msg.innerHTML = '';
    const pName = $('odcsPName').value.trim(), pEmail = $('odcsPEmail').value.trim(), cEmail = $('odcsCEmail').value.trim(), pin = $('odcsPin').value;
    const emailOk = s => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
    if (!pName) return fail('Please enter the participant\'s name.');
    if (pEmail && !emailOk(pEmail)) return fail('The participant e-mail does not look right.');
    if (!emailOk(cEmail)) return fail('Please enter your (coordinator) e-mail.');
    if (!pin) return fail('Please enter the coordinator PIN.');
    localStorage.setItem(LS_COORD, JSON.stringify({ email: cEmail, pin: pin }));
    const btn = $('odcsCreate'); btn.disabled = true; btn.textContent = 'Creating…';
    try {
      const payload = adapter.collect();
      const res = await api({ action: 'create', pin: pin, page: PAGE, coordinatorEmail: cEmail, participantName: pName, participantEmail: pEmail, payload: payload });
      const url = location.origin + location.pathname + '#sid=' + res.id;
      $('odcsLink').value = url;
      $('odcsLinkBox').style.display = '';
      $('odcsOpen').href = url;
      const first = pName.split(' ')[0];
      const subject = 'One Dream Community – your ' + adapter.docLabel + ' to review and sign';
      const body = 'Dear ' + first + ',\n\nPlease open the link below to review and sign your ' + adapter.docLabel + ' with One Dream Community. No login is needed — it works on a phone, tablet or computer.\n\n' + url
        + '\n\nOnce you submit, a signed copy will be e-mailed to you. The link is valid for ' + (CFG.linkValidDays || 30) + ' days.\n\nIf anything in the documents is incorrect, or you have any questions, please call us on ' + (CFG.officePhone || '') + ' before signing.\n\nKind regards,\nOne Dream Community\n' + (CFG.officeEmail || '');
      $('odcsMail').href = 'mailto:' + encodeURIComponent(pEmail) + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body);
      msg.className = 'msg ok'; msg.innerHTML = '✓ Link created' + (res.expiresAt ? ' — valid until ' + esc(new Date(res.expiresAt).toLocaleDateString('en-AU')) : '') + '. Copy it or e-mail it to the participant.';
      btn.textContent = 'Create another link';
      btn.disabled = false;
      setTimeout(() => { try { $('odcsLink').select(); } catch (e) {} }, 50);
    } catch (e) {
      fail(e.message || String(e));
      btn.disabled = false; btn.textContent = 'Create link';
    }
    function fail(t) { msg.className = 'msg err'; msg.innerHTML = '✗ ' + esc(t); return false; }
  }
})();
