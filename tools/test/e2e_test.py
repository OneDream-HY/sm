import json, os, re, sys, time
from playwright.sync_api import sync_playwright

S = __import__('os').path.dirname(__import__('os').path.abspath(__file__))
LIB = S + '/node_modules'
BASE = 'http://localhost:8765'
PAGE = sys.argv[1] if len(sys.argv) > 1 else 'contract'
CDN = {
    'bootstrap.bundle.min.js': LIB + '/bootstrap/dist/js/bootstrap.bundle.min.js',
    'bootstrap.min.css': LIB + '/bootstrap/dist/css/bootstrap.min.css',
    'signature_pad.umd.min.js': LIB + '/signature_pad/dist/signature_pad.umd.min.js',
    'lz-string.min.js': LIB + '/lz-string/libs/lz-string.min.js',
    'jspdf.umd.min.js': LIB + '/jspdf/dist/jspdf.umd.min.js',
    'html2canvas.min.js': LIB + '/html2canvas/dist/html2canvas.min.js',
    'html2pdf.bundle.min.js': LIB + '/html2pdf.js/dist/html2pdf.bundle.min.js',
}

def route_cdn(route, request):
    name = request.url.split('/')[-1].split('?')[0]
    if name in CDN:
        ct = 'text/css' if name.endswith('.css') else 'text/javascript'
        return route.fulfill(path=CDN[name], content_type=ct)
    return route.fulfill(status=404, body='')

def attach_logging(page, tag):
    page.on('console', lambda m: print(f'  [{tag} console.{m.type}] {m.text[:220]}') if m.type in ('error', 'warning') or 'odc-sign' in m.text or 'mock' in m.text else None)
    page.on('pageerror', lambda e: print(f'  [{tag} PAGEERROR] {e}'))
    page.on('dialog', lambda d: (print(f'  [{tag} dialog] {d.type}: {d.message[:120]}'), d.accept()))

def draw(page, canvas_sel):
    box = page.locator(canvas_sel).bounding_box()
    x, y = box['x'] + 20, box['y'] + box['height'] / 2
    page.mouse.move(x, y); page.mouse.down()
    for i in range(1, 30):
        page.mouse.move(x + i * 6, y + (10 if i % 2 else -10), steps=2)
    page.mouse.up()

def new_page(browser, tag):
    ctx = browser.new_context(viewport={'width': 1300, 'height': 900})
    pg = ctx.new_page()
    pg.route(re.compile(r'https://(cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net)/.*'), route_cdn)
    attach_logging(pg, tag)
    return ctx, pg

with sync_playwright() as p:
    browser = p.chromium.launch()
    # ───────────── COORDINATOR ─────────────
    ctx, pg = new_page(browser, 'coord')
    pg.goto(f'{BASE}/{PAGE}/', wait_until='load')
    pg.wait_for_timeout(1200)
    assert pg.locator('#loginGate').count() == 0, 'login gate still present'
    assert pg.locator('.odcs-banner').count() == 1, 'coordinator banner missing'
    assert pg.locator('#odcsCreateBtn').is_visible()
    print('coordinator page loaded, banner + button present')
    if PAGE == 'contract':
        pg.fill('#f_name', 'Test Participant'); pg.fill('#f_email', 'participant@example.com'); pg.fill('#f_mobile', '0400 000 000')
        pg.fill('#f_street', '1 Test St'); pg.fill('#f_suburb', 'Burwood NSW'); pg.fill('#f_postcode', '2134')
        pg.fill('#f_startDate', '2026-11-01'); pg.fill('#f_officer', 'Coordinator Kim')
        pg.check('input[name="contractType"][value="selfmanaged"]')
        pg.click('text=Next: Capture Signature'); pg.wait_for_timeout(400)
        draw(pg, '#sigCanvasStaff'); pg.click("button:has-text('Accept') >> nth=2"); pg.wait_for_timeout(300)
        assert pg.locator('#staffSigBanner').is_visible(), 'staff sig not accepted'
        pg.click('text=Next: View Documents'); pg.wait_for_timeout(1500)
        pg.click('text=Apply Signature to ALL Fields'); pg.wait_for_timeout(500)
        n_staff = pg.evaluate('appliedStaffIds.size'); print('staff slots applied:', n_staff)
        assert n_staff >= 1
        # Edit an inline ff field (price list "Review Date" etc.) to check ffCache round-trips
        pg.evaluate("(() => { const i = document.querySelector('#docsContainer .ff-inline'); i.value = 'FF-ROUNDTRIP'; i.dispatchEvent(new Event('input')); })()")
    else:
        pg.fill('[data-shared="p_name"]', 'Test Participant')
        # pick consent + emergency + eligibility
        pg.click('text=Pick Forms'); pg.wait_for_timeout(300)
        for fid in ['consent', 'emergency', 'eligibility', 'intake']:
            pg.check(f'.fp-pick[data-fid="{fid}"]')
        pg.wait_for_timeout(300)
        pg.click('#formTabs .nav-btn[data-target="consent"]'); pg.wait_for_timeout(300)
        pg.fill('[data-shared="p_email"]', 'participant@example.com')
        # staff signs sig_10 (emergency staff) via modal
        pg.click('#formTabs .nav-btn[data-target="emergency"]'); pg.wait_for_timeout(300)
        pg.click('[data-sigid="sig_10"]'); pg.wait_for_timeout(400)
        draw(pg, '#sigCanvas'); pg.click("#sigModalBackdrop button:has-text('Apply to This One')"); pg.wait_for_timeout(300)
        assert pg.locator('[data-sigid="sig_10"].applied').count() == 1
    pg.click('#odcsCreateBtn'); pg.wait_for_timeout(400)
    assert pg.locator('#odcsModalBackdrop.show').count() == 1
    assert pg.input_value('#odcsPName') == 'Test Participant'
    assert pg.input_value('#odcsPEmail') == 'participant@example.com'
    pg.fill('#odcsCEmail', 'coordinator@onedreamcommunity.com.au'); pg.fill('#odcsPin', '0000')
    pg.click('#odcsCreate'); pg.wait_for_timeout(800)
    assert 'Incorrect coordinator PIN' in pg.inner_text('#odcsMsg'), pg.inner_text('#odcsMsg')
    pg.fill('#odcsPin', '1234'); pg.click('#odcsCreate'); pg.wait_for_timeout(1200)
    link = pg.input_value('#odcsLink')
    print('client link:', link)
    assert '#sid=' in link
    assert 'mailto:participant%40example.com' in pg.get_attribute('#odcsMail', 'href')
    ctx.close()

    # ───────────── CLIENT ─────────────
    ctx, pg = new_page(browser, 'client')
    pg.goto(link, wait_until='load')
    pg.wait_for_selector('.odcs-banner.client', timeout=20000)
    pg.wait_for_timeout(2500)
    assert 'Hello Test Participant' in pg.inner_text('.odcs-banner.client')
    assert pg.locator('#odcsSubmitFloat').is_visible()
    if PAGE == 'contract':
        assert not pg.locator('#panel1').is_visible() and not pg.locator('#tab1').is_visible()
        assert pg.locator('#panel2.active').count() == 1, 'client should start on signature step'
        assert not pg.locator('#sigCanvasStaff').is_visible()
        assert pg.input_value('#f_name') == 'Test Participant'
        assert pg.evaluate("document.querySelector('input[name=contractType]:checked').value") == 'selfmanaged'
        assert pg.evaluate('!!staffSigUrl'), 'staff signature not restored'
        draw(pg, '#sigCanvasClient'); pg.click("#panel2 button:has-text('Accept') >> nth=0"); pg.wait_for_timeout(300)
        assert pg.locator('#clientSigBanner').is_visible()
        pg.click('text=Next: View Documents'); pg.wait_for_timeout(2500)
        assert pg.locator('#docsContainer .doc-card').count() >= 1
        assert 'Self-managed' in pg.inner_text('#docsContainer')
        ff = pg.evaluate("Array.from(document.querySelectorAll('#docsContainer .ff-inline')).map(i=>i.value).includes('FF-ROUNDTRIP')")
        print('ff-inline round-trip:', ff); assert ff
        # staff slot already signed, participant cannot remove it
        assert pg.locator('#ss_sa-provider-sig.applied').count() == 1, 'provider sig not painted'
        pg.evaluate("removeSigFrom('sa-provider-sig')"); pg.wait_for_timeout(200)
        assert pg.locator('#ss_sa-provider-sig.applied').count() == 1, 'client removed staff sig!'
        # read-only check
        pe = pg.evaluate("getComputedStyle(document.querySelector('#docsContainer input:not([type=date])')).pointerEvents"); assert pe == 'none', pe
        # submit without placing signature → blocked
        pg.click('#odcsSubmitBar'); pg.wait_for_timeout(400)
        assert pg.locator('#odcsDone').count() == 0
        pg.click('text=Place my signature on all fields'); pg.wait_for_timeout(500)
        assert pg.evaluate('appliedClientIds.size') >= 3
        assert pg.locator('#ss_sa-participant-sig.applied').count() == 1
    else:
        # staff slot restored, participant can't open modal on it
        pg.click('#formTabs .nav-btn[data-target="emergency"]'); pg.wait_for_timeout(300)
        assert pg.locator('[data-sigid="sig_10"].applied').count() == 1, 'staff sig not restored'
        assert pg.evaluate("document.querySelector('[data-sigid=\"sig_10\"] .sig-clear-x') ? getComputedStyle(document.querySelector('[data-sigid=\"sig_10\"] .sig-clear-x')).display : 'none'") == 'none'
        pg.evaluate("openSigModal('sig_10','x')"); pg.wait_for_timeout(200)
        assert pg.locator('#sigModalBackdrop.show').count() == 0, 'client could open staff slot modal'
        assert pg.input_value('[data-shared="p_name"]') == 'Test Participant'
        tabs = pg.evaluate("Array.from(document.querySelectorAll('#formTabs .nav-btn')).filter(b=>b.style.display!=='none').map(b=>b.dataset.target)")
        print('visible tabs:', tabs); assert set(tabs) == {'client', 'consent', 'emergency', 'eligibility', 'intake'}
        pg.click('#odcsSubmitFloat'); pg.wait_for_timeout(400); assert pg.locator('#odcsDone').count() == 0
        # sign sig_8 (client) on emergency then "apply to all my fields"
        pg.click('[data-sigid="sig_8"]'); pg.wait_for_timeout(400)
        draw(pg, '#sigCanvas'); pg.click("#sigModalBackdrop button:has-text('Apply to all my fields')"); pg.wait_for_timeout(500)
        applied = pg.evaluate("Array.from(document.querySelectorAll('[data-sigid].applied')).map(s=>s.dataset.sigid)")
        print('applied after apply-all:', applied)
        assert 'sig_8' in applied and 'sig_3' in applied and 'sig_5' in applied
        assert 'sig_6' not in applied and 'sig_7' not in applied and 'sig_4' not in applied and 'sig_9' not in applied and 'sig_2' not in applied
        # consent choice signed individually
        pg.click('#formTabs .nav-btn[data-target="consent"]'); pg.wait_for_timeout(300)
        pg.click('[data-sigid="sig_6"]'); pg.wait_for_timeout(400); draw(pg, '#sigCanvas'); pg.click("#sigModalBackdrop button:has-text('Apply to This One')"); pg.wait_for_timeout(300)
        assert pg.locator('[data-sigid="sig_6"].applied').count() == 1
    t0 = time.time()
    pg.click('#odcsSubmitFloat')
    pg.wait_for_selector('#odcsDone', timeout=240000)
    print(f'submitted in {time.time()-t0:.0f}s')
    txt = pg.inner_text('#odcsDone')
    print(txt[:300].replace('\n', ' | '))
    assert 'has been submitted' in txt and 'participant@example.com' in txt
    # reopen → already submitted
    pg2 = ctx.new_page(); pg2.route(re.compile(r'https://(cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net)/.*'), route_cdn); attach_logging(pg2, 'reopen')
    pg2.goto(link, wait_until='load'); pg2.wait_for_selector('#odcsDone', timeout=20000)
    assert 'Already submitted' in pg2.inner_text('#odcsDone')
    print('reopen → already submitted OK')
    browser.close()
print('ALL OK', PAGE)
