const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ts = require('typescript');
const { JSDOM } = require('jsdom');

const repo = 'https://github.com/drghassanabukar/modern-hillbilly-tabs-365';
const cacheKey = 'modern-hillbilly-tabs-365:release-check:v1';
const tick = () => new Promise(resolve => setTimeout(resolve, 15));

function release(version = '1.2.0') {
  return {
    tag_name: 'v' + version, html_url: repo + '/releases/tag/v' + version,
    draft: false, prerelease: false,
    assets: [{ name: 'modern-hillbilly-tabs.sppkg', state: 'uploaded', browser_download_url: repo + '/releases/download/v' + version + '/modern-hillbilly-tabs.sppkg' }]
  };
}

function response(payload, status = 200, headers = {}) {
  return { ok: status >= 200 && status < 300, status, headers: { get: name => headers[name.toLowerCase()] ?? null }, json: async () => payload };
}

function page() {
  const dom = new JSDOM('<div id="first"></div><div id="second"></div>', { url: 'https://example.sharepoint.com/sites/test' });
  const { window } = dom;
  const module = { exports: {} };
  const source = ts.transpileModule(fs.readFileSync(path.join(__dirname, '../src/webparts/modernHillbillyTabs/UpdateNotice.ts'), 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2018 }
  }).outputText;
  const context = vm.createContext({ document: window.document, window, URL: window.URL, AbortController: window.AbortController, console });
  vm.runInContext('(function(module,exports){' + source + '\n})', context)(module, module.exports);
  const notices = [];
  const make = (version = '1.1.0', options = {}, id = 'first') => {
    const host = window.document.getElementById(id);
    const notice = new module.exports.UpdateNotice(host, version, { alwaysShow: true, enabled: true, ...options });
    notices.push(notice);
    return { notice, host };
  };
  return { window, make, close: () => { notices.forEach(notice => notice.dispose()); window.close(); } };
}

test('an available release shows versions and safe installation/reload actions without sending credentials or page content', async () => {
  const fixture = page();
  const calls = [];
  let reloads = 0;
  try {
    const { host } = fixture.make('1.1.0', {
      fetcher: async (...args) => { calls.push(args); return response(release()); },
      reload: () => reloads++
    });
    await tick();
    assert.equal(host.hidden, false);
    assert.equal(host.dataset.state, 'available');
    assert.match(host.textContent, /Loaded version: 1\.1\.0/);
    assert.match(host.textContent, /Version 1\.2\.0 is available/);
    assert.match(host.textContent, /Reloading does not install an update/);
    assert.match(host.textContent, /Save or publish your page changes before reloading/);
    assert.equal(calls.length, 1);
    assert.equal(calls[0][0], 'https://api.github.com/repos/drghassanabukar/modern-hillbilly-tabs-365/releases/latest');
    const options = calls[0][1];
    assert.equal(options.credentials, 'omit');
    assert.equal(options.referrerPolicy, 'no-referrer');
    assert.equal(options.method, 'GET');
    assert.equal(options.body, undefined);
    assert.equal(options.headers.Authorization, undefined);
    assert.ok(options.signal);
    const links = Array.from(host.querySelectorAll('a'));
    assert.equal(links.find(link => link.textContent === 'Download update').href, repo + '/releases/download/v1.2.0/modern-hillbilly-tabs.sppkg');
    assert.equal(links.find(link => link.textContent === 'Release notes').href, repo + '/releases/tag/v1.2.0');
    assert.ok(links.every(link => link.rel === 'noopener noreferrer'));
    assert.equal(reloads, 0, 'update discovery must never reload automatically');
    Array.from(host.querySelectorAll('button')).find(button => /Already deployed/.test(button.textContent)).click();
    assert.equal(reloads, 1);
  } finally { fixture.close(); }
});

test('numeric versions treat three and four segments equally and compare 1.10 after 1.9', async () => {
  for (const [loaded, latest, expected] of [['1.2.0.0','1.2.0','current'], ['1.9.0','1.10.0','available'], ['1.10.0','1.9.0','current']]) {
    const fixture = page();
    try {
      const { host } = fixture.make(loaded, { fetcher: async () => response(release(latest)) });
      await tick();
      assert.equal(host.dataset.state, expected);
    } finally { fixture.close(); }
  }
});

test('read-mode admin notices stay hidden when current and appear only for an available update', async () => {
  for (const [latest, hidden] of [['1.1.0', true], ['1.2.0', false]]) {
    const fixture = page();
    try {
      const { host } = fixture.make('1.1.0', { alwaysShow: false, fetcher: async () => response(release(latest)) });
      assert.equal(host.hidden, true, 'checking must not add a notice to the read view');
      await tick();
      assert.equal(host.hidden, hidden);
    } finally { fixture.close(); }
  }
});

test('disabled checking displays the edit-mode version but makes no network request', async () => {
  const fixture = page();
  let calls = 0;
  try {
    const { host } = fixture.make('1.1.0', { enabled: false, fetcher: async () => { calls++; return response(release()); } });
    await tick();
    assert.equal(calls, 0);
    assert.equal(host.dataset.state, 'disabled');
    assert.match(host.textContent, /Loaded version: 1\.1\.0/);
    assert.equal(host.querySelector('button').disabled, true);
  } finally { fixture.close(); }
});

test('instances share a single pending release request and the one-hour cache', async () => {
  const fixture = page();
  let calls = 0;
  let finish;
  const fetcher = () => { calls++; return new Promise(resolve => { finish = resolve; }); };
  try {
    const first = fixture.make('1.1.0', { fetcher });
    const second = fixture.make('1.1.0', { fetcher }, 'second');
    assert.equal(calls, 1);
    finish(response(release()));
    await tick();
    assert.equal(first.host.dataset.state, 'available');
    assert.equal(second.host.dataset.state, 'available');
    first.notice.dispose();
    const cached = fixture.make('1.1.0', { fetcher });
    await tick();
    assert.equal(calls, 1);
    assert.equal(cached.host.dataset.state, 'available');
    assert.ok(fixture.window.sessionStorage.getItem(cacheKey));
  } finally { fixture.close(); }
});

test('manual checks bypass fresh cache after the minimum interval and respect immediate backoff', async () => {
  const fixture = page();
  let calls = 0;
  let now = 1000000;
  try {
    const { host } = fixture.make('1.1.0', { now: () => now, fetcher: async () => { calls++; return response(release(calls === 1 ? '1.1.0' : '1.2.0')); } });
    await tick();
    host.querySelector('button').click();
    await tick();
    assert.equal(calls, 1);
    assert.match(host.textContent, /wait at least a minute/);
    now += 61000;
    host.querySelector('button').click();
    await tick();
    assert.equal(calls, 2);
    assert.equal(host.dataset.state, 'available');
  } finally { fixture.close(); }
});

test('server rate limits prevent repeated requests until the Retry-After delay expires', async () => {
  const fixture = page();
  let calls = 0;
  let now = 1000000;
  try {
    const { host } = fixture.make('1.1.0', { now: () => now, fetcher: async () => { calls++; return calls === 1 ? response({}, 429, { 'retry-after': '180' }) : response(release()); } });
    await tick();
    assert.match(host.textContent, /rate limited/);
    now += 61000;
    host.querySelector('button').click();
    await tick();
    assert.equal(calls, 1);
    now += 120000;
    host.querySelector('button').click();
    await tick();
    assert.equal(calls, 2);
  } finally { fixture.close(); }
});

test('network errors and unavailable fetch are nonfatal and retain a safe releases link', async () => {
  for (const fetcher of [async () => { throw new Error('Network unavailable'); }, undefined]) {
    const fixture = page();
    try {
      let result;
      assert.doesNotThrow(() => { result = fixture.make('1.1.0', { fetcher }); });
      await tick();
      assert.equal(result.host.dataset.state, 'error');
      assert.equal(result.host.querySelector('a').href, repo + '/releases/latest');
      assert.equal(result.host.querySelector('button').disabled, false);
    } finally { fixture.close(); }
  }
});

test('an unknown loaded version is shown safely without a network request or update claim', async () => {
  const fixture = page();
  let calls = 0;
  try {
    const { host } = fixture.make('unknown', { fetcher: async () => { calls++; return response(release()); } });
    await tick();
    assert.equal(calls, 0);
    assert.equal(host.dataset.state, 'error');
    assert.match(host.textContent, /Loaded version: unknown/);
    assert.match(host.textContent, /loaded version could not be determined/);
    assert.equal(Array.from(host.querySelectorAll('a')).some(link => link.textContent === 'Download update'), false);
  } finally { fixture.close(); }
});

test('a nonresponsive request aborts at the configured timeout and offers a manual retry', async () => {
  const fixture = page();
  const originalTimeout = fixture.window.setTimeout.bind(fixture.window);
  let signal;
  let timeout;
  fixture.window.setTimeout = (callback, milliseconds) => { timeout = milliseconds; return originalTimeout(callback, 2); };
  try {
    const { host } = fixture.make('1.1.0', { fetcher: (_, options) => { signal = options.signal; return new Promise(() => {}); } });
    await tick();
    assert.equal(timeout, 5000);
    assert.equal(signal.aborted, true);
    assert.equal(host.dataset.state, 'error');
    assert.match(host.textContent, /timed out/);
  } finally { fixture.close(); }
});

test('blocked browser storage does not prevent update checks', async () => {
  const fixture = page();
  Object.defineProperty(fixture.window, 'sessionStorage', { get: () => { throw new Error('Storage blocked'); } });
  try {
    const { host } = fixture.make('1.1.0', { fetcher: async () => response(release()) });
    await tick();
    assert.equal(host.dataset.state, 'available');
  } finally { fixture.close(); }
});

test('a stalled release response body also reaches the configured timeout', async () => {
  const fixture = page();
  const originalTimeout = fixture.window.setTimeout.bind(fixture.window);
  fixture.window.setTimeout = (callback) => originalTimeout(callback, 2);
  try {
    const { host } = fixture.make('1.1.0', { fetcher: async () => ({ ...response(release()), json: () => new Promise(() => {}) }) });
    await tick();
    assert.equal(host.dataset.state, 'error');
    assert.match(host.textContent, /timed out/);
  } finally { fixture.close(); }
});

test('a validated session cache survives a fresh module without fetching again', async () => {
  const fixture = page();
  let calls = 0;
  try {
    fixture.window.sessionStorage.setItem(cacheKey, JSON.stringify({ checkedAt: 1000000, release: { tag: 'v1.2.0', version: '1.2.0', releaseUrl: repo + '/releases/tag/v1.2.0', downloadUrl: repo + '/releases/download/v1.2.0/modern-hillbilly-tabs.sppkg' } }));
    const { host } = fixture.make('1.1.0', { now: () => 1000001, fetcher: async () => { calls++; return response(release()); } });
    await tick();
    assert.equal(calls, 0);
    assert.equal(host.dataset.state, 'available');
  } finally { fixture.close(); }
});

test('invalid cached URLs and impossible stored backoffs are ignored safely', async () => {
  const fixture = page();
  let calls = 0;
  try {
    fixture.window.sessionStorage.setItem(cacheKey, JSON.stringify({ checkedAt: 1000000, release: { tag: 'v1.2.0', version: '1.2.0', releaseUrl: 'https://example.invalid', downloadUrl: 'javascript:unexpected()' } }));
    fixture.window.sessionStorage.setItem(cacheKey + ':retry', 'Infinity');
    const { host } = fixture.make('1.1.0', { now: () => 1000001, fetcher: async () => { calls++; return response(release()); } });
    await tick();
    assert.equal(calls, 1);
    assert.equal(host.dataset.state, 'available');
    assert.equal(host.querySelector('a[href^="javascript:"]'), null);
  } finally { fixture.close(); }
});

test('unknown versions, prereleases, unsafe release URLs, and mismatched assets never create an update action', async () => {
  const invalid = [
    { ...release(), tag_name: 'next' },
    { ...release(), tag_name: 'v1.2.0-beta' },
    { ...release(), prerelease: true },
    { ...release(), draft: true },
    { ...release(), html_url: 'javascript:unexpected()' },
    { ...release(), html_url: 'https://github.com.evil.invalid/drghassanabukar/modern-hillbilly-tabs-365/releases/tag/v1.2.0' },
    { ...release(), assets: [{ name: 'other.sppkg', state: 'uploaded', browser_download_url: repo + '/releases/download/v1.2.0/other.sppkg' }] },
    { ...release(), assets: [{ name: 'modern-hillbilly-tabs.sppkg', state: 'uploaded', browser_download_url: 'https://example.invalid/payload.sppkg' }] }
  ];
  for (const payload of invalid) {
    const fixture = page();
    try {
      const { host } = fixture.make('1.1.0', { fetcher: async () => response(payload) });
      await tick();
      assert.equal(host.dataset.state, 'error');
      assert.ok(!Array.from(host.querySelectorAll('a')).some(link => link.textContent === 'Download update'));
      assert.equal(host.querySelector('a').href, repo + '/releases/latest');
    } finally { fixture.close(); }
  }
});

test('disposing a notice prevents late network results and reload actions from changing the page', async () => {
  const fixture = page();
  let finish;
  try {
    const { notice, host } = fixture.make('1.1.0', { fetcher: () => new Promise(resolve => { finish = resolve; }) });
    const before = host.innerHTML;
    notice.dispose();
    finish(response(release()));
    await tick();
    assert.equal(host.innerHTML, before);
    assert.equal(host.querySelector('button').disabled, true);
  } finally { fixture.close(); }
});

test('release body HTML is never inserted into the notice', async () => {
  const fixture = page();
  try {
    const { host } = fixture.make('1.1.0', { fetcher: async () => response({ ...release(), body: '<img src=x onerror="unexpected()">' }) });
    await tick();
    assert.equal(host.querySelector('img'), null);
    assert.equal(host.dataset.state, 'available');
  } finally { fixture.close(); }
});

test('check and reload keyboard actions survive host cancellation and activate once without repeats', async () => {
  const fixture = page();
  let calls = 0;
  let reloads = 0;
  let checkClicks = 0;
  let now = 1000000;
  const cancelNativeKey = event => event.preventDefault();
  fixture.window.document.body.addEventListener('keydown', cancelNativeKey, true);
  try {
    const { notice, host } = fixture.make('1.1.0', {
      now: () => now,
      fetcher: async () => { calls++; return response(release(calls === 1 ? '1.1.0' : '1.2.0')); },
      reload: () => reloads++
    });
    await tick();
    host.addEventListener('click', event => {
      if (event.target.textContent === 'Check for updates') checkClicks++;
    });
    const press = (button, key, repeat = false) => {
      const event = new fixture.window.KeyboardEvent('keydown', { key, repeat, bubbles: true, cancelable: true });
      button.dispatchEvent(event);
      assert.equal(event.defaultPrevented, true, 'native activation must be suppressed to avoid duplicate clicks');
    };
    now += 61000;
    press(host.querySelector('button'), 'Enter');
    await tick();
    assert.equal(calls, 2, 'one initial request and one keyboard-triggered request');
    assert.equal(checkClicks, 1);
    press(host.querySelector('button'), 'Enter', true);
    await tick();
    assert.equal(calls, 2);
    assert.equal(checkClicks, 1);
    const reload = Array.from(host.querySelectorAll('button')).find(button => /Already deployed/.test(button.textContent));
    press(reload, 'Enter');
    assert.equal(reloads, 1);
    press(reload, 'Enter', true);
    assert.equal(reloads, 1);
    press(reload, ' ');
    reload.dispatchEvent(new fixture.window.KeyboardEvent('keyup', { key: ' ', bubbles: true, cancelable: true }));
    assert.equal(reloads, 2, 'Space keydown activates once and keyup does not duplicate it');
    notice.dispose();
    fixture.window.document.body.removeEventListener('keydown', cancelNativeKey, true);
    const disposedKey = new fixture.window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    reload.dispatchEvent(disposedKey);
    assert.equal(disposedKey.defaultPrevented, false, 'disposal removes the delegated handler');
    assert.equal(reloads, 2);
  } finally {
    fixture.window.document.body.removeEventListener('keydown', cancelNativeKey, true);
    fixture.close();
  }
});

test('update links retain native navigation when the canvas cancels bubbling clicks', async () => {
  const fixture = page();
  let cancelled = 0;
  const cancelCanvasClick = event => { cancelled++; event.preventDefault(); };
  fixture.window.document.body.addEventListener('click', cancelCanvasClick);
  try {
    const { notice, host } = fixture.make('1.1.0', { fetcher: async () => response(release()) });
    await tick();
    const link = Array.from(host.querySelectorAll('a')).find(anchor => anchor.textContent === 'Download update');
    const click = new fixture.window.MouseEvent('click', { bubbles: true, cancelable: true });
    link.dispatchEvent(click);
    assert.equal(click.defaultPrevented, false, 'the download must keep its native anchor action');
    assert.equal(cancelled, 0, 'the canvas bubble listener must not receive addon actions');
    assert.equal(link.target, '_blank');
    assert.equal(link.href, repo + '/releases/download/v1.2.0/modern-hillbilly-tabs.sppkg');
    notice.dispose();
    const disposedClick = new fixture.window.MouseEvent('click', { bubbles: true, cancelable: true });
    link.dispatchEvent(disposedClick);
    assert.equal(cancelled, 1, 'disposal removes the click shield');
    assert.equal(disposedClick.defaultPrevented, true);
  } finally {
    fixture.window.document.body.removeEventListener('click', cancelCanvasClick);
    fixture.close();
  }
});
