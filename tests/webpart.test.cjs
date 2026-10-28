const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
let ts, JSDOM, jquery;
try {
  ts = require('typescript');
  ({ JSDOM } = require('jsdom'));
  jquery = require('jquery');
} catch {
  ts = require('../.tools/test-harness/node_modules/typescript');
  ({ JSDOM } = require('../.tools/test-harness/node_modules/jsdom'));
  jquery = require('../.tools/test-harness/node_modules/jquery');
}

const ids = ['11111111-1111-4111-8111-111111111111', '22222222-2222-4222-8222-222222222222'];
const tick = () => new Promise(resolve => setTimeout(resolve, 30));

function createPage(modern = true, content = true) {
  const control = (id, text) => modern
    ? `<div data-automation-id="CanvasControl"><div data-sp-feature-instance-id="${id}">${text}</div></div>`
    : `<div class="ControlZone" id="${id}">${text}</div>`;
  const dom = new JSDOM(`<section ${modern ? 'data-automation-id="CanvasSection"' : 'class="CanvasSection"'}>
    ${control('tabs-host', '<div id="host"></div>')}
    ${content ? ids.map((id, i) => control(id, `Content ${i + 1}`)).join('') : ''}
    </section>`, { url: 'https://example.sharepoint.com/sites/test' });
  const { window } = dom;
  const context = vm.createContext({
    window, document: window.document, HTMLElement: window.HTMLElement,
    MutationObserver: window.MutationObserver, console, setTimeout, clearTimeout,
    $: jquery(window), jQuery: jquery(window)
  });
  class BaseClientSideWebPart {}
  const cache = new Map();
  function load(file) {
    file = path.resolve(file);
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} };
    cache.set(file, module);
    const source = fs.readFileSync(file, 'utf8');
    const compiled = file.endsWith('.ts') ? ts.transpileModule(source, {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2018 }
    }).outputText : source;
    function localRequire(name) {
      if (name.endsWith('.scss')) return { default: new Proxy({}, { get: (_, key) => key }) };
      if (name.endsWith('.css')) return {};
      if (name === 'jquery') return context.$;
      if (name === '@microsoft/sp-core-library') return { DisplayMode: { Read: 1, Edit: 2 }, Version: { parse: x => x } };
      if (name === '@microsoft/sp-webpart-base') return { BaseClientSideWebPart };
      if (name === '@microsoft/sp-lodash-subset') return { escape: x => x };
      if (name === '@microsoft/sp-property-pane') return { PropertyPaneTextField: () => ({}) };
      if (name.includes('PropertyFieldCollectionData')) return {
        PropertyFieldCollectionData: (_, options) => options,
        CustomCollectionFieldType: { dropdown: 'dropdown', string: 'string' }
      };
      if (name === 'ModernHillbillyTabsWebPartStrings') return {};
      if (name.startsWith('.')) {
        const resolved = path.resolve(path.dirname(file), name);
        return load(fs.existsSync(resolved) ? resolved : `${resolved}.ts`);
      }
      throw new Error(`Unexpected dependency ${name}`);
    }
    vm.runInContext(`(function(require, module, exports) { ${compiled}\n})`, context, { filename: file })(localRequire, module, module.exports);
    return module.exports;
  }
  const WebPart = load(path.join(__dirname, '../src/webparts/modernHillbillyTabs/ModernHillbillyTabsWebPart.ts')).default;
  const webpart = new WebPart();
  webpart.domElement = window.document.getElementById('host');
  webpart.instanceId = 'tabs-test';
  webpart.displayMode = 1;
  webpart.properties = {
    sectionClass: 'CanvasSection', webpartClass: 'ControlZone',
    tabData: ids.map((id, i) => ({ WebPartID: id, TabLabel: `Tab ${i + 1}` }))
  };
  const target = i => modern
    ? window.document.querySelector(`[data-sp-feature-instance-id="${ids[i]}"]`).parentElement
    : window.document.getElementById(ids[i]);
  const hidden = element => element.hidden || element.style.display === 'none' || element.classList.contains('addui-Tabs-content') && !element.classList.contains('addui-Tabs-active');
  return { window, webpart, target, hidden, control, close: () => { webpart.onDispose?.(); window.close(); } };
}

for (const modern of [true, false]) {
  test(`only selected content is visible and switching works (${modern ? 'modern markup' : 'legacy markup'})`, async () => {
    const page = createPage(modern);
    try {
      page.webpart.render();
      await tick();
      assert.equal(page.hidden(page.target(0)), false, 'first tab content must be visible');
      assert.equal(page.hidden(page.target(1)), true, 'inactive web part must be hidden');
      const tabs = page.webpart.domElement.querySelectorAll('[role="tab"], .addui-Tabs-tab');
      tabs[1].click();
      assert.equal(page.hidden(page.target(0)), true);
      assert.equal(page.hidden(page.target(1)), false, 'selected tab content must appear');
    } finally { page.close(); }
  });
}

test('late SharePoint rendering still applies the selected tab', async () => {
  const page = createPage(true, false);
  try {
    page.webpart.render();
    page.webpart.domElement.closest('section').insertAdjacentHTML('beforeend', ids.map((id, i) => page.control(id, `Late ${i}`)).join(''));
    await tick();
    assert.equal(page.hidden(page.target(1)), true, 'late inactive content must be hidden');
  } finally { page.close(); }
});

function originalState(element) {
  return {
    attributes: Object.fromEntries(['id', 'role', 'aria-labelledby', 'aria-hidden', 'hidden', 'tabindex']
      .map(name => [name, element.getAttribute(name)])),
    display: element.style.getPropertyValue('display'),
    displayPriority: element.style.getPropertyPriority('display')
  };
}

test('render, read/edit transitions, and disposal preserve canvas nodes and restore original state', async () => {
  const page = createPage();
  try {
    const content = [page.target(0), page.target(1)];
    const parents = content.map(element => element.parentElement);
    content[0].id = 'existing-panel-id';
    content[0].setAttribute('role', 'region');
    content[0].setAttribute('aria-labelledby', 'original-heading');
    content[0].setAttribute('aria-hidden', 'false');
    content[0].setAttribute('tabindex', '-1');
    content[0].style.setProperty('display', 'grid', 'important');
    content[1].setAttribute('hidden', 'hidden');
    content[1].style.setProperty('display', 'flex');
    const before = content.map(originalState);

    page.webpart.render();
    page.webpart.domElement.querySelectorAll('[role="tab"]')[1].click();
    page.webpart.render();
    await tick();
    content.forEach((element, index) => {
      assert.equal(element.parentElement, parents[index], 'SharePoint-owned canvas nodes must never be reparented');
      assert.equal(element.isConnected, true, 're-render must not delete canvas content');
    });

    page.webpart.displayMode = 2;
    page.webpart.render();
    await tick();
    content.forEach((element, index) => assert.deepEqual(originalState(element), before[index], 'edit mode must restore author state'));
    assert.equal(page.webpart.domElement.querySelector('[role="tab"]'), null);

    page.webpart.displayMode = 1;
    page.webpart.render();
    page.webpart.onDispose();
    await tick();
    content.forEach((element, index) => {
      assert.deepEqual(originalState(element), before[index], 'disposal must restore author state');
      assert.equal(element.parentElement, parents[index]);
    });
  } finally { page.close(); }
});

test('SharePoint replacing a canvas control reapplies the current tab without moving content', async () => {
  const page = createPage();
  try {
    page.webpart.render();
    const removed = page.target(1);
    const parent = removed.parentElement;
    removed.outerHTML = page.control(ids[1], 'Replacement content');
    await tick();
    const replacement = page.target(1);
    assert.notEqual(replacement, removed);
    assert.equal(replacement.parentElement, parent);
    assert.equal(page.hidden(replacement), true, 'replacement for inactive tab must be hidden');
    assert.equal(removed.style.display, '', 'detached replaced node must be released');
    page.webpart.domElement.querySelectorAll('[role="tab"]')[1].click();
    assert.equal(page.hidden(replacement), false);
    assert.equal(page.hidden(page.target(0)), true);
  } finally { page.close(); }
});

test('same-node SharePoint visibility resets are corrected while the tab is inactive', async () => {
  const page = createPage();
  try {
    page.webpart.render();
    const inactive = page.target(1);
    inactive.hidden = false;
    inactive.style.setProperty('display', 'block', 'important');
    await tick();
    assert.equal(page.hidden(inactive), true, 'SharePoint resetting styles must not expose inactive content');
    assert.equal(inactive.getAttribute('aria-hidden'), 'true');
  } finally { page.close(); }
});

test('tabs expose matching ARIA relationships and keyboard navigation skips unavailable content', async () => {
  const page = createPage();
  try {
    page.webpart.properties.tabData.splice(1, 0, { WebPartID: 'not-rendered', TabLabel: 'Unavailable' });
    page.webpart.render();
    const tabs = page.webpart.domElement.querySelectorAll('[role="tab"]');
    assert.ok(page.webpart.domElement.querySelector('[role="tablist"]').getAttribute('aria-label'));
    assert.equal(tabs[1].disabled, true);
    for (const index of [0, 2]) {
      const panel = page.window.document.getElementById(tabs[index].getAttribute('aria-controls'));
      assert.equal(panel.getAttribute('role'), 'tabpanel');
      assert.equal(panel.getAttribute('aria-labelledby'), tabs[index].id);
    }
    tabs[0].focus();
    const press = key => page.window.document.activeElement.dispatchEvent(new page.window.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
    press('ArrowRight');
    assert.equal(page.window.document.activeElement, tabs[2]);
    assert.equal(tabs[2].getAttribute('aria-selected'), 'true');
    assert.equal(tabs[2].tabIndex, 0);
    assert.equal(tabs[0].tabIndex, -1);
    assert.equal(page.hidden(page.target(1)), false);
    press('ArrowRight');
    assert.equal(page.window.document.activeElement, tabs[0], 'keyboard navigation must wrap');
    press('End');
    assert.equal(page.window.document.activeElement, tabs[2]);
    press('Home');
    assert.equal(page.window.document.activeElement, tabs[0]);
    press('ArrowLeft');
    assert.equal(page.window.document.activeElement, tabs[2]);
  } finally { page.close(); }
});

test('tab labels remain literal text and malformed saved IDs do not become selectors', async () => {
  const page = createPage();
  try {
    const label = '<img src=x onerror="window.unexpected = true"> & content';
    page.webpart.properties.tabData[0].TabLabel = label;
    page.webpart.properties.tabData.push({ WebPartID: 'bad"id\\]', TabLabel: 'Missing' });
    assert.doesNotThrow(() => page.webpart.render());
    const tabs = page.webpart.domElement.querySelectorAll('[role="tab"]');
    assert.equal(tabs[0].textContent, label);
    assert.equal(tabs[0].querySelector('img'), null);
    assert.equal(tabs[2].disabled, true);
    assert.equal(page.window.unexpected, undefined);
    assert.equal(page.hidden(page.target(1)), true);
  } finally { page.close(); }
});

test('configured content and property options remain scoped to the host section', async () => {
  const page = createPage();
  try {
    const outsideId = '33333333-3333-4333-8333-333333333333';
    page.window.document.body.insertAdjacentHTML('beforeend', `<section data-automation-id="CanvasSection">${page.control(outsideId, 'Other section')}</section>`);
    const outside = page.window.document.querySelector(`[data-sp-feature-instance-id="${outsideId}"]`).parentElement;
    const before = originalState(outside);
    page.webpart.properties.tabData.push({ WebPartID: outsideId, TabLabel: 'Outside' });
    page.webpart.render();
    await tick();
    const tabs = page.webpart.domElement.querySelectorAll('[role="tab"]');
    assert.equal(tabs[2].disabled, true);
    assert.deepEqual(originalState(outside), before, 'another section must not be modified');
    const field = page.webpart.getPropertyPaneConfiguration().pages[0].groups[0].groupFields[0].fields[0];
    assert.match(field.options.find(option => option.key === outsideId).text, /unavailable in this section/);
  } finally { page.close(); }
});

test('unrecognized section markup fails closed without hiding page content', async () => {
  const page = createPage();
  try {
    page.webpart.domElement.closest('section').removeAttribute('data-automation-id');
    page.webpart.render();
    await tick();
    assert.equal(page.hidden(page.target(0)), false);
    assert.equal(page.hidden(page.target(1)), false);
    assert.ok(Array.from(page.webpart.domElement.querySelectorAll('[role="tab"]')).every(tab => tab.disabled));
    assert.match(page.webpart.domElement.querySelector('[role="status"]').textContent, /Cannot find this page section/);
  } finally { page.close(); }
});

test('legacy control_GUID settings match feature-instance IDs and retain saved property keys', async () => {
  const page = createPage();
  try {
    page.webpart.properties.tabData = ids.map((id, i) => ({ WebPartID: `control_${id.toUpperCase()}`, TabLabel: `Saved ${i}` }));
    page.webpart.render();
    assert.equal(page.hidden(page.target(1)), true);
    const tabs = page.webpart.domElement.querySelectorAll('[role="tab"]');
    assert.equal(tabs[1].disabled, false);
    tabs[1].click();
    assert.equal(page.hidden(page.target(1)), false);
    const options = page.webpart.getPropertyPaneConfiguration().pages[0].groups[0].groupFields[0].fields[0].options;
    assert.equal(options.length, 2);
    assert.equal(options[0].key, page.webpart.properties.tabData[0].WebPartID);
    assert.equal(options[1].key, page.webpart.properties.tabData[1].WebPartID);
  } finally { page.close(); }
});

test('saved legacy wrapper IDs still match inside newer CanvasControl wrappers', async () => {
  const page = createPage();
  try {
    const outer = page.target(1);
    outer.innerHTML = `<div class="ControlZone" id="legacy-wrapper-id"><div data-sp-feature-instance-id="${ids[1]}">Mixed markup content</div></div>`;
    const inner = page.window.document.getElementById('legacy-wrapper-id');
    page.webpart.properties.tabData[1].WebPartID = 'legacy-wrapper-id';
    page.webpart.render();
    await tick();
    const tabs = page.webpart.domElement.querySelectorAll('[role="tab"]');
    assert.equal(tabs[1].disabled, false, 'a nested legacy wrapper remains a valid saved identity');
    assert.equal(page.hidden(outer), true, 'visibility must apply to the whole outer canvas control');
    assert.equal(inner.parentElement, outer, 'mixed canvas markup must retain its structure');
    tabs[1].click();
    assert.equal(page.hidden(outer), false);
    assert.equal(page.hidden(page.target(0)), true);
    const options = page.webpart.getPropertyPaneConfiguration().pages[0].groups[0].groupFields[0].fields[0].options;
    assert.equal(options.length, 2, 'nested markup must produce one option per canvas control');
    assert.equal(options[1].key, 'legacy-wrapper-id');
  } finally { page.close(); }
});

test('active web part layout changes survive observer updates and tab switching', async () => {
  const page = createPage();
  try {
    page.webpart.render();
    const active = page.target(0);
    active.style.display = 'grid';
    await tick();
    assert.equal(active.style.display, 'grid', 'observer must preserve active SharePoint layout changes');
    const tabs = page.webpart.domElement.querySelectorAll('[role="tab"]');
    tabs[1].click();
    assert.equal(page.hidden(active), true);
    tabs[0].click();
    await tick();
    assert.equal(page.hidden(active), false);
    assert.equal(active.style.display, 'grid', 'activation must restore the layout captured when hiding');
  } finally { page.close(); }
});

test('multiple tab instances cannot take ownership of the same canvas content', async () => {
  const page = createPage();
  let second;
  try {
    const section = page.webpart.domElement.closest('section');
    section.insertAdjacentHTML('beforeend', page.control('second-host', '<div id="second-host-root"></div>'));
    second = new page.webpart.constructor();
    second.domElement = page.window.document.getElementById('second-host-root');
    second.instanceId = 'second-tabs';
    second.displayMode = 1;
    second.properties = { ...page.webpart.properties, tabData: page.webpart.properties.tabData.map(tab => ({ ...tab })) };
    page.webpart.render();
    second.render();
    await tick();
    assert.ok(Array.from(second.domElement.querySelectorAll('[role="tab"]')).every(tab => tab.disabled));
    assert.equal(page.target(0).getAttribute('aria-labelledby'), 'tabs-test-tab-0');
    assert.equal(page.hidden(page.target(1)), true);
    page.webpart.domElement.querySelectorAll('[role="tab"]')[1].click();
    await tick();
    assert.equal(page.hidden(page.target(0)), true);
    assert.equal(page.hidden(page.target(1)), false);
    let statusMutations = 0;
    const observer = new page.window.MutationObserver(records => {
      statusMutations += records.filter(record => record.target.classList.contains('hillbilly-tabs-status')).length;
    });
    observer.observe(section, { subtree: true, attributes: true, attributeFilter: ['hidden'] });
    try {
      await tick();
      statusMutations = 0;
      await tick();
      assert.equal(statusMutations, 0, 'multiple observers must settle instead of triggering each other indefinitely');
    } finally { observer.disconnect(); }
  } finally { second?.onDispose(); page.close(); }
});

test('different saved aliases for one control cannot hide the selected content', async () => {
  const page = createPage();
  try {
    const first = page.target(0);
    first.id = 'legacy-first-wrapper';
    page.webpart.properties.tabData.splice(1, 0, { WebPartID: 'legacy-first-wrapper', TabLabel: 'Duplicate alias' });
    page.webpart.render();
    const tabs = page.webpart.domElement.querySelectorAll('[role="tab"]');
    assert.equal(tabs[0].getAttribute('aria-selected'), 'true');
    assert.equal(tabs[1].disabled, true);
    assert.equal(page.hidden(first), false);
    tabs[2].click();
    assert.equal(page.hidden(first), true);
    assert.equal(page.hidden(page.target(1)), false);
  } finally { page.close(); }
});

module.exports = { createPage, ids, tick };
