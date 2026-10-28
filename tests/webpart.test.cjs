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
    AbortController: window.AbortController, URL: window.URL,
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
      if (name === '@microsoft/sp-page-context') return { SPPermission: { manageWeb: 'manageWeb' } };
      if (name === '@microsoft/sp-lodash-subset') return { escape: x => x };
      if (name === '@microsoft/sp-property-pane') return {
        PropertyPaneTextField: () => ({}),
        PropertyPaneToggle: (targetProperty, options) => ({ type: 'toggle', targetProperty, ...options }),
        PropertyPaneDropdown: (targetProperty, options) => ({ type: 'dropdown', targetProperty, ...options })
      };
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
  webpart.manifest = { version: '1.2.0' };
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
    assert.ok(Array.from(page.webpart.domElement.querySelectorAll('[role="tab"]')).every(tab => tab.closest('.hillbilly-tabs-preview')),
      'edit mode may render local preview tabs only');

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

const tabDesigns = ['underline', 'pills', 'outlined', 'segmented', 'blocks', 'minimal'];

test('existing saved pages use the original underline design and start alignment', () => {
  const page = createPage();
  try {
    assert.equal(page.webpart.properties.tabDesign, undefined);
    assert.equal(page.webpart.properties.tabAlignment, undefined);
    page.webpart.render();
    const tablist = page.webpart.domElement.querySelector('.hillbilly-tabs-list');
    assert.equal(tablist.dataset.tabDesign, 'underline');
    assert.equal(tablist.dataset.tabAlignment, 'start');
    assert.equal(page.hidden(page.target(0)), false);
    assert.equal(page.hidden(page.target(1)), true);
  } finally { page.close(); }
});

test('the Appearance settings offer every design and alignment without changing saved tab options', () => {
  const page = createPage();
  try {
    page.webpart.properties.tabDesign = 'pills';
    page.webpart.properties.tabAlignment = 'stretch';
    const groups = page.webpart.getPropertyPaneConfiguration().pages[0].groups;
    const appearance = groups.find(group => group.groupName === 'Appearance');
    assert.ok(appearance, 'page editors must be able to find Appearance settings');
    const design = appearance.groupFields.find(field => field.targetProperty === 'tabDesign');
    const alignment = appearance.groupFields.find(field => field.targetProperty === 'tabAlignment');
    assert.equal(design.type, 'dropdown');
    assert.deepEqual(Array.from(design.options, option => option.key), tabDesigns);
    assert.ok(design.options.every(option => typeof option.text === 'string' && option.text.length > 0));
    assert.equal(design.selectedKey, 'pills');
    assert.deepEqual(Array.from(alignment.options, option => option.key), ['start', 'center', 'stretch']);
    assert.equal(alignment.selectedKey, 'stretch');
    const configuredParts = groups[0].groupFields[0].fields[0].options;
    assert.deepEqual(Array.from(configuredParts, option => option.key), ids);
  } finally { page.close(); }
});

for (const design of tabDesigns) {
  test(`${design} design preserves mouse switching, keyboard navigation, and panel relationships`, () => {
    const page = createPage();
    try {
      page.webpart.properties.tabDesign = design;
      page.webpart.properties.tabAlignment = 'stretch';
      page.webpart.render();
      const tablist = page.webpart.domElement.querySelector('.hillbilly-tabs-list');
      assert.equal(tablist.dataset.tabDesign, design);
      assert.equal(tablist.dataset.tabAlignment, 'stretch');
      const tabs = tablist.querySelectorAll('[role="tab"]');
      tabs[1].click();
      assert.equal(page.hidden(page.target(0)), true);
      assert.equal(page.hidden(page.target(1)), false);
      assert.equal(page.window.document.getElementById(tabs[1].getAttribute('aria-controls')), page.target(1));
      tabs[1].focus();
      tabs[1].dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true, cancelable: true }));
      assert.equal(page.window.document.activeElement, tabs[0]);
      assert.equal(tabs[0].getAttribute('aria-selected'), 'true');
      assert.equal(page.hidden(page.target(0)), false);
      assert.equal(page.hidden(page.target(1)), true);
    } finally { page.close(); }
  });
}

test('invalid persisted appearance values fall back safely', () => {
  const page = createPage();
  try {
    page.webpart.properties.tabDesign = 'pills" onclick="unexpected()';
    page.webpart.properties.tabAlignment = '__proto__';
    page.webpart.render();
    const tablist = page.webpart.domElement.querySelector('.hillbilly-tabs-list');
    assert.equal(tablist.dataset.tabDesign, 'underline');
    assert.equal(tablist.dataset.tabAlignment, 'start');
    assert.equal(tablist.hasAttribute('onclick'), false);
    assert.equal(page.hidden(page.target(1)), true);
    const appearance = page.webpart.getPropertyPaneConfiguration().pages[0].groups.find(group => group.groupName === 'Appearance');
    assert.equal(appearance.groupFields.find(field => field.targetProperty === 'tabDesign').selectedKey, 'underline');
    assert.equal(appearance.groupFields.find(field => field.targetProperty === 'tabAlignment').selectedKey, 'start');
  } finally { page.close(); }
});

test('appearance updates keep saved labels and SharePoint content in place', () => {
  const page = createPage();
  try {
    const originalParts = [page.target(0), page.target(1)];
    const originalParents = originalParts.map(part => part.parentElement);
    const savedSettings = JSON.stringify(page.webpart.properties.tabData);
    page.webpart.render();
    page.webpart.properties.tabDesign = 'outlined';
    page.webpart.properties.tabAlignment = 'center';
    page.webpart.render();
    const tablist = page.webpart.domElement.querySelector('.hillbilly-tabs-list');
    assert.equal(tablist.dataset.tabDesign, 'outlined');
    assert.equal(tablist.dataset.tabAlignment, 'center');
    assert.equal(JSON.stringify(page.webpart.properties.tabData), savedSettings);
    originalParts.forEach((part, index) => {
      assert.equal(page.target(index), part);
      assert.equal(part.parentElement, originalParents[index]);
    });
  } finally { page.close(); }
});

test('different web part instances keep their own design, alignment, and selection', () => {
  const page = createPage();
  let second;
  try {
    const secondIds = ['33333333-3333-4333-8333-333333333333', '44444444-4444-4444-8444-444444444444'];
    page.window.document.body.insertAdjacentHTML('beforeend', `<section data-automation-id="CanvasSection">${page.control('second-tabs-host', '<div id="second-tabs-root"></div>')}${secondIds.map(id => page.control(id, 'Separate section content')).join('')}</section>`);
    second = new page.webpart.constructor();
    second.domElement = page.window.document.getElementById('second-tabs-root');
    second.instanceId = 'second-tabs';
    second.displayMode = 1;
    second.properties = {
      ...page.webpart.properties, tabDesign: 'blocks', tabAlignment: 'center',
      tabData: secondIds.map((id, index) => ({ WebPartID: id, TabLabel: `Other ${index + 1}` }))
    };
    page.webpart.properties.tabDesign = 'pills';
    page.webpart.properties.tabAlignment = 'stretch';
    page.webpart.render();
    second.render();
    page.webpart.domElement.querySelectorAll('[role="tab"]')[1].click();
    const firstList = page.webpart.domElement.querySelector('.hillbilly-tabs-list');
    const secondList = second.domElement.querySelector('.hillbilly-tabs-list');
    assert.equal(firstList.dataset.tabDesign, 'pills');
    assert.equal(firstList.dataset.tabAlignment, 'stretch');
    assert.equal(secondList.dataset.tabDesign, 'blocks');
    assert.equal(secondList.dataset.tabAlignment, 'center');
    assert.equal(firstList.querySelectorAll('[role="tab"]')[1].getAttribute('aria-selected'), 'true');
    assert.equal(secondList.querySelectorAll('[role="tab"]')[0].getAttribute('aria-selected'), 'true');
  } finally { second?.onDispose(); page.close(); }
});

test('edit mode previews the selected appearance without controlling real page content', async () => {
  const page = createPage();
  try {
    const actualParts = [page.target(0), page.target(1)];
    const initialState = actualParts.map(originalState);
    const initialParents = actualParts.map(part => part.parentElement);
    page.webpart.properties.tabDesign = 'segmented';
    page.webpart.properties.tabAlignment = 'center';
    page.webpart.render();
    page.webpart.displayMode = 2;
    page.webpart.render();
    const preview = page.webpart.domElement.querySelector('.hillbilly-tabs-preview');
    assert.ok(preview, 'Appearance settings need an in-place edit preview');
    const tablist = preview.querySelector('.hillbilly-tabs-list');
    assert.equal(tablist.dataset.tabDesign, 'segmented');
    assert.equal(tablist.dataset.tabAlignment, 'center');
    const previewTabs = tablist.querySelectorAll('[role="tab"]');
    assert.ok(previewTabs.length >= 2, 'preview should show the contrast between selected and unselected tabs');
    previewTabs[1].click();
    assert.equal(previewTabs[1].getAttribute('aria-selected'), 'true', 'preview buttons must demonstrate the selected appearance');
    const previewPanel = page.window.document.getElementById(previewTabs[1].getAttribute('aria-controls'));
    assert.ok(previewPanel && preview.contains(previewPanel), 'preview tabs must reference local preview content');
    assert.equal(previewPanel.hidden, false);
    previewTabs[1].focus();
    previewTabs[1].dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'Home', bubbles: true, cancelable: true }));
    assert.equal(page.window.document.activeElement, previewTabs[0]);
    assert.equal(previewTabs[0].getAttribute('aria-selected'), 'true');
    page.webpart.properties.tabDesign = 'outlined';
    page.webpart.properties.tabAlignment = 'stretch';
    page.webpart.render();
    const updatedPreview = page.webpart.domElement.querySelector('.hillbilly-tabs-preview');
    assert.notEqual(updatedPreview, preview, 'editing Appearance must refresh the local preview');
    assert.equal(updatedPreview.querySelector('.hillbilly-tabs-list').dataset.tabDesign, 'outlined');
    assert.equal(updatedPreview.querySelector('.hillbilly-tabs-list').dataset.tabAlignment, 'stretch');
    updatedPreview.querySelectorAll('[role="tab"]')[2].click();
    await tick();
    actualParts.forEach((part, index) => {
      assert.deepEqual(originalState(part), initialState[index], 'preview interaction must not alter real SharePoint controls');
      assert.equal(part.parentElement, initialParents[index]);
    });
    assert.match(page.webpart.domElement.textContent, /About/);
    assert.match(page.webpart.domElement.textContent, /Dr\. Ghassan Abukar, Ph\.D\./);
    assert.match(page.webpart.domElement.textContent, /Original Modern Hillbilly Tabs project by Mark Rackley/);
  } finally { page.close(); }
});

for (const themeCase of [
  { name: 'light', accent: '#f7dc6f', accentText: '#000000', background: '#fff9e6', text: '#302b1d' },
  { name: 'dark', accent: '#003366', accentText: '#ffffff', background: '#131c27', text: '#edf2f7' }
]) {
  test(`${themeCase.name} SharePoint theme updates contrast and semantic colors without resetting selection`, async () => {
    const page = createPage();
    try {
      page.webpart.properties.tabDesign = 'blocks';
      page.webpart.render();
      const tablist = page.webpart.domElement.querySelector('.hillbilly-tabs-list');
      const tabs = tablist.querySelectorAll('[role="tab"]');
      const savedTabData = JSON.stringify(page.webpart.properties.tabData);
      tabs[1].click();
      page.webpart.onThemeChanged({
        palette: { themePrimary: themeCase.accent, neutralTertiary: '#708090' },
        semanticColors: {
          bodyBackground: themeCase.background, bodyText: themeCase.text,
          bodyStandoutBackground: '#667788', inputBorder: '#8899aa',
          buttonBackgroundHovered: '#556677', focusBorder: '#445566'
        }
      });
      await tick();
      const style = page.webpart.domElement.style;
      assert.equal(style.getPropertyValue('--hillbilly-accent'), themeCase.accent);
      assert.equal(style.getPropertyValue('--hillbilly-accent-text'), themeCase.accentText);
      assert.equal(style.getPropertyValue('--hillbilly-surface'), themeCase.background);
      assert.equal(style.getPropertyValue('--hillbilly-text'), themeCase.text);
      assert.equal(style.getPropertyValue('--hillbilly-surface-muted'), '#667788');
      assert.equal(style.getPropertyValue('--hillbilly-border'), '#8899aa');
      assert.equal(style.getPropertyValue('--hillbilly-hover'), '#556677');
      assert.equal(style.getPropertyValue('--hillbilly-focus'), '#445566');
      assert.equal(page.webpart.domElement.querySelector('.hillbilly-tabs-list'), tablist, 'theme changes must not re-render the tab list');
      assert.equal(tabs[1].getAttribute('aria-selected'), 'true');
      assert.equal(page.hidden(page.target(0)), true);
      assert.equal(page.hidden(page.target(1)), false);
      assert.equal(JSON.stringify(page.webpart.properties.tabData), savedTabData);
    } finally { page.close(); }
  });
}

function configureCategories(page, count = 24, prefix = 'category', webpart = page.webpart) {
  const section = webpart.domElement.closest('section');
  const categories = Array.from({ length: count }, (_, index) => ({
    WebPartID: `${prefix}-${index + 1}`,
    TabLabel: `Category ${String(index + 1).padStart(2, '0')}`
  }));
  section.insertAdjacentHTML('beforeend', categories.map(category => page.control(category.WebPartID, `${category.TabLabel} content`)).join(''));
  webpart.properties.tabData = categories;
  const parts = categories.map(category => section.querySelector(`[data-sp-feature-instance-id="${category.WebPartID}"]`).parentElement);
  return { categories, parts, parents: parts.map(part => part.parentElement) };
}

function categoryMenu(webpart) {
  const summary = webpart.domElement.querySelector('summary.hillbilly-category-summary');
  assert.ok(summary, 'category navigation must expose a native summary control');
  const details = summary.closest('details');
  assert.ok(details);
  const search = details.querySelector('.hillbilly-category-search');
  assert.ok(search, 'category navigation needs an accessible search input');
  assert.ok(search.getAttribute('aria-label') || details.querySelector(`label[for="${search.id}"]`), 'search must have an accessible name');
  const buttons = Array.from(details.querySelectorAll('[role="tab"]'));
  return { summary, details, search, buttons, selected: summary.querySelector('.hillbilly-category-selected') };
}

function filterCategories(page, menu, query) {
  menu.search.value = query;
  menu.search.dispatchEvent(new page.window.Event('input', { bubbles: true }));
}

test('new web parts default to automatic navigation while existing undefined settings keep tabs', () => {
  const page = createPage(true, false);
  try {
    configureCategories(page);
    assert.equal(page.webpart.properties.tabLayout, undefined);
    page.webpart.render();
    assert.equal(page.webpart.domElement.querySelector('.hillbilly-category-summary'), null);
    const appearance = page.webpart.getPropertyPaneConfiguration().pages[0].groups.find(group => group.groupName === 'Appearance');
    const navigation = appearance.groupFields.find(field => field.targetProperty === 'tabLayout');
    assert.ok(navigation);
    assert.equal(navigation.selectedKey, 'tabs');
    assert.deepEqual(Array.from(navigation.options, option => option.key), ['tabs', 'menu', 'auto']);
    const manifestPath = path.join(__dirname, '../src/webparts/modernHillbillyTabs/ModernHillbillyTabsWebPart.manifest.json');
    const parsed = ts.parseConfigFileTextToJson(manifestPath, fs.readFileSync(manifestPath, 'utf8'));
    assert.equal(parsed.error, undefined);
    assert.equal(parsed.config.preconfiguredEntries[0].properties.tabLayout, 'auto');
  } finally { page.close(); }
});

for (const layout of [
  { setting: 'tabs', count: 24, menu: false },
  { setting: 'menu', count: 2, menu: true },
  { setting: 'auto', count: 8, menu: false },
  { setting: 'auto', count: 9, menu: true },
  { setting: 'auto', count: 24, menu: true },
  { setting: 'unexpected-layout', count: 24, menu: false }
]) {
  test(`${layout.setting} navigation with ${layout.count} categories uses ${layout.menu ? 'a category menu' : 'tabs'}`, () => {
    const page = createPage(true, false);
    try {
      configureCategories(page, layout.count);
      page.webpart.properties.tabLayout = layout.setting;
      page.webpart.render();
      assert.equal(!!page.webpart.domElement.querySelector('.hillbilly-category-summary'), layout.menu);
      assert.equal(page.webpart.domElement.querySelectorAll('[role="tab"]').length, layout.count);
    } finally { page.close(); }
  });
}

test('a 24-category menu shows the current category and filters labels without altering content or settings', async () => {
  const page = createPage(true, false);
  try {
    const { categories, parts, parents } = configureCategories(page);
    page.webpart.properties.tabLayout = 'menu';
    const savedData = JSON.stringify(categories);
    page.webpart.render();
    const menu = categoryMenu(page.webpart);
    assert.ok(menu.selected);
    assert.equal(menu.selected.textContent.trim(), 'Category 01');
    assert.match(menu.summary.textContent, /24/);
    assert.equal(menu.details.open, false);
    menu.details.open = true;
    await tick();
    filterCategories(page, menu, 'category 2');
    assert.deepEqual(menu.buttons.filter(button => !button.hidden).map(button => button.textContent), ['Category 20', 'Category 21', 'Category 22', 'Category 23', 'Category 24']);
    assert.equal(menu.details.open, true);
    assert.equal(menu.selected.textContent.trim(), 'Category 01', 'filtering alone must not change the current category');
    assert.equal(page.hidden(parts[0]), false);
    assert.ok(parts.slice(1).every(page.hidden));
    assert.equal(JSON.stringify(page.webpart.properties.tabData), savedData);
    parts.forEach((part, index) => assert.equal(part.parentElement, parents[index]));
  } finally { page.close(); }
});

test('choosing a filtered category updates the summary, closes the menu, and focuses its summary', async () => {
  const page = createPage(true, false);
  try {
    const { parts } = configureCategories(page);
    page.webpart.properties.tabLayout = 'menu';
    page.webpart.render();
    const menu = categoryMenu(page.webpart);
    menu.details.open = true;
    await tick();
    filterCategories(page, menu, 'Category 24');
    menu.buttons[23].click();
    assert.equal(menu.selected.textContent.trim(), 'Category 24');
    assert.equal(menu.details.open, false);
    assert.equal(page.window.document.activeElement, menu.summary);
    assert.equal(menu.buttons[23].getAttribute('aria-selected'), 'true');
    assert.equal(page.hidden(parts[23]), false);
    assert.ok(parts.slice(0, 23).every(page.hidden));
  } finally { page.close(); }
});

test('category search and arrow keys navigate only visible usable choices while Escape returns to the summary', async () => {
  const page = createPage(true, false);
  try {
    const { parts } = configureCategories(page);
    // The first search match is configured but unavailable; keyboard navigation must skip it.
    parts[19].remove();
    page.webpart.properties.tabLayout = 'menu';
    page.webpart.render();
    const menu = categoryMenu(page.webpart);
    menu.details.open = true;
    await tick();
    filterCategories(page, menu, 'Category 2');
    menu.search.focus();
    menu.search.dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true }));
    assert.equal(page.window.document.activeElement, menu.buttons[20]);
    assert.equal(menu.buttons[19].disabled, true);
    menu.buttons[20].dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true }));
    assert.equal(page.window.document.activeElement, menu.buttons[21]);
    assert.equal(menu.buttons[21].getAttribute('aria-selected'), 'true');
    assert.equal(menu.details.open, true, 'arrow navigation must not close the category menu');
    menu.buttons[21].dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'Home', bubbles: true, cancelable: true }));
    assert.equal(page.window.document.activeElement, menu.buttons[20]);
    menu.buttons[20].dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'End', bubbles: true, cancelable: true }));
    assert.equal(page.window.document.activeElement, menu.buttons[23]);
    menu.buttons[23].dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    assert.equal(menu.details.open, false);
    assert.equal(page.window.document.activeElement, menu.summary);
    assert.equal(page.hidden(parts[23]), false);
  } finally { page.close(); }
});

test('Enter from category search chooses the first available match', async () => {
  const page = createPage(true, false);
  try {
    const { parts } = configureCategories(page);
    parts[19].remove();
    page.webpart.properties.tabLayout = 'menu';
    page.webpart.render();
    const menu = categoryMenu(page.webpart);
    menu.details.open = true;
    await tick();
    filterCategories(page, menu, 'Category 2');
    menu.search.dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
    assert.equal(menu.selected.textContent.trim(), 'Category 21');
    assert.equal(menu.details.open, false);
    assert.equal(page.window.document.activeElement, menu.summary);
    assert.equal(page.hidden(parts[20]), false);
  } finally { page.close(); }
});

test('no matching categories announces an empty result and recovers without changing the active content', async () => {
  const page = createPage(true, false);
  try {
    const { parts } = configureCategories(page);
    page.webpart.properties.tabLayout = 'menu';
    page.webpart.render();
    const menu = categoryMenu(page.webpart);
    menu.details.open = true;
    await tick();
    filterCategories(page, menu, 'does not exist');
    assert.ok(menu.buttons.every(button => button.hidden));
    const status = Array.from(menu.details.querySelectorAll('[role="status"]')).find(element => /no.*(categor|match)|no results/i.test(element.textContent));
    assert.ok(status && !status.hidden, 'empty results must be announced accessibly');
    assert.equal(page.hidden(parts[0]), false);
    const emptyEnter = new page.window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    menu.search.dispatchEvent(emptyEnter);
    assert.equal(emptyEnter.defaultPrevented, true, 'Enter with no matches must not submit an enclosing page form');
    assert.equal(menu.selected.textContent.trim(), 'Category 01');
    assert.equal(menu.details.open, true);
    filterCategories(page, menu, '');
    assert.equal(menu.buttons.filter(button => !button.hidden).length, 24);
    assert.ok(status.hidden || !status.textContent.trim(), 'the empty result message must clear when matches return');
    assert.equal(page.hidden(parts[0]), false);
  } finally { page.close(); }
});

test('category filtering and selection stay independent between two tab web parts', async () => {
  const page = createPage(true, false);
  let second;
  try {
    configureCategories(page);
    page.webpart.properties.tabLayout = 'menu';
    page.window.document.body.insertAdjacentHTML('beforeend', `<section data-automation-id="CanvasSection">${page.control('second-menu-host', '<div id="second-menu-root"></div>')}</section>`);
    second = new page.webpart.constructor();
    second.domElement = page.window.document.getElementById('second-menu-root');
    second.instanceId = 'second-menu';
    second.displayMode = 1;
    second.properties = { ...page.webpart.properties, tabData: [] };
    configureCategories(page, 12, 'second-category', second);
    page.webpart.render();
    second.render();
    const firstMenu = categoryMenu(page.webpart);
    const secondMenu = categoryMenu(second);
    firstMenu.details.open = true;
    await tick();
    filterCategories(page, firstMenu, 'Category 24');
    firstMenu.buttons[23].click();
    assert.equal(secondMenu.selected.textContent.trim(), 'Category 01');
    assert.equal(secondMenu.search.value, '');
    assert.equal(secondMenu.buttons.filter(button => !button.hidden).length, 12);
    secondMenu.details.open = true;
    await tick();
    filterCategories(page, secondMenu, 'Category 11');
    secondMenu.buttons[10].click();
    assert.equal(secondMenu.selected.textContent.trim(), 'Category 11');
    assert.equal(firstMenu.selected.textContent.trim(), 'Category 24');
    assert.equal(firstMenu.buttons[23].getAttribute('aria-selected'), 'true');
  } finally { second?.onDispose(); page.close(); }
});

test('category menu edit preview stays local to its sample content', async () => {
  const page = createPage();
  try {
    const parts = [page.target(0), page.target(1)];
    const states = parts.map(originalState);
    page.webpart.properties.tabLayout = 'menu';
    page.webpart.displayMode = 2;
    page.webpart.render();
    const preview = page.webpart.domElement.querySelector('.hillbilly-tabs-preview');
    const menu = categoryMenu({ domElement: preview });
    menu.details.open = true;
    await tick();
    menu.buttons[2].click();
    assert.equal(menu.selected.textContent.trim(), menu.buttons[2].textContent);
    assert.equal(menu.details.open, false);
    const panel = page.window.document.getElementById(menu.buttons[2].getAttribute('aria-controls'));
    assert.ok(preview.contains(panel));
    parts.forEach((part, index) => assert.deepEqual(originalState(part), states[index]));
  } finally { page.close(); }
});

test('edit mode shows the loaded version and update setting without the original-project link', () => {
  const page = createPage();
  try {
    page.webpart.displayMode = 2;
    page.webpart.properties.checkForUpdates = false;
    page.webpart.render();
    assert.match(page.webpart.domElement.querySelector('.hillbilly-update-version').textContent, /1\.2\.0/);
    assert.match(page.webpart.domElement.querySelector('.hillbilly-update-status').textContent, /turned off/);
    assert.equal(page.webpart.domElement.querySelector('a[href*="mrackley"]'), null);
    assert.match(page.webpart.domElement.textContent, /Original Modern Hillbilly Tabs project by Mark Rackley/);
    const group = page.webpart.getPropertyPaneConfiguration().pages[0].groups.find(item => item.groupName === 'Updates');
    assert.equal(group.groupFields[0].targetProperty, 'checkForUpdates');
    assert.equal(group.groupFields[0].checked, false);
  } finally { page.close(); }
});

test('ordinary page readers do not create an update notice or make a release request', async () => {
  const page = createPage();
  try {
    let requests = 0;
    page.window.fetch = async () => { requests++; throw new Error('must not fetch for ordinary readers'); };
    page.webpart.context = { pageContext: { web: { permissions: { hasPermission: () => false } } } };
    page.webpart.render();
    await tick();
    assert.equal(requests, 0);
    assert.equal(page.webpart.domElement.querySelector('.hillbilly-update-notice'), null);
    assert.equal(page.hidden(page.target(1)), true);
  } finally { page.close(); }
});

test('site owners see a newer-release notice while published tabs keep switching correctly', async () => {
  const page = createPage();
  try {
    let requests = 0;
    page.window.fetch = async () => {
      requests++;
      return { ok: true, status: 200, headers: { get: () => null }, json: async () => ({
        tag_name: 'v1.3.0', draft: false, prerelease: false,
        html_url: 'https://github.com/drghassanabukar/modern-hillbilly-tabs-365/releases/tag/v1.3.0',
        assets: [{ name: 'modern-hillbilly-tabs.sppkg', state: 'uploaded', browser_download_url: 'https://github.com/drghassanabukar/modern-hillbilly-tabs-365/releases/download/v1.3.0/modern-hillbilly-tabs.sppkg' }]
      }) };
    };
    page.webpart.context = { pageContext: { web: { permissions: { hasPermission: permission => permission === 'manageWeb' } } } };
    page.webpart.render();
    await tick();
    const notice = page.webpart.domElement.querySelector('.hillbilly-update-notice');
    assert.equal(requests, 1);
    assert.equal(notice.hidden, false);
    assert.equal(notice.dataset.state, 'available');
    assert.match(notice.textContent, /1\.3\.0 is available/);
    assert.ok(Array.from(notice.querySelectorAll('a')).some(link => link.textContent === 'Download update'));
    page.webpart.domElement.querySelectorAll('[role="tab"]')[1].click();
    assert.equal(page.hidden(page.target(1)), false);
    assert.equal(page.hidden(page.target(0)), true);
  } finally { page.close(); }
});

module.exports = { createPage, ids, tick };
